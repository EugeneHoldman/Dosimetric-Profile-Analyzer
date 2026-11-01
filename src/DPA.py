
#!/usr/bin/env python3
from __future__ import annotations

import argparse
import math
import re
import sys
from collections import defaultdict
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, List, Optional, Sequence, Tuple

import numpy as np
import pydicom
from pydicom.misc import is_dicom
from openpyxl import Workbook
from openpyxl.chart import LineChart, Reference, ScatterChart, Series
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.drawing.line import LineProperties
from openpyxl.styles import PatternFill
from scipy.interpolate import RegularGridInterpolator

SAD_MM = 1000.0
LINE_WIDTH_1PT = 12700
YELLOW_FILL = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
RED_FILL = PatternFill(start_color="F4CCCC", end_color="F4CCCC", fill_type="solid")
SMALL_FIELD_MAX_MM = 30.0
SMALL_FIELD_DOSE_THRESHOLD_PERCENT = 20.0
PROFILE_TYPE_ORDER = {"crossline": 1, "inline": 2, "diagonal": 3}


@dataclass
class AscPoint:
    x: float
    y: float
    z: float
    dose: float


@dataclass
class Meta:
    machine: str = ""
    field_comment: str = ""
    scan_id: str = ""
    date: str = ""
    typ: str = ""
    axis: str = ""
    pnts: str = ""
    step: str = ""
    ssd: str = ""
    flsz: str = ""
    dpth: str = ""


@dataclass
class Profile:
    asc_path: Path
    index: int
    meta: Meta
    points: List[AscPoint] = field(default_factory=list)
    parse_error: Optional[str] = None


@dataclass
class DicomIndex:
    rtdose_by_uid: Dict[str, Path]
    rtplan_by_uid: Dict[str, Path]
    candidate_by_key: Dict[Tuple[str, float, str], "CandidatePair"]


@dataclass
class CandidatePair:
    rtdose_path: Path
    rtplan_path: Path
    plan_label: str
    energy: float
    mode: str


@dataclass
class Grid:
    x: np.ndarray
    y: np.ndarray
    z: np.ndarray
    interp: RegularGridInterpolator


@dataclass
class SummaryRecord:
    profile_name: str
    family: str
    depth_mm: float
    field_mm: float
    max_diff_left: Optional[float]
    max_diff_right: Optional[float]
    n_points_1_2: int
    n_points_gt2: int
    pct_points_1_2: Optional[float]
    pct_points_gt2: Optional[float]
    avg_diff_left: Optional[float]
    avg_diff_right: Optional[float]


class StartupError(RuntimeError):
    pass


class ProfileSkip(RuntimeError):
    pass


def discover(root: Path) -> Tuple[List[Path], List[Path]]:
    asc, dcm = [], []
    try:
        for p in root.rglob("*"):
            if not p.is_file():
                continue
            s = p.suffix.lower()
            if s == ".asc":
                asc.append(p)
            elif s == ".dcm":
                dcm.append(p)
    except Exception as exc:
        raise StartupError(f"Global folder/file discovery failed: {exc}") from exc
    return sorted(asc), sorted(dcm)


def index_dicoms(dcm_files: Sequence[Path]) -> DicomIndex:
    dose_coll: Dict[str, List[Path]] = defaultdict(list)
    dose_lookup: Dict[str, Path] = {}
    plan_coll: Dict[str, List[Path]] = defaultdict(list)
    plan_lookup: Dict[str, Path] = {}
    for path in dcm_files:
        try:
            ds = pydicom.dcmread(path, stop_before_pixels=True, force=True)
        except Exception as exc:
            # Keep non-DICOM *.dcm placeholders from aborting startup indexing.
            try:
                dicom_like = is_dicom(path)
            except Exception:
                dicom_like = True
            if dicom_like:
                raise StartupError(f"Unreadable critical DICOM during startup indexing: {path} ({exc})") from exc
            print(
                f"WARN: Skipping unreadable non-DICOM .dcm file during startup indexing: {path} ({exc})",
                file=sys.stderr,
            )
            continue
        mod = str(getattr(ds, "Modality", "")).upper()
        if mod == "RTDOSE":
            uid = str(getattr(ds, "SOPInstanceUID", "")).strip()
            if not uid:
                raise StartupError(f"RTDOSE missing SOPInstanceUID: {path}")
            dose_coll[uid].append(path)
            dose_lookup.setdefault(uid, path)
        elif mod == "RTPLAN":
            uid = str(getattr(ds, "SOPInstanceUID", "")).strip()
            if not uid:
                raise StartupError(f"RTPLAN missing SOPInstanceUID: {path}")
            plan_coll[uid].append(path)
            plan_lookup.setdefault(uid, path)

    dup_dose_uid = {k: v for k, v in dose_coll.items() if len(v) > 1}
    if dup_dose_uid:
        details = "\n".join(f"  {k}: {', '.join(str(p) for p in v)}" for k, v in sorted(dup_dose_uid.items()))
        raise StartupError(
            "Duplicated RTDOSE SOPInstanceUID values were found. RTDOSE SOPInstanceUID values must be unique.\n"
            + details
        )

    dup_uid = {k: v for k, v in plan_coll.items() if len(v) > 1}
    if dup_uid:
        details = "\n".join(f"  {k}: {', '.join(str(p) for p in v)}" for k, v in sorted(dup_uid.items()))
        raise StartupError(
            "Duplicated RTPLAN SOPInstanceUID values were found. RTPLAN SOPInstanceUID values must be unique.\n"
            + details
        )

    candidate_by_key = build_candidate_matching_index(dose_lookup, plan_lookup)
    return DicomIndex(dose_lookup, plan_lookup, candidate_by_key)


def parse_asc(path: Path) -> List[Profile]:
    lines = path.read_text(encoding="utf-8", errors="replace").splitlines()
    blocks: List[Tuple[List[str], Optional[str]]] = []
    cur: List[str] = []
    in_block = False
    for ln in lines:
        s = ln.strip()
        if s == "$STOM":
            if in_block and cur:
                blocks.append((cur, "Malformed ASC profile block: missing $ENOM terminator before next $STOM"))
            in_block = True
            cur = [ln]
            continue
        if in_block:
            cur.append(ln)
            if s == "$ENOM":
                blocks.append((cur, None))
                cur = []
                in_block = False
    if in_block and cur:
        blocks.append((cur, "Malformed ASC profile block: missing $ENOM terminator"))
    out = []
    for i, (raw, parse_error) in enumerate(blocks):
        out.append(parse_block(path, i, raw, parse_error=parse_error))
    return out


def looks_like_field_comment(comment: str) -> bool:
    cu = comment.upper()
    return bool(
        re.search(
            r"\d+(?:\.\d+)?\s*(?:CM|MM)?\s*[X*]\s*\d+(?:\.\d+)?\s*(?:CM|MM)?",
            cu,
        )
    )


def parse_block(path: Path, idx: int, raw: Sequence[str], parse_error: Optional[str] = None) -> Profile:
    m = Meta()
    pts: List[AscPoint] = []
    comments: List[str] = []
    kv = re.compile(r"^%(\w+)\s+(.*)$")
    pt = re.compile(r"^<\s*([+-]?\d+(?:\.\d+)?)\s+([+-]?\d+(?:\.\d+)?)\s+([+-]?\d+(?:\.\d+)?)\s+([+-]?\d+(?:\.\d+)?)\s*>$")
    for ln in raw:
        s = ln.strip()
        if not s:
            continue
        if s.startswith("#"):
            comments.append(s.lstrip("#").strip())
            continue
        mk = kv.match(s)
        if mk:
            k, v = mk.group(1).upper(), mk.group(2).strip()
            if k == "DATE":
                m.date = v
            elif k == "TYPE":
                m.typ = v.upper()
            elif k == "AXIS":
                m.axis = v.upper()
            elif k == "PNTS":
                m.pnts = v
            elif k == "STEP":
                m.step = v
            elif k == "SSD":
                m.ssd = v
            elif k == "FLSZ":
                m.flsz = v
            elif k == "DPTH":
                m.dpth = v
            continue
        mp = pt.match(s)
        if mp:
            pts.append(AscPoint(float(mp.group(1)), float(mp.group(2)), float(mp.group(3)), float(mp.group(4))))

    for c in comments:
        cu = c.upper()
        if not m.machine and "MV" in cu:
            m.machine = c
        elif not m.scan_id and "SCAN ID" in cu:
            m.scan_id = c
        elif not m.field_comment and looks_like_field_comment(c):
            m.field_comment = c

    if not m.field_comment:
        for c in comments:
            if not c or c == m.machine or c == m.scan_id:
                continue
            m.field_comment = c
            break

    return Profile(path, idx, m, pts, parse_error=parse_error)


def classify(profile: Profile) -> str:
    t, a = profile.meta.typ.upper(), profile.meta.axis.upper()
    if t == "OPP" and a == "Y":
        return "inline"
    if t == "OPP" and a == "X":
        return "crossline"
    if t == "OPD":
        return "pdd"
    if t == "DPR":
        return "diagonal"
    raise ProfileSkip(f"Unsupported %TYPE/%AXIS combination: %TYPE={t}, %AXIS={a}")


def normalize_plan_label(label: str) -> str:
    compact = re.sub(r"\s+", "", label.strip())
    if not compact:
        raise ProfileSkip("RT Plan Label is missing/empty")
    return compact.lower()


def normalize_energy(value: float) -> float:
    if not np.isfinite(value):
        raise ProfileSkip("Invalid nominal beam energy value")
    if math.isclose(value, round(value), abs_tol=1e-6):
        return float(int(round(value)))
    return float(round(value, 6))


def format_compact_number(value: float) -> str:
    if math.isclose(value, round(value), abs_tol=1e-6):
        return str(int(round(value)))
    return f"{value:.6f}".rstrip("0").rstrip(".")


def parse_expected_plan_label_from_asc(profile: Profile) -> str:
    m = re.fullmatch(r"\s*(\d+(?:\.\d+)?)\s*\*\s*(\d+(?:\.\d+)?)\s*", profile.meta.flsz)
    if not m:
        raise ProfileSkip(f"Malformed %FLSZ: '{profile.meta.flsz}'")
    a_mm = float(m.group(1))
    b_mm = float(m.group(2))
    if not math.isclose(a_mm, b_mm, rel_tol=0.0, abs_tol=1e-6):
        raise ProfileSkip(f"Only square symmetric fields are supported: '{profile.meta.flsz}'")

    ssd_match = re.fullmatch(r"\s*(\d+(?:\.\d+)?)\s*", profile.meta.ssd)
    if not ssd_match:
        raise ProfileSkip(f"Malformed %SSD: '{profile.meta.ssd}'")
    ssd_mm_value = float(ssd_match.group(1))

    field_cm = a_mm / 10.0
    ssd_cm = ssd_mm_value / 10.0
    return f"{format_compact_number(field_cm)}x{format_compact_number(field_cm)}_SSD{format_compact_number(ssd_cm)}"


def parse_expected_energy_mode_from_asc_comment(comment: str) -> Tuple[float, str]:
    if not comment:
        raise ProfileSkip("Missing machine / energy comment line")
    mv_match = re.search(r"(\d+(?:\.\d+)?)\s*MV\b", comment, re.IGNORECASE)
    if not mv_match:
        raise ProfileSkip(f"Cannot parse energy from machine/energy comment line: '{comment}'")
    energy = normalize_energy(float(mv_match.group(1)))
    mode = "FFF" if re.search(r"\(\s*FFF\s*\)", comment, re.IGNORECASE) else "FF"
    return energy, mode


def extract_referenced_rtplan_uid_from_rtdose(rtdose_ds) -> str:
    seq = getattr(rtdose_ds, "ReferencedRTPlanSequence", None)
    if not seq:
        raise ProfileSkip("RTDOSE -> RTPLAN link missing (ReferencedRTPlanSequence)")
    if len(seq) != 1:
        raise ProfileSkip("Ambiguous RTDOSE -> RTPLAN link")
    uid = str(getattr(seq[0], "ReferencedSOPInstanceUID", "")).strip()
    if not uid:
        raise ProfileSkip("RTDOSE ReferencedSOPInstanceUID is missing")
    return uid


def extract_plan_label_from_rtplan(rtplan_ds) -> str:
    label = str(getattr(rtplan_ds, "RTPlanLabel", "")).strip()
    if not label:
        raise ProfileSkip("RTPLAN missing RT Plan Label (300A,0002)")
    return label


def _extract_beam_nominal_energy(beam) -> float:
    val = getattr(beam, "NominalBeamEnergy", None)
    if val is None:
        cps = getattr(beam, "ControlPointSequence", None)
        if cps and len(cps) > 0:
            val = getattr(cps[0], "NominalBeamEnergy", None)
    if val is None:
        raise ProfileSkip("RTPLAN treatment beam missing Nominal Beam Energy (300A,0114)")
    try:
        return normalize_energy(float(val))
    except (TypeError, ValueError) as exc:
        raise ProfileSkip(f"Invalid Nominal Beam Energy value: '{val}'") from exc


def _extract_beam_mode(beam) -> str:
    mode = None
    mode_id = None

    # Preferred source in RTPLAN beam metadata.
    primary_seq = getattr(beam, "PrimaryFluenceModeSequence", None)
    if primary_seq and len(primary_seq) > 0:
        item = primary_seq[0]
        mode = getattr(item, "FluenceMode", None)
        mode_id = getattr(item, "FluenceModeID", None)

    # Backward-compatible fallbacks.
    if mode is None:
        mode = getattr(beam, "FluenceMode", None)
        mode_id = getattr(beam, "FluenceModeID", None)
    if mode is None:
        cps = getattr(beam, "ControlPointSequence", None)
        if cps and len(cps) > 0:
            mode = getattr(cps[0], "FluenceMode", None)
            mode_id = getattr(cps[0], "FluenceModeID", None)
    mode_txt = str(mode).strip().upper()
    mode_id_txt = str(mode_id).strip().upper()

    if mode_txt == "STANDARD":
        return "FF"
    if mode_txt == "NON_STANDARD":
        if mode_id_txt == "FFF":
            return "FFF"
        raise ProfileSkip(
            f"RTPLAN NON_STANDARD fluence mode without FluenceModeID=FFF (found '{mode_id_txt or 'missing'}')"
        )
    raise ProfileSkip(f"Unsupported/missing Fluence Mode (3002,0051): '{mode_txt or 'missing'}'")


def extract_beam_energy_mode_from_rtplan(rtplan_ds) -> Tuple[float, str]:
    beam_seq = getattr(rtplan_ds, "BeamSequence", None)
    if not beam_seq:
        raise ProfileSkip("RTPLAN missing BeamSequence")

    usable = []
    for beam in beam_seq:
        radiation = str(getattr(beam, "RadiationType", "")).strip().upper()
        delivery = str(getattr(beam, "TreatmentDeliveryType", "")).strip().upper()
        if radiation != "PHOTON":
            continue
        if delivery and delivery != "TREATMENT":
            continue
        usable.append(beam)

    if len(usable) != 1:
        raise ProfileSkip(
            f"RTPLAN must contain exactly one usable photon treatment beam for matching (found {len(usable)})"
        )

    beam = usable[0]
    return _extract_beam_nominal_energy(beam), _extract_beam_mode(beam)


def build_candidate_matching_index(
    rtdose_by_uid: Dict[str, Path], rtplan_by_uid: Dict[str, Path]
) -> Dict[Tuple[str, float, str], CandidatePair]:
    candidate_by_key: Dict[Tuple[str, float, str], CandidatePair] = {}

    for _dose_uid, rtdose_path in sorted(rtdose_by_uid.items()):
        try:
            rtdose_ds = pydicom.dcmread(rtdose_path, stop_before_pixels=True, force=True)
            plan_uid = extract_referenced_rtplan_uid_from_rtdose(rtdose_ds)
            rtplan_path = rtplan_by_uid.get(plan_uid)
            if rtplan_path is None:
                raise ProfileSkip(f"Referenced RTPLAN not found for SOPInstanceUID '{plan_uid}'")

            rtplan_ds = pydicom.dcmread(rtplan_path, stop_before_pixels=True, force=True)
            plan_label = extract_plan_label_from_rtplan(rtplan_ds)
            energy, mode = extract_beam_energy_mode_from_rtplan(rtplan_ds)
            key = (normalize_plan_label(plan_label), energy, mode)
            pair = CandidatePair(
                rtdose_path=rtdose_path,
                rtplan_path=rtplan_path,
                plan_label=plan_label,
                energy=energy,
                mode=mode,
            )

            prev = candidate_by_key.get(key)
            if prev is not None:
                raise StartupError(
                    "Ambiguous duplicate RTDOSE/RTPLAN candidates found for the same "
                    f"field/SSD/energy key {key}: '{prev.rtdose_path}' and '{rtdose_path}'"
                )
            candidate_by_key[key] = pair
        except ProfileSkip as exc:
            print(
                f"WARN: Excluding RTDOSE candidate '{rtdose_path}' from matching index: {exc}",
                file=sys.stderr,
            )
        except StartupError:
            raise
        except Exception as exc:
            raise StartupError(
                f"Unreadable critical DICOM during startup indexing: {rtdose_path} ({exc})"
            ) from exc

    return candidate_by_key

def origin_from_rtplan(rtplan_ds) -> np.ndarray:
    seq = getattr(rtplan_ds, "DoseReferenceSequence", None)
    if not seq:
        raise ProfileSkip("RTPLAN missing DoseReferenceSequence")
    matches = []
    for item in seq:
        name = str(getattr(item, "DoseReferenceDescription", "")).strip().lower()
        if name != "origin":
            continue
        coords = getattr(item, "DoseReferencePointCoordinates", None)
        if coords is None:
            continue
        arr = np.asarray(coords, dtype=float).reshape(-1)
        if arr.size == 3:
            matches.append(arr)
    if len(matches) == 0:
        raise ProfileSkip("ORIGIN dose-reference point not found")
    if len(matches) > 1:
        raise ProfileSkip("ORIGIN dose-reference point is ambiguous")
    return matches[0]


def _ascending(axis: np.ndarray, cube: np.ndarray, axis_index: int) -> Tuple[np.ndarray, np.ndarray]:
    if axis[0] <= axis[-1]:
        return axis, cube
    return axis[::-1], np.flip(cube, axis=axis_index)


def build_grid(rtdose_ds, origin_mm: np.ndarray) -> Grid:
    scale = float(getattr(rtdose_ds, "DoseGridScaling", 0.0))
    if scale <= 0:
        raise ProfileSkip("Invalid/non-positive DoseGridScaling")

    ipp = np.asarray(getattr(rtdose_ds, "ImagePositionPatient", None), dtype=float)
    if ipp.size != 3:
        raise ProfileSkip("Invalid ImagePositionPatient")

    spacing = np.asarray(getattr(rtdose_ds, "PixelSpacing", None), dtype=float)
    if spacing.size != 2:
        raise ProfileSkip("Invalid PixelSpacing")

    rows = int(getattr(rtdose_ds, "Rows", 0))
    cols = int(getattr(rtdose_ds, "Columns", 0))
    gfov = np.asarray(getattr(rtdose_ds, "GridFrameOffsetVector", None), dtype=float)
    if rows <= 0 or cols <= 0 or gfov.size == 0:
        raise ProfileSkip("Invalid RTDOSE grid geometry metadata")

    arr = np.asarray(rtdose_ds.pixel_array, dtype=np.float64)
    if arr.ndim == 2:
        arr = arr[np.newaxis, :, :]
    if arr.ndim != 3:
        raise ProfileSkip("Unexpected RTDOSE pixel_array dimensionality")
    if arr.shape != (gfov.size, rows, cols):
        raise ProfileSkip("RTDOSE pixel_array shape does not match metadata")

    cube = arr * scale

    # Fixed RTDOSE axis reconstruction in mm.
    x_abs = ipp[0] + spacing[1] * np.arange(cols, dtype=np.float64)
    y_abs = ipp[1] + spacing[0] * np.arange(rows, dtype=np.float64)
    z_abs = ipp[2] + gfov

    x = x_abs - origin_mm[0]
    y = y_abs - origin_mm[1]
    z = z_abs - origin_mm[2]

    z, cube = _ascending(z, cube, 0)
    y, cube = _ascending(y, cube, 1)
    x, cube = _ascending(x, cube, 2)

    interp = RegularGridInterpolator((z, y, x), cube, method="linear", bounds_error=False, fill_value=np.nan)
    return Grid(x, y, z, interp)


def asc_to_query(point: AscPoint) -> Tuple[float, float, float]:
    # ASC axes are beam coordinates for this workflow:
    # crossline -> DICOM X, depth -> DICOM Y, inline -> negative DICOM Z.
    return point.x, point.z, -point.y


def mapping_sanity(points: Sequence[AscPoint], family: str) -> None:
    x0, y0, z0 = asc_to_query(AscPoint(0.0, 0.0, 0.0, 0.0))
    if (x0, y0, z0) != (0.0, 0.0, 0.0):
        raise ProfileSkip("Coordinate-transform bug: ASC(0,0,0) does not map to DICOM-relative(0,0,0)")
    if asc_to_query(AscPoint(0.0, 0.0, 1.0, 0.0))[1] <= 0.0:
        raise ProfileSkip("Coordinate-transform bug: positive ASC depth must map toward positive DICOM Y")
    if asc_to_query(AscPoint(0.0, 1.0, 0.0, 0.0))[2] >= 0.0:
        raise ProfileSkip("Coordinate-transform bug: positive ASC inline must map toward negative DICOM Z")
    if not points:
        raise ProfileSkip("Malformed profile block: no points")

    xs = np.array([p.x for p in points])
    ys = np.array([p.y for p in points])
    zs = np.array([p.z for p in points])
    sx, sy, sz = float(xs.max() - xs.min()), float(ys.max() - ys.min()), float(zs.max() - zs.min())
    tol = 1.0

    if family == "pdd":
        if np.any(np.abs(xs) > tol) or np.any(np.abs(ys) > tol):
            raise ProfileSkip("Coordinate-transform bug: central-axis PDD must keep crossline=0 and inline=0")
        if sz <= tol:
            raise ProfileSkip("Coordinate-transform bug: PDD depth must vary")
    elif family == "inline":
        if sy <= tol or sx > tol or sz > tol:
            raise ProfileSkip("Coordinate-transform bug: inline should vary inline only")
    elif family == "crossline":
        if sx <= tol or sy > tol or sz > tol:
            raise ProfileSkip("Coordinate-transform bug: crossline should vary crossline only")
    elif family == "diagonal":
        if sz > tol:
            raise ProfileSkip("Coordinate-transform bug: diagonal should have near-constant depth")


def sample(grid: Grid, x: float, y: float, z: float) -> Optional[float]:
    if not (grid.x[0] <= x <= grid.x[-1] and grid.y[0] <= y <= grid.y[-1] and grid.z[0] <= z <= grid.z[-1]):
        return None
    val = float(grid.interp(np.array([[z, y, x]], dtype=np.float64))[0])
    if not np.isfinite(val):
        return None
    return val


def diagonal_position(points: Sequence[AscPoint]) -> np.ndarray:
    x = np.array([p.x for p in points], dtype=float)
    y = np.array([p.y for p in points], dtype=float)
    if np.mean(np.abs(x - y)) <= np.mean(np.abs(x + y)):
        return (x + y) / math.sqrt(2.0)
    return (x - y) / math.sqrt(2.0)


def plane_depth(points: Sequence[AscPoint], family: str) -> float:
    d = np.array([p.z for p in points], dtype=float)
    if family in {"inline", "crossline", "diagonal"} and float(d.max() - d.min()) > 1.0:
        raise ProfileSkip(f"Expected near-constant depth for {family} profile")
    return float(np.median(d))


def field_size_mm(flsz: str) -> float:
    m = re.fullmatch(r"(\d+)\*(\d+)", flsz.strip())
    if not m:
        raise ProfileSkip(f"Malformed %FLSZ: '{flsz}'")
    a, b = float(m.group(1)), float(m.group(2))
    if not math.isclose(a, b, rel_tol=0.0, abs_tol=1e-6):
        raise ProfileSkip(f"Only square symmetric fields are supported: '{flsz}'")
    return a


def ssd_mm(ssd: str) -> float:
    if not re.fullmatch(r"\d+", ssd.strip()):
        raise ProfileSkip(f"Malformed %SSD: '{ssd}'")
    return float(ssd)


def is_small_field_profile(profile: Profile, family: str) -> bool:
    if family not in {"inline", "crossline", "diagonal"}:
        return False
    return field_size_mm(profile.meta.flsz) <= SMALL_FIELD_MAX_MM + 1e-6


def _sorted_valid_profile_curves(
    rows,
) -> Optional[Tuple[np.ndarray, np.ndarray, np.ndarray]]:
    """Sort acquired positions for derived metrics, retaining independent dose gaps."""
    vals = []
    for p1d_mm, _depth_mm, meas, rt_norm, _diff, _ok in rows:
        if not np.isfinite(p1d_mm):
            return None
        vals.append((float(p1d_mm), float(meas) if meas is not None else np.nan,
                     float(rt_norm) if rt_norm is not None else np.nan))
    if len(vals) < 2:
        return None
    arr = np.asarray(sorted(vals, key=lambda t: t[0]), dtype=float)
    return arr[:, 0], arr[:, 1], arr[:, 2]


def _prepare_profile_curve(
    positions_mm: np.ndarray, values: np.ndarray,
) -> Optional[Tuple[np.ndarray, np.ndarray]]:
    """Prepare one derived curve without interpolating across invalid dose samples.

    Repeated identical samples are harmless; conflicting doses at the same
    position make interpolation ambiguous and leave this curve unavailable.
    """
    x = np.asarray(positions_mm, dtype=float)
    y = np.asarray(values, dtype=float)
    if x.ndim != 1 or y.ndim != 1 or x.size != y.size or x.size < 2 or not np.all(np.isfinite(x)):
        return None
    order = np.argsort(x, kind="stable")
    x, y = x[order], y[order]
    unique_x, starts = np.unique(x, return_index=True)
    unique_y = []
    for doses in np.split(y, starts[1:]):
        finite_doses = doses[np.isfinite(doses)]
        if finite_doses.size and not np.allclose(finite_doses, finite_doses[0], rtol=1e-9, atol=1e-12):
            return None
        unique_y.append(float(finite_doses[0]) if finite_doses.size else np.nan)
    if unique_x.size < 2:
        return None
    return unique_x, np.asarray(unique_y, dtype=float)


def _linear_crossing(x1: float, y1: float, x2: float, y2: float, level: float) -> Optional[float]:
    if y1 == y2:
        return None
    return float(x1 + (level - y1) * (x2 - x1) / (y2 - y1))


def _find_crossing_from_peak(
    positions_mm: np.ndarray, values: np.ndarray, peak_idx: int, level: float, side: str
) -> Optional[float]:
    direction = -1 if side == "left" else 1
    if side == "left":
        iterator = range(peak_idx - 1, -1, -1)
    else:
        iterator = range(peak_idx + 1, len(values))

    inner = peak_idx
    for outer in iterator:
        x1 = float(positions_mm[inner])
        x2 = float(positions_mm[outer])
        y1 = float(values[inner])
        y2 = float(values[outer])
        if not np.isfinite(y1) or not np.isfinite(y2):
            # A crossing beyond an unknown dose interval cannot be located reliably.
            return None
        if math.isclose(y2, level, rel_tol=0.0, abs_tol=1e-9):
            next_outer = outer + direction
            if 0 <= next_outer < len(values) and math.isclose(
                float(values[next_outer]), level, rel_tol=0.0, abs_tol=1e-9
            ):
                # An entire interval at this level has no unique crossing.
                return None
            return x2
        d1 = y1 - level
        d2 = y2 - level
        if d1 * d2 < 0:
            return _linear_crossing(x1, y1, x2, y2, level)
        inner = outer
    return None


def compute_profile_crossing_positions(positions_mm: np.ndarray, values: np.ndarray) -> Dict[str, Optional[float]]:
    """Find the first outward D80/D50/D20 crossings from the sampled maximum.

    Levels are relative to this profile's own maximum, independently of the
    CAX normalization used for the primary measured-versus-TPS comparison.
    No flat central region is assumed, including for triangular narrow fields.
    Flat threshold intervals and separated equal maxima with a sub-threshold
    valley leave the affected level unavailable rather than selecting one lobe.
    """
    out: Dict[str, Optional[float]] = {
        "curve_max": None,
        "left80": None,
        "left50": None,
        "left20": None,
        "right80": None,
        "right50": None,
        "right20": None,
    }
    prepared = _prepare_profile_curve(positions_mm, values)
    if prepared is None:
        return out
    positions_mm, values = prepared
    finite = np.isfinite(values)
    if not np.any(finite):
        return out
    peak_idx = int(np.argmax(np.where(finite, values, -np.inf)))
    curve_max = float(values[peak_idx])
    if not np.isfinite(curve_max) or curve_max <= 0:
        return out
    out["curve_max"] = curve_max

    levels = {
        "80": 0.8 * curve_max,
        "50": 0.5 * curve_max,
        "20": 0.2 * curve_max,
    }
    maxima = np.flatnonzero(finite & np.isclose(values, curve_max, rtol=1e-9, atol=1e-12))
    between_maxima = values[maxima[0]:maxima[-1] + 1]
    for label, level in levels.items():
        if not np.all(np.isfinite(between_maxima)) or np.any(between_maxima <= level):
            continue
        out[f"left{label}"] = _find_crossing_from_peak(positions_mm, values, peak_idx, level, side="left")
        out[f"right{label}"] = _find_crossing_from_peak(positions_mm, values, peak_idx, level, side="right")
    return out


def _delta(a: Optional[float], b: Optional[float]) -> Optional[float]:
    if a is None or b is None:
        return None
    return float(a - b)


def _width(left: Optional[float], right: Optional[float]) -> Optional[float]:
    if left is None or right is None:
        return None
    return float(right - left)


def _abs_width(a: Optional[float], b: Optional[float]) -> Optional[float]:
    if a is None or b is None:
        return None
    return float(abs(b - a))


def _field_center(crossings: Dict[str, Optional[float]]) -> Optional[float]:
    left, right = crossings["left50"], crossings["right50"]
    if left is None or right is None:
        return None
    return float((left + right) / 2.0)


def compute_small_field_symmetry(positions_mm: np.ndarray, values: np.ndarray, crossings):
    """Compare mirrored doses about this curve's D50 center in percentage points.

    Positive offsets are the unique distances of acquired positions from the
    center, plus the common mirrored D20 boundary. These are the knots and
    endpoints of the piecewise-linear mirrored difference, so no fine grid is
    needed to find its maximum. The mean is an unweighted mean over these
    comparisons; the pair count is not a count of independent acquisitions.
    At least two distinct nonzero valid offsets are required for a metric.
    """
    result = {"max_abs_percent": None, "mean_abs_percent": None, "pair_count": 0}
    center = _field_center(crossings)
    left20, right20 = crossings["left20"], crossings["right20"]
    peak = crossings["curve_max"]
    if center is None or left20 is None or right20 is None or peak is None or peak <= 0:
        return result
    prepared = _prepare_profile_curve(positions_mm, values)
    if prepared is None:
        return result
    x, y = prepared
    extent = min(center - left20, right20 - center)
    if not np.isfinite(extent) or extent <= 1e-9:
        return result
    offsets = sorted([float(abs(pos - center)) for pos in x if 1e-9 < abs(pos - center) < extent] + [extent])
    unique_offsets = []
    for offset in offsets:
        if not unique_offsets or not math.isclose(offset, unique_offsets[-1], rel_tol=0.0, abs_tol=1e-9):
            unique_offsets.append(offset)
    differences = []
    threshold = 0.2 * peak
    dose_epsilon = 1e-9 * max(1.0, peak)
    for offset in unique_offsets:
        # np.interp uses adjacent acquired samples. Retained NaNs prevent gaps
        # from being bridged; no values are extrapolated beyond the D20 domain.
        left = float(np.interp(center - offset, x, y))
        right = float(np.interp(center + offset, x, y))
        if not np.isfinite(left) or not np.isfinite(right):
            continue
        if left < threshold - dose_epsilon or right < threshold - dose_epsilon:
            continue
        differences.append(abs(100.0 * (right - left) / peak))
    result["pair_count"] = len(differences)
    if len(differences) >= 2:
        result["max_abs_percent"] = float(np.max(differences))
        result["mean_abs_percent"] = float(np.mean(differences))
    return result


def compute_small_field_area_symmetry(positions_mm: np.ndarray, values: np.ndarray, crossings) -> Optional[float]:
    """Return signed area symmetry (%) over equal halves about the own D50 center.

    The nearer D20 crossing limits both halves. Every original in-domain knot
    is retained, with only the center and boundaries interpolated as necessary.
    Trapezoids therefore integrate the existing piecewise-linear curve without
    resampling it. Missing support in either half leaves the metric unavailable.
    """
    bounds = [crossings.get(key) for key in ("left20", "left50", "right50", "right20")]
    if any(value is None or not np.isfinite(value) for value in bounds):
        return None
    left20, left50, right50, right20 = bounds
    if not left20 <= left50 < right50 <= right20:
        return None
    center = _field_center(crossings)
    extent = min(center - left20, right20 - center)
    if not np.isfinite(extent) or extent <= 0:
        return None
    prepared = _prepare_profile_curve(positions_mm, values)
    if prepared is None:
        return None
    x, y = prepared
    # Reconstructing a limiting endpoint can overshoot its exact D20 crossing
    # by one ulp; keep the mathematically equal bounds inside those crossings.
    left_bound = max(left20, center - extent)
    right_bound = min(right20, center + extent)
    if left_bound < x[0] or right_bound > x[-1]:
        return None

    areas = []
    for lower, upper in [(left_bound, center), (center, right_bound)]:
        half_x = np.concatenate(([lower], x[(x > lower) & (x < upper)], [upper]))
        if half_x.size < 2 or np.any(np.diff(half_x) <= 0):
            return None
        # Adjacent acquired samples supply the interpolated endpoints. Retained
        # NaNs reject missing knots/brackets instead of bridging unknown doses.
        half_y = np.interp(half_x, x, y)
        if not np.all(np.isfinite(half_y)) or np.any(half_y < 0):
            return None
        area = float(np.sum(np.diff(half_x) * (half_y[:-1] + half_y[1:]) / 2.0))
        if not np.isfinite(area) or area <= 0:
            return None
        areas.append(area)
    area_left, area_right = areas
    total_area = area_right + area_left
    if not np.isfinite(total_area) or total_area <= 0:
        return None
    return float(200.0 * ((area_right - area_left) / total_area))


def compute_small_field_fwhm_and_penumbra_metrics(rows, profile_label: str):
    report = {
        "measured_crossings": {
            "left80": None,
            "left50": None,
            "left20": None,
            "right80": None,
            "right50": None,
            "right20": None,
        },
        "rtdose_crossings": {
            "left80": None,
            "left50": None,
            "left20": None,
            "right80": None,
            "right50": None,
            "right20": None,
        },
        "measured_fwhm_mm": None,
        "rtdose_fwhm_mm": None,
        "delta_fwhm_mm": None,
        "left_50_shift_mm": None,
        "right_50_shift_mm": None,
        "measured_center_mm": None,
        "rtdose_center_mm": None,
        "delta_center_mm": None,
        "measured_left_penumbra_mm": None,
        "rtdose_left_penumbra_mm": None,
        "delta_left_penumbra_mm": None,
        "measured_right_penumbra_mm": None,
        "rtdose_right_penumbra_mm": None,
        "delta_right_penumbra_mm": None,
        "measured_symmetry_max_abs_percent": None,
        "rtdose_symmetry_max_abs_percent": None,
        "delta_symmetry_max_abs_percent": None,
        "measured_symmetry_mean_abs_percent": None,
        "rtdose_symmetry_mean_abs_percent": None,
        "delta_symmetry_mean_abs_percent": None,
        "measured_symmetry_pair_count": 0,
        "rtdose_symmetry_pair_count": 0,
        "measured_symmetry_area_percent": None,
        "rtdose_symmetry_area_percent": None,
        "delta_symmetry_area_percent": None,
    }

    curves = _sorted_valid_profile_curves(rows)
    if curves is None:
        print(
            f"WARN: {profile_label}: small-field geometric metrics unavailable (not enough valid sampled points)",
            file=sys.stderr,
        )
        return report

    positions_mm, measured_vals, rtdose_vals = curves
    measured_cross = compute_profile_crossing_positions(positions_mm, measured_vals)
    rtdose_cross = compute_profile_crossing_positions(positions_mm, rtdose_vals)
    report["measured_crossings"] = {k: measured_cross[k] for k in report["measured_crossings"]}
    report["rtdose_crossings"] = {k: rtdose_cross[k] for k in report["rtdose_crossings"]}

    report["measured_fwhm_mm"] = _width(measured_cross["left50"], measured_cross["right50"])
    report["rtdose_fwhm_mm"] = _width(rtdose_cross["left50"], rtdose_cross["right50"])
    report["delta_fwhm_mm"] = _delta(report["measured_fwhm_mm"], report["rtdose_fwhm_mm"])

    report["left_50_shift_mm"] = _delta(measured_cross["left50"], rtdose_cross["left50"])
    report["right_50_shift_mm"] = _delta(measured_cross["right50"], rtdose_cross["right50"])

    report["measured_center_mm"] = _field_center(measured_cross)
    report["rtdose_center_mm"] = _field_center(rtdose_cross)
    report["delta_center_mm"] = _delta(report["measured_center_mm"], report["rtdose_center_mm"])

    report["measured_left_penumbra_mm"] = _abs_width(measured_cross["left80"], measured_cross["left20"])
    report["rtdose_left_penumbra_mm"] = _abs_width(rtdose_cross["left80"], rtdose_cross["left20"])
    report["delta_left_penumbra_mm"] = _delta(
        report["measured_left_penumbra_mm"], report["rtdose_left_penumbra_mm"]
    )

    report["measured_right_penumbra_mm"] = _abs_width(measured_cross["right80"], measured_cross["right20"])
    report["rtdose_right_penumbra_mm"] = _abs_width(rtdose_cross["right80"], rtdose_cross["right20"])
    report["delta_right_penumbra_mm"] = _delta(
        report["measured_right_penumbra_mm"], report["rtdose_right_penumbra_mm"]
    )

    for prefix, values, crossings in [
        ("measured", measured_vals, measured_cross),
        ("rtdose", rtdose_vals, rtdose_cross),
    ]:
        symmetry = compute_small_field_symmetry(positions_mm, values, crossings)
        report[f"{prefix}_symmetry_max_abs_percent"] = symmetry["max_abs_percent"]
        report[f"{prefix}_symmetry_mean_abs_percent"] = symmetry["mean_abs_percent"]
        report[f"{prefix}_symmetry_pair_count"] = symmetry["pair_count"]
        report[f"{prefix}_symmetry_area_percent"] = compute_small_field_area_symmetry(positions_mm, values, crossings)
    for statistic in ["max_abs", "mean_abs"]:
        report[f"delta_symmetry_{statistic}_percent"] = _delta(
            report[f"measured_symmetry_{statistic}_percent"], report[f"rtdose_symmetry_{statistic}_percent"]
        )
    report["delta_symmetry_area_percent"] = _delta(
        report["measured_symmetry_area_percent"], report["rtdose_symmetry_area_percent"]
    )

    missing = []
    for label, value in [
        ("measured FWHM", report["measured_fwhm_mm"]),
        ("RTDOSE FWHM", report["rtdose_fwhm_mm"]),
        ("left 50% shift", report["left_50_shift_mm"]),
        ("right 50% shift", report["right_50_shift_mm"]),
        ("measured field center", report["measured_center_mm"]),
        ("RTDOSE field center", report["rtdose_center_mm"]),
        ("measured left 80-20 penumbra", report["measured_left_penumbra_mm"]),
        ("RTDOSE left 80-20 penumbra", report["rtdose_left_penumbra_mm"]),
        ("measured right 80-20 penumbra", report["measured_right_penumbra_mm"]),
        ("RTDOSE right 80-20 penumbra", report["rtdose_right_penumbra_mm"]),
        ("measured symmetry", report["measured_symmetry_max_abs_percent"]),
        ("RTDOSE symmetry", report["rtdose_symmetry_max_abs_percent"]),
        ("measured area symmetry", report["measured_symmetry_area_percent"]),
        ("RTDOSE area symmetry", report["rtdose_symmetry_area_percent"]),
    ]:
        if value is None:
            missing.append(label)
    if missing:
        print(
            f"WARN: {profile_label}: small-field metrics partially unavailable: {', '.join(missing)}",
            file=sys.stderr,
        )

    return report


def summarize_small_field_differences(diff_rows):
    """Describe signed relative dose differences at the evaluated ASC positions."""
    signed_vals = np.asarray([float(d) for _p, _m, _r, d in diff_rows], dtype=float)
    signed_vals = signed_vals[np.isfinite(signed_vals)]
    if not signed_vals.size:
        return {
            "count": 0,
            "mean_signed_diff_percent": None,
            "mean_abs_diff_percent": None,
            "max_abs_diff_percent": None,
        }
    abs_vals = np.abs(signed_vals)
    return {
        "count": int(abs_vals.size),
        "mean_signed_diff_percent": float(np.mean(signed_vals)),
        "mean_abs_diff_percent": float(np.mean(abs_vals)),
        "max_abs_diff_percent": float(np.max(abs_vals)),
    }


def _mm_to_cm(value: Optional[float]) -> Optional[float]:
    if value is None:
        return None
    return float(value / 10.0)


def append_small_field_report_to_worksheet(ws, report, difference_summary) -> None:
    start_col = 21  # U
    row = 1
    ws.cell(row=row, column=start_col, value="SMALL-FIELD SUMMARY")
    row += 1

    summary_rows = [
        ("measured FWHM (cm)", _mm_to_cm(report["measured_fwhm_mm"])),
        ("RTDOSE FWHM (cm)", _mm_to_cm(report["rtdose_fwhm_mm"])),
        ("delta FWHM (cm)", _mm_to_cm(report["delta_fwhm_mm"])),
        ("left 50% edge shift (cm)", _mm_to_cm(report["left_50_shift_mm"])),
        ("right 50% edge shift (cm)", _mm_to_cm(report["right_50_shift_mm"])),
        ("measured field center (cm)", _mm_to_cm(report["measured_center_mm"])),
        ("RTDOSE field center (cm)", _mm_to_cm(report["rtdose_center_mm"])),
        ("delta field center (cm)", _mm_to_cm(report["delta_center_mm"])),
        ("measured left 80-20 penumbra (cm)", _mm_to_cm(report["measured_left_penumbra_mm"])),
        ("RTDOSE left 80-20 penumbra (cm)", _mm_to_cm(report["rtdose_left_penumbra_mm"])),
        ("delta left penumbra (cm)", _mm_to_cm(report["delta_left_penumbra_mm"])),
        ("measured right 80-20 penumbra (cm)", _mm_to_cm(report["measured_right_penumbra_mm"])),
        ("RTDOSE right 80-20 penumbra (cm)", _mm_to_cm(report["rtdose_right_penumbra_mm"])),
        ("delta right penumbra (cm)", _mm_to_cm(report["delta_right_penumbra_mm"])),
        ("measured Symmetry (point), max abs (pp)", report["measured_symmetry_max_abs_percent"]),
        ("RTDOSE Symmetry (point), max abs (pp)", report["rtdose_symmetry_max_abs_percent"]),
        ("delta Symmetry (point), max abs (pp)", report["delta_symmetry_max_abs_percent"]),
        ("measured Symmetry (point), mean abs (pp)", report["measured_symmetry_mean_abs_percent"]),
        ("RTDOSE Symmetry (point), mean abs (pp)", report["rtdose_symmetry_mean_abs_percent"]),
        ("delta Symmetry (point), mean abs (pp)", report["delta_symmetry_mean_abs_percent"]),
        ("measured Symmetry (point) pair count", report["measured_symmetry_pair_count"]),
        ("RTDOSE Symmetry (point) pair count", report["rtdose_symmetry_pair_count"]),
        ("measured Symmetry (area) (%)", report["measured_symmetry_area_percent"]),
        ("RTDOSE Symmetry (area) (%)", report["rtdose_symmetry_area_percent"]),
        ("delta Symmetry (area) (pp)", report["delta_symmetry_area_percent"]),
        ("point count (RTDOSE >= 20%)", difference_summary["count"]),
        ("mean signed difference (%) (RTDOSE >= 20%)", difference_summary["mean_signed_diff_percent"]),
        ("mean abs difference (%) (RTDOSE >= 20%)", difference_summary["mean_abs_diff_percent"]),
        ("max abs difference (%) (RTDOSE >= 20%)", difference_summary["max_abs_diff_percent"]),
    ]
    for label, value in summary_rows:
        ws.cell(row=row, column=start_col, value=label)
        ws.cell(row=row, column=start_col + 1, value=value if value is not None else "N/A")
        row += 1

    pen_row = row + 1
    ws.cell(row=pen_row, column=start_col, value="PENUMBRA TABLE")
    ws.cell(row=pen_row + 1, column=start_col, value="metric label")
    ws.cell(row=pen_row + 1, column=start_col + 1, value="measured position (cm)")
    ws.cell(row=pen_row + 1, column=start_col + 2, value="RTDOSE position (cm)")
    ws.cell(row=pen_row + 1, column=start_col + 3, value="shift (cm)")

    pen_rows = [
        ("left 80%", "left80"),
        ("left 50%", "left50"),
        ("left 20%", "left20"),
        ("right 80%", "right80"),
        ("right 50%", "right50"),
        ("right 20%", "right20"),
    ]
    r = pen_row + 2
    for label, key in pen_rows:
        m = report["measured_crossings"].get(key)
        d = report["rtdose_crossings"].get(key)
        ws.cell(row=r, column=start_col, value=label)
        for col_offset, value in enumerate((_mm_to_cm(m), _mm_to_cm(d), _mm_to_cm(_delta(m, d))), start=1):
            ws.cell(row=r, column=start_col + col_offset, value=value if value is not None else "N/A")
        r += 1

    notes = [
        ("missing metrics", "N/A = unavailable; no acceptance criteria are applied."),
        ("dose comparison", "Original ASC positions and measured doses; RTDOSE normalized to TPS CAX = 100%."),
        ("difference (%)", "100 * (measured - RTDOSE normalized) / RTDOSE normalized; TPS >= 20%."),
        ("geometry", "D20/D50/D80 relative to each profile maximum; center = (left D50 + right D50) / 2."),
        ("delta convention", "All deltas and shifts are measured minus RTDOSE."),
        ("Symmetry (point) (pp)", "100 * (right dose - left dose) / own maximum; pairs about own D50 center, both >= 20%."),
        ("Symmetry (point) pairs", "Distinct nonzero original-sample offsets plus the shared mirrored D20 boundary; linear interpolation within D20 bounds."),
        ("Symmetry (point) availability", "At least two valid mirrored offsets; missing D20/D50 crossings give N/A."),
        ("Symmetry (point) mean/count", "Unweighted comparisons at sample offsets and the D20 boundary; pair count is not an independent acquisition count."),
        ("Symmetry (area) (%)", "200 * (right area - left area) / (right area + left area); positive = larger right area."),
        ("Symmetry (area) extent", "Equal distances about own D50 center, limited by the nearer D20 crossing."),
        ("Symmetry (area) integration", "Trapezoids at original sample positions plus interpolated center/boundaries; missing required doses or crossings give N/A."),
    ]
    for label, note in notes:
        r += 1
        ws.cell(row=r, column=start_col, value=label)
        ws.cell(row=r, column=start_col + 1, value=note)


def interpolate_pdd_value_at_depth(
    depths_mm: Sequence[float], values: Sequence[float], target_depth_mm: float
) -> Optional[float]:
    pairs: Dict[float, List[float]] = defaultdict(list)
    for d, v in zip(depths_mm, values):
        if not np.isfinite(d) or not np.isfinite(v):
            continue
        pairs[float(round(float(d), 6))].append(float(v))

    if not pairs:
        return None

    xs = sorted(pairs.keys())
    ys = [float(np.mean(pairs[x])) for x in xs]
    if len(xs) == 1:
        if math.isclose(target_depth_mm, xs[0], abs_tol=1e-6):
            return ys[0]
        return None

    if target_depth_mm < xs[0] - 1e-6 or target_depth_mm > xs[-1] + 1e-6:
        return None

    for x, y in zip(xs, ys):
        if math.isclose(target_depth_mm, x, abs_tol=1e-6):
            return y

    idx = int(np.searchsorted(np.asarray(xs, dtype=float), target_depth_mm, side="left"))
    if idx <= 0 or idx >= len(xs):
        return None
    x0, x1 = xs[idx - 1], xs[idx]
    y0, y1 = ys[idx - 1], ys[idx]
    if math.isclose(x0, x1, abs_tol=1e-12):
        return None
    return float(y0 + (target_depth_mm - x0) * (y1 - y0) / (x1 - x0))


def compute_tpr_from_pdd_pair(pdd10: Optional[float], pdd20: Optional[float]) -> Optional[float]:
    if pdd10 is None or pdd20 is None:
        return None
    if math.isclose(pdd10, 0.0, abs_tol=1e-12):
        return None
    return float(1.2661 * pdd20 / pdd10 - 0.0595)


def compute_ratio_diff(measured_value: Optional[float], tps_value: Optional[float]) -> Optional[float]:
    if measured_value is None or tps_value is None:
        return None
    if math.isclose(tps_value, 0.0, abs_tol=1e-12):
        return None
    return float(100.0 * (measured_value - tps_value) / tps_value)


def compute_pdd_tpr_summary(rows, profile_label: str):
    measured_depths = []
    measured_vals = []
    rtdose_depths = []
    rtdose_vals = []

    for _p1d_mm, depth_mm_value, meas, rt_norm, _diff, _ok in rows:
        if np.isfinite(depth_mm_value) and np.isfinite(meas):
            measured_depths.append(float(depth_mm_value))
            measured_vals.append(float(meas))
        if rt_norm is not None and np.isfinite(depth_mm_value) and np.isfinite(rt_norm):
            rtdose_depths.append(float(depth_mm_value))
            rtdose_vals.append(float(rt_norm))

    summary = {
        "measured_pdd10": interpolate_pdd_value_at_depth(measured_depths, measured_vals, 100.0),
        "measured_pdd20": interpolate_pdd_value_at_depth(measured_depths, measured_vals, 200.0),
        "rtdose_pdd10": interpolate_pdd_value_at_depth(rtdose_depths, rtdose_vals, 100.0),
        "rtdose_pdd20": interpolate_pdd_value_at_depth(rtdose_depths, rtdose_vals, 200.0),
    }
    summary["tpr_measured"] = compute_tpr_from_pdd_pair(summary["measured_pdd10"], summary["measured_pdd20"])
    summary["tpr_rtdose"] = compute_tpr_from_pdd_pair(summary["rtdose_pdd10"], summary["rtdose_pdd20"])
    summary["diff_tpr_ratio"] = compute_ratio_diff(summary["tpr_measured"], summary["tpr_rtdose"])

    missing = []
    for label, value in [
        ("measured PDD(10)", summary["measured_pdd10"]),
        ("measured PDD(20)", summary["measured_pdd20"]),
        ("RTDOSE PDD(10)", summary["rtdose_pdd10"]),
        ("RTDOSE PDD(20)", summary["rtdose_pdd20"]),
        ("TPR_measured", summary["tpr_measured"]),
        ("TPR_rtdose", summary["tpr_rtdose"]),
    ]:
        if value is None:
            missing.append(label)
    if missing:
        print(
            f"WARN: {profile_label}: PDD/TPR summary partially unavailable: {', '.join(missing)}",
            file=sys.stderr,
        )
    return summary


def append_pdd_tpr_summary_to_worksheet(ws, summary) -> None:
    start_col = 18  # R
    row = 1
    ws.cell(row=row, column=start_col, value="PDD/TPR SUMMARY")
    row += 1
    rows = [
        ("measured PDD(10)", summary["measured_pdd10"]),
        ("measured PDD(20)", summary["measured_pdd20"]),
        ("RTDOSE PDD(10)", summary["rtdose_pdd10"]),
        ("RTDOSE PDD(20)", summary["rtdose_pdd20"]),
        ("TPR_measured", summary["tpr_measured"]),
        ("TPR_rtdose", summary["tpr_rtdose"]),
        ("diff_TPR (%)", summary["diff_tpr_ratio"]),
    ]
    for label, value in rows:
        ws.cell(row=row, column=start_col, value=label)
        ws.cell(row=row, column=start_col + 1, value=value)
        row += 1


def in_central(family: str, pos_mm: float, depth_mm: float, field_mm: float, ssd_value_mm: float) -> bool:
    width = field_mm * (ssd_value_mm + depth_mm) / SAD_MM
    if family in {"inline", "crossline"}:
        cb = 0.8 * width
    elif family == "diagonal":
        cb = 0.8 * math.sqrt(2.0) * width
    else:
        return True
    return abs(pos_mm) <= 0.5 * cb


def process_profile(profile: Profile, family: str, grid: Grid):
    mapping_sanity(profile.points, family)
    if family == "inline":
        pos = np.array([p.y for p in profile.points], dtype=float)
    elif family == "crossline":
        pos = np.array([p.x for p in profile.points], dtype=float)
    elif family == "pdd":
        pos = np.array([p.z for p in profile.points], dtype=float)
    else:
        pos = diagonal_position(profile.points)

    dep = plane_depth(profile.points, family)
    fmm = field_size_mm(profile.meta.flsz)
    ssdv = ssd_mm(profile.meta.ssd)
    small_field = is_small_field_profile(profile, family)

    sampled: List[Optional[float]] = []
    for p in profile.points:
        xq, yq, zq = asc_to_query(p)
        s = sample(grid, xq, yq, zq)
        sampled.append(s)
        if s is None:
            print(
                f"WARN: Out-of-bounds point skipped: x={p.x:.3f} mm, y={p.y:.3f} mm, depth={p.z:.3f} mm"
            )

    if family in {"inline", "crossline", "diagonal"}:
        cax_x, cax_y, cax_z = asc_to_query(AscPoint(0.0, 0.0, dep, 0.0))
        cax = sample(grid, cax_x, cax_y, cax_z)
        if cax is None or cax <= 0:
            raise ProfileSkip("Normalization failure: invalid/non-positive CAX RTDOSE value")
        # Dose comparison uses the TPS CAX, not each profile's peak used for geometry.
        norm = [(100.0 * s / cax) if s is not None else None for s in sampled]
    else:
        valid = [s for s in sampled if s is not None and s > 0]
        if not valid:
            raise ProfileSkip("Normalization failure: no valid RTDOSE samples for PDD")
        dmax = max(valid)
        if dmax <= 0:
            raise ProfileSkip("Normalization failure: invalid/non-positive PDD Dmax")
        norm = [(100.0 * s / dmax) if s is not None else None for s in sampled]

    rows = []
    diff_rows = []
    for p, p1d, rn in zip(profile.points, pos, norm):
        diff = None
        valid_diff = False
        if rn is not None and rn > 0:
            if family == "pdd":
                valid_diff = True
            elif small_field:
                valid_diff = (
                    np.isfinite(p1d) and np.isfinite(p.dose) and np.isfinite(rn)
                    and rn >= SMALL_FIELD_DOSE_THRESHOLD_PERCENT
                )
            else:
                valid_diff = in_central(family, float(p1d), dep, fmm, ssdv)
            if valid_diff:
                diff = 100.0 * (p.dose - rn) / rn
                diff_rows.append((float(p1d), p.dose, rn, diff))
        rows.append((float(p1d), p.z, p.dose, rn, diff, valid_diff))
    if small_field and family in {"inline", "crossline", "diagonal"} and not diff_rows:
        print(
            f"WARN: {profile.asc_path.name} profile #{profile.index + 1}: "
            f"no valid dose-comparison points found (requires RTDOSE normalized >= {SMALL_FIELD_DOSE_THRESHOLD_PERCENT:.0f}%)",
            file=sys.stderr,
        )
    return dep, rows, diff_rows, small_field


def _side_diffs(diff_rows, side: str) -> List[float]:
    out = []
    for pos_mm, _meas, _rt, diff in diff_rows:
        p = float(pos_mm)
        if side == "left" and p < 0.0:
            out.append(float(diff))
        elif side == "right" and p > 0.0:
            out.append(float(diff))
    return out


def _max_abs_signed(values: Sequence[float]) -> Optional[float]:
    if not values:
        return None
    return float(max(values, key=lambda v: abs(v)))


def _mean_signed(values: Sequence[float]) -> Optional[float]:
    if not values:
        return None
    return float(np.mean(np.asarray(values, dtype=float)))


def build_conventional_summary_record(
    profile_name: str,
    family: str,
    depth_mm: float,
    field_mm: float,
    diff_rows,
) -> SummaryRecord:
    all_diffs = [float(d) for _pos, _m, _r, d in diff_rows]
    total = len(all_diffs)

    n_1_2 = 0
    n_gt2 = 0
    for d in all_diffs:
        a = abs(d)
        if a >= 1.0 and a <= 2.0:
            n_1_2 += 1
        elif a > 2.0:
            n_gt2 += 1

    pct_1_2 = (100.0 * n_1_2 / total) if total > 0 else None
    pct_gt2 = (100.0 * n_gt2 / total) if total > 0 else None

    left_diffs = _side_diffs(diff_rows, side="left")
    right_diffs = _side_diffs(diff_rows, side="right")

    return SummaryRecord(
        profile_name=profile_name,
        family=family,
        depth_mm=float(depth_mm),
        field_mm=float(field_mm),
        max_diff_left=_max_abs_signed(left_diffs),
        max_diff_right=_max_abs_signed(right_diffs),
        n_points_1_2=n_1_2,
        n_points_gt2=n_gt2,
        pct_points_1_2=pct_1_2,
        pct_points_gt2=pct_gt2,
        avg_diff_left=_mean_signed(left_diffs),
        avg_diff_right=_mean_signed(right_diffs),
    )


def _summary_headers() -> List[str]:
    return [
        "Profile",
        "Max diff L",
        "Max diff R",
        "N points 1-2%",
        "N points >2%",
        "% of points 1-2%",
        "% of points >2%",
        "Average diff L",
        "Average diff R",
    ]


def _write_summary_headers(ws, row: int) -> int:
    for col, header in enumerate(_summary_headers(), start=1):
        ws.cell(row=row, column=col, value=header)
    return row + 1


def _write_summary_record_row(ws, row: int, record: SummaryRecord) -> int:
    ws.cell(row=row, column=1, value=record.profile_name)
    ws.cell(row=row, column=2, value=record.max_diff_left)
    ws.cell(row=row, column=3, value=record.max_diff_right)
    ws.cell(row=row, column=4, value=record.n_points_1_2)
    ws.cell(row=row, column=5, value=record.n_points_gt2)
    ws.cell(row=row, column=6, value=record.pct_points_1_2)
    ws.cell(row=row, column=7, value=record.pct_points_gt2)
    ws.cell(row=row, column=8, value=record.avg_diff_left)
    ws.cell(row=row, column=9, value=record.avg_diff_right)
    return row + 1


def _ordered_profile_families() -> List[str]:
    return [k for k, _v in sorted(PROFILE_TYPE_ORDER.items(), key=lambda kv: kv[1])]


def _write_summary_section_type_depth_field(ws, row: int, records: Sequence[SummaryRecord]) -> int:
    wrote_any_family = False
    for family in _ordered_profile_families():
        fam = [r for r in records if r.family == family]
        if not fam:
            continue
        if wrote_any_family:
            row += 1
        fam_sorted = sorted(fam, key=lambda r: (r.depth_mm, r.field_mm, r.profile_name.lower()))
        prev_depth: Optional[float] = None
        for rec in fam_sorted:
            if prev_depth is not None and not math.isclose(rec.depth_mm, prev_depth, abs_tol=1e-6):
                row += 1
            row = _write_summary_record_row(ws, row, rec)
            prev_depth = rec.depth_mm
        wrote_any_family = True
    return row


def _write_summary_section_type_field_depth(ws, row: int, records: Sequence[SummaryRecord]) -> int:
    wrote_any_family = False
    for family in _ordered_profile_families():
        fam = [r for r in records if r.family == family]
        if not fam:
            continue
        if wrote_any_family:
            row += 1
        fam_sorted = sorted(fam, key=lambda r: (r.field_mm, r.depth_mm, r.profile_name.lower()))
        prev_field: Optional[float] = None
        for rec in fam_sorted:
            if prev_field is not None and not math.isclose(rec.field_mm, prev_field, abs_tol=1e-6):
                row += 1
            row = _write_summary_record_row(ws, row, rec)
            prev_field = rec.field_mm
        wrote_any_family = True
    return row


def write_summary_sheet(wb: Workbook, records: Sequence[SummaryRecord]) -> None:
    if not records:
        return

    ws = wb.create_sheet(title="Summary")
    row = 1
    row = _write_summary_headers(ws, row)
    row = _write_summary_section_type_depth_field(ws, row, records)
    row += 2
    row = _write_summary_headers(ws, row)
    _write_summary_section_type_field_depth(ws, row, records)


def unique_sheet_name(wb: Workbook, base: str) -> str:
    names = {n.lower() for n in wb.sheetnames}
    name, i = base, 2
    while name.lower() in names:
        name = f"{base}_{i}"
        i += 1
    return name


def unique_workbook_path(out_dir: Path, base_name: str) -> Path:
    candidate = out_dir / base_name
    if not candidate.exists():
        return candidate
    stem = Path(base_name).stem
    suffix = Path(base_name).suffix
    i = 1
    while True:
        candidate = out_dir / f"{stem}_{i}{suffix}"
        if not candidate.exists():
            return candidate
        i += 1


def depth_token(mm: float) -> str:
    if math.isclose(mm, round(mm), abs_tol=1e-6):
        return str(int(round(mm)))
    return f"{mm:.1f}".rstrip("0").rstrip(".")


def sheet_base(meta: Meta, family: str, depth_mm_value: float) -> str:
    fl = meta.flsz.replace("*", "x") if meta.flsz else "unknown"
    ssd = meta.ssd if meta.ssd else "unknown"
    if family == "pdd":
        return f"{fl}_{ssd}_pdd"
    return f"{fl}_{ssd}_{depth_token(depth_mm_value)}_{family}"


def write_meta(ws, m: Meta) -> int:
    items = [
        ("machine / energy", m.machine),
        ("field", m.field_comment),
        ("Scan ID", m.scan_id),
        ("%DATE", m.date),
        ("%TYPE", m.typ),
        ("%AXIS", m.axis),
        ("%PNTS", m.pnts),
        ("%STEP", m.step),
        ("%SSD", m.ssd),
        ("%FLSZ", m.flsz),
        ("%DPTH", m.dpth),
    ]
    r = 1
    for k, v in items:
        ws.cell(row=r, column=1, value=k)
        ws.cell(row=r, column=2, value=v)
        r += 1
    return r + 1


def color_diff_profile(cell, diff: float) -> None:
    a = abs(diff)
    if 1 < a <= 2:
        cell.fill = YELLOW_FILL
    elif a > 2:
        cell.fill = RED_FILL


def color_diff_pdd(cell, diff: float) -> None:
    # PDD uses a single 1% action limit (no intermediate yellow band).
    if abs(diff) > 1:
        cell.fill = RED_FILL


def style_primary_axes(chart) -> None:
    # Ensure visible primary horizontal/vertical axes.
    chart.x_axis.delete = False
    chart.y_axis.delete = False
    chart.x_axis.axPos = "b"
    chart.y_axis.axPos = "l"
    chart.x_axis.majorTickMark = "out"
    chart.y_axis.majorTickMark = "out"
    chart.x_axis.graphicalProperties = GraphicalProperties()
    chart.y_axis.graphicalProperties = GraphicalProperties()
    if chart.x_axis.graphicalProperties.line is None:
        chart.x_axis.graphicalProperties.line = LineProperties()
    if chart.y_axis.graphicalProperties.line is None:
        chart.y_axis.graphicalProperties.line = LineProperties()
    chart.x_axis.graphicalProperties.line.solidFill = "000000"
    chart.y_axis.graphicalProperties.line.solidFill = "000000"
    chart.x_axis.graphicalProperties.line.width = LINE_WIDTH_1PT
    chart.y_axis.graphicalProperties.line.width = LINE_WIDTH_1PT


def make_overlay(ws, min_row: int, max_row: int, x_col: int, y1_col: int, y2_col: int, title: str, x_title: str, anchor: str):
    c = LineChart()
    c.title = title
    c.x_axis.title = x_title
    c.y_axis.title = "Dose (%)"
    style_primary_axes(c)
    c.x_axis.crosses = "min"
    c.y_axis.crosses = "min"
    c.x_axis.crossAx = c.y_axis.axId
    c.y_axis.crossAx = c.x_axis.axId
    c.height = 7
    c.width = 12
    data = Reference(ws, min_col=y1_col, max_col=y2_col, min_row=min_row, max_row=max_row)
    cats = Reference(ws, min_col=x_col, min_row=min_row + 1, max_row=max_row)
    c.add_data(data, titles_from_data=True)
    c.set_categories(cats)
    c.x_axis.majorGridlines = None
    c.y_axis.majorGridlines = None
    c.x_axis.minorGridlines = None
    c.y_axis.minorGridlines = None
    for s in c.series:
        s.graphicalProperties.line.width = LINE_WIDTH_1PT
    ws.add_chart(c, anchor)


def make_diff_chart(ws, min_row: int, max_row: int, x_col: int, y_col: int, x_title: str, anchor: str):
    c = ScatterChart()
    c.title = "Difference"
    c.x_axis.title = x_title
    c.y_axis.title = "Difference (%)"
    style_primary_axes(c)
    c.x_axis.crossAx = c.y_axis.axId
    c.y_axis.crossAx = c.x_axis.axId
    c.height = 7
    c.width = 12
    xref = Reference(ws, min_col=x_col, min_row=min_row + 1, max_row=max_row)
    yref = Reference(ws, min_col=y_col, min_row=min_row, max_row=max_row)
    s = Series(yref, xref, title_from_data=True)
    s.marker.symbol = "circle"
    s.marker.size = 5
    s.smooth = False
    s.graphicalProperties.line.width = LINE_WIDTH_1PT
    c.series.append(s)
    c.legend = None
    c.x_axis.majorGridlines = None
    c.y_axis.majorGridlines = None
    c.x_axis.minorGridlines = None
    c.y_axis.minorGridlines = None

    x_vals = []
    y_vals = []
    for r in range(min_row + 1, max_row + 1):
        xv = ws.cell(row=r, column=x_col).value
        yv = ws.cell(row=r, column=y_col).value
        if isinstance(xv, (int, float)):
            x_vals.append(float(xv))
        if isinstance(yv, (int, float)):
            y_vals.append(float(yv))

    if x_vals:
        xmin, xmax = min(x_vals), max(x_vals)
        if xmin < 0.0 < xmax:
            x_abs = max(abs(xmin), abs(xmax))
            c.x_axis.scaling.min = -x_abs
            c.x_axis.scaling.max = x_abs
        x_span = max(xmax - xmin, 0.0)
        if x_span <= 2.0:
            c.x_axis.majorUnit = 0.2
        elif x_span <= 10.0:
            c.x_axis.majorUnit = 1.0
        elif x_span <= 40.0:
            c.x_axis.majorUnit = 5.0
        else:
            c.x_axis.majorUnit = 10.0

    if y_vals:
        y_abs = max(max(abs(v) for v in y_vals), 0.2)
        y_lim = math.ceil((1.1 * y_abs) / 0.2) * 0.2
        c.y_axis.scaling.min = -y_lim
        c.y_axis.scaling.max = y_lim
        if y_lim <= 0.6:
            c.y_axis.majorUnit = 0.1
        elif y_lim <= 2.0:
            c.y_axis.majorUnit = 0.2
        elif y_lim <= 5.0:
            c.y_axis.majorUnit = 0.5
        else:
            c.y_axis.majorUnit = 1.0

    ws.add_chart(c, anchor)


def write_skipped_sheet(wb: Workbook, profile: Profile, family_hint: str, depth_hint: float, reason: str) -> None:
    ws = wb.create_sheet(title=unique_sheet_name(wb, sheet_base(profile.meta, family_hint, depth_hint)))
    r = write_meta(ws, profile.meta)
    ws.cell(row=r, column=1, value="STATUS")
    ws.cell(row=r, column=2, value="SKIPPED")
    ws.cell(row=r + 1, column=1, value="REASON")
    ws.cell(row=r + 1, column=2, value=reason)


def write_profile_sheet(
    wb: Workbook, profile: Profile, family: str, depth_mm_value: float, rows, diff_rows, small_field: bool
) -> str:
    ws = wb.create_sheet(title=unique_sheet_name(wb, sheet_base(profile.meta, family, depth_mm_value)))
    start = write_meta(ws, profile.meta)

    if family in {"inline", "crossline", "diagonal"}:
        ws.cell(row=start, column=1, value="position (cm)")
        ws.cell(row=start, column=2, value="measured dose")
        ws.cell(row=start, column=3, value="RTDOSE dose (% normalized)")
        r = start + 1
        for p1d_mm, _depth_mm, meas, rt_norm, _diff, _ok in rows:
            ws.cell(row=r, column=1, value=p1d_mm / 10.0)
            ws.cell(row=r, column=2, value=meas)
            ws.cell(row=r, column=3, value=rt_norm)
            r += 1

        if small_field:
            ws.cell(row=start, column=6, value="position (cm)")
            ws.cell(row=start, column=7, value="measured dose")
            ws.cell(row=start, column=8, value="RTDOSE dose (% normalized)")
            ws.cell(row=start, column=9, value="difference (%)")
            ws.cell(row=start, column=10, value="abs difference (%)")
        else:
            ws.cell(row=start, column=6, value="position (cm)")
            ws.cell(row=start, column=7, value="measured dose")
            ws.cell(row=start, column=8, value="RTDOSE dose (% normalized)")
            ws.cell(row=start, column=9, value="difference (%)")
        rd = start + 1
        for p1d_mm, meas, rt_norm, d in diff_rows:
            ws.cell(row=rd, column=6, value=p1d_mm / 10.0)
            ws.cell(row=rd, column=7, value=meas)
            ws.cell(row=rd, column=8, value=rt_norm)
            c = ws.cell(row=rd, column=9, value=d)
            if small_field:
                ws.cell(row=rd, column=10, value=abs(d))
            else:
                color_diff_profile(c, d)
            rd += 1

        if small_field:
            ws.cell(row=start - 1, column=6, value="DOSE-DIFFERENCE TABLE (RTDOSE >= 20%)")
            sf_metrics = compute_small_field_fwhm_and_penumbra_metrics(
                rows, f"{profile.asc_path.name} profile #{profile.index + 1}"
            )
            difference_summary = summarize_small_field_differences(diff_rows)
            append_small_field_report_to_worksheet(ws, sf_metrics, difference_summary)

        if r - 1 >= start + 1:
            make_overlay(ws, start, r - 1, 1, 2, 3, "Measured vs RTDOSE", "Distance (cm)", "I1")
        if rd - 1 >= start + 1:
            make_diff_chart(ws, start, rd - 1, 6, 9, "Distance (cm)", "I20")

    else:
        ws.cell(row=start, column=1, value="depth (cm)")
        ws.cell(row=start, column=2, value="measured dose")
        ws.cell(row=start, column=3, value="RTDOSE dose (% normalized)")
        ws.cell(row=start, column=4, value="difference (%)")
        r = start + 1
        for _p1d_mm, depth_mm_value_row, meas, rt_norm, d, _ok in rows:
            ws.cell(row=r, column=1, value=depth_mm_value_row / 10.0)
            ws.cell(row=r, column=2, value=meas)
            ws.cell(row=r, column=3, value=rt_norm)
            c = ws.cell(row=r, column=4, value=d)
            if d is not None:
                color_diff_pdd(c, d)
            r += 1
        pdd_summary = compute_pdd_tpr_summary(rows, f"{profile.asc_path.name} profile #{profile.index + 1}")
        append_pdd_tpr_summary_to_worksheet(ws, pdd_summary)
        if r - 1 >= start + 1:
            make_overlay(ws, start, r - 1, 1, 2, 3, "Measured vs RTDOSE", "Depth (cm)", "F1")
            make_diff_chart(ws, start, r - 1, 1, 4, "Depth (cm)", "F20")
    return ws.title


def process_asc(asc_path: Path, di: DicomIndex, out_dir: Path) -> None:
    profiles = parse_asc(asc_path)
    wb = Workbook()
    wb.remove(wb.active)
    summary_records: List[SummaryRecord] = []

    for pf in profiles:
        fam_hint, dep_hint = "profile", 0.0
        try:
            if pf.parse_error:
                raise ProfileSkip(pf.parse_error)
            family = classify(pf)
            fam_hint = family
            dep_hint = plane_depth(pf.points, family) if pf.points else 0.0

            expected_plan_label = parse_expected_plan_label_from_asc(pf)
            expected_energy, expected_mode = parse_expected_energy_mode_from_asc_comment(pf.meta.machine)
            expected_key = (normalize_plan_label(expected_plan_label), expected_energy, expected_mode)

            candidate = di.candidate_by_key.get(expected_key)
            if candidate is None:
                raise ProfileSkip(
                    "No RTDOSE/RTPLAN candidate matched ASC metadata key "
                    f"(plan_label='{expected_plan_label}', energy={expected_energy:g}, mode={expected_mode})"
                )

            rtdose_ds = pydicom.dcmread(candidate.rtdose_path, force=True)
            rtplan_ds = pydicom.dcmread(candidate.rtplan_path, stop_before_pixels=True, force=True)

            origin = origin_from_rtplan(rtplan_ds)
            grid = build_grid(rtdose_ds, origin)

            dep, rows, diff_rows, small_field = process_profile(pf, family, grid)
            sheet_name = write_profile_sheet(wb, pf, family, dep, rows, diff_rows, small_field)
            if family in {"inline", "crossline", "diagonal"} and not small_field:
                summary_records.append(
                    build_conventional_summary_record(
                        profile_name=sheet_name,
                        family=family,
                        depth_mm=dep,
                        field_mm=field_size_mm(pf.meta.flsz),
                        diff_rows=diff_rows,
                    )
                )

        except ProfileSkip as exc:
            print(f"WARN: {asc_path.name} profile #{pf.index + 1} skipped: {exc}", file=sys.stderr)
            write_skipped_sheet(wb, pf, fam_hint, dep_hint, str(exc))
        except Exception as exc:
            reason = f"Unexpected profile processing error: {exc}"
            print(f"WARN: {asc_path.name} profile #{pf.index + 1} skipped: {reason}", file=sys.stderr)
            write_skipped_sheet(wb, pf, fam_hint, dep_hint, reason)

    if not wb.sheetnames:
        ws = wb.create_sheet("no_profiles")
        ws["A1"] = "No profile blocks parsed"
    elif summary_records:
        write_summary_sheet(wb, summary_records)

    out_path = unique_workbook_path(out_dir, f"{asc_path.stem}_comparison.xlsx")
    wb.save(out_path)
    print(f"INFO: Wrote workbook: {out_path}")


def run(root: Path) -> int:
    root = root.resolve()
    if not root.exists() or not root.is_dir():
        raise StartupError(f"Target folder does not exist or is not a directory: '{root}'")

    asc_files, dcm_files = discover(root)
    if not asc_files:
        raise StartupError(f"No ASC files found under '{root}'")

    di = index_dicoms(dcm_files)
    for asc in asc_files:
        process_asc(asc, di, root)
    print("INFO: Processing complete")
    return 0


def parse_args(argv: Optional[Sequence[str]] = None):
    ap = argparse.ArgumentParser(
        description=(
            "Compare ASC profiles vs RTDOSE using RTPLAN-linked metadata matching and RTPLAN ORIGIN "
            "for coordinate navigation"
        )
    )
    ap.add_argument(
        "root",
        nargs="?",
        default=".",
        help="Target folder to scan recursively; output workbooks are saved in this same folder",
    )
    return ap.parse_args(argv)


def main(argv: Optional[Sequence[str]] = None) -> int:
    args = parse_args(argv)
    try:
        return run(Path(args.root))
    except StartupError as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 2
    except Exception as exc:
        print(f"ERROR: Unexpected fatal error: {exc}", file=sys.stderr)
        return 3


if __name__ == "__main__":
    raise SystemExit(main())
