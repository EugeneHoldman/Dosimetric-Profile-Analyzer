# Water-phantom dose comparison

[DPC small field new.py](<DPC small field new.py>) compares measured water-phantom scans in ASC files with Eclipse-exported DICOM RTDOSE volumes and writes Excel reports.

The central idea was to support selected dosimetric QA requirements of **[AAPM Medical Physics Practice Guideline 8.b: Linear accelerator performance tests](https://doi.org/10.1002/acm2.14160)** (Krauss et al., 2023), particularly beam-profile and depth-dose comparisons with TPS reference data.

## Run

Download `DPC.exe` from this GitHub repository. The executable runs on Windows without a separate Python or library installation.

To run it directly, copy `DPC.exe` into the folder containing your ASC profiles and DICOM files, then double-click it. The program automatically searches that folder and its subfolders and saves the Excel reports alongside the executable.

Alternatively, open PowerShell in the folder containing `DPC.exe` and specify a data folder:

```powershell
.\DPC.exe "C:\path\to\data"
```

When using PowerShell, run `.\DPC.exe` without a folder argument to process the current working directory. Running from PowerShell lets you review progress, warnings, and skipped-profile messages after processing finishes.

Each ASC produces `<ASC name>_comparison.xlsx` in the selected folder, with profile sheets, tables, overlay/difference charts, and an optional conventional-field `Summary`. Existing filenames receive a numeric suffix. Skipped profiles show their reason; out-of-grid TPS values remain blank and are excluded from comparisons. Internal geometry is in mm; analysis tables and chart axes use cm, while ASC metadata are preserved.

## Inputs and matching

- ASC profiles are delimited by `$STOM` / `$ENOM`. Point columns are **crossline, inline, depth, measured dose**, with coordinates in mm. Supported scans are crossline (`OPP/X`), inline (`OPP/Y`), diagonal (`DPR`), and PDD (`OPD`).
- The workflow assumes photon beams, square symmetric fields, SAD = 100 cm, and standardized HFS, axis-aligned Eclipse geometry. Field size and SSD use integer mm, e.g. `%FLSZ 100*100` and `%SSD 0900`.
- Each RTDOSE links to its RTPLAN by SOP Instance UID. The plan must have exactly one usable photon treatment beam, energy and FF/FFF metadata, and exactly one valid dose-reference point named `ORIGIN` (case-insensitive).
- Matching uses **RTPlanLabel + nominal energy + FF/FFF**. For example, `%FLSZ 100*100`, `%SSD 0900`, and `# Edge - 6 MV (FFF)` require plan label `10x10_SSD90`, energy 6, and FFF mode. Label matching ignores case and whitespace; machine names and DICOM filenames are not matching keys. Separate datasets with colliding keys into different runs.
- Duplicate or missing RTDOSE/RTPLAN SOP Instance UIDs, ambiguous matching keys, and unreadable critical DICOMs stop the batch. Candidates with unusable matching metadata and unmatched or malformed profiles are reported and skipped.

## Sampling and calculations

RTDOSE geometry comes from `ImagePositionPatient`, `PixelSpacing`, and `GridFrameOffsetVector`, with `DoseGridScaling` applied and coordinates shifted relative to RTPLAN `ORIGIN`. The **current fixed mapping** is:

```text
(DICOM X, Y, Z) relative to ORIGIN = (ASC crossline, ASC depth, -ASC inline)
```

Trilinear interpolation samples TPS dose at the original measurement coordinates, preserving measured positions, order, and doses. Lateral TPS values are normalized to the TPS central-axis dose at the profile depth; PDD TPS values to the maximum positive TPS dose sampled at the available ASC depths. Measurements must already be on the intended comparison scale.

All pointwise differences use `100 × (measured − normalized TPS) / normalized TPS`.

- **Conventional lateral fields >30 mm:** compare the central 80% of the depth-diverged nominal field width, with a diagonal correction. Cells are yellow for `1 < |difference| ≤ 2%` and red above 2%. `Summary` reports signed left/right extrema and means, counts/percentages for `1 ≤ |difference| ≤ 2%` and >2%, sorted by type → depth → field and type → field → depth. PDD, small fields, and skipped profiles are excluded.
- **Small lateral fields ≤30 mm**, including 3 × 3 cm²: compare points with CAX-normalized TPS dose **≥20%**. Report point count, signed mean, mean absolute and maximum absolute dose differences; FWHM; left/right D50 shifts; field-center displacement; separate 80–20% penumbras; D20/D50/D80 crossing positions; and point and area symmetry. Results are descriptive, without tolerances, highlighting, or pass/fail.
- **PDD:** report interpolated PDD at 10 and 20 cm, empirical `TPR20,10 = 1.2661 × PDD20/PDD10 − 0.0595`, and relative measured-versus-TPS TPR difference (%). Pointwise PDD differences are red above 1% in magnitude.

Small-field crossings use each curve's own maximum. Both symmetry methods use its D50-derived center and equal extents limited by the nearer D20 crossing. Point symmetry reports maximum and mean absolute mirrored-dose differences in percentage points; area symmetry uses trapezoidal integration and signed `200 × (right area − left area)/(right area + left area)`. Derived calculations interpolate between acquired samples without replacing the primary sampling grid. Shifts and metric deltas are **measured minus TPS**; unavailable metrics show `N/A`.

## Development notes: what remains in this version

The sequence below follows dependencies between the six project Markdown files; it is not a verified commit chronology.

| Notes | Current status |
| --- | --- |
| [Prompt.md](Prompt.md) | Base parsing, sampling, normalization, conventional comparison, and Excel reporting retained. Its original axis mapping was superseded by the mapping documented above. |
| [Patch.md](Patch.md) | Metadata-based matching implemented; replaces machine/filename matching and RTDOSE filename-uniqueness checks with UID/key checks. |
| [Patch2.md](Patch2.md) | Small-field geometry and PDD-derived TPR implemented. Its ≥80% comparison region is now ≥20%; TPR difference is relative (%) rather than an absolute delta. |
| [Patch3.md](Patch3.md) | Conventional-field `Summary` implemented. A generated conventional profile with no evaluated central points still gets a row with zero counts and blank metrics. |
| [Small field refactoring.md](<Small field refactoring.md>) | ≥20% comparison, field-center displacement, point symmetry, and expanded dose-difference statistics implemented. |
| [Follow up task for SF.md](<Follow up task for SF.md>) | Area symmetry added alongside point symmetry. |

All major requested feature groups are present; the notes identify no separate unfinished feature backlog. PDD highlighting currently uses the single >1% red threshold, superseding the original shared yellow/red scheme. Gamma analysis, rectangular fields, general beam-orientation handling, measured-data smoothing/recentering/renormalization, and the old `water_phantom.txt`/voxel-walking approach are outside the current design.

## Acknowledgements

The idea of sampling DICOM RTDOSE at measured scan coordinates was taken from **Mark Geurts' [Water Tank TPS Comparison Tool (`mwgeurts/water_tank`)](https://github.com/mwgeurts/water_tank)**.

Scientific libraries used here that provide citation guidance:

- **NumPy** — Harris, C. R., et al. (2020). *Array programming with NumPy*. Nature, 585, 357–362. [doi:10.1038/s41586-020-2649-2](https://doi.org/10.1038/s41586-020-2649-2).
- **SciPy** — Virtanen, P., et al. (2020). *SciPy 1.0: Fundamental algorithms for scientific computing in Python*. Nature Methods, 17, 261–272. [doi:10.1038/s41592-019-0686-2](https://doi.org/10.1038/s41592-019-0686-2). Used for dose-grid interpolation.
- **pydicom** — Mason, D. L., et al. *pydicom: An open source DICOM library*. [Software](https://github.com/pydicom/pydicom); [official citation guidance](https://pydicom.github.io/pydicom/dev/faq.html#how-do-i-cite-pydicom). Used to read RTDOSE and RTPLAN.
