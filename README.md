# GIS-Tool-QC_PMMS

**QC_PMMS** is a standalone Windows tool that checks the quality of street-network data stored in an Esri **File Geodatabase** — **no ArcGIS / ArcPy required**.

It is **read-only**: your geodatabase is never modified. Findings are exported to an Excel report and to Shapefiles you can open directly on the map.

> Created by **Ahmed Yasser**

---

## Features

- Runs without ArcGIS: reads `.gdb` through GDAL (`pyogrio`) and does geometry work with `shapely` / `pyproj`.
- Simple desktop UI (Tkinter): pick a GDB or a folder of GDBs, scan, select layers, run.
- 40+ rules covering identifiers, relations between layers, geometry and administrative codes.
- Excel report + one Shapefile per layer containing only the problem features.
- Suggests direction flips for wrongly-oriented sections (exported as a separate Shapefile).
- Fully configurable through `config.json` (rules on/off, field names, thresholds, layer mapping).
- Packages into a single `QC_PMMS.exe`.

## Rules overview

| Layer | Checks |
|---|---|
| **All layers** | Unwanted whitespace · Arabic normalization review · Z/M geometry · duplicate identifiers |
| **SECTION** | `SECTION_NO` = REGION + MAIN + TYPE + SERIAL · duplicates · base vs base+serial · serial sequence `01..n` · REGION reference · `MUNIC_CODE` · `SEC_TYPE` domain · max length · crossing INTERSECTION · street-name normalization · direction (region on the right, `FROM_STREET` at start, lane on the left) |
| **LANE** | `LANE_NO` = SECTION_NO + LANE_TYPE · lane-type domain · `MUNIC_CODE` · SECTION reference · lane longer than section (≥ 30 m) · lane type vs section type · lane length order · lane direction vs section |
| **INTERSECTION** | `INTERSECTION_NO` = MUNIC + MAIN1 + MAIN2 · `MUNIC_CODE` vs `DISTRICT_NO` · both main streets exist within 300 m |
| **INTERSEC_SAMPLE** | `INTSAMPLE_NO` = INTERSECTION_NO + sample code · duplicates |
| **REGION** | `REGION_NO` = MUNIC + BRANCH + DIST + REGION_CODE · `DISTRICT_NO` vs `REGION_NO` · sliver polygons · municipality / district / branch codes vs polygons |

Some checks are intentionally reported as `NOT_CONFIGURED` because they need project-specific conventions (e.g. numbering direction). The tool does not guess.

### Result statuses

`PASS` · `FAIL` · `REVIEW` · `DEPENDENCY_MISSING` · `NOT_RUN` · `NOT_CONFIGURED` · `ERROR`

## Requirements

- Windows 64-bit
- To run from source: Python 3.11 / 3.12 (64-bit) with Tk, and:

```bash
pip install -r requirements.txt
```

- Spatial checks (length, direction, intersections) need layers in a **projected CRS in metres**, using the **same CRS** across the layers that interact.

## Quick start

### Run from source

```bash
python QC_PMMS.py
```

### Build `QC_PMMS.exe`

Keep these files in one folder: `QC_PMMS.py`, `NAMAA_PNG_main_3.ico`, `build_QC_PMMS.bat`, then run:

```bat
build_QC_PMMS.bat
```

The script creates a clean virtual environment (`build_env`), installs the dependencies, and builds `dist\QC_PMMS.exe`. Progress and errors are written to `build_log.txt`. The first build takes several minutes.

### Using the tool

1. Browse to a `.gdb` (or a folder that contains several).
2. **Scan Layers** and select the layers to check.
3. **Run QC** and choose where to save the Excel report.
4. Review the results, fix the data in the original GDB (use `SRC_FID`), and run again.

## Outputs

- **Excel** — `QC Findings` (all results) and `Summary` (counts per layer / rule / status).
- **Shapefiles** — folder `<report name>_shp` next to the Excel file:
  - `<GDB>_<Layer>_findings.shp` — features with `FAIL` / `REVIEW` / `ERROR`, fields: `SRC_FID`, `RULE_ID`, `STATUS`, `SEVERITY`, `FIELD`, `VALUE`, `MESSAGE`.
  - `<GDB>_SECTION_suggested_flip.shp` — sections whose direction is suggested to be reversed (geometry already reversed), with a `REASON` field.
- **`qc.log`** — technical log, written next to the program.

## Configuration (`config.json`)

Created automatically on first run next to the program (or the `.exe`).

| Key | Purpose |
|---|---|
| `options.section_no_includes_serial` | `true` (default): `SECTION_NO` already contains the serial |
| `layer_prefixes` | Prefixes stripped from layer names (default `NAMAA_GIS_SDE_`) |
| `layer_map` | Map logical names to real layers, e.g. `{"SECTION": "Section_N", "LANE": "Namaa_GIS_SDE_LANE_N"}` |
| `rules` | Turn any check on / off |
| `fields` | Field names used by the checks |
| `thresholds` | Distances / lengths (defaults below) |

| Threshold | Default |
|---|---|
| `section_max_km` | 1.3 |
| `lane_longer_m` | 30 |
| `lane_order_diff_m` | 25 |
| `intersection_search_m` | 300 |
| `from_street_tol_m` | 10 |
| `region_side_offset_m` | 5 |
| `sliver_area_m2` | 1 |

### Layer names

Layers are recognised by name after removing the prefix: `SECTION`, `LANE`, `INTERSECTION`, `INTERSEC_SAMPLE`, `REGION` (checked layers) and `MAIN_STREETS`, `MUNICIPALITY`, `DISTRICT`, `BRANCH` (helper layers). Layers with other suffixes (e.g. `Section_N`) must be mapped with `layer_map`. If the same layer name exists in more than one GDB, the layer is skipped as `SOURCE_AMBIGUOUS`.

## Project structure

```
GIS-Tool-QC_PMMS/
├── QC_PMMS.py              # the tool (UI + rules)
├── NAMAA_PNG_main_3.ico    # application icon
├── build_QC_PMMS.bat       # builds QC_PMMS.exe in a clean environment
├── requirements.txt
└── README.md
```

## Known limitations

- Read-only by design: corrections (including direction flips) are done manually in the source GDB.
- Each layer is loaded fully into memory; very large layers (millions of features) may need a chunked reader.
- Direction/flip results are suggestions — review the Shapefile before editing data.
- Builds are time-limited; contact the author for a renewed build.

## Author

**Ahmed Yasser** — for questions, issues or a renewed build, please contact the author.

---

## نبذة بالعربي

**QC_PMMS** أداة مستقلة بتفحص جودة بيانات شبكة الشوارع في الـ File Geodatabase **من غير ArcGIS**. هي للقراءة فقط ومبتعدّلش في البيانات، وبتطلّع تقرير Excel وملفات Shapefile فيها الـ features الغلط بس، مع اقتراح بقلب اتجاه القطاعات الغلط. للتشغيل: افتح البرنامج، اختار الـ GDB، **Scan Layers**، اختار الطبقات، ثم **Run QC**.

للتواصل أو التجديد: **أحمد ياسر**.
