# Machine Monitor — Excel VBA Dashboard

## Original problem statement
> Build a machine monitoring dashboard for up to 16 machines where I can batch
> upload performance data via CSV. Each machine has 3 metrics that can vary,
> but I need to group machines by tags to compare their performance
> side-by-side on one screen. Include a date range selector and a simple
> report generation button. There will be live update in the future.
> **This must run in Microsoft Excel.**

## Confirmed user choices (2026-01)
- **Platform:** Pure Excel `.xlsm` workbook + VBA macros (no server, no add-in).
- **CSV columns:** `machine_id, tag, timestamp, metric1, metric2, metric3`.
- **Metric names:** user-defined per machine (Config sheet).
- **Report output:** New sheet snapshot **and** PDF export.
- **Tag assignment:** manual, via Config sheet.

## Architecture
Delivered as a two-file package built by a Python generator:
| Artifact | Role |
|---|---|
| `build_workbook.py` | openpyxl builder — writes `MachineMonitor.xlsx` with all sheets, tables, charts, sample data, KPI formulas, named ranges, data-validation dropdowns. |
| `vba/MachineMonitor.bas` | VBA module — imported once via VBE, then user saves as `.xlsm`. Runs `SetupWorkbook` to install shape-buttons + hotkeys. |
| `MachineMonitor.xlsx` | Generated base workbook. |
| `sample_data.csv` | 32-row sample for CSV import demo. |
| `README.md` | 2-minute setup + daily usage guide. |

## Sheets
1. **Dashboard** — title strip, filters (Start/End Date, Tag, Metric focus),
   4 action buttons, 5 KPI tiles, 3 side-by-side comparison line charts,
   hidden staging area (cols R:AF) populated by `RefreshDashboard`.
2. **Config** — 16-row `tblMachines` (Machine ID, Name, Tag, 3 Metric names, Active).
3. **Data** — `tblData` raw store (auto-appended by `ImportCSV`), seeded with 480 sample points.
4. **Reports** — log of every generated report + PDF path.
5. **README** — inline setup guide inside Excel.

## VBA public entry points
| Macro | Hotkey | Behaviour |
|---|---|---|
| `SetupWorkbook` | — | One-time install: adds 4 amber shape-buttons + registers hotkeys. |
| `ImportCSV` | `Ctrl+Shift+I` | File picker → append to `tblData` → auto-refresh. |
| `RefreshDashboard` | `Ctrl+Shift+R` | Recompute staging + rebind chart series to active tag/date range. |
| `GenerateReport` | `Ctrl+Shift+G` | Duplicate Dashboard as static sheet + export PDF alongside workbook + log to Reports. |
| `ClearData` | — | Wipe `tblData` (with confirm). |
| `ToggleLive` / `AutoRefresh` | — | 60-sec live refresh scaffold, ready for future feed. |

## Design system
- Industrial palette: `PANEL #13232B`, `AMBER #E8A33D`, `TEAL #3FA7B8`, `LIME #9BC53D`, `PAPER #F4F7F8`.
- Consolas monospace throughout — reads on control-room monitors.
- Chart series colours (teal / amber / lime / red) support up to 4 machines per tag before repeating.

## What's implemented (2026-01)
- End-to-end Excel workflow: configure → import → filter → compare → report.
- Sample data + charts render immediately on first open, before any CSV import.
- Live-update hook in place (off by default).

## Prioritised backlog
- **P1** Real live feed: PowerQuery source or shared drop-folder polling in `AutoRefresh`.
- **P2** Alert thresholds per metric (highlight KPI red when out-of-band).
- **P2** Multi-tag comparison (show 2+ tags side-by-side, not just one).
- **P3** Machine drill-down sheet on click.
- **P3** Anomaly flag column in the report PDF.

## Not applicable
- No web frontend, no backend service, no auth, no DB — pure Excel.
