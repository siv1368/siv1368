# Machine Monitor  —  Excel VBA Dashboard

Pure Microsoft Excel workbook (`.xlsm`) for monitoring up to **16 machines**,
grouped by tags, with **batch CSV upload**, **side-by-side comparison charts**,
a **date-range filter**, and **one-click report generation** (new sheet + PDF).

No server, no add-ins, no internet. Just Excel + VBA.

---

## Files in this repository

| File | Purpose |
|---|---|
| `build_workbook.py` | Regenerates `MachineMonitor.xlsx` from scratch (openpyxl). |
| `MachineMonitor.xlsx` | Pre-formatted starter workbook with sheets, tables, sample data, charts. |
| `vba/MachineMonitor.bas` | VBA module — you import this once into the workbook. |
| `sample_data.csv` | Example CSV for the Import macro. |

---

## One-time setup (2 minutes)

1. Open `MachineMonitor.xlsx` in Microsoft Excel (Windows or Mac).
2. Press **`Alt`+`F11`** to open the VBA editor.
3. **File → Import File…** → select `vba/MachineMonitor.bas` → close the editor.
4. **File → Save As** → pick **Excel Macro-Enabled Workbook (`*.xlsm`)** → save as `MachineMonitor.xlsm`.
5. Press **`Alt`+`F8`**, choose **`SetupWorkbook`**, click **Run**.
   This installs the four action buttons on the Dashboard and registers the
   hotkeys.

You're done. From now on always open the `.xlsm` version.

---

## Daily usage

### 1. Configure machines
On the **Config** sheet, edit each of the 16 rows:
| Machine ID | Machine Name | Tag | Metric 1 Name | Metric 2 Name | Metric 3 Name | Active |
|---|---|---|---|---|---|---|
| M01 | Boiler A | Line-A | Temperature (°C) | Vibration (mm/s) | Output (units/h) | Yes |

The **Tag** column is what you'll group by on the Dashboard.
Metric names are user-defined per machine.

### 2. Import a batch CSV
On the **Dashboard**, click **IMPORT CSV** (or press `Ctrl+Shift+I`).
CSV format (header row optional):

```
machine_id,tag,timestamp,metric1,metric2,metric3
M01,Line-A,2026-02-01 08:00,52.4,12.1,105.2
```

Rows append to the `tblData` table on the **Data** sheet.

### 3. Filter & compare
On the **Dashboard**:
- Pick **Start Date** and **End Date**.
- Pick a **Tag Group** — every machine in that tag appears as its own line
  on all three comparison charts.
- Click **REFRESH DASHBOARD** (or `Ctrl+Shift+R`).

The KPI strip recomputes automatically from formulas.

### 4. Generate a report
Click **GENERATE REPORT** (or `Ctrl+Shift+G`).

You get:
- A new sheet `Rpt_<Tag>_<timestamp>` — a frozen snapshot of the Dashboard.
- A PDF next to the workbook file.
- A log row on the **Reports** sheet.

### 5. Clear data
**CLEAR DATA** empties the `tblData` table (with confirmation).

### 6. Live updates (later)
The `ToggleLive` macro is already wired in. Run it once you have a live feed
(e.g. a PowerQuery refresh or a shared drop-folder). It calls
`RefreshDashboard` every 60 seconds until you toggle it off.

---

## Regenerating the base workbook

If you want to modify layout, colours, or sample data, edit `build_workbook.py`
and re-run:

```bash
python3 build_workbook.py
```

Then re-import the `.bas` module (steps 2–5 above).

---

## Design notes

- Palette: industrial deep-teal panels, amber action buttons, cool paper body.
  Charts use teal / amber / lime / red so up to 4 machines per tag stay
  distinguishable at a glance.
- Font: Consolas throughout — reads well on control-room monitors.
- All data lives in Excel tables (`tblMachines`, `tblData`) so formulas &
  charts auto-expand.
- Chart data is computed into a hidden staging area (columns R:AF on the
  Dashboard) by `RefreshDashboard`, so the charts stay simple and fast.
