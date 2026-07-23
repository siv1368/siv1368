"""
Builds MachineMonitor.xlsx — the base workbook that ships with all sheets,
tables, sample data, formatting, and named ranges pre-configured.

After running this script, open MachineMonitor.xlsx in Excel, import the
VBA module from vba/MachineMonitor.bas, run the SetupWorkbook macro once,
then Save As MachineMonitor.xlsm (macro-enabled).

See README.md for the full setup guide.
"""

from datetime import datetime, timedelta
from pathlib import Path
import random

from openpyxl import Workbook
from openpyxl.chart import LineChart, Reference
from openpyxl.styles import (
    Alignment, Border, Font, PatternFill, Side, NamedStyle
)
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.worksheet.datavalidation import DataValidation

OUT_XLSX = Path(__file__).parent / "MachineMonitor.xlsx"

# ---------- Palette (industrial dark-teal + amber accents) ----------
INK        = "0F1A20"
PANEL      = "13232B"
PANEL_2    = "1B303A"
DIVIDER    = "24404C"
STEEL      = "6E8A96"
FROST      = "E6EEF1"
PAPER      = "F4F7F8"
AMBER      = "E8A33D"
AMBER_2    = "C87A1F"
LIME       = "9BC53D"
RED        = "D64545"
CHART_A    = "3FA7B8"
CHART_B    = "E8A33D"
CHART_C    = "9BC53D"


# ---------- Common styles ----------
def make_styles(wb: Workbook) -> None:
    styles = {
        "hdr_title": NamedStyle(
            name="hdr_title",
            font=Font(name="Consolas", size=22, bold=True, color=FROST),
            fill=PatternFill("solid", fgColor=PANEL),
            alignment=Alignment(horizontal="left", vertical="center", indent=1),
        ),
        "hdr_sub": NamedStyle(
            name="hdr_sub",
            font=Font(name="Consolas", size=10, color=STEEL),
            fill=PatternFill("solid", fgColor=PANEL),
            alignment=Alignment(horizontal="left", vertical="center", indent=1),
        ),
        "section": NamedStyle(
            name="section",
            font=Font(name="Consolas", size=11, bold=True, color=AMBER),
            fill=PatternFill("solid", fgColor=PANEL_2),
            alignment=Alignment(horizontal="left", vertical="center", indent=1),
        ),
        "label": NamedStyle(
            name="label",
            font=Font(name="Consolas", size=10, color=STEEL),
            alignment=Alignment(horizontal="left", vertical="center", indent=1),
        ),
        "value": NamedStyle(
            name="value",
            font=Font(name="Consolas", size=11, bold=True, color=INK),
            fill=PatternFill("solid", fgColor=FROST),
            alignment=Alignment(horizontal="center", vertical="center"),
            border=Border(
                left=Side(style="thin", color=DIVIDER),
                right=Side(style="thin", color=DIVIDER),
                top=Side(style="thin", color=DIVIDER),
                bottom=Side(style="thin", color=DIVIDER),
            ),
        ),
        "input": NamedStyle(
            name="input",
            font=Font(name="Consolas", size=10, color=INK),
            fill=PatternFill("solid", fgColor="FFFDF3"),
            alignment=Alignment(horizontal="left", vertical="center", indent=1),
            border=Border(
                left=Side(style="thin", color=AMBER_2),
                right=Side(style="thin", color=AMBER_2),
                top=Side(style="thin", color=AMBER_2),
                bottom=Side(style="thin", color=AMBER_2),
            ),
        ),
        "kpi_label": NamedStyle(
            name="kpi_label",
            font=Font(name="Consolas", size=9, color=STEEL),
            fill=PatternFill("solid", fgColor=PANEL_2),
            alignment=Alignment(horizontal="left", vertical="center", indent=1),
        ),
        "kpi_val": NamedStyle(
            name="kpi_val",
            font=Font(name="Consolas", size=18, bold=True, color=AMBER),
            fill=PatternFill("solid", fgColor=PANEL_2),
            alignment=Alignment(horizontal="left", vertical="center", indent=1),
            number_format="0.00",
        ),
        "btn": NamedStyle(
            name="btn",
            font=Font(name="Consolas", size=11, bold=True, color=INK),
            fill=PatternFill("solid", fgColor=AMBER),
            alignment=Alignment(horizontal="center", vertical="center"),
            border=Border(
                left=Side(style="medium", color=AMBER_2),
                right=Side(style="medium", color=AMBER_2),
                top=Side(style="medium", color=AMBER_2),
                bottom=Side(style="medium", color=AMBER_2),
            ),
        ),
        "th": NamedStyle(
            name="th",
            font=Font(name="Consolas", size=10, bold=True, color=FROST),
            fill=PatternFill("solid", fgColor=PANEL_2),
            alignment=Alignment(horizontal="center", vertical="center"),
        ),
        "th_left": NamedStyle(
            name="th_left",
            font=Font(name="Consolas", size=10, bold=True, color=FROST),
            fill=PatternFill("solid", fgColor=PANEL_2),
            alignment=Alignment(horizontal="left", vertical="center", indent=1),
        ),
        "pick_id": NamedStyle(
            name="pick_id",
            font=Font(name="Consolas", size=10, bold=True, color=INK),
            fill=PatternFill("solid", fgColor=FROST),
            alignment=Alignment(horizontal="left", vertical="center", indent=1),
            border=Border(
                left=Side(style="thin", color=DIVIDER),
                right=Side(style="thin", color=DIVIDER),
                top=Side(style="thin", color=DIVIDER),
                bottom=Side(style="thin", color=DIVIDER),
            ),
        ),
        "pick_show": NamedStyle(
            name="pick_show",
            font=Font(name="Consolas", size=10, bold=True, color=INK),
            fill=PatternFill("solid", fgColor="E9F3E4"),
            alignment=Alignment(horizontal="center", vertical="center"),
            border=Border(
                left=Side(style="thin", color=DIVIDER),
                right=Side(style="thin", color=DIVIDER),
                top=Side(style="thin", color=DIVIDER),
                bottom=Side(style="thin", color=DIVIDER),
            ),
        ),
    }
    for s in styles.values():
        if s.name not in wb.named_styles:
            wb.add_named_style(s)


# ---------- Sheets ----------
def build_config(wb: Workbook) -> None:
    ws = wb.create_sheet("Config")
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = AMBER

    ws.row_dimensions[1].height = 46
    ws.row_dimensions[2].height = 22
    ws.merge_cells("B1:H1")
    ws["B1"] = "MACHINE  CONFIGURATION"
    ws["B1"].style = "hdr_title"
    ws.merge_cells("B2:H2")
    ws["B2"] = "Register up to 16 machines. Assign a tag to group them for side-by-side comparison."
    ws["B2"].style = "hdr_sub"

    headers = ["Machine ID", "Machine Name", "Tag",
               "Metric 1 Name", "Metric 2 Name", "Metric 3 Name", "Active"]
    for i, h in enumerate(headers, start=2):
        c = ws.cell(row=4, column=i, value=h)
        c.style = "th"
    ws.row_dimensions[4].height = 26

    tags = ["Line-A", "Line-B", "Line-C", "Line-D"]
    metric_sets = [
        ("Temperature (°C)", "Vibration (mm/s)", "Output (units/h)"),
        ("Pressure (bar)",   "Flow (L/min)",     "RPM"),
        ("Torque (Nm)",      "Current (A)",      "Voltage (V)"),
        ("Speed (m/s)",      "Load (%)",         "Cycles/hr"),
    ]
    for row_i in range(16):
        r = 5 + row_i
        tag_idx = row_i // 4
        ws.cell(row=r, column=2, value=f"M{row_i+1:02d}")
        ws.cell(row=r, column=3, value=f"Machine {row_i+1:02d}")
        ws.cell(row=r, column=4, value=tags[tag_idx])
        ws.cell(row=r, column=5, value=metric_sets[tag_idx][0])
        ws.cell(row=r, column=6, value=metric_sets[tag_idx][1])
        ws.cell(row=r, column=7, value=metric_sets[tag_idx][2])
        ws.cell(row=r, column=8, value="Yes")
        for col in range(2, 9):
            ws.cell(row=r, column=col).style = "value"

    widths = {"A": 2, "B": 12, "C": 18, "D": 12, "E": 22, "F": 22, "G": 22, "H": 10}
    for col, w in widths.items():
        ws.column_dimensions[col].width = w

    tbl = Table(displayName="tblMachines", ref="B4:H20")
    tbl.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium2", showRowStripes=True
    )
    ws.add_table(tbl)

    dv = DataValidation(type="list", formula1='"Yes,No"', allow_blank=True)
    ws.add_data_validation(dv)
    dv.add("H5:H20")


def build_data(wb: Workbook) -> None:
    ws = wb.create_sheet("Data")
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = CHART_A

    ws.row_dimensions[1].height = 46
    ws.merge_cells("B1:G1")
    ws["B1"] = "RAW  DATA  STORE"
    ws["B1"].style = "hdr_title"
    ws.merge_cells("B2:G2")
    ws["B2"] = "Populated automatically by the CSV import macro. Do not edit manually."
    ws["B2"].style = "hdr_sub"

    headers = ["Machine ID", "Tag", "Timestamp", "Metric 1", "Metric 2", "Metric 3"]
    for i, h in enumerate(headers, start=2):
        ws.cell(row=4, column=i, value=h).style = "th"
    ws.row_dimensions[4].height = 26

    tags = ["Line-A"] * 4 + ["Line-B"] * 4 + ["Line-C"] * 4 + ["Line-D"] * 4
    rng = random.Random(7)
    start = datetime(2026, 1, 1)
    row = 5
    for day in range(30):
        for m_idx in range(16):
            mid = f"M{m_idx+1:02d}"
            tag = tags[m_idx]
            ts = start + timedelta(days=day, hours=8)
            base = 50 + m_idx * 2
            m1 = round(base + rng.uniform(-5, 5), 2)
            m2 = round(10 + rng.uniform(0, 8), 2)
            m3 = round(100 + m_idx * 5 + rng.uniform(-15, 15), 2)
            ws.cell(row=row, column=2, value=mid)
            ws.cell(row=row, column=3, value=tag)
            ws.cell(row=row, column=4, value=ts).number_format = "yyyy-mm-dd hh:mm"
            ws.cell(row=row, column=5, value=m1)
            ws.cell(row=row, column=6, value=m2)
            ws.cell(row=row, column=7, value=m3)
            row += 1

    end_row = row - 1
    tbl = Table(displayName="tblData", ref=f"B4:G{end_row}")
    tbl.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium9", showRowStripes=True
    )
    ws.add_table(tbl)

    widths = {"A": 2, "B": 12, "C": 12, "D": 20, "E": 12, "F": 12, "G": 12}
    for col, w in widths.items():
        ws.column_dimensions[col].width = w


def build_dashboard(wb: Workbook) -> None:
    ws = wb.create_sheet("Dashboard", 0)
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = INK

    for col in range(1, 40):
        ws.column_dimensions[get_column_letter(col)].width = 12
    ws.column_dimensions["A"].width = 2

    # ---------- Title strip ----------
    ws.row_dimensions[1].height = 8
    ws.row_dimensions[2].height = 44
    ws.row_dimensions[3].height = 22
    ws.merge_cells("B2:P2")
    ws["B2"] = "MACHINE  MONITORING  ///  CONTROL  ROOM"
    ws["B2"].style = "hdr_title"
    ws.merge_cells("B3:P3")
    ws["B3"] = "Batch-upload CSV → filter by tag & date → pick machines → compare side-by-side → export report."
    ws["B3"].style = "hdr_sub"

    # ---------- Filters ----------
    ws.row_dimensions[5].height = 22
    ws.merge_cells("B5:P5")
    ws["B5"] = "  FILTERS"
    ws["B5"].style = "section"
    ws.row_dimensions[6].height = 8
    ws.row_dimensions[7].height = 24
    ws.row_dimensions[8].height = 24

    ws["B7"] = "Start Date"; ws["B7"].style = "label"
    ws["B8"] = "End Date";   ws["B8"].style = "label"
    ws.merge_cells("C7:D7"); ws.merge_cells("C8:D8")
    ws["C7"] = datetime(2026, 1, 1);  ws["C7"].number_format = "yyyy-mm-dd"
    ws["C8"] = datetime(2026, 1, 30); ws["C8"].number_format = "yyyy-mm-dd"
    ws["C7"].style = "value"; ws["C8"].style = "value"

    ws["F7"] = "Tag Group"; ws["F7"].style = "label"
    ws.merge_cells("G7:H7")
    ws["G7"] = "Line-A"; ws["G7"].style = "value"
    dv_tag = DataValidation(type="list", formula1="=Config!$D$5:$D$20", allow_blank=False)
    ws.add_data_validation(dv_tag)
    dv_tag.add("G7:H7")

    ws["F8"] = "Metric Focus"; ws["F8"].style = "label"
    ws.merge_cells("G8:H8")
    ws["G8"] = "Metric 1"; ws["G8"].style = "value"
    dv_metric = DataValidation(type="list", formula1='"Metric 1,Metric 2,Metric 3"', allow_blank=False)
    ws.add_data_validation(dv_metric)
    dv_metric.add("G8:H8")

    ws.merge_cells("J7:L7"); ws["J7"] = "▸  IMPORT CSV";            ws["J7"].style = "btn"
    ws.merge_cells("J8:L8"); ws["J8"] = "▸  REFRESH DASHBOARD";     ws["J8"].style = "btn"
    ws.merge_cells("N7:P7"); ws["N7"] = "▸  GENERATE REPORT";       ws["N7"].style = "btn"
    ws.merge_cells("N8:P8"); ws["N8"] = "▸  CLEAR DATA";            ws["N8"].style = "btn"

    # ---------- KPI row ----------
    ws.row_dimensions[10].height = 22
    ws.merge_cells("B10:P10")
    ws["B10"] = "  KEY  METRICS  (filtered)"
    ws["B10"].style = "section"
    ws.row_dimensions[11].height = 18
    ws.row_dimensions[12].height = 34

    kpi_specs = [
        ("MACHINES",
         "=SUMPRODUCT((Config!D5:D20=G7)*(Config!H5:H20=\"Yes\"))", "0"),
        ("DATA POINTS",
         "=SUMPRODUCT((Data!C5:C20000=G7)*(Data!D5:D20000>=C7)*(Data!D5:D20000<=C8+1))",
         "#,##0"),
        ("AVG METRIC 1",
         "=IFERROR(SUMPRODUCT((Data!C5:C20000=G7)*(Data!D5:D20000>=C7)*(Data!D5:D20000<=C8+1)*Data!E5:E20000)"
         "/MAX(1,SUMPRODUCT((Data!C5:C20000=G7)*(Data!D5:D20000>=C7)*(Data!D5:D20000<=C8+1))),0)", "0.00"),
        ("AVG METRIC 2",
         "=IFERROR(SUMPRODUCT((Data!C5:C20000=G7)*(Data!D5:D20000>=C7)*(Data!D5:D20000<=C8+1)*Data!F5:F20000)"
         "/MAX(1,SUMPRODUCT((Data!C5:C20000=G7)*(Data!D5:D20000>=C7)*(Data!D5:D20000<=C8+1))),0)", "0.00"),
        ("AVG METRIC 3",
         "=IFERROR(SUMPRODUCT((Data!C5:C20000=G7)*(Data!D5:D20000>=C7)*(Data!D5:D20000<=C8+1)*Data!G5:G20000)"
         "/MAX(1,SUMPRODUCT((Data!C5:C20000=G7)*(Data!D5:D20000>=C7)*(Data!D5:D20000<=C8+1))),0)", "0.00"),
    ]
    for (label, formula, fmt), col in zip(kpi_specs, ["B", "E", "H", "K", "N"]):
        end_col = get_column_letter(ord(col) - ord("A") + 3)
        ws.merge_cells(f"{col}11:{end_col}11")
        ws.merge_cells(f"{col}12:{end_col}12")
        ws[f"{col}11"] = "  " + label; ws[f"{col}11"].style = "kpi_label"
        ws[f"{col}12"] = formula;      ws[f"{col}12"].style = "kpi_val"
        ws[f"{col}12"].number_format = fmt

    # ---------- ADD NEW MACHINE ----------
    ws.row_dimensions[14].height = 22
    ws.merge_cells("B14:P14")
    ws["B14"] = "  ADD  NEW  MACHINE"
    ws["B14"].style = "section"

    ws.row_dimensions[15].height = 20
    add_labels = ["Machine ID", "Machine Name", "Tag",
                  "Metric 1 Name", "Metric 2 Name", "Metric 3 Name"]
    add_cols = ["B", "D", "F", "H", "J", "L"]
    for lab, col in zip(add_labels, add_cols):
        end_col = get_column_letter(ord(col) - ord("A") + 2)
        ws.merge_cells(f"{col}15:{end_col}15")
        ws[f"{col}15"] = lab
        ws[f"{col}15"].style = "th_left"

    ws.row_dimensions[16].height = 26
    # input cells
    for col in add_cols:
        end_col = get_column_letter(ord(col) - ord("A") + 2)
        ws.merge_cells(f"{col}16:{end_col}16")
        ws[f"{col}16"].style = "input"
    # placeholder text (grey hint)
    ws["B16"] = ""
    ws["D16"] = ""
    ws["F16"] = ""
    ws["H16"] = ""
    ws["J16"] = ""
    ws["L16"] = ""
    # ADD button at N16:P16
    ws.merge_cells("N16:P16")
    ws["N16"] = "▸  ADD MACHINE"
    ws["N16"].style = "btn"

    # ---------- MACHINES TO COMPARE (picker) ----------
    ws.row_dimensions[18].height = 22
    ws.merge_cells("B18:P18")
    ws["B18"] = "  MACHINES  TO  COMPARE   (tick / untick, then REFRESH)"
    ws["B18"].style = "section"

    ws.row_dimensions[19].height = 20
    # two side-by-side picker groups: B:D and F:H, rows 20-27 → 16 slots
    pick_headers = [("B", "C", "D"), ("F", "G", "H")]
    for grp in pick_headers:
        ws[f"{grp[0]}19"] = "Show";         ws[f"{grp[0]}19"].style = "th"
        ws[f"{grp[1]}19"] = "Machine ID";   ws[f"{grp[1]}19"].style = "th"
        ws[f"{grp[2]}19"] = "Machine Name"; ws[f"{grp[2]}19"].style = "th_left"

    dv_show = DataValidation(type="list", formula1='"Yes,No"', allow_blank=True)
    ws.add_data_validation(dv_show)

    for i in range(8):
        r = 20 + i
        ws.row_dimensions[r].height = 20
        for grp in pick_headers:
            ws[f"{grp[0]}{r}"].style = "pick_show"
            ws[f"{grp[1]}{r}"].style = "pick_id"
            ws[f"{grp[2]}{r}"].style = "value"
            ws.column_dimensions[grp[2]].width = 18
        dv_show.add(f"B{r}:B{r}")
        dv_show.add(f"F{r}:F{r}")

    # Pre-populate picker so charts render on first open (all 4 machines of Line-A)
    default_mach = [("M01", "Machine 01"), ("M02", "Machine 02"),
                    ("M03", "Machine 03"), ("M04", "Machine 04")]
    for i, (mid, mname) in enumerate(default_mach):
        r = 20 + i
        ws[f"B{r}"] = "Yes"
        ws[f"C{r}"] = mid
        ws[f"D{r}"] = mname

    # J18:P27 will hold summary hint / legend
    ws.merge_cells("J20:P20"); ws["J20"] = "  Tip: change the Tag Group above"; ws["J20"].style = "label"
    ws.merge_cells("J21:P21"); ws["J21"] = "  then click REFRESH DASHBOARD to"; ws["J21"].style = "label"
    ws.merge_cells("J22:P22"); ws["J22"] = "  re-populate this machine list."; ws["J22"].style = "label"
    ws.merge_cells("J24:P24"); ws["J24"] = "  Untick any machine you want"; ws["J24"].style = "label"
    ws.merge_cells("J25:P25"); ws["J25"] = "  to exclude from the charts,"; ws["J25"].style = "label"
    ws.merge_cells("J26:P26"); ws["J26"] = "  then click REFRESH again."; ws["J26"].style = "label"

    # Track last tag so RefreshMachinePicker knows when to reset selections
    ws["AH1"] = "Line-A"   # hidden cell (col AH = 34)

    # ---------- Comparison charts ----------
    ws.row_dimensions[29].height = 22
    ws.merge_cells("B29:P29")
    ws["B29"] = "  SIDE-BY-SIDE  COMPARISON"
    ws["B29"].style = "section"

    # Staging area (hidden cols R:AF, rows 30+)
    # Rebuilt by VBA RefreshDashboard.
    ws["R30"] = "Date"
    ws["R30"].style = "th"
    for i in range(1, 5):
        ws.cell(row=30, column=17 + i, value=f"M{i:02d}").style = "th"

    start = datetime(2026, 1, 1)
    for d in range(30):
        ws.cell(row=31 + d, column=18, value=start + timedelta(days=d)).number_format = "yyyy-mm-dd"
        for m in range(1, 5):
            ws.cell(row=31 + d, column=18 + m,
                    value=50 + m * 3 + (d % 7) - 3 + random.random() * 4)

    def make_chart(title, top_left):
        chart = LineChart()
        chart.title = title
        chart.style = 2
        chart.height = 8
        chart.width = 14
        data = Reference(ws, min_col=19, max_col=22, min_row=30, max_row=60)
        cats = Reference(ws, min_col=18, min_row=31, max_row=60)
        chart.add_data(data, titles_from_data=True)
        chart.set_categories(cats)
        palette = [CHART_A, CHART_B, CHART_C, "D64545"]
        for i, ser in enumerate(chart.series):
            ser.graphicalProperties.line.solidFill = palette[i % 4]
            ser.graphicalProperties.line.width = 20000
            ser.smooth = True
        ws.add_chart(chart, top_left)

    make_chart("Metric 1 — comparison", "B30")
    make_chart("Metric 2 — comparison", "I30")

    chart3 = LineChart()
    chart3.title = "Metric 3 — comparison"
    chart3.style = 2
    chart3.height = 8; chart3.width = 28
    data = Reference(ws, min_col=19, max_col=22, min_row=30, max_row=60)
    cats = Reference(ws, min_col=18, min_row=31, max_row=60)
    chart3.add_data(data, titles_from_data=True)
    chart3.set_categories(cats)
    for i, ser in enumerate(chart3.series):
        ser.graphicalProperties.line.solidFill = [CHART_A, CHART_B, CHART_C, "D64545"][i]
        ser.graphicalProperties.line.width = 20000
        ser.smooth = True
    ws.add_chart(chart3, "B47")

    # Footer
    ws.row_dimensions[63].height = 22
    ws.merge_cells("B63:P63")
    ws["B63"] = "  TIP: Ctrl+Shift+I = Import   |   Ctrl+Shift+R = Refresh   |   Ctrl+Shift+G = Generate Report"
    ws["B63"].style = "hdr_sub"

    # Hide the staging columns (R:AF)
    for col in range(18, 35):
        ws.column_dimensions[get_column_letter(col)].hidden = True

    # Named ranges for VBA
    wb.defined_names["StartDate"]   = DefinedName("StartDate",   attr_text="Dashboard!$C$7")
    wb.defined_names["EndDate"]     = DefinedName("EndDate",     attr_text="Dashboard!$C$8")
    wb.defined_names["TagFilter"]   = DefinedName("TagFilter",   attr_text="Dashboard!$G$7")
    wb.defined_names["MetricFocus"] = DefinedName("MetricFocus", attr_text="Dashboard!$G$8")
    wb.defined_names["AddMID"]      = DefinedName("AddMID",      attr_text="Dashboard!$B$16")
    wb.defined_names["AddName"]     = DefinedName("AddName",     attr_text="Dashboard!$D$16")
    wb.defined_names["AddTag"]      = DefinedName("AddTag",      attr_text="Dashboard!$F$16")
    wb.defined_names["AddM1"]       = DefinedName("AddM1",       attr_text="Dashboard!$H$16")
    wb.defined_names["AddM2"]       = DefinedName("AddM2",       attr_text="Dashboard!$J$16")
    wb.defined_names["AddM3"]       = DefinedName("AddM3",       attr_text="Dashboard!$L$16")
    wb.defined_names["LastTag"]     = DefinedName("LastTag",     attr_text="Dashboard!$AH$1")


def build_reports(wb: Workbook) -> None:
    ws = wb.create_sheet("Reports")
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = LIME
    ws.row_dimensions[1].height = 46
    ws.merge_cells("B1:H1")
    ws["B1"] = "REPORT  ARCHIVE"
    ws["B1"].style = "hdr_title"
    ws.merge_cells("B2:H2")
    ws["B2"] = "Every 'Generate Report' click appends a snapshot below and exports a PDF next to the workbook."
    ws["B2"].style = "hdr_sub"

    headers = ["Generated At", "Tag", "Start", "End", "Machines", "Points", "Report Sheet", "PDF File"]
    for i, h in enumerate(headers, start=2):
        ws.cell(row=4, column=i, value=h).style = "th"
    ws.row_dimensions[4].height = 26

    widths = {"A": 2, "B": 20, "C": 12, "D": 12, "E": 12, "F": 10, "G": 10, "H": 24, "I": 32}
    for col, w in widths.items():
        ws.column_dimensions[col].width = w


def build_readme_sheet(wb: Workbook) -> None:
    ws = wb.create_sheet("README")
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = STEEL
    ws.row_dimensions[1].height = 46
    ws.merge_cells("B1:J1")
    ws["B1"] = "SETUP  &  USAGE"
    ws["B1"].style = "hdr_title"

    steps = [
        ("1.", "ONE-TIME SETUP",  "Alt+F11 → File → Import File… → pick vba/MachineMonitor.bas. Save As Macro-Enabled Workbook (*.xlsm). Then Alt+F8 → SetupWorkbook → Run."),
        ("2.", "CONFIGURE",       "Config sheet holds up to 16 machines. Or add machines live from the Dashboard using the ADD MACHINE panel."),
        ("3.", "IMPORT DATA",     "Dashboard → IMPORT CSV. CSV columns: machine_id, tag, timestamp, metric1, metric2, metric3."),
        ("4.", "FILTER",          "Set Start / End dates + Tag. Click REFRESH DASHBOARD to populate the machine picker."),
        ("5.", "PICK MACHINES",   "In 'MACHINES TO COMPARE', untick any machine you want to exclude, then REFRESH again."),
        ("6.", "REPORT",          "GENERATE REPORT creates a frozen snapshot sheet + a PDF next to the workbook."),
        ("7.", "SHORTCUTS",       "Ctrl+Shift+I Import   |   Ctrl+Shift+R Refresh   |   Ctrl+Shift+G Report."),
        ("8.", "LIVE UPDATES",    "Run ToggleLive to auto-refresh every 60 sec (for when you wire in a live feed later)."),
    ]
    for i, (n, title, body) in enumerate(steps, start=3):
        ws.cell(row=i, column=2, value=n).style = "kpi_label"
        ws.cell(row=i, column=3, value=title).style = "section"
        ws.merge_cells(start_row=i, start_column=4, end_row=i, end_column=10)
        ws.cell(row=i, column=4, value=body).style = "label"

    widths = {"A": 2, "B": 4, "C": 22, "D": 14, "E": 14, "F": 14, "G": 14, "H": 14, "I": 14, "J": 22}
    for col, w in widths.items():
        ws.column_dimensions[col].width = w


def main() -> None:
    wb = Workbook()
    default = wb.active
    wb.remove(default)

    make_styles(wb)

    build_config(wb)
    build_data(wb)
    build_reports(wb)
    build_readme_sheet(wb)
    build_dashboard(wb)  # inserted at index 0

    wb.active = 0
    wb.save(OUT_XLSX)
    print(f"✔ Wrote {OUT_XLSX}")


if __name__ == "__main__":
    main()
