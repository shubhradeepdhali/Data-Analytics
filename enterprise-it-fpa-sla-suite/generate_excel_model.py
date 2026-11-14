"""
generate_excel_model.py
=======================
Builds FPA_Scenario_Governance_Model.xlsx with openpyxl, containing:

  Tab 1: Regional_Summary
         – Service master data + SUMIFS / XLOOKUP formulas
         – 2-variable What-If sensitivity grid template

  Tab 2: Sensitivity_Grid
         – Full 2-variable data table with revenue growth × cost reduction

  Tab 3: Data_Dictionary_Governance
         – Complete documentation of every table, field, data type,
           key status, and Power Query / DAX governance rule
"""

from pathlib import Path
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import (Font, PatternFill, Alignment, Border, Side,
                              numbers)
from openpyxl.utils import get_column_letter

# ── Paths ────────────────────────────────────────────────────────────────────
BASE_DIR = Path(__file__).resolve().parent
CLEAN_DIR = BASE_DIR / "cleaned_data"
OUT_FILE = BASE_DIR / "FPA_Scenario_Governance_Model.xlsx"

# ── Load cleaned data ────────────────────────────────────────────────────────
sl = pd.read_csv(CLEAN_DIR / "service_line_master.csv")
gl = pd.read_csv(CLEAN_DIR / "actual_gl_financials.csv")

# ── Styles ───────────────────────────────────────────────────────────────────
HEADER_FONT = Font(name="Calibri", bold=True, color="FFFFFF", size=11)
HEADER_FILL = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
ACCENT_FILL = PatternFill(start_color="D6E4F0", end_color="D6E4F0", fill_type="solid")
GREEN_FILL  = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
YELLOW_FILL = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
THIN_BORDER = Border(
    left=Side(style="thin"), right=Side(style="thin"),
    top=Side(style="thin"), bottom=Side(style="thin")
)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT_WRAP = Alignment(horizontal="left", vertical="top", wrap_text=True)

def style_header_row(ws, row, max_col):
    for col in range(1, max_col + 1):
        cell = ws.cell(row=row, column=col)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = CENTER
        cell.border = THIN_BORDER

def apply_borders(ws, min_row, max_row, max_col):
    for r in range(min_row, max_row + 1):
        for c in range(1, max_col + 1):
            ws.cell(row=r, column=c).border = THIN_BORDER


wb = Workbook()

# ═══════════════════════════════════════════════════════════════════════════════
#  TAB 1: Regional_Summary  (SUMIFS / XLOOKUP formulas + sensitivity grid)
# ═══════════════════════════════════════════════════════════════════════════════
ws1 = wb.active
ws1.title = "Regional_Summary"
ws1.sheet_properties.tabColor = "1F4E79"

# ── Write Raw_GL data to a helper sheet ──────────────────────────────────────
ws_raw = wb.create_sheet("Raw_GL")
ws_raw.sheet_properties.tabColor = "808080"
raw_headers = ["GL_Txn_ID", "Region", "Posting_Date", "Service_ID",
               "Revenue_USD", "OpEx_Cost_USD"]
for c, h in enumerate(raw_headers, 1):
    ws_raw.cell(row=1, column=c, value=h)
style_header_row(ws_raw, 1, len(raw_headers))

for i, row in gl.iterrows():
    r = i + 2
    ws_raw.cell(row=r, column=1, value=row["GL_Txn_ID"])
    ws_raw.cell(row=r, column=2, value=row["Region"])
    ws_raw.cell(row=r, column=3, value=str(row["Posting_Date"]))
    ws_raw.cell(row=r, column=4, value=row["Service_ID"])
    ws_raw.cell(row=r, column=5, value=row["Revenue_USD"])
    ws_raw.cell(row=r, column=6, value=row["OpEx_Cost_USD"])

# ── Section A: Service Line Master + SUMIFS formulas ─────────────────────────
headers_a = ["Service_ID", "Service_Line_Name", "Practice_Head",
             "Target_SLA_Hours", "Target_Margin_Pct",
             "Actual_Revenue", "Actual_OpEx", "EBITDA_Margin_%"]
for c, h in enumerate(headers_a, 1):
    ws1.cell(row=1, column=c, value=h)
style_header_row(ws1, 1, len(headers_a))

for i, row in sl.iterrows():
    r = i + 2  # rows 2-6
    ws1.cell(row=r, column=1, value=row["Service_ID"])
    ws1.cell(row=r, column=2, value=row["Service_Line_Name"])
    ws1.cell(row=r, column=3, value=row["Practice_Head"])
    ws1.cell(row=r, column=4, value=row["Target_SLA_Hours"])
    ws1.cell(row=r, column=5, value=row["Target_Margin_Pct"])
    ws1.cell(row=r, column=5).number_format = '0.00%'

    # SUMIFS formulas referencing Raw_GL sheet
    ws1.cell(row=r, column=6).value = f'=SUMIFS(Raw_GL!E:E,Raw_GL!D:D,A{r})'
    ws1.cell(row=r, column=6).number_format = '#,##0.00'

    ws1.cell(row=r, column=7).value = f'=SUMIFS(Raw_GL!F:F,Raw_GL!D:D,A{r})'
    ws1.cell(row=r, column=7).number_format = '#,##0.00'

    # EBITDA Margin % = (Revenue - OpEx) / Revenue
    ws1.cell(row=r, column=8).value = f'=IF(F{r}=0,0,(F{r}-G{r})/F{r})'
    ws1.cell(row=r, column=8).number_format = '0.00%'

apply_borders(ws1, 1, 6, len(headers_a))

# ── Section B: XLOOKUP Lookup Box ────────────────────────────────────────────
ws1.cell(row=8, column=9, value="XLOOKUP Lookup").font = Font(bold=True, size=12, color="1F4E79")
ws1.cell(row=9, column=9, value="Service ID →")
ws1.cell(row=9, column=10, value="SL-104")
ws1.cell(row=9, column=10).fill = YELLOW_FILL
ws1.cell(row=10, column=9, value="Target Margin →")
ws1.cell(row=10, column=10).value = '=XLOOKUP(J9,A2:A6,E2:E6,"Not Found")'
ws1.cell(row=10, column=10).number_format = '0.00%'
ws1.cell(row=10, column=10).fill = GREEN_FILL

# ── Section C: What-If Sensitivity Grid ──────────────────────────────────────
ws1.cell(row=8, column=1, value="═══ WHAT-IF SENSITIVITY ANALYSIS ═══").font = Font(bold=True, size=12, color="1F4E79")
ws1.merge_cells("A8:E8")

ws1.cell(row=9, column=1, value="Revenue Growth %")
ws1.cell(row=9, column=2, value=0).number_format = '0%'   # I5 equivalent → B9
ws1.cell(row=10, column=1, value="OpEx Reduction %")
ws1.cell(row=10, column=2, value=0).number_format = '0%'  # I6 equivalent → B10
ws1.cell(row=11, column=1, value="Projected EBITDA Margin")
ws1.cell(row=11, column=1).font = Font(bold=True)

# Projected EBITDA Margin formula
ws1.cell(row=11, column=2).value = '=((SUM(F2:F6)*(1+B9))-(SUM(G2:G6)*(1+B10)))/(SUM(F2:F6)*(1+B9))'
ws1.cell(row=11, column=2).number_format = '0.00%'
ws1.cell(row=11, column=2).fill = GREEN_FILL

# ── 2-Variable Sensitivity Grid Template ─────────────────────────────────────
# Row 13: headers
ws1.cell(row=13, column=1, value="EBITDA Margin Scenarios").font = Font(bold=True, size=11, color="1F4E79")
ws1.merge_cells("A13:F13")
ws1.cell(row=14, column=1, value="Rev Growth ↓ \\ OpEx Δ →").font = Font(bold=True, size=9)
ws1.cell(row=14, column=1).fill = HEADER_FILL
ws1.cell(row=14, column=1).font = HEADER_FONT

opex_deltas = [-0.02, -0.05, -0.08, -0.10, -0.15]
rev_growths = [0.0, 0.05, 0.10, 0.15, 0.20, 0.25]

for j, opex in enumerate(opex_deltas):
    col = j + 2
    cell = ws1.cell(row=14, column=col, value=opex)
    cell.number_format = '0%'
    cell.fill = HEADER_FILL
    cell.font = HEADER_FONT
    cell.alignment = CENTER

for i, rev_g in enumerate(rev_growths):
    row = 15 + i
    cell_rg = ws1.cell(row=row, column=1, value=rev_g)
    cell_rg.number_format = '0%'
    cell_rg.fill = ACCENT_FILL
    cell_rg.font = Font(bold=True)
    cell_rg.alignment = CENTER

    for j, opex in enumerate(opex_deltas):
        col = j + 2
        # Inline formula: ((sum_rev*(1+rev_g)) - (sum_opex*(1+opex))) / (sum_rev*(1+rev_g))
        formula = (
            f'=((SUM(F$2:F$6)*(1+$A{row}))'
            f'-(SUM(G$2:G$6)*(1+{get_column_letter(col)}$14)))'
            f'/(SUM(F$2:F$6)*(1+$A{row}))'
        )
        cell = ws1.cell(row=row, column=col, value=formula)
        cell.number_format = '0.00%'
        cell.alignment = CENTER

apply_borders(ws1, 14, 14 + len(rev_growths), 1 + len(opex_deltas))

# Column widths
for col, w in {1: 22, 2: 20, 3: 16, 4: 16, 5: 16, 6: 18, 7: 18, 8: 16, 9: 18, 10: 16}.items():
    ws1.column_dimensions[get_column_letter(col)].width = w

# ═══════════════════════════════════════════════════════════════════════════════
#  TAB 3: Data_Dictionary_Governance
# ═══════════════════════════════════════════════════════════════════════════════
ws3 = wb.create_sheet("Data_Dictionary_Governance")
ws3.sheet_properties.tabColor = "2E7D32"

dict_headers = ["Table_Name", "Field_Name", "Data_Type",
                "Primary_or_Foreign_Key", "Power_Query_Governance_Rule"]
for c, h in enumerate(dict_headers, 1):
    ws3.cell(row=1, column=c, value=h)
style_header_row(ws3, 1, len(dict_headers))

# Complete data dictionary entries
dict_data = [
    # service_line_master
    ("service_line_master", "Service_ID", "Text (SL-NNN)", "PRIMARY KEY",
     "No cleaning needed. Used as join key to all fact tables."),
    ("service_line_master", "Service_Line_Name", "Text", "—",
     "PQ: Transform → Trim → Capitalize Each Word. Removes leading/trailing spaces and normalizes casing."),
    ("service_line_master", "Practice_Head", "Text", "—",
     "No cleaning needed."),
    ("service_line_master", "Target_SLA_Hours", "Decimal (hours)", "—",
     "No cleaning needed. Used to evaluate SLA Met/Breached."),
    ("service_line_master", "Target_Margin_Pct", "Decimal (0-1)", "—",
     "No cleaning needed. Reference benchmark for EBITDA margin analysis."),

    # regional_budgets_wide → long
    ("regional_budgets_wide", "Region", "Text", "FOREIGN KEY → (implicit)",
     "No cleaning needed. Identifies geographic business unit."),
    ("regional_budgets_wide", "Service_ID", "Text (SL-NNN)", "FOREIGN KEY → service_line_master",
     "No cleaning needed. Join key to dimension table."),
    ("regional_budgets_wide", "[Month columns]", "Decimal (wide)", "—",
     "PQ: Select Region + Service_ID → Right-click → Unpivot Other Columns. Rename Attribute → Month_Str, Value → Budget_Revenue_USD."),
    ("regional_budgets_long", "Month_Str", "Text (YYYY-MM)", "—",
     "PQ: Add Column → Custom Column: Date.FromText([Month_Str] & \"-01\"). Change type to Date."),
    ("regional_budgets_long", "Budget_Date", "Date", "FOREIGN KEY → Dim_Calendar",
     "Derived from Month_Str. Links to calendar dimension for time intelligence DAX."),
    ("regional_budgets_long", "Budget_Revenue_USD", "Decimal", "—",
     "DAX: Total Budget Revenue = SUM(regional_budgets_wide[Budget_Revenue_USD])"),

    # actual_gl_financials
    ("actual_gl_financials", "GL_Txn_ID", "Text (GL-NNNNN)", "PRIMARY KEY",
     "No cleaning needed. Unique transaction identifier."),
    ("actual_gl_financials", "Region", "Text", "FOREIGN KEY → (implicit)",
     "No cleaning needed."),
    ("actual_gl_financials", "Posting_Date", "Date", "FOREIGN KEY → Dim_Calendar",
     "PQ: Change type to Date after cleaning Revenue_USD."),
    ("actual_gl_financials", "Service_ID", "Text (SL-NNN)", "FOREIGN KEY → service_line_master",
     "No cleaning needed. Join key to dimension table."),
    ("actual_gl_financials", "Revenue_USD", "Decimal", "—",
     "PQ: Replace Values '$' → blank, Replace ',' → blank, Change type to Fixed Decimal. DAX: Total Actual Revenue = SUM(actual_gl_financials[Revenue_USD])"),
    ("actual_gl_financials", "OpEx_Cost_USD", "Decimal", "—",
     "No cleaning needed. DAX: EBITDA Margin % = DIVIDE([Total Actual Revenue] - SUM(OpEx_Cost_USD), [Total Actual Revenue], 0)"),

    # it_sla_tickets_log
    ("it_sla_tickets_log", "Ticket_ID", "Text (TKT-NNNNN)", "PRIMARY KEY",
     "PQ: Home → Remove Rows → Remove Duplicates. Drops 45 injected duplicates (3045 → 3000)."),
    ("it_sla_tickets_log", "Region", "Text", "FOREIGN KEY → (implicit)",
     "No cleaning needed."),
    ("it_sla_tickets_log", "Service_ID", "Text (SL-NNN)", "FOREIGN KEY → service_line_master",
     "No cleaning needed. Join key."),
    ("it_sla_tickets_log", "Incident_Date", "Date", "FOREIGN KEY → Dim_Calendar",
     "PQ: Change type to Date. Links to calendar dimension."),
    ("it_sla_tickets_log", "SLA_Target_Hours", "Decimal (hours)", "—",
     "No cleaning needed. Lookup from service_line_master."),
    ("it_sla_tickets_log", "Resolution_Hours", "Decimal (hours)", "—",
     "No cleaning needed. Used for SLA_Status derivation and IQR outlier detection."),
    ("it_sla_tickets_log", "SLA_Status", "Text (Met/Breached)", "—",
     "PQ: Add Column → Conditional Column. If Resolution_Hours ≤ SLA_Target_Hours → 'Met', Else → 'Breached'. DAX: SLA Compliance Rate % = DIVIDE(CALCULATE(COUNTROWS(...), SLA_Status=\"Met\"), COUNTROWS(...), 0)"),

    # Dim_Calendar (DAX-generated)
    ("Dim_Calendar", "Date", "Date", "PRIMARY KEY",
     "DAX Table: Dim_Calendar = CALENDAR(DATE(2025,1,1), DATE(2026,12,31)). Links 1:* to Budget_Date, Posting_Date, Incident_Date."),

    # DAX Measures
    ("_Measures", "Total Actual Revenue", "DAX Measure", "—",
     "= SUM(actual_gl_financials[Revenue_USD])"),
    ("_Measures", "Total Budget Revenue", "DAX Measure", "—",
     "= SUM(regional_budgets_wide[Budget_Revenue_USD])"),
    ("_Measures", "Budget Variance $", "DAX Measure", "—",
     "= [Total Actual Revenue] - [Total Budget Revenue]"),
    ("_Measures", "Budget Variance %", "DAX Measure", "—",
     "= DIVIDE([Budget Variance $], [Total Budget Revenue], 0)"),
    ("_Measures", "EBITDA Margin %", "DAX Measure", "—",
     "= DIVIDE([Total Actual Revenue] - SUM(actual_gl_financials[OpEx_Cost_USD]), [Total Actual Revenue], 0)"),
    ("_Measures", "YoY Revenue Growth %", "DAX Measure", "—",
     "= VAR PrevYearRev = CALCULATE([Total Actual Revenue], SAMEPERIODLASTYEAR(Dim_Calendar[Date])) RETURN DIVIDE([Total Actual Revenue] - PrevYearRev, PrevYearRev, 0)"),
    ("_Measures", "SLA Compliance Rate %", "DAX Measure", "—",
     "= DIVIDE(CALCULATE(COUNTROWS(it_sla_tickets_log), it_sla_tickets_log[SLA_Status]=\"Met\"), COUNTROWS(it_sla_tickets_log), 0)"),
]

for i, (table, field, dtype, key, rule) in enumerate(dict_data, 2):
    ws3.cell(row=i, column=1, value=table)
    ws3.cell(row=i, column=2, value=field)
    ws3.cell(row=i, column=3, value=dtype)
    ws3.cell(row=i, column=4, value=key)
    ws3.cell(row=i, column=5, value=rule)

    # Alternate row shading
    if i % 2 == 0:
        for c in range(1, 6):
            ws3.cell(row=i, column=c).fill = ACCENT_FILL

    for c in range(1, 6):
        ws3.cell(row=i, column=c).alignment = LEFT_WRAP
        ws3.cell(row=i, column=c).border = THIN_BORDER

# Column widths for readability
for col, w in {1: 26, 2: 28, 3: 22, 4: 30, 5: 80}.items():
    ws3.column_dimensions[get_column_letter(col)].width = w

ws3.freeze_panes = "A2"

# ── Save ─────────────────────────────────────────────────────────────────────
wb.save(OUT_FILE)
print(f"✅ {OUT_FILE.name} saved with tabs:")
print(f"   • Regional_Summary  (SUMIFS + XLOOKUP + Sensitivity Grid)")
print(f"   • Raw_GL            (GL data backing SUMIFS)")
print(f"   • Data_Dictionary_Governance ({len(dict_data)} entries)")
