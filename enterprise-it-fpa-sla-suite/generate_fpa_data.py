"""
generate_fpa_data.py
====================
Generates 4 realistic raw CSV files with intentional data flaws for the
Enterprise IT FP&A & SLA Performance Suite.

Outputs (inside raw_data/):
  1. service_line_master.csv   – Dimension table (whitespace & casing flaws)
  2. regional_budgets_wide.csv – Wide-format monthly budgets (needs unpivot)
  3. actual_gl_financials.csv  – GL revenue & OpEx ($ signs, commas in 15% rows)
  4. it_sla_tickets_log.csv    – SLA tickets (45 duplicates appended)
"""

from pathlib import Path
import pandas as pd
import numpy as np

np.random.seed(101)

# ── Anchor all paths to the folder where this script lives ──────────────────
BASE_DIR = Path(__file__).resolve().parent
RAW_DIR = BASE_DIR / "raw_data"
RAW_DIR.mkdir(exist_ok=True)

# ═══════════════════════════════════════════════════════════════════════════════
# 1. Dimension Table: Service Line Master
#    (intentional whitespace & inconsistent casing for Power Query cleaning)
# ═══════════════════════════════════════════════════════════════════════════════
service_lines = pd.DataFrame({
    "Service_ID": ["SL-101", "SL-102", "SL-103", "SL-104", "SL-105"],
    "Service_Line_Name": [
        "  Network Integration ", "unified communications",
        "Data Center Infrastructure  ", "Cybersecurity Managed SLA",
        "cloud edge deployment "
    ],
    "Practice_Head": ["R. Sharma", "M. Vance", "K. Nair", "D. Miller", "S. Kulkarni"],
    "Target_SLA_Hours": [4.0, 2.0, 6.0, 1.0, 4.0],
    "Target_Margin_Pct": [0.28, 0.34, 0.25, 0.42, 0.31]
})
service_lines.to_csv(RAW_DIR / "service_line_master.csv", index=False)
print("[✓] service_line_master.csv  → 5 rows")

# ═══════════════════════════════════════════════════════════════════════════════
# 2. Fact Table 1: Wide Regional Budgets (requires Unpivot in Power Query)
#    4 Regions × 5 Services × 12 monthly columns → 20 rows, wide format
# ═══════════════════════════════════════════════════════════════════════════════
regions = ["North America", "EMEA", "APAC - India", "LATAM"]
months = [
    "2025-01", "2025-02", "2025-03", "2025-04", "2025-05", "2025-06",
    "2026-01", "2026-02", "2026-03", "2026-04", "2026-05", "2026-06"
]

budget_rows = []
for reg in regions:
    for sid in service_lines["Service_ID"]:
        row = {"Region": reg, "Service_ID": sid}
        base = np.random.randint(85000, 220000)
        for m in months:
            growth = 1.08 if m.startswith("2026") else 1.0
            row[m] = round(base * growth * np.random.uniform(0.95, 1.05), 2)
        budget_rows.append(row)

pd.DataFrame(budget_rows).to_csv(RAW_DIR / "regional_budgets_wide.csv", index=False)
print("[✓] regional_budgets_wide.csv → 20 rows × 14 columns (wide)")

# ═══════════════════════════════════════════════════════════════════════════════
# 3. Fact Table 2: Actual General Ledger (GL) Revenue & Operating Costs
#    15% of Revenue_USD values injected with "$" and "," for dirty-text cleaning
# ═══════════════════════════════════════════════════════════════════════════════
gl_rows = []
txn_id = 50000
for reg in regions:
    for sid in service_lines["Service_ID"]:
        for m in months:
            for _ in range(12):
                txn_id += 1
                day = np.random.randint(1, 28)
                rev = round(np.random.uniform(8000, 24000), 2)
                cost = round(rev * np.random.uniform(0.58, 0.82), 2)
                # Inject dirty "$" + comma formatting in ~15% of rows
                rev_str = f"${rev:,.2f}" if np.random.rand() < 0.15 else str(rev)
                gl_rows.append({
                    "GL_Txn_ID": f"GL-{txn_id}",
                    "Region": reg,
                    "Posting_Date": f"{m}-{day:02d}",
                    "Service_ID": sid,
                    "Revenue_USD": rev_str,
                    "OpEx_Cost_USD": cost
                })

pd.DataFrame(gl_rows).to_csv(RAW_DIR / "actual_gl_financials.csv", index=False)
print(f"[✓] actual_gl_financials.csv → {len(gl_rows)} rows")

# ═══════════════════════════════════════════════════════════════════════════════
# 4. Fact Table 3: IT SLA Tickets Log
#    ~86% Met / ~14% Breached, plus 45 duplicate rows appended
# ═══════════════════════════════════════════════════════════════════════════════
sla_target_map = dict(zip(service_lines["Service_ID"],
                          service_lines["Target_SLA_Hours"]))

ticket_rows = []
ticket_id = 10000
for _ in range(3000):
    ticket_id += 1
    sid = np.random.choice(service_lines["Service_ID"].values)
    reg = np.random.choice(regions)
    m = np.random.choice(months)
    day = np.random.randint(1, 28)
    sla_target = sla_target_map[sid]

    if np.random.rand() < 0.86:
        resolution = round(sla_target * np.random.uniform(0.3, 0.95), 2)
    else:
        resolution = round(sla_target * np.random.uniform(1.1, 2.4), 2)

    ticket_rows.append({
        "Ticket_ID": f"TKT-{ticket_id}",
        "Region": reg,
        "Service_ID": sid,
        "Incident_Date": f"{m}-{day:02d}",
        "SLA_Target_Hours": sla_target,
        "Resolution_Hours": resolution
    })

df_sla = pd.DataFrame(ticket_rows)
# Inject 45 duplicate rows (intentional flaw for deduplication exercise)
df_sla = pd.concat([df_sla, df_sla.sample(45, random_state=42)], ignore_index=True)
df_sla.to_csv(RAW_DIR / "it_sla_tickets_log.csv", index=False)
print(f"[✓] it_sla_tickets_log.csv   → {len(df_sla)} rows (45 duplicates)")

print(f"\n✅ All 4 CSVs saved to: {RAW_DIR}")
