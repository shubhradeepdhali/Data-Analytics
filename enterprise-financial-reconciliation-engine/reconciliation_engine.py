import sqlite3
import pandas as pd
import numpy as np
from datetime import datetime, timedelta

np.random.seed(42)

def seed_financial_database(conn):
    """Seeds realistic IT Infrastructure invoices, payments, and regional budgets with deliberate anomalies."""
    regions = ['North America', 'EMEA', 'APAC - India', 'LATAM']
    service_lines = [
        'Network Integration', 'Unified Communications', 
        'Data Center Infrastructure', 'Cybersecurity Managed SLA', 
        'Cloud Edge Deployment'
    ]
    
    # 1. Create Regional Cost Center Budgets
    budgets = []
    for region in regions:
        for service in service_lines:
            allocated = np.random.randint(120000, 350000)
            budgets.append({
                'region': region,
                'service_line': service,
                'monthly_budget_usd': float(allocated)
            })
    df_budgets = pd.DataFrame(budgets)
    df_budgets.to_sql('cost_center_budgets', conn, if_exists='replace', index=False)

    # 2. Generate 1,500 Enterprise Client/Vendor Invoices
    invoices = []
    payments = []
    base_date = datetime(2026, 10, 1)

    for i in range(1, 1501):
        inv_id = f"INV-2026-{i:04d}"
        region = np.random.choice(regions, p=[0.4, 0.25, 0.25, 0.1])
        service = np.random.choice(service_lines)
        inv_date = base_date + timedelta(days=int(np.random.randint(0, 28)))
        amount = round(float(np.random.uniform(1500, 18000)), 2)

        invoices.append({
            'invoice_id': inv_id,
            'client_code': f"CLT-{np.random.randint(100, 250)}",
            'region': region,
            'service_line': service,
            'invoice_date': inv_date.strftime('%Y-%m-%d'),
            'invoice_amount_usd': amount
        })

        # Simulate payment behavior (90% exact match, 10% financial anomalies)
        anomaly_roll = np.random.rand()
        if anomaly_roll < 0.88:
            # Clean match
            payments.append({
                'payment_id': f"PAY-{i:05d}",
                'invoice_id': inv_id,
                'payment_date': (inv_date + timedelta(days=int(np.random.randint(2, 15)))).strftime('%Y-%m-%d'),
                'paid_amount_usd': amount,
                'payment_method': np.random.choice(['ACH Wire', 'Corporate EFT', 'SWIFT'])
            })
        elif anomaly_roll < 0.93:
            # Short / Partial Payment Discrepancy
            short_amt = round(amount * np.random.uniform(0.65, 0.92), 2)
            payments.append({
                'payment_id': f"PAY-{i:05d}",
                'invoice_id': inv_id,
                'payment_date': (inv_date + timedelta(days=10)).strftime('%Y-%m-%d'),
                'paid_amount_usd': short_amt,
                'payment_method': 'ACH Wire'
            })
        elif anomaly_roll < 0.96:
            # Duplicate Payment Error (Paid twice!)
            for dup in ['A', 'B']:
                payments.append({
                    'payment_id': f"PAY-{i:05d}-{dup}",
                    'invoice_id': inv_id,
                    'payment_date': (inv_date + timedelta(days=7)).strftime('%Y-%m-%d'),
                    'paid_amount_usd': amount,
                    'payment_method': 'Corporate EFT'
                })
        else:
            # Unpaid / Missing Ledger Entry (No payment record generated)
            pass

    pd.DataFrame(invoices).to_sql('invoices_ledger', conn, if_exists='replace', index=False)
    pd.DataFrame(payments).to_sql('payments_ledger', conn, if_exists='replace', index=False)
    print(f"[INFO] Seeded 1,500 invoices and {len(payments)} payment records into SQLite.")


def run_sql_reconciliation(conn):
    """Executes multi-CTE SQL reconciliation and budget variance queries."""
    reconciliation_sql = """
    WITH AggregatedPayments AS (
        SELECT 
            invoice_id,
            COUNT(payment_id) AS payment_txn_count,
            ROUND(SUM(paid_amount_usd), 2) AS total_paid_usd,
            MAX(payment_date) AS last_payment_date
        FROM payments_ledger
        GROUP BY invoice_id
    ),
    InvoiceReconciliation AS (
        SELECT 
            i.invoice_id,
            i.client_code,
            i.region,
            i.service_line,
            i.invoice_date,
            i.invoice_amount_usd,
            COALESCE(p.total_paid_usd, 0.0) AS total_paid_usd,
            COALESCE(p.payment_txn_count, 0) AS payment_txn_count,
            ROUND(i.invoice_amount_usd - COALESCE(p.total_paid_usd, 0.0), 2) AS variance_usd,
            CASE 
                WHEN p.invoice_id IS NULL THEN 'CRITICAL: Missing Payment'
                WHEN p.payment_txn_count > 1 THEN 'ALERT: Duplicate Payment'
                WHEN ABS(i.invoice_amount_usd - p.total_paid_usd) > 0.05 THEN 'WARNING: Partial/Short Payment'
                ELSE 'Reconciled'
            END AS audit_status
        FROM invoices_ledger i
        LEFT JOIN AggregatedPayments p ON i.invoice_id = p.invoice_id
    )
    SELECT * FROM InvoiceReconciliation
    ORDER BY ABS(variance_usd) DESC;
    """
    df_reconciled = pd.read_sql_query(reconciliation_sql, conn)

    budget_variance_sql = """
    WITH ActualSpend AS (
        SELECT 
            region,
            service_line,
            COUNT(invoice_id) AS total_invoices,
            ROUND(SUM(invoice_amount_usd), 2) AS actual_billed_usd
        FROM invoices_ledger
        GROUP BY region, service_line
    )
    SELECT 
        b.region AS Region,
        b.service_line AS Service_Line,
        a.total_invoices AS Invoice_Count,
        b.monthly_budget_usd AS Monthly_Budget_USD,
        a.actual_billed_usd AS Actual_Billed_USD,
        ROUND(a.actual_billed_usd - b.monthly_budget_usd, 2) AS Variance_USD,
        ROUND(((a.actual_billed_usd - b.monthly_budget_usd) / b.monthly_budget_usd) * 100, 2) AS Variance_Pct,
        CASE 
            WHEN ((a.actual_billed_usd - b.monthly_budget_usd) / b.monthly_budget_usd) > 0.10 THEN 'OVER BUDGET (>10%)'
            WHEN ((a.actual_billed_usd - b.monthly_budget_usd) / b.monthly_budget_usd) < -0.10 THEN 'UNDER UTILIZED (<-10%)'
            ELSE 'Within Threshold'
        END AS Governance_Flag
    FROM cost_center_budgets b
    LEFT JOIN ActualSpend a ON b.region = a.region AND b.service_line = a.service_line
    ORDER BY Variance_Pct DESC;
    """
    df_budget_var = pd.read_sql_query(budget_variance_sql, conn)
    return df_reconciled, df_budget_var


def export_staging_workbook(df_reconciled, df_budget_var, output_path):
    """Exports multi-sheet Excel staging workbook ready for VBA automation."""
    df_exceptions = df_reconciled[df_reconciled['audit_status'] != 'Reconciled'].copy()

    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        df_budget_var.to_excel(writer, sheet_name='Executive_Budget_Variance', index=False)
        df_exceptions.to_excel(writer, sheet_name='Audit_Exception_Log', index=False)
        df_reconciled.to_excel(writer, sheet_name='Full_Reconciled_Ledger', index=False)

    print(f"[SUCCESS] Staging workbook generated: {output_path}")
    print(f" -> Total Invoices Processed: {len(df_reconciled)}")
    print(f" -> Audit Exceptions Flagged: {len(df_exceptions)} ({round(len(df_exceptions)/len(df_reconciled)*100, 1)}%)")


from pathlib import Path

if __name__ == "__main__":
    # Always resolve paths relative to the folder where this script lives
    BASE_DIR = Path(__file__).resolve().parent
    db_path = BASE_DIR / "enterprise_finance.db"
    excel_path = BASE_DIR / "Month_End_Financial_Reconciliation.xlsx"

    conn = sqlite3.connect(db_path)
    seed_financial_database(conn)
    df_rec, df_bud = run_sql_reconciliation(conn)
    export_staging_workbook(df_rec, df_bud, excel_path)
    conn.close()