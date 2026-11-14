# Enterprise Financial Reconciliation Engine

A lightweight Python-based financial reconciliation and audit engine that generates realistic enterprise finance data, reconciles invoice and payment records in SQLite, flags exceptions, and exports a staged Excel workbook for downstream month-end reporting.

## Overview

This project simulates a month-end reconciliation workflow for an IT infrastructure business across multiple regions and service lines. It does the following:

- Seeds a SQLite database with realistic regional budgets, invoices, and payments
- Identifies reconciliation variances between invoice totals and payment totals
- Detects duplicate payments, short/partial payments, and missing-payment conditions
- Compares actual billed spend to budget allocations by region/service line
- Exports a multi-sheet Excel workbook for finance review and automation

## Business Use Case

The workflow mirrors common month-end finance controls used by enterprise teams:

- Invoice-to-payment matching
- Budget versus actual variance analysis
- Governance flagging for over/under budget conditions
- Exception logging for manual investigation

## Project Structure

- `reconciliation_engine.py` — main script that generates the data, runs reconciliation logic, and exports Excel output
- `enterprise_finance.db` — SQLite database created and populated by the script
- `Month_End_Financial_Reconciliation.xlsx` — generated staging workbook
- `Executive_Month_End_Variance_Report.pdf` — sample finance report artifact
- `.venv/` — local Python virtual environment for dependencies

## Prerequisites

- Python 3.9+
- Access to the included virtual environment or installation of the required packages

Required Python libraries:

- `pandas`
- `numpy`
- `openpyxl`

## Setup

From the project root, activate the virtual environment and run the script:

```bash
# Windows PowerShell
.\.venv\Scripts\Activate.ps1
python reconciliation_engine.py
```

If you are not using the bundled environment, install dependencies first:

```bash
pip install pandas numpy openpyxl
```

## How It Works

The script performs the following in order:

1. Creates or replaces a SQLite database named `enterprise_finance.db`
2. Seeds synthetic data for:
   - regional budget tables
   - invoice ledger entries
   - payment ledger entries with deliberate anomalies
3. Runs reconciliation queries to compute:
   - invoice variance vs. payment totals
   - audit statuses
   - budget variance and governance flags
4. Exports an Excel workbook with these sheets:
   - `Executive_Budget_Variance`
   - `Audit_Exception_Log`
   - `Full_Reconciled_Ledger`

## Output Files

After execution, the script writes:

- `Month_End_Financial_Reconciliation.xlsx`
- `enterprise_finance.db`

These are the primary assets for a month-end finance review workflow.

## Notes

This project intentionally contains synthetic data and anomaly patterns to demonstrate how a reconciliation engine can detect issues without requiring production financial data.

## License

This project is provided as a local working example for financial reconciliation workflows and is not intended for production use without review and validation.
