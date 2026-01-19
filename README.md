# Reporting automation (Excel + Power Query + VBA)

# Reporting Automation — P&L & Balance Sheet

**Goal:** Automate monthly reporting refresh and consolidation for P&L and Balance Sheet.

## Highlights
- **Power Query:** Data ingestion, cleaning, and model refresh.
- **Power Pivot:** Data model with relationships for IS/BS.
- **VBA Macros:** One-click refresh and export to reporting packs.
- **Outcome:** Reduced manual effort and ensured consistent, audit-ready outputs.

## Files
- `reporting_automation.xlsm` — macro-enabled workbook
- `screenshots/automation_overview.png` — process flow and output sample

## How to use
1. Open the workbook and enable macros.
2. Update the sample data in the `Data` sheet.
3. Click `Refresh All` to update queries and pivots.
4. Use the `Export` macro to generate reporting outputs.

## Data notes
- Sample data is anonymized and synthetic.
- Replace with your own dataset in the `Data` sheet.

## KPIs & Outputs
- P&L summary, BS summary, period-over-period variance
- Exported reporting pack (PDF/Excel) via macro
