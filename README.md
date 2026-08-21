# Whitmore Sleeve KPI Streamlit Dashboard

## Files

- `app.py` — complete Streamlit dashboard.
- `requirements.txt` — Python dependencies.
- `Whitmore_Sleeve_KPI_Workings.xlsx` — supplied source workbook.

## Run locally

1. Put the three files in the same folder.
2. Create/activate a Python virtual environment.
3. Install dependencies:

```bash
pip install -r requirements.txt
```

4. Start Streamlit:

```bash
streamlit run app.py
```

The dashboard automatically loads `Whitmore_Sleeve_KPI_Workings.xlsx` if it is beside `app.py`. You can also upload the workbook from the sidebar.

## Dashboard features

- KPI cards for gross income, annualized gross yield, shown leverage cost and concentration headroom.
- Filters for period, asset class, and obligor/counterparty where available.
- Loan-level drill-down with funded exposure, cash interest, recognized/accrued interest and recognition-vs-cash gap.
- Hedge analysis for IRS, CDS, FX forwards and options.
- Explicit recognition-vs-cash timing table.
- Concentration scenario analysis against the 12.5% limit.
- Visible assumptions / judgment calls from the workbook.
- Raw workbook-tab audit view with CSV export.
- Defensive parsing and missing-sheet handling so the app fails gracefully rather than crashing.

## Important workbook behavior

The source workbook itself identifies several genuine data gaps and judgment calls. The dashboard intentionally surfaces them rather than estimating them. In particular, the Q4 fund-level Greystone revolver interest cost is not isolated in the supplied data, so the dashboard does not manufacture a Q4 figure.
