import io
import os
from pathlib import Path

import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st


st.set_page_config(
    page_title="Whitmore Sleeve KPI Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

DEFAULT_FILE = Path(__file__).with_name("Whitmore_Sleeve_KPI_Workings.xlsx")


# -----------------------------
# Helpers
# -----------------------------
def money(x, decimals=0):
    if x is None or (isinstance(x, float) and np.isnan(x)):
        return "—"
    try:
        x = float(x)
        sign = "-" if x < 0 else ""
        return f"{sign}${abs(x):,.{decimals}f}"
    except Exception:
        return "—"


def pct(x, decimals=1):
    if x is None or (isinstance(x, float) and np.isnan(x)):
        return "—"
    try:
        return f"{float(x) * 100:.{decimals}f}%"
    except Exception:
        return "—"


def clean_df(df):
    if df is None:
        return pd.DataFrame()
    df = df.copy()
    df = df.dropna(axis=0, how="all").dropna(axis=1, how="all")
    df.columns = [str(c).strip() if c is not None else "" for c in df.columns]
    return df


def read_workbook(source):
    """Read workbook defensively. Returns {sheet_name: dataframe}."""
    try:
        if hasattr(source, "seek"):
            source.seek(0)
        book = pd.ExcelFile(source, engine="openpyxl")
        sheets = {}
        for sheet in book.sheet_names:
            try:
                df = pd.read_excel(book, sheet_name=sheet, engine="openpyxl")
                sheets[sheet] = clean_df(df)
            except Exception:
                sheets[sheet] = pd.DataFrame()
        return sheets, None
    except Exception as exc:
        return {}, f"Could not read the Excel workbook: {exc}"


def find_row(df, first_col_value):
    if df.empty:
        return None
    first = df.iloc[:, 0].astype(str).str.strip()
    hits = df.index[first.eq(str(first_col_value).strip())]
    return hits[0] if len(hits) else None


def value_after_label(df, label, value_col=1):
    """Get the value in the next column for a label in the first column."""
    row = find_row(df, label)
    if row is None or df.shape[1] <= value_col:
        return np.nan
    return df.loc[row].iloc[value_col]


def get_sheet(sheets, name):
    return sheets.get(name, pd.DataFrame())


def loan_data(sheets):
    df = get_sheet(sheets, "Loan_Book")
    if df.empty or "Loan ID" not in df.columns:
        return pd.DataFrame()
    df = df[df["Loan ID"].notna()].copy()
    df = df[~df["Loan ID"].astype(str).str.startswith("TOTAL", na=False)]
    if "Include?" in df.columns:
        df["Included"] = df["Include?"].astype(str).str.upper().eq("Y")
    else:
        df["Included"] = True
    for c in ["Funded Bal. (USD)", "Q4 Cash Int. Recv'd (USD)",
              "Q4 Accrued Int. Income (USD)", "Recog. vs Cash Gap", "DPD"]:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)
    df["Period"] = "Q4-2025"
    df["Asset Class"] = "Loans"
    return df[df["Included"]].copy()


def cash_data(sheets):
    df = get_sheet(sheets, "Cash_Recon")
    if df.empty or "Description" not in df.columns:
        return pd.DataFrame()
    df = df[df["Description"].notna()].copy()
    # Exclude workbook total row.
    df = df[~df["Description"].astype(str).str.startswith("NET CASH MOVEMENT", na=False)]
    if "Value Date" in df.columns:
        df["Value Date"] = pd.to_datetime(df["Value Date"], errors="coerce")
    if "Amount (USD)" in df.columns:
        df["Amount (USD)"] = pd.to_numeric(df["Amount (USD)"], errors="coerce").fillna(0)
    df["Period"] = "Q4-2025"
    df["Asset Class"] = "Cash / Other"
    return df


def derivative_tables(sheets):
    result = {}

    cds = get_sheet(sheets, "CDS")
    if not cds.empty and "Trade Ref" in cds.columns:
        cds = cds[cds["Trade Ref"].notna()].copy()
        for c in ["Notional (USD)", "Fair Value 31-Dec-25 (USD)", "Q4 Premium Cash (USD)"]:
            if c in cds.columns:
                cds[c] = pd.to_numeric(cds[c], errors="coerce").fillna(0)
        cds["Asset Class"] = "CDS"
        cds["Period"] = "Q4-2025"
        result["CDS"] = cds

    fx = get_sheet(sheets, "FX_Forwards")
    if not fx.empty and "Client Ref" in fx.columns:
        fx = fx[fx["Client Ref"].notna()].copy()
        for c in ["Notional", "Contract Rate", "Fixing/Current Rate", "Settlement / MTM P&L (USD)"]:
            if c in fx.columns:
                fx[c] = pd.to_numeric(fx[c], errors="coerce").fillna(0)
        fx["Asset Class"] = "FX Forwards"
        fx["Period"] = "Q4-2025"
        result["FX Forwards"] = fx

    opts = get_sheet(sheets, "Options")
    if not opts.empty and "Ticker" in opts.columns:
        # Locate the open-position header and keep only actual position rows.
        opts = opts[opts["Ticker"].notna()].copy()
        for c in ["Contracts", "Strike", "Fair Value (USD)"]:
            if c in opts.columns:
                opts[c] = pd.to_numeric(opts[c], errors="coerce").fillna(0)
        opts["Asset Class"] = "Options"
        opts["Period"] = "Q4-2025"
        result["Options"] = opts

    return result


def assumptions_table(sheets):
    df = get_sheet(sheets, "Cover & Assumptions")
    if df.empty:
        return pd.DataFrame()
    # The assumption table has # / Open Item / Assumption... headers.
    header_idx = None
    for i in range(min(len(df), 15)):
        row = [str(x).strip() for x in df.iloc[i].tolist()]
        if row and row[0] == "#" and "Open Item" in row:
            header_idx = i
            break
    if header_idx is None:
        return pd.DataFrame()
    out = df.iloc[header_idx + 1:].copy()
    out.columns = df.iloc[header_idx].tolist()
    out = out[out["#"].notna()]
    return clean_df(out)


def kpi_values(sheets):
    kpi = get_sheet(sheets, "KPI_Summary")
    vals = {}
    if kpi.empty:
        return vals
    labels = [
        "Loan book — cash interest received, Q4",
        "Loan book — accrued interest income, Q4 (accrual basis)",
        "Content-library sub-licence royalty income, Q4 (per Assumption A4)",
        "Less: content-library lease cost, Q4 (Longwood + Nishimura, per A4)",
        "Less: CDS net premium paid/received, Q4",
        "Gross sleeve income, Q4 (accrual basis)",
        "Average funded loan balance (proxy for invested capital)",
        "Gross annualized yield on funded loan balance (accrual basis)",
        "Total fund-level leverage cost shown (partial-period, see A13)",
        "Net sleeve income after shown leverage cost (partial-period)",
        "Orion Studios Holdings LLC exposure, mezz at recovery + CDS netted (recommended, A5)",
        "Greystone 12.5% single-name limit (USD)",
        "Headroom / (breach)",
        "Q4 net cash movement, all activity (Cash_Recon tab)",
        "Net LP capital called, inception-to-date (LP_Capital tab)",
        "Total LP distributions, inception-to-date (LP_Capital tab)",
    ]
    for label in labels:
        vals[label] = value_after_label(kpi, label)
    return vals


# -----------------------------
# Load workbook
# -----------------------------
st.title("Whitmore Structured Credit Opportunities Fund II")
st.caption("Structured Debt & Derivatives Sleeve • Q4-2025 KPI Dashboard")

with st.sidebar:
    st.header("Workbook")
    uploaded = st.file_uploader(
        "Upload the KPI workings Excel file",
        type=["xlsx"],
        help="The dashboard is designed for the supplied Whitmore_Sleeve_KPI_Workings.xlsx structure.",
    )
    source = uploaded if uploaded is not None else (DEFAULT_FILE if DEFAULT_FILE.exists() else None)

    if source is None:
        st.error("No workbook found. Upload the Excel file to continue.")
        st.stop()

    with st.spinner("Loading workbook…"):
        sheets, error = read_workbook(source)

    if error:
        st.error(error)
        st.stop()

    st.success(f"Loaded {len(sheets)} workbook tabs.")

    st.divider()
    st.header("Filters")

    # User-facing filters. The supplied workbook is Q4-2025, but the filter is
    # intentionally explicit so the app remains usable with a refreshed workbook.
    period_options = ["Q4-2025"]
    if get_sheet(sheets, "LP_Capital").shape[0] > 0:
        period_options.append("FY2025 / inception-to-date")
    selected_period = st.selectbox("Period", period_options)

    loans = loan_data(sheets)
    obligors = sorted(loans["Obligor"].dropna().astype(str).unique().tolist()) if not loans.empty else []
    selected_obligor = st.multiselect("Counterparty / obligor", obligors)

    asset_options = ["All", "Loans", "CDS", "FX Forwards", "Options", "Leases", "Cash / Other"]
    selected_assets = st.multiselect("Asset class", asset_options, default=["All"])
    if not selected_assets:
        selected_assets = ["All"]

# -----------------------------
# KPI cards
# -----------------------------
vals = kpi_values(sheets)
gross_yield = vals.get("Gross annualized yield on funded loan balance (accrual basis)")
gross_income = vals.get("Gross sleeve income, Q4 (accrual basis)")
net_income = vals.get("Net sleeve income after shown leverage cost (partial-period)")
headroom = vals.get("Headroom / (breach)")

cards = st.columns(4)
cards[0].metric("Gross sleeve income", money(gross_income))
cards[1].metric("Gross annualized yield", pct(gross_yield))
cards[2].metric("After shown leverage cost", money(net_income))
cards[3].metric("Orion concentration headroom", money(headroom))

st.info(
    "Recognition and cash are deliberately shown separately. "
    "The workbook contains timing/recognition judgments (for example PIK income, "
    "accrued Orion mezz interest, Longwood true-up, and the IRS duplicate entry); "
    "these are surfaced below rather than silently embedded in the final KPI."
)

# -----------------------------
# Main tabs
# -----------------------------
tab_overview, tab_loans, tab_hedges, tab_recognition, tab_concentration, tab_adjustments, tab_raw = st.tabs(
    ["Overview", "Loan Book", "Hedge Book", "Recognition vs Cash", "Concentration", "Adjustments & Judgments", "Raw Workbook"]
)

with tab_overview:
    st.subheader("1. Sleeve performance")

    c1, c2 = st.columns(2)
    with c1:
        perf = pd.DataFrame(
            {
                "Metric": [
                    "Cash interest received",
                    "Accrued interest income",
                    "Royalty income",
                    "Lease cost",
                    "Net CDS premium",
                    "Gross sleeve income",
                    "Shown leverage cost",
                    "Net income after shown leverage cost",
                ],
                "USD": [
                    vals.get("Loan book — cash interest received, Q4"),
                    vals.get("Loan book — accrued interest income, Q4 (accrual basis)"),
                    vals.get("Content-library sub-licence royalty income, Q4 (per Assumption A4)"),
                    vals.get("Less: content-library lease cost, Q4 (Longwood + Nishimura, per A4)"),
                    vals.get("Less: CDS net premium paid/received, Q4"),
                    gross_income,
                    vals.get("Total fund-level leverage cost shown (partial-period, see A13)"),
                    net_income,
                ],
            }
        )
        perf["USD"] = pd.to_numeric(perf["USD"], errors="coerce").fillna(0)
        fig = px.bar(perf, x="USD", y="Metric", orientation="h", title="Income bridge")
        fig.update_layout(height=430, margin=dict(l=10, r=10, t=50, b=10))
        st.plotly_chart(fig, use_container_width=True)

    with c2:
        st.subheader("Recognition / cash bridge")
        loan_cash = vals.get("Loan book — cash interest received, Q4")
        loan_accr = vals.get("Loan book — accrued interest income, Q4 (accrual basis)")
        bridge = pd.DataFrame(
            {
                "Basis": ["Cash", "Recognized / accrued"],
                "USD": [loan_cash, loan_accr],
            }
        )
        fig = px.bar(bridge, x="Basis", y="USD", text_auto=".3s", title="Loan interest: cash vs recognition")
        fig.update_layout(height=430, margin=dict(l=10, r=10, t=50, b=10))
        st.plotly_chart(fig, use_container_width=True)

    st.subheader("What the workbook is telling the user")
    st.markdown(
        """
- **Gross yield is accrual-based:** it uses recognized/accrued loan income rather than only cash received.
- **Leverage cost is not a clean Q4 measure:** the Greystone revolver interest shown is Q3 actual cash, because Q4 fund-level revolver interest is not separately available in the workbook.
- **Hedge conclusions are directional where the source data is incomplete:** the IRS is correctly signed but cannot be quantitatively proven over Q4 from the available facility-interest time series.
- **Recognition timing is a first-class KPI:** PIK, accrued-but-unpaid interest, and the Longwood true-up are displayed explicitly.
"""
    )

with tab_loans:
    st.subheader("Loan book drill-down")
    if loans.empty:
        st.warning("Loan_Book could not be parsed.")
    else:
        loan_view = loans.copy()
        if selected_obligor:
            loan_view = loan_view[loan_view["Obligor"].astype(str).isin(selected_obligor)]
        if "Loans" not in selected_assets and "All" not in selected_assets:
            st.info("The selected asset-class filter excludes Loans.")
        else:
            total_funded = loan_view["Funded Bal. (USD)"].sum()
            cash_int = loan_view["Q4 Cash Int. Recv'd (USD)"].sum()
            accr_int = loan_view["Q4 Accrued Int. Income (USD)"].sum()
            gap = loan_view["Recog. vs Cash Gap"].sum()

            a, b, c, d = st.columns(4)
            a.metric("Funded balance", money(total_funded))
            b.metric("Cash interest", money(cash_int))
            c.metric("Recognized interest", money(accr_int))
            d.metric("Recognition − cash", money(gap))

            left, right = st.columns(2)
            with left:
                by_obligor = (
                    loan_view.groupby("Obligor", as_index=False)["Funded Bal. (USD)"]
                    .sum()
                    .sort_values("Funded Bal. (USD)", ascending=False)
                )
                fig = px.bar(by_obligor, x="Funded Bal. (USD)", y="Obligor", orientation="h", title="Funded exposure by obligor")
                fig.update_layout(height=450)
                st.plotly_chart(fig, use_container_width=True)
            with right:
                rec = loan_view.groupby("Obligor", as_index=False)[
                    ["Q4 Cash Int. Recv'd (USD)", "Q4 Accrued Int. Income (USD)", "Recog. vs Cash Gap"]
                ].sum()
                rec_long = rec.melt("Obligor", var_name="Measure", value_name="USD")
                fig = px.bar(rec_long, x="Obligor", y="USD", color="Measure", barmode="group", title="Recognition vs cash by obligor")
                fig.update_layout(height=450, xaxis_tickangle=-35)
                st.plotly_chart(fig, use_container_width=True)

            display_cols = [
                "Loan ID", "Obligor", "Ccy", "Funded Bal. (USD)", "Coupon", "Pay Freq",
                "Q4 Cash Int. Recv'd (USD)", "Q4 Accrued Int. Income (USD)",
                "Recog. vs Cash Gap", "DPD", "Notes / Assumption"
            ]
            display_cols = [c for c in display_cols if c in loan_view.columns]
            st.dataframe(
                loan_view[display_cols],
                use_container_width=True,
                hide_index=True,
                column_config={
                    "Funded Bal. (USD)": st.column_config.NumberColumn(format="$%,.0f"),
                    "Q4 Cash Int. Recv'd (USD)": st.column_config.NumberColumn(format="$%,.0f"),
                    "Q4 Accrued Int. Income (USD)": st.column_config.NumberColumn(format="$%,.0f"),
                    "Recog. vs Cash Gap": st.column_config.NumberColumn(format="$%,.0f"),
                },
            )

with tab_hedges:
    st.subheader("Hedge effectiveness and direction")
    dts = derivative_tables(sheets)

    # IRS
    irs = get_sheet(sheets, "IRS")
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("IRS notional", money(value_after_label(irs, "Notional (USD)")))
    c2.metric("PM effective exposure", money(value_after_label(irs, "PM-estimated effective floating exposure (~50% avg revolver util. blended)")))
    c3.metric("Under-hedged notional", money(value_after_label(irs, "Under-hedged notional gap")))
    c4.metric("IRS YE fair value", money(value_after_label(irs, "Swap fair value to Fund, 31-Dec-25 (MTM)")))

    st.markdown("**IRS conclusion:** Directionally correct, but full Q4 quantitative effectiveness cannot be proven from the supplied time series.")
    st.caption("The dashboard intentionally does not invent a Q4 facility-interest change where the workbook says the data is unavailable.")

    if "CDS" in dts:
        st.markdown("### CDS")
        cds = dts["CDS"]
        cds_show = [c for c in [
            "Trade Ref", "Reference Entity", "Notional (USD)", "Fund Direction",
            "Fair Value 31-Dec-25 (USD)", "Q4 Premium Cash (USD)", "Purpose", "Notes"
        ] if c in cds.columns]
        st.dataframe(cds[cds_show], use_container_width=True, hide_index=True)

    if "FX Forwards" in dts:
        st.markdown("### FX forwards")
        fx = dts["FX Forwards"]
        fx_show = [c for c in [
            "Client Ref", "Pair", "Fund Position (as executed)", "Notional",
            "Contract Rate", "Fixing/Current Rate", "Settlement / MTM P&L (USD)",
            "Hedges", "Notes / Direction Analysis"
        ] if c in fx.columns]
        st.dataframe(fx[fx_show], use_container_width=True, hide_index=True)

    if "Options" in dts:
        st.markdown("### Options — recognized fair value")
        opts = dts["Options"]
        st.dataframe(opts, use_container_width=True, hide_index=True)

with tab_recognition:
    st.subheader("Recognition timing and allocation adjustments")
    kpi = get_sheet(sheets, "KPI_Summary")

    # Find the recognition table by its header row.
    if not kpi.empty:
        mask = kpi.apply(
            lambda r: r.astype(str).str.contains("Accrued/Recognized", regex=False).any()
            and r.astype(str).str.contains("Cash Received/Paid", regex=False).any(),
            axis=1,
        )
        idxs = kpi.index[mask]
        if len(idxs):
            idx = idxs[0]
            rec = kpi.iloc[idx + 1:].copy()
            rec.columns = kpi.iloc[idx].tolist()
            rec = rec[rec.iloc[:, 0].notna()].copy()
            st.dataframe(
                rec,
                use_container_width=True,
                hide_index=True,
                column_config={
                    "Accrued/Recognized (USD)": st.column_config.NumberColumn(format="$%,.0f"),
                    "Cash Received/Paid (USD)": st.column_config.NumberColumn(format="$%,.0f"),
                    "Gap (USD)": st.column_config.NumberColumn(format="$%,.0f"),
                },
            )
            if "Gap (USD)" in rec.columns:
                chart = rec[["Item", "Gap (USD)"]].copy()
                chart["Gap (USD)"] = pd.to_numeric(chart["Gap (USD)"], errors="coerce").fillna(0)
                fig = px.bar(chart, x="Gap (USD)", y="Item", orientation="h", title="Recognition vs cash gaps")
                fig.update_layout(height=420)
                st.plotly_chart(fig, use_container_width=True)
        else:
            st.warning("Recognition table not found in KPI_Summary.")

    st.markdown(
        """
**Key interpretation:** a positive gap means income/cost has been recognized or accrued but has not yet appeared as cash in the ledger. 
The workbook also contains processing/booking adjustments that are not economic timing gaps; these are kept separate in the Adjustments tab.
"""
    )

with tab_concentration:
    st.subheader("Counterparty concentration")
    ct = get_sheet(sheets, "Concentration_Test")

    if not ct.empty:
        nav = value_after_label(ct, "Fund Total NAV (estimated, pending audit)")
        limit_pct = value_after_label(ct, "Single-name concentration limit")
        limit_usd = value_after_label(ct, "Limit in USD")
        recovery = value_after_label(kpi, "Orion Studios Holdings LLC exposure, mezz at recovery + CDS netted (recommended, A5)")
        head = value_after_label(kpi, "Headroom / (breach)")

        a, b, c, d = st.columns(4)
        a.metric("Estimated NAV", money(nav))
        b.metric("Limit", f"{pct(limit_pct)} / {money(limit_usd)}")
        c.metric("Recommended Orion exposure", money(recovery))
        d.metric("Headroom", money(head))

        # Scenario rows
        scenario_labels = [
            "Total exposure — mezz AT PAR, CDS netted",
            "Total exposure — mezz AT RECOVERY (A5), CDS netted",
            "Total exposure — mezz AT PAR, CDS NOT netted (position treated as closed)",
        ]
        scenario_vals = [value_after_label(ct, x) for x in scenario_labels]
        scen = pd.DataFrame({"Scenario": scenario_labels, "Exposure (USD)": pd.to_numeric(scenario_vals, errors="coerce")})
        scen["Limit (USD)"] = pd.to_numeric(limit_usd, errors="coerce")
        fig = px.bar(scen, x="Scenario", y="Exposure (USD)", title="Orion concentration scenarios")
        fig.add_hline(y=float(limit_usd), line_dash="dash", annotation_text="12.5% limit")
        fig.update_layout(height=450, xaxis_tickangle=-25)
        st.plotly_chart(fig, use_container_width=True)

        st.dataframe(ct, use_container_width=True, hide_index=True)

with tab_adjustments:
    st.subheader("Workbook assumptions, recognition adjustments and judgment calls")
    st.caption("These are intentionally visible so a user can understand why the dashboard metric differs from raw source cash or raw marks.")

    assumptions = assumptions_table(sheets)
    if assumptions.empty:
        st.warning("Assumptions log could not be parsed.")
    else:
        st.dataframe(assumptions, use_container_width=True, hide_index=True)

    st.markdown("### Important adjustments surfaced by the workbook")
    adjustments = pd.DataFrame(
        [
            ["A1", "IRS cash duplicate", "$1.25mm receipt excluded; $1.725mm payment adopted", "Cash / recognition"],
            ["A2", "Orion CDS settlement direction", "$14.4375mm treated as receipt to Fund", "Recognition / settlement"],
            ["A3", "GBP forward direction", "Modelled as executed; hedge is non-offsetting", "Hedge allocation"],
            ["A4", "Longwood lease + royalty", "Lease treated as sleeve investment cost and paired with royalty income", "Allocation"],
            ["A5", "Orion mezz recovery", "25mm face carried at 42.25% for concentration scenario", "Recognition / valuation"],
            ["A6", "JPY forward sizing", "337.8mm JPY under-hedge retained rather than retroactively resized", "Hedge effectiveness"],
            ["A9", "FRGF option", "USD-converted broker figure used instead of mislabeled raw GBp value", "Valuation"],
            ["A10", "SLRA option", "$117.5k memo mark excluded from recognized P&L", "Recognition"],
            ["A11", "LMVU exercised calls", "Excluded from option P&L; rolled into equity cost basis", "Allocation"],
            ["A13", "Greystone revolver Q4 cost", "Not estimated because Q4 utilization / interest is not isolated", "Recognition / data gap"],
            ["A14", "IRS sizing", "$250mm notional vs ~$350mm PM effective exposure", "Hedge effectiveness"],
        ],
        columns=["Ref", "Topic", "Treatment", "Type"],
    )
    st.dataframe(adjustments, use_container_width=True, hide_index=True)

with tab_raw:
    st.subheader("Raw workbook tabs")
    st.caption("Use this section to audit the dashboard back to the supplied Excel workbook.")
    sheet_choice = st.selectbox("Select workbook tab", list(sheets.keys()))
    raw = get_sheet(sheets, sheet_choice)
    st.dataframe(raw, use_container_width=True, hide_index=True)

    # Download current sheet as CSV
    csv = raw.to_csv(index=False).encode("utf-8")
    st.download_button(
        "Download selected tab as CSV",
        data=csv,
        file_name=f"{sheet_choice.replace(' ', '_')}.csv",
        mime="text/csv",
    )

# -----------------------------
# Footer / audit note
# -----------------------------
st.divider()
st.caption(
    "Source: supplied Whitmore_Sleeve_KPI_Workings.xlsx. "
    "The dashboard reads the workbook's calculated/cached values and does not invent missing Q4 figures. "
    "For production use, refresh the workbook formulas in Excel before uploading if source inputs have changed."
)
