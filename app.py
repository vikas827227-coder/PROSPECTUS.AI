"""
IPOReady AI — SME IPO Readiness, Disclosure & Financial Simulation Platform
Run with:  streamlit run app.py
Install deps:  pip install streamlit pandas plotly openpyxl
"""

import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
import io

try:
    import docx  # python-docx
    DOCX_SUPPORT = True
except ImportError:
    DOCX_SUPPORT = False

st.set_page_config(page_title="IPOReady AI", layout="wide")

# ----------------------------------------------------------------------
# SESSION STATE (acts as our "database" for the prototype)
# ----------------------------------------------------------------------
if "company" not in st.session_state:
    st.session_state.company = {
        "name": "",
        "industry": "",
        "business_model": "",
        "years_of_operation": 0,
        "employees": 0,
        "locations": "",
        "promoter_names": "",
    }

if "financials" not in st.session_state:
    # 3 years of data: FY24, FY25, FY26
    st.session_state.financials = pd.DataFrame({
        "Year": ["FY26"],
        "Revenue": [0.0],
        "EBITDA": [0.0],
        "PAT": [0.0],
        "Assets": [0.0],
        "Liabilities": [0.0],
        "Equity": [0.0],
        "Debt": [0.0],
        "Cash": [0.0],
        "Receivables": [0.0],
        "Inventory": [0.0],
        "InterestExpense": [0.0],
    })

if "ipo_info" not in st.session_state:
    st.session_state.ipo_info = {
        "issue_size": 0.0,
        "fresh_issue": 0.0,
        "offer_for_sale": 0.0,
        "use_of_funds": "",
    }

if "disclosures" not in st.session_state:
    # Simple checklist: item -> status (Complete / Needs Review / Missing)
    # All start as "Missing" until real data is entered — no fake sample statuses.
    st.session_state.disclosures = {
        "Historical financial statements": "Missing",
        "Promoter background & shareholding": "Missing",
        "Related party transactions": "Missing",
        "Risk factors": "Missing",
        "Legal proceedings / litigations": "Missing",
        "Business description & industry overview": "Missing",
        "Use of IPO proceeds": "Missing",
        "Statutory approvals & licenses": "Missing",
        "Material contracts": "Missing",
        "Corporate governance structure": "Missing",
    }

if "companies_data" not in st.session_state:
    # Populated when the user uploads the multi-company (10-company) workbook.
    # Maps company display name -> {"financials": DataFrame, "profile": str}
    st.session_state.companies_data = {}

if "data_loaded" not in st.session_state:
    # False until the user uploads a workbook/file or enters real data.
    st.session_state.data_loaded = False

if "cross_check" not in st.session_state:
    # user-entered vs document-extracted values, for the inconsistency detector — empty until the user adds rows
    st.session_state.cross_check = pd.DataFrame({
        "Field": pd.Series(dtype="str"),
        "Entered Value": pd.Series(dtype="float"),
        "Document Value": pd.Series(dtype="float"),
    })


# ----------------------------------------------------------------------
# WORD (.docx) DRAFT EXPORT
# ----------------------------------------------------------------------
def build_draft_docx(comp: dict, ipo: dict, fin_table: pd.DataFrame, risk_items: list,
                      ratio_rows: dict, disclosures: dict, complete_n: int, review_n: int, missing_n: int) -> bytes:
    if not DOCX_SUPPORT:
        raise RuntimeError("python-docx is not installed. Run: pip install python-docx")
    d = docx.Document()
    d.add_heading(comp["name"] or "[Company Name]", level=0)
    d.add_heading("Preliminary IPO Information Draft", level=1)
    d.add_paragraph("Prepared using the company data currently loaded in this app.").italic = True

    d.add_heading("1. Company Overview", level=2)
    d.add_paragraph(
        f"{comp['name'] or '[Company Name]'} is a company operating in the "
        f"{comp['industry'] or '[industry not specified]'} sector, engaged in "
        f"{comp['business_model'] or '[business model not specified]'}. The company has been in "
        f"operation for approximately {comp['years_of_operation']} years"
        + (f", employing around {comp['employees']:,} people" if comp['employees'] else "")
        + (f", with operations across {comp['locations']}" if comp['locations'] else "")
        + "."
    )

    d.add_heading("2. Business Model", level=2)
    d.add_paragraph(
        comp["business_model"]
        or "[Describe the company's core business model, principal revenue streams, target customer "
           "segments, and competitive positioning within its industry.]"
    )

    d.add_heading("3. Promoters & Capital Structure", level=2)
    d.add_paragraph(f"Promoter(s): {comp['promoter_names'] or '[promoter details not provided]'}")

    d.add_heading("4. Financial Summary (₹ Cr)", level=2)
    table = d.add_table(rows=1, cols=len(fin_table.columns) + 1)
    table.style = "Light Grid Accent 1"
    hdr = table.rows[0].cells
    hdr[0].text = "Metric"
    for i, col in enumerate(fin_table.columns):
        hdr[i + 1].text = str(col)
    for metric, row in fin_table.iterrows():
        cells = table.add_row().cells
        cells[0].text = str(metric)
        for i, val in enumerate(row):
            cells[i + 1].text = f"{val:.1f}"

    d.add_heading("5. Key Financial Ratios (latest year)", level=2)
    ratio_table = d.add_table(rows=1, cols=2)
    ratio_table.style = "Light Grid Accent 1"
    ratio_table.rows[0].cells[0].text = "Metric"
    ratio_table.rows[0].cells[1].text = "Value"
    for metric, val in ratio_rows.items():
        cells = ratio_table.add_row().cells
        cells[0].text = metric
        cells[1].text = str(val)

    d.add_heading("6. Disclosure Status Summary", level=2)
    d.add_paragraph(
        f"Of the {len(disclosures)} disclosure items tracked for this filing, {complete_n} are complete, "
        f"{review_n} require further review, and {missing_n} remain missing. All missing and under-review "
        f"items should be resolved prior to submission to authorised intermediaries."
    )
    for item, status in disclosures.items():
        d.add_paragraph(f"{item} — {status}", style="List Bullet")

    d.add_heading("7. Risk Factors", level=2)
    for r in risk_items:
        d.add_paragraph(r, style="List Bullet")

    d.add_heading("8. Use of IPO Proceeds", level=2)
    d.add_paragraph(
        f"The company proposes an IPO issue size of ₹{ipo['issue_size']:.1f} crore, comprising a fresh "
        f"issue of ₹{ipo['fresh_issue']:.1f} crore and an offer for sale of ₹{ipo['offer_for_sale']:.1f} crore."
    )
    d.add_paragraph(
        ipo["use_of_funds"] or "[Describe the intended use of IPO proceeds — e.g. debt repayment, working "
                                "capital, capital expenditure, or general corporate purposes.]"
    )

    d.add_paragraph()
    warn_run = d.add_paragraph().add_run(
        "⚠️ Preliminary AI-generated draft for review — requires certification by authorised intermediaries."
    )
    warn_run.italic = True

    buf = io.BytesIO()
    d.save(buf)
    buf.seek(0)
    return buf.getvalue()


# ----------------------------------------------------------------------
# MULTI-COMPANY WORKBOOK PARSER
# ----------------------------------------------------------------------
REQUIRED_FIN_COLS = ["Year", "Revenue", "EBITDA", "PAT", "Assets", "Liabilities",
                     "Equity", "Debt", "Cash", "Receivables", "Inventory", "InterestExpense"]


def parse_multi_company_workbook(file, rupees_to_crore: bool = True) -> dict:
    """
    Parses a workbook containing one sheet per company (sheet names starting
    with 'Company'), each with a title row, a header row starting with 'Year',
    and one data row per financial year. Figures are assumed to be in plain
    rupees and converted to ₹ crore (divide by 1e7) unless rupees_to_crore=False.
    Returns {company_name: {"financials": DataFrame, "profile": str}}.
    """
    xls = pd.ExcelFile(file)
    companies = {}
    for sheet in xls.sheet_names:
        if not sheet.strip().lower().startswith("company"):
            continue
        raw = pd.read_excel(xls, sheet, header=None)
        if raw.empty:
            continue

        title_cell = str(raw.iloc[0, 0])
        if "—" in title_cell:
            name_part, profile_part = title_cell.split("—", 1)
        elif " - " in title_cell:
            name_part, profile_part = title_cell.split(" - ", 1)
        else:
            name_part, profile_part = title_cell, ""
        company_name = name_part.strip() or sheet.strip()
        profile_desc = profile_part.replace("sample profile:", "").strip()

        header_row_idx = None
        for i in range(len(raw)):
            if str(raw.iloc[i, 0]).strip().lower() == "year":
                header_row_idx = i
                break
        if header_row_idx is None:
            continue

        headers = [str(h).strip() for h in raw.iloc[header_row_idx].tolist()]
        data = raw.iloc[header_row_idx + 1:].copy()
        data.columns = headers
        data = data.dropna(subset=["Year"], how="any")
        if data.empty:
            continue

        missing = [c for c in REQUIRED_FIN_COLS if c not in data.columns]
        if missing:
            continue

        numeric_cols = [c for c in REQUIRED_FIN_COLS if c != "Year"]
        for c in numeric_cols:
            data[c] = pd.to_numeric(data[c], errors="coerce")
            if rupees_to_crore:
                data[c] = data[c] / 1e7  # plain rupees -> ₹ crore
        data["Year"] = data["Year"].astype(str)
        data = data[REQUIRED_FIN_COLS].dropna(subset=numeric_cols, how="all").reset_index(drop=True)
        if data.empty:
            continue

        # Optional: a second table further down the same sheet, headed "Disclosure Item" in
        # column A and "Status" in column B, listing that company's checklist statuses.
        disclosures = {}
        disc_header_idx = None
        for i in range(len(raw)):
            if str(raw.iloc[i, 0]).strip().lower() == "disclosure item":
                disc_header_idx = i
                break
        if disc_header_idx is not None:
            valid_statuses = {"complete", "needs review", "missing"}
            for i in range(disc_header_idx + 1, len(raw)):
                item = raw.iloc[i, 0]
                status = raw.iloc[i, 1] if raw.shape[1] > 1 else None
                if pd.isna(item) or str(item).strip() == "":
                    break
                item = str(item).strip()
                status = str(status).strip() if pd.notna(status) else ""
                # normalize case so "complete" / "Complete" / "COMPLETE" all match
                matched = next((s for s in ["Complete", "Needs Review", "Missing"]
                                if s.lower() == status.lower()), None)
                disclosures[item] = matched or "Missing"

        companies[company_name] = {"financials": data, "profile": profile_desc, "disclosures": disclosures}
    return companies


# ----------------------------------------------------------------------
# FINANCIAL ENGINE
# ----------------------------------------------------------------------
def _safe_div(numerator: pd.Series, denominator: pd.Series) -> pd.Series:
    """Element-wise division that returns NaN instead of inf/-inf when the denominator is 0."""
    result = numerator / denominator.replace(0, pd.NA)
    return pd.to_numeric(result, errors="coerce")


def compute_ratios(df: pd.DataFrame) -> pd.DataFrame:
    r = df.copy()
    r["Revenue Growth %"] = r["Revenue"].pct_change() * 100
    r["EBITDA Margin %"] = _safe_div(r["EBITDA"], r["Revenue"]) * 100
    r["Net Profit Margin %"] = _safe_div(r["PAT"], r["Revenue"]) * 100
    r["ROE %"] = _safe_div(r["PAT"], r["Equity"]) * 100
    r["Current Ratio"] = _safe_div(
        r["Cash"] + r["Receivables"] + r["Inventory"], r["Liabilities"] * 0.4
    )  # simplified proxy for current liabilities
    r["Debt-to-Equity"] = _safe_div(r["Debt"], r["Equity"])
    r["Interest Coverage"] = _safe_div(r["EBITDA"], r["InterestExpense"])
    r["Receivable Days"] = _safe_div(r["Receivables"], r["Revenue"]) * 365
    r["Inventory Days"] = _safe_div(r["Inventory"], r["Revenue"]) * 365
    return r


def interpret_ratio(name: str, value: float) -> tuple[str, str]:
    """Returns (flag_emoji, plain-English explanation)."""
    if pd.isna(value):
        return "⚪", "Not enough historical data yet."
    if name == "Debt-to-Equity":
        if value > 2.0:
            return "🔴", f"Debt-to-Equity is {value:.2f}x — relatively high leverage. This may need closer review before pursuing an IPO."
        elif value > 1.0:
            return "🟡", f"Debt-to-Equity is {value:.2f}x — moderate leverage, acceptable but worth monitoring."
        return "🟢", f"Debt-to-Equity is {value:.2f}x — healthy leverage position."
    if name == "Interest Coverage":
        if value < 2:
            return "🔴", f"Interest coverage of {value:.2f}x is low — earnings may struggle to cover interest obligations."
        elif value < 4:
            return "🟡", f"Interest coverage of {value:.2f}x is moderate."
        return "🟢", f"Interest coverage of {value:.2f}x is strong."
    if name == "Current Ratio":
        if value < 1:
            return "🔴", f"Current ratio of {value:.2f} suggests potential short-term liquidity stress."
        elif value < 1.5:
            return "🟡", f"Current ratio of {value:.2f} is adequate but not strong."
        return "🟢", f"Current ratio of {value:.2f} indicates healthy liquidity."
    if name == "Net Profit Margin %":
        if value < 5:
            return "🔴", f"Net profit margin of {value:.1f}% is thin."
        elif value < 10:
            return "🟡", f"Net profit margin of {value:.1f}% is reasonable."
        return "🟢", f"Net profit margin of {value:.1f}% is strong."
    if name == "Revenue Growth %":
        if value < 0:
            return "🔴", f"Revenue declined by {abs(value):.1f}% — needs explanation for investors."
        elif value < 10:
            return "🟡", f"Revenue growth of {value:.1f}% is modest."
        return "🟢", f"Revenue growth of {value:.1f}% is strong."
    return "⚪", f"{name}: {value:.2f}"


# ----------------------------------------------------------------------
# IPO READINESS SCORE
# ----------------------------------------------------------------------
def compute_readiness_score(ratios_latest: pd.Series, disclosures: dict) -> dict:
    has_financial_data = pd.notna(ratios_latest["Revenue"]) and ratios_latest["Revenue"] > 0

    # Financial Health (0-100) — based on margin, growth, leverage
    if not has_financial_data:
        fin_score = 0
    else:
        fin_score = 100
        if ratios_latest["Debt-to-Equity"] > 2:
            fin_score -= 25
        elif ratios_latest["Debt-to-Equity"] > 1:
            fin_score -= 10
        if ratios_latest["Net Profit Margin %"] < 5:
            fin_score -= 20
        elif ratios_latest["Net Profit Margin %"] < 10:
            fin_score -= 8
        if ratios_latest["Revenue Growth %"] < 0:
            fin_score -= 20
        elif ratios_latest["Revenue Growth %"] < 10:
            fin_score -= 8
        if ratios_latest["Interest Coverage"] < 2:
            fin_score -= 15
        fin_score = max(0, fin_score)

    # Financial Consistency — placeholder tied to cross-check status (filled in later)
    consistency_score = st.session_state.get("consistency_score_cache", 0)

    # Disclosure Completeness
    total = len(disclosures)
    complete = sum(1 for v in disclosures.values() if v == "Complete")
    needs_review = sum(1 for v in disclosures.values() if v == "Needs Review")
    disclosure_score = round((complete * 1.0 + needs_review * 0.5) / total * 100)

    # Debt & Risk
    if not has_financial_data:
        debt_score = 0
    else:
        debt_score = 100
        if ratios_latest["Debt-to-Equity"] > 2:
            debt_score -= 30
        elif ratios_latest["Debt-to-Equity"] > 1:
            debt_score -= 10
        if ratios_latest["Receivable Days"] > 90:
            debt_score -= 15
        debt_score = max(0, debt_score)

    # Corporate Information — proxy: are promoter/company fields filled
    corp_fields = [st.session_state.company["name"], st.session_state.company["industry"],
                   st.session_state.company["promoter_names"]]
    corp_score = round(sum(1 for f in corp_fields if f.strip()) / len(corp_fields) * 100) if any(corp_fields) else 0

    # IPO Preparation — proxy: whether issue size / use of funds filled
    ipo_fields = [st.session_state.ipo_info["issue_size"], st.session_state.ipo_info["use_of_funds"]]
    prep_score = 80 if st.session_state.ipo_info["use_of_funds"].strip() else 0

    categories = {
        "Financial Health": fin_score,
        "Financial Consistency": consistency_score,
        "Disclosure Completeness": disclosure_score,
        "Debt & Risk": debt_score,
        "Corporate Information": corp_score,
        "IPO Preparation": prep_score,
    }
    overall = round(sum(categories.values()) / len(categories))
    return {"overall": overall, "categories": categories}


# ----------------------------------------------------------------------
# SIDEBAR NAVIGATION
# ----------------------------------------------------------------------
st.sidebar.title("🚀 IPOReady AI")
st.sidebar.caption("Prepare. Detect. Simulate. Go Public.")
page = st.sidebar.radio(
    "Navigate",
    ["🏠 Dashboard", "📊 Financial Health", "🔍 IPO Readiness", "⚠️ Gap & Risk Detector",
     "📄 Draft Generator", "🧪 What-If Simulator", "⚙️ Data Input"],
)

ratios = compute_ratios(st.session_state.financials)
latest = ratios.iloc[-1]

# Consistency check (used by readiness score + gap detector)
cc = st.session_state.cross_check.copy()
cc["Difference"] = (cc["Entered Value"] - cc["Document Value"]).abs()
cc["Status"] = cc["Difference"].apply(lambda d: "✅ Match" if d < 0.01 else "🚨 Mismatch")
mismatches = cc[cc["Status"] == "🚨 Mismatch"]
st.session_state["consistency_score_cache"] = round(
    (len(cc) - len(mismatches)) / len(cc) * 100
) if len(cc) else 0

score = compute_readiness_score(latest, st.session_state.disclosures)


# ----------------------------------------------------------------------
# PAGE: DASHBOARD
# ----------------------------------------------------------------------
if page == "🏠 Dashboard":
    st.title(f"{st.session_state.company['name'] or 'Your Company'} — IPO Readiness Overview")
    if not st.session_state.data_loaded:
        st.warning(
            "📌 No company data has been uploaded or entered yet — all figures below show ₹0. "
            "Go to **⚙️ Data Input** to upload or enter a company's real data."
        )

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Revenue (latest)", f"₹{latest['Revenue']:.1f} Cr")
    col2.metric("PAT (latest)", f"₹{latest['PAT']:.1f} Cr")
    col3.metric("Debt (latest)", f"₹{latest['Debt']:.1f} Cr")
    col4.metric("Proposed IPO Size", f"₹{st.session_state.ipo_info['issue_size']:.1f} Cr")

    st.divider()
    c1, c2 = st.columns([1, 2])
    with c1:
        fig = go.Figure(go.Indicator(
            mode="gauge+number",
            value=score["overall"],
            title={"text": "IPO Readiness Score"},
            gauge={"axis": {"range": [0, 100]},
                   "bar": {"color": "#F0F921"},  # bright plasma yellow needle/bar
                   "bgcolor": "#0D0887",  # deep plasma purple background
                   "steps": [
                       {"range": [0, 50], "color": "#0D0887"},   # dark purple
                       {"range": [50, 75], "color": "#CC4778"},  # magenta/pink
                       {"range": [75, 100], "color": "#F89441"}]}))  # vibrant orange
        st.plotly_chart(fig, use_container_width=True)
    with c2:
        cat_df = pd.DataFrame(list(score["categories"].items()), columns=["Category", "Score"])
        fig2 = px.bar(cat_df, x="Score", y="Category", orientation="h", range_x=[0, 100],
                      color="Score", color_continuous_scale="Plasma")
        st.plotly_chart(fig2, use_container_width=True)

    st.subheader("Quick Summary")
    strengths, concerns = [], []
    if latest["Revenue Growth %"] > 10:
        strengths.append("Revenue growing consistently")
    if latest["EBITDA"] > 0:
        strengths.append("Positive operating profitability")
    if latest["Net Profit Margin %"] > 5:
        strengths.append("Stable profitability")
    if latest["Debt-to-Equity"] > 1.5:
        concerns.append("High debt relative to equity")
    if any(v == "Missing" for v in st.session_state.disclosures.values()):
        concerns.append("Missing disclosure items")
    if latest["Receivable Days"] > 75:
        concerns.append("High receivable days")
    if len(mismatches):
        concerns.append("Data inconsistencies between sources")

    colA, colB = st.columns(2)
    with colA:
        st.markdown("**🟢 Strengths**")
        for s in strengths or ["— none flagged yet —"]:
            st.write(f"- {s}")
    with colB:
        st.markdown("**🔴 Concerns**")
        for c in concerns or ["— none flagged yet —"]:
            st.write(f"- {c}")


# ----------------------------------------------------------------------
# PAGE: FINANCIAL HEALTH
# ----------------------------------------------------------------------
elif page == "📊 Financial Health":
    st.title("📊 Financial Health")

    st.dataframe(ratios.set_index("Year").round(2), use_container_width=True)

    c1, c2 = st.columns(2)
    with c1:
        st.plotly_chart(px.line(ratios, x="Year", y=["Revenue", "EBITDA", "PAT"], markers=True,
                                 title="Revenue / EBITDA / PAT (₹ Cr)"), use_container_width=True)
        st.plotly_chart(px.line(ratios, x="Year", y="Debt", markers=True, title="Debt (₹ Cr)"),
                         use_container_width=True)
    with c2:
        st.plotly_chart(px.line(ratios, x="Year", y=["EBITDA Margin %", "Net Profit Margin %"], markers=True,
                                 title="Margins (%)"), use_container_width=True)
        st.plotly_chart(px.bar(ratios, x="Year", y=["Receivable Days", "Inventory Days"], barmode="group",
                                title="Working Capital Days"), use_container_width=True)

    st.subheader("Plain-English Interpretation (latest year)")
    for metric in ["Revenue Growth %", "EBITDA Margin %", "Net Profit Margin %", "Debt-to-Equity",
                   "Interest Coverage", "Current Ratio"]:
        flag, text = interpret_ratio(metric, latest[metric])
        st.write(f"{flag} **{metric}**: {text}")


# ----------------------------------------------------------------------
# PAGE: IPO READINESS
# ----------------------------------------------------------------------
elif page == "🔍 IPO Readiness":
    st.title("🔍 IPO Readiness Breakdown")
    st.metric("Overall IPO Readiness", f"{score['overall']} / 100")
    for cat, val in score["categories"].items():
        st.write(f"**{cat}**: {val}/100")
        st.progress(val / 100)


# ----------------------------------------------------------------------
# PAGE: GAP & RISK DETECTOR
# ----------------------------------------------------------------------
elif page == "⚠️ Gap & Risk Detector":
    st.title("⚠️ Gap & Risk Detector")

    st.subheader("Disclosure Checklist")
    st.caption("Click any dropdown below to update that item's status directly.")
    complete = sum(1 for v in st.session_state.disclosures.values() if v == "Complete")
    review = sum(1 for v in st.session_state.disclosures.values() if v == "Needs Review")
    missing = sum(1 for v in st.session_state.disclosures.values() if v == "Missing")
    c1, c2, c3 = st.columns(3)
    c1.metric("🟢 Complete", complete)
    c2.metric("🟡 Needs Review", review)
    c3.metric("🔴 Missing", missing)

    status_options = ["Complete", "Needs Review", "Missing"]
    for item in list(st.session_state.disclosures.keys()):
        current = st.session_state.disclosures[item]
        col_a, col_b = st.columns([3, 1])
        with col_a:
            severity = "HIGH" if current == "Missing" else ("MEDIUM" if current == "Needs Review" else "—")
            emoji = {"Complete": "🟢", "Needs Review": "🟡", "Missing": "🔴"}[current]
            st.write(f"{emoji} **{item}** (Severity: {severity})")
        with col_b:
            st.session_state.disclosures[item] = st.selectbox(
                item, status_options, index=status_options.index(current),
                key=f"gap_disc_{item}", label_visibility="collapsed",
            )

    st.divider()
    st.subheader("🚨 Inconsistency Detector")
    st.caption(
        "Enter the same figure as it appears in two different places — what you typed into the system "
        "vs. what an uploaded document shows — and this table flags any mismatch automatically."
    )
    edited_cc = st.data_editor(
        st.session_state.cross_check, num_rows="dynamic", use_container_width=True, key="gap_cross_check_editor"
    )
    st.session_state.cross_check = edited_cc

    cc_live = edited_cc.copy()
    cc_live["Difference"] = (cc_live["Entered Value"] - cc_live["Document Value"]).abs()
    cc_live["Status"] = cc_live["Difference"].apply(lambda d: "✅ Match" if d < 0.01 else "🚨 Mismatch")
    live_mismatches = cc_live[cc_live["Status"] == "🚨 Mismatch"]

    if len(live_mismatches):
        for _, row in live_mismatches.iterrows():
            st.error(
                f"Potential inconsistency in **{row['Field']}**: entered ₹{row['Entered Value']} Cr "
                f"vs document ₹{row['Document Value']} Cr (difference ₹{row['Difference']:.1f} Cr). "
                f"Status: Requires verification."
            )
    else:
        st.success("No inconsistencies detected across the checked fields.")


# ----------------------------------------------------------------------
# PAGE: DRAFT GENERATOR
# ----------------------------------------------------------------------
elif page == "📄 Draft Generator":
    st.title("📄 Preliminary IPO Draft Generator")
    st.info("This produces a **Preliminary AI-generated draft for review** — not a SEBI-compliant final document.")

    if st.button("GENERATE PRELIMINARY IPO DRAFT", type="primary"):
        comp = st.session_state.company
        ipo = st.session_state.ipo_info
        fin_table = st.session_state.financials.set_index("Year")[
            ["Revenue", "EBITDA", "PAT", "Debt", "Equity"]
        ].T

        risk_items = []
        if pd.notna(latest["Debt-to-Equity"]) and latest["Debt-to-Equity"] > 1.5:
            risk_items.append(
                f"The company carries a Debt-to-Equity ratio of {latest['Debt-to-Equity']:.2f}x, "
                f"reflecting relatively high financial leverage that investors and regulators may scrutinize closely."
            )
        if pd.notna(latest["Receivable Days"]) and latest["Receivable Days"] > 75:
            risk_items.append(
                f"Trade receivables stand at approximately {latest['Receivable Days']:.0f} days of revenue, "
                f"indicating a degree of customer concentration or collection risk."
            )
        if pd.notna(latest["Interest Coverage"]) and latest["Interest Coverage"] < 3:
            risk_items.append(
                "Interest coverage is relatively thin, meaning a downturn in operating profit could strain "
                "the company's ability to service its debt obligations."
            )
        risk_items.append(
            "Input costs, raw material prices, and broader macroeconomic conditions may affect margins."
        )
        risk_items.append(
            "The business remains subject to sector-specific regulatory changes and competitive pressure."
        )

        complete_n = sum(1 for v in st.session_state.disclosures.values() if v == "Complete")
        review_n = sum(1 for v in st.session_state.disclosures.values() if v == "Needs Review")
        missing_n = sum(1 for v in st.session_state.disclosures.values() if v == "Missing")

        st.markdown(f"## {comp['name'] or '[Company Name]'}")
        st.markdown("### Preliminary IPO Information Draft")
        st.caption("Prepared using the company data currently loaded in this app.")

        st.markdown("**1. Company Overview**")
        st.write(
            f"{comp['name'] or '[Company Name]'} is a company operating in the "
            f"{comp['industry'] or '[industry not specified]'} sector, engaged in "
            f"{comp['business_model'] or '[business model not specified]'}. The company has been in "
            f"operation for approximately {comp['years_of_operation']} years"
            + (f", employing around {comp['employees']:,} people" if comp['employees'] else "")
            + (f", with operations across {comp['locations']}" if comp['locations'] else "")
            + f". Based on the most recent reported financial year, the company generated revenue of "
            f"₹{latest['Revenue']:.1f} crore with a profit after tax of ₹{latest['PAT']:.1f} crore."
        )

        st.markdown("**2. Business Model**")
        st.write(
            comp["business_model"]
            or "[Describe the company's core business model, principal revenue streams, target customer "
               "segments, and competitive positioning within its industry.]"
        )

        st.markdown("**3. Promoters & Capital Structure**")
        st.write(
            f"Promoter(s): {comp['promoter_names'] or '[promoter details not provided]'}. "
            f"As of the latest reported period, the company's total equity (net worth) stood at "
            f"₹{latest['Equity']:.1f} crore against total debt of ₹{latest['Debt']:.1f} crore, "
            f"implying a Debt-to-Equity ratio of "
            + (f"{latest['Debt-to-Equity']:.2f}x." if pd.notna(latest["Debt-to-Equity"]) else "not meaningful with the data currently available.")
        )

        st.markdown("**4. Financial Summary (₹ Cr)**")
        st.dataframe(fin_table.round(1), use_container_width=True)

        st.markdown("**5. Key Financial Ratios (latest year)**")
        ratio_rows = {
            "Revenue Growth": f"{latest['Revenue Growth %']:.1f}%" if pd.notna(latest["Revenue Growth %"]) else "N/A",
            "EBITDA Margin": f"{latest['EBITDA Margin %']:.1f}%" if pd.notna(latest["EBITDA Margin %"]) else "N/A",
            "Net Profit Margin": f"{latest['Net Profit Margin %']:.1f}%" if pd.notna(latest["Net Profit Margin %"]) else "N/A",
            "Return on Equity (ROE)": f"{latest['ROE %']:.1f}%" if pd.notna(latest["ROE %"]) else "N/A",
            "Debt-to-Equity": f"{latest['Debt-to-Equity']:.2f}x" if pd.notna(latest["Debt-to-Equity"]) else "N/A",
            "Interest Coverage": f"{latest['Interest Coverage']:.2f}x" if pd.notna(latest["Interest Coverage"]) else "N/A",
            "Current Ratio": f"{latest['Current Ratio']:.2f}" if pd.notna(latest["Current Ratio"]) else "N/A",
        }
        st.table(pd.DataFrame(ratio_rows.items(), columns=["Metric", "Value"]).set_index("Metric"))

        st.markdown("**6. Disclosure Status Summary**")
        st.write(
            f"Of the {len(st.session_state.disclosures)} disclosure items tracked for this filing, "
            f"{complete_n} are complete, {review_n} require further review, and {missing_n} remain missing. "
            f"All missing and under-review items should be resolved prior to submission to authorised intermediaries."
        )
        for item, status in st.session_state.disclosures.items():
            emoji = {"Complete": "🟢", "Needs Review": "🟡", "Missing": "🔴"}[status]
            st.write(f"{emoji} {item} — {status}")

        st.markdown("**7. Risk Factors**")
        for r in risk_items:
            st.write(f"- {r}")

        st.markdown("**8. Use of IPO Proceeds**")
        st.write(
            f"The company proposes an IPO issue size of ₹{ipo['issue_size']:.1f} crore, comprising a fresh "
            f"issue of ₹{ipo['fresh_issue']:.1f} crore and an offer for sale of ₹{ipo['offer_for_sale']:.1f} crore. "
        )
        st.write(ipo["use_of_funds"] or "[Describe the intended use of IPO proceeds — e.g. debt repayment, "
                                          "working capital, capital expenditure, or general corporate purposes.]")

        st.warning("⚠️ Preliminary AI-generated draft for review — requires certification by authorised intermediaries.")

        if DOCX_SUPPORT:
            docx_bytes = build_draft_docx(comp, ipo, fin_table, risk_items, ratio_rows,
                                            st.session_state.disclosures, complete_n, review_n, missing_n)
            st.download_button(
                "⬇️ Download this draft as a Word document (.docx)",
                data=docx_bytes,
                file_name=f"{(comp['name'] or 'IPO_Draft').replace(' ', '_')}_Preliminary_IPO_Draft.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            )
        else:
            st.info("Install python-docx (`pip install python-docx`) to enable downloading this draft as a Word file.")


# ----------------------------------------------------------------------
# PAGE: WHAT-IF SIMULATOR
# ----------------------------------------------------------------------
elif page == "🧪 What-If Simulator":
    st.title("🧪 What-If Financial Simulator")
    st.caption("Adjust assumptions to see the projected impact on the latest year's financials.")

    base = latest.copy()
    debt_change = st.slider("Change in Debt (₹ Cr)", -10.0, 20.0, 0.0, 0.5)
    interest_rate = st.slider("Assumed interest rate on incremental debt (%)", 5.0, 15.0, 9.0, 0.5)
    revenue_change_pct = st.slider("Change in Revenue (%)", -20.0, 30.0, 0.0, 1.0)

    new_debt = base["Debt"] + debt_change
    new_interest_expense = base["InterestExpense"] + (debt_change * interest_rate / 100)
    new_revenue = base["Revenue"] * (1 + revenue_change_pct / 100)
    new_ebitda = base["EBITDA"] * (new_revenue / base["Revenue"]) if base["Revenue"] else base["EBITDA"]
    new_pat = base["PAT"] + (new_ebitda - base["EBITDA"]) - (new_interest_expense - base["InterestExpense"])
    new_de = new_debt / base["Equity"]
    new_interest_coverage = new_ebitda / new_interest_expense if new_interest_expense else float("inf")

    sim_latest = base.copy()
    sim_latest["Debt"] = new_debt
    sim_latest["Revenue"] = new_revenue
    sim_latest["PAT"] = new_pat
    sim_latest["Debt-to-Equity"] = new_de
    sim_latest["Interest Coverage"] = new_interest_coverage
    sim_latest["Net Profit Margin %"] = new_pat / new_revenue * 100 if new_revenue else 0
    sim_latest["Revenue Growth %"] = revenue_change_pct

    sim_score = compute_readiness_score(sim_latest, st.session_state.disclosures)

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Debt-to-Equity", f"{new_de:.2f}x", f"{new_de - base['Debt-to-Equity']:+.2f}")
    c2.metric("Interest Expense", f"₹{new_interest_expense:.2f} Cr",
              f"{new_interest_expense - base['InterestExpense']:+.2f}")
    c3.metric("PAT", f"₹{new_pat:.2f} Cr", f"{new_pat - base['PAT']:+.2f}")
    c4.metric("IPO Readiness", f"{sim_score['overall']}", f"{sim_score['overall'] - score['overall']:+d}")


# ----------------------------------------------------------------------
# PAGE: DATA INPUT
# ----------------------------------------------------------------------
elif page == "⚙️ Data Input":
    st.title("⚙️ Company & Financial Data Input")
    st.caption(
        "Upload your company's data below — either a 10-company workbook or a single company's file. "
        "Manual entry is available further down for anything an upload doesn't cover, or if you don't "
        "have a file to upload."
    )

    st.subheader("📤 Option 1: Upload 10-Company Workbook")
    st.caption(
        "Upload the multi-sheet Excel workbook (one tab per company, tabs named 'Company 1', 'Company 2', "
        "etc.). Figures are read as plain rupees and auto-converted to ₹ crore."
    )
    multi_file = st.file_uploader("Upload the 10-company Excel workbook", type=["xlsx"], key="multi_company_upload")
    if multi_file is not None:
        try:
            parsed = parse_multi_company_workbook(multi_file)
            if not parsed:
                st.error("No valid 'Company' sheets found. Check that each tab has a 'Year' header row "
                          "and the required columns.")
            else:
                st.session_state.companies_data = parsed
                st.success(f"Loaded {len(parsed)} companies: {', '.join(parsed.keys())}")
        except Exception as e:
            st.error(f"Couldn't read that workbook: {e}")

    if st.session_state.companies_data:
        chosen_name = st.selectbox(
            "Select a company to load into the app",
            list(st.session_state.companies_data.keys()),
            key="company_selector",
        )
        if st.button("Load selected company into the app", type="primary"):
            chosen = st.session_state.companies_data[chosen_name]
            st.session_state.financials = chosen["financials"].copy()
            st.session_state.company["name"] = chosen_name
            if chosen["profile"]:
                st.session_state.company["business_model"] = chosen["profile"]
            if chosen.get("disclosures"):
                # Only overwrite items the workbook actually specified — keep existing
                # statuses for anything the workbook's checklist didn't mention.
                st.session_state.disclosures.update(chosen["disclosures"])
                st.info(f"Loaded {len(chosen['disclosures'])} disclosure statuses from the workbook too.")
            st.session_state.data_loaded = True
            st.success(f"Loaded {chosen_name}. Switch to another page from the sidebar to see it reflected.")
            st.dataframe(st.session_state.financials, use_container_width=True)

    st.divider()
    st.subheader("📤 Option 2: Upload a Single Company's Data (CSV/Excel)")
    st.caption(
        "Upload a CSV or Excel file with one company's financial data. "
        "The file must have these exact column headers: Year, Revenue, EBITDA, PAT, Assets, Liabilities, "
        "Equity, Debt, Cash, Receivables, Inventory, InterestExpense — one row per financial year."
    )

    template_csv = (
        "Year,Revenue,EBITDA,PAT,Assets,Liabilities,Equity,Debt,Cash,Receivables,Inventory,InterestExpense\n"
        "FY24,35,6,3,40,20,20,12,4,6,5,1.4\n"
        "FY25,42,7.5,3.4,48,24,24,15,5,8,6,1.7\n"
        "FY26,50,9,4,55,28,27,18,5.5,11,7,2.1\n"
    )
    st.download_button(
        "⬇️ Download blank template (CSV)",
        data=template_csv,
        file_name="company_financials_template.csv",
        mime="text/csv",
    )

    uploaded_file = st.file_uploader("Upload financial data", type=["csv", "xlsx"])
    if uploaded_file is not None:
        try:
            if uploaded_file.name.endswith(".csv"):
                new_df = pd.read_csv(uploaded_file)
            else:
                new_df = pd.read_excel(uploaded_file)

            required_cols = ["Year", "Revenue", "EBITDA", "PAT", "Assets", "Liabilities",
                              "Equity", "Debt", "Cash", "Receivables", "Inventory", "InterestExpense"]
            missing = [c for c in required_cols if c not in new_df.columns]
            if missing:
                st.error(f"Your file is missing these columns: {', '.join(missing)}. "
                          f"Use the template above to check the exact spelling.")
            else:
                st.session_state.financials = new_df[required_cols].reset_index(drop=True)
                st.session_state.data_loaded = True
                st.success(f"Loaded {len(new_df)} year(s) of data from {uploaded_file.name}. "
                            f"Switch to another page from the sidebar to see it reflected.")
                st.dataframe(st.session_state.financials, use_container_width=True)
        except Exception as e:
            st.error(f"Couldn't read that file: {e}")

    st.divider()
    with st.expander("✏️ Manually enter or edit company, IPO & financial details (optional)"):
        st.caption(
            "Use this if you don't have a file to upload, or want to fill in details an upload doesn't "
            "cover (like company name, promoters, or IPO issue structure)."
        )
        st.markdown("**Company Information**")
        comp = st.session_state.company
        comp["name"] = st.text_input("Company name", comp["name"])
        if comp["name"].strip():
            st.session_state.data_loaded = True
        comp["industry"] = st.text_input("Industry", comp["industry"])
        comp["business_model"] = st.text_area("Business model", comp["business_model"])
        comp["years_of_operation"] = st.number_input("Years of operation", 0, 100, comp["years_of_operation"])
        comp["employees"] = st.number_input("Number of employees", 0, 100000, comp["employees"])
        comp["locations"] = st.text_input("Locations", comp["locations"])
        comp["promoter_names"] = st.text_input("Promoter name(s)", comp["promoter_names"])

        st.markdown("**IPO Information**")
        ipo = st.session_state.ipo_info
        ipo["issue_size"] = st.number_input("Proposed issue size (₹ Cr)", 0.0, value=float(ipo["issue_size"]))
        ipo["fresh_issue"] = st.number_input("Fresh issue (₹ Cr)", 0.0, value=float(ipo["fresh_issue"]))
        ipo["offer_for_sale"] = st.number_input("Offer for sale (₹ Cr)", 0.0, value=float(ipo["offer_for_sale"]))
        ipo["use_of_funds"] = st.text_area("Intended use of funds", ipo["use_of_funds"])

        st.markdown("**Financial Data**")
        edited_financials = st.data_editor(
            st.session_state.financials, num_rows="dynamic", use_container_width=True, key="manual_financials_editor"
        )
        st.session_state.financials = edited_financials
        if edited_financials.drop(columns=["Year"]).fillna(0).sum().sum() > 0:
            st.session_state.data_loaded = True

    st.caption(
        "For the **Disclosure Checklist** and **Inconsistency Detector**, go to the "
        "⚠️ Gap & Risk Detector page — both are directly editable there."
    )
