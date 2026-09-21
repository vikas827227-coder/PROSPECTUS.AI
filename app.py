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

# "prep_applicable": True  = unlisted company (single-company upload)  -> IPO Preparation included, score out of 100
#                    False = listed company (Banker/Auditor upload)     -> IPO Preparation removed, score out of 95
DEFAULT_IPO_INFO = {
    "issue_size": 0.0,
    "fresh_issue": 0.0,
    "offer_for_sale": 0.0,
    "use_of_funds": "",
    "prep_applicable": True,
}

if "ipo_info" not in st.session_state:
    st.session_state.ipo_info = dict(DEFAULT_IPO_INFO)

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

if "legal_compliance" not in st.session_state:
    # Self-reported legal/compliance flags. "No" is the good answer for all of these.
    st.session_state.legal_compliance = {
        "Pending litigation against the company": "Not disclosed",
        "Unresolved tax disputes": "Not disclosed",
        "Past regulatory violations / SEBI action": "Not disclosed",
        "Related-party transactions properly disclosed": "Not disclosed",
    }

if "governance_info" not in st.session_state:
    st.session_state.governance_info = {
        "auditor_name": "",
        "auditor_is_reputable": "Not disclosed",  # Yes / No / Not disclosed
        "independent_directors_pct": 0,
        "audit_committee_exists": "Not disclosed",
        "qualified_audit_opinion_last3yrs": "Not disclosed",  # "No" is good
    }

if "industry_risk" not in st.session_state:
    st.session_state.industry_risk = {
        "top_customer_concentration_pct": 0,  # % of revenue from largest single customer
        "key_regulatory_license_status": "Not disclosed",  # Obtained / Pending / Not Applicable / Not disclosed
        "market_sentiment": "Not assessed",  # Favorable / Neutral / Unfavorable / Not assessed — informational only
    }

if "valuation_info" not in st.session_state:
    st.session_state.valuation_info = {
        "peer_pe_multiple": 0.0,
    }

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
# SHARED SHEET PARSER (used by both single-company and multi-company uploads)
# ----------------------------------------------------------------------
REQUIRED_FIN_COLS = ["Year", "Revenue", "EBITDA", "PAT", "Assets", "Liabilities",
                     "Equity", "Debt", "Cash", "Receivables", "Inventory", "InterestExpense"]


NA_TOKENS = {"n/a", "na", "n.a.", "n.a", "not applicable", "-", "—", "nil"}


def _is_na(value) -> bool:
    """True if a cell contains an explicit 'N/A'-style marker (not just an empty cell)."""
    return pd.notna(value) and str(value).strip().lower() in NA_TOKENS


def _find_row(raw: pd.DataFrame, label: str):
    """Returns the row index where column A matches `label` (case-insensitive), or None."""
    for i in range(len(raw)):
        if str(raw.iloc[i, 0]).strip().lower() == label.lower():
            return i
    return None


def parse_company_sheet(raw: pd.DataFrame, rupees_to_crore: bool = True) -> dict:
    """
    Parses ONE sheet containing everything for one company:
      - Row 1: title, format "Company Name — sample profile: one-line description"
      - A financial table headed by a "Year" row (required)
      - Optional "Disclosure Item" / "Status" table
      - Optional "IPO Info" key/value table (IssueSize, FreshIssue, OfferForSale, UseOfFunds)
      - Optional "Company Info" key/value table (Industry, YearsOfOperation, Employees,
        Locations, PromoterNames)
      - Optional "Cross Check" table with Field / Entered Value / Document Value columns
    Returns a dict with keys: name, profile, financials, disclosures, ipo_info,
    company_info, cross_check — any section not found is left as an empty dict/DataFrame.
    """
    result = {
        "name": "", "profile": "", "financials": None, "disclosures": {},
        "ipo_info": {}, "company_info": {}, "cross_check": None,
        "legal_compliance": {}, "governance_info": {}, "industry_risk": {}, "valuation_info": {},
    }
    if raw.empty:
        return result

    title_cell = str(raw.iloc[0, 0])
    if "—" in title_cell:
        name_part, profile_part = title_cell.split("—", 1)
    elif " - " in title_cell:
        name_part, profile_part = title_cell.split(" - ", 1)
    else:
        name_part, profile_part = title_cell, ""
    result["name"] = name_part.strip()
    result["profile"] = profile_part.replace("sample profile:", "").strip()

    # --- Financial table (required) ---
    header_row_idx = _find_row(raw, "year")
    if header_row_idx is not None:
        headers = [str(h).strip() for h in raw.iloc[header_row_idx].tolist()]
        data_rows = []
        for i in range(header_row_idx + 1, len(raw)):
            year_cell = raw.iloc[i, 0]
            if pd.isna(year_cell) or str(year_cell).strip() == "":
                break  # stop at the first blank row — don't swallow later sections
            data_rows.append(raw.iloc[i].tolist())
        if data_rows:
            data = pd.DataFrame(data_rows, columns=headers)
            missing = [c for c in REQUIRED_FIN_COLS if c not in data.columns]
            if not missing:
                numeric_cols = [c for c in REQUIRED_FIN_COLS if c != "Year"]
                for c in numeric_cols:
                    data[c] = pd.to_numeric(data[c], errors="coerce")
                    if rupees_to_crore:
                        data[c] = data[c] / 1e7  # plain rupees -> ₹ crore
                data["Year"] = data["Year"].astype(str)
                data = data[REQUIRED_FIN_COLS].dropna(subset=numeric_cols, how="all").reset_index(drop=True)
                if not data.empty:
                    result["financials"] = data

    # --- Disclosure checklist (optional) ---
    disc_header_idx = _find_row(raw, "disclosure item")
    if disc_header_idx is not None:
        for i in range(disc_header_idx + 1, len(raw)):
            item = raw.iloc[i, 0]
            status = raw.iloc[i, 1] if raw.shape[1] > 1 else None
            if pd.isna(item) or str(item).strip() == "":
                break
            item = str(item).strip()
            status = str(status).strip() if pd.notna(status) else ""
            matched = next((s for s in ["Complete", "Needs Review", "Missing"]
                            if s.lower() == status.lower()), None)
            result["disclosures"][item] = matched or "Missing"

    # --- IPO Info key/value table (optional) ---
    # NOTE: an "IPO Preparation" row, if present, is ignored. Whether a company is scored out of
    # 95 (listed) or 100 (unlisted) is decided by WHICH UPLOAD OPTION is used, not by the sheet.
    ipo_header_idx = _find_row(raw, "ipo info")
    if ipo_header_idx is not None:
        ipo_keys = {"issuesize": "issue_size", "freshissue": "fresh_issue",
                    "offerforsale": "offer_for_sale", "useoffunds": "use_of_funds"}
        for i in range(ipo_header_idx + 1, len(raw)):
            key_cell = raw.iloc[i, 0]
            val_cell = raw.iloc[i, 1] if raw.shape[1] > 1 else None
            if pd.isna(key_cell) or str(key_cell).strip() == "":
                break
            key_norm = str(key_cell).strip().lower().replace(" ", "").replace("_", "")
            if key_norm in ipo_keys:
                target = ipo_keys[key_norm]
                if target == "use_of_funds":
                    result["ipo_info"][target] = (
                        "" if _is_na(val_cell)
                        else (str(val_cell).strip() if pd.notna(val_cell) else "")
                    )
                else:
                    result["ipo_info"][target] = (
                        float(val_cell) if pd.notna(val_cell) and not _is_na(val_cell) else 0.0
                    )

    # --- Company Info key/value table (optional) ---
    comp_header_idx = _find_row(raw, "company info")
    if comp_header_idx is not None:
        comp_keys = {"industry": "industry", "yearsofoperation": "years_of_operation",
                     "employees": "employees", "locations": "locations",
                     "promoternames": "promoter_names"}
        for i in range(comp_header_idx + 1, len(raw)):
            key_cell = raw.iloc[i, 0]
            val_cell = raw.iloc[i, 1] if raw.shape[1] > 1 else None
            if pd.isna(key_cell) or str(key_cell).strip() == "":
                break
            key_norm = str(key_cell).strip().lower().replace(" ", "").replace("_", "")
            if key_norm in comp_keys:
                target = comp_keys[key_norm]
                if target in ("years_of_operation", "employees"):
                    result["company_info"][target] = int(val_cell) if pd.notna(val_cell) else 0
                else:
                    result["company_info"][target] = str(val_cell).strip() if pd.notna(val_cell) else ""

    # --- Cross-check table (optional) ---
    cross_header_idx = _find_row(raw, "cross check")
    if cross_header_idx is not None:
        col_header_idx = cross_header_idx + 1
        rows = []
        for i in range(col_header_idx + 1, len(raw)):
            field = raw.iloc[i, 0]
            entered = raw.iloc[i, 1] if raw.shape[1] > 1 else None
            document = raw.iloc[i, 2] if raw.shape[1] > 2 else None
            if pd.isna(field) or str(field).strip() == "":
                break
            rows.append({
                "Field": str(field).strip(),
                "Entered Value": float(entered) if pd.notna(entered) else 0.0,
                "Document Value": float(document) if pd.notna(document) else 0.0,
            })
        if rows:
            result["cross_check"] = pd.DataFrame(rows)

    # --- Legal Info key/value table (optional) ---
    legal_header_idx = _find_row(raw, "legal info")
    if legal_header_idx is not None:
        legal_keys = {
            "pendinglitigation": "Pending litigation against the company",
            "taxdisputes": "Unresolved tax disputes",
            "regulatoryviolations": "Past regulatory violations / SEBI action",
            "relatedpartydisclosed": "Related-party transactions properly disclosed",
        }
        for i in range(legal_header_idx + 1, len(raw)):
            key_cell = raw.iloc[i, 0]
            val_cell = raw.iloc[i, 1] if raw.shape[1] > 1 else None
            if pd.isna(key_cell) or str(key_cell).strip() == "":
                break
            key_norm = str(key_cell).strip().lower().replace(" ", "").replace("_", "")
            if key_norm in legal_keys:
                val = str(val_cell).strip() if pd.notna(val_cell) else "Not disclosed"
                matched = next((s for s in ["Yes", "No", "Not disclosed"] if s.lower() == val.lower()), "Not disclosed")
                result["legal_compliance"][legal_keys[key_norm]] = matched

    # --- Governance Info key/value table (optional) ---
    gov_header_idx = _find_row(raw, "governance info")
    if gov_header_idx is not None:
        for i in range(gov_header_idx + 1, len(raw)):
            key_cell = raw.iloc[i, 0]
            val_cell = raw.iloc[i, 1] if raw.shape[1] > 1 else None
            if pd.isna(key_cell) or str(key_cell).strip() == "":
                break
            key_norm = str(key_cell).strip().lower().replace(" ", "").replace("_", "")
            yn_options = ["Yes", "No", "Not disclosed"]
            if key_norm == "auditorname":
                result["governance_info"]["auditor_name"] = str(val_cell).strip() if pd.notna(val_cell) else ""
            elif key_norm == "auditorreputable":
                val = str(val_cell).strip() if pd.notna(val_cell) else "Not disclosed"
                result["governance_info"]["auditor_is_reputable"] = next(
                    (s for s in yn_options if s.lower() == val.lower()), "Not disclosed")
            elif key_norm == "independentdirectorspct":
                result["governance_info"]["independent_directors_pct"] = int(val_cell) if pd.notna(val_cell) else 0
            elif key_norm == "auditcommittee":
                val = str(val_cell).strip() if pd.notna(val_cell) else "Not disclosed"
                result["governance_info"]["audit_committee_exists"] = next(
                    (s for s in yn_options if s.lower() == val.lower()), "Not disclosed")
            elif key_norm == "qualifiedopinion":
                val = str(val_cell).strip() if pd.notna(val_cell) else "Not disclosed"
                result["governance_info"]["qualified_audit_opinion_last3yrs"] = next(
                    (s for s in yn_options if s.lower() == val.lower()), "Not disclosed")

    # --- Industry Risk key/value table (optional) ---
    ind_header_idx = _find_row(raw, "industry risk")
    if ind_header_idx is not None:
        for i in range(ind_header_idx + 1, len(raw)):
            key_cell = raw.iloc[i, 0]
            val_cell = raw.iloc[i, 1] if raw.shape[1] > 1 else None
            if pd.isna(key_cell) or str(key_cell).strip() == "":
                break
            key_norm = str(key_cell).strip().lower().replace(" ", "").replace("_", "")
            if key_norm == "customerconcentrationpct":
                result["industry_risk"]["top_customer_concentration_pct"] = int(val_cell) if pd.notna(val_cell) else 0
            elif key_norm == "regulatorylicensestatus":
                val = str(val_cell).strip() if pd.notna(val_cell) else "Not disclosed"
                options = ["Obtained", "Pending", "Not Applicable", "Not disclosed"]
                result["industry_risk"]["key_regulatory_license_status"] = next(
                    (s for s in options if s.lower() == val.lower()), "Not disclosed")
            elif key_norm == "marketsentiment":
                val = str(val_cell).strip() if pd.notna(val_cell) else "Not assessed"
                options = ["Favorable", "Neutral", "Unfavorable", "Not assessed"]
                result["industry_risk"]["market_sentiment"] = next(
                    (s for s in options if s.lower() == val.lower()), "Not assessed")

    # --- Valuation Info key/value table (optional) ---
    val_header_idx = _find_row(raw, "valuation info")
    if val_header_idx is not None:
        for i in range(val_header_idx + 1, len(raw)):
            key_cell = raw.iloc[i, 0]
            val_cell = raw.iloc[i, 1] if raw.shape[1] > 1 else None
            if pd.isna(key_cell) or str(key_cell).strip() == "":
                break
            key_norm = str(key_cell).strip().lower().replace(" ", "").replace("_", "")
            if key_norm == "peerpemultiple":
                result["valuation_info"]["peer_pe_multiple"] = float(val_cell) if pd.notna(val_cell) else 0.0

    return result


def parse_multi_company_workbook(file, rupees_to_crore: bool = True) -> dict:
    """
    Parses a workbook containing one sheet per company (sheet names starting
    with 'Company'). Each sheet is parsed with parse_company_sheet, which reads
    financials plus any optional Disclosure/IPO Info/Company Info/Cross Check
    sections present on that same sheet.
    Returns {company_name: {"financials", "profile", "disclosures", "ipo_info",
    "company_info", "cross_check"}}.
    """
    xls = pd.ExcelFile(file)
    companies = {}
    for sheet in xls.sheet_names:
        if not sheet.strip().lower().startswith("company"):
            continue
        raw = pd.read_excel(xls, sheet, header=None, keep_default_na=False, na_values=[""])
        parsed = parse_company_sheet(raw, rupees_to_crore=rupees_to_crore)
        if parsed["financials"] is None:
            continue
        company_name = parsed["name"] or sheet.strip()
        companies[company_name] = {
            "financials": parsed["financials"],
            "profile": parsed["profile"],
            "disclosures": parsed["disclosures"],
            "ipo_info": parsed["ipo_info"],
            "company_info": parsed["company_info"],
            "cross_check": parsed["cross_check"],
            "legal_compliance": parsed.get("legal_compliance", {}),
            "governance_info": parsed.get("governance_info", {}),
            "industry_risk": parsed.get("industry_risk", {}),
            "valuation_info": parsed.get("valuation_info", {}),
        }
    return companies


# ----------------------------------------------------------------------
# EXCEL FORMAT GUIDE + DOWNLOADABLE TEMPLATES (shown on the Data Input page)
# ----------------------------------------------------------------------
DEFAULT_DISCLOSURE_ITEMS = [
    "Historical financial statements", "Promoter background & shareholding", "Related party transactions",
    "Risk factors", "Legal proceedings / litigations", "Business description & industry overview",
    "Use of IPO proceeds", "Statutory approvals & licenses", "Material contracts",
    "Corporate governance structure",
]

OPTION1_NOTES = """
**Option 1 — Banker/Auditor view: LISTED companies, one workbook with many sheets**

- **File:** one `.xlsx` workbook, **one sheet (tab) per company**.
- **Tab names:** must **start with `Company`** — e.g. `Company 1`, `Company 2`, … (any number of tabs, not limited to 10).
  Tabs that don't start with `Company`, or that have no valid financial table, are **skipped silently**.
- **Company name:** the text before the dash in cell `A1`. If A1 has no name, the tab name is used.
  Two tabs with the same company name → the later one overwrites the earlier one.
- **Money in the financial table:** plain rupees (e.g. `500000000` = ₹50 Cr). The app converts to ₹ crore automatically.
- **Scoring:** every company loaded through this option is treated as **listed** and scored **out of 95** — the IPO Preparation category
  (5 points) is not included. This depends only on using Option 1; nothing in the sheet has to be marked. Any `IPO Preparation` row in the sheet is ignored.
- **After uploading:** pick a company in the dropdown, then click **Load selected company into the app**. The score only changes on load.
"""

OPTION2_NOTES = """
**Option 2 — Single UNLISTED (IPO-bound) company, one file**

- **Excel (`.xlsx`):** only the **first sheet** is read; the tab name doesn't matter. One company per file.
  Money in the financial table is **plain rupees** and is converted to ₹ crore automatically (same as Option 1).
- **CSV (`.csv`):** contains **only** the 12-column financial table (header row exactly as below, no title row, no other sections).
  Unlike Excel, CSV values are **not converted** — enter them **already in ₹ crore** (e.g. `50` for ₹50 Cr).
- **Scoring:** treated as **unlisted** → scored **out of 100**. IPO Preparation (5 points) is always included and is scored by **how many of the 4 IPO Info fields are filled in** (`IssueSize`, `FreshIssue`, `OfferForSale`, `UseOfFunds`): 4 filled = 5/5, 3 = 4/5, 2 = 2/5, 1 = 1/5, none = 0/5. A number counts as filled only if it is above 0; text counts only if it is not empty, `0`, `nil`, `none` or `N/A`. Any `IPO Preparation` row in the sheet is ignored.
- **Loads immediately** — no company-selection step.
"""

SHEET_FORMAT_GUIDE = """
#### Sheet layout (identical for every company sheet)

Row 1 is the title. The financial table is the only required section; all others are optional and can appear in
**any order** below the title. Put **at least one completely blank row between sections** — each section stops at the
first empty cell in column A. Section labels and keys are **not case-sensitive**, and spaces/underscores in keys are
ignored (`Years Of Operation` = `YearsOfOperation`). Labels/keys go in **column A**, values in **column B**
(the financial table uses columns A–L, Cross Check uses A–C).

| # | Section | Cell in column A that starts it | What goes underneath (column A key → column B value) | Required? |
|---|---|---|---|---|
| 1 | Title | `A1` | `Company Name — sample profile: one-line description`. Text before the dash = company name; text after it (with "sample profile:" removed) = Business Model. A plain ` - ` hyphen also works. | Yes |
| 2 | Financial table | `Year` | A header row with exactly these 12 column names: `Year, Revenue, EBITDA, PAT, Assets, Liabilities, Equity, Debt, Cash, Receivables, Inventory, InterestExpense`. Then one row per financial year, **oldest first, latest year last**. Ends at the first blank Year cell. | **Yes** |
| 3 | Company Info | `Company Info` | `Industry`, `YearsOfOperation`, `Employees`, `Locations`, `PromoterNames` | Optional |
| 4 | IPO Info | `IPO Info` | `IssueSize`, `FreshIssue`, `OfferForSale`, `UseOfFunds` | Optional |
| 5 | Disclosure checklist | `Disclosure Item` | Column A = item name, column B = its status. The header row itself is `Disclosure Item` in A and `Status` in B. | Optional |
| 6 | Cross Check | `Cross Check` | The next row is a header: `Field`, `Entered Value`, `Document Value` (columns A–C). Then one row per figure compared. | Optional |
| 7 | Legal Info | `Legal Info` | `PendingLitigation`, `TaxDisputes`, `RegulatoryViolations`, `RelatedPartyDisclosed` | Optional |
| 8 | Governance Info | `Governance Info` | `AuditorName`, `AuditorReputable`, `IndependentDirectorsPct`, `AuditCommittee`, `QualifiedOpinion` | Optional |
| 9 | Industry Risk | `Industry Risk` | `CustomerConcentrationPct`, `RegulatoryLicenseStatus`, `MarketSentiment` | Optional |
| 10 | Valuation Info | `Valuation Info` | `PeerPEMultiple` | Optional |

#### What to type in each cell (types, units, allowed values)

| Field | Type | Allowed values / unit |
|---|---|---|
| `Year` | text | Any label, e.g. `FY24`, `FY25`, `FY26` |
| `Revenue, EBITDA, PAT, Assets, Liabilities, Equity, Debt, Cash, Receivables, Inventory, InterestExpense` | number | **Plain rupees** in Excel (e.g. `500000000` = ₹50 Cr) — no commas, no `₹`, no text. Column order doesn't matter and extra columns are ignored, but the names must be spelled **exactly** as above (case-sensitive). |
| `IssueSize`, `FreshIssue`, `OfferForSale` | number | **₹ crore** (not rupees), e.g. `25`. `N/A` is read as 0. Each one above 0 earns IPO Preparation points (Option 2 only). |
| `UseOfFunds` | text | Free text. `N/A`, `nil`, `none`, `0` or blank = not filled. Filled earns IPO Preparation points (Option 2 only). |
| `Industry`, `Locations`, `PromoterNames`, `AuditorName` | text | Free text. Separate several locations/promoters with commas. |
| `YearsOfOperation`, `Employees` | whole number | Digits only, e.g. `12`, `250` (text like "250+" makes the file fail to load). |
| Disclosure `Status` | text | `Complete`, `Needs Review` or `Missing` (anything else is read as `Missing`). |
| Cross Check `Entered Value`, `Document Value` | number | **₹ crore.** Flagged as a mismatch when they differ by 0.01 or more. `Field` is free text. |
| `PendingLitigation`, `TaxDisputes`, `RegulatoryViolations` | text | `Yes`, `No` or `Not disclosed`. **`No` is the good answer.** |
| `RelatedPartyDisclosed` | text | `Yes`, `No` or `Not disclosed`. **`Yes` is the good answer.** |
| `AuditorReputable`, `AuditCommittee` | text | `Yes`, `No` or `Not disclosed`. `Yes` is the good answer. |
| `QualifiedOpinion` | text | `Yes`, `No` or `Not disclosed` — a qualified audit opinion in the last 3 years. **`No` is the good answer.** |
| `IndependentDirectorsPct` | whole number | 0–100, written as `40` (not `40%` or `0.4`). 33 or above scores full marks. |
| `CustomerConcentrationPct` | whole number | 0–100 — share of revenue from the largest customer, e.g. `18`. Above 25 triggers a warning. |
| `RegulatoryLicenseStatus` | text | `Obtained`, `Pending`, `Not Applicable` or `Not disclosed` |
| `MarketSentiment` | text | `Favorable`, `Neutral`, `Unfavorable` or `Not assessed` (informational only, not scored) |
| `PeerPEMultiple` | number | Peer P/E multiple, e.g. `22.5` (used by the valuation estimator) |

Values that don't match the allowed list fall back to `Not disclosed` / `Not assessed`, and an empty value cell falls back to `0`, blank text or `Not disclosed`.

#### Disclosure checklist — the 10 default item names
Use these exact names in column A so they update the built-in items (a new name is added as an extra item):
""" + "\n".join(f"{n}. `{item}`" for n, item in enumerate(DEFAULT_DISCLOSURE_ITEMS, start=1)) + """

#### Most common mistakes
- **Numbers saved as text** (`5,00,00,000`, `₹50 Cr`, `50 Cr`) are read as blank — type plain numbers.
- **Years in the wrong order** — put the oldest year first and the latest last (growth and "latest year" use the last row).
- **No blank row between sections**, or a blank row *inside* a section — the section ends early.
- **Wrong units** — financial table in rupees, but IPO Info and Cross Check in ₹ crore.
- **Misspelled column names** in the financial header row (e.g. `Interest Expense` instead of `InterestExpense`).
"""


def build_template_workbook(listed: bool) -> bytes:
    """Builds a ready-to-fill example workbook in exactly the format the parser reads.
    listed=True  -> Option 1 template: two example tabs ('Company 1', 'Company 2')
    listed=False -> Option 2 template: one example tab."""
    from openpyxl import Workbook
    from openpyxl.styles import Font

    bold = Font(bold=True)
    wb = Workbook()
    wb.remove(wb.active)

    def write_sheet(ws, title, prep_row_value):
        r = [1]

        def put(values, is_bold=False):
            for j, v in enumerate(values, start=1):
                c = ws.cell(row=r[0], column=j, value=v)
                if is_bold:
                    c.font = bold
            r[0] += 1

        def blank():
            r[0] += 1

        put([title], True)
        blank()
        put(REQUIRED_FIN_COLS, True)
        put(["FY24", 350000000, 60000000, 30000000, 400000000, 200000000, 200000000, 120000000, 40000000, 60000000, 50000000, 14000000])
        put(["FY25", 420000000, 75000000, 34000000, 480000000, 240000000, 240000000, 150000000, 50000000, 80000000, 60000000, 17000000])
        put(["FY26", 500000000, 90000000, 40000000, 550000000, 280000000, 270000000, 180000000, 55000000, 110000000, 70000000, 21000000])
        blank()
        put(["Company Info", "Value"], True)
        put(["Industry", "Manufacturing"])
        put(["YearsOfOperation", 12])
        put(["Employees", 250])
        put(["Locations", "Mumbai, Pune"])
        put(["PromoterNames", "Promoter One, Promoter Two"])
        blank()
        put(["IPO Info", "Value"], True)
        put(["IssueSize", 25])
        put(["FreshIssue", 20])
        put(["OfferForSale", 5])
        put(["UseOfFunds", "Debt repayment, working capital and general corporate purposes"])
        if prep_row_value is not None:
            put(["IPO Preparation", prep_row_value])
        blank()
        put(["Disclosure Item", "Status"], True)
        for i, item in enumerate(DEFAULT_DISCLOSURE_ITEMS):
            put([item, ["Complete", "Needs Review", "Missing"][i % 3]])
        blank()
        put(["Cross Check"], True)
        put(["Field", "Entered Value", "Document Value"], True)
        put(["Revenue FY26 (₹ Cr)", 50, 50])
        put(["PAT FY26 (₹ Cr)", 4, 4.5])
        blank()
        put(["Legal Info", "Value"], True)
        put(["PendingLitigation", "No"])
        put(["TaxDisputes", "No"])
        put(["RegulatoryViolations", "No"])
        put(["RelatedPartyDisclosed", "Yes"])
        blank()
        put(["Governance Info", "Value"], True)
        put(["AuditorName", "Example & Associates"])
        put(["AuditorReputable", "Yes"])
        put(["IndependentDirectorsPct", 40])
        put(["AuditCommittee", "Yes"])
        put(["QualifiedOpinion", "No"])
        blank()
        put(["Industry Risk", "Value"], True)
        put(["CustomerConcentrationPct", 18])
        put(["RegulatoryLicenseStatus", "Obtained"])
        put(["MarketSentiment", "Neutral"])
        blank()
        put(["Valuation Info", "Value"], True)
        put(["PeerPEMultiple", 22.5])
        ws.column_dimensions["A"].width = 44
        ws.column_dimensions["B"].width = 22
        ws.column_dimensions["C"].width = 18

    if listed:
        write_sheet(wb.create_sheet("Company 1"),
                    "Example Listed Company A — sample profile: replace with a one-line business description", None)
        write_sheet(wb.create_sheet("Company 2"),
                    "Example Listed Company B — sample profile: replace with a one-line business description", None)
    else:
        write_sheet(wb.create_sheet("Company 1"),
                    "Example Unlisted Company — sample profile: replace with a one-line business description", None)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


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

    # IPO Preparation — scored by how much of the IPO Info is filled in (4 fields: Issue Size,
    # Fresh Issue, Offer for Sale, Use of Funds). A field counts as filled if it is > 0 / non-empty
    # (text like N/A, nil, 0, none counts as NOT filled). 4 filled = 5/5, 3 = 4/5, 2 = 2/5, 1 = 1/5, 0 = 0/5.
    # Only included for unlisted companies (single-company upload). For listed companies
    # (Banker/Auditor upload) the category is removed entirely, so the total becomes 95 instead of 100.
    prep_applicable = st.session_state.ipo_info.get("prep_applicable", True)
    _ipo = st.session_state.ipo_info
    _empty_text = NA_TOKENS | {"", "0", "none"}
    prep_items_filled = sum([
        float(_ipo.get("issue_size") or 0) > 0,
        float(_ipo.get("fresh_issue") or 0) > 0,
        float(_ipo.get("offer_for_sale") or 0) > 0,
        str(_ipo.get("use_of_funds") or "").strip().lower() not in _empty_text,
    ])
    prep_score = prep_items_filled / 4 * 100

    # Legal & Compliance — self-reported flags; "No" is the good answer for the first 3,
    # "Yes" is good for the 4th (related-party disclosure). Unanswered items count as unresolved.
    legal = st.session_state.legal_compliance
    legal_points = 0
    good_answers = {
        "Pending litigation against the company": "No",
        "Unresolved tax disputes": "No",
        "Past regulatory violations / SEBI action": "No",
        "Related-party transactions properly disclosed": "Yes",
    }
    for item, good_answer in good_answers.items():
        if legal.get(item) == good_answer:
            legal_points += 1
    legal_score = round(legal_points / len(good_answers) * 100) if any(v != "Not disclosed" for v in legal.values()) else 0

    # Governance & Auditor Quality
    gov = st.session_state.governance_info
    if not any([gov["auditor_name"].strip(), gov["auditor_is_reputable"] != "Not disclosed",
                gov["audit_committee_exists"] != "Not disclosed", gov["independent_directors_pct"] > 0]):
        gov_score = 0
    else:
        gov_score = 100
        if gov["auditor_is_reputable"] != "Yes":
            gov_score -= 25
        if gov["audit_committee_exists"] != "Yes":
            gov_score -= 25
        if gov["independent_directors_pct"] < 33:
            gov_score -= 25
        if gov["qualified_audit_opinion_last3yrs"] == "Yes":
            gov_score -= 25
        gov_score = max(0, gov_score)

    # Each category has its own maximum points (weights sum to 100), reflecting that
    # some factors matter more than others for IPO readiness — not a flat average.
    category_weights = {
        "Financial Health": 25,
        "Debt & Risk": 20,
        "Disclosure Completeness": 15,
        "Legal & Compliance": 15,
        "Governance & Auditor Quality": 10,
        "Financial Consistency": 5,
        "Corporate Information": 5,
    }
    if prep_applicable:
        category_weights["IPO Preparation"] = 5
    raw_scores = {
        "Financial Health": fin_score,
        "Financial Consistency": consistency_score,
        "Disclosure Completeness": disclosure_score,
        "Debt & Risk": debt_score,
        "Corporate Information": corp_score,
        "IPO Preparation": prep_score,  # dropped below for listed companies
        "Legal & Compliance": legal_score,
        "Governance & Auditor Quality": gov_score,
    }
    # Convert each 0-100 raw score into points out of that category's own max weight
    categories = {
        name: round(raw_scores[name] / 100 * max_points)
        for name, max_points in category_weights.items()
    }
    category_max = category_weights
    overall = sum(categories.values())
    max_total = sum(category_weights.values())  # 100 with IPO Preparation, 95 without
    return {"overall": overall, "max_total": max_total,
            "categories": categories, "category_max": category_max}


def check_sebi_eligibility(financials: pd.DataFrame, ratios: pd.DataFrame) -> list:
    """
    A pass/fail table against a few of SEBI's actual, commonly-cited SME IPO eligibility
    criteria — NOT the full legal requirement list (see the caption on the page that uses
    this), just the handful that can be checked directly from financial statements.
    """
    checks = []
    if financials.empty or financials["Revenue"].fillna(0).sum() == 0:
        return [{"Requirement": "No financial data entered yet", "Company Value": "—", "Status": "⚪ N/A"}]

    last3 = ratios.tail(3)
    positive_years = (last3["PAT"] > 0).sum()
    checks.append({
        "Requirement": "Positive PAT in at least 2 of the last 3 years",
        "Company Value": f"{positive_years} of {len(last3)} year(s) positive",
        "Status": "✅ Pass" if positive_years >= 2 else "❌ Fail",
    })

    latest_networth = ratios.iloc[-1]["Equity"]
    checks.append({
        "Requirement": "Positive net worth (Equity)",
        "Company Value": f"₹{latest_networth:.1f} Cr",
        "Status": "✅ Pass" if pd.notna(latest_networth) and latest_networth > 0 else "❌ Fail",
    })

    net_tangible_assets = ratios.iloc[-1]["Assets"]
    checks.append({
        "Requirement": "Net Tangible Assets ≥ ₹3 crore (SME IPO minimum, illustrative)",
        "Company Value": f"₹{net_tangible_assets:.1f} Cr",
        "Status": "✅ Pass" if pd.notna(net_tangible_assets) and net_tangible_assets >= 3 else "❌ Fail",
    })

    latest_debt_equity = ratios.iloc[-1]["Debt-to-Equity"]
    checks.append({
        "Requirement": "Debt-to-Equity below 3x (general prudence guideline, not a hard SEBI rule)",
        "Company Value": f"{latest_debt_equity:.2f}x" if pd.notna(latest_debt_equity) else "N/A",
        "Status": "✅ Pass" if pd.isna(latest_debt_equity) or latest_debt_equity < 3 else "❌ Fail",
    })

    return checks


# ----------------------------------------------------------------------
# SIDEBAR NAVIGATION
# ----------------------------------------------------------------------
st.sidebar.title("🚀 IPOReady AI")
st.sidebar.caption("Prepare. Detect. Simulate. Go Public.")
page = st.sidebar.radio(
    "Navigate",
    ["🏠 Dashboard", "📊 Financial Health", "🔍 IPO Readiness", "⚠️ Gap & Risk Detector",
     "⚖️ Legal, Governance & Valuation", "📄 Draft Generator", "🧪 What-If Simulator", "⚙️ Data Input"],
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
    if st.session_state.ipo_info.get("prep_applicable", True):
        col4.metric("Proposed IPO Size", f"₹{st.session_state.ipo_info['issue_size']:.1f} Cr")
    else:
        col4.metric("Proposed IPO Size", "N/A (listed)")

    st.divider()
    c1, c2 = st.columns([1, 2])
    with c1:
        fig = go.Figure(go.Indicator(
            mode="gauge+number",
            value=score["overall"],
            number={"suffix": f" / {score['max_total']}"},
            title={"text": "IPO Readiness Score"},
            gauge={"axis": {"range": [0, score["max_total"]]},
                   "bar": {"color": "#F0F921"},  # bright plasma yellow needle/bar
                   "bgcolor": "#0D0887",  # deep plasma purple background
                   "steps": [
                       {"range": [0, score["max_total"] * 0.50], "color": "#0D0887"},   # dark purple
                       {"range": [score["max_total"] * 0.50, score["max_total"] * 0.75], "color": "#CC4778"},  # magenta/pink
                       {"range": [score["max_total"] * 0.75, score["max_total"]], "color": "#F89441"}]}))  # vibrant orange
        st.plotly_chart(fig, use_container_width=True)
    with c2:
        cat_df = pd.DataFrame([
            {"Category": cat, "Points": val, "Max": score["category_max"][cat],
             "Percent": round(val / score["category_max"][cat] * 100) if score["category_max"][cat] else 0,
             "Label": f"{val}/{score['category_max'][cat]}"}
            for cat, val in score["categories"].items()
        ])
        fig2 = px.bar(cat_df, x="Percent", y="Category", orientation="h", range_x=[0, 100],
                      color="Percent", color_continuous_scale="Plasma", text="Label")
        fig2.update_layout(xaxis_title="% of that category's max points")
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
    st.metric("Overall IPO Readiness", f"{score['overall']} / {score['max_total']}")
    st.caption("Each category has its own maximum weight, reflecting that some factors matter more than others.")
    if not st.session_state.ipo_info.get("prep_applicable", True):
        st.info("Listed company (loaded via the Banker/Auditor upload): the IPO Preparation category is not "
                "included, so the score is out of 95.")
    else:
        st.caption("Unlisted company: IPO Preparation is included and scored by how much IPO Info is filled in "
                   "(Issue Size, Fresh Issue, Offer for Sale, Use of Funds), so the score is out of 100.")
    for cat, val in score["categories"].items():
        max_pts = score["category_max"][cat]
        st.write(f"**{cat}**: {val}/{max_pts}")
        st.progress(val / max_pts if max_pts else 0)


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
# PAGE: LEGAL, GOVERNANCE & VALUATION
# ----------------------------------------------------------------------
elif page == "⚖️ Legal, Governance & Valuation":
    st.title("⚖️ Legal, Governance & Valuation")
    st.caption(
        "These checks cover areas real IPO due diligence looks at beyond the financial numbers. "
        "This is still self-reported and simplified — it does not replace an actual legal or "
        "audit review, and does not check SEBI's full eligibility rulebook."
    )

    st.subheader("Legal & Compliance")
    st.caption("Answer honestly — 'Not disclosed' counts the same as an unresolved risk in the score.")
    legal = st.session_state.legal_compliance
    options_legal = ["Yes", "No", "Not disclosed"]
    for item in list(legal.keys()):
        legal[item] = st.selectbox(
            item, options_legal, index=options_legal.index(legal[item]), key=f"legal_{item}"
        )

    st.divider()
    st.subheader("Governance & Auditor Quality")
    gov = st.session_state.governance_info
    gov["auditor_name"] = st.text_input("Auditor name", gov["auditor_name"])
    options_yn = ["Yes", "No", "Not disclosed"]
    gov["auditor_is_reputable"] = st.selectbox(
        "Is the auditor a reputable/well-established firm?", options_yn,
        index=options_yn.index(gov["auditor_is_reputable"]),
    )
    gov["independent_directors_pct"] = st.number_input(
        "Independent directors on the board (%)", 0, 100, gov["independent_directors_pct"]
    )
    gov["audit_committee_exists"] = st.selectbox(
        "Does the company have a functioning audit committee?", options_yn,
        index=options_yn.index(gov["audit_committee_exists"]),
    )
    gov["qualified_audit_opinion_last3yrs"] = st.selectbox(
        "Any qualified audit opinion in the last 3 years? (a red flag if Yes)", options_yn,
        index=options_yn.index(gov["qualified_audit_opinion_last3yrs"]),
    )

    st.divider()
    st.subheader("Industry & Concentration Risk")
    ind = st.session_state.industry_risk
    ind["top_customer_concentration_pct"] = st.number_input(
        "Revenue share from the single largest customer (%)", 0, 100,
        ind["top_customer_concentration_pct"]
    )
    if ind["top_customer_concentration_pct"] > 25:
        st.warning("⚠️ High customer concentration (>25% from one customer) is a common red flag in real due diligence.")
    options_license = ["Obtained", "Pending", "Not Applicable", "Not disclosed"]
    ind["key_regulatory_license_status"] = st.selectbox(
        "Status of key industry-specific regulatory license/approval (if applicable)",
        options_license, index=options_license.index(ind["key_regulatory_license_status"])
    )
    options_sentiment = ["Favorable", "Neutral", "Unfavorable", "Not assessed"]
    ind["market_sentiment"] = st.selectbox(
        "Current IPO market sentiment (informational only — not scored, since it's external "
        "to the company)", options_sentiment, index=options_sentiment.index(ind["market_sentiment"])
    )

    st.divider()
    st.subheader("📊 SEBI Eligibility Check (illustrative)")
    st.caption(
        "A pass/fail check against a few commonly-cited SME IPO criteria, computed directly from "
        "your financial data — kept separate from the 0-100 score above since these are meant to "
        "be closer to fixed regulatory thresholds, not weighted factors. This is NOT the complete "
        "SEBI rulebook — verify actual eligibility with a merchant banker."
    )
    sebi_checks = check_sebi_eligibility(st.session_state.financials, ratios)
    st.table(pd.DataFrame(sebi_checks).set_index("Requirement"))

    st.divider()
    st.subheader("💰 Valuation Estimator (illustrative)")
    st.caption(
        "A rough valuation range using your company's PAT and a peer P/E multiple — a real IPO "
        "valuation involves far more (DCF models, EV/EBITDA, growth-adjusted multiples, banker "
        "negotiation), so treat this as a conversation starter, not a valuation opinion."
    )
    val = st.session_state.valuation_info
    val["peer_pe_multiple"] = st.number_input(
        "Typical P/E multiple for similar companies in this industry", 0.0, 200.0,
        float(val["peer_pe_multiple"]), step=0.5
    )
    if val["peer_pe_multiple"] > 0 and pd.notna(latest["PAT"]) and latest["PAT"] > 0:
        low_pe = val["peer_pe_multiple"] * 0.8
        high_pe = val["peer_pe_multiple"] * 1.2
        low_val = latest["PAT"] * low_pe
        high_val = latest["PAT"] * high_pe
        st.metric("Estimated Pre-IPO Valuation Range",
                   f"₹{low_val:.0f} Cr – ₹{high_val:.0f} Cr")
        st.caption(
            f"Based on PAT of ₹{latest['PAT']:.1f} Cr × a {low_pe:.1f}x–{high_pe:.1f}x P/E range "
            f"(±20% around your entered peer multiple of {val['peer_pe_multiple']:.1f}x)."
        )
    else:
        st.info("Enter a peer P/E multiple above, and make sure PAT is loaded, to see an estimated valuation range.")


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

    st.subheader("📤 Option 1: Banker/Auditor View — Listed Companies (Compare Multiple)")
    st.caption(
        "For merchant bankers, auditors, or due-diligence teams screening several **listed** companies at once. "
        "Upload the multi-sheet Excel workbook (one tab per company, tabs named 'Company 1', 'Company 2', "
        "etc.). Figures are read as plain rupees and auto-converted to ₹ crore. "
        "Every company loaded here is treated as listed and scored **out of 95** (the IPO Preparation category is not included). "
        "*(An unlisted, IPO-bound company checking only its own readiness should use Option 2 below instead.)*"
    )
    st.download_button(
        "⬇️ Download Option 1 Excel template (listed companies, 2 example tabs)",
        data=build_template_workbook(listed=True),
        file_name="option1_listed_companies_template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="dl_template_opt1",
    )
    with st.expander("📋 Exact Excel format — Option 1 (listed companies, multi-sheet workbook)"):
        st.markdown(OPTION1_NOTES)
        st.markdown(SHEET_FORMAT_GUIDE)
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
            st.session_state.ipo_info = dict(DEFAULT_IPO_INFO)  # clear any previous company's IPO info
            st.session_state.ipo_info["prep_applicable"] = False  # Option 1 = LISTED company -> scored out of 95
            st.session_state.company["name"] = chosen_name
            if chosen["profile"]:
                st.session_state.company["business_model"] = chosen["profile"]
            extras_loaded = []
            if chosen.get("disclosures"):
                st.session_state.disclosures.update(chosen["disclosures"])
                extras_loaded.append(f"{len(chosen['disclosures'])} disclosure statuses")
            if chosen.get("ipo_info"):
                st.session_state.ipo_info.update(chosen["ipo_info"])
                extras_loaded.append("IPO info")
            if chosen.get("company_info"):
                st.session_state.company.update(chosen["company_info"])
                extras_loaded.append("company info")
            if chosen.get("cross_check") is not None and not chosen["cross_check"].empty:
                st.session_state.cross_check = chosen["cross_check"].copy()
                extras_loaded.append(f"{len(chosen['cross_check'])} cross-check rows")
            if chosen.get("legal_compliance"):
                st.session_state.legal_compliance.update(chosen["legal_compliance"])
                extras_loaded.append("legal & compliance info")
            if chosen.get("governance_info"):
                st.session_state.governance_info.update(chosen["governance_info"])
                extras_loaded.append("governance info")
            if chosen.get("industry_risk"):
                st.session_state.industry_risk.update(chosen["industry_risk"])
                extras_loaded.append("industry risk info")
            if chosen.get("valuation_info"):
                st.session_state.valuation_info.update(chosen["valuation_info"])
                extras_loaded.append("valuation info")
            if extras_loaded:
                st.info(f"Also loaded from the workbook: {', '.join(extras_loaded)}.")
            st.session_state.data_loaded = True
            st.success(f"Loaded {chosen_name}. Switch to another page from the sidebar to see it reflected.")
            st.dataframe(st.session_state.financials, use_container_width=True)

    st.divider()
    st.subheader("📤 Option 2: Upload a Single Unlisted Company's Data (CSV/Excel)")
    st.caption(
        "For an individual **unlisted** company preparing for its IPO — scored **out of 100**, with IPO Preparation (5 points) based on how much IPO Info is filled in. "
        "A CSV needs just the 12-column financial table below. An Excel (.xlsx) file can include "
        "everything the app uses in one sheet — company info, IPO info, 10 years of financials, "
        "the disclosure checklist, and the cross-check table. See the format guide below."
    )

    template_csv = (
        "Year,Revenue,EBITDA,PAT,Assets,Liabilities,Equity,Debt,Cash,Receivables,Inventory,InterestExpense\n"
        "FY24,35,6,3,40,20,20,12,4,6,5,1.4\n"
        "FY25,42,7.5,3.4,48,24,24,15,5,8,6,1.7\n"
        "FY26,50,9,4,55,28,27,18,5.5,11,7,2.1\n"
    )
    st.download_button(
        "⬇️ Download blank CSV template (financials only)",
        data=template_csv,
        file_name="company_financials_template.csv",
        mime="text/csv",
    )

    st.download_button(
        "⬇️ Download Option 2 Excel template (unlisted company, everything in one sheet)",
        data=build_template_workbook(listed=False),
        file_name="option2_unlisted_company_template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="dl_template_opt2",
    )
    with st.expander("📋 Exact Excel format — Option 2 (single unlisted company: Excel or CSV)"):
        st.markdown(OPTION2_NOTES)
        st.markdown(SHEET_FORMAT_GUIDE)

    uploaded_file = st.file_uploader("Upload company data", type=["csv", "xlsx"])
    if uploaded_file is not None:
        try:
            if uploaded_file.name.endswith(".csv"):
                new_df = pd.read_csv(uploaded_file)
                required_cols = ["Year", "Revenue", "EBITDA", "PAT", "Assets", "Liabilities",
                                  "Equity", "Debt", "Cash", "Receivables", "Inventory", "InterestExpense"]
                missing = [c for c in required_cols if c not in new_df.columns]
                if missing:
                    st.error(f"Your file is missing these columns: {', '.join(missing)}. "
                              f"Use the template above to check the exact spelling.")
                else:
                    st.session_state.financials = new_df[required_cols].reset_index(drop=True)
                    st.session_state.ipo_info["prep_applicable"] = True  # Option 2 = unlisted -> out of 100
                    st.session_state.data_loaded = True
                    st.success(f"Loaded {len(new_df)} year(s) of financial data from {uploaded_file.name}.")
                    st.dataframe(st.session_state.financials, use_container_width=True)
            else:
                raw = pd.read_excel(uploaded_file, sheet_name=0, header=None, keep_default_na=False, na_values=[""])
                parsed = parse_company_sheet(raw)
                if parsed["financials"] is None:
                    st.error("Couldn't find a valid financial table (a row starting with 'Year' "
                              "followed by the required columns). Check the format guide above.")
                else:
                    st.session_state.financials = parsed["financials"]
                    st.session_state.ipo_info = dict(DEFAULT_IPO_INFO)  # clear any previous company's IPO info
                    if parsed["name"]:
                        st.session_state.company["name"] = parsed["name"]
                    if parsed["profile"]:
                        st.session_state.company["business_model"] = parsed["profile"]
                    loaded_extras = ["financial data"]
                    if parsed["company_info"]:
                        st.session_state.company.update(parsed["company_info"])
                        loaded_extras.append("company info")
                    if parsed["ipo_info"]:
                        st.session_state.ipo_info.update(parsed["ipo_info"])
                        loaded_extras.append("IPO info")
                    # Option 2 = UNLISTED company -> IPO Preparation included, scored out of 100
                    st.session_state.ipo_info["prep_applicable"] = True
                    if parsed["disclosures"]:
                        st.session_state.disclosures.update(parsed["disclosures"])
                        loaded_extras.append(f"{len(parsed['disclosures'])} disclosure statuses")
                    if parsed["cross_check"] is not None and not parsed["cross_check"].empty:
                        st.session_state.cross_check = parsed["cross_check"]
                        loaded_extras.append(f"{len(parsed['cross_check'])} cross-check rows")
                    if parsed.get("legal_compliance"):
                        st.session_state.legal_compliance.update(parsed["legal_compliance"])
                        loaded_extras.append("legal & compliance info")
                    if parsed.get("governance_info"):
                        st.session_state.governance_info.update(parsed["governance_info"])
                        loaded_extras.append("governance info")
                    if parsed.get("industry_risk"):
                        st.session_state.industry_risk.update(parsed["industry_risk"])
                        loaded_extras.append("industry risk info")
                    if parsed.get("valuation_info"):
                        st.session_state.valuation_info.update(parsed["valuation_info"])
                        loaded_extras.append("valuation info")
                    st.session_state.data_loaded = True
                    st.success(f"Loaded from {uploaded_file.name}: {', '.join(loaded_extras)}.")
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
