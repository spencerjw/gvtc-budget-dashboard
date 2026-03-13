import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import csv
import io
import re
import json

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------
st.set_page_config(
    page_title="GVTC Dept 44 - Web Management Budget",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

MONTH_LABELS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
                "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]

# Friendly category names keyed by GL account
GL_NAMES = {
    "6124.32": "Software Licenses",
    "6124.34": "GPC Office Supplies",
    "6623.15": "Transportation & Lodging",
    "6623.16": "Travel Meals",
    "6623.17": "Auto Reimbursement",
    "6623.18": "Business Meals",
    "6623.19": "Employee Appreciation",
    "6623.22": "Seminars & Training",
    "6623.34": "Office Supplies",
    "6623.36": "Cellular Phone",
    "6623.74": "Contract Labor",
}

PLOTLY_LAYOUT = dict(
    paper_bgcolor="rgba(0,0,0,0)",
    plot_bgcolor="rgba(0,0,0,0)",
    font_color="#FAFAFA",
    margin=dict(l=20, r=20, t=40, b=20),
)

COLORS = px.colors.qualitative.Set3


def _to_float(val: str) -> float:
    """Convert a string like '1,167 ' or '(1,265)' to float."""
    if not val or not isinstance(val, str):
        return 0.0
    s = val.strip().replace(",", "").replace(" ", "")
    if not s:
        return 0.0
    # Handle parenthetical negatives: (1265) -> -1265
    if s.startswith("(") and s.endswith(")"):
        s = "-" + s[1:-1]
    try:
        return float(s)
    except ValueError:
        return 0.0


# ---------------------------------------------------------------------------
# Parsers
# ---------------------------------------------------------------------------

def _xlsx_bytes_to_rows(file_bytes: bytes) -> list[list[str]]:
    """Convert xlsx bytes to list of rows using xlsx2csv."""
    try:
        from xlsx2csv import Xlsx2csv
    except ImportError:
        return []

    buf_in = io.BytesIO(file_bytes)
    buf_out = io.StringIO()

    try:
        Xlsx2csv(buf_in, skip_empty_lines=False).convert(buf_out)
    except Exception:
        return []

    buf_out.seek(0)
    return list(csv.reader(buf_out))


def parse_forecast_file(file_bytes: bytes) -> dict:
    """Parse the Web Management Yearly Budget Forecast file.

    This single file contains all GL accounts with:
    - 2025 Total (actual/forecast)
    - 2026/2027/2028 annual budgets + monthly breakdowns
    - Comments per year

    Returns dict keyed by GL account number, each with:
        name, years (year -> annual), monthly (year -> [12 months]), comments (year -> str)
    """
    rows = _xlsx_bytes_to_rows(file_bytes)
    if not rows:
        return {}

    # Header row 0 has column labels
    header = rows[0] if rows else []

    # Detect year columns from header: find "YYYY Total" and "YYYY" + "YYYY-01"..."YYYY-12"
    year_columns = {}  # year -> {total_col, annual_col, monthly_start, comments_col}

    for ci, cell in enumerate(header):
        cell = cell.strip()

        # "2025 Total" -> 2025 total column (no monthly)
        m = re.match(r"^(\d{4}) Total$", cell)
        if m:
            year = int(m.group(1))
            if year not in year_columns:
                year_columns[year] = {"total_col": ci}
            else:
                year_columns[year]["total_col"] = ci
            continue

        # "2026" standalone (annual budget column, before monthly)
        m = re.match(r"^(\d{4})$", cell)
        if m:
            year = int(m.group(1))
            # Check if next col is YYYY-01
            if ci + 1 < len(header) and re.match(rf"^{year}-01$", header[ci + 1].strip()):
                if year not in year_columns:
                    year_columns[year] = {}
                year_columns[year]["annual_col"] = ci
                year_columns[year]["monthly_start"] = ci + 1

        # "Comments YYYY"
        m = re.match(r"^Comments (\d{4})$", cell)
        if m:
            year = int(m.group(1))
            if year not in year_columns:
                year_columns[year] = {}
            year_columns[year]["comments_col"] = ci

    budgets = {}

    for r in rows[2:]:  # Skip header rows
        if len(r) < 11:
            continue

        # Account info in col 4: "6124.32 - General Purpose Computers - ..."
        acct_cell = r[4].strip() if len(r) > 4 else ""
        m = re.match(r"^(\d{4}\.\d{2})\s*-\s*(.+)$", acct_cell)
        if not m:
            continue

        account = m.group(1)
        full_name = m.group(2).strip()

        entry = {
            "name": GL_NAMES.get(account, full_name),
            "full_name": full_name,
            "years": {},
            "monthly": {},
            "comments": {},
        }

        for year, cols in year_columns.items():
            # Annual total
            if "total_col" in cols:
                entry["years"][year] = _to_float(r[cols["total_col"]]) if len(r) > cols["total_col"] else 0.0
            elif "annual_col" in cols:
                entry["years"][year] = _to_float(r[cols["annual_col"]]) if len(r) > cols["annual_col"] else 0.0

            # Monthly breakdown
            if "monthly_start" in cols:
                monthly = []
                for m_i in range(12):
                    idx = cols["monthly_start"] + m_i
                    monthly.append(_to_float(r[idx]) if len(r) > idx else 0.0)
                entry["monthly"][year] = monthly

            # Comments
            if "comments_col" in cols:
                comment = r[cols["comments_col"]].strip() if len(r) > cols["comments_col"] else ""
                if comment:
                    entry["comments"][year] = comment

        budgets[account] = entry

    return budgets


def parse_variance_report(file_bytes: bytes) -> dict | None:
    """Parse a monthly Actual-to-Budget Expense Variance Report.

    Returns dict with:
        year       - e.g. 2026
        month      - e.g. 1 (January)
        month_label - e.g. "Jan"
        accounts   - list of dicts with:
            account, name, actual, budget, variance, variance_pct,
            ytd_actual, ytd_budget, ytd_variance, ytd_variance_pct,
            explanation, ytd_explanation
        total      - dict with total actual, budget, variance for the month
        ytd_total  - dict with total YTD actual, budget, variance
    """
    rows = _xlsx_bytes_to_rows(file_bytes)
    if not rows:
        return None

    # Detect year and month from metadata rows
    year = None
    month_str = None
    for r in rows[:12]:
        if len(r) > 2 and r[0] == "Database":
            # Row like: Database, localhost/GVTC, 2026-01
            ym = r[2].strip()
            match = re.match(r"(\d{4})-(\d{2})", ym)
            if match:
                year = int(match.group(1))
                month_str = match.group(2)
                break

    if year is None or month_str is None:
        return None

    month_num = int(month_str)
    month_label = MONTH_LABELS[month_num - 1] if 1 <= month_num <= 12 else month_str

    # Find data rows (rows with GL account numbers like 6124.32)
    accounts = []
    total_row = None

    for r in rows:
        # GL account data rows have account in col 2 (6124.xx or 6623.xx)
        if len(r) > 14:
            acct_candidate = r[2].strip() if len(r) > 2 else ""
            if re.match(r"\d{4}\.\d{2}", acct_candidate):
                actual_month = _to_float(r[4])
                budget_month = _to_float(r[5])
                variance_month = _to_float(r[6])
                variance_pct_str = r[7].strip().rstrip("%") if len(r) > 7 else "0"

                # Account name from col 10
                acct_name = r[10].strip() if len(r) > 10 else ""

                # YTD columns (11=actual, 12=budget, 13=variance, 14=pct)
                ytd_actual = _to_float(r[11]) if len(r) > 11 else 0.0
                ytd_budget = _to_float(r[12]) if len(r) > 12 else 0.0
                ytd_variance = _to_float(r[13]) if len(r) > 13 else 0.0
                ytd_pct_str = r[14].strip().rstrip("%") if len(r) > 14 else "0"

                # Explanations (col 8 for current month, col 16 for YTD)
                explanation = r[8].strip() if len(r) > 8 else ""
                ytd_explanation = r[16].strip() if len(r) > 16 else ""

                accounts.append({
                    "account": acct_candidate,
                    "name": acct_name or GL_NAMES.get(acct_candidate, acct_candidate),
                    "actual": actual_month,
                    "budget": budget_month,
                    "variance": variance_month,
                    "variance_pct": variance_pct_str,
                    "ytd_actual": ytd_actual,
                    "ytd_budget": ytd_budget,
                    "ytd_variance": ytd_variance,
                    "ytd_variance_pct": ytd_pct_str,
                    "explanation": explanation,
                    "ytd_explanation": ytd_explanation,
                })

        # Total row
        if len(r) > 4 and r[0] == "#_Department":
            total_row = {
                "actual": _to_float(r[3]),
                "budget": _to_float(r[4]),
                "variance": _to_float(r[5]),
                "ytd_actual": _to_float(r[11]) if len(r) > 11 else 0.0,
                "ytd_budget": _to_float(r[12]) if len(r) > 12 else 0.0,
                "ytd_variance": _to_float(r[13]) if len(r) > 13 else 0.0,
            }

    if not accounts:
        return None

    # Calculate totals from accounts if total_row not found
    if total_row is None:
        total_row = {
            "actual": sum(a["actual"] for a in accounts),
            "budget": sum(a["budget"] for a in accounts),
            "variance": sum(a["variance"] for a in accounts),
            "ytd_actual": sum(a["ytd_actual"] for a in accounts),
            "ytd_budget": sum(a["ytd_budget"] for a in accounts),
            "ytd_variance": sum(a["ytd_variance"] for a in accounts),
        }

    return {
        "year": year,
        "month": month_num,
        "month_label": month_label,
        "accounts": accounts,
        "total": total_row,
        "ytd_total": {
            "actual": total_row["ytd_actual"],
            "budget": total_row["ytd_budget"],
            "variance": total_row["ytd_variance"],
        },
    }


# ---------------------------------------------------------------------------
# Google Drive integration
# ---------------------------------------------------------------------------

def _build_drive_service():
    """Build an authenticated Google Drive API service from st.secrets."""
    from google.oauth2.credentials import Credentials
    from googleapiclient.discovery import build

    creds = Credentials(
        token=None,
        refresh_token=st.secrets["google_drive"]["refresh_token"],
        client_id=st.secrets["google_drive"]["client_id"],
        client_secret=st.secrets["google_drive"]["client_secret"],
        token_uri=st.secrets["google_drive"]["token_uri"],
    )
    return build("drive", "v3", credentials=creds)


@st.cache_data(ttl=300, show_spinner="Loading budget data from Google Drive...")
def load_drive_data(folder_id: str) -> dict:
    """Load all budget files from Google Drive folder.

    Returns dict with:
        budgets     - dict of GL account -> budget data (annual + monthly)
        variance    - list of parsed variance reports (one per month)
        all_years   - sorted list of budget years found
        all_accounts - sorted list of GL accounts found
    """
    service = _build_drive_service()

    results = service.files().list(
        q=f"'{folder_id}' in parents and trashed = false",
        fields="files(id, name, mimeType)",
        pageSize=100,
    ).execute()
    files = results.get("files", [])

    budgets = {}
    variance_reports = []
    export_mime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

    for f in files:
        file_id = f["id"]
        name = f["name"]
        mime = f["mimeType"]

        if mime == "application/vnd.google-apps.spreadsheet":
            resp = service.files().export(fileId=file_id, mimeType=export_mime).execute()
        elif mime == export_mime:
            resp = service.files().get_media(fileId=file_id).execute()
        else:
            continue

        file_bytes = resp if isinstance(resp, bytes) else resp.encode("latin-1")
        name_lower = name.lower()

        # Variance reports: contain "variance" or "actual" + "budget" in name
        if "variance" in name_lower or ("actual" in name_lower and "budget" in name_lower):
            vr = parse_variance_report(file_bytes)
            if vr:
                variance_reports.append(vr)
                continue

        # Budget forecast file: contains all GL accounts
        if "budget" in name_lower or "forecast" in name_lower:
            parsed = parse_forecast_file(file_bytes)
            if parsed:
                for acct, data in parsed.items():
                    if acct in budgets:
                        for y, val in data["years"].items():
                            if val != 0:
                                budgets[acct]["years"][y] = val
                        for y, monthly in data["monthly"].items():
                            if sum(monthly) != 0:
                                budgets[acct]["monthly"][y] = monthly
                    else:
                        budgets[acct] = data

    # Sort variance reports by year, month
    variance_reports.sort(key=lambda v: (v["year"], v["month"]))

    # Collect all years from budgets
    all_years = set()
    for acct, data in budgets.items():
        all_years.update(data["years"].keys())
    all_years = sorted(all_years)

    # Collect all accounts
    all_accounts = sorted(set(list(budgets.keys()) +
        [a["account"] for vr in variance_reports for a in vr["accounts"]]))

    return {
        "budgets": budgets,
        "variance": variance_reports,
        "all_years": all_years,
        "all_accounts": all_accounts,
    }


# ---------------------------------------------------------------------------
# Sidebar
# ---------------------------------------------------------------------------
with st.sidebar:
    st.markdown("### GVTC Department 44")
    st.markdown("**Web Management**")
    st.markdown("Entity: P0100")
    st.divider()

# Load data
try:
    _ = st.secrets["google_drive"]
    has_secrets = True
except (KeyError, FileNotFoundError):
    has_secrets = False

if not has_secrets:
    st.error(
        "**Google Drive credentials not configured.** "
        "Add `[google_drive]` section to `.streamlit/secrets.toml` or Streamlit Cloud secrets "
        "with `refresh_token`, `client_id`, `client_secret`, `token_uri`, and `folder_id`."
    )
    st.stop()

try:
    data = load_drive_data(st.secrets["google_drive"]["folder_id"])
except Exception as e:
    st.error(f"**Failed to load data from Google Drive:** {e}")
    st.stop()

budgets = data["budgets"]
variance_reports = data["variance"]
all_years = data["all_years"]
all_accounts = data["all_accounts"]

if not budgets and not variance_reports:
    st.warning("No valid budget or variance data found in the Google Drive folder.")
    st.stop()

# Build actuals data: month -> account -> {actual, budget, variance}
actuals_by_month = {}
for vr in variance_reports:
    key = (vr["year"], vr["month"])
    actuals_by_month[key] = {}
    for a in vr["accounts"]:
        actuals_by_month[key][a["account"]] = a

# Latest variance report for YTD totals
latest_vr = variance_reports[-1] if variance_reports else None

with st.sidebar:
    if st.button("🔄 Refresh Data"):
        load_drive_data.clear()
        st.rerun()

    st.divider()

    page = st.radio(
        "Navigate",
        ["Budget Overview", "Monthly Actuals", "Variance Analysis", "Year Comparison"],
        index=0,
    )

    st.divider()

    # Category filter
    _name_to_acct = {GL_NAMES.get(a, a): a for a in sorted(all_accounts)}
    all_short_names = list(_name_to_acct.keys())

    if st.button("Select All Categories"):
        st.session_state["cat_filter"] = all_short_names

    selected_names = st.multiselect(
        "Filter Categories",
        all_short_names,
        default=st.session_state.get("cat_filter", all_short_names),
        key="cat_filter",
    )

    if not selected_names:
        selected_names = all_short_names
        st.session_state["cat_filter"] = all_short_names
        st.rerun()

filtered_accounts = [_name_to_acct[n] for n in selected_names if n in _name_to_acct]


def short_name(acct: str) -> str:
    return GL_NAMES.get(acct, acct)


# ---------------------------------------------------------------------------
# Page 1: Budget Overview
# ---------------------------------------------------------------------------
if page == "Budget Overview":
    st.title("Budget Overview")

    # Top-level KPIs
    if latest_vr:
        ytd = latest_vr["ytd_total"]
        through_month = latest_vr["month_label"]
        yr = latest_vr["year"]

        # Annual budget for the current year
        annual_budget = sum(
            budgets.get(a, {}).get("years", {}).get(yr, 0)
            for a in filtered_accounts
        )

        c1, c2, c3, c4 = st.columns(4)
        c1.metric(f"{yr} Annual Budget", f"${annual_budget:,.0f}")
        c2.metric(f"YTD Spent (thru {through_month})",
                  f"${ytd['actual']:,.0f}")
        c3.metric(f"YTD Budget",
                  f"${ytd['budget']:,.0f}")

        remaining = annual_budget - ytd['actual']
        c4.metric("Remaining", f"${remaining:,.0f}",
                  delta=f"${ytd['variance']:,.0f} {'under' if ytd['variance'] > 0 else 'over'} YTD")

        st.divider()

        # Progress bar
        pct_used = ytd['actual'] / annual_budget * 100 if annual_budget else 0
        month_pct = latest_vr["month"] / 12 * 100
        st.markdown(f"**Budget Utilization:** {pct_used:.1f}% used ({month_pct:.0f}% through the year)")
        st.progress(min(pct_used / 100, 1.0))

        if pct_used > month_pct + 10:
            st.warning(f"⚠️ Spending is ahead of pace. {pct_used:.1f}% used but only {month_pct:.0f}% through the year.")
        elif pct_used < month_pct - 15:
            st.info(f"💡 Spending is well under pace. {pct_used:.1f}% used, {month_pct:.0f}% through the year.")

        st.divider()

    # Budget by category (bar chart)
    st.subheader("Annual Budget by Category")

    if all_years:
        selected_year = st.selectbox("Year", all_years, index=len(all_years) - 1,
                                      key="overview_year")
    else:
        selected_year = None

    if selected_year:
        cat_data = []
        for acct in filtered_accounts:
            budget_val = budgets.get(acct, {}).get("years", {}).get(selected_year, 0)
            row = {"Category": short_name(acct), "Budget": budget_val}

            # Add YTD actual if we have variance data for this year
            ytd_actual = 0
            for vr in variance_reports:
                if vr["year"] == selected_year:
                    for a in vr["accounts"]:
                        if a["account"] == acct:
                            ytd_actual = a["ytd_actual"]
            row["YTD Actual"] = ytd_actual
            cat_data.append(row)

        df_cat = pd.DataFrame(cat_data)

        if not df_cat.empty:
            fig = px.bar(
                df_cat.melt(id_vars="Category", var_name="Type", value_name="Amount"),
                x="Category", y="Amount", color="Type",
                barmode="group",
                color_discrete_map={"Budget": "#5B8DEF", "YTD Actual": "#2ED573"},
                text_auto="$,.0f",
            )
            fig.update_layout(**PLOTLY_LAYOUT, height=450)
            fig.update_traces(textposition="outside", textfont_size=10)
            st.plotly_chart(fig, use_container_width=True)

    # Budget allocation donut
    if selected_year:
        st.subheader("Budget Allocation")
        alloc_data = []
        for acct in filtered_accounts:
            val = budgets.get(acct, {}).get("years", {}).get(selected_year, 0)
            if val > 0:
                alloc_data.append({"Category": short_name(acct), "Budget": val})

        if alloc_data:
            df_alloc = pd.DataFrame(alloc_data)
            fig_donut = px.pie(
                df_alloc, values="Budget", names="Category",
                color_discrete_sequence=COLORS,
                hole=0.45,
            )
            fig_donut.update_layout(**PLOTLY_LAYOUT, height=400, showlegend=True)
            fig_donut.update_traces(textinfo="percent", textposition="outside")
            st.plotly_chart(fig_donut, use_container_width=True)


# ---------------------------------------------------------------------------
# Page 2: Monthly Actuals
# ---------------------------------------------------------------------------
elif page == "Monthly Actuals":
    st.title("Monthly Actuals vs Budget")

    if not variance_reports:
        st.info("No variance reports uploaded yet. Upload monthly Actual-to-Budget reports to see actuals data.")
        st.stop()

    # Monthly bars: actual vs budget per month
    monthly_data = []
    for vr in variance_reports:
        for a in vr["accounts"]:
            if a["account"] in filtered_accounts:
                monthly_data.append({
                    "Month": vr["month_label"],
                    "Month_Num": vr["month"],
                    "Category": short_name(a["account"]),
                    "Actual": a["actual"],
                    "Budget": a["budget"],
                })

    df_monthly = pd.DataFrame(monthly_data)

    if df_monthly.empty:
        st.info("No actuals data for the selected categories.")
        st.stop()

    # Aggregate by month
    df_month_totals = df_monthly.groupby(["Month", "Month_Num"]).agg(
        Actual=("Actual", "sum"),
        Budget=("Budget", "sum"),
    ).reset_index().sort_values("Month_Num")
    df_month_totals["Variance"] = df_month_totals["Budget"] - df_month_totals["Actual"]

    st.subheader("Total Spend by Month")
    fig_monthly = go.Figure()
    fig_monthly.add_trace(go.Bar(
        x=df_month_totals["Month"], y=df_month_totals["Actual"],
        name="Actual", marker_color="#2ED573",
        text=[f"${v:,.0f}" for v in df_month_totals["Actual"]],
        textposition="outside",
    ))
    fig_monthly.add_trace(go.Bar(
        x=df_month_totals["Month"], y=df_month_totals["Budget"],
        name="Budget", marker_color="#5B8DEF",
        text=[f"${v:,.0f}" for v in df_month_totals["Budget"]],
        textposition="outside",
    ))
    fig_monthly.update_layout(**PLOTLY_LAYOUT, barmode="group", height=400)
    st.plotly_chart(fig_monthly, use_container_width=True)

    st.divider()

    # Per-category monthly breakdown
    st.subheader("By Category")
    for acct in filtered_accounts:
        name = short_name(acct)
        acct_data = df_monthly[df_monthly["Category"] == name]
        if acct_data.empty:
            continue

        with st.expander(f"{name} ({acct})"):
            fig_cat = go.Figure()
            fig_cat.add_trace(go.Bar(
                x=acct_data["Month"], y=acct_data["Actual"],
                name="Actual", marker_color="#2ED573",
            ))
            fig_cat.add_trace(go.Bar(
                x=acct_data["Month"], y=acct_data["Budget"],
                name="Budget", marker_color="#5B8DEF",
            ))
            fig_cat.update_layout(**PLOTLY_LAYOUT, barmode="group", height=300,
                                   title=name)
            st.plotly_chart(fig_cat, use_container_width=True)


# ---------------------------------------------------------------------------
# Page 3: Variance Analysis
# ---------------------------------------------------------------------------
elif page == "Variance Analysis":
    st.title("Variance Analysis")

    if not variance_reports:
        st.info("No variance reports uploaded yet.")
        st.stop()

    # Month selector
    vr_options = {f"{vr['month_label']} {vr['year']}": vr for vr in variance_reports}
    selected_vr_label = st.selectbox("Report Period", list(vr_options.keys()),
                                      index=len(vr_options) - 1)
    vr = vr_options[selected_vr_label]

    # Summary cards
    c1, c2, c3 = st.columns(3)
    c1.metric(f"{vr['month_label']} Actual", f"${vr['total']['actual']:,.0f}")
    c2.metric(f"{vr['month_label']} Budget", f"${vr['total']['budget']:,.0f}")
    variance_val = vr['total']['variance']
    c3.metric("Variance",
              f"${abs(variance_val):,.0f} {'under' if variance_val > 0 else 'over'}",
              delta=f"${variance_val:+,.0f}",
              delta_color="normal")

    st.divider()

    # YTD summary
    if vr.get("ytd_total"):
        st.subheader(f"Year-to-Date (thru {vr['month_label']})")
        c1, c2, c3 = st.columns(3)
        c1.metric("YTD Actual", f"${vr['ytd_total']['actual']:,.0f}")
        c2.metric("YTD Budget", f"${vr['ytd_total']['budget']:,.0f}")
        ytd_var = vr['ytd_total']['variance']
        c3.metric("YTD Variance",
                  f"${abs(ytd_var):,.0f} {'under' if ytd_var > 0 else 'over'}",
                  delta=f"${ytd_var:+,.0f}",
                  delta_color="normal")

    st.divider()

    # Variance by category
    st.subheader("Variance by Category")
    var_data = []
    for a in vr["accounts"]:
        if a["account"] in filtered_accounts:
            var_data.append({
                "Category": short_name(a["account"]),
                "Account": a["account"],
                "Actual": a["actual"],
                "Budget": a["budget"],
                "Variance ($)": a["variance"],
                "Variance (%)": a["variance_pct"],
                "Explanation": a["explanation"],
            })

    df_var = pd.DataFrame(var_data)

    if not df_var.empty:
        # Horizontal bar chart
        fig_var = px.bar(
            df_var, x="Variance ($)", y="Category", orientation="h",
            color=df_var["Variance ($)"].apply(lambda v: "Under Budget" if v > 0 else "Over Budget"),
            color_discrete_map={"Under Budget": "#2ED573", "Over Budget": "#FC5C65"},
            text_auto="$,.0f",
        )
        fig_var.update_layout(**PLOTLY_LAYOUT, height=max(300, len(df_var) * 45),
                               showlegend=True)
        fig_var.update_traces(textposition="outside", textfont_size=11)
        st.plotly_chart(fig_var, use_container_width=True)

        st.divider()

        # Detail table with explanations
        st.subheader("Detail")
        display_cols = ["Category", "Actual", "Budget", "Variance ($)", "Variance (%)"]
        has_explanations = any(a["explanation"] for a in vr["accounts"])
        if has_explanations:
            display_cols.append("Explanation")

        def color_variance(val):
            if isinstance(val, (int, float)):
                if val > 0:
                    return "color: #2ED573"
                elif val < 0:
                    return "color: #FC5C65"
            return ""

        fmt = {"Actual": "${:,.0f}", "Budget": "${:,.0f}", "Variance ($)": "${:+,.0f}"}
        styled = df_var[display_cols].style.format(fmt).map(
            color_variance, subset=["Variance ($)"]
        )
        st.dataframe(styled, use_container_width=True, hide_index=True)

    # YTD variance table
    st.divider()
    st.subheader(f"Year-to-Date Detail (thru {vr['month_label']})")
    ytd_data = []
    for a in vr["accounts"]:
        if a["account"] in filtered_accounts:
            ytd_data.append({
                "Category": short_name(a["account"]),
                "YTD Actual": a["ytd_actual"],
                "YTD Budget": a["ytd_budget"],
                "YTD Variance ($)": a["ytd_variance"],
                "YTD Variance (%)": a["ytd_variance_pct"],
            })

    df_ytd = pd.DataFrame(ytd_data)
    if not df_ytd.empty:
        fmt_ytd = {
            "YTD Actual": "${:,.0f}",
            "YTD Budget": "${:,.0f}",
            "YTD Variance ($)": "${:+,.0f}",
        }
        styled_ytd = df_ytd.style.format(fmt_ytd).map(
            color_variance, subset=["YTD Variance ($)"]
        )
        st.dataframe(styled_ytd, use_container_width=True, hide_index=True)


# ---------------------------------------------------------------------------
# Page 4: Year Comparison
# ---------------------------------------------------------------------------
elif page == "Year Comparison":
    st.title("Year-over-Year Budget Comparison")

    if not all_years:
        st.info("No budget data loaded.")
        st.stop()

    # Summary cards per year
    year_strs = [str(y) for y in all_years]
    cols = st.columns(len(year_strs))
    for col, year in zip(cols, all_years):
        total = sum(budgets.get(a, {}).get("years", {}).get(year, 0) for a in filtered_accounts)
        col.metric(f"{year} Budget", f"${total:,.0f}")

    st.divider()

    # YoY grouped bar chart
    yoy_data = []
    for acct in filtered_accounts:
        for year in all_years:
            val = budgets.get(acct, {}).get("years", {}).get(year, 0)
            yoy_data.append({
                "Category": short_name(acct),
                "Year": str(year),
                "Budget": val,
            })

    df_yoy = pd.DataFrame(yoy_data)

    if not df_yoy.empty:
        yoy_colors = ["#2ED573", "#5B8DEF", "#F7B731", "#FC5C65"][:len(all_years)]
        fig_yoy = px.bar(
            df_yoy, x="Category", y="Budget", color="Year",
            barmode="group",
            color_discrete_sequence=yoy_colors,
            text_auto="$,.0f",
        )
        fig_yoy.update_layout(**PLOTLY_LAYOUT, height=450)
        fig_yoy.update_traces(textposition="outside", textfont_size=10)
        st.plotly_chart(fig_yoy, use_container_width=True)

    st.divider()

    # YoY table with change columns
    st.subheader("Year-over-Year Summary")
    yoy_rows = []
    for acct in filtered_accounts:
        row = {"Category": short_name(acct)}
        year_totals = {}
        for y in all_years:
            t = budgets.get(acct, {}).get("years", {}).get(y, 0)
            row[str(y)] = t
            year_totals[y] = t
        for i in range(1, len(all_years)):
            prev_y = all_years[i - 1]
            curr_y = all_years[i]
            short_prev = str(prev_y)[-2:]
            short_curr = str(curr_y)[-2:]
            row[f"{short_prev}→{short_curr}"] = year_totals[curr_y] - year_totals[prev_y]
        yoy_rows.append(row)

    df_yoy_table = pd.DataFrame(yoy_rows)

    if not df_yoy_table.empty:
        change_cols = [c for c in df_yoy_table.columns if "→" in c]

        def color_change(val):
            if isinstance(val, (int, float)):
                if val > 0:
                    return "color: #FC5C65"  # increase = red (more spending)
                elif val < 0:
                    return "color: #2ED573"  # decrease = green (less spending)
            return ""

        fmt_cols = {c: "${:,.0f}" for c in df_yoy_table.columns if c != "Category"}
        for c in change_cols:
            fmt_cols[c] = "${:+,.0f}"

        styled = df_yoy_table.style.format(fmt_cols).map(
            color_change, subset=change_cols
        )
        st.dataframe(styled, use_container_width=True, hide_index=True)
