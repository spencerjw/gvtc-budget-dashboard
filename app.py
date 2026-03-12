import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import csv
import io
import os
import re
import shutil

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------
st.set_page_config(
    page_title="GVTC Dept 44 – Web Management Budget",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

DATA_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data")

MONTHS = [f"{m:02d}" for m in range(1, 13)]
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


# ---------------------------------------------------------------------------
# xlsx2csv helper – use Python API, fall back to CLI, then error
# ---------------------------------------------------------------------------
def _xlsx_bytes_to_rows(file_bytes: bytes) -> list[list[str]]:
    """Convert xlsx bytes to a list of CSV rows using xlsx2csv."""
    try:
        from xlsx2csv import Xlsx2csv
    except ImportError:
        # Try CLI fallback
        xlsx2csv_bin = shutil.which("xlsx2csv")
        if xlsx2csv_bin is None:
            st.error(
                "**xlsx2csv** is not installed. Install it with: "
                "`pip install xlsx2csv`"
            )
            return []
        import subprocess
        import tempfile
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
            tmp.write(file_bytes)
            tmp_path = tmp.name
        try:
            result = subprocess.run(
                [xlsx2csv_bin, tmp_path],
                capture_output=True, text=True, timeout=30,
            )
            reader = csv.reader(io.StringIO(result.stdout))
            return list(reader)
        finally:
            os.unlink(tmp_path)

    # Use Python API
    buf = io.BytesIO(file_bytes)
    out = io.StringIO()
    Xlsx2csv(buf).convert(out)
    out.seek(0)
    reader = csv.reader(out)
    return list(reader)


# ---------------------------------------------------------------------------
# Data loading helpers
# ---------------------------------------------------------------------------
def _to_float(val: str) -> float:
    try:
        return float(val.replace(",", "").strip())
    except (ValueError, AttributeError):
        return 0.0


def parse_gl_bytes(file_bytes: bytes) -> dict | None:
    """Parse a single GL account xlsx export from bytes.

    Returns dict with keys:
        account   – e.g. '6623.18'
        desc      – full GL description from Line000
        line_items – list of dicts with name, annual, monthly, per year
    """
    rows = _xlsx_bytes_to_rows(file_bytes)
    if not rows:
        return None

    account = None
    for r in rows:
        if len(r) > 1 and r[0] == "" and r[1] == "Account":
            account = r[2].strip() if len(r) > 2 else None
            break

    # Find the header row (contains "Line Item")
    header_idx = None
    for i, r in enumerate(rows):
        if any("Line Item" in c for c in r):
            header_idx = i
            break

    if header_idx is None or account is None:
        return None

    # Line000 description
    desc = ""
    line_items = []

    for r in rows[header_idx + 2:]:  # skip header + "Working Budget" row
        if len(r) < 5:
            continue
        line_id = r[2].strip() if len(r) > 2 else ""
        if not re.match(r"Line\d+", line_id):
            continue

        name = r[4].strip() if len(r) > 4 else ""

        if line_id == "Line000":
            desc = name
            continue

        # Skip empty items (no name and zero annual)
        annual_2026 = _to_float(r[8]) if len(r) > 8 else 0
        if not name and annual_2026 == 0:
            continue

        if not name:
            continue

        monthly_2026 = [_to_float(r[9 + i]) if len(r) > 9 + i else 0.0 for i in range(12)]
        annual_2027 = _to_float(r[22]) if len(r) > 22 else 0.0
        annual_2028 = _to_float(r[25]) if len(r) > 25 else 0.0

        line_items.append({
            "name": name,
            "annual_2026": annual_2026,
            "monthly_2026": monthly_2026,
            "annual_2027": annual_2027,
            "annual_2028": annual_2028,
        })

    return {
        "account": account,
        "desc": desc,
        "line_items": line_items,
    }


def parse_master_bytes(file_bytes: bytes) -> list[dict]:
    """Parse Dept_44_Budget.xlsx for 2025 budget vs actual data."""
    rows = _xlsx_bytes_to_rows(file_bytes)
    if not rows:
        return []

    records = []
    current_account = None
    current_desc = None

    for r in rows:
        if len(r) < 3:
            continue

        # Detect category header rows (have account number in col 0)
        if r[0].strip() and re.match(r"\d{4}\.\d{2}", r[0].strip()):
            current_account = r[0].strip()
            current_desc = r[1].strip() if len(r) > 1 else ""
            # Category summary row – extract budget/actual per month
            budgets = []
            actuals = []
            for m in range(12):
                b_idx = 2 + m * 2
                a_idx = 3 + m * 2
                budgets.append(_to_float(r[b_idx]) if len(r) > b_idx else 0)
                actuals.append(_to_float(r[a_idx]) if len(r) > a_idx else 0)
            yearly_budget = _to_float(r[26]) if len(r) > 26 else sum(budgets)
            yearly_actual = _to_float(r[27]) if len(r) > 27 else sum(actuals)
            records.append({
                "account": current_account,
                "category": current_desc,
                "line_item": "(Category Total)",
                "is_category": True,
                "monthly_budget": budgets,
                "monthly_actual": actuals,
                "yearly_budget": yearly_budget,
                "yearly_actual": yearly_actual,
            })
            continue

        # Detect sub-line items (no account number, name in col 0 or 1)
        if current_account and not r[0].strip():
            name = r[1].strip() if len(r) > 1 and r[1].strip() else ""
            if not name:
                name = r[0].strip()
            if not name:
                continue
            if name in ("Account", "", "Net Income"):
                continue

            budgets = []
            actuals = []
            for m in range(12):
                b_idx = 2 + m * 2
                a_idx = 3 + m * 2
                budgets.append(_to_float(r[b_idx]) if len(r) > b_idx else 0)
                actuals.append(_to_float(r[a_idx]) if len(r) > a_idx else 0)

            if sum(budgets) == 0 and sum(actuals) == 0:
                continue

            yearly_budget = _to_float(r[26]) if len(r) > 26 else sum(budgets)
            yearly_actual = _to_float(r[27]) if len(r) > 27 else sum(actuals)
            records.append({
                "account": current_account,
                "category": current_desc,
                "line_item": name,
                "is_category": False,
                "monthly_budget": budgets,
                "monthly_actual": actuals,
                "yearly_budget": yearly_budget,
                "yearly_actual": yearly_actual,
            })
        elif r[0].strip() and not re.match(r"\d{4}\.\d{2}", r[0].strip()):
            if current_account:
                name = r[0].strip()
                if name in ("Account", "Net Income"):
                    continue
                budgets = []
                actuals = []
                for m in range(12):
                    b_idx = 2 + m * 2
                    a_idx = 3 + m * 2
                    budgets.append(_to_float(r[b_idx]) if len(r) > b_idx else 0)
                    actuals.append(_to_float(r[a_idx]) if len(r) > a_idx else 0)
                if sum(budgets) == 0 and sum(actuals) == 0:
                    continue
                yearly_budget = _to_float(r[26]) if len(r) > 26 else sum(budgets)
                yearly_actual = _to_float(r[27]) if len(r) > 27 else sum(actuals)
                records.append({
                    "account": current_account,
                    "category": current_desc,
                    "line_item": name,
                    "is_category": False,
                    "monthly_budget": budgets,
                    "monthly_actual": actuals,
                    "yearly_budget": yearly_budget,
                    "yearly_actual": yearly_actual,
                })

    return records


def category_label(account: str) -> str:
    name = GL_NAMES.get(account, account)
    return f"{account} – {name}"


# ---------------------------------------------------------------------------
# Data source: uploaded files or local data/ directory
# ---------------------------------------------------------------------------
def _load_local_files() -> dict[str, bytes] | None:
    """If xlsx files exist in data/ subdirectory, load them."""
    if not os.path.isdir(DATA_DIR):
        # Also check app root for backwards compat during migration
        app_dir = os.path.dirname(os.path.abspath(__file__))
        xlsx_in_root = [f for f in os.listdir(app_dir) if f.endswith(".xlsx")]
        if xlsx_in_root:
            return {f: open(os.path.join(app_dir, f), "rb").read() for f in xlsx_in_root}
        return None
    xlsx_files = [f for f in os.listdir(DATA_DIR) if f.endswith(".xlsx")]
    if not xlsx_files:
        return None
    return {f: open(os.path.join(DATA_DIR, f), "rb").read() for f in xlsx_files}


def _get_file_data() -> dict[str, bytes] | None:
    """Return dict of filename->bytes from uploads or local files."""
    if "uploaded_files" in st.session_state and st.session_state["uploaded_files"]:
        return st.session_state["uploaded_files"]
    return _load_local_files()


def _process_files(file_data: dict[str, bytes]) -> tuple[dict, list]:
    """Parse all files and return (gl_data dict, actuals_2025 list)."""
    gl_data = {}
    actuals_2025 = []

    for fname, fbytes in sorted(file_data.items()):
        if fname == "Dept_44_Budget.xlsx":
            actuals_2025 = parse_master_bytes(fbytes)
        else:
            parsed = parse_gl_bytes(fbytes)
            if parsed:
                gl_data[parsed["account"]] = parsed

    return gl_data, actuals_2025


# ---------------------------------------------------------------------------
# Sidebar: file uploader + navigation
# ---------------------------------------------------------------------------
with st.sidebar:
    st.markdown("### GVTC Department 44")
    st.markdown("**Web Management**")
    st.markdown("Entity: P0100")
    st.divider()

    # File uploader
    uploaded = st.file_uploader(
        "Upload GL Account xlsx files",
        type=["xlsx"],
        accept_multiple_files=True,
        help="Upload your GVTC GL account Excel exports and the Dept_44_Budget.xlsx master file.",
    )

    if uploaded:
        st.session_state["uploaded_files"] = {f.name: f.read() for f in uploaded}

    # Show loaded file status
    file_data = _get_file_data()
    if file_data:
        source = "uploaded" if "uploaded_files" in st.session_state and st.session_state["uploaded_files"] else "local"
        st.caption(f"{len(file_data)} file(s) loaded ({source})")
        with st.expander("Loaded files", expanded=False):
            for fname in sorted(file_data.keys()):
                st.text(f"  {fname}")
        if st.button("Clear Data", type="secondary"):
            st.session_state.pop("uploaded_files", None)
            st.rerun()

    st.divider()


# ---------------------------------------------------------------------------
# Welcome page when no data is loaded
# ---------------------------------------------------------------------------
if file_data is None:
    st.title("GVTC Dept 44 – Budget Dashboard")
    st.markdown("---")
    st.markdown(
        """
        ### Welcome

        No budget data is currently loaded. To get started, upload your GVTC
        GL account Excel exports using the **file uploader in the sidebar**.

        **Expected files:**
        - **Individual GL account exports** (e.g. `Contract_labor_and_services.xlsx`,
          `Cellular_phone.xlsx`, `GPC_Software_licenses.xlsx`, etc.)
          — one file per GL account with line-item budget details for 2026-2028.
        - **Master budget file** (`Dept_44_Budget.xlsx`)
          — contains 2025 budget vs actuals data across all categories.

        **File format:** Standard GVTC budget Excel exports with account numbers,
        line items, and monthly/annual budget columns. The parser expects the
        standard layout with an "Account" field and "Line Item" headers.

        **Local development:** Place xlsx files in a `data/` subdirectory and
        they will be loaded automatically.
        """
    )
    st.stop()


# ---------------------------------------------------------------------------
# Load & process data
# ---------------------------------------------------------------------------
gl_data, actuals_2025 = _process_files(file_data)

if not gl_data:
    st.warning(
        "No valid GL account data found in the uploaded files. "
        "Please check that you uploaded the correct GVTC budget xlsx exports."
    )
    st.stop()

# Build summary DataFrames
summary_rows = []
for acct, info in sorted(gl_data.items()):
    total_2026 = sum(li["annual_2026"] for li in info["line_items"])
    total_2027 = sum(li["annual_2027"] for li in info["line_items"])
    total_2028 = sum(li["annual_2028"] for li in info["line_items"])
    summary_rows.append({
        "account": acct,
        "category": category_label(acct),
        "short_name": GL_NAMES.get(acct, acct),
        "2026": total_2026,
        "2027": total_2027,
        "2028": total_2028,
    })

df_summary = pd.DataFrame(summary_rows)

# ---------------------------------------------------------------------------
# Sidebar: navigation & filters (continued)
# ---------------------------------------------------------------------------
with st.sidebar:
    page = st.radio(
        "Navigate",
        ["Budget Overview", "Monthly View", "Line Item Detail", "2025 Actuals"],
        index=0,
    )

    st.divider()
    selected_year = st.selectbox("Year", [2026, 2027, 2028], index=0)

    all_categories = [category_label(a) for a in sorted(gl_data.keys())]
    selected_categories = st.multiselect(
        "Filter Categories",
        all_categories,
        default=all_categories,
    )

    st.divider()
    total_budget = df_summary[str(selected_year)].sum()
    st.metric("Total Budget", f"${total_budget:,.0f}", delta=None)

# Filter helper
filtered_accounts = [
    c.split(" – ")[0] for c in selected_categories
]
df_filtered = df_summary[df_summary["account"].isin(filtered_accounts)]


# ---------------------------------------------------------------------------
# Page 1: Budget Overview
# ---------------------------------------------------------------------------
if page == "Budget Overview":
    st.title("Budget Overview")

    # Summary cards
    c1, c2, c3 = st.columns(3)
    for col, year in zip([c1, c2, c3], ["2026", "2027", "2028"]):
        val = df_filtered[year].sum()
        prev_val = df_filtered[str(int(year) - 1)].sum() if year != "2026" else None
        delta = f"${val - prev_val:+,.0f}" if prev_val is not None else None
        col.metric(f"{year} Total Budget", f"${val:,.0f}", delta=delta)

    st.divider()

    col_left, col_right = st.columns(2)

    # Horizontal bar chart – sorted by budget
    with col_left:
        st.subheader("Budget by GL Category")
        df_bar = df_filtered[["short_name", str(selected_year)]].copy()
        df_bar.columns = ["Category", "Budget"]
        df_bar = df_bar.sort_values("Budget")
        fig_bar = px.bar(
            df_bar, x="Budget", y="Category", orientation="h",
            color="Category", color_discrete_sequence=COLORS,
            text_auto="$,.0f",
        )
        fig_bar.update_layout(**PLOTLY_LAYOUT, showlegend=False, height=450)
        fig_bar.update_traces(textposition="outside")
        st.plotly_chart(fig_bar, use_container_width=True)

    # Pie chart
    with col_right:
        st.subheader("Budget Distribution")
        df_pie = df_filtered[["short_name", str(selected_year)]].copy()
        df_pie.columns = ["Category", "Budget"]
        df_pie = df_pie[df_pie["Budget"] > 0]
        fig_pie = px.pie(
            df_pie, values="Budget", names="Category",
            color_discrete_sequence=COLORS,
            hole=0.4,
        )
        fig_pie.update_layout(**PLOTLY_LAYOUT, height=450)
        fig_pie.update_traces(textinfo="label+percent", textposition="outside")
        st.plotly_chart(fig_pie, use_container_width=True)

    # Year-over-year comparison
    st.subheader("Year-over-Year Comparison")
    df_yoy = df_filtered[["short_name", "2026", "2027", "2028"]].melt(
        id_vars="short_name", var_name="Year", value_name="Budget"
    )
    fig_yoy = px.bar(
        df_yoy, x="short_name", y="Budget", color="Year",
        barmode="group",
        color_discrete_sequence=["#5B8DEF", "#F7B731", "#FC5C65"],
        labels={"short_name": "Category", "Budget": "Budget ($)"},
        text_auto="$,.0f",
    )
    fig_yoy.update_layout(**PLOTLY_LAYOUT, height=420)
    fig_yoy.update_traces(textposition="outside", textfont_size=10)
    st.plotly_chart(fig_yoy, use_container_width=True)


# ---------------------------------------------------------------------------
# Page 2: Monthly View
# ---------------------------------------------------------------------------
elif page == "Monthly View":
    st.title(f"Monthly Budget Breakdown – {selected_year}")

    # Build monthly data per category
    monthly_data = []
    for acct in filtered_accounts:
        if acct not in gl_data:
            continue
        info = gl_data[acct]
        monthly_totals = [0.0] * 12
        for li in info["line_items"]:
            if selected_year == 2026:
                for m in range(12):
                    monthly_totals[m] += li["monthly_2026"][m]
            else:
                # 2027/2028 only have annual totals, spread evenly
                annual = li[f"annual_{selected_year}"]
                for m in range(12):
                    monthly_totals[m] += annual / 12
        for m in range(12):
            monthly_data.append({
                "Month": MONTH_LABELS[m],
                "Month_Num": m + 1,
                "Category": GL_NAMES.get(acct, acct),
                "Budget": monthly_totals[m],
            })

    df_monthly = pd.DataFrame(monthly_data)

    # Stacked area chart
    st.subheader("Monthly Spending by Category")
    fig_area = px.area(
        df_monthly, x="Month", y="Budget", color="Category",
        color_discrete_sequence=COLORS,
        labels={"Budget": "Budget ($)"},
        category_orders={"Month": MONTH_LABELS},
    )
    fig_area.update_layout(**PLOTLY_LAYOUT, height=450)
    st.plotly_chart(fig_area, use_container_width=True)

    # Monthly totals bar chart highlighting concentrated months
    st.subheader("Total Monthly Spending")
    df_month_total = df_monthly.groupby(["Month", "Month_Num"])["Budget"].sum().reset_index()
    df_month_total = df_month_total.sort_values("Month_Num")
    avg_spend = df_month_total["Budget"].mean()
    df_month_total["Highlight"] = df_month_total["Budget"].apply(
        lambda x: "Above Average" if x > avg_spend * 1.2 else "Normal"
    )
    fig_totals = px.bar(
        df_month_total, x="Month", y="Budget", color="Highlight",
        color_discrete_map={"Above Average": "#FC5C65", "Normal": "#5B8DEF"},
        text_auto="$,.0f",
        category_orders={"Month": MONTH_LABELS},
    )
    fig_totals.update_layout(**PLOTLY_LAYOUT, height=350, showlegend=True)
    fig_totals.update_traces(textposition="outside")
    st.plotly_chart(fig_totals, use_container_width=True)

    # Monthly breakdown table
    st.subheader("Monthly Breakdown Table")
    pivot_data = df_monthly.pivot_table(
        index="Category", columns="Month", values="Budget", aggfunc="sum"
    )
    pivot_data = pivot_data[MONTH_LABELS]
    pivot_data["Annual"] = pivot_data.sum(axis=1)
    pivot_data.loc["TOTAL"] = pivot_data.sum()
    st.dataframe(
        pivot_data.style.format("${:,.0f}"),
        use_container_width=True,
        height=450,
    )


# ---------------------------------------------------------------------------
# Page 3: Line Item Detail
# ---------------------------------------------------------------------------
elif page == "Line Item Detail":
    st.title(f"Line Item Detail – {selected_year}")

    for acct in sorted(filtered_accounts):
        if acct not in gl_data:
            continue
        info = gl_data[acct]
        label = category_label(acct)
        total = sum(li[f"annual_{selected_year}"] for li in info["line_items"])

        with st.expander(f"{label}  —  ${total:,.0f}", expanded=False):
            rows = []
            for li in info["line_items"]:
                annual = li[f"annual_{selected_year}"]
                if annual == 0 and not li["name"]:
                    continue
                row = {"Line Item": li["name"], "Annual": annual}
                if selected_year == 2026:
                    for m in range(12):
                        row[MONTH_LABELS[m]] = li["monthly_2026"][m]
                else:
                    for m in range(12):
                        row[MONTH_LABELS[m]] = annual / 12

                # YoY change
                prev_year = selected_year - 1
                if prev_year >= 2026:
                    prev = li[f"annual_{prev_year}"]
                    change = annual - prev
                    row["YoY Change"] = change
                elif selected_year == 2026:
                    row["vs 2027"] = li["annual_2027"] - annual
                    row["vs 2028"] = li["annual_2028"] - annual

                rows.append(row)

            if not rows:
                st.info("No line items with budget.")
                continue

            df_detail = pd.DataFrame(rows)

            # Highlight changes
            def color_change(val):
                if not isinstance(val, (int, float)):
                    return ""
                if val > 0:
                    return "color: #2ED573"
                elif val < 0:
                    return "color: #FC5C65"
                return ""

            change_cols = [c for c in df_detail.columns if "Change" in c or "vs " in c]
            format_dict = {c: "${:,.0f}" for c in df_detail.columns if c != "Line Item"}
            styled = df_detail.style.format(format_dict)
            if change_cols:
                styled = styled.map(color_change, subset=change_cols)

            st.dataframe(styled, use_container_width=True, hide_index=True)

            # Mini bar chart for this category
            fig_items = px.bar(
                df_detail, x="Line Item", y="Annual",
                color_discrete_sequence=COLORS,
                text_auto="$,.0f",
            )
            fig_items.update_layout(**PLOTLY_LAYOUT, height=300, showlegend=False)
            fig_items.update_traces(textposition="outside")
            st.plotly_chart(fig_items, use_container_width=True)

    # Year comparison summary
    st.divider()
    st.subheader("Year-over-Year Summary")
    yoy_rows = []
    for acct in sorted(filtered_accounts):
        if acct not in gl_data:
            continue
        info = gl_data[acct]
        t26 = sum(li["annual_2026"] for li in info["line_items"])
        t27 = sum(li["annual_2027"] for li in info["line_items"])
        t28 = sum(li["annual_2028"] for li in info["line_items"])
        yoy_rows.append({
            "Category": GL_NAMES.get(acct, acct),
            "2026": t26,
            "2027": t27,
            "2028": t28,
            "26→27": t27 - t26,
            "27→28": t28 - t27,
        })
    df_yoy = pd.DataFrame(yoy_rows)
    totals = df_yoy.select_dtypes(include="number").sum()
    totals["Category"] = "TOTAL"
    df_yoy = pd.concat([df_yoy, pd.DataFrame([totals])], ignore_index=True)

    def color_change(val):
        if not isinstance(val, (int, float)):
            return ""
        if val > 0:
            return "color: #2ED573"
        elif val < 0:
            return "color: #FC5C65"
        return ""

    fmt = {c: "${:,.0f}" for c in df_yoy.columns if c != "Category"}
    styled = df_yoy.style.format(fmt).map(color_change, subset=["26→27", "27→28"])
    st.dataframe(styled, use_container_width=True, hide_index=True)


# ---------------------------------------------------------------------------
# Page 4: 2025 Actuals
# ---------------------------------------------------------------------------
elif page == "2025 Actuals":
    st.title("2025 Budget vs Actuals")

    if not actuals_2025:
        st.warning("Could not load 2025 actuals data. Make sure Dept_44_Budget.xlsx is included.")
    else:
        # Category-level summary
        cat_records = [r for r in actuals_2025 if r["is_category"]]

        st.subheader("Variance Analysis by Category")
        var_rows = []
        for r in cat_records:
            budget = r["yearly_budget"]
            actual = r["yearly_actual"]
            variance = budget - actual
            util = (actual / budget * 100) if budget else 0
            var_rows.append({
                "Account": r["account"],
                "Category": r["category"],
                "Budget": budget,
                "Actual": actual,
                "Variance": variance,
                "Utilization": util,
            })

        df_var = pd.DataFrame(var_rows)

        # Summary cards
        total_budget = df_var["Budget"].sum()
        total_actual = df_var["Actual"].sum()
        total_variance = total_budget - total_actual

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Total Budget", f"${total_budget:,.0f}")
        c2.metric("Total Actual", f"${total_actual:,.0f}")
        c3.metric("Total Variance", f"${total_variance:,.0f}",
                  delta=f"{'Under' if total_variance > 0 else 'Over'} budget")
        c4.metric("Utilization", f"{total_actual / total_budget * 100:.1f}%" if total_budget else "N/A")

        st.divider()

        # Budget vs Actual grouped bar
        col_l, col_r = st.columns(2)

        with col_l:
            st.subheader("Budget vs Actual by Category")
            df_bva = df_var[["Category", "Budget", "Actual"]].melt(
                id_vars="Category", var_name="Type", value_name="Amount"
            )
            df_bva["Category"] = df_bva["Category"].str.replace(
                r"^CC - |^GP Comp - ", "", regex=True
            ).str[:35]
            fig_bva = px.bar(
                df_bva, x="Category", y="Amount", color="Type",
                barmode="group",
                color_discrete_map={"Budget": "#5B8DEF", "Actual": "#F7B731"},
                text_auto="$,.0f",
            )
            fig_bva.update_layout(**PLOTLY_LAYOUT, height=450)
            fig_bva.update_traces(textposition="outside", textfont_size=9)
            fig_bva.update_xaxes(tickangle=45)
            st.plotly_chart(fig_bva, use_container_width=True)

        with col_r:
            st.subheader("Utilization Rate")
            df_util = df_var[["Category", "Utilization"]].copy()
            df_util["Category"] = df_util["Category"].str.replace(
                r"^CC - |^GP Comp - ", "", regex=True
            ).str[:35]
            df_util = df_util.sort_values("Utilization")
            df_util["Color"] = df_util["Utilization"].apply(
                lambda x: "#FC5C65" if x > 100 else "#2ED573" if x < 80 else "#F7B731"
            )
            fig_util = px.bar(
                df_util, x="Utilization", y="Category", orientation="h",
                color="Color",
                color_discrete_map="identity",
                text_auto=".1f",
            )
            fig_util.update_layout(**PLOTLY_LAYOUT, height=450, showlegend=False)
            fig_util.update_traces(texttemplate="%{x:.1f}%", textposition="outside")
            fig_util.add_vline(x=100, line_dash="dash", line_color="#FAFAFA", opacity=0.4)
            st.plotly_chart(fig_util, use_container_width=True)

        # Monthly budget vs actual trend
        st.subheader("Monthly Budget vs Actual Trend")
        monthly_b = [0.0] * 12
        monthly_a = [0.0] * 12
        for r in cat_records:
            for m in range(12):
                monthly_b[m] += r["monthly_budget"][m]
                monthly_a[m] += r["monthly_actual"][m]

        df_trend = pd.DataFrame({
            "Month": MONTH_LABELS,
            "Budget": monthly_b,
            "Actual": monthly_a,
        })
        fig_trend = go.Figure()
        fig_trend.add_trace(go.Scatter(
            x=df_trend["Month"], y=df_trend["Budget"],
            name="Budget", line=dict(color="#5B8DEF", width=3),
            mode="lines+markers",
        ))
        fig_trend.add_trace(go.Scatter(
            x=df_trend["Month"], y=df_trend["Actual"],
            name="Actual", line=dict(color="#F7B731", width=3),
            mode="lines+markers",
        ))
        fig_trend.update_layout(**PLOTLY_LAYOUT, height=350)
        st.plotly_chart(fig_trend, use_container_width=True)

        # Line item detail table
        st.subheader("Line Item Utilization")
        li_records = [r for r in actuals_2025 if not r["is_category"]]
        li_rows = []
        for r in li_records:
            budget = r["yearly_budget"]
            actual = r["yearly_actual"]
            variance = budget - actual
            util = (actual / budget * 100) if budget else (100 if actual > 0 else 0)
            li_rows.append({
                "Account": r["account"],
                "Category": r["category"],
                "Line Item": r["line_item"],
                "Budget": budget,
                "Actual": actual,
                "Variance": variance,
                "Util %": util,
            })

        df_li = pd.DataFrame(li_rows)

        def color_variance(val):
            if not isinstance(val, (int, float)):
                return ""
            if val < 0:
                return "color: #FC5C65"
            elif val > 0:
                return "color: #2ED573"
            return ""

        def color_util(val):
            if not isinstance(val, (int, float)):
                return ""
            if val > 100:
                return "color: #FC5C65"
            elif val < 50:
                return "color: #2ED573"
            return ""

        fmt = {"Budget": "${:,.2f}", "Actual": "${:,.2f}", "Variance": "${:,.2f}", "Util %": "{:.1f}%"}
        styled = (df_li.style
                  .format(fmt)
                  .map(color_variance, subset=["Variance"])
                  .map(color_util, subset=["Util %"]))
        st.dataframe(styled, use_container_width=True, hide_index=True, height=500)
