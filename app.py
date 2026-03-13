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

    Detects years dynamically from header row.
    Returns dict with keys:
        account   – e.g. '6623.18'
        desc      – full GL description from Line000
        years     – list of years found (e.g. [2025] or [2026, 2027, 2028])
        line_items – list of dicts with name + per-year data
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
    header_row = []
    for i, r in enumerate(rows):
        if any("Line Item" in c for c in r):
            header_idx = i
            header_row = r
            break

    if header_idx is None or account is None:
        return None

    # Detect year columns from header
    # Format: ..., "Comments YYYY", "YYYY", "YYYY-01", ..., "YYYY-12", ...
    year_columns = {}  # year -> {'annual_col': int, 'monthly_start': int}
    for ci, cell in enumerate(header_row):
        cell = cell.strip()
        m = re.match(r"^(\d{4})$", cell)
        if m:
            year = int(m.group(1))
            # Check if next columns are monthly (YYYY-01, etc.)
            if ci + 1 < len(header_row) and re.match(rf"^{year}-\d{{2}}$", header_row[ci + 1].strip()):
                year_columns[year] = {'annual_col': ci, 'monthly_start': ci + 1}
            else:
                # Annual-only year (like 2027, 2028 in the original files)
                year_columns[year] = {'annual_col': ci, 'monthly_start': None}

    if not year_columns:
        return None

    # Primary year = first year with monthly data
    primary_year = None
    for y in sorted(year_columns.keys()):
        if year_columns[y]['monthly_start'] is not None:
            primary_year = y
            break
    if primary_year is None:
        primary_year = min(year_columns.keys())

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

        # Get primary year annual
        pcol = year_columns[primary_year]['annual_col']
        primary_annual = _to_float(r[pcol]) if len(r) > pcol else 0
        if not name and primary_annual == 0:
            continue
        if not name:
            continue

        item = {"name": name}
        for year, cols in sorted(year_columns.items()):
            acol = cols['annual_col']
            item[f"annual_{year}"] = _to_float(r[acol]) if len(r) > acol else 0.0
            if cols['monthly_start'] is not None:
                item[f"monthly_{year}"] = [
                    _to_float(r[cols['monthly_start'] + m]) if len(r) > cols['monthly_start'] + m else 0.0
                    for m in range(12)
                ]

        line_items.append(item)

    return {
        "account": account,
        "desc": desc,
        "years": sorted(year_columns.keys()),
        "line_items": line_items,
    }


def category_label(account: str) -> str:
    name = GL_NAMES.get(account, account)
    return f"{account} – {name}"


def li_annual(li: dict, year: int) -> float:
    """Get annual amount for a line item, safe for any year."""
    key = f"annual_{year}"
    return li.get(key, 0.0)


# ---------------------------------------------------------------------------
# Google Drive data source
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
    """Load all budget files from Google Drive folder, parse, and return gl_data.

    Multiple files for the same GL account (e.g. 2025 actuals + 2026 budget) are merged.
    Cached for 5 minutes. folder_id is passed as param so cache key is stable.
    """
    service = _build_drive_service()

    results = service.files().list(
        q=f"'{folder_id}' in parents and trashed = false",
        fields="files(id, name, mimeType)",
        pageSize=100,
    ).execute()
    files = results.get("files", [])

    gl_data = {}
    export_mime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

    for f in files:
        file_id = f["id"]
        mime = f["mimeType"]

        if mime == "application/vnd.google-apps.spreadsheet":
            resp = service.files().export(fileId=file_id, mimeType=export_mime).execute()
        elif mime == export_mime:
            resp = service.files().get_media(fileId=file_id).execute()
        else:
            continue

        file_bytes = resp if isinstance(resp, bytes) else resp.encode("latin-1")
        parsed = parse_gl_bytes(file_bytes)
        if parsed:
            acct = parsed["account"]
            if acct in gl_data:
                existing = gl_data[acct]
                existing["line_items"].extend(parsed["line_items"])
                existing["years"] = sorted(set(existing["years"]) | set(parsed["years"]))
            else:
                gl_data[acct] = parsed

    return gl_data


def _process_files(file_data: dict[str, bytes]) -> dict:
    """Parse all files and return gl_data dict keyed by account.
    
    Multiple files for the same account are merged (line items combined).
    """
    gl_data = {}

    for fname, fbytes in sorted(file_data.items()):
        parsed = parse_gl_bytes(fbytes)
        if parsed:
            acct = parsed["account"]
            if acct in gl_data:
                # Merge: combine line items and years
                existing = gl_data[acct]
                existing["line_items"].extend(parsed["line_items"])
                existing["years"] = sorted(set(existing["years"]) | set(parsed["years"]))
            else:
                gl_data[acct] = parsed

    return gl_data


# ---------------------------------------------------------------------------
# Sidebar: department info + navigation
# ---------------------------------------------------------------------------
with st.sidebar:
    st.markdown("### GVTC Department 44")
    st.markdown("**Web Management**")
    st.markdown("Entity: P0100")
    st.divider()

# ---------------------------------------------------------------------------
# Load data from Google Drive
# ---------------------------------------------------------------------------
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
    gl_data = load_drive_data(st.secrets["google_drive"]["folder_id"])
except Exception as e:
    st.error(f"**Failed to load data from Google Drive:** {e}")
    st.stop()

if not gl_data:
    st.warning(
        "No valid GL account data found in the Google Drive folder. "
        "Please check that the folder contains the correct GVTC budget files."
    )
    st.stop()

# Collect all years across all GL files
all_years = set()
for acct, info in gl_data.items():
    all_years.update(info.get("years", []))
all_years = sorted(all_years)

# Build summary DataFrames
summary_rows = []
for acct, info in sorted(gl_data.items()):
    row = {
        "account": acct,
        "category": category_label(acct),
        "short_name": GL_NAMES.get(acct, acct),
    }
    for year in all_years:
        row[str(year)] = sum(li_annual(li, year) for li in info["line_items"])
    summary_rows.append(row)

df_summary = pd.DataFrame(summary_rows)

# ---------------------------------------------------------------------------
# Sidebar: navigation & filters (continued)
# ---------------------------------------------------------------------------
with st.sidebar:
    if st.button("🔄 Refresh Data"):
        load_drive_data.clear()
        st.rerun()

    st.divider()

    page = st.radio(
        "Navigate",
        ["Budget Overview", "Monthly View", "Line Item Detail"],
        index=0,
    )

    st.divider()
    selected_year = st.selectbox("Year", all_years, index=min(1, len(all_years) - 1))

    # Build mapping: short name -> account number
    _name_to_acct = {GL_NAMES.get(a, a): a for a in sorted(gl_data.keys())}
    all_short_names = list(_name_to_acct.keys())

    if st.button("Select All Categories"):
        st.session_state["cat_filter"] = all_short_names

    selected_names = st.multiselect(
        "Filter Categories",
        all_short_names,
        default=st.session_state.get("cat_filter", all_short_names),
        key="cat_filter",
    )

    # Prevent empty selection from crashing charts
    if not selected_names:
        selected_names = all_short_names
        st.session_state["cat_filter"] = all_short_names
        st.rerun()

    st.divider()
    total_budget = df_summary[str(selected_year)].sum()
    st.metric("Total Budget", f"${total_budget:,.0f}", delta=None)

# Filter helper
filtered_accounts = [_name_to_acct[n] for n in selected_names if n in _name_to_acct]
df_filtered = df_summary[df_summary["account"].isin(filtered_accounts)]


# ---------------------------------------------------------------------------
# Page 1: Budget Overview
# ---------------------------------------------------------------------------
if page == "Budget Overview":
    st.title("Budget Overview")

    if df_filtered.empty:
        st.info("No data for the selected categories.")
        st.stop()

    # Summary cards
    year_strs = [str(y) for y in all_years]
    cols = st.columns(len(year_strs))
    for col, year_s in zip(cols, year_strs):
        val = df_filtered[year_s].sum()
        prev_s = str(int(year_s) - 1)
        prev_val = df_filtered[prev_s].sum() if prev_s in df_filtered.columns else None
        delta = f"${val - prev_val:+,.0f}" if prev_val is not None else None
        col.metric(f"{year_s} Budget", f"${val:,.0f}", delta=delta)

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

    # Category drill-down
    with col_right:
        st.subheader("Category Drill-Down")
        # Build list of categories with budget > 0
        drill_options = df_filtered[df_filtered[str(selected_year)] > 0][["account", "short_name", str(selected_year)]].copy()
        drill_options.columns = ["account", "name", "budget"]
        drill_options = drill_options.sort_values("budget", ascending=False)

        drill_choices = ["All Categories"] + drill_options["name"].tolist()
        drill_choice = st.selectbox(
            "Select a category to see line items",
            drill_choices,
            index=0,
        )

        if drill_choice == "All Categories":
            # Show overall distribution donut
            df_pie = drill_options[["name", "budget"]].copy()
            df_pie.columns = ["Category", "Budget"]
            total = df_pie["Budget"].sum()
            st.metric("Total Budget", f"${total:,.0f}")
            fig_all = px.pie(
                df_pie, values="Budget", names="Category",
                color_discrete_sequence=COLORS,
                hole=0.35,
            )
            fig_all.update_layout(**PLOTLY_LAYOUT, height=350,
                                  legend=dict(font=dict(size=11)))
            fig_all.update_traces(
                textinfo="none",
                textposition="inside",
                textfont_size=11,
                hovertemplate="%{label}<br>$%{value:,.0f}<br>%{percent}<extra></extra>",
            )
            st.plotly_chart(fig_all, use_container_width=True)
        else:
            # Drill into specific category
            drill_acct = drill_options[drill_options["name"] == drill_choice]["account"].iloc[0]

            if drill_acct in gl_data:
                info = gl_data[drill_acct]
                li_data = []
                for li in info["line_items"]:
                    annual = li_annual(li, selected_year)
                    if annual > 0:
                        li_data.append({"Line Item": li["name"], "Budget": annual})

                if li_data:
                    df_drill = pd.DataFrame(li_data).sort_values("Budget", ascending=False)
                    total = df_drill["Budget"].sum()

                    st.metric(f"{drill_choice} Total", f"${total:,.0f}")

                    fig_drill = px.pie(
                        df_drill, values="Budget", names="Line Item",
                        color_discrete_sequence=COLORS,
                        hole=0.35,
                    )
                    fig_drill.update_layout(**PLOTLY_LAYOUT, height=350,
                                            legend=dict(font=dict(size=11)))
                    fig_drill.update_traces(
                        textinfo="none",
                        texttemplate="%{label}<br>$%{value:,.0f}",
                        textposition="inside" if len(li_data) > 4 else "outside",
                        textfont_size=11,
                    )
                    st.plotly_chart(fig_drill, use_container_width=True)
                else:
                    st.info("No line items with budget for this year.")

    # Year-over-year comparison
    st.subheader("Year-over-Year Comparison")
    yoy_cols = ["short_name"] + [str(y) for y in all_years]
    df_yoy = df_filtered[yoy_cols].melt(
        id_vars="short_name", var_name="Year", value_name="Budget"
    )
    yoy_colors = ["#2ED573", "#5B8DEF", "#F7B731", "#FC5C65"][:len(all_years)]
    fig_yoy = px.bar(
        df_yoy, x="short_name", y="Budget", color="Year",
        barmode="group",
        color_discrete_sequence=yoy_colors,
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
        monthly_key = f"monthly_{selected_year}"
        for li in info["line_items"]:
            if monthly_key in li:
                for m in range(12):
                    monthly_totals[m] += li[monthly_key][m]
            else:
                # Year only has annual total, spread evenly
                annual = li_annual(li, selected_year)
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

    if df_monthly.empty:
        st.info("No data for the selected year and categories.")
        st.stop()

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
        total = sum(li_annual(li, selected_year) for li in info["line_items"])

        with st.expander(f"{label}  —  ${total:,.0f}", expanded=False):
            rows = []
            monthly_key = f"monthly_{selected_year}"

            for li in info["line_items"]:
                annual = li_annual(li, selected_year)
                if annual == 0 and not li["name"]:
                    continue
                row = {"Line Item": li["name"], "Annual": annual}
                if monthly_key in li:
                    for m in range(12):
                        row[MONTH_LABELS[m]] = li[monthly_key][m]
                else:
                    for m in range(12):
                        row[MONTH_LABELS[m]] = annual / 12

                # YoY change vs previous year if available
                prev_year = selected_year - 1
                if str(prev_year) in df_summary.columns:
                    prev = li_annual(li, prev_year)
                    row["YoY Change"] = annual - prev

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
        info = gl_data[acct]
        row_data = {"Category": GL_NAMES.get(acct, acct)}
        year_totals = {}
        for y in all_years:
            t = sum(li_annual(li, y) for li in info["line_items"])
            row_data[str(y)] = t
            year_totals[y] = t
        for i in range(1, len(all_years)):
            prev_y = all_years[i - 1]
            curr_y = all_years[i]
            short_prev = str(prev_y)[-2:]
            short_curr = str(curr_y)[-2:]
            row_data[f"{short_prev}→{short_curr}"] = year_totals[curr_y] - year_totals[prev_y]
        yoy_rows.append(row_data)
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
    change_cols = [c for c in df_yoy.columns if "→" in c]
    styled = df_yoy.style.format(fmt).map(color_change, subset=change_cols)
    st.dataframe(styled, use_container_width=True, hide_index=True)


