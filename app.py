import io
import re
from datetime import datetime
from zoneinfo import ZoneInfo

import pandas as pd
import plotly.graph_objects as go
import requests
import streamlit as st

# -----------------------------------------------------------------------------
# PAGE CONFIGURATION
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="Overdue Works Dashboard",
    page_icon="📊",
    layout="wide",
)

# -----------------------------------------------------------------------------
# CONSTANTS & TARGET SHEETS
# -----------------------------------------------------------------------------
TARGET_SHEETS = [
    "4702_MOT",
    "4702-New Tanks",
    "4702-LIS",
    "4702_A&P",
    "PMKSY-SMI-RRR",
    "4702_Paschima Vahini",
    "4711__Other Flood Control",
    "4702-Kharland",
    "NABARD - New Tank",
    "SDMF",
    "KMERC",
    "NABARD - A&P",
    "ATAL BHUJAL",
    "NABARD_MOT",
    "NABARD - LIS",
    "other dept funds",
    "SCP (LIS)",
    "SCP (Borewell)",
    "SCP (Checkdam)",
    "TSP (Checkdam)",
    "TSP (Borewell)",
    "TSP (LIS)",
    "NABARD - Other Flood Control",
]

COL_SO = "ಶಾಖಾಧಿಕಾರಿಯ ಹೆಸರು"
COL_AEE = "ಸಹಾಯಕ ಕಾರ್ಯಪಾಲಕ ಇಂಜಿನಿಯರ್ ಹೆಸರು"
COL_TARGET_DATE = "ಕಾಮಗಾರಿ ಮುಕ್ತಾಯಗೊಳಿಸಬೇಕಾದ ದಿನಾಂಕ"
COL_ACTUAL_DATE = "ಮುಕ್ತಾಯವಾದ ದಿನಾಂಕ"

DEFAULT_SHEET_URL = (
    "https://docs.google.com/spreadsheets/d/"
    "13W04z-fdU3RhuVJ5vQFRpK4BBSQRJWfscPqpr7Oj2Y8/edit?gid=1055199869#gid=1055199869"
)


# -----------------------------------------------------------------------------
# HELPER FUNCTIONS
# -----------------------------------------------------------------------------
def normalize_text(val) -> str:
    """Collapse multiple spaces/newlines and strip text."""
    if pd.isna(val):
        return ""
    return re.sub(r"\s+", " ", str(val)).strip()


def extract_export_url(sheet_url: str) -> str:
    """Convert a standard Google Sheets edit link into an xlsx export URL."""
    match = re.search(r"/spreadsheets/d/([a-zA-Z0-9-_]+)", sheet_url)
    if not match:
        raise ValueError("Invalid Google Sheet URL. Could not extract spreadsheet key.")
    spreadsheet_key = match.group(1)
    return f"https://docs.google.com/spreadsheets/d/{spreadsheet_key}/export?format=xlsx"


def find_matching_column_indices(headers: list[str]) -> dict:
    """
    Find required column indices (0-based positions) across sheets:
    - ಶಾಖಾಧಿಕಾರಿಯ ಹೆಸರು
    - ಸಹಾಯಕ ಕಾರ್ಯಪಾಲಕ ಇಂಜಿನಿಯರ್ ಹೆಸರು
    - ಕಾಮಗಾರಿ ಮುಕ್ತಾಯಗೊಳಿಸಬೇಕಾದ ದಿನಾಂಕ
    - ಮುಕ್ತಾಯವಾದ ದಿನಾಂಕ
    - Status column containing 'CP/CF' anywhere in the header
    """
    col_idx_map = {}
    for idx, col in enumerate(headers):
        norm_col = normalize_text(col)
        upper_col = norm_col.upper()
        compact_upper = re.sub(r"\s+", "", upper_col)

        if (COL_SO in norm_col or "ಶಾಖಾಧಿಕಾರಿ" in norm_col) and "so" not in col_idx_map:
            col_idx_map["so"] = idx
        elif (
            COL_AEE in norm_col or "ಸಹಾಯಕ ಕಾರ್ಯಪಾಲಕ ಇಂಜಿನಿಯರ್" in norm_col
        ) and "aee" not in col_idx_map:
            col_idx_map["aee"] = idx
        elif (
            COL_TARGET_DATE in norm_col or "ಮುಕ್ತಾಯಗೊಳಿಸಬೇಕಾದ ದಿನಾಂಕ" in norm_col
        ) and "target_date" not in col_idx_map:
            col_idx_map["target_date"] = idx
        elif (
            COL_ACTUAL_DATE in norm_col
            and "ಮುಕ್ತಾಯಗೊಳಿಸಬೇಕಾದ" not in norm_col
            and "actual_date" not in col_idx_map
        ):
            col_idx_map["actual_date"] = idx
        elif "CP/CF" in compact_upper and "status" not in col_idx_map:
            col_idx_map["status"] = idx

    return col_idx_map


def parse_indian_date(series: pd.Series) -> pd.Series:
    """Safely parse dates (handling datetime objects and DD-MM-YYYY / DD.MM.YYYY strings)."""
    def _clean_date(val):
        if pd.isna(val):
            return pd.NaT
        if isinstance(val, datetime):
            return pd.Timestamp(val.date())
        s = normalize_text(val)
        if not s or s.upper() in {"-", "--", "NIL", "NA", "N/A", "NONE", ""}:
            return pd.NaT
        s = re.sub(r"[./]", "-", s)
        return s

    cleaned = series.map(_clean_date)
    return pd.to_datetime(cleaned, dayfirst=True, errors="coerce")


def is_empty_date(series: pd.Series) -> pd.Series:
    """Check if the 'ಮುಕ್ತಾಯವಾದ ದಿನಾಂಕ' cell is empty or placeholder text."""
    def _check_empty(val):
        if pd.isna(val):
            return True
        s = normalize_text(val)
        return s.upper() in {"", "-", "--", "NIL", "NA", "N/A", "NONE", "NAT"}

    return series.map(_check_empty)


def is_valid_officer_name(name: str) -> bool:
    """Filter out empty cells, serial numbers, or summary rows like 'ಒಟ್ಟು' / 'Total'."""
    if not name:
        return False
    upper = name.upper()
    invalid_tokens = {"NAN", "NONE", "-", "--", "NIL", "NA", "N/A", "0"}
    if upper in invalid_tokens or name.isdigit():
        return False
    if "ಒಟ್ಟು" in name or "TOTAL" in upper or COL_SO in name or COL_AEE in name:
        return False
    return True


@st.cache_data(ttl=600, show_spinner=False)
def load_and_process_sheets(sheet_url: str):
    """Download the workbook and extract records from the 23 specified sheets."""
    export_url = extract_export_url(sheet_url)
    response = requests.get(export_url, timeout=60)
    response.raise_for_status()

    excel_file = pd.ExcelFile(io.BytesIO(response.content), engine="openpyxl")

    # Lookup map normalized by whitespace and case
    available_sheets = {
        normalize_text(name).lower(): name for name in excel_file.sheet_names
    }

    today_ist = pd.Timestamp(datetime.now(ZoneInfo("Asia/Kolkata")).date())

    compiled_rows = []
    sheet_diagnostics = []

    for target in TARGET_SHEETS:
        norm_target = normalize_text(target).lower()
        actual_sheet_name = available_sheets.get(norm_target)

        if not actual_sheet_name:
            sheet_diagnostics.append(
                {"Sheet": target, "Status": "Missing in Workbook", "Rows Loaded": 0}
            )
            continue

        try:
            raw_df = pd.read_excel(excel_file, sheet_name=actual_sheet_name, header=None)
            if raw_df.empty:
                sheet_diagnostics.append(
                    {"Sheet": target, "Status": "Empty Sheet", "Rows Loaded": 0}
                )
                continue

            # Dynamically locate the header row within the first 25 rows
            header_idx = None
            for idx in range(min(25, len(raw_df))):
                row_text = " | ".join(normalize_text(v) for v in raw_df.iloc[idx].values)
                if (
                    COL_SO in row_text
                    or COL_AEE in row_text
                    or "ಶಾಖಾಧಿಕಾರಿ" in row_text
                ):
                    header_idx = idx
                    break

            if header_idx is None:
                sheet_diagnostics.append(
                    {"Sheet": target, "Status": "Header Row Not Found", "Rows Loaded": 0}
                )
                continue

            headers = [normalize_text(col) for col in raw_df.iloc[header_idx].values]
            col_idx_map = find_matching_column_indices(headers)

            required_keys = {"so", "aee", "target_date", "actual_date", "status"}
            missing_keys = required_keys - set(col_idx_map.keys())

            if missing_keys:
                sheet_diagnostics.append(
                    {
                        "Sheet": target,
                        "Status": f"Missing Columns: {', '.join(sorted(missing_keys))}",
                        "Rows Loaded": 0,
                    }
                )
                continue

            # Reset index immediately after slicing below header row so indices start at 0
            df = raw_df.iloc[header_idx + 1 :].reset_index(drop=True).copy()

            s_so = df.iloc[:, col_idx_map["so"]].map(normalize_text)
            s_aee = df.iloc[:, col_idx_map["aee"]].map(normalize_text)
            s_status = df.iloc[:, col_idx_map["status"]].map(normalize_text)
            s_target_raw = df.iloc[:, col_idx_map["target_date"]]
            s_actual_raw = df.iloc[:, col_idx_map["actual_date"]]

            sub = pd.DataFrame(
                {
                    "sheet_name": target,
                    "so_name": s_so,
                    "aee_name": s_aee,
                    "status_raw": s_status,
                    "target_date_raw": s_target_raw.map(normalize_text),
                    "actual_date_raw": s_actual_raw.map(normalize_text),
                    "target_dt": parse_indian_date(s_target_raw),
                    "actual_empty": is_empty_date(s_actual_raw),
                }
            )

            # Keep only rows that have an officer name or status entry
            sub = sub[
                (sub["status_raw"] != "")
                | (sub["so_name"] != "")
                | (sub["aee_name"] != "")
            ].reset_index(drop=True)

            sub["is_status_p"] = sub["status_raw"].str.upper() == "P"

            # Overdue Logic:
            # 1. Status is 'P' or 'p' under CP/CF column
            # 2. ಕಾಮಗಾರಿ ಮುಕ್ತಾಯಗೊಳಿಸಬೇಕಾದ ದಿನಾಂಕ < Today's Date
            # 3. ಮುಕ್ತಾಯವಾದ ದಿನಾಂಕ is empty
            sub["is_overdue"] = (
                sub["is_status_p"]
                & sub["target_dt"].notna()
                & (sub["target_dt"] < today_ist)
                & sub["actual_empty"]
            )

            compiled_rows.append(sub)
            sheet_diagnostics.append(
                {
                    "Sheet": target,
                    "Status": "Loaded",
                    "Rows Loaded": len(sub),
                }
            )

        except Exception as sheet_err:
            sheet_diagnostics.append(
                {
                    "Sheet": target,
                    "Status": f"Error: {sheet_err}",
                    "Rows Loaded": 0,
                }
            )

    master_df = (
        pd.concat(compiled_rows, ignore_index=True)
        if compiled_rows
        else pd.DataFrame()
    )
    diag_df = pd.DataFrame(sheet_diagnostics)
    return master_df, diag_df, today_ist.strftime("%d-%m-%Y")


# -----------------------------------------------------------------------------
# SIDEBAR CONTROLS
# -----------------------------------------------------------------------------
st.sidebar.header("⚙️ Dashboard Controls")

sheet_url_input = st.sidebar.text_input(
    "Google Spreadsheet Link",
    value=DEFAULT_SHEET_URL,
    help="Ensure the Google Sheet is set to 'Anyone with the link can view'.",
)

denominator_mode = st.sidebar.radio(
    "Total Works Assigned Count Basis:",
    options=[
        "All Works Assigned (All Statuses)",
        "Only Ongoing Works (Status = 'P')",
    ],
    index=0,
    help=(
        "Choose whether 'Total Works Assigned' counts every work row assigned to the officer "
        "across the 23 sheets, or only works currently marked 'P' / 'p'."
    ),
)

if st.sidebar.button("🔄 Refresh Live Data"):
    st.cache_data.clear()

# -----------------------------------------------------------------------------
# MAIN UI
# -----------------------------------------------------------------------------
st.title("📊 Minor Irrigation — Overdue Works Dashboard")

if not sheet_url_input:
    st.info("Please paste your Google Sheet link in the sidebar to load the dashboard.")
    st.stop()

with st.spinner("Fetching and processing 23 worksheets from Google Sheets..."):
    try:
        master_df, diag_df, today_str = load_and_process_sheets(sheet_url_input)
    except requests.HTTPError as http_err:
        st.error(
            f"Google Sheets HTTP Error: Make sure link sharing is set to 'Anyone with the link can view'.\n\nDetails: {http_err}"
        )
        st.stop()
    except Exception as e:
        st.error(f"Error while processing workbook: {e}")
        st.stop()

if master_df.empty:
    st.warning("No matching data rows were found across the target worksheets.")
    st.dataframe(diag_df, use_container_width=True)
    st.stop()

# -----------------------------------------------------------------------------
# SWITCH BETWEEN SECTION OFFICER & AEE
# -----------------------------------------------------------------------------
view_choice = st.radio(
    "ವೀಕ್ಷಣೆ ಆಯ್ಕೆಮಾಡಿ / Select Officer View:",
    options=[COL_SO, COL_AEE],
    horizontal=True,
)

group_col = "so_name" if view_choice == COL_SO else "aee_name"

# Filter valid officer names for the selected view
valid_df = master_df[master_df[group_col].map(is_valid_officer_name)].copy()

if denominator_mode == "Only Ongoing Works (Status = 'P')":
    valid_df = valid_df[valid_df["is_status_p"]].copy()

# Aggregate per officer
summary_df = (
    valid_df.groupby(group_col, as_index=False)
    .agg(
        total_works=("is_overdue", "count"),
        ongoing_p_works=("is_status_p", "sum"),
        overdue_works=("is_overdue", "sum"),
    )
    .rename(columns={group_col: view_choice})
)

summary_df["overdue_pct"] = (
    (summary_df["overdue_works"] / summary_df["total_works"]) * 100
).round(1).fillna(0.0)

# Sort highest overdue at the top
summary_df = summary_df.sort_values(
    by=["overdue_works", "overdue_pct", "total_works"],
    ascending=[False, False, False],
).reset_index(drop=True)

# -----------------------------------------------------------------------------
# KPI SUMMARY METRICS
# -----------------------------------------------------------------------------
total_assigned_all = int(summary_df["total_works"].sum())
total_overdue_all = int(summary_df["overdue_works"].sum())
overall_overdue_pct = (
    round((total_overdue_all / total_assigned_all) * 100, 1)
    if total_assigned_all > 0
    else 0.0
)

k1, k2, k3, k4 = st.columns(4)
k1.metric("📅 Comparison Date (Today)", today_str)
k2.metric(f"👥 Total Officers ({view_choice})", len(summary_df))
k3.metric("📋 Total Works Assigned", f"{total_assigned_all:,}")
k4.metric(
    "⚠️ Total Overdue Works",
    f"{total_overdue_all:,} ({overall_overdue_pct}%)",
)

st.divider()

# -----------------------------------------------------------------------------
# HORIZONTAL BAR CHART (HIGHEST OVERDUE AT THE TOP)
# -----------------------------------------------------------------------------
st.subheader(f"📈 Overdue Works by {view_choice}")

# Reverse order for Plotly horizontal bar so highest overdue renders at the TOP
plot_df = summary_df.iloc[::-1].copy()

plot_df["bar_label"] = plot_df.apply(
    lambda r: (
        f"  <b>{int(r['overdue_works'])} Overdue</b> / "
        f"{int(r['total_works'])} Total ({r['overdue_pct']:.1f}%)"
    ),
    axis=1,
)

fig = go.Figure()

# Background / companion bar showing Total Works Assigned
fig.add_trace(
    go.Bar(
        y=plot_df[view_choice],
        x=plot_df["total_works"],
        name="Total Works Assigned",
        orientation="h",
        marker=dict(
            color="rgba(148, 163, 184, 0.35)",
            line=dict(color="#64748B", width=1),
        ),
        hovertemplate=(
            f"<b>{view_choice}:</b> %{{y}}<br>"
            "<b>Total Works Assigned:</b> %{x}<extra></extra>"
        ),
    )
)

# Foreground bar showing Overdue Works + rich text annotation
fig.add_trace(
    go.Bar(
        y=plot_df[view_choice],
        x=plot_df["overdue_works"],
        name="Overdue Works",
        orientation="h",
        marker=dict(color="#DC2626"),
        text=plot_df["bar_label"],
        textposition="outside",
        cliponaxis=False,
        customdata=plot_df[["total_works", "ongoing_p_works", "overdue_pct"]],
        hovertemplate=(
            f"<b>{view_choice}:</b> %{{y}}<br>"
            "<b>Overdue Works:</b> %{x}<br>"
            "<b>Total Works Assigned:</b> %{customdata[0]}<br>"
            "<b>Ongoing ('P') Works:</b> %{customdata[1]}<br>"
            "<b>Overdue Percentage:</b> %{customdata[2]:.1f}%<extra></extra>"
        ),
    )
)

max_x = max(plot_df["total_works"].max(), plot_df["overdue_works"].max(), 1)
chart_height = max(450, len(plot_df) * 38 + 120)

fig.update_layout(
    barmode="overlay",
    height=chart_height,
    xaxis=dict(
        title="Number of Works",
        range=[0, max_x * 1.38],
        gridcolor="rgba(200, 200, 200, 0.25)",
    ),
    yaxis=dict(
        title=view_choice,
        automargin=True,
    ),
    legend=dict(
        orientation="h",
        yanchor="bottom",
        y=1.02,
        xanchor="right",
        x=1,
    ),
    margin=dict(l=20, r=180, t=50, b=40),
)

st.plotly_chart(fig, use_container_width=True)

# -----------------------------------------------------------------------------
# DETAILED SUMMARY TABLE & SHEET DIAGNOSTICS
# -----------------------------------------------------------------------------
st.subheader("📋 Officer-Wise Summary Table")

display_table = summary_df.rename(
    columns={
        "total_works": "Total Works Assigned",
        "ongoing_p_works": "Ongoing Works ('P')",
        "overdue_works": "Overdue Works",
        "overdue_pct": "Overdue %",
    }
)
display_table.index = display_table.index + 1

st.dataframe(
    display_table.style.format({"Overdue %": "{:.1f}%"}),
    use_container_width=True,
)

with st.expander("🔍 Inspect Worksheet Loading Status (23 Target Sheets)"):
    st.dataframe(diag_df, use_container_width=True)
