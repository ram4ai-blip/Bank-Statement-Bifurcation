"""
Vendor Transaction Tracker
──────────────────────────
Drop a monthly bank XLS → matches debits to vendors → pushes to Google Sheets.

Project layout:
    app.py                ← this file
    List_of_Vendors.csv   ← your keyword map (edit to add/update vendors)
    credentials.json      ← Google Service Account key (place once, never commit)
    requirements.txt
"""

import io
import math
import os
import re
import openpyxl

import gspread
import pandas as pd
import streamlit as st
from google.oauth2.service_account import Credentials
from rapidfuzz import fuzz

# ── Configuration ─────────────────────────────────────────────────────────────

VENDOR_CSV = "List_of_Vendors.csv"
CREDENTIALS = "credentials.json"
FUZZY_THRESHOLD = 88
HEADER_ROW = 6
DR_FLAG = "DR"

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

OUTPUT_COLS = [
    "Vendor Name",
    "Transaction Date",
    "Transaction ID",
    "Amount (INR)",
    "Original Description",
]

PIVOT_COLS = [
    "Vendor Name",
    "Month",
    "Total Amount (INR)",
    "No. of Transactions",
]

# ── Salary / Stipend exclusion ────────────────────────────────────────────────

EXCLUDE_SUBSTRINGS = [
    "salary", "stipend", "incentive", "bonus", "fnf", "fnfsettl",
    "full time", "fte salary", "salarystipend", "buld",
    "onetime", "one time", "partsalary", "reimburs" "Interns" "Ops" "ops" "interns" "intern",
]

EXCLUDE_TOKENS = {"ceo", "intern"}

# ── Tax / Government exclusion ────────────────────────────────────────────────
# Add substrings or whole tokens here to flag more tax-related descriptions.

TAX_SUBSTRINGS = ["dtax"]
TAX_TOKENS: set[str] = set()


def is_tax(description: str) -> bool:
    """Return True if the transaction looks like a tax / government payment."""
    desc_lower = description.lower()
    for kw in TAX_SUBSTRINGS:
        if kw in desc_lower:
            return True
    tokens = set(re.split(r"[/\s_\-]+", desc_lower))
    return bool(tokens & TAX_TOKENS)


def is_payroll(description: str) -> bool:
    desc_lower = description.lower()
    for kw in EXCLUDE_SUBSTRINGS:
        if kw in desc_lower:
            return True
    tokens = set(re.split(r"[/\s_\-]+", desc_lower))
    return bool(tokens & EXCLUDE_TOKENS)


# ── Data helpers ──────────────────────────────────────────────────────────────

def get_engine(file_bytes: bytes) -> str:
    """Detect XLS vs XLSX from magic bytes and return the correct pandas engine."""
    # XLSX is a ZIP file → starts with PK magic bytes
    # XLS is OLE2 compound doc → starts with 0xD0 0xCF
    return "openpyxl" if file_bytes[:2] == b'PK' else "xlrd"


@st.cache_data
def load_vendors() -> pd.DataFrame:
    df = pd.read_csv(VENDOR_CSV)
    df["Key Word"] = df["Key Word"].astype(str).str.strip()
    df["Vendor Name"] = df["Vendor Name"].astype(str).str.strip()
    return df


def build_display_map(vendors_df: pd.DataFrame) -> dict:
    kw_to_names: dict = {}
    for _, row in vendors_df.iterrows():
        kw_to_names.setdefault(row["Key Word"], []).append(row["Vendor Name"])
    return {
        kw: (kw if len(names) > 1 else names[0])
        for kw, names in kw_to_names.items()
    }


def get_sheet_names(file_bytes: bytes) -> list[str]:
    xl = pd.ExcelFile(io.BytesIO(file_bytes), engine=get_engine(file_bytes))
    return xl.sheet_names


def score_keyword(kw: str, description: str) -> int:
    kw_lower = kw.lower()
    desc_lower = description.lower()
    kw_nospace = re.sub(r"\s+", "", kw_lower)

    if len(kw) <= 4:
        return 100 if kw_lower in desc_lower else 0
    if kw_lower in desc_lower:
        return 100
    if len(kw_nospace) > 4 and kw_nospace in desc_lower:
        return 95
    return fuzz.partial_ratio(kw_lower, desc_lower)


def process_sheet(
    file_bytes:  bytes,
    sheet_name:  str,
    vendors_df:  pd.DataFrame,
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    raw = pd.read_excel(
        io.BytesIO(file_bytes),
        sheet_name=sheet_name,
        engine=get_engine(file_bytes),
        header=HEADER_ROW,
    )
    raw.columns = raw.columns.str.strip()

    raw = raw[[
        "Transaction ID",
        "Value Date",
        "Description",
        "Cr/Dr",
        "Transaction Amount(INR)",
    ]].copy()
    raw.rename(
        columns={"Transaction Amount(INR)": "Amount (INR)"}, inplace=True)
    raw.dropna(subset=["Transaction ID", "Description"], inplace=True)
    raw = raw[raw["Cr/Dr"].astype(str).str.strip() == DR_FLAG].copy()

    if raw.empty:
        empty = pd.DataFrame(columns=OUTPUT_COLS)
        return empty, empty, empty, empty

    raw["Value Date"] = (
        pd.to_datetime(raw["Value Date"], dayfirst=True, errors="coerce")
        .dt.strftime("%d/%m/%Y")
    )

    display_map = build_display_map(vendors_df)
    keywords = vendors_df["Key Word"].tolist()
    matched, unclassified, excluded, tax_rows = [], [], [], []

    for _, row in raw.iterrows():
        description = str(row["Description"])

        # ── Tax check (runs before payroll so DTAX never bleeds into Salary tab)
        if is_tax(description):
            tax_rows.append({
                "Vendor Name":          "Tax & Government",
                "Transaction Date":     str(row["Value Date"]),
                "Transaction ID":       str(row["Transaction ID"]),
                "Amount (INR)":         row["Amount (INR)"],
                "Original Description": description,
            })
            continue

        if is_payroll(description):
            excluded.append({
                "Vendor Name":          "Salary / Stipend / Internal",
                "Transaction Date":     str(row["Value Date"]),
                "Transaction ID":       str(row["Transaction ID"]),
                "Amount (INR)":         row["Amount (INR)"],
                "Original Description": description,
            })
            continue

        best_score, best_kw = 0, None
        for kw in keywords:
            score = score_keyword(kw, description)
            if score > best_score:
                best_score, best_kw = score, kw

        record = {
            "Vendor Name":          (
                display_map.get(best_kw, best_kw)
                if best_score >= FUZZY_THRESHOLD else "Unclassified"
            ),
            "Transaction Date":     str(row["Value Date"]),
            "Transaction ID":       str(row["Transaction ID"]),
            "Amount (INR)":         row["Amount (INR)"],
            "Original Description": description,
        }
        (matched if best_score >= FUZZY_THRESHOLD else unclassified).append(record)

    return (
        pd.DataFrame(matched,       columns=OUTPUT_COLS),
        pd.DataFrame(unclassified,  columns=OUTPUT_COLS),
        pd.DataFrame(excluded,      columns=OUTPUT_COLS),
        pd.DataFrame(tax_rows,      columns=OUTPUT_COLS),
    )


def build_pivot(matched_df: pd.DataFrame, month_name: str) -> pd.DataFrame:
    """Pivot with a Month column so all months consolidate in one tab."""
    if matched_df.empty:
        return pd.DataFrame(columns=PIVOT_COLS)
    pivot = (
        matched_df
        .groupby("Vendor Name", as_index=False)
        .agg(**{
            "Total Amount (INR)":  ("Amount (INR)", "sum"),
            "No. of Transactions": ("Transaction ID", "count"),
        })
        .sort_values("Total Amount (INR)", ascending=False)
    )
    pivot.insert(1, "Month", month_name)
    return pivot


# ── Google Sheets helpers ─────────────────────────────────────────────────────

def get_gspread_client():
    creds = Credentials.from_service_account_file(CREDENTIALS, scopes=SCOPES)
    return gspread.authorize(creds)


def _safe_value(val):
    if val is None:
        return ""
    if isinstance(val, float):
        return "" if math.isnan(val) else round(val, 2)
    if isinstance(val, bool):
        return int(val)
    return val


def to_gspread_values(df: pd.DataFrame) -> list[list]:
    """DataFrame → list-of-lists WITH header row (for initial writes)."""
    headers = df.columns.tolist()
    if df.empty:
        return [headers]
    return [headers] + [
        [_safe_value(v) for v in record.values()]
        for record in df.to_dict("records")
    ]


def to_gspread_rows(df: pd.DataFrame) -> list[list]:
    """DataFrame → list-of-lists WITHOUT header (for appending)."""
    if df.empty:
        return []
    return [
        [_safe_value(v) for v in record.values()]
        for record in df.to_dict("records")
    ]


def safe_tab_name(ss, desired: str) -> str:
    """Return desired name if available; otherwise desired(1), desired(2) …"""
    existing = {ws.title for ws in ss.worksheets()}
    if desired not in existing:
        return desired
    counter = 1
    while f"{desired}({counter})" in existing:
        counter += 1
    return f"{desired}({counter})"


def append_to_tab(
    ss,
    title:       str,
    df:          pd.DataFrame,
    col_headers: list,
) -> None:
    """
    Append df rows to an existing tab (no header re-write).
    Creates the tab with headers if it doesn't exist yet.
    """
    try:
        ws = ss.worksheet(title)
        rows = to_gspread_rows(df)
        if rows:
            ws.append_rows(rows, value_input_option="USER_ENTERED")
    except gspread.exceptions.WorksheetNotFound:
        ws = ss.add_worksheet(
            title=title,
            rows=max(len(df) + 50, 100),
            cols=len(col_headers) + 1,
        )
        write_df = df if not df.empty else pd.DataFrame(columns=col_headers)
        ws.update(to_gspread_values(write_df))
        col_letter = chr(64 + len(col_headers))
        ws.format(f"A1:{col_letter}1", {"textFormat": {"bold": True}})


def push_to_sheets(
    client,
    sheet_url:       str,
    tab_name:        str,
    matched_df:      pd.DataFrame,
    unclassified_df: pd.DataFrame,
    pivot_df:        pd.DataFrame,
    excluded_df:     pd.DataFrame,
    tax_df:          pd.DataFrame,
) -> tuple[str, str]:
    ss = client.open_by_url(sheet_url)

    actual_name = safe_tab_name(ss, tab_name)
    ws_main = ss.add_worksheet(
        title=actual_name,
        rows=max(len(matched_df) + 20, 50),
        cols=6,
    )
    ws_main.update(to_gspread_values(matched_df))
    ws_main.format("A1:E1", {"textFormat": {"bold": True}})

    try:
        blank = ss.worksheet("Sheet1")
        if not blank.get_all_values():
            ss.del_worksheet(blank)
    except Exception:
        pass

    append_to_tab(ss, "Unclassified",      unclassified_df, OUTPUT_COLS)
    append_to_tab(ss, "Pivot",             pivot_df,        PIVOT_COLS)
    append_to_tab(ss, "Salary & Internal", excluded_df,     OUTPUT_COLS)
    append_to_tab(ss, "Tax & Government",  tax_df,          OUTPUT_COLS)

    return ss.url, actual_name


# ── Streamlit UI ──────────────────────────────────────────────────────────────

def _init_state():
    defaults = {
        "processed":   False,
        "push_done":   False,
        "pushing":     False,
        "results":     None,
        "pushed_url":  None,
        "pushed_tab":  None,
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v


def main():
    st.set_page_config(
        page_title="Vendor Transaction Tracker",
        page_icon="💳",
        layout="wide",
    )
    _init_state()

    st.title("💳 Vendor Transaction Tracker")
    st.caption(
        "Upload a monthly bank XLS/XLSX → filter debits → exclude payroll & tax → "
        "match vendors → preview → push to Google Sheet."
    )

    # ── Guard: required files ─────────────────────────────────────────────────
    missing = []
    if not os.path.exists(VENDOR_CSV):
        missing.append(
            f"**`{VENDOR_CSV}`** — place your vendor keyword CSV next to `app.py`")
    if not os.path.exists(CREDENTIALS):
        missing.append(
            f"**`{CREDENTIALS}`** — place your Google Service Account key next to `app.py`")
    if missing:
        st.error("Missing required files:")
        for m in missing:
            st.markdown(f"- {m}")
        st.stop()

    vendors_df = load_vendors()

    # ── Sidebar ───────────────────────────────────────────────────────────────
    with st.sidebar:
        st.header("⚙️ Settings")
        try:
            import json
            with open(CREDENTIALS) as f:
                sa_email = json.load(f).get("client_email", "")
            st.info(
                f"**Service account email:**\n\n`{sa_email}`\n\n"
                "Share your Google Sheet with this address (Editor access)."
            )
        except Exception:
            st.warning(
                "Could not read service account email from credentials.json")

        st.markdown("---")
        st.caption(f"**{len(vendors_df)}** vendors loaded from `{VENDOR_CSV}`")
        st.caption(f"Fuzzy threshold: **{FUZZY_THRESHOLD} / 100**")
        st.caption(
            f"Payroll exclusion keywords: **{len(EXCLUDE_SUBSTRINGS) + len(EXCLUDE_TOKENS)}**")
        st.caption(
            f"Tax exclusion keywords: **{len(TAX_SUBSTRINGS) + len(TAX_TOKENS)}**")
        if st.button("🔄 Reload vendor list"):
            st.cache_data.clear()
            st.rerun()

    # ── Step 1: Upload ────────────────────────────────────────────────────────
    st.markdown("---")
    st.markdown("### Step 1 — Upload transaction file")
    uploaded = st.file_uploader(
        "Drag & drop your bank statement here (XLS or XLSX format)", type=["xls", "xlsx"]
    )
    if not uploaded:
        st.info("Waiting for a file…")
        st.stop()

    file_bytes = uploaded.read()
    try:
        sheet_names = get_sheet_names(file_bytes)
    except Exception as exc:
        st.error(f"Could not read file: {exc}")
        st.stop()

    # ── Step 2: Sheet selection ───────────────────────────────────────────────
    st.markdown("### Step 2 — Select sheet to process")
    selected_sheet = st.selectbox("Sheet found in this file:", sheet_names)

    # ── Step 3: Google Sheet target ───────────────────────────────────────────
    st.markdown("### Step 3 — Target Google Sheet")
    st.caption(
        "Create a blank Google Sheet, share it with the service account email "
        "(Editor access), then paste the URL below."
    )
    col_a, col_b = st.columns(2)
    with col_a:
        sheet_url = st.text_input(
            "Google Sheet URL",
            placeholder="https://docs.google.com/spreadsheets/d/...",
        )
    with col_b:
        tab_name = st.text_input(
            "Month tab name",
            placeholder="Jun 2025",
        )

    if not sheet_url.strip() or not tab_name.strip():
        st.warning("Enter the Google Sheet URL and a month tab name to continue.")
        st.stop()

    # ── Button 1: Process ─────────────────────────────────────────────────────
    st.markdown("---")
    if st.button("▶ Process Transactions", type="primary"):
        with st.spinner("Reading and classifying transactions…"):
            try:
                matched_df, unclassified_df, excluded_df, tax_df = process_sheet(
                    file_bytes, selected_sheet, vendors_df
                )
                pivot_df = build_pivot(matched_df, tab_name.strip())
                st.session_state.results = {
                    "matched_df":      matched_df,
                    "unclassified_df": unclassified_df,
                    "excluded_df":     excluded_df,
                    "tax_df":          tax_df,
                    "pivot_df":        pivot_df,
                }
                st.session_state.processed = True
                st.session_state.push_done = False
                st.session_state.pushed_url = None
                st.session_state.pushed_tab = None
            except Exception as exc:
                st.error(f"Processing failed: {exc}")
                st.session_state.processed = False

    # ── Preview (shown after successful process) ──────────────────────────────
    if st.session_state.processed and st.session_state.results:
        r = st.session_state.results
        matched_df = r["matched_df"]
        unclassified_df = r["unclassified_df"]
        excluded_df = r["excluded_df"]
        tax_df = r["tax_df"]
        pivot_df = r["pivot_df"]

        # Metrics
        st.markdown("---")
        c1, c2, c3, c4, c5, c6 = st.columns(6)
        total_debit = (
            matched_df["Amount (INR)"].sum()
            + unclassified_df["Amount (INR)"].sum()
            + excluded_df["Amount (INR)"].sum()
            + tax_df["Amount (INR)"].sum()
        )
        c1.metric("Total DR transactions", len(matched_df) +
                  len(unclassified_df) + len(excluded_df) + len(tax_df))
        c2.metric("Matched to vendors",    len(matched_df))
        c3.metric("Unclassified",          len(unclassified_df))
        c4.metric("Salary / Stipend",      len(excluded_df))
        c5.metric("Tax & Government",      len(tax_df))
        c6.metric("Total debit (INR)",     f"₹{total_debit:,.0f}")

        with st.expander("📋 Matched transactions", expanded=True):
            st.dataframe(matched_df, use_container_width=True, height=300)

        with st.expander("❓ Unclassified transactions"):
            if unclassified_df.empty:
                st.success("No unclassified transactions!")
            else:
                st.info(
                    "These didn't match any keyword. "
                    "Add keywords to `List_of_Vendors.csv` to classify them."
                )
                st.dataframe(unclassified_df, use_container_width=True)

        with st.expander("💼 Salary / Stipend / Internal"):
            if excluded_df.empty:
                st.success("No payroll transactions detected.")
            else:
                st.info(
                    f"{len(excluded_df)} payroll/internal transactions excluded from vendor matching.")
                st.dataframe(excluded_df, use_container_width=True)

        with st.expander("🏛️ Tax & Government"):
            if tax_df.empty:
                st.success("No tax transactions detected.")
            else:
                st.info(
                    f"{len(tax_df)} tax/government transactions excluded from vendor matching.")
                st.dataframe(tax_df, use_container_width=True)

        with st.expander("📊 Pivot summary"):
            st.dataframe(pivot_df, use_container_width=True)

        # ── Button 2: Push (only after process, hidden after push) ────────────
        if not st.session_state.push_done:
            st.markdown("---")
            st.markdown("**✅ Preview looks good? Push to Google Sheets.**")

            push_clicked = st.button(
                "📤 Push to Google Sheets",
                type="primary",
                disabled=st.session_state.pushing,
            )

            if push_clicked:
                st.session_state.pushing = True
                with st.spinner("Writing data to Google Sheet…"):
                    try:
                        client = get_gspread_client()
                        url, actual_tab = push_to_sheets(
                            client,
                            sheet_url.strip(),
                            tab_name.strip(),
                            matched_df,
                            unclassified_df,
                            pivot_df,
                            excluded_df,
                            tax_df,
                        )
                        st.session_state.push_done = True
                        st.session_state.pushing = False
                        st.session_state.pushed_url = url
                        st.session_state.pushed_tab = actual_tab
                        st.rerun()
                    except Exception as exc:
                        st.error(f"Google Sheets error: {exc}")
                        st.session_state.pushing = False

        # ── Success state ─────────────────────────────────────────────────────
        if st.session_state.push_done:
            st.markdown("---")
            st.success("✅ All done! Data pushed successfully.")

            actual_tab = st.session_state.pushed_tab
            if actual_tab and actual_tab != tab_name.strip():
                st.info(
                    f"A tab named **{tab_name.strip()}** already existed — "
                    f"data saved as **{actual_tab}** instead."
                )

            if st.session_state.pushed_url:
                st.markdown(
                    f"## [Open Google Sheet →]({st.session_state.pushed_url})")

            st.caption(
                "Run **▶ Process Transactions** again to process another month.")


if __name__ == "__main__":
    main()
