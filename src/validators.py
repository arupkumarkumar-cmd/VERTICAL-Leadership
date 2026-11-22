"""
validators.py
=============
Input validation, date range validation, source column validation,
and data quality checking for DS Weekly Operations Dashboard.
"""

from datetime import datetime, date
import re
import pandas as pd


REQUIRED_COLUMNS = [
    "TRACKER ID",
    "TRACKER HIST ID",
    "SHORT_DESCRIPTION",
    "ASSIGNED_TO_GROUP",
    "ASSIGNED TO USER",
    "CREATE_DT",
    "FOLLOW_UP_DT",
    "UPDATE_DT",
    "TRACKER STATUS",
    "TRACKER STATUS HISTORY",
]


def parse_date(date_str, allow_future=False, ref_date=None):
    """
    Parse a date string in DD-MM-YYYY or common variations.
    Returns datetime.date or raises ValueError.
    """
    if not date_str or not str(date_str).strip():
        raise ValueError("Date string cannot be empty.")

    clean_str = str(date_str).strip()

    formats = [
        "%d-%m-%Y",
        "%d/%m/%Y",
        "%Y-%m-%d",
        "%d-%b-%Y",
        "%d-%B-%Y",
        "%d.%m.%Y",
    ]

    parsed_date = None
    for fmt in formats:
        try:
            parsed_date = datetime.strptime(clean_str, fmt).date()
            break
        except ValueError:
            continue

    if parsed_date is None:
        raise ValueError(
            f"Invalid date format: '{clean_str}'. Expected DD-MM-YYYY (e.g., 09-11-2026)."
        )

    if not allow_future:
        check_ref = ref_date if ref_date is not None else date.today()
        if parsed_date > check_ref:
            raise ValueError(
                f"Date {parsed_date.strftime('%d-%b-%Y')} is in the future. "
                f"Future reporting dates are not permitted."
            )

    return parsed_date


def validate_date_range(start_date, end_date):
    """
    Validate that start_date and end_date form a valid period.
    """
    if start_date is None or end_date is None:
        return False, "Start date and End date must be provided."

    if start_date > end_date:
        return False, (
            f"Start date ({start_date.strftime('%d-%b-%Y')}) cannot be after "
            f"End date ({end_date.strftime('%d-%b-%Y')})."
        )

    return True, ""


def validate_source_columns(df):
    """
    Validate presence of all required columns in the dataframe (case-insensitive).
    Returns (is_valid, missing_columns, column_mapping).
    """
    df_cols_norm = {re.sub(r'[\s_]+', ' ', str(c).strip().upper()): c for c in df.columns}

    col_mapping = {}
    missing_columns = []

    for req in REQUIRED_COLUMNS:
        req_norm = re.sub(r'[\s_]+', ' ', req.strip().upper())
        if req_norm in df_cols_norm:
            col_mapping[req] = df_cols_norm[req_norm]
        else:
            # Fallback for minor variations
            found = False
            for actual_norm, actual_orig in df_cols_norm.items():
                if req_norm.replace(" ", "") == actual_norm.replace(" ", ""):
                    col_mapping[req] = actual_orig
                    found = True
                    break
            if not found:
                missing_columns.append(req)

    is_valid = len(missing_columns) == 0
    return is_valid, missing_columns, col_mapping


def run_data_quality_checks(df_raw, col_mapping, start_date=None, end_date=None, reconciliation_diff=0):
    """
    Perform thorough data quality validation.
    Returns:
    - summary: dict of counts
    - issues_df: DataFrame listing specific records with issues
    """
    issues = []

    id_col = col_mapping.get("TRACKER ID")
    hist_col = col_mapping.get("TRACKER HIST ID")
    grp_col = col_mapping.get("ASSIGNED_TO_GROUP")
    upd_col = col_mapping.get("UPDATE_DT")
    follow_col = col_mapping.get("FOLLOW_UP_DT")
    desc_col = col_mapping.get("SHORT_DESCRIPTION")

    # 1. Missing required identifiers
    if id_col:
        null_ids = df_raw[df_raw[id_col].isna()]
        for idx, row in null_ids.iterrows():
            issues.append({
                "TRACKER_ID": None,
                "TRACKER_HIST_ID": row.get(hist_col),
                "ISSUE_TYPE": "Missing TRACKER ID",
                "SEVERITY": "Critical",
                "DETAILS": f"Row index {idx} has empty TRACKER ID."
            })

    # 2. Duplicate or Invalid History IDs
    if hist_col and id_col:
        # Check non-numeric history IDs
        non_num_hist = df_raw[pd.to_numeric(df_raw[hist_col], errors='coerce').isna()]
        for idx, row in non_num_hist.iterrows():
            issues.append({
                "TRACKER_ID": row.get(id_col),
                "TRACKER_HIST_ID": row.get(hist_col),
                "ISSUE_TYPE": "Invalid History ID",
                "SEVERITY": "High",
                "DETAILS": f"Non-numeric TRACKER HIST ID: '{row.get(hist_col)}'"
            })

        # Check duplicate history IDs within same Tracker ID
        dupes = df_raw[df_raw.duplicated(subset=[id_col, hist_col], keep=False)]
        for idx, row in dupes.iterrows():
            issues.append({
                "TRACKER_ID": row.get(id_col),
                "TRACKER_HIST_ID": row.get(hist_col),
                "ISSUE_TYPE": "Duplicate History ID",
                "SEVERITY": "Medium",
                "DETAILS": f"Duplicate TRACKER HIST ID {row.get(hist_col)} for TRACKER ID {row.get(id_col)}"
            })

    # 3. Missing or invalid UPDATE_DT
    if upd_col:
        parsed_upd = pd.to_datetime(df_raw[upd_col], format='%d-%b-%y', errors='coerce')
        invalid_upd = df_raw[parsed_upd.isna()]
        for idx, row in invalid_upd.iterrows():
            issues.append({
                "TRACKER_ID": row.get(id_col),
                "TRACKER_HIST_ID": row.get(hist_col),
                "ISSUE_TYPE": "Missing/Invalid UPDATE_DT",
                "SEVERITY": "High",
                "DETAILS": f"Unparseable or empty UPDATE_DT: '{row.get(upd_col)}'"
            })

    # 4. Missing FOLLOW_UP_DT for SLA tickets
    if follow_col and desc_col:
        is_sla = df_raw[desc_col].astype(str).str.lower().str.contains("sla", na=False)
        sla_rows = df_raw[is_sla]
        parsed_follow = pd.to_datetime(sla_rows[follow_col], format='%d-%b-%y', errors='coerce')
        missing_follow = sla_rows[parsed_follow.isna()]
        for idx, row in missing_follow.iterrows():
            issues.append({
                "TRACKER_ID": row.get(id_col),
                "TRACKER_HIST_ID": row.get(hist_col),
                "ISSUE_TYPE": "Missing FOLLOW_UP_DT on SLA Ticket",
                "SEVERITY": "Medium",
                "DETAILS": f"SLA Ticket without valid follow-up date: '{row.get(follow_col)}'"
            })

    # 5. Missing ASSIGNED_TO_GROUP
    if grp_col:
        empty_grp = df_raw[df_raw[grp_col].isna() | (df_raw[grp_col].astype(str).str.strip() == "")]
        for idx, row in empty_grp.iterrows():
            issues.append({
                "TRACKER_ID": row.get(id_col),
                "TRACKER_HIST_ID": row.get(hist_col),
                "ISSUE_TYPE": "Missing ASSIGNED_TO_GROUP",
                "SEVERITY": "Medium",
                "DETAILS": f"Empty assigned to group."
            })

    # 6. Reconciliation difference notification
    if reconciliation_diff != 0:
        issues.append({
            "TRACKER_ID": "N/A",
            "TRACKER_HIST_ID": "N/A",
            "ISSUE_TYPE": "Reconciliation Difference",
            "SEVERITY": "Information",
            "DETAILS": (
                f"Movement calculation produced a variance of {reconciliation_diff} ticket(s) "
                f"relative to snapshot closing balance due to multi-transition tickets in reporting window."
            )
        })

    issues_df = pd.DataFrame(issues)
    if issues_df.empty:
        issues_df = pd.DataFrame(columns=["TRACKER_ID", "TRACKER_HIST_ID", "ISSUE_TYPE", "SEVERITY", "DETAILS"])

    summary = {
        "total_issues": len(issues),
        "critical_issues": len(issues_df[issues_df["SEVERITY"] == "Critical"]) if not issues_df.empty else 0,
        "high_issues": len(issues_df[issues_df["SEVERITY"] == "High"]) if not issues_df.empty else 0,
        "medium_issues": len(issues_df[issues_df["SEVERITY"] == "Medium"]) if not issues_df.empty else 0,
        "info_issues": len(issues_df[issues_df["SEVERITY"] == "Information"]) if not issues_df.empty else 0,
    }

    return summary, issues_df
