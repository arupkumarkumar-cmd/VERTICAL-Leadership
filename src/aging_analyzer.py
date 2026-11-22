"""
aging_analyzer.py
=================
Computes ticket aging for tickets open in DS at the closing snapshot,
segments by aging buckets and category (New Format vs Others/Maintenance),
and validates bucket totals against closing balance.
"""

from datetime import date
import pandas as pd
import numpy as np


DEFAULT_BUCKET_DEFS = [
    {"name": "<2 days", "min_days": 0, "max_days": 1},
    {"name": "2-4 days", "min_days": 2, "max_days": 3},
    {"name": "4-7 days", "min_days": 4, "max_days": 6},
    {"name": "2 weeks", "min_days": 7, "max_days": 13},
    {"name": "3 weeks", "min_days": 14, "max_days": 20},
    {"name": "4 weeks", "min_days": 21, "max_days": 27},
    {"name": "More than 1 month", "min_days": 28, "max_days": None},
]


def assign_bucket(age_days, bucket_defs=None):
    """Assigns an age in calendar days to an aging bucket."""
    defs = bucket_defs if bucket_defs else DEFAULT_BUCKET_DEFS
    for b in defs:
        min_d = b["min_days"]
        max_d = b["max_days"]
        if max_d is None:
            if age_days >= min_d:
                return b["name"]
        else:
            if min_d <= age_days <= max_d:
                return b["name"]
    return defs[-1]["name"]


def analyze_aging(closing_ticket_ids, df_ds, ds_ticket_master, end_date, bucket_defs=None, new_format_keyword="New Format"):
    """
    Performs aging analysis for all tickets owned by DS at closing snapshot.
    
    Parameters:
    - closing_ticket_ids: list/set of TRACKER IDs owned by DS at closing snapshot.
    - df_ds: history DataFrame.
    - ds_ticket_master: master DataFrame with first DS assignment info.
    - end_date: date object (Report End Date).
    - bucket_defs: list of bucket definitions.
    - new_format_keyword: keyword for New Format categorization.
    
    Returns:
    - aging_summary_df: DataFrame with Category vs Bucket counts and Total column.
    - ticket_aging_df: DataFrame with individual open ticket aging details.
    """
    if bucket_defs is None:
        bucket_defs = DEFAULT_BUCKET_DEFS

    bucket_names = [b["name"] for b in bucket_defs]
    nf_kw_lower = new_format_keyword.strip().lower()

    master_indexed = ds_ticket_master.set_index("TRACKER_ID")
    end_date_ts = pd.to_datetime(end_date)

    ticket_records = []

    for tid in closing_ticket_ids:
        if tid not in master_indexed.index:
            continue
        t_info = master_indexed.loc[tid]

        first_ds_dt = t_info["FIRST_DS_DATE"]
        if pd.isna(first_ds_dt):
            age_days = 0
        else:
            age_days = max(0, (end_date_ts - first_ds_dt).days)

        bucket_name = assign_bucket(age_days, bucket_defs)

        # Category
        desc = str(t_info.get("SHORT_DESCRIPTION", "")).strip()
        is_nf = nf_kw_lower in desc.lower()
        category = "New Format" if is_nf else "Maintenance"

        ticket_records.append({
            "TRACKER_ID": tid,
            "FIRST_DS_DATE": first_ds_dt,
            "AGE_DAYS": age_days,
            "AGING_BUCKET": bucket_name,
            "CATEGORY": category,
            "SHORT_DESCRIPTION": desc,
            "CUSTOMER": t_info.get("CUSTOMER", ""),
            "PRIORITY": t_info.get("PRIORITY", ""),
            "OPENED_BY_GROUP": t_info.get("OPENED_BY_GROUP", ""),
            "LATEST_USER": t_info.get("LATEST_USER", ""),
            "LATEST_FOLLOW_UP": t_info.get("LATEST_FOLLOW_UP", pd.NaT)
        })

    ticket_aging_df = pd.DataFrame(ticket_records)

    # Build Summary Table
    # Rows: Maintenance, New Format, Total
    summary_data = {
        "Maintenance": {b: 0 for b in bucket_names},
        "New Format": {b: 0 for b in bucket_names}
    }

    if not ticket_aging_df.empty:
        for _, row in ticket_aging_df.iterrows():
            cat = row["CATEGORY"]
            bkt = row["AGING_BUCKET"]
            if cat in summary_data and bkt in summary_data[cat]:
                summary_data[cat][bkt] += 1

    summary_rows = []
    for cat in ["Maintenance", "New Format"]:
        row_dict = {"Category": cat}
        row_total = 0
        for bkt in bucket_names:
            cnt = summary_data[cat][bkt]
            row_dict[bkt] = cnt
            row_total += cnt
        row_dict["Total"] = row_total
        summary_rows.append(row_dict)

    # Add Total row
    total_row = {"Category": "Total"}
    grand_total = 0
    for bkt in bucket_names:
        bkt_total = summary_data["Maintenance"][bkt] + summary_data["New Format"][bkt]
        total_row[bkt] = bkt_total
        grand_total += bkt_total
    total_row["Total"] = grand_total
    summary_rows.append(total_row)

    aging_summary_df = pd.DataFrame(summary_rows)

    # Validation: Sum of aging must equal closing balance
    closing_count = len(closing_ticket_ids)
    if grand_total != closing_count:
        print(f"[WARNING] Aging total ({grand_total}) differs from Closing Balance ({closing_count})!")
    else:
        print(f"[*] Aging Validation Passed: Total aging ({grand_total}) matches Closing Balance ({closing_count}).")

    return aging_summary_df, ticket_aging_df
