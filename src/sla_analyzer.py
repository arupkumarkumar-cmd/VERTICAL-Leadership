"""
sla_analyzer.py
===============
Analyzes SLA compliance for tickets open at the closing snapshot.
Determines SLA breaches, missing follow-up dates, compliance rates,
and breakdown by Opened By Group (FP, EP, COM, etc.).
"""

import pandas as pd
import numpy as np


def analyze_sla(closing_ticket_ids, df_ds, ds_ticket_master, end_date, sla_keyword="SLA"):
    """
    Performs SLA analysis for tickets owned by DS at closing snapshot.
    
    Parameters:
    - closing_ticket_ids: set/list of TRACKER IDs owned by DS at close snapshot.
    - df_ds: history records.
    - ds_ticket_master: master DataFrame.
    - end_date: date object (Report End Date).
    - sla_keyword: str keyword to detect SLA tickets.
    
    Returns:
    - sla_summary: dict of overall metrics.
    - sla_group_df: DataFrame broken down by Opened By Group (FP, EP, COM, etc.).
    - sla_ticket_df: DataFrame listing details of each SLA ticket.
    """
    sla_kw_lower = sla_keyword.strip().lower()
    end_date_ts = pd.to_datetime(end_date)
    master_indexed = ds_ticket_master.set_index("TRACKER_ID")

    sla_records = []

    for tid in closing_ticket_ids:
        if tid not in master_indexed.index:
            continue
        t_info = master_indexed.loc[tid]

        desc = str(t_info.get("SHORT_DESCRIPTION", "")).strip()
        is_sla = sla_kw_lower in desc.lower()

        follow_up_val = t_info.get("LATEST_FOLLOW_UP")
        follow_up_dt = pd.to_datetime(follow_up_val) if pd.notna(follow_up_val) else pd.NaT

        sla_status = "Not SLA"
        sla_missed = False
        missing_follow_up = False

        if is_sla:
            if pd.isna(follow_up_dt):
                sla_status = "Missing Follow-up Date"
                missing_follow_up = True
            elif follow_up_dt < end_date_ts:
                sla_status = "SLA Missed"
                sla_missed = True
            else:
                sla_status = "SLA On Track"

        sla_records.append({
            "TRACKER_ID": tid,
            "SHORT_DESCRIPTION": desc,
            "IS_SLA": is_sla,
            "FOLLOW_UP_DATE": follow_up_dt,
            "SLA_STATUS": sla_status,
            "SLA_MISSED": sla_missed,
            "MISSING_FOLLOW_UP": missing_follow_up,
            "OPENED_BY_GROUP": str(t_info.get("OPENED_BY_GROUP", "")).strip().upper(),
            "CUSTOMER": t_info.get("CUSTOMER", ""),
            "PRIORITY": t_info.get("PRIORITY", ""),
            "FIRST_DS_DATE": t_info.get("FIRST_DS_DATE")
        })

    all_tickets_sla_df = pd.DataFrame(sla_records)

    # Filter to SLA tickets only for SLA KPIs
    sla_tickets_only = all_tickets_sla_df[all_tickets_sla_df["IS_SLA"]].copy()

    total_open_sla = len(sla_tickets_only)
    sla_missed_count = len(sla_tickets_only[sla_tickets_only["SLA_MISSED"]])
    missing_follow_count = len(sla_tickets_only[sla_tickets_only["MISSING_FOLLOW_UP"]])
    sla_on_track_count = total_open_sla - sla_missed_count - missing_follow_count

    sla_missed_pct = (sla_missed_count / total_open_sla * 100) if total_open_sla > 0 else 0.0

    sla_summary = {
        "total_open_sla_tickets": total_open_sla,
        "sla_missed_count": sla_missed_count,
        "sla_on_track_count": sla_on_track_count,
        "missing_follow_up_count": missing_follow_count,
        "sla_missed_pct": round(sla_missed_pct, 2)
    }

    # Breakdown by Opened By Group
    # Standard groups seen in template: FP, EP, COM
    standard_groups = ["FP", "EP", "COM"]
    found_groups = sorted(list(all_tickets_sla_df["OPENED_BY_GROUP"].unique()))
    all_groups = list(dict.fromkeys(standard_groups + [g for g in found_groups if g and g != "UNKNOWN"]))

    group_rows = []
    for grp in all_groups:
        grp_all = all_tickets_sla_df[all_tickets_sla_df["OPENED_BY_GROUP"] == grp]
        grp_sla = grp_all[grp_all["IS_SLA"]]
        grp_missed = len(grp_sla[grp_sla["SLA_MISSED"]])
        grp_total_sla = len(grp_sla)
        grp_on_track = len(grp_sla[grp_sla["SLA_STATUS"] == "SLA On Track"])
        grp_missing = len(grp_sla[grp_sla["MISSING_FOLLOW_UP"]])

        group_rows.append({
            "OPENED_BY_GROUP": grp,
            "TOTAL_OPEN_TICKETS": len(grp_all),
            "TOTAL_SLA_TICKETS": grp_total_sla,
            "SLA_MISS_COUNT": grp_missed,
            "SLA_ON_TRACK": grp_on_track,
            "MISSING_FOLLOW_UP": grp_missing,
            "SLA_MISS_PCT": round((grp_missed / grp_total_sla * 100), 2) if grp_total_sla > 0 else 0.0
        })

    sla_group_df = pd.DataFrame(group_rows)

    print("\n[SLA Analysis Summary]")
    print(f"  Total Open SLA Tickets : {total_open_sla}")
    print(f"  SLA Missed Count       : {sla_missed_count}")
    print(f"  SLA On Track Count     : {sla_on_track_count}")
    print(f"  Missing Follow-up Date : {missing_follow_count}")
    print(f"  SLA Missed Percentage  : {sla_summary['sla_missed_pct']}%")

    return sla_summary, sla_group_df, all_tickets_sla_df
