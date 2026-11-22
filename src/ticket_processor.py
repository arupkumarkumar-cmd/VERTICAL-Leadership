"""
ticket_processor.py
===================
Processes DART records, filters DS tickets, sorts history numerically,
determines first DS assignment and overall ownership status.
Considers TRACKER STATUS: if 'Completed', marks ticket as Done/closed from DS team
even if the latest assigned_to_group is DS.
"""

import pandas as pd
import numpy as np


def process_ds_tickets(df_raw, col_mapping, team_identifier="DS"):
    """
    Extracts all records for tickets that have touched the DS team at least once.
    Numerically sorts history per ticket.
    Determines first DS assignment details and current overall ownership.
    
    Returns:
    - df_ds_history: DataFrame of all historical records for DS tickets, sorted numerically by hist ID.
    - ds_ticket_master: DataFrame with one row per distinct TRACKER ID containing summary attributes.
    """
    id_col = col_mapping["TRACKER ID"]
    hist_col = col_mapping["TRACKER HIST ID"]
    grp_col = col_mapping["ASSIGNED_TO_GROUP"]
    user_col = col_mapping["ASSIGNED TO USER"]
    desc_col = col_mapping["SHORT_DESCRIPTION"]
    upd_col = col_mapping["UPDATE_DT"]
    status_col = col_mapping["TRACKER STATUS"]
    follow_col = col_mapping["FOLLOW_UP_DT"]
    create_col = col_mapping["CREATE_DT"]

    team_upper = team_identifier.strip().upper()

    # Clean group string
    grp_clean = df_raw[grp_col].astype(str).str.strip().str.upper()

    # Identify distinct TRACKER IDs that ever touched DS
    ds_mask = (grp_clean == team_upper)
    ds_tracker_ids = df_raw.loc[ds_mask, id_col].unique()

    print(f"[*] Total distinct TRACKER IDs in dataset: {df_raw[id_col].nunique()}")
    print(f"[*] Distinct TRACKER IDs that touched team '{team_identifier}': {len(ds_tracker_ids)}")

    # Filter to only DS tickets
    df_ds = df_raw[df_raw[id_col].isin(ds_tracker_ids)].copy()

    # Numeric Hist ID
    df_ds["_HIST_NUM"] = pd.to_numeric(df_ds[hist_col], errors="coerce").fillna(0)
    df_ds["_GRP_CLEAN"] = df_ds[grp_col].astype(str).str.strip().str.upper()
    df_ds["_UPD_DT_PARSED"] = pd.to_datetime(df_ds[upd_col], format="%d-%b-%y", errors="coerce")
    df_ds["_FOLLOW_DT_PARSED"] = pd.to_datetime(df_ds[follow_col], format="%d-%b-%y", errors="coerce")
    df_ds["_CREATE_DT_PARSED"] = pd.to_datetime(df_ds[create_col], format="%d-%b-%y", errors="coerce")

    # Sort numerically by TRACKER ID, _HIST_NUM ascending
    df_ds = df_ds.sort_values(by=[id_col, "_HIST_NUM"], ascending=[True, True]).reset_index(drop=True)

    # Build Ticket Master: 1 row per distinct TRACKER ID
    master_rows = []

    for tid, group in df_ds.groupby(id_col, sort=False):
        # 1. First DS Assignment
        ds_records = group[group["_GRP_CLEAN"] == team_upper]
        if ds_records.empty:
            continue
        first_ds = ds_records.iloc[0]

        first_ds_hist_id = first_ds[hist_col]
        first_ds_user = str(first_ds[user_col]) if pd.notna(first_ds[user_col]) else ""
        first_ds_date = first_ds["_UPD_DT_PARSED"]
        first_ds_date_raw = first_ds[upd_col]
        short_desc = str(first_ds[desc_col]) if pd.notna(first_ds[desc_col]) else ""

        # 2. Latest Ownership & Status
        latest_rec = group.iloc[-1]
        latest_hist_id = latest_rec[hist_col]
        latest_grp = latest_rec["_GRP_CLEAN"]
        latest_user = str(latest_rec[user_col]) if pd.notna(latest_rec[user_col]) else ""
        latest_upd_date = latest_rec["_UPD_DT_PARSED"]
        latest_status = latest_rec[status_col]
        latest_follow_up = latest_rec["_FOLLOW_DT_PARSED"]
        latest_desc = str(latest_rec[desc_col]) if pd.notna(latest_rec[desc_col]) else short_desc

        # User Instruction:
        # If the last Tracker_hist_id has assigned_to_group is "DS":
        # Check tracker_status: if tracker_status is "Completed" (case-insensitive),
        # make it Done/closed from DS team.
        # Otherwise: Open.
        # If latest group is not DS: Done.
        is_completed = (str(latest_status).strip().upper() == "COMPLETED")
        if latest_grp == team_upper and not is_completed:
            ds_ticket_status = "Open"
        else:
            ds_ticket_status = "Done"

        # Additional metadata
        customer = str(latest_rec.get("CUSTOMER", "")) if pd.notna(latest_rec.get("CUSTOMER")) else ""
        priority = str(latest_rec.get("PRIORITY", "")) if pd.notna(latest_rec.get("PRIORITY")) else ""
        opened_by = str(latest_rec.get("OPENED_BY_GROUP", "")) if pd.notna(latest_rec.get("OPENED_BY_GROUP")) else ""

        master_rows.append({
            "TRACKER_ID": tid,
            "FIRST_DS_HIST_ID": first_ds_hist_id,
            "FIRST_DS_DATE": first_ds_date,
            "FIRST_DS_DATE_RAW": first_ds_date_raw,
            "FIRST_DS_USER": first_ds_user,
            "LATEST_HIST_ID": latest_hist_id,
            "LATEST_GROUP": latest_grp,
            "LATEST_USER": latest_user,
            "LATEST_UPDATE_DATE": latest_upd_date,
            "LATEST_STATUS": latest_status,
            "LATEST_FOLLOW_UP": latest_follow_up,
            "SHORT_DESCRIPTION": latest_desc,
            "CUSTOMER": customer,
            "PRIORITY": priority,
            "OPENED_BY_GROUP": opened_by,
            "DS_TICKET_STATUS": ds_ticket_status,
            "TOTAL_HISTORY_RECORDS": len(group)
        })

    ds_ticket_master = pd.DataFrame(master_rows)
    print(f"[*] Built Master Table for {len(ds_ticket_master)} distinct DS tickets.")

    return df_ds, ds_ticket_master
