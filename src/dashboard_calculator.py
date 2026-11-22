"""
dashboard_calculator.py
========================
Aggregates calculations from movement, aging, and SLA analyzers into a unified
dashboard dataset ready for Excel generation, logging, and reporting.
"""

import re
import pandas as pd


def extract_new_format_projects(ticket_aging_df):
    """
    Identifies distinct New Format projects and milestones from open tickets.
    Looks for pattern: 'New Format - <PROJECT_NAME>'
    """
    projects = []
    if ticket_aging_df.empty:
        return pd.DataFrame(columns=["PROJECT_NAME", "TICKET_COUNT", "SAMPLE_TICKETS", "REMARKS"])

    nf_tickets = ticket_aging_df[ticket_aging_df["CATEGORY"] == "New Format"]
    if nf_tickets.empty:
        return pd.DataFrame(columns=["PROJECT_NAME", "TICKET_COUNT", "SAMPLE_TICKETS", "REMARKS"])

    grouped = {}
    for _, row in nf_tickets.iterrows():
        desc = row["SHORT_DESCRIPTION"]
        # Extract project name after 'New Format'
        match = re.search(r'new\s*format\s*[-–:]\s*([^\s–-]+(?:\s+[^\s–-]+)*)', desc, re.IGNORECASE)
        if match:
            p_name = match.group(1).strip().upper()
        else:
            p_name = desc.strip() if desc.strip() else "METRONET"

        if p_name not in grouped:
            grouped[p_name] = []
        grouped[p_name].append(row["TRACKER_ID"])

    proj_rows = []
    for p_name, tids in grouped.items():
        sample_str = ", ".join(str(t) for t in tids[:5])
        if len(tids) > 5:
            sample_str += f" (+{len(tids) - 5} more)"
        proj_rows.append({
            "PROJECT_NAME": p_name,
            "TICKET_COUNT": len(tids),
            "SAMPLE_TICKETS": sample_str,
            "REMARKS": "Test in progress" if len(tids) > 0 else "Completed"
        })

    return pd.DataFrame(proj_rows)


def build_dashboard_payload(
    movement_results,
    aging_summary_df,
    ticket_aging_df,
    sla_summary,
    sla_group_df,
    all_tickets_sla_df,
    df_ds_history,
    ds_ticket_master,
    df_raw,
    data_quality_summary,
    data_quality_df
):
    """
    Assembles all analyzed data into a single comprehensive dictionary payload.
    """
    nf_projects_df = extract_new_format_projects(ticket_aging_df)

    # Process health indicator
    # GREEN if SLA missed is 0 or low, and reconciliation diff <= 5
    sla_missed = sla_summary.get("sla_missed_count", 0)
    if sla_missed == 0:
        health_status = "GREEN"
    elif sla_missed <= 3:
        health_status = "AMBER"
    else:
        health_status = "RED"

    payload = {
        "start_date": movement_results["start_date"],
        "end_date": movement_results["end_date"],
        "opening_snapshot_date": movement_results["opening_snapshot_date"],
        "closing_snapshot_date": movement_results["closing_snapshot_date"],
        "health_status": health_status,
        "movement": movement_results,
        "aging_summary": aging_summary_df,
        "ticket_aging_details": ticket_aging_df,
        "sla_summary": sla_summary,
        "sla_group_breakdown": sla_group_df,
        "sla_ticket_details": all_tickets_sla_df,
        "new_format_projects": nf_projects_df,
        "ticket_master": ds_ticket_master,
        "movement_audit": None, # Populated by caller
        "raw_data": df_raw,
        "data_quality_summary": data_quality_summary,
        "data_quality_details": data_quality_df
    }

    return payload
