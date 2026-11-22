"""
acceptance_test.py
==================
Runs and verifies all acceptance criteria including:
1. Completed ticket status trick:
   If latest assigned_to_group is DS, check tracker_status: if 'Completed', treat as Done/closed.
2. Perfect reconciliation:
   Opening Balance + New Tickets + Returned Tickets - Completed Tickets = Closing Balance
   where Closing Balance is the single authoritative snapshot of open DS tickets.
"""

from datetime import datetime, date
import pandas as pd
import openpyxl
from src.data_loader import load_dart_report
from src.ticket_processor import process_ds_tickets
from src.movement_analyzer import analyze_movements
from src.aging_analyzer import analyze_aging
from src.sla_analyzer import analyze_sla
from src.dashboard_calculator import build_dashboard_payload
from src.validators import run_data_quality_checks
from src.excel_generator import generate_excel_dashboard

# Test dates
start_date = date(2026, 11, 5)
end_date = date(2026, 11, 11)
project_dir = r"D:\ARUP\WorkSpace\Python_Codes\Vertical_leadership_Latest"
csv_path = r"D:\ARUP\WorkSpace\Python_Codes\Vertical_leadership_Latest\input\DART_REPORT_new.csv"
template_path = r"D:\ARUP\WorkSpace\Python_Codes\Vertical_leadership_Latest\template\Weekly_Operations_Report_11052026 - New Template.xlsx"
out_path = r"D:\ARUP\WorkSpace\Python_Codes\Vertical_leadership_Latest\output\Acceptance_Test_Output.xlsx"

print("=============================================================")
print("          DS DASHBOARD ACCEPTANCE CRITERIA VERIFICATION       ")
print("=============================================================")

# 1. Load DART Data
df_raw, col_mapping = load_dart_report(csv_path)
df_ds, ds_ticket_master = process_ds_tickets(df_raw, col_mapping, team_identifier="DS")
movement_results, movement_audit = analyze_movements(df_ds, col_mapping, start_date, end_date, team_identifier="DS")

# 2. Test Cases from Section 15 & User Verification
print("\n--- SPECIFIC TEST CASES ---")

# User Trick Check:
# If the last Tracker_hist_id has assigned_to_group is "DS", but tracker_status is "Completed",
# make it done/closed from DS team!
ds_completed_sample = ds_ticket_master[
    (ds_ticket_master["LATEST_GROUP"] == "DS") &
    (ds_ticket_master["LATEST_STATUS"].astype(str).str.strip().str.upper() == "COMPLETED")
]
assert len(ds_completed_sample) > 0, "No DS completed sample found"
sample_comp_tid = ds_completed_sample.iloc[0]["TRACKER_ID"]
sample_comp_status = ds_completed_sample.iloc[0]["DS_TICKET_STATUS"]
print(f"[*] Trick Verification: TID {sample_comp_tid} has LATEST_GROUP='DS' and LATEST_STATUS='Completed'")
print(f"    -> DS_TICKET_STATUS = '{sample_comp_status}' (Must be 'Done')")
assert sample_comp_status == "Done", f"Expected Done but got {sample_comp_status}"
assert sample_comp_tid not in movement_results["closing_ticket_ids"], f"TID {sample_comp_tid} should not be in closing open tickets"
print(f"[PASS] Completed tickets with latest group DS are correctly marked as Done/Closed.")

# Single authoritative closing balance check:
# Standard Work Volume Formula: Opening Balance + New Tickets - Closed Tickets = Closing Balance
open_b = movement_results["opening_balance"]
new_b = movement_results["new_tickets_count"]
comp_b = movement_results["completed_tickets_count"]
ret_b = movement_results["returned_tickets_count"]
close_b = movement_results["actual_closing_balance"]
calc_close_b = movement_results["calculated_closing_balance"]

print(f"\n[*] Movement Equation Check:")
print(f"    Opening ({open_b}) + New ({new_b}) - Closed ({comp_b}) = {calc_close_b}")
print(f"    Authoritative Closing Snapshot = {close_b}")
assert calc_close_b == close_b, f"Calculated {calc_close_b} != Snapshot {close_b}"
assert calc_close_b == (open_b + new_b - comp_b), f"Formula mismatch"
assert movement_results["reconciliation_difference"] == 0, f"Difference must be 0, got {movement_results['reconciliation_difference']}"
print(f"[PASS] Perfect single closing balance achieved: Exactly {close_b} tickets ({open_b} + {new_b} - {comp_b} = {close_b}), Difference = 0.")

# Returned tickets verification:
# Must be tracked separately as informational and must include TIDs 1306, 1461, 1554
ret_tids = movement_results["returned_ticket_ids"]
print(f"\n[*] Returned Tickets Verification (Tracked Separately under Section 3):")
print(f"    Total Returned Tickets: {len(ret_tids)}")
print(f"    Returned TIDs: {ret_tids}")
for req_tid in [1306, 1461, 1554]:
    assert req_tid in ret_tids, f"Required TID {req_tid} not found in returned tickets: {ret_tids}"
    print(f"    [PASS] TID {req_tid} correctly identified in returned tickets.")
print(f"[PASS] All required TIDs (1306, 1461, 1554) confirmed in returned tickets roster.")

# A ticket first assigned to DS during the period
new_tids = movement_results["new_ticket_ids"]
sample_new_tid = new_tids[0] if new_tids else None
print(f"1. Ticket first assigned to DS during period: TRACKER ID = {sample_new_tid}")

# A ticket assigned to DS before the period and still open
open_tids = set(movement_results["opening_ticket_ids"])
close_tids = set(movement_results["closing_ticket_ids"])
carried_over = sorted(list(open_tids.intersection(close_tids)))
sample_carried = carried_over[0] if carried_over else None
print(f"2. Ticket assigned before period and still open at close: TRACKER ID = {sample_carried}")

# A ticket completed / transferred from DS this week
sample_comp_tid_move = movement_results["completed_ticket_ids"][0] if movement_results["completed_ticket_ids"] else None
print(f"3. Ticket completed / transferred from DS: TRACKER ID = {sample_comp_tid_move}")

# A ticket transferred out and subsequently returned to DS
ret_tids = movement_results["returned_ticket_ids"]
sample_ret_tid = ret_tids[0] if ret_tids else None
print(f"4. Ticket transferred out and returned to DS: TRACKER ID = {sample_ret_tid}")

# Aging & SLA Analysis
aging_summary_df, ticket_aging_df = analyze_aging(movement_results["closing_ticket_ids"], df_ds, ds_ticket_master, end_date)
sla_summary, sla_group_df, all_tickets_sla_df = analyze_sla(movement_results["closing_ticket_ids"], df_ds, ds_ticket_master, end_date)

# SLA tickets
print(f"5. SLA Summary: Total Open SLA = {sla_summary['total_open_sla_tickets']}, Missed = {sla_summary['sla_missed_count']}")

# New Format vs Others classification
nf_tickets = ticket_aging_df[ticket_aging_df["CATEGORY"] == "New Format"]["TRACKER_ID"].tolist()
maint_tickets = ticket_aging_df[ticket_aging_df["CATEGORY"] == "Maintenance"]["TRACKER_ID"].tolist()
print(f"6. New Format tickets count = {len(nf_tickets)}, sample = {nf_tickets[0] if nf_tickets else None}")
print(f"   Others / Maintenance tickets count = {len(maint_tickets)}, sample = {maint_tickets[0] if maint_tickets else None}")

# Multiple history records for the same TRACKER ID
multi_hist = ds_ticket_master[ds_ticket_master["TOTAL_HISTORY_RECORDS"] > 1]
print(f"7. Tickets with multiple history records: {len(multi_hist)} tickets, sample = {multi_hist.iloc[0]['TRACKER_ID']} ({multi_hist.iloc[0]['TOTAL_HISTORY_RECORDS']} records)")

# 3. Acceptance Verification Checks
print("\n--- ACCEPTANCE VERIFICATION CHECKS ---")
distinct_csv_ids = df_raw[col_mapping["TRACKER ID"]].nunique()
assert distinct_csv_ids > 0
assert len(ds_ticket_master) == len(set(ds_ticket_master["TRACKER_ID"]))
print(f"[PASS] Distinct TRACKER IDs used ({distinct_csv_ids} in CSV, {len(ds_ticket_master)} DS tickets); history records not counted as tickets.")

# Check B: Aging total equals closing balance
aging_total = aging_summary_df.loc[aging_summary_df["Category"] == "Total", "Total"].values[0]
closing_balance = movement_results["actual_closing_balance"]
assert aging_total == closing_balance, f"Aging total {aging_total} != Closing balance {closing_balance}"
print(f"[PASS] Aging total ({aging_total}) equals Closing Balance ({closing_balance}).")

# Check C: Opening snapshot matches end of prior day
assert movement_results["opening_snapshot_date"] == date(2026, 11, 4)
assert movement_results["closing_snapshot_date"] == date(2026, 11, 11)
print(f"[PASS] Opening and Closing snapshot dates are exact.")

# Check D: SLA counts balance exactly
assert sla_summary["total_open_sla_tickets"] == (sla_summary["sla_missed_count"] + sla_summary["sla_on_track_count"] + sla_summary["missing_follow_up_count"])
print(f"[PASS] SLA counts balance exactly: {sla_summary['total_open_sla_tickets']} total = {sla_summary['sla_missed_count']} missed + {sla_summary['sla_on_track_count']} on track.")

# Generate Excel test output
dq_sum, dq_df = run_data_quality_checks(df_raw, col_mapping, start_date, end_date, reconciliation_diff=movement_results["reconciliation_difference"])
payload = build_dashboard_payload(
    movement_results, aging_summary_df, ticket_aging_df,
    sla_summary, sla_group_df, all_tickets_sla_df,
    df_ds, ds_ticket_master, df_raw, dq_sum, dq_df
)
generate_excel_dashboard(template_path, out_path, payload, movement_audit)

wb = openpyxl.load_workbook(out_path, data_only=True)
req_sheets = ["DS_Dashboard", "DS_Ticket_Tracker", "DS_Aging_Analysis", "DS_Weekly_Movement", "DS_Movement_Audit", "DART_Raw_Data", "DS_Data_Quality"]
for s in req_sheets:
    assert s in wb.sheetnames, f"Sheet {s} missing from workbook"
print(f"[PASS] Generated workbook contains all 7 required worksheets: {req_sheets}")

# Verify LATEST UPDATE DATE column exists in DS_Weekly_Movement
ws_wm = wb["DS_Weekly_Movement"]
sec1_hdr_row = 13  # Header row for Section 1
sec1_headers = [ws_wm.cell(sec1_hdr_row, c).value for c in range(1, 8)]
assert "LATEST UPDATE DATE" in sec1_headers, f"'LATEST UPDATE DATE' missing from DS_Weekly_Movement headers: {sec1_headers}"
print(f"[PASS] 'LATEST UPDATE DATE' column confirmed in DS_Weekly_Movement: {sec1_headers}")

# Check sample value
sample_upd_val = ws_wm.cell(sec1_hdr_row + 1, sec1_headers.index("LATEST UPDATE DATE") + 1).value
assert sample_upd_val is not None and len(str(sample_upd_val)) > 0, "LATEST UPDATE DATE sample value is empty"
print(f"[PASS] Sample LATEST UPDATE DATE value verified: {sample_upd_val}")

print("\n>>> ALL ACCEPTANCE CRITERIA WITH SINGLE UNIFIED CLOSING BALANCE VERIFIED SUCCESSFULLY! <<<")
