"""
excel_generator.py
==================
Generates the comprehensive weekly operations Excel report using OpenPyXL.
Preserves the original template, updates DS_Dashboard with all movement metrics,
aging tables, and charts, and generates all 6 supporting detail sheets:
1. DS_Dashboard
2. DS_Ticket_Tracker
3. DS_Aging_Analysis
4. DS_Weekly_Movement
5. DS_Movement_Audit
6. DART_Raw_Data
7. DS_Data_Quality
"""

import os
import copy
from datetime import datetime
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, Reference


# Styling Constants
NAVY_HEADER_FILL = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
SUBHEADER_FILL = PatternFill(start_color="2F5597", end_color="2F5597", fill_type="solid")
LIGHT_BLUE_FILL = PatternFill(start_color="E9EDF4", end_color="E9EDF4", fill_type="solid")
ZEBRA_FILL = PatternFill(start_color="F2F5F9", end_color="F2F5F9", fill_type="solid")
TOTAL_ROW_FILL = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")

FONT_TITLE = Font(name="Calibri", size=16, bold=True, color="1F4E78")
FONT_SECTION = Font(name="Calibri", size=13, bold=True, color="365F91")
FONT_SUBSECTION = Font(name="Calibri", size=11, bold=True, color="1F4E78")
FONT_HEADER_WHITE = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
FONT_BOLD = Font(name="Calibri", size=11, bold=True, color="000000")
FONT_REGULAR = Font(name="Calibri", size=11, bold=False, color="000000")
FONT_MUTED = Font(name="Calibri", size=10, italic=True, color="595959")

THIN_BORDER_SIDE = Side(border_style="thin", color="D9D9D9")
THIN_BORDER = Border(left=THIN_BORDER_SIDE, right=THIN_BORDER_SIDE, top=THIN_BORDER_SIDE, bottom=THIN_BORDER_SIDE)
TOTAL_BORDER = Border(
    left=THIN_BORDER_SIDE, right=THIN_BORDER_SIDE,
    top=Side(border_style="thin", color="1F4E78"),
    bottom=Side(border_style="double", color="1F4E78")
)

ALIGN_LEFT = Alignment(horizontal="left", vertical="center")
ALIGN_CENTER = Alignment(horizontal="center", vertical="center")
ALIGN_RIGHT = Alignment(horizontal="right", vertical="center")


def auto_fit_columns(ws, min_width=12, max_width=60, padding=3):
    """Auto-fit column widths based on maximum contents."""
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            val = str(cell.value or '')
            if '\n' in val:
                lines = val.split('\n')
                max_len = max(max_len, max(len(l) for l in lines))
            else:
                max_len = max(max_len, len(val))
        ws.column_dimensions[col_letter].width = min(max_width, max(min_width, max_len + padding))


def style_cell(cell, font=None, fill=None, alignment=None, border=None, num_format=None):
    """Helper to apply styles to an openpyxl cell."""
    if font:
        cell.font = font
    if fill:
        cell.fill = fill
    if alignment:
        cell.alignment = alignment
    if border:
        cell.border = border
    if num_format:
        cell.number_format = num_format


def find_row_containing(ws, text, search_col=1, max_search_row=100):
    """Dynamically finds the row index containing text in the specified column."""
    text_lower = text.lower()
    for r in range(1, min(ws.max_row + 1, max_search_row)):
        val = ws.cell(r, search_col).value
        if val and text_lower in str(val).lower():
            return r
    return None


def populate_dashboard_sheet(ws, payload):
    """
    Populates Sheet 1 (renamed from 'Weekly Operations Summary' to 'DS_Dashboard').
    Presents ONE clear, authoritative Closing Balance reconciled with 100% precision:
    Opening Balance + New Tickets + Returned Tickets - Closed Tickets = Closing Balance
    """
    ws.title = "DS_Dashboard"

    start_date = payload["start_date"]
    end_date = payload["end_date"]
    report_date = payload["end_date"]
    health = payload["health_status"]

    start_str = start_date.strftime("%d-%b-%Y")
    end_str = end_date.strftime("%d-%b-%Y")
    rep_str = report_date.strftime("%d-%b-%Y")

    # Row 1: Title
    ws["A1"] = "Weekly Operations Review - Data Services (DS) Team"
    ws["A1"].font = FONT_TITLE

    # Row 4: Reporting Window Banner
    banner_text = f"Report Date: {rep_str}  |  Reporting Window: {start_str} to {end_str}  |  Process Health: {health}"
    ws["A4"] = banner_text
    ws["A4"].font = Font(name="Calibri", size=11, bold=True, color="595959")

    # Movement KPIs
    m = payload["movement"]
    open_bal = m["opening_balance"]
    new_t = m["new_tickets_count"]
    comp_t = m["completed_tickets_count"]
    ret_t = m["returned_tickets_count"]
    close_bal = m["actual_closing_balance"]

    prev_review_str = (start_date - pd.Timedelta(days=1)).strftime("%d-%b-%Y")

    # 1. Work Volume Table
    wv_row = find_row_containing(ws, "work volume") or 7
    r_open = wv_row + 2     # Row 9
    r_new = wv_row + 3      # Row 10
    r_closed = wv_row + 4   # Row 11
    r_close = wv_row + 5    # Row 12 (Matches template layout)

    ws.cell(r_open, 1, "Opening Balance")
    ws.cell(r_open, 2, open_bal)
    ws.cell(r_open, 3, f"Pending carryover balance from previous review ({prev_review_str})")
    style_cell(ws.cell(r_open, 1), font=FONT_REGULAR, fill=None, alignment=ALIGN_LEFT, border=THIN_BORDER)
    style_cell(ws.cell(r_open, 2), font=FONT_BOLD, fill=None, alignment=ALIGN_CENTER, border=THIN_BORDER)
    style_cell(ws.cell(r_open, 3), font=FONT_REGULAR, fill=None, alignment=ALIGN_LEFT, border=THIN_BORDER)

    ws.cell(r_new, 1, "New tickets opened this week")
    ws.cell(r_new, 2, new_t)
    ws.cell(r_new, 3, f"New tickets first assigned to DS within period ({start_str} to {end_str})")
    style_cell(ws.cell(r_new, 1), font=FONT_REGULAR, fill=None, alignment=ALIGN_LEFT, border=THIN_BORDER)
    style_cell(ws.cell(r_new, 2), font=FONT_BOLD, fill=None, alignment=ALIGN_CENTER, border=THIN_BORDER)
    style_cell(ws.cell(r_new, 3), font=FONT_REGULAR, fill=None, alignment=ALIGN_LEFT, border=THIN_BORDER)

    ws.cell(r_closed, 1, "Tickets Closed this Week")
    ws.cell(r_closed, 2, comp_t)
    ws.cell(r_closed, 3, f"Tickets successfully completed or transferred out within period ({start_str} to {end_str})")
    style_cell(ws.cell(r_closed, 1), font=FONT_REGULAR, fill=None, alignment=ALIGN_LEFT, border=THIN_BORDER)
    style_cell(ws.cell(r_closed, 2), font=FONT_BOLD, fill=None, alignment=ALIGN_CENTER, border=THIN_BORDER)
    style_cell(ws.cell(r_closed, 3), font=FONT_REGULAR, fill=None, alignment=ALIGN_LEFT, border=THIN_BORDER)

    # Single Authoritative Closing Balance (Matches template row 12)
    ws.cell(r_close, 1, "Closing Balance")
    ws.cell(r_close, 2, close_bal)
    ws.cell(r_close, 3, f"Standard formula: Opening Balance ({open_bal}) + Opened ({new_t}) - Closed ({comp_t}) = {close_bal}")
    style_cell(ws.cell(r_close, 1), font=FONT_BOLD, fill=LIGHT_BLUE_FILL, alignment=ALIGN_LEFT, border=THIN_BORDER)
    style_cell(ws.cell(r_close, 2), font=FONT_BOLD, fill=LIGHT_BLUE_FILL, alignment=ALIGN_CENTER, border=THIN_BORDER)
    style_cell(ws.cell(r_close, 3), font=FONT_REGULAR, fill=LIGHT_BLUE_FILL, alignment=ALIGN_LEFT, border=THIN_BORDER)

    # 2. Aging Breakdown Table
    aging_hdr_row = find_row_containing(ws, "aging breakdown")
    if aging_hdr_row:
        table_hdr_row = aging_hdr_row + 1
        maint_row = table_hdr_row + 1
        nf_row = table_hdr_row + 2

        aging_summary = payload["aging_summary"]
        bucket_cols = ["<2 days", "2-4 days", "4-7 days", "2 weeks", "3 weeks", "4 weeks", "More than 1 month"]

        # Ensure Header has 'Total' column
        ws.cell(table_hdr_row, 9, "Total")
        style_cell(ws.cell(table_hdr_row, 9), font=FONT_HEADER_WHITE, fill=NAVY_HEADER_FILL, alignment=ALIGN_CENTER, border=THIN_BORDER)

        # Maintenance Row
        maint_data = aging_summary[aging_summary["Category"] == "Maintenance"].iloc[0] if len(aging_summary[aging_summary["Category"] == "Maintenance"]) > 0 else {}
        ws.cell(maint_row, 1, "Maintenance")
        for col_idx, bkt in enumerate(bucket_cols, start=2):
            val = int(maint_data.get(bkt, 0))
            c = ws.cell(maint_row, col_idx, val)
            style_cell(c, font=FONT_REGULAR, alignment=ALIGN_CENTER, border=THIN_BORDER)
        c_tot = ws.cell(maint_row, 9, int(maint_data.get("Total", 0)))
        style_cell(c_tot, font=FONT_BOLD, fill=LIGHT_BLUE_FILL, alignment=ALIGN_CENTER, border=THIN_BORDER)

        # New Format Row
        nf_data = aging_summary[aging_summary["Category"] == "New Format"].iloc[0] if len(aging_summary[aging_summary["Category"] == "New Format"]) > 0 else {}
        ws.cell(nf_row, 1, "New Format")
        for col_idx, bkt in enumerate(bucket_cols, start=2):
            val = int(nf_data.get(bkt, 0))
            c = ws.cell(nf_row, col_idx, val)
            style_cell(c, font=FONT_REGULAR, alignment=ALIGN_CENTER, border=THIN_BORDER)
        c_tot2 = ws.cell(nf_row, 9, int(nf_data.get("Total", 0)))
        style_cell(c_tot2, font=FONT_BOLD, fill=LIGHT_BLUE_FILL, alignment=ALIGN_CENTER, border=THIN_BORDER)

        # Insert Total Row for Aging
        tot_row = nf_row + 1
        ws.insert_rows(tot_row, 1)
        ws.cell(tot_row, 1, "Total")
        style_cell(ws.cell(tot_row, 1), font=FONT_BOLD, fill=TOTAL_ROW_FILL, alignment=ALIGN_LEFT, border=TOTAL_BORDER)

        tot_data = aging_summary[aging_summary["Category"] == "Total"].iloc[0] if len(aging_summary[aging_summary["Category"] == "Total"]) > 0 else {}
        for col_idx, bkt in enumerate(bucket_cols, start=2):
            val = int(tot_data.get(bkt, 0))
            c = ws.cell(tot_row, col_idx, val)
            style_cell(c, font=FONT_BOLD, fill=TOTAL_ROW_FILL, alignment=ALIGN_CENTER, border=TOTAL_BORDER)
        c_grand = ws.cell(tot_row, 9, int(tot_data.get("Total", 0)))
        style_cell(c_grand, font=FONT_BOLD, fill=TOTAL_ROW_FILL, alignment=ALIGN_CENTER, border=TOTAL_BORDER)

        # Add Aging Bar Chart
        chart = BarChart()
        chart.type = "col"
        chart.style = 10
        chart.title = "Tickets Aging Breakdown (Open Tickets)"
        chart.y_axis.title = "Ticket Count"
        chart.x_axis.title = "Aging Bucket"
        chart.width = 16
        chart.height = 9.5

        data_ref = Reference(ws, min_col=2, min_row=table_hdr_row, max_col=8, max_row=nf_row)
        cats_ref = Reference(ws, min_col=2, min_row=table_hdr_row, max_col=8, max_row=table_hdr_row)
        chart.add_data(data_ref, titles_from_data=True, from_rows=True)
        chart.set_categories(cats_ref)
        ws.add_chart(chart, "J8")

    # 3. New Format Project Tracking Table
    nf_hdr_row = find_row_containing(ws, "new format project")
    if nf_hdr_row:
        table_hdr = nf_hdr_row + 1
        ws.cell(table_hdr, 1, "Milestone / Metric")
        ws.cell(table_hdr, 2, "Project / Description")
        ws.cell(table_hdr, 3, "Ticket Count")
        ws.cell(table_hdr, 4, "Remarks")
        for c in range(1, 5):
            style_cell(ws.cell(table_hdr, c), font=FONT_HEADER_WHITE, fill=NAVY_HEADER_FILL, border=THIN_BORDER)

        nf_projects = payload["new_format_projects"]
        num_projects = len(nf_projects) if not nf_projects.empty else 1

        if num_projects > 1:
            ws.insert_rows(table_hdr + 2, num_projects - 1)

        curr_row = table_hdr + 1
        if not nf_projects.empty:
            for idx, p in nf_projects.iterrows():
                ws.cell(curr_row, 1, "New Format")
                ws.cell(curr_row, 2, p["PROJECT_NAME"])
                ws.cell(curr_row, 3, p["TICKET_COUNT"])
                ws.cell(curr_row, 4, p["REMARKS"])
                for c in range(1, 5):
                    style_cell(ws.cell(curr_row, c), font=FONT_REGULAR, border=THIN_BORDER)
                curr_row += 1
        else:
            ws.cell(curr_row, 1, "New Format")
            ws.cell(curr_row, 2, "METRONET")
            ws.cell(curr_row, 3, 0)
            ws.cell(curr_row, 4, "No active new format tickets")
            for c in range(1, 5):
                style_cell(ws.cell(curr_row, c), font=FONT_REGULAR, border=THIN_BORDER)

    # 4. SLA Compliance Table
    sla_hdr_row = find_row_containing(ws, "service level agreement")
    if sla_hdr_row:
        table_hdr = sla_hdr_row + 1
        ws.cell(table_hdr, 1, "Opened By Group")
        ws.cell(table_hdr, 2, "Total Open SLA Tickets")
        ws.cell(table_hdr, 3, "SLA Miss Count")
        ws.cell(table_hdr, 4, "SLA On Track")
        ws.cell(table_hdr, 5, "SLA Miss %")
        for c in range(1, 6):
            style_cell(ws.cell(table_hdr, c), font=FONT_HEADER_WHITE, fill=NAVY_HEADER_FILL, alignment=ALIGN_CENTER, border=THIN_BORDER)

        sla_df = payload["sla_group_breakdown"]
        num_sla_rows = len(sla_df) + 1  # groups + total row

        delta = num_sla_rows - 3
        if delta > 0:
            ws.insert_rows(table_hdr + 4, delta)
        elif delta < 0:
            ws.delete_rows(table_hdr + 4, abs(delta))

        row_idx = table_hdr + 1
        for _, s_row in sla_df.iterrows():
            ws.cell(row_idx, 1, s_row["OPENED_BY_GROUP"])
            ws.cell(row_idx, 2, int(s_row["TOTAL_SLA_TICKETS"]))
            ws.cell(row_idx, 3, int(s_row["SLA_MISS_COUNT"]))
            ws.cell(row_idx, 4, int(s_row["SLA_ON_TRACK"]))
            ws.cell(row_idx, 5, f"{s_row['SLA_MISS_PCT']:.1f}%")
            for c in range(1, 6):
                style_cell(ws.cell(row_idx, c), font=FONT_REGULAR, alignment=ALIGN_CENTER, border=THIN_BORDER)
            row_idx += 1

        # SLA Total Row
        tot_sla = payload["sla_summary"]["total_open_sla_tickets"]
        tot_miss = payload["sla_summary"]["sla_missed_count"]
        tot_track = payload["sla_summary"]["sla_on_track_count"]
        tot_pct = payload["sla_summary"]["sla_missed_pct"]

        ws.cell(row_idx, 1, "Total SLA Compliance")
        ws.cell(row_idx, 2, tot_sla)
        ws.cell(row_idx, 3, tot_miss)
        ws.cell(row_idx, 4, tot_track)
        ws.cell(row_idx, 5, f"{tot_pct:.1f}%")
        for c in range(1, 6):
            style_cell(ws.cell(row_idx, c), font=FONT_BOLD, fill=TOTAL_ROW_FILL, alignment=ALIGN_CENTER, border=TOTAL_BORDER)

        # SLA Bar Chart
        sla_chart = BarChart()
        sla_chart.type = "col"
        sla_chart.style = 11
        sla_chart.title = "SLA Compliance by Opened Group"
        sla_chart.y_axis.title = "Ticket Count"
        sla_chart.x_axis.title = "Group"
        sla_chart.width = 16
        sla_chart.height = 9.5

        data_ref = Reference(ws, min_col=3, min_row=table_hdr, max_col=4, max_row=row_idx - 1)
        cats_ref = Reference(ws, min_col=1, min_row=table_hdr + 1, max_col=1, max_row=row_idx - 1)
        sla_chart.add_data(data_ref, titles_from_data=True)
        sla_chart.set_categories(cats_ref)
        ws.add_chart(sla_chart, "J24")


def create_ticket_tracker_sheet(wb, payload):
    """
    Creates Sheet 2: DS_Ticket_Tracker.
    One row per distinct TRACKER ID.
    """
    ws = wb.create_sheet(title="DS_Ticket_Tracker")
    ws.views.sheetView[0].showGridLines = True

    # Title Banner
    ws["A1"] = "DS Distinct Ticket Tracker & Master Inventory"
    ws["A1"].font = FONT_TITLE
    ws["A2"] = f"Reporting Window: {payload['start_date'].strftime('%d-%b-%Y')} to {payload['end_date'].strftime('%d-%b-%Y')}"
    ws["A2"].font = FONT_MUTED

    headers = [
        "TRACKER ID",
        "SHORT DESCRIPTION",
        "CUSTOMER",
        "PRIORITY",
        "OPENED BY GROUP",
        "FIRST DS ASSIGNED DATE",
        "FIRST DS HIST ID",
        "FIRST DS ASSIGNED USER",
        "LATEST GROUP",
        "LATEST ASSIGNED USER",
        "LATEST UPDATE DATE",
        "DS OWNERSHIP STATUS",
        "AGE (CALENDAR DAYS)",
        "AGING BUCKET",
        "CATEGORY",
        "IS SLA TICKET",
        "FOLLOW UP DATE",
        "SLA STATUS",
        "TOTAL HISTORY RECORDS"
    ]

    header_row = 4
    for col_idx, h in enumerate(headers, start=1):
        cell = ws.cell(header_row, col_idx, h)
        style_cell(cell, font=FONT_HEADER_WHITE, fill=NAVY_HEADER_FILL, alignment=ALIGN_CENTER, border=THIN_BORDER)

    # Combine master data with aging & SLA data
    master_df = payload["ticket_master"]
    aging_map = {row["TRACKER_ID"]: row for _, row in payload["ticket_aging_details"].iterrows()}
    sla_map = {row["TRACKER_ID"]: row for _, row in payload["sla_ticket_details"].iterrows()}

    curr_row = header_row + 1
    for _, t in master_df.iterrows():
        tid = t["TRACKER_ID"]
        ag = aging_map.get(tid, {})
        sl = sla_map.get(tid, {})

        first_ds_dt_str = t["FIRST_DS_DATE"].strftime("%d-%b-%Y") if pd.notna(t["FIRST_DS_DATE"]) else ""
        latest_upd_str = t["LATEST_UPDATE_DATE"].strftime("%d-%b-%Y") if pd.notna(t["LATEST_UPDATE_DATE"]) else ""
        follow_up_str = t["LATEST_FOLLOW_UP"].strftime("%d-%b-%Y") if pd.notna(t["LATEST_FOLLOW_UP"]) else ""

        age_val = ag.get("AGE_DAYS", "")
        bkt_val = ag.get("AGING_BUCKET", "")
        cat_val = ag.get("CATEGORY", "Others")
        is_sla_str = "Yes" if sl.get("IS_SLA") else "No"
        sla_stat = sl.get("SLA_STATUS", "N/A")

        row_vals = [
            tid,
            t["SHORT_DESCRIPTION"],
            t["CUSTOMER"],
            t["PRIORITY"],
            t["OPENED_BY_GROUP"],
            first_ds_dt_str,
            t["FIRST_DS_HIST_ID"],
            t["FIRST_DS_USER"],
            t["LATEST_GROUP"],
            t["LATEST_USER"],
            latest_upd_str,
            t["DS_TICKET_STATUS"],
            age_val,
            bkt_val,
            cat_val,
            is_sla_str,
            follow_up_str,
            sla_stat,
            t["TOTAL_HISTORY_RECORDS"]
        ]

        row_fill = ZEBRA_FILL if (curr_row % 2 == 0) else None
        for col_idx, val in enumerate(row_vals, start=1):
            cell = ws.cell(curr_row, col_idx, val)
            align = ALIGN_CENTER if col_idx in [1, 4, 5, 6, 7, 9, 11, 12, 13, 14, 16, 17, 18, 19] else ALIGN_LEFT
            style_cell(cell, font=FONT_REGULAR, fill=row_fill, alignment=align, border=THIN_BORDER)

        curr_row += 1

    ws.freeze_panes = "A5"
    auto_fit_columns(ws, min_width=12, max_width=45)


def create_aging_analysis_sheet(wb, payload):
    """
    Creates Sheet 3: DS_Aging_Analysis.
    Shows summary aging table and open ticket listings grouped by aging bucket and category.
    """
    ws = wb.create_sheet(title="DS_Aging_Analysis")
    ws.views.sheetView[0].showGridLines = True

    # Title
    ws["A1"] = "DS Open Ticket Aging Analysis & Bucket Distribution"
    ws["A1"].font = FONT_TITLE
    ws["A2"] = f"Evaluation as of Closing Snapshot: {payload['end_date'].strftime('%d-%b-%Y')}"
    ws["A2"].font = FONT_MUTED

    # Section 1: Summary Pivot
    ws["A4"] = "1. Aging Breakdown Summary Matrix"
    ws["A4"].font = FONT_SECTION

    aging_summary = payload["aging_summary"]
    cols = list(aging_summary.columns)

    for col_idx, col_name in enumerate(cols, start=1):
        cell = ws.cell(5, col_idx, col_name)
        style_cell(cell, font=FONT_HEADER_WHITE, fill=NAVY_HEADER_FILL, alignment=ALIGN_CENTER, border=THIN_BORDER)

    for r_idx, (_, row) in enumerate(aging_summary.iterrows(), start=6):
        is_total = (row["Category"] == "Total")
        r_fill = TOTAL_ROW_FILL if is_total else None
        r_font = FONT_BOLD if is_total else FONT_REGULAR
        r_border = TOTAL_BORDER if is_total else THIN_BORDER

        for col_idx, col_name in enumerate(cols, start=1):
            cell = ws.cell(r_idx, col_idx, row[col_name])
            align = ALIGN_LEFT if col_idx == 1 else ALIGN_CENTER
            style_cell(cell, font=r_font, fill=r_fill, alignment=align, border=r_border)

    # Section 2: Ticket Detail by Bucket
    start_row = 11
    ws.cell(start_row, 1, "2. Detailed Ticket Listing by Aging Bucket").font = FONT_SECTION

    details_df = payload["ticket_aging_details"]
    bucket_order = ["<2 days", "2-4 days", "4-7 days", "2 weeks", "3 weeks", "4 weeks", "More than 1 month"]

    curr_row = start_row + 2
    detail_headers = [
        "TRACKER ID", "CATEGORY", "SHORT DESCRIPTION", "FIRST DS DATE",
        "AGE (DAYS)", "AGING BUCKET", "LATEST ASSIGNEE", "FOLLOW UP DATE"
    ]

    for bkt in bucket_order:
        bkt_tickets = details_df[details_df["AGING_BUCKET"] == bkt]

        ws.cell(curr_row, 1, f"Aging Bucket: {bkt} ({len(bkt_tickets)} Ticket(s))")
        ws.cell(curr_row, 1).font = FONT_SUBSECTION
        ws.cell(curr_row, 1).fill = LIGHT_BLUE_FILL
        curr_row += 1

        for col_idx, h in enumerate(detail_headers, start=1):
            cell = ws.cell(curr_row, col_idx, h)
            style_cell(cell, font=FONT_HEADER_WHITE, fill=SUBHEADER_FILL, alignment=ALIGN_CENTER, border=THIN_BORDER)
        curr_row += 1

        if bkt_tickets.empty:
            ws.cell(curr_row, 1, "No open tickets in this aging bucket.")
            style_cell(ws.cell(curr_row, 1), font=FONT_MUTED)
            curr_row += 2
            continue

        for _, t in bkt_tickets.iterrows():
            first_ds_str = t["FIRST_DS_DATE"].strftime("%d-%b-%Y") if pd.notna(t["FIRST_DS_DATE"]) else ""
            follow_str = t["LATEST_FOLLOW_UP"].strftime("%d-%b-%Y") if pd.notna(t["LATEST_FOLLOW_UP"]) else ""

            row_data = [
                t["TRACKER_ID"],
                t["CATEGORY"],
                t["SHORT_DESCRIPTION"],
                first_ds_str,
                t["AGE_DAYS"],
                t["AGING_BUCKET"],
                t["LATEST_USER"],
                follow_str
            ]
            for col_idx, val in enumerate(row_data, start=1):
                cell = ws.cell(curr_row, col_idx, val)
                align = ALIGN_CENTER if col_idx in [1, 2, 4, 5, 6, 8] else ALIGN_LEFT
                style_cell(cell, font=FONT_REGULAR, alignment=align, border=THIN_BORDER)
            curr_row += 1

        curr_row += 1

    auto_fit_columns(ws, min_width=12, max_width=50)


def create_weekly_movement_sheet(wb, payload):
    """
    Creates Sheet 4: DS_Weekly_Movement.
    Shows distinct ticket IDs contributing to Opening Balance, New Tickets,
    Completed Tickets, Returned Tickets, and Closing Balance.
    """
    ws = wb.create_sheet(title="DS_Weekly_Movement")
    ws.views.sheetView[0].showGridLines = True

    # Title
    ws["A1"] = "DS Weekly Ticket Movement Breakdown & Reconciliation"
    ws["A1"].font = FONT_TITLE
    ws["A2"] = f"Reporting Window: {payload['start_date'].strftime('%d-%b-%Y')} to {payload['end_date'].strftime('%d-%b-%Y')}"
    ws["A2"].font = FONT_MUTED

    m = payload["movement"]
    master_map = {row["TRACKER_ID"]: row for _, row in payload["ticket_master"].iterrows()}

    # KPI Summary Cards
    ws["A4"] = "Movement Metric"
    ws["B4"] = "Count"
    ws["C4"] = "Reconciliation & Snapshot Notes"
    for c in range(1, 4):
        style_cell(ws.cell(4, c), font=FONT_HEADER_WHITE, fill=NAVY_HEADER_FILL, border=THIN_BORDER)

    kpi_rows = [
        ("Opening Balance", m["opening_balance"], f"Authoritative snapshot of open DS tickets as of {m['opening_snapshot_date'].strftime('%d-%b-%Y')}"),
        ("New Tickets", m["new_tickets_count"], f"First DS assignment within reporting period"),
        ("Tickets Closed this Week", m["completed_tickets_count"], f"Tickets successfully completed or transferred out of DS in period"),
        ("Closing Balance", m["actual_closing_balance"], f"Standard formula: {m['opening_balance']} + {m['new_tickets_count']} - {m['completed_tickets_count']} = {m['actual_closing_balance']}"),
        ("Returned to DS (Informational)", m["returned_tickets_count"], f"Tickets returning to DS from another team (tracked separately below in Section 3; does not impact net work volume)")
    ]

    for idx, (lbl, cnt, note) in enumerate(kpi_rows, start=5):
        is_close = ("Closing Balance" in lbl)
        f_fill = TOTAL_ROW_FILL if is_close else None
        f_font = FONT_BOLD if is_close else FONT_REGULAR
        f_border = TOTAL_BORDER if is_close else THIN_BORDER

        ws.cell(idx, 1, lbl)
        ws.cell(idx, 2, cnt)
        ws.cell(idx, 3, note)
        for c in range(1, 4):
            style_cell(ws.cell(idx, c), font=f_font, fill=f_fill, border=f_border)

    # Detailed KPI Sections
    curr_row = 12
    sections = [
        ("1. Opening Balance Ticket Roster", m["opening_ticket_ids"], "Tickets open in DS at start of period"),
        ("2. New Tickets Opened This Week", m["new_ticket_ids"], "Tickets with first DS assignment in period"),
        ("3. Tickets Returned to DS This Week", m["returned_ticket_ids"], "Tickets that returned to DS during period"),
        ("4. Tickets Closed / Completed This Week", m["completed_ticket_ids"], "Tickets completed or transferred out of DS in period"),
        ("5. Closing Balance Ticket Roster", m["closing_ticket_ids"], "Tickets open in DS at end of period")
    ]

    sec_headers = ["TRACKER ID", "SHORT DESCRIPTION", "FIRST DS DATE", "LATEST GROUP", "LATEST ASSIGNEE", "LATEST UPDATE DATE", "STATUS"]
    latest_upd_map = m.get("latest_update_dates", {})

    for sec_title, tids, sec_desc in sections:
        ws.cell(curr_row, 1, f"{sec_title} ({len(tids)} Tickets)").font = FONT_SECTION
        ws.cell(curr_row, 3, sec_desc).font = FONT_MUTED
        curr_row += 1

        for c_idx, h in enumerate(sec_headers, start=1):
            cell = ws.cell(curr_row, c_idx, h)
            style_cell(cell, font=FONT_HEADER_WHITE, fill=SUBHEADER_FILL, alignment=ALIGN_CENTER, border=THIN_BORDER)
        curr_row += 1

        if not tids:
            ws.cell(curr_row, 1, "None recorded.").font = FONT_MUTED
            curr_row += 2
            continue

        for tid in tids:
            t = master_map.get(tid, {})
            first_dt_str = t.get("FIRST_DS_DATE", "").strftime("%d-%b-%Y") if pd.notna(t.get("FIRST_DS_DATE")) else ""
            upd_dt = latest_upd_map.get(tid)
            if pd.isna(upd_dt):
                upd_dt = t.get("LATEST_UPDATE_DATE")
            latest_upd_str = upd_dt.strftime("%d-%b-%Y") if pd.notna(upd_dt) else ""

            row_data = [
                tid,
                t.get("SHORT_DESCRIPTION", ""),
                first_dt_str,
                t.get("LATEST_GROUP", ""),
                t.get("LATEST_USER", ""),
                latest_upd_str,
                t.get("DS_TICKET_STATUS", "")
            ]
            for c_idx, val in enumerate(row_data, start=1):
                cell = ws.cell(curr_row, c_idx, val)
                align = ALIGN_CENTER if c_idx in [1, 3, 4, 6, 7] else ALIGN_LEFT
                style_cell(cell, font=FONT_REGULAR, alignment=align, border=THIN_BORDER)
            curr_row += 1

        curr_row += 1

    auto_fit_columns(ws, min_width=12, max_width=50)


def create_movement_audit_sheet(wb, movement_audit_df):
    """
    Creates Sheet 5: DS_Movement_Audit.
    Complete audit trail of all historical movement records with transitions.
    """
    ws = wb.create_sheet(title="DS_Movement_Audit")
    ws.views.sheetView[0].showGridLines = True

    ws["A1"] = "DS Historical Movement Audit Trail"
    ws["A1"].font = FONT_TITLE
    ws["A2"] = "Complete chronological record of all group and user transitions."
    ws["A2"].font = FONT_MUTED

    headers = [
        "TRACKER ID",
        "TRACKER HIST ID",
        "PREVIOUS ASSIGNED GROUP",
        "CURRENT ASSIGNED GROUP",
        "PREVIOUS ASSIGNED USER",
        "CURRENT ASSIGNED USER",
        "UPDATE DT",
        "MOVEMENT TYPE",
        "IN REPORTING PERIOD"
    ]

    header_row = 4
    for c_idx, h in enumerate(headers, start=1):
        cell = ws.cell(header_row, c_idx, h)
        style_cell(cell, font=FONT_HEADER_WHITE, fill=NAVY_HEADER_FILL, alignment=ALIGN_CENTER, border=THIN_BORDER)

    curr_row = header_row + 1
    for _, row in movement_audit_df.iterrows():
        row_fill = ZEBRA_FILL if (curr_row % 2 == 0) else None
        vals = [
            row["TRACKER ID"],
            row["TRACKER HIST ID"],
            row["PREVIOUS ASSIGNED GROUP"],
            row["CURRENT ASSIGNED GROUP"],
            row["PREVIOUS ASSIGNED USER"],
            row["CURRENT ASSIGNED USER"],
            row["UPDATE_DT"],
            row["MOVEMENT TYPE"],
            "Yes" if row["IN_PERIOD"] else "No"
        ]
        for c_idx, v in enumerate(vals, start=1):
            cell = ws.cell(curr_row, c_idx, v)
            align = ALIGN_CENTER if c_idx in [1, 2, 3, 4, 7, 9] else ALIGN_LEFT
            style_cell(cell, font=FONT_REGULAR, fill=row_fill, alignment=align, border=THIN_BORDER)
        curr_row += 1

    ws.freeze_panes = "A5"
    auto_fit_columns(ws, min_width=12, max_width=45)


def create_dart_raw_data_sheet(wb, df_raw, ds_ticket_ids):
    """
    Creates Sheet 6: DART_Raw_Data.
    Preserves all records for DS tickets from source CSV.
    """
    ws = wb.create_sheet(title="DART_Raw_Data")
    ws.views.sheetView[0].showGridLines = True

    df_filtered = df_raw[df_raw["TRACKER ID"].isin(ds_ticket_ids)].copy() if "TRACKER ID" in df_raw.columns else df_raw

    cols = list(df_filtered.columns)

    for c_idx, col_name in enumerate(cols, start=1):
        cell = ws.cell(1, c_idx, col_name)
        style_cell(cell, font=FONT_HEADER_WHITE, fill=NAVY_HEADER_FILL, alignment=ALIGN_CENTER, border=THIN_BORDER)

    for r_idx, (_, row) in enumerate(df_filtered.iterrows(), start=2):
        row_fill = ZEBRA_FILL if (r_idx % 2 == 0) else None
        for c_idx, col_name in enumerate(cols, start=1):
            val = row[col_name]
            val_clean = "" if pd.isna(val) else val
            cell = ws.cell(r_idx, c_idx, val_clean)
            style_cell(cell, font=FONT_REGULAR, fill=row_fill, border=THIN_BORDER)

    ws.freeze_panes = "A2"
    auto_fit_columns(ws, min_width=10, max_width=40)


def create_data_quality_sheet(wb, payload):
    """
    Creates Sheet 7: DS_Data_Quality.
    Logs missing fields, invalid IDs, duplicate history rows, and reconciliation metrics.
    """
    ws = wb.create_sheet(title="DS_Data_Quality")
    ws.views.sheetView[0].showGridLines = True

    ws["A1"] = "DS Data Quality & Integrity Audit"
    ws["A1"].font = FONT_TITLE
    ws["A2"] = "Exceptions detected during source CSV ingestion and historical reconstruction."
    ws["A2"].font = FONT_MUTED

    # Summary KPI Cards
    dq_sum = payload["data_quality_summary"]
    ws["A4"] = "Quality Category"
    ws["B4"] = "Exception Count"
    ws["C4"] = "Integrity Severity"
    for c in range(1, 4):
        style_cell(ws.cell(4, c), font=FONT_HEADER_WHITE, fill=NAVY_HEADER_FILL, border=THIN_BORDER)

    cards = [
        ("Total Data Quality Exceptions", dq_sum["total_issues"], "Overall"),
        ("Critical Integrity Issues", dq_sum["critical_issues"], "Critical"),
        ("High Severity Anomalies", dq_sum["high_issues"], "High"),
        ("Medium Severity Warnings", dq_sum["medium_issues"], "Medium"),
        ("Informational / Reconciliation Notes", dq_sum["info_issues"], "Information")
    ]

    for idx, (cat, cnt, sev) in enumerate(cards, start=5):
        ws.cell(idx, 1, cat)
        ws.cell(idx, 2, cnt)
        ws.cell(idx, 3, sev)
        for c in range(1, 4):
            style_cell(ws.cell(idx, c), font=FONT_REGULAR if idx > 5 else FONT_BOLD, border=THIN_BORDER)

    # Detailed Table
    ws.cell(12, 1, "Detailed Exception Records").font = FONT_SECTION

    headers = ["TRACKER ID", "TRACKER HIST ID", "ISSUE TYPE", "SEVERITY", "DETAILS"]
    for c_idx, h in enumerate(headers, start=1):
        cell = ws.cell(13, c_idx, h)
        style_cell(cell, font=FONT_HEADER_WHITE, fill=SUBHEADER_FILL, alignment=ALIGN_CENTER, border=THIN_BORDER)

    issues_df = payload["data_quality_details"]
    curr_row = 14

    if issues_df.empty:
        ws.cell(curr_row, 1, "No data quality exceptions identified. Source dataset passed all checks.").font = FONT_MUTED
    else:
        for _, row in issues_df.iterrows():
            row_fill = ZEBRA_FILL if (curr_row % 2 == 0) else None
            vals = [
                row["TRACKER_ID"],
                row["TRACKER HIST ID"],
                row["ISSUE TYPE"],
                row["SEVERITY"],
                row["DETAILS"]
            ]
            for c_idx, v in enumerate(vals, start=1):
                cell = ws.cell(curr_row, c_idx, v)
                align = ALIGN_CENTER if c_idx in [1, 2, 4] else ALIGN_LEFT
                style_cell(cell, font=FONT_REGULAR, fill=row_fill, alignment=align, border=THIN_BORDER)
            curr_row += 1

    ws.freeze_panes = "A14"
    auto_fit_columns(ws, min_width=12, max_width=60)


def generate_excel_dashboard(template_path, output_path, payload, movement_audit_df):
    """
    Main workbook generation coordinator:
    1. Loads master template.
    2. Populates Sheet 1 (DS_Dashboard).
    3. Adds sheets 2-7.
    4. Saves to output_path preserving original template intact.
    """
    print(f"\n[*] Opening Excel Template: {template_path}")
    wb = openpyxl.load_workbook(template_path)

    # 1. Sheet 1: DS_Dashboard
    print("    -> Populating Sheet 1: DS_Dashboard")
    ws_dash = wb.active
    populate_dashboard_sheet(ws_dash, payload)

    # 2. Sheet 2: DS_Ticket_Tracker
    print("    -> Generating Sheet 2: DS_Ticket_Tracker")
    create_ticket_tracker_sheet(wb, payload)

    # 3. Sheet 3: DS_Aging_Analysis
    print("    -> Generating Sheet 3: DS_Aging_Analysis")
    create_aging_analysis_sheet(wb, payload)

    # 4. Sheet 4: DS_Weekly_Movement
    print("    -> Generating Sheet 4: DS_Weekly_Movement")
    create_weekly_movement_sheet(wb, payload)

    # 5. Sheet 5: DS_Movement_Audit
    print("    -> Generating Sheet 5: DS_Movement_Audit")
    create_movement_audit_sheet(wb, movement_audit_df)

    # 6. Sheet 6: DART_Raw_Data
    print("    -> Generating Sheet 6: DART_Raw_Data")
    create_dart_raw_data_sheet(wb, payload["raw_data"], payload["movement"]["closing_ticket_ids"] + payload["movement"]["opening_ticket_ids"])

    # 7. Sheet 7: DS_Data_Quality
    print("    -> Generating Sheet 7: DS_Data_Quality")
    create_data_quality_sheet(wb, payload)

    # Ensure output directory exists
    os.makedirs(os.path.dirname(output_path), exist_ok=True)

    print(f"[*] Saving Final Workbook: {output_path}")
    wb.save(output_path)
    print("    -> Workbook generated successfully!")
    return output_path
