"""
main.py
=======
Main entry point and orchestrator for DS Weekly Operations Dashboard.
Handles command line options, interactive date entry, logging, pipeline stages,
and final Excel validation.
"""

import os
import sys
import json
import logging
import argparse
from datetime import datetime, date
import openpyxl

# Local package imports
from .validators import parse_date, validate_date_range, run_data_quality_checks
from .data_loader import find_input_csv, load_dart_report
from .ticket_processor import process_ds_tickets
from .movement_analyzer import analyze_movements
from .aging_analyzer import analyze_aging
from .sla_analyzer import analyze_sla
from .dashboard_calculator import build_dashboard_payload
from .excel_generator import generate_excel_dashboard


def setup_logger(log_dir, base_filename="processing_log.txt"):
    """Configures file and console logging."""
    os.makedirs(log_dir, exist_ok=True)
    log_file_path = os.path.join(log_dir, base_filename)

    logger = logging.getLogger("DS_Dashboard")
    logger.setLevel(logging.INFO)

    # Avoid adding duplicate handlers if logger already initialized
    if not logger.handlers:
        file_handler = logging.FileHandler(log_file_path, mode="a", encoding="utf-8")
        file_formatter = logging.Formatter("[%(asctime)s] [%(levelname)s] %(message)s")
        file_handler.setFormatter(file_formatter)
        logger.addHandler(file_handler)

    return logger, log_file_path


def load_configuration(config_path):
    """Loads JSON config file or provides standard defaults."""
    if os.path.exists(config_path):
        with open(config_path, "r", encoding="utf-8") as f:
            return json.load(f)
    return {
        "paths": {
            "input_dir": "input",
            "output_dir": "output",
            "template_path": "template/Weekly_Operations_Report_11052026 - New Template.xlsx",
            "log_dir": "logs"
        },
        "team_identifier": "DS",
        "keywords": {"new_format": "New Format", "sla": "SLA"},
        "aging_buckets": None,
        "date_settings": {"allow_future_dates": False}
    }


def prompt_user_for_dates(allow_future=False):
    """Interactive console prompt for reporting start and end dates."""
    print("==========================================")
    print("       DS WEEKLY OPERATIONS DASHBOARD     ")
    print("==========================================")
    print("\nEnter Reporting Period (or 'Q' to quit):")

    while True:
        try:
            start_in = input("\nStart Date (DD-MM-YYYY): ").strip()
            if start_in.upper() in ["Q", "QUIT", "CANCEL"]:
                print("Operation cancelled by user.")
                sys.exit(0)
            start_dt = parse_date(start_in, allow_future=allow_future)
            break
        except ValueError as err:
            print(f"[Error] {err} Please try again.")

    while True:
        try:
            end_in = input("End Date (DD-MM-YYYY)  : ").strip()
            if end_in.upper() in ["Q", "QUIT", "CANCEL"]:
                print("Operation cancelled by user.")
                sys.exit(0)
            end_dt = parse_date(end_in, allow_future=allow_future)

            is_valid, msg = validate_date_range(start_dt, end_dt)
            if not is_valid:
                print(f"[Error] {msg} Please re-enter End Date.")
                continue
            break
        except ValueError as err:
            print(f"[Error] {err} Please try again.")

    print("\nSelected Period:")
    print(f"  {start_dt.strftime('%d-%b-%Y')} to {end_dt.strftime('%d-%b-%Y')}")

    while True:
        confirm = input("\nProceed with processing? (Y/N): ").strip().upper()
        if confirm in ["Y", "YES"]:
            return start_dt, end_dt
        elif confirm in ["N", "NO", "Q", "QUIT"]:
            print("Operation cancelled by user.")
            sys.exit(0)
        else:
            print("Please enter Y or N.")


def validate_output_workbook(output_path):
    """
    Validates that the generated Excel workbook exists, can be opened by openpyxl,
    and contains all 7 required worksheets.
    """
    required_sheets = [
        "DS_Dashboard",
        "DS_Ticket_Tracker",
        "DS_Aging_Analysis",
        "DS_Weekly_Movement",
        "DS_Movement_Audit",
        "DART_Raw_Data",
        "DS_Data_Quality"
    ]

    if not os.path.exists(output_path):
        raise FileNotFoundError(f"Output workbook not found at: {output_path}")

    wb = openpyxl.load_workbook(output_path, data_only=True)
    sheet_names = wb.sheetnames

    missing = [s for s in required_sheets if s not in sheet_names]
    if missing:
        raise ValueError(f"Generated workbook is missing required sheets: {missing}")

    return True


def run_pipeline(project_dir, start_date=None, end_date=None, explicit_csv=None, explicit_config=None, non_interactive=False):
    """Executes the complete dashboard generation pipeline."""
    # Resolve config
    config_path = explicit_config or os.path.join(project_dir, "config", "config.json")
    config = load_configuration(config_path)

    paths_cfg = config.get("paths", {})
    input_dir_rel = paths_cfg.get("input_dir", "input")
    output_dir_rel = paths_cfg.get("output_dir", "output")
    template_path_rel = paths_cfg.get("template_path", "template/Weekly_Operations_Report_11052026 - New Template.xlsx")
    log_dir_rel = paths_cfg.get("log_dir", "logs")

    log_dir = os.path.join(project_dir, log_dir_rel)
    logger, log_file = setup_logger(log_dir)

    team_id = config.get("team_identifier", "DS")
    kw_cfg = config.get("keywords", {})
    new_format_kw = kw_cfg.get("new_format", "New Format")
    sla_kw = kw_cfg.get("sla", "SLA")
    bucket_defs = config.get("aging_buckets")
    allow_future = config.get("date_settings", {}).get("allow_future_dates", False)

    logger.info("=" * 60)
    logger.info("DS WEEKLY OPERATIONS DASHBOARD EXECUTION STARTED")

    # Dates
    if start_date is None or end_date is None:
        if non_interactive:
            raise ValueError("Start date and End date must be supplied in non-interactive mode.")
        start_date, end_date = prompt_user_for_dates(allow_future=allow_future)

    print(f"\nTarget Reporting Window: {start_date.strftime('%d-%b-%Y')} to {end_date.strftime('%d-%b-%Y')}")
    logger.info(f"Reporting Window: {start_date.strftime('%d-%b-%Y')} to {end_date.strftime('%d-%b-%Y')}")

    # Stage 1: Load DART Report
    print("\n[1/7] Loading DART report...")
    csv_path = find_input_csv(project_dir, input_dir_rel=input_dir_rel, explicit_path=explicit_csv)
    df_raw, col_mapping = load_dart_report(csv_path)
    logger.info(f"Loaded input CSV: {os.path.basename(csv_path)} ({len(df_raw)} records)")

    # Stage 2: Validate source columns
    print("[2/7] Validating source columns...")
    logger.info(f"Validated all required columns successfully. Col mappings: {col_mapping}")

    # Stage 3: Process DS Tickets and Reconstruct movements
    print("[3/7] Reconstructing ticket movements...")
    df_ds_history, ds_ticket_master = process_ds_tickets(df_raw, col_mapping, team_identifier=team_id)
    movement_results, movement_audit_df = analyze_movements(
        df_ds_history, col_mapping, start_date, end_date, team_identifier=team_id
    )
    logger.info(f"Reconstructed movements: Opening={movement_results['opening_balance']}, "
                f"New={movement_results['new_tickets_count']}, "
                f"Completed={movement_results['completed_tickets_count']}, "
                f"Returned={movement_results['returned_tickets_count']}, "
                f"Actual Closing={movement_results['actual_closing_balance']}")

    # Stage 4: Balances & Reconciliation
    print("[4/7] Calculating opening and closing balances...")
    recon_diff = movement_results["reconciliation_difference"]
    if recon_diff != 0:
        logger.warning(f"Reconciliation difference noted: {recon_diff}")
    else:
        logger.info("Reconciliation balanced perfectly.")

    # Stage 5: Aging & SLA Analysis
    print("[5/7] Performing aging and SLA analysis...")
    closing_ids = movement_results["closing_ticket_ids"]
    aging_summary_df, ticket_aging_df = analyze_aging(
        closing_ids, df_ds_history, ds_ticket_master, end_date,
        bucket_defs=bucket_defs, new_format_keyword=new_format_kw
    )
    sla_summary, sla_group_df, all_tickets_sla_df = analyze_sla(
        closing_ids, df_ds_history, ds_ticket_master, end_date, sla_keyword=sla_kw
    )
    logger.info(f"Aging & SLA completed: Total Aging={aging_summary_df.loc[aging_summary_df['Category']=='Total', 'Total'].values[0]}, "
                f"SLA Missed={sla_summary['sla_missed_count']}")

    # Data Quality Validation
    dq_summary, dq_df = run_data_quality_checks(
        df_raw, col_mapping, start_date, end_date, reconciliation_diff=recon_diff
    )
    logger.info(f"Data quality checks completed: {dq_summary['total_issues']} issues logged.")

    # Assemble Dashboard Payload
    payload = build_dashboard_payload(
        movement_results=movement_results,
        aging_summary_df=aging_summary_df,
        ticket_aging_df=ticket_aging_df,
        sla_summary=sla_summary,
        sla_group_df=sla_group_df,
        all_tickets_sla_df=all_tickets_sla_df,
        df_ds_history=df_ds_history,
        ds_ticket_master=ds_ticket_master,
        df_raw=df_raw,
        data_quality_summary=dq_summary,
        data_quality_df=dq_df
    )

    # Stage 6: Generate Excel Dashboard
    print("[6/7] Generating Excel dashboard...")
    template_path = os.path.join(project_dir, template_path_rel)
    if not os.path.exists(template_path):
        # Fallback search
        template_path = os.path.join(project_dir, "Weekly_Operations_Report_11052026 - New Template.xlsx")

    out_filename = f"DS_Weekly_Operations_Dashboard_{start_date.strftime('%Y%m%d')}_{end_date.strftime('%Y%m%d')}.xlsx"
    output_path = os.path.join(project_dir, output_dir_rel, out_filename)

    generate_excel_dashboard(template_path, output_path, payload, movement_audit_df)
    logger.info(f"Workbook generated at: {output_path}")

    # Stage 7: Validate Output Workbook
    print("[7/7] Validating output workbook...")
    validate_output_workbook(output_path)
    logger.info("Workbook validated successfully.")

    # Summary completion banner
    print("\n" + "=" * 50)
    print("    DS DASHBOARD GENERATED SUCCESSFULLY!")
    print("=" * 50)
    print(f"  Reporting Window     : {start_date.strftime('%d-%b-%Y')} to {end_date.strftime('%d-%b-%Y')}")
    print(f"  Opening Balance      : {movement_results['opening_balance']}")
    print(f"  New Tickets Opened   : {movement_results['new_tickets_count']}")
    print(f"  Tickets Completed    : {movement_results['completed_tickets_count']}")
    print(f"  Tickets Returned     : {movement_results['returned_tickets_count']}")
    print(f"  Calculated Closing   : {movement_results['calculated_closing_balance']}")
    print(f"  Actual Closing       : {movement_results['actual_closing_balance']}")
    print(f"  Reconciliation Diff  : {movement_results['reconciliation_difference']}")
    print(f"  Open SLA Tickets     : {sla_summary['total_open_sla_tickets']}")
    print(f"  SLA Missed Count     : {sla_summary['sla_missed_count']}")
    print(f"  Output Excel File    : {output_path}")
    print(f"  Log File             : {log_file}")
    print("=" * 50 + "\n")

    logger.info("EXECUTION COMPLETED SUCCESSFULLY.\n")
    return output_path


def main():
    """CLI Argument Parsing and Entry Point."""
    parser = argparse.ArgumentParser(description="DS Team Weekly Operations Reporting Application")
    parser.add_argument("--start-date", "-s", type=str, help="Start Date in DD-MM-YYYY format")
    parser.add_argument("--end-date", "-e", type=str, help="End Date in DD-MM-YYYY format")
    parser.add_argument("--input-csv", "-i", type=str, help="Explicit path to DART CSV report")
    parser.add_argument("--config", "-c", type=str, help="Path to config.json")
    parser.add_argument("--non-interactive", action="store_true", help="Run without prompts (fails if required args missing)")

    args = parser.parse_args()

    # Determine Project Directory (assumed to be root of this script's parent folder)
    script_dir = os.path.dirname(os.path.abspath(__file__))
    project_dir = os.path.abspath(os.path.join(script_dir, ".."))

    start_dt = None
    end_dt = None

    if args.start_date:
        start_dt = parse_date(args.start_date, allow_future=True)
    if args.end_date:
        end_dt = parse_date(args.end_date, allow_future=True)

    try:
        run_pipeline(
            project_dir=project_dir,
            start_date=start_dt,
            end_date=end_dt,
            explicit_csv=args.input_csv,
            explicit_config=args.config,
            non_interactive=args.non_interactive
        )
    except Exception as e:
        print(f"\n[CRITICAL ERROR] Execution failed: {e}")
        logging.getLogger("DS_Dashboard").exception(f"Fatal error during execution: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
