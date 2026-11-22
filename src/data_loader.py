"""
data_loader.py
==============
Handles input CSV discovery, selection, loading, and column validation.
Never modifies source data.
"""

import os
import sys
import pandas as pd
from .validators import validate_source_columns


def find_input_csv(project_dir, input_dir_rel="input", explicit_path=None):
    """
    Locates the input CSV file based on rules:
    - If explicit_path is provided and exists, use it.
    - Check input_dir_rel (e.g. 'input').
    - If empty, check project_dir.
    - If exactly 1 CSV found, return it.
    - If multiple found, prompt user to select.
    - If none found, raise descriptive FileNotFoundError.
    """
    if explicit_path:
        if os.path.exists(explicit_path):
            return os.path.abspath(explicit_path)
        raise FileNotFoundError(f"Specified input file not found: {explicit_path}")

    search_dirs = [
        os.path.join(project_dir, input_dir_rel),
        project_dir
    ]

    csv_files = []
    found_in_dir = None

    for s_dir in search_dirs:
        if os.path.exists(s_dir):
            files = [
                os.path.join(s_dir, f) for f in os.listdir(s_dir)
                if f.lower().endswith(".csv") and not f.startswith("~$")
            ]
            if files:
                csv_files = files
                found_in_dir = s_dir
                break

    if not csv_files:
        expected_dir = os.path.join(project_dir, input_dir_rel)
        raise FileNotFoundError(
            f"\n[ERROR] No input CSV file found!\n"
            f"Please place your DART report CSV file in:\n"
            f"  -> {expected_dir}\n"
            f"Example filename: DART_REPORT_new.csv"
        )

    if len(csv_files) == 1:
        chosen = csv_files[0]
        print(f"[*] Found input CSV: {os.path.basename(chosen)} in '{os.path.relpath(found_in_dir, project_dir)}'")
        return os.path.abspath(chosen)

    # Multiple CSV files found
    print(f"\nMultiple CSV files found in '{os.path.relpath(found_in_dir, project_dir)}':")
    for idx, f in enumerate(csv_files, 1):
        size_kb = os.path.getsize(f) / 1024
        print(f"  [{idx}] {os.path.basename(f)} ({size_kb:.1f} KB)")

    while True:
        try:
            choice = input(f"Select file [1-{len(csv_files)}] (or 'Q' to quit): ").strip()
            if choice.upper() == 'Q':
                print("Operation cancelled by user.")
                sys.exit(0)
            idx = int(choice)
            if 1 <= idx <= len(csv_files):
                chosen = csv_files[idx - 1]
                print(f"[*] Selected: {os.path.basename(chosen)}")
                return os.path.abspath(chosen)
            print(f"Please enter a number between 1 and {len(csv_files)}.")
        except ValueError:
            print("Invalid input. Please enter a number or 'Q'.")


def load_dart_report(csv_path):
    """
    Loads DART CSV file safely with multiple encoding fallbacks.
    Validates required columns.
    Returns (df_raw, col_mapping).
    """
    print(f"[*] Loading DART CSV: {csv_path}")
    
    encodings = ["utf-8", "utf-8-sig", "latin-1", "cp1252"]
    df = None
    last_err = None

    for enc in encodings:
        try:
            df = pd.read_csv(csv_path, encoding=enc, low_memory=False)
            break
        except Exception as e:
            last_err = e
            continue

    if df is None:
        raise IOError(f"Failed to read CSV file {csv_path} with encodings {encodings}. Error: {last_err}")

    # Validate columns
    is_valid, missing_cols, col_mapping = validate_source_columns(df)
    if not is_valid:
        raise ValueError(
            f"DART CSV file is missing required columns:\n"
            f"  Missing: {missing_cols}\n"
            f"  Present in CSV: {list(df.columns)}"
        )

    print(f"[*] Successfully loaded {len(df)} rows with {len(df.columns)} columns.")
    return df, col_mapping
