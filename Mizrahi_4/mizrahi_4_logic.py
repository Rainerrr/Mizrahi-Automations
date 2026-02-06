#!/usr/bin/env python3
"""
Mizrahi_4 Automation - Daily Tracking Data Generator

Generates an Excel workbook with daily tracking data for Mizrahi-scope funds.
Creates one sheet per fund with dates from January 1 of the current year to today.

Inputs:
  - Mutual_Funds_List.xlsx - Worksheet.csv: Fund metadata (fees, trustee info)
  - Mizrahi_4/fund-index-table.xlsx: Fund-to-index mapping (86 funds)

Outputs:
  - XLSX workbook with one sheet per fund

Usage:
    python Mizrahi_4/mizrahi_4_logic.py \
        --mutual-funds-list "Mutual_Funds_List.xlsx - Worksheet.csv" \
        --fund-index-table "Mizrahi_4/fund-index-table.xlsx" \
        --output-xlsx "Mizrahi_4/output.xlsx"
"""

from __future__ import annotations

import argparse
import csv
import datetime as dt
import logging
import sys
import uuid
from dataclasses import dataclass
from pathlib import Path
from typing import Optional

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

# Global log directory for current run
LOG_RUN_DIR: Optional[Path] = None

# Logger
logger = logging.getLogger(__name__)

# Global counter for unique table names
_table_counter = 0

# Excel styling constants
HEADER_FONT = Font(name='Arial', bold=True)
DATA_ALIGNMENT = Alignment(horizontal='right', vertical='center')

# Column definitions (1-indexed)
# (column_letter, header_text, number_format)
COLUMN_HEADERS = {
    1: ("A", "תאריך", 'd"/"m"/"yyyy'),
    2: ("B", "מסחר בקרן", None),
    3: ("C", "יום ללא חישוב", None),
    4: ("D", "מדד", "#,##0.00"),
    5: ("E", "שע\"ח", "#,##0.0000"),
    6: ("F", "שיעור שכר מנהל", "0.000%"),
    7: ("G", "שיעור שכר נאמן", "0.000%"),
    8: ("H", "מקדם שכר מצטבר", None),
    9: ("I", "שווי נכסי קרן לפני ניקויים", None),
    10: ("J", "יחידות בידי ציבור", None),
    11: ("K", "יחידות SWAP כשרה", None),
    12: ("L", "שער קרן לפני ניקויים", None),
    13: ("M", "שווי שכר מנהל יומי", None),
    14: ("N", "שווי שכר נאמן יומי", None),
    15: ("O", "שווי נכסי קרן בניקוי שכר מנהל ונאמן", None),
    16: ("P", "שער קרן בניקוי שכר מנהל ונאמן", None),
    17: ("Q", "מדד יחס", None),
    18: ("R", "הפרש עקיבה יומי", None),
    19: ("S", "שיעור דמי ניהול משתנים", "0.000%"),
    20: ("T", "דמי ניהול משתנים יומי", None),
    21: ("U", "דמי ניהול משתנים מצטבר", None),
    22: ("V", "שער קרן לאחר ניקויים", None),
    23: ("W", "עמלת פדיון ויצירה", None),
    24: ("X", "שער יצירה", None),
    25: ("Y", "שער פדיון", None),
    26: ("Z", "שווי דמי ניהול משתנים יומי", None),
    27: ("AA", "שווי נכסים לאחר ניקויים", None),
    28: ("AB", "שווי דמי ניהול משתנים מצטבר", None),
    29: ("AC", "שיעור ערבות בנקאית", None),
    30: ("AD", "שווי ערבות בנקאית", None),
    31: ("AE", "רצועת ביטחון", None),
}


# -----------------------------
# Data structures
# -----------------------------

@dataclass
class MutualFund:
    """A fund from the mutual funds list."""
    fund_id: int
    fund_name: str
    trustee_name: str
    manager_fee: float  # שכר המנהל (as percentage, e.g., 0.65 for 0.65%)
    trustee_fee: float  # שכר הנאמן (as percentage)
    variable_fee: float  # דמי ניהול משתנים (as percentage)


@dataclass
class FundData:
    """A fund with its index mapping for output."""
    fund_id: int
    fund_name: str
    index_id: str
    manager_fee: float  # Already divided by 100 (e.g., 0.0065 for 0.65%)
    trustee_fee: float  # Already divided by 100
    variable_fee: float  # Already divided by 100
    currency_code: Optional[str] = None  # סוג מטבע (קוד bfix) - None means use 1


# -----------------------------
# Helpers
# -----------------------------

def _to_str(v) -> Optional[str]:
    """Convert value to string, handling None and pandas NaN."""
    import pandas as pd
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return None
    s = str(v).strip()
    # Also check for string "nan" from pandas
    if s.lower() == "nan" or not s:
        return None
    return s


def _to_int(v) -> Optional[int]:
    """Convert value to int, handling None and non-numeric values."""
    if v is None or v == "":
        return None
    try:
        return int(float(str(v).strip()))
    except Exception:
        return None


def _to_float(v) -> Optional[float]:
    """Convert value to float, handling None and non-numeric values."""
    if v is None or v == "":
        return None
    try:
        return float(str(v).strip())
    except Exception:
        return None


# -----------------------------
# Logging setup
# -----------------------------

def setup_logging(log_base_dir: Path = Path(__file__).parent.parent / "log") -> Path:
    """Set up logging with timestamped directory."""
    global LOG_RUN_DIR

    timestamp = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
    short_uuid = str(uuid.uuid4())[:8]
    run_id = f"{timestamp}_{short_uuid}"

    run_dir = log_base_dir / run_id
    run_dir.mkdir(parents=True, exist_ok=True)
    LOG_RUN_DIR = run_dir

    log_format = logging.Formatter(
        '%(asctime)s - %(levelname)s - %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S'
    )

    # Console handler
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setLevel(logging.INFO)
    console_handler.setFormatter(log_format)

    # File handler
    file_handler = logging.FileHandler(run_dir / "mizrahi_4.log", encoding='utf-8')
    file_handler.setLevel(logging.DEBUG)
    file_handler.setFormatter(log_format)

    logger.setLevel(logging.DEBUG)
    logger.addHandler(console_handler)
    logger.addHandler(file_handler)

    logger.info("Log directory created: %s", run_dir)
    return run_dir


# -----------------------------
# Data loaders
# -----------------------------

def load_mutual_funds(csv_path: Path) -> dict[int, MutualFund]:
    """
    Load mutual funds list from CSV.

    Returns dict keyed by fund ID (מספר בורסה).
    """
    funds: dict[int, MutualFund] = {}

    with open(csv_path, 'r', encoding='utf-8-sig') as f:
        reader = csv.DictReader(f)
        if reader.fieldnames:
            reader.fieldnames = [h.strip() for h in reader.fieldnames]

        for row in reader:
            fund_id = _to_int(row.get("מספר בורסה"))
            if fund_id is None:
                continue

            fund = MutualFund(
                fund_id=fund_id,
                fund_name=_to_str(row.get("שם קרן בעברית")) or "",
                trustee_name=_to_str(row.get("שם נאמן")) or "",
                manager_fee=_to_float(row.get("שכר המנהל")) or 0.0,
                trustee_fee=_to_float(row.get("שכר הנאמן")) or 0.0,
                variable_fee=_to_float(row.get("דמי ניהול משתנים")) or 0.0,
            )
            funds[fund_id] = fund

    logger.info("Loaded %d funds from mutual funds list", len(funds))
    return funds


def load_fund_index_table(xlsx_path: Path) -> dict[int, tuple[str, Optional[str]]]:
    """
    Load fund-to-index mapping from Excel file.

    Returns dict mapping fund_id -> (index_id, currency_code).
    currency_code is None if not specified (use default 1).
    """
    import pandas as pd

    df = pd.read_excel(xlsx_path)

    # Column names from the file
    fund_col = "מס' קרן"
    index_col = "מספר מדד"
    currency_col = "סוג מטבע (קוד bfix)"

    fund_index_map: dict[int, tuple[str, Optional[str]]] = {}

    for _, row in df.iterrows():
        fund_id = _to_int(row.get(fund_col))
        index_id = _to_str(row.get(index_col))
        currency_code = _to_str(row.get(currency_col))

        if fund_id is not None and index_id is not None:
            fund_index_map[fund_id] = (index_id, currency_code)

    logger.info("Loaded %d fund-index mappings", len(fund_index_map))
    return fund_index_map


def is_fund_in_scope(fund: MutualFund) -> bool:
    """
    Check if a fund is in scope based on:
    1. Mizrahi trustee
    2. Has non-zero variable management fees (דמי ניהול משתנים)
    """
    # Check Mizrahi trustee
    if "מזרחי טפחות" not in fund.trustee_name:
        return False

    # Check variable fees - must be non-zero and non-blank
    if fund.variable_fee is None or fund.variable_fee == 0:
        return False

    return True


def get_in_scope_funds(
    mutual_funds: dict[int, MutualFund],
    fund_index_map: dict[int, tuple[str, Optional[str]]]
) -> list[FundData]:
    """
    Filter to in-scope funds: Mizrahi trustee + has index mapping + has variable fees.

    Returns list of FundData objects with fees divided by 100.
    """
    in_scope: list[FundData] = []

    for fund_id, (index_id, currency_code) in fund_index_map.items():
        fund = mutual_funds.get(fund_id)

        if fund is None:
            logger.warning("Fund %d in index table but not in mutual funds list", fund_id)
            continue

        # Check if fund is in scope (Mizrahi + has variable fees)
        if not is_fund_in_scope(fund):
            logger.debug("Fund %d (%s) excluded: not in scope (trustee=%s, variable_fee=%s)",
                        fund_id, fund.fund_name, fund.trustee_name, fund.variable_fee)
            continue

        # Create FundData with fees divided by 100
        fund_data = FundData(
            fund_id=fund_id,
            fund_name=fund.fund_name,
            index_id=index_id,
            manager_fee=fund.manager_fee / 100.0,
            trustee_fee=fund.trustee_fee / 100.0,
            variable_fee=fund.variable_fee / 100.0,
            currency_code=currency_code,
        )
        in_scope.append(fund_data)

    logger.info("Found %d in-scope funds (Mizrahi + index mapping + variable fees)", len(in_scope))
    return in_scope


def generate_date_range() -> list[dt.date]:
    """
    Generate list of dates from January 1 of current year to today.
    """
    today = dt.date.today()
    start_date = dt.date(today.year, 1, 1)

    dates: list[dt.date] = []
    current = start_date
    while current <= today:
        dates.append(current)
        current += dt.timedelta(days=1)

    logger.info("Generated date range: %s to %s (%d days)",
               start_date, today, len(dates))
    return dates


# -----------------------------
# Excel generation
# -----------------------------

def create_fund_sheet(ws, fund: FundData, dates: list[dt.date]) -> None:
    """
    Populate a worksheet for a single fund.

    Structure:
    - Row 1: Title with fund name and code (merged across columns)
    - Row 2: Headers (תאריך, מסחר בקרן, מדד, etc.)
    - Row 3+: Data rows (one per date)
    """
    global _table_counter

    # Set RTL
    ws.sheet_view.rightToLeft = True

    # Row 1: Title with fund name and code (merged across all 31 columns)
    title_text = f"{fund.fund_name} ({fund.fund_id})"
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=31)
    title_cell = ws.cell(1, 1, title_text)
    title_cell.font = Font(name='Arial', bold=True, size=18, color="FFFFFF")
    title_cell.alignment = Alignment(horizontal='right', vertical='center')
    title_cell.fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")

    # Row 2: Headers
    current_year = dt.date.today().year
    header_alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    for col_idx, (_, header_text, _) in COLUMN_HEADERS.items():
        # Special case for date column - show current year
        if col_idx == 1:
            header_text = f"תאריך ({current_year})"
        cell = ws.cell(2, col_idx, header_text if header_text else None)
        cell.font = HEADER_FONT
        cell.alignment = header_alignment

    # Write data rows
    for row_idx, date in enumerate(dates, start=3):
        # Column A: תאריך
        cell = ws.cell(row_idx, 1, date)
        cell.number_format = 'd"/"m"/"yyyy'
        cell.alignment = DATA_ALIGNMENT

        # Column D: מדד (index_id)
        cell = ws.cell(row_idx, 4, fund.index_id)
        cell.number_format = "@"  # Text format
        cell.alignment = DATA_ALIGNMENT

        # Column E: שע"ח (currency code if specified, otherwise 1)
        if fund.currency_code:
            cell = ws.cell(row_idx, 5, fund.currency_code)
            cell.number_format = "@"  # Text format for currency codes
        else:
            cell = ws.cell(row_idx, 5, 1.0)
            cell.number_format = "#,##0.0000"
        cell.alignment = DATA_ALIGNMENT

        # Column F: שיעור שכר מנהל
        cell = ws.cell(row_idx, 6, fund.manager_fee)
        cell.number_format = "0.000%"
        cell.alignment = DATA_ALIGNMENT

        # Column G: שיעור שכר נאמן
        cell = ws.cell(row_idx, 7, fund.trustee_fee)
        cell.number_format = "0.000%"
        cell.alignment = DATA_ALIGNMENT

        # Column S: שיעור דמי ניהול משתנים
        cell = ws.cell(row_idx, 19, fund.variable_fee)
        cell.number_format = "0.000%"
        cell.alignment = DATA_ALIGNMENT

    # Set column widths
    column_widths = {
        1: 15,   # A - תאריך
        2: 12,   # B - מסחר בקרן
        3: 14,   # C - יום ללא חישוב
        4: 12,   # D - מדד
        5: 10,   # E - שע"ח
        6: 16,   # F - שכר מנהל
        7: 16,   # G - שכר נאמן
        8: 16,   # H - מקדם שכר
        9: 22,   # I - שווי נכסי קרן
        10: 16,  # J - יחידות ציבור
        11: 16,  # K - יחידות SWAP
        12: 18,  # L - שער קרן לפני
        16: 26,  # M - שווי שכר מנהל
        14: 16,  # N - שווי שכר נאמן
        15: 28,  # O - שווי נכסי קרן בניקוי
        16: 26,  # P - שער קרן בניקוי
        17: 12,  # Q - מדד יחס
        18: 16,  # R - הפרש עקיבה
        19: 22,  # S - דמי ניהול משתנים
        20: 20,  # T - דמי ניהול יומי
        21: 22,  # U - דמי ניהול מצטבר
        22: 20,  # V - שער קרן לאחר
        23: 16,  # W - עמלת פדיון
        24: 12,  # X - שער יצירה
        25: 12,  # Y - שער פדיון
        26: 22,  # Z - שווי דמי ניהול יומי
        27: 22,  # AA - שווי נכסים לאחר
        28: 24,  # AB - שווי דמי ניהול מצטבר
        29: 18,  # AC - שיעור ערבות
        30: 18,  # AD - שווי ערבות
        31: 16,  # AE - רצועת ביטחון
    }

    for col_idx, width in column_widths.items():
        ws.column_dimensions[get_column_letter(col_idx)].width = width

    # Add Excel table (starting from row 2 to avoid merged title row)
    _table_counter += 1
    table_name = f"Fund_{fund.fund_id}_{_table_counter}"
    # Sanitize table name
    table_name = "".join(c if c.isalnum() or c == "_" else "_" for c in table_name)

    last_row = 2 + len(dates)  # Header row + data rows
    table_range = f"A2:AE{last_row}"

    table = Table(displayName=table_name, ref=table_range)
    style = TableStyleInfo(
        name="TableStyleMedium2",
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=True,
        showColumnStripes=False
    )
    table.tableStyleInfo = style
    ws.add_table(table)


def create_missing_funds_sheet(ws, missing_funds: list[MutualFund]) -> None:
    """
    Create a sheet listing Mizrahi funds that are missing from the index table.
    """
    global _table_counter

    # Set RTL
    ws.sheet_view.rightToLeft = True

    # Row 1: Title
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=4)
    title_cell = ws.cell(1, 1, "קרנות חסרות במקרא")
    title_cell.font = Font(name='Arial', bold=True, size=18, color="FFFFFF")
    title_cell.alignment = Alignment(horizontal='right', vertical='center')
    title_cell.fill = PatternFill(start_color="C00000", end_color="C00000", fill_type="solid")

    # Row 2: Headers
    headers = ["מספר בורסה", "שם קרן בעברית", "שם נאמן", "מצב הקרן"]
    header_alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    for col_idx, header in enumerate(headers, start=1):
        cell = ws.cell(2, col_idx, header)
        cell.font = HEADER_FONT
        cell.alignment = header_alignment

    # Data rows
    for row_idx, fund in enumerate(sorted(missing_funds, key=lambda f: f.fund_id), start=3):
        ws.cell(row_idx, 1, fund.fund_id).alignment = DATA_ALIGNMENT
        ws.cell(row_idx, 2, fund.fund_name).alignment = DATA_ALIGNMENT
        ws.cell(row_idx, 3, fund.trustee_name).alignment = DATA_ALIGNMENT
        # We don't have status in MutualFund, so leave empty or add it

    # Column widths
    ws.column_dimensions['A'].width = 14
    ws.column_dimensions['B'].width = 60
    ws.column_dimensions['C'].width = 30
    ws.column_dimensions['D'].width = 12

    # Add table
    if len(missing_funds) > 0:
        _table_counter += 1
        table_name = f"MissingFunds_{_table_counter}"
        last_row = 2 + len(missing_funds)
        table_range = f"A2:D{last_row}"

        table = Table(displayName=table_name, ref=table_range)
        style = TableStyleInfo(
            name="TableStyleMedium3",  # Red-ish style
            showFirstColumn=False,
            showLastColumn=False,
            showRowStripes=True,
            showColumnStripes=False
        )
        table.tableStyleInfo = style
        ws.add_table(table)


def generate_output_workbook(
    funds: list[FundData],
    dates: list[dt.date],
    output_path: Path,
    missing_funds: list[MutualFund] = None
) -> None:
    """
    Generate the output Excel workbook with one sheet per fund.
    """
    wb = openpyxl.Workbook()

    # Remove default sheet
    if 'Sheet' in wb.sheetnames:
        del wb['Sheet']

    # Sort funds by fund_id for consistent ordering
    sorted_funds = sorted(funds, key=lambda f: f.fund_id)

    for fund in sorted_funds:
        # Sheet name is the fund's מספר בורסה
        sheet_name = str(fund.fund_id)

        # Excel sheet names max 31 chars
        if len(sheet_name) > 31:
            sheet_name = sheet_name[:31]

        ws = wb.create_sheet(title=sheet_name)
        create_fund_sheet(ws, fund, dates)

        logger.debug("Created sheet for fund %d (%s)", fund.fund_id, fund.fund_name)

    # Add missing funds sheet at the end
    if missing_funds:
        ws_missing = wb.create_sheet(title="קרנות חסרות במקרא")
        create_missing_funds_sheet(ws_missing, missing_funds)
        logger.info("Added missing funds sheet with %d funds", len(missing_funds))

    # Save workbook
    wb.save(output_path)
    logger.info("Saved workbook to: %s", output_path)
    logger.info("Total sheets: %d", len(wb.sheetnames))


# -----------------------------
# Main
# -----------------------------

def main():
    parser = argparse.ArgumentParser(
        description="Generate daily tracking Excel workbook for Mizrahi-scope funds"
    )
    parser.add_argument(
        "--mutual-funds-list",
        required=True,
        help="Path to Mutual_Funds_List CSV file"
    )
    parser.add_argument(
        "--fund-index-table",
        required=True,
        help="Path to fund-index-table.xlsx file"
    )
    parser.add_argument(
        "--output-xlsx",
        required=True,
        help="Path for output Excel file"
    )

    args = parser.parse_args()

    # Setup logging
    setup_logging()

    logger.info("=" * 60)
    logger.info("Mizrahi_4 Daily Tracking Data Generator")
    logger.info("=" * 60)

    # Load data
    mutual_funds_path = Path(args.mutual_funds_list)
    fund_index_path = Path(args.fund_index_table)
    output_path = Path(args.output_xlsx)

    logger.info("Loading mutual funds from: %s", mutual_funds_path)
    mutual_funds = load_mutual_funds(mutual_funds_path)

    logger.info("Loading fund-index table from: %s", fund_index_path)
    fund_index_map = load_fund_index_table(fund_index_path)

    # Get in-scope funds
    in_scope_funds = get_in_scope_funds(mutual_funds, fund_index_map)

    if not in_scope_funds:
        logger.error("No in-scope funds found!")
        sys.exit(1)

    # Find in-scope funds missing from index table (Mizrahi + has variable fees)
    index_fund_ids = set(fund_index_map.keys())
    missing_funds: list[MutualFund] = []
    for fund_id, fund in mutual_funds.items():
        if is_fund_in_scope(fund) and fund_id not in index_fund_ids:
            missing_funds.append(fund)
    logger.info("Found %d in-scope funds missing from index table", len(missing_funds))

    # Generate date range
    dates = generate_date_range()

    # Generate output
    logger.info("Generating output workbook...")
    generate_output_workbook(in_scope_funds, dates, output_path, missing_funds)

    logger.info("=" * 60)
    logger.info("Complete!")
    logger.info("  Funds processed: %d", len(in_scope_funds))
    logger.info("  Date range: %s to %s (%d days)",
               dates[0], dates[-1], len(dates))
    logger.info("  Output file: %s", output_path)
    logger.info("=" * 60)


if __name__ == "__main__":
    main()
