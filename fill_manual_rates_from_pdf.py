"""
Stage 2: fill blank Manual Rates in Expiring Dump (KeyRisk) from policy PDF.

This script does NOT repeat the UW transfer or Rates Payroll lookups. Run it
after `uw_to_template_rates.py` and after the policy PDF has been downloaded.

For every blank cell in Expiring Dump Column H (Manual Rate):
- State comes from Column F.
- Class Code comes from Column G.
- The matching Rate Per $100 of Remuneration comes from the newest PDF in the
  local downloads folder.
- The rate is written back to Column H for the same row.

Column K (All-In Fund Rate) is not changed by this stage; its blank count is
reported for visibility.

Run:
    python fill_manual_rates_from_pdf.py
"""

from datetime import datetime
from pathlib import Path
import re
import shutil
import sys

try:
    import win32com.client as win32
except ImportError:
    print("ERROR: Install pywin32 using: pip install pywin32", flush=True)
    raise

try:
    import pdfplumber
except ImportError:
    print(
        "ERROR: Install PDF dependency using: "
        "python -m pip install -r requirements.txt",
        flush=True,
    )
    raise


BASE_DIR = Path(r"C:\Users\Samarth\Documents\Key Risk automation")
PDF_DOWNLOAD_DIR = BASE_DIR / "downloads"
TEMPLATE_FILE = (
    BASE_DIR
    / "template"
    / "Lotus HR Inc WC 26-27 Net Rate Comparision (KeyRisk)_Internal.xlsx"
)

EXPIRING_SHEET = "Expiring Dump (KeyRisk)"
DATA_START_ROW = 3

# Leave as None to select the newest PDF in PDF_DOWNLOAD_DIR. Assign a full
# Path here if you need to force one specific policy PDF.
POLICY_PDF_FILE = None

XL_UP = -4162


def log(message):
    print(f"[{datetime.now():%H:%M:%S}] {message}", flush=True)


def cell_text(value):
    """Convert Excel/PDF content to a stable trimmed string."""
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def as_rows(value):
    """Normalize Excel COM range output to a list of row tuples."""
    if value is None:
        return []
    if not isinstance(value, tuple):
        return [(value,)]
    if value and not isinstance(value[0], tuple):
        return [value]
    return list(value)


def normalize(value):
    """Normalize table headers for tolerant matching."""
    return "".join(
        character.casefold()
        for character in cell_text(value)
        if character.isalnum()
    )


def normalize_state(value):
    """Normalize and validate a two-letter state code."""
    state = cell_text(value).upper()
    return state if re.fullmatch(r"[A-Z]{2}", state) else ""


def normalize_class_code(value):
    """Normalize numeric class codes from Excel or PDF."""
    digits = re.sub(r"\D", "", cell_text(value))
    return (digits.lstrip("0") or "0") if digits else ""


def parse_rate(value):
    """Convert a PDF rate cell to a non-negative number."""
    if isinstance(value, (int, float)):
        rate = float(value)
        return rate if rate >= 0 else None
    text = cell_text(value).replace(",", "").replace("$", "")
    match = re.search(r"-?\d+(?:\.\d+)?", text)
    if not match:
        return None
    rate = float(match.group(0))
    return rate if rate >= 0 else None


def add_rate(rates, state, class_code, rate, page_number):
    """Store the first rate found for each normalized State/Class pair."""
    state_key = normalize_state(state)
    class_key = normalize_class_code(class_code)
    parsed_rate = parse_rate(rate)
    if not state_key or not class_key or parsed_rate is None:
        return

    key = (state_key, class_key)
    if key not in rates:
        rates[key] = parsed_rate
    elif rates[key] != parsed_rate:
        log(
            f"WARNING: multiple PDF rates for {state_key} class {class_key}: "
            f"{rates[key]} and {parsed_rate} (page {page_number}). "
            "Keeping the first."
        )


def find_table_headers(table):
    """Locate State, Class Code, and Rate Per $100 columns in a PDF table."""
    if not table:
        return None

    for start in range(min(10, len(table))):
        for depth in range(1, 4):
            rows = table[start:start + depth]
            if len(rows) < depth:
                continue
            width = max((len(row or []) for row in rows), default=0)
            headers = []
            for column in range(width):
                parts = [
                    cell_text(row[column])
                    for row in rows
                    if row and column < len(row) and row[column] is not None
                ]
                headers.append(normalize(" ".join(parts)))

            state_column = next(
                (i for i, header in enumerate(headers) if header == "state"),
                None,
            )
            class_column = next(
                (i for i, header in enumerate(headers) if "classcode" in header),
                None,
            )
            rate_column = next(
                (
                    i
                    for i, header in enumerate(headers)
                    if "rateper100" in header
                    or ("rate" in header and "remuneration" in header)
                ),
                None,
            )
            if None not in (state_column, class_column, rate_column):
                return start + depth, state_column, class_column, rate_column
    return None


def extract_page_tables(page, page_number, rates):
    """Extract rate rows from structured PDF tables."""
    for table in page.extract_tables() or []:
        header = find_table_headers(table)
        if header is None:
            continue
        data_start, state_column, class_column, rate_column = header
        required_column = max(state_column, class_column, rate_column)
        for row in table[data_start:]:
            if not row or len(row) <= required_column:
                continue
            add_rate(
                rates,
                row[state_column],
                row[class_column],
                row[rate_column],
                page_number,
            )


def extract_page_text(page, page_number, rates):
    """Fallback extraction when a PDF table has no detectable borders."""
    text = page.extract_text(layout=True) or page.extract_text() or ""
    row_start = re.compile(r"^\s*([A-Z]{2})\s+(\d{3,8})\s+(.+?)\s*$")
    # Typical ending: annual payroll, rate per $100, annual premium.
    row_end = re.compile(
        r"(?:\$?\(?-?[\d,]+(?:\.\d+)?\)?)\s+"
        r"(\d+(?:\.\d+)?)\s+"
        r"\$?\(?-?[\d,]+(?:\.\d+)?\)?\s*$"
    )

    for line in text.splitlines():
        match = row_start.match(line)
        if not match:
            continue
        state, class_code, remainder = match.groups()
        rate_match = row_end.search(remainder)
        if rate_match:
            add_rate(
                rates,
                state,
                class_code,
                rate_match.group(1),
                page_number,
            )


def select_policy_pdf():
    """Choose the configured PDF or newest PDF in the downloads folder."""
    if POLICY_PDF_FILE is not None:
        selected = Path(POLICY_PDF_FILE)
        if not selected.is_file():
            raise FileNotFoundError(f"Configured policy PDF not found: {selected}")
        return selected

    if not PDF_DOWNLOAD_DIR.is_dir():
        raise FileNotFoundError(
            f"Policy download folder not found: {PDF_DOWNLOAD_DIR}"
        )
    candidates = [
        path
        for path in PDF_DOWNLOAD_DIR.iterdir()
        if path.is_file() and path.suffix.casefold() == ".pdf"
    ]
    if not candidates:
        raise FileNotFoundError(f"No PDF found in: {PDF_DOWNLOAD_DIR}")
    return max(candidates, key=lambda path: path.stat().st_mtime)


def extract_policy_rates(pdf_path):
    """Build a mapping of (State, Class Code) to policy manual rate."""
    rates = {}
    with pdfplumber.open(pdf_path) as pdf:
        for page_number, page in enumerate(pdf.pages, start=1):
            extract_page_tables(page, page_number, rates)
            extract_page_text(page, page_number, rates)
    if not rates:
        raise ValueError(
            f"No State/Class/Rate Per $100 rows were extracted from: {pdf_path}"
        )
    return rates


def last_data_row(ws):
    """Return the final Expiring Dump data row based on State F and Class G."""
    last_state = ws.Cells(ws.Rows.Count, 6).End(XL_UP).Row
    last_class = ws.Cells(ws.Rows.Count, 7).End(XL_UP).Row
    return max(last_state, last_class)


def count_blanks(ws, column, last_row):
    """Count blank cells over the active Expiring Dump row range."""
    if last_row < DATA_START_ROW:
        return 0
    values = as_rows(
        ws.Range(
            ws.Cells(DATA_START_ROW, column),
            ws.Cells(last_row, column),
        ).Value2
    )
    return sum(1 for row in values if row[0] in (None, ""))


def collect_blank_manual_rates(ws, last_row):
    """Collect blank H rows and their State F / Class Code G keys."""
    rows = as_rows(
        ws.Range(
            ws.Cells(DATA_START_ROW, 6),
            ws.Cells(last_row, 8),
        ).Value2
    )
    missing = []
    for offset, row in enumerate(rows):
        if row[2] not in (None, ""):
            continue
        missing.append(
            (
                DATA_START_ROW + offset,
                normalize_state(row[0]),
                normalize_class_code(row[1]),
            )
        )
    return missing


def worksheet_by_name(workbook, name):
    """Return a workbook sheet by case-insensitive name."""
    for index in range(1, workbook.Worksheets.Count + 1):
        sheet = workbook.Worksheets(index)
        if sheet.Name.casefold() == name.casefold():
            return sheet
    return None


def main():
    if not TEMPLATE_FILE.is_file():
        raise FileNotFoundError(f"Template workbook not found: {TEMPLATE_FILE}")

    excel = None
    workbook = None
    try:
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False

        # Inspect the existing Stage 1 output before requiring a PDF or creating
        # a backup. If H has no blanks, Stage 2 exits without modifying files.
        workbook = excel.Workbooks.Open(
            str(TEMPLATE_FILE), UpdateLinks=0, ReadOnly=True
        )
        sheet = worksheet_by_name(workbook, EXPIRING_SHEET)
        if sheet is None:
            raise ValueError(f"Sheet not found: {EXPIRING_SHEET}")

        final_row = last_data_row(sheet)
        if final_row < DATA_START_ROW:
            raise ValueError("No Expiring Dump records found from row 3.")

        blank_h_before = count_blanks(sheet, 8, final_row)
        blank_k = count_blanks(sheet, 11, final_row)
        log(f"Blank H Manual Rates before PDF: {blank_h_before}")
        log(f"Blank K All-In Fund Rates (unchanged): {blank_k}")

        workbook.Close(SaveChanges=False)
        workbook = None

        if blank_h_before == 0:
            log("No blank H Manual Rates require PDF processing.")
            return

        pdf_path = select_policy_pdf()
        log(f"Selected newest policy PDF: {pdf_path}")
        rates = extract_policy_rates(pdf_path)
        log(f"Unique State/Class PDF rates extracted: {len(rates)}")

        backup = TEMPLATE_FILE.with_name(
            f"{TEMPLATE_FILE.stem}_before_pdf_rates_"
            f"{datetime.now():%Y%m%d_%H%M%S}{TEMPLATE_FILE.suffix}"
        )
        shutil.copy2(TEMPLATE_FILE, backup)
        log(f"Template backup created: {backup}")

        workbook = excel.Workbooks.Open(
            str(TEMPLATE_FILE), UpdateLinks=0, ReadOnly=False
        )
        sheet = worksheet_by_name(workbook, EXPIRING_SHEET)
        if sheet is None:
            raise ValueError(f"Sheet not found: {EXPIRING_SHEET}")

        # Re-evaluate rows after reopening the workbook for editing.
        final_row = last_data_row(sheet)
        blank_h_before = count_blanks(sheet, 8, final_row)
        filled = 0
        unmatched = 0
        for excel_row, state, class_code in collect_blank_manual_rates(
            sheet, final_row
        ):
            rate = rates.get((state, class_code))
            if rate is None:
                unmatched += 1
                log(
                    f"No PDF rate for H{excel_row}: "
                    f"state={state or '<blank>'}, "
                    f"class={class_code or '<blank>'}"
                )
                continue
            sheet.Cells(excel_row, 8).Value = rate
            filled += 1
            log(
                f"Filled H{excel_row}: state={state}, "
                f"class={class_code}, manual_rate={rate}"
            )

        blank_h_after = count_blanks(sheet, 8, final_row)
        expected_after = blank_h_before - filled
        if blank_h_after != expected_after:
            raise RuntimeError(
                "PDF fill validation failed: expected "
                f"{expected_after} blank H cells, found {blank_h_after}."
            )

        workbook.Save()
        log(f"Manual Rates filled from PDF: {filled}")
        log(f"PDF State/Class keys not found: {unmatched}")
        log(f"Blank H Manual Rates after PDF: {blank_h_after}")
        log(f"Blank K All-In Fund Rates still present: {blank_k}")
        log("STAGE 2 SUCCESS: PDF Manual Rate processing completed.")

    finally:
        if workbook is not None:
            try:
                # Changes are explicitly saved only after validation succeeds.
                workbook.Close(SaveChanges=False)
            except Exception:
                pass
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass


if __name__ == "__main__":
    try:
        main()
    except Exception as error:
        log(f"ERROR: {error}")
        sys.exit(1)
