"""
Extract E-Mod values from the expiring and quote PDFs and fill the template.

Mappings
--------
1. downloads/Lotus HR, Inc.pdf
   "Experience Modification" adjacent value -> Master Comparision Column I.

2. Quote/Q-279399_Lotus_HR_Inc_Master_v2.pdf
   "Experience Mod Final" row, value under Rate -> Master Comparision Column O.

Both values are written from row 3 through the final nonblank record in Column C.
The template is backed up before modification and saved only after validation.

Run:
    python fill_emod_from_pdfs.py
"""

from datetime import datetime
from pathlib import Path
import re
import shutil
import sys

try:
    import pdfplumber
except ImportError:
    print(
        "ERROR: Install PDF dependency using: "
        "python -m pip install -r requirements.txt",
        flush=True,
    )
    raise

try:
    import win32com.client as win32
except ImportError:
    print("ERROR: Install pywin32 using: pip install pywin32", flush=True)
    raise


BASE_DIR = Path(r"C:\Users\Samarth\Documents\Key Risk automation")
EXPIRING_PDF = BASE_DIR / "downloads" / "Lotus HR, Inc.pdf"
QUOTE_PDF = BASE_DIR / "Quote" / "Q-279399_Lotus_HR_Inc_Master_v2.pdf"
TEMPLATE_FILE = (
    BASE_DIR
    / "template"
    / "Lotus HR Inc WC 26-27 Net Rate Comparision (KeyRisk)_Internal.xlsx"
)

MASTER_SHEET = "Master Comparision"
DATA_START_ROW = 3
RECORD_COLUMN = 3       # C
EXPIRING_EMOD_COLUMN = 9  # I
QUOTE_EMOD_COLUMN = 15    # O

XL_UP = -4162


DECIMAL_PATTERN = re.compile(r"(?<!\d)(\d+\.\d+)(?!\d)")


def log(message):
    print(f"[{datetime.now():%H:%M:%S}] {message}", flush=True)


def validate_emod(value, label):
    """Require a plausible positive E-Mod value."""
    try:
        emod = float(value)
    except (TypeError, ValueError) as error:
        raise ValueError(f"{label} E-Mod is not numeric: {value!r}") from error
    if not 0 < emod < 10:
        raise ValueError(f"{label} E-Mod is outside expected range: {emod}")
    return emod


def page_layout_text(page):
    """Extract text while preserving enough horizontal layout for table rows."""
    return page.extract_text(layout=True) or page.extract_text() or ""


def extract_expiring_emod(pdf_path):
    """Extract the number immediately following 'Experience Modification'."""
    label = re.compile(r"Experience\s+Modification", re.IGNORECASE)

    with pdfplumber.open(pdf_path) as pdf:
        for page_number, page in enumerate(pdf.pages, start=1):
            text = page_layout_text(page)
            for line in text.splitlines():
                match = label.search(line)
                if not match:
                    continue
                adjacent = DECIMAL_PATTERN.search(line[match.end():])
                if adjacent:
                    value = validate_emod(adjacent.group(1), "Expiring PDF")
                    log(
                        f"Expiring E-Mod found on PDF page {page_number}: "
                        f"{value:.3f}"
                    )
                    return value

            # Position-aware fallback: find the label's line and take the
            # nearest decimal to its right.
            words = page.extract_words() or []
            for index, word in enumerate(words):
                if word["text"].casefold() != "experience":
                    continue
                same_line = [
                    item
                    for item in words
                    if abs(item["top"] - word["top"]) <= 4
                    and item["x0"] >= word["x0"]
                ]
                same_line.sort(key=lambda item: item["x0"])
                line_text = " ".join(item["text"] for item in same_line)
                if not label.search(line_text):
                    continue
                numeric_words = [
                    item
                    for item in same_line
                    if re.fullmatch(r"\d+\.\d+", item["text"].replace(",", ""))
                ]
                if numeric_words:
                    raw = numeric_words[-1]["text"].replace(",", "")
                    value = validate_emod(raw, "Expiring PDF")
                    log(
                        f"Expiring E-Mod found positionally on page "
                        f"{page_number}: {value:.3f}"
                    )
                    return value

    raise ValueError(
        f"Could not find an adjacent Experience Modification value in: {pdf_path}"
    )


def extract_quote_emod_from_tables(page):
    """Find an Experience Mod Final table row and return its decimal rate."""
    for table in page.extract_tables() or []:
        for row in table or []:
            if not row:
                continue
            cells = ["" if value is None else str(value).strip() for value in row]
            row_text = " ".join(cells)
            normalized = " ".join(row_text.split()).casefold()
            if "experience mod" not in normalized or "final" not in normalized:
                continue

            # Code 9898 and premium 0 are integers. The Rate is the decimal
            # value such as 0.600, so select the first decimal-valued cell.
            for cell in cells:
                match = DECIMAL_PATTERN.search(cell.replace(",", ""))
                if match:
                    return match.group(1)
    return None


def extract_quote_emod(pdf_path):
    """Extract the decimal Rate from the 'Experience Mod Final' row."""
    label = re.compile(r"Experience\s+Mod\s+Final", re.IGNORECASE)

    with pdfplumber.open(pdf_path) as pdf:
        for page_number, page in enumerate(pdf.pages, start=1):
            table_value = extract_quote_emod_from_tables(page)
            if table_value is not None:
                value = validate_emod(table_value, "Quote PDF")
                log(f"Quote E-Mod found in table on page {page_number}: {value:.3f}")
                return value

            text = page_layout_text(page)
            lines = text.splitlines()
            for line_index, line in enumerate(lines):
                if not label.search(line):
                    continue

                # In the shown Quote row, 9898 is the code and 0 is premium;
                # the only decimal on the row is the Rate (for example 0.600).
                decimal = DECIMAL_PATTERN.search(line)
                if decimal:
                    value = validate_emod(decimal.group(1), "Quote PDF")
                    log(
                        f"Quote E-Mod found on PDF page {page_number}: "
                        f"{value:.3f}"
                    )
                    return value

                # Tolerate PDF extraction splitting the row across nearby lines.
                nearby = " ".join(lines[line_index:line_index + 3])
                decimal = DECIMAL_PATTERN.search(nearby)
                if decimal:
                    value = validate_emod(decimal.group(1), "Quote PDF")
                    log(
                        f"Quote E-Mod found near row on page {page_number}: "
                        f"{value:.3f}"
                    )
                    return value

            # Position-aware fallback for split words on the same visual row.
            words = page.extract_words() or []
            for word in words:
                if word["text"].casefold() != "experience":
                    continue
                same_line = [
                    item
                    for item in words
                    if abs(item["top"] - word["top"]) <= 5
                    and item["x0"] >= word["x0"]
                ]
                same_line.sort(key=lambda item: item["x0"])
                line_text = " ".join(item["text"] for item in same_line)
                if not label.search(line_text):
                    continue
                decimals = [
                    item["text"].replace(",", "")
                    for item in same_line
                    if re.fullmatch(r"\d+\.\d+", item["text"].replace(",", ""))
                ]
                if decimals:
                    value = validate_emod(decimals[0], "Quote PDF")
                    log(
                        f"Quote E-Mod found positionally on page "
                        f"{page_number}: {value:.3f}"
                    )
                    return value

    raise ValueError(
        f"Could not find the Rate on the Experience Mod Final row in: {pdf_path}"
    )


def worksheet_by_name(workbook, name):
    """Return a worksheet by case-insensitive name."""
    for index in range(1, workbook.Worksheets.Count + 1):
        worksheet = workbook.Worksheets(index)
        if worksheet.Name.casefold() == name.casefold():
            return worksheet
    return None


def write_and_validate_emods(sheet, last_row, expiring_emod, quote_emod):
    """Fill I and O over every active Column C record and validate results."""
    expiring_range = sheet.Range(
        sheet.Cells(DATA_START_ROW, EXPIRING_EMOD_COLUMN),
        sheet.Cells(last_row, EXPIRING_EMOD_COLUMN),
    )
    quote_range = sheet.Range(
        sheet.Cells(DATA_START_ROW, QUOTE_EMOD_COLUMN),
        sheet.Cells(last_row, QUOTE_EMOD_COLUMN),
    )

    expiring_range.Value = expiring_emod
    quote_range.Value = quote_emod
    expiring_range.NumberFormat = "0.000"
    quote_range.NumberFormat = "0.000"

    # Validate the first and last row of each filled range.
    checks = (
        (sheet.Cells(DATA_START_ROW, EXPIRING_EMOD_COLUMN).Value2, expiring_emod, "I3"),
        (sheet.Cells(last_row, EXPIRING_EMOD_COLUMN).Value2, expiring_emod, f"I{last_row}"),
        (sheet.Cells(DATA_START_ROW, QUOTE_EMOD_COLUMN).Value2, quote_emod, "O3"),
        (sheet.Cells(last_row, QUOTE_EMOD_COLUMN).Value2, quote_emod, f"O{last_row}"),
    )
    for actual, expected, address in checks:
        if abs(float(actual) - expected) > 0.0000001:
            raise RuntimeError(
                f"E-Mod write validation failed at {address}: "
                f"expected {expected}, found {actual}"
            )


def main():
    for path, label in (
        (EXPIRING_PDF, "Expiring policy PDF"),
        (QUOTE_PDF, "Quote PDF"),
        (TEMPLATE_FILE, "Template workbook"),
    ):
        if not path.is_file():
            raise FileNotFoundError(f"{label} not found: {path}")

    log(f"Reading Expiring E-Mod from: {EXPIRING_PDF}")
    expiring_emod = extract_expiring_emod(EXPIRING_PDF)
    log(f"Reading Quote E-Mod from: {QUOTE_PDF}")
    quote_emod = extract_quote_emod(QUOTE_PDF)

    backup = TEMPLATE_FILE.with_name(
        f"{TEMPLATE_FILE.stem}_before_emod_"
        f"{datetime.now():%Y%m%d_%H%M%S}{TEMPLATE_FILE.suffix}"
    )
    shutil.copy2(TEMPLATE_FILE, backup)
    log(f"Template backup created: {backup}")

    excel = None
    workbook = None
    try:
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False

        workbook = excel.Workbooks.Open(
            str(TEMPLATE_FILE), UpdateLinks=0, ReadOnly=False
        )
        sheet = worksheet_by_name(workbook, MASTER_SHEET)
        if sheet is None:
            raise ValueError(f"Sheet not found: {MASTER_SHEET}")

        last_row = sheet.Cells(sheet.Rows.Count, RECORD_COLUMN).End(XL_UP).Row
        if last_row < DATA_START_ROW:
            raise ValueError("No Master Comparision records found in Column C.")

        record_count = last_row - DATA_START_ROW + 1
        log(f"Master Comparision records found: {record_count} (rows 3:{last_row})")
        log(f"Writing Expiring E-Mod {expiring_emod:.3f} to I3:I{last_row}")
        log(f"Writing Quote E-Mod {quote_emod:.3f} to O3:O{last_row}")

        write_and_validate_emods(
            sheet,
            last_row,
            expiring_emod,
            quote_emod,
        )
        workbook.Save()
        log("E-Mod ranges validated and template workbook saved.")
        log("SUCCESS: Expiring and Quote E-Mod values were filled.")

    finally:
        if workbook is not None:
            try:
                # Save occurs explicitly only after validation succeeds.
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
