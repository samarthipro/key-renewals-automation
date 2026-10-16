"""
Extract E-Mod values from the expiring and quote PDFs and fill the template.

Mappings
--------
1. downloads/Lotus HR, Inc.pdf
   "Experience Modification" adjacent value -> Master Comparision Column I.

2. Quote/Q-279399_Lotus_HR_Inc_Master_v2.pdf
   "Experience Mod Final" row, value under Rate -> Master Comparision Column O.

3. Quote/Q-279399_Lotus_HR_Inc_Master_v2.pdf
   Every "Expense Constant" and "Expense Constant Applied to" pair is matched
   to State in Column F. For each state, the value is written to Column W on
   the row with the highest Payroll in Column H.

E-Mod values are written from row 3 through the final nonblank record in
Column C. Expense constants support multiple states in one PDF. The template
is backed up before modification and saved only after validation.

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
RECORD_COLUMN = 3         # C
STATE_COLUMN = 6          # F
PAYROLL_COLUMN = 8        # H
EXPIRING_EMOD_COLUMN = 9  # I
QUOTE_EMOD_COLUMN = 15    # O
EXPENSE_CONSTANT_COLUMN = 23  # W

XL_UP = -4162


DECIMAL_PATTERN = re.compile(r"(?<!\d)(\d+\.\d+)(?!\d)")
EXPENSE_CONSTANT_STATE_LABEL = re.compile(
    r"Expense\s+Constant\s+Applied\s+to", re.IGNORECASE
)
EXPENSE_CONSTANT_LABEL = re.compile(
    r"Expense\s+Constant(?!\s+Applied\s+to)", re.IGNORECASE
)
EXPENSE_NUMBER_PATTERN = re.compile(
    r"(?<![\w.])\$?\s*(-?\d+(?:,\d{3})*(?:\.\d+)?)(?![\w.])"
)
STATE_VALUE_PATTERN = re.compile(r"\b([A-Za-z]{2})\b")


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


def normalize_state(value):
    """Return a normalized two-letter state code, or an empty string."""
    if value is None:
        return ""
    state = str(value).strip().upper()
    return state if re.fullmatch(r"[A-Z]{2}", state) else ""


def parse_nonnegative_number(value):
    """Parse an Excel/PDF number while tolerating currency and commas."""
    if isinstance(value, bool) or value is None:
        return None
    if isinstance(value, (int, float)):
        number = float(value)
    else:
        match = EXPENSE_NUMBER_PATTERN.search(str(value).strip())
        if not match:
            return None
        number = float(match.group(1).replace(",", ""))
    return number if number >= 0 and number != float("inf") else None


def expense_records_from_lines(lines):
    """Extract all labeled values from rows without crossing into other fields."""
    records = []
    amount_label_count = 0
    state_label_count = 0

    for line in lines:
        labels = []
        for match in EXPENSE_CONSTANT_STATE_LABEL.finditer(line):
            labels.append((match.start(), match.end(), "state"))
            state_label_count += 1
        for match in EXPENSE_CONSTANT_LABEL.finditer(line):
            labels.append((match.start(), match.end(), "amount"))
            amount_label_count += 1
        labels.sort(key=lambda item: item[0])

        # A flattened table can contain more than one labeled field on a row.
        # Limit each value search to the text before the next label so an
        # unrelated class code, payroll, or premium cannot be used.
        for index, (_, label_end, record_type) in enumerate(labels):
            value_end = labels[index + 1][0] if index + 1 < len(labels) else None
            value_text = line[label_end:value_end]
            if record_type == "state":
                match = STATE_VALUE_PATTERN.search(value_text)
                state = normalize_state(match.group(1) if match else None)
                if state:
                    records.append(("state", state))
            else:
                match = EXPENSE_NUMBER_PATTERN.search(value_text)
                amount = parse_nonnegative_number(match.group(1) if match else None)
                if amount is not None:
                    records.append(("amount", amount))

    return records, amount_label_count, state_label_count


def pair_expense_records(records, source):
    """Pair adjacent amount/state records and reject ambiguous PDF content."""
    constants = {}
    pending_amount = None
    pending_state = None

    for record_type, value in records:
        if record_type == "amount":
            if pending_amount is not None:
                raise ValueError(
                    f"Found consecutive unpaired Expense Constant values in {source}."
                )
            pending_amount = value
        else:
            if pending_state is not None:
                raise ValueError(
                    f"Found consecutive unpaired expense-constant states in {source}."
                )
            pending_state = value

        if pending_amount is None or pending_state is None:
            continue

        existing = constants.get(pending_state)
        if existing is not None and abs(existing - pending_amount) > 0.0000001:
            raise ValueError(
                f"Conflicting Expense Constant values for {pending_state} in "
                f"{source}: {existing} and {pending_amount}"
            )
        constants[pending_state] = pending_amount
        pending_amount = None
        pending_state = None

    if pending_amount is not None or pending_state is not None:
        raise ValueError(f"Found an unpaired Expense Constant entry in {source}.")
    return constants


def merge_expense_constants(target, additions, source):
    """Merge independently extracted representations without hiding conflicts."""
    for state, amount in additions.items():
        existing = target.get(state)
        if existing is not None and abs(existing - amount) > 0.0000001:
            raise ValueError(
                f"Conflicting Expense Constant values for {state} across PDF "
                f"representations: {existing} and {amount} ({source})"
            )
        target[state] = amount


def extract_expense_constants(pdf_path):
    """Extract every Expense Constant and Applied-to state pair from a PDF."""
    text_lines = []
    table_lines = []

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text_lines.extend(page_layout_text(page).splitlines())
            for table in page.extract_tables() or []:
                for row in table or []:
                    if row:
                        table_lines.append(
                            " ".join(
                                "" if cell is None else str(cell).strip()
                                for cell in row
                            )
                        )

    source = str(pdf_path)
    constants = {}
    found_label = False
    maximum_expected_pairs = 0
    maximum_complete_pairs = 0
    extraction_errors = []

    # Reconcile table and layout-text representations. A partial text match
    # must not suppress a complete table extraction (or vice versa).
    for representation, lines in (
        ("PDF tables", table_lines),
        ("PDF layout text", text_lines),
    ):
        records, amount_count, state_count = expense_records_from_lines(lines)
        if amount_count == 0 and state_count == 0:
            continue
        found_label = True
        maximum_expected_pairs = max(
            maximum_expected_pairs, amount_count, state_count
        )
        try:
            extracted = pair_expense_records(
                records, f"{source} ({representation})"
            )
        except ValueError as error:
            extraction_errors.append(str(error))
            continue

        complete_pairs = len(records) // 2
        maximum_complete_pairs = max(maximum_complete_pairs, complete_pairs)
        merge_expense_constants(constants, extracted, representation)

    if not found_label:
        raise ValueError(f"No Expense Constant entries were found in: {pdf_path}")
    if not constants or maximum_complete_pairs < maximum_expected_pairs:
        details = "; ".join(extraction_errors) or "labels and values were incomplete"
        raise ValueError(
            f"Could not safely pair every Expense Constant in {pdf_path}: {details}"
        )

    for state, amount in constants.items():
        log(f"Expense Constant found: {state} -> {amount:g}")
    return constants


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


def write_and_validate_expense_constants(sheet, last_row, constants):
    """Write each constant to W on its state's highest-payroll row."""
    state_rows = {state: [] for state in constants}
    payroll_rows = {state: [] for state in constants}

    for row in range(DATA_START_ROW, last_row + 1):
        state = normalize_state(sheet.Cells(row, STATE_COLUMN).Value2)
        if state not in constants:
            continue
        state_rows[state].append(row)
        payroll = parse_nonnegative_number(
            sheet.Cells(row, PAYROLL_COLUMN).Value2
        )
        if payroll is not None:
            payroll_rows[state].append((payroll, row))

    plan = {}
    for state, amount in constants.items():
        if not state_rows[state]:
            raise ValueError(
                f"Expense Constant state {state} was not found in Column F."
            )
        if not payroll_rows[state]:
            raise ValueError(
                f"State {state} has no numeric Payroll values in Column H."
            )

        highest_payroll = max(payroll for payroll, _ in payroll_rows[state])
        highest_rows = [
            row
            for payroll, row in payroll_rows[state]
            if abs(payroll - highest_payroll) <= 0.0000001
        ]
        target_row = min(highest_rows)
        if len(highest_rows) > 1:
            log(
                f"State {state} has {len(highest_rows)} rows tied for highest "
                f"Payroll {highest_payroll:g}; using topmost row {target_row}."
            )
        plan[state] = (target_row, highest_payroll, amount)

    # Clear earlier values for only the states being updated so a repeated run
    # cannot leave a constant on a row that is no longer the highest payroll.
    for state, rows in state_rows.items():
        for row in rows:
            sheet.Cells(row, EXPENSE_CONSTANT_COLUMN).ClearContents()

        target_row, highest_payroll, amount = plan[state]
        sheet.Cells(target_row, EXPENSE_CONSTANT_COLUMN).Value = amount
        log(
            f"W{target_row} <- {amount:g} for {state}; highest Payroll in "
            f"H{target_row} is {highest_payroll:g}"
        )

    for state, rows in state_rows.items():
        target_row, _, expected = plan[state]
        for row in rows:
            actual = sheet.Cells(row, EXPENSE_CONSTANT_COLUMN).Value2
            if row == target_row:
                parsed = parse_nonnegative_number(actual)
                if parsed is None or abs(parsed - expected) > 0.0000001:
                    raise RuntimeError(
                        f"Expense Constant validation failed at W{row}: "
                        f"expected {expected}, found {actual!r}"
                    )
            elif actual not in (None, ""):
                raise RuntimeError(
                    f"Unexpected additional Expense Constant for {state} at "
                    f"W{row}: {actual!r}"
                )


def clear_state_column_filter(sheet):
    """Clear an active filter criterion for absolute worksheet Column F."""
    if not bool(sheet.AutoFilterMode):
        log("No worksheet AutoFilter is present; Column F needs no cleanup.")
        return

    auto_filter = sheet.AutoFilter
    filter_range = auto_filter.Range
    first_column = int(filter_range.Column)
    last_column = first_column + int(filter_range.Columns.Count) - 1
    if not first_column <= STATE_COLUMN <= last_column:
        log("The worksheet AutoFilter does not include Column F.")
        return

    field = STATE_COLUMN - first_column + 1
    try:
        filter_is_on = bool(auto_filter.Filters.Item(field).On)
    except Exception as error:
        raise RuntimeError(
            "Could not inspect the Column F AutoFilter criterion."
        ) from error

    if not filter_is_on:
        log("Column F has no active filter criterion.")
        return

    filter_range.AutoFilter(Field=field)
    if bool(auto_filter.Filters.Item(field).On):
        raise RuntimeError("Column F filter criterion could not be removed.")
    log("Column F filter criterion removed after Expense Constant updates.")


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
    log(f"Reading Expense Constants from: {QUOTE_PDF}")
    expense_constants = extract_expense_constants(QUOTE_PDF)

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
        log(
            f"Writing {len(expense_constants)} Expense Constant value(s) "
            "to Column W."
        )
        write_and_validate_expense_constants(
            sheet, last_row, expense_constants
        )
        clear_state_column_filter(sheet)

        workbook.Save()
        log("E-Mod and Expense Constant values validated; workbook saved.")
        log("SUCCESS: PDF values were filled in Master Comparision.")

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
