"""
Compare UW payroll against Quote payroll before running the main automation.

Rules
-----
UW workbook / Master sheet:
- Data rows: row 2 through the last nonblank row in Column A.
- State: Column P.
- Class Code: Column Q.
- Payroll: Column R.

Quote workbook:
- Include only rows where Column F contains "Classcode" (spacing and case are
  ignored, so "Class Code" also qualifies).
- State: Column G.
- Class Code: Column H.
- Payroll: Column J.

If the totals match after standard whole-dollar rounding, the script prints MATCH
and does not modify a workbook. If they differ, it creates/replaces a "Payroll Comparison" sheet in
the Quote workbook showing payroll differences by State and Class Code. A
backup of the Quote workbook is created before that workbook is modified.

Run on Windows with Microsoft Excel installed:
    python compare_payroll_totals.py
"""

from collections import defaultdict
from datetime import datetime
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from pathlib import Path
import re
import shutil
import sys

try:
    import win32com.client as win32
except ImportError:
    print("ERROR: Install pywin32 using: pip install pywin32", flush=True)
    raise


BASE_DIR = Path(r"C:\Users\Samarth\Documents\Key Risk automation")
UW_FILE = BASE_DIR / "UW" / "Lotus HR Inc Template 26-27 UW V2.xlsx"
QUOTE_FILE = BASE_DIR / "Quote" / "Q-279399_Lotus_HR_Inc_Master_v2.xlsx"

UW_SHEET = "Master"
REPORT_SHEET = "Payroll Comparison"

DATA_START_ROW = 2
WHOLE_DOLLAR = Decimal("1")

XL_UP = -4162
XL_TO_LEFT = -4159
XL_SHEET_VISIBLE = -1


def log(message):
    print(f"[{datetime.now():%H:%M:%S}] {message}", flush=True)


def as_rows(value):
    """Normalize Excel COM range output to a list of row tuples."""
    if value is None:
        return []
    if not isinstance(value, tuple):
        return [(value,)]
    if value and not isinstance(value[0], tuple):
        return [value]
    return list(value)


def cell_text(value):
    """Convert an Excel value to a stable trimmed string."""
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def normalize_marker(value):
    """Normalize text for matching 'Classcode' regardless of spaces/case."""
    return "".join(character.casefold() for character in cell_text(value) if character.isalnum())


def normalize_state(value):
    """Normalize state values to uppercase codes/text."""
    return cell_text(value).upper()


def normalize_class_code(value):
    """Normalize numeric class codes and preserve nonnumeric codes safely."""
    text = cell_text(value).upper()
    if not text:
        return ""

    # Excel often returns a numeric class code as 8742.0.
    if re.fullmatch(r"\d+\.0+", text):
        text = text.split(".", 1)[0]

    compact = re.sub(r"\s+", "", text)
    if compact.isdigit():
        return compact.lstrip("0") or "0"
    return compact


def parse_amount(value, workbook_label, excel_row):
    """Convert numeric/currency Excel content to Decimal payroll."""
    if value in (None, ""):
        return Decimal("0")
    if isinstance(value, (int, float)):
        return Decimal(str(value))

    text = cell_text(value)
    negative = text.startswith("(") and text.endswith(")")
    cleaned = text.replace("$", "").replace(",", "").strip("() ")
    try:
        amount = Decimal(cleaned)
    except InvalidOperation as error:
        raise ValueError(
            f"Invalid payroll value in {workbook_label} row {excel_row}: "
            f"{value!r}"
        ) from error
    return -amount if negative else amount


def money(value):
    """Round a payroll total to the nearest whole dollar, half up."""
    return value.quantize(WHOLE_DOLLAR, rounding=ROUND_HALF_UP)


def last_used_row(ws, columns):
    """Return the maximum last nonblank row among one or more column numbers."""
    return max(
        ws.Cells(ws.Rows.Count, column).End(XL_UP).Row
        for column in columns
    )


def read_uw_payroll(ws):
    """Return UW total and payroll grouped by (State P, Class Code Q)."""
    last_row = ws.Cells(ws.Rows.Count, 1).End(XL_UP).Row  # Column A controls extent.
    if last_row < DATA_START_ROW:
        raise ValueError(f"No UW records found in {UW_SHEET!r} below row 1.")

    # Read P:R: State, Class Code, Payroll.
    rows = as_rows(
        ws.Range(
            ws.Cells(DATA_START_ROW, 16),
            ws.Cells(last_row, 18),
        ).Value2
    )

    grouped = defaultdict(Decimal)
    total = Decimal("0")
    for offset, row in enumerate(rows):
        excel_row = DATA_START_ROW + offset
        state = normalize_state(row[0])
        class_code = normalize_class_code(row[1])
        payroll = parse_amount(row[2], "UW", excel_row)
        total += payroll
        grouped[(state, class_code)] += payroll

    return money(total), grouped, last_row - DATA_START_ROW + 1


def find_quote_data_sheet(workbook):
    """Find the Quote sheet with the most qualifying Column F records."""
    best_sheet = None
    best_count = 0

    for index in range(1, workbook.Worksheets.Count + 1):
        worksheet = workbook.Worksheets(index)
        if worksheet.Name.casefold() == REPORT_SHEET.casefold():
            continue
        try:
            last_row = worksheet.Cells(
                worksheet.Rows.Count, 6
            ).End(XL_UP).Row
            if last_row < DATA_START_ROW:
                continue
            values = as_rows(
                worksheet.Range(
                    worksheet.Cells(DATA_START_ROW, 6),
                    worksheet.Cells(last_row, 6),
                ).Value2
            )
            qualifying = sum(
                1
                for row in values
                if "classcode" in normalize_marker(row[0])
            )
            if qualifying > best_count:
                best_sheet = worksheet
                best_count = qualifying
        except Exception:
            continue

    if best_sheet is None:
        raise ValueError(
            "No Quote worksheet contains qualifying 'Classcode' values "
            "in Column F."
        )
    log(
        f"Using Quote worksheet {best_sheet.Name!r} "
        f"({best_count} qualifying rows)."
    )
    return best_sheet


def read_quote_payroll(ws):
    """Return filtered Quote total and payroll grouped by (State G, Class H)."""
    last_row = last_used_row(ws, (6, 7, 8, 10))  # F, G, H, J
    if last_row < DATA_START_ROW:
        raise ValueError("No Quote records found below row 1.")

    # Read F:J: marker, State, Class Code, unused I, Payroll.
    rows = as_rows(
        ws.Range(
            ws.Cells(DATA_START_ROW, 6),
            ws.Cells(last_row, 10),
        ).Value2
    )

    grouped = defaultdict(Decimal)
    total = Decimal("0")
    included_count = 0
    for offset, row in enumerate(rows):
        if "classcode" not in normalize_marker(row[0]):
            continue

        excel_row = DATA_START_ROW + offset
        state = normalize_state(row[1])
        class_code = normalize_class_code(row[2])
        payroll = parse_amount(row[4], "Quote", excel_row)
        total += payroll
        grouped[(state, class_code)] += payroll
        included_count += 1

    if included_count == 0:
        raise ValueError(
            "No Quote rows qualified: Column F did not contain 'Classcode'."
        )

    return money(total), grouped, included_count


def worksheet_by_name(workbook, name):
    """Return a worksheet by name, or None when absent."""
    for index in range(1, workbook.Worksheets.Count + 1):
        worksheet = workbook.Worksheets(index)
        if worksheet.Name.casefold() == name.casefold():
            return worksheet
    return None


def difference_rows(uw_grouped, quote_grouped):
    """Return only state/class combinations whose payroll differs."""
    rows = []
    all_keys = sorted(
        set(uw_grouped) | set(quote_grouped),
        key=lambda key: (key[0], key[1]),
    )
    for state, class_code in all_keys:
        uw_payroll = money(uw_grouped.get((state, class_code), Decimal("0")))
        quote_payroll = money(quote_grouped.get((state, class_code), Decimal("0")))
        difference = money(uw_payroll - quote_payroll)
        if difference == Decimal("0"):
            continue
        rows.append(
            (
                state or "<BLANK>",
                class_code or "<BLANK>",
                float(uw_payroll),
                float(quote_payroll),
                float(difference),
                "Only in UW" if quote_payroll == 0 else (
                    "Only in Quote" if uw_payroll == 0 else "Payroll mismatch"
                ),
            )
        )
    return rows


def write_report(workbook, uw_total, quote_total, rows):
    """Create or replace the formatted Payroll Comparison worksheet."""
    existing = worksheet_by_name(workbook, REPORT_SHEET)
    if existing is not None:
        existing.Delete()

    report = workbook.Worksheets.Add(After=workbook.Worksheets(workbook.Worksheets.Count))
    report.Name = REPORT_SHEET
    report.Visible = XL_SHEET_VISIBLE

    report.Range("A1:F1").Merge()
    report.Range("A1").Value = "UW vs Quote Payroll Comparison"
    report.Range("A1").Font.Bold = True
    report.Range("A1").Font.Size = 14

    report.Range("A3").Value = "Rounded UW Total (Master!R)"
    report.Range("B3").Value = float(uw_total)
    report.Range("A4").Value = "Rounded Quote Total (F contains Classcode; sum J)"
    report.Range("B4").Value = float(quote_total)
    report.Range("A5").Value = "Rounded Difference (UW - Quote)"
    report.Range("B5").Value = float(money(uw_total - quote_total))
    report.Range("A6").Value = "Generated"
    report.Range("B6").Value = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    headers = (
        "State",
        "Class Code",
        "Rounded UW Payroll",
        "Rounded Quote Payroll",
        "Rounded Difference (UW - Quote)",
        "Status",
    )
    report.Range("A8:F8").Value = (headers,)
    report.Range("A8:F8").Font.Bold = True

    if rows:
        last_report_row = 8 + len(rows)
        report.Range(f"A9:F{last_report_row}").Value = tuple(rows)
        report.Range(f"C9:E{last_report_row}").NumberFormat = "$#,##0;[Red]-$#,##0"
        report.Range(f"A8:F{last_report_row}").Borders.LineStyle = 1
        report.Range(f"A8:F{last_report_row}").AutoFilter()
    else:
        report.Range("A9").Value = "Totals differ, but no grouped difference was found."

    report.Range("B3:B5").NumberFormat = "$#,##0;[Red]-$#,##0"
    report.Columns("A:F").AutoFit()
    report.Activate()
    report.Range("A9").Select()
    workbook.Application.ActiveWindow.FreezePanes = True
    return report


def main():
    for path, label in ((UW_FILE, "UW workbook"), (QUOTE_FILE, "Quote workbook")):
        if not path.exists():
            raise FileNotFoundError(f"{label} not found: {path}")

    excel = None
    uw_workbook = None
    quote_workbook = None

    try:
        log("Starting Excel...")
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False

        log("Opening UW workbook read-only...")
        uw_workbook = excel.Workbooks.Open(
            str(UW_FILE), UpdateLinks=0, ReadOnly=True
        )
        uw_sheet = worksheet_by_name(uw_workbook, UW_SHEET)
        if uw_sheet is None:
            raise ValueError(f"UW sheet not found: {UW_SHEET}")

        log("Opening Quote workbook...")
        quote_workbook = excel.Workbooks.Open(
            str(QUOTE_FILE), UpdateLinks=0, ReadOnly=False
        )
        quote_sheet = find_quote_data_sheet(quote_workbook)

        uw_total, uw_grouped, uw_count = read_uw_payroll(uw_sheet)
        quote_total, quote_grouped, quote_count = read_quote_payroll(quote_sheet)

        log(f"UW rows checked (through last Column A record): {uw_count}")
        log(f"Quote rows included (Column F contains Classcode): {quote_count}")
        log(f"UW Column R total (rounded): {uw_total:,.0f}")
        log(f"Filtered Quote Column J total (rounded): {quote_total:,.0f}")

        if uw_total == quote_total:
            log("MATCH: Rounded UW and Quote payroll totals are equal.")
            return

        total_difference = money(uw_total - quote_total)
        rows = difference_rows(uw_grouped, quote_grouped)
        log(f"MISMATCH: Rounded UW - Quote payroll difference: {total_difference:,.0f}")
        log(f"State/Class combinations with differences: {len(rows)}")

        backup = QUOTE_FILE.with_name(
            f"{QUOTE_FILE.stem}_backup_{datetime.now():%Y%m%d_%H%M%S}"
            f"{QUOTE_FILE.suffix}"
        )
        # The workbook must be closed before file-level backup to avoid copying
        # an in-memory or locked partial representation.
        quote_workbook.Close(SaveChanges=False)
        quote_workbook = None
        shutil.copy2(QUOTE_FILE, backup)
        log(f"Quote backup created: {backup}")

        quote_workbook = excel.Workbooks.Open(
            str(QUOTE_FILE), UpdateLinks=0, ReadOnly=False
        )
        write_report(
            quote_workbook,
            uw_total,
            quote_total,
            rows,
        )
        quote_workbook.Save()
        log(f"Comparison sheet written to Quote workbook: {REPORT_SHEET}")
        log("Review the detailed State/Class payroll differences before continuing.")

    finally:
        if uw_workbook is not None:
            try:
                uw_workbook.Close(SaveChanges=False)
            except Exception:
                pass
        if quote_workbook is not None:
            try:
                quote_workbook.Close(SaveChanges=False)
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
