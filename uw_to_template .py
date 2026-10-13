from pathlib import Path
from datetime import datetime
import shutil
import sys

try:
    import win32com.client as win32
except ImportError:
    print("ERROR: Install pywin32 using: pip install pywin32", flush=True)
    raise


BASE_DIR = Path(r"C:\Users\Samarth\Documents\Key Risk automation")

SOURCE_FILE = BASE_DIR / "UW" / "Lotus HR Inc Template 26-27 UW V2.xlsx"
RATES_FILE = BASE_DIR / "Expiry" / "Rates Payroll.xlsx"
TEMPLATE_FILE = (
    BASE_DIR
    / "template"
    / "Lotus HR Inc WC 26-27 Net Rate Comparision (KeyRisk)_Internal.xlsx"
)

SOURCE_SHEET = "Master"
RATES_SHEET = "Lotus Merged Payroll"
DESTINATION_SHEET = "Master Comparision"
EXPIRING_SHEET = "Expiring Dump (KeyRisk)"

HEADERS = [
    "Client Id",
    "Name",
    "Tax PayerId",
    "State",
    "Class Code",
    "Payroll",
]

DEST_HEADER_ROW = 2
DEST_START_ROW = 3
DEST_START_COL = 3
EXISTING_FORMULA_ROW = 83

XL_UP = -4162
XL_TO_LEFT = -4159
XL_SHIFT_DOWN = -4121
XL_CALCULATION_MANUAL = -4135


def log(message):
    print(f"[{datetime.now():%H:%M:%S}] {message}", flush=True)


def normalize(value):
    if value is None:
        return ""
    return "".join(
        ch.lower() for ch in str(value).strip() if ch.isalnum()
    )


def cell_text(value):
    """Convert an Excel cell value to a trimmed string for lookup keys."""
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def make_lookup_key(value):
    """Normalize lookup keys for case-insensitive matching."""
    return cell_text(value).casefold()


def as_rows(value):
    """Normalize Excel COM range output to a list of row tuples."""
    if value is None:
        return []
    if not isinstance(value, tuple):
        return [(value,)]
    if value and not isinstance(value[0], tuple):
        return [value]
    return list(value)


def find_header_columns(ws, headers, header_row=1):
    last_col = ws.Cells(header_row, ws.Columns.Count).End(XL_TO_LEFT).Column
    found = {}

    for col in range(1, last_col + 1):
        key = normalize(ws.Cells(header_row, col).Value)
        if key and key not in found:
            found[key] = col

    missing = [h for h in headers if normalize(h) not in found]
    if missing:
        raise ValueError(
            f"Missing headers in '{ws.Name}' row {header_row}: {missing}"
        )

    return [found[normalize(h)] for h in headers]


def main():
    for path, label in (
        (SOURCE_FILE, "UW workbook"),
        (RATES_FILE, "Rates Payroll workbook"),
        (TEMPLATE_FILE, "Template workbook"),
    ):
        if not path.exists():
            raise FileNotFoundError(f"{label} not found: {path}")

    backup = TEMPLATE_FILE.with_name(
        f"{TEMPLATE_FILE.stem}_backup_{datetime.now():%Y%m%d_%H%M%S}"
        f"{TEMPLATE_FILE.suffix}"
    )

    log("Creating template backup...")
    shutil.copy2(TEMPLATE_FILE, backup)
    log(f"Backup created: {backup.name}")

    excel = None
    source_wb = None
    rates_wb = None
    dest_wb = None
    old_calculation = None
    old_screen = None
    old_events = None

    try:
        log("Starting Excel...")
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False

        try:
            old_calculation = excel.Calculation
            old_screen = excel.ScreenUpdating
            old_events = excel.EnableEvents
            excel.Calculation = XL_CALCULATION_MANUAL
            excel.ScreenUpdating = False
            excel.EnableEvents = False
        except Exception:
            pass

        # STEP 1: UW -> Master Comparision
        log("Opening UW workbook...")
        source_wb = excel.Workbooks.Open(
            str(SOURCE_FILE), UpdateLinks=0, ReadOnly=True
        )

        try:
            source_ws = source_wb.Worksheets(SOURCE_SHEET)
        except Exception as exc:
            raise ValueError(
                f"Source sheet '{SOURCE_SHEET}' was not found."
            ) from exc

        log("Finding source columns...")
        source_columns = find_header_columns(source_ws, HEADERS, header_row=1)

        last_source_row = source_ws.Cells(
            source_ws.Rows.Count, 1
        ).End(XL_UP).Row

        if last_source_row < 2:
            raise ValueError("No UW data found below row 1.")

        source_values = source_ws.Range(
            source_ws.Cells(2, 1),
            source_ws.Cells(last_source_row, max(source_columns))
        ).Value

        records = []
        for row in as_rows(source_values):
            record = tuple(row[col - 1] for col in source_columns)
            if any(value not in (None, "") for value in record):
                records.append(record)

        if not records:
            raise ValueError("No nonblank UW records found.")

        count = len(records)
        last_data_row = DEST_START_ROW + count - 1
        formula_target_row = last_data_row + 1

        log(f"UW records found: {count}")
        log(f"Master Comparision destination: C3:H{last_data_row}")

        log("Opening template workbook...")
        dest_wb = excel.Workbooks.Open(
            str(TEMPLATE_FILE), UpdateLinks=0, ReadOnly=False
        )

        try:
            master_ws = dest_wb.Worksheets(DESTINATION_SHEET)
            expiring_ws = dest_wb.Worksheets(EXPIRING_SHEET)
        except Exception as exc:
            available = [
                dest_wb.Worksheets(i).Name
                for i in range(1, dest_wb.Worksheets.Count + 1)
            ]
            raise ValueError(
                f"Could not find a required destination sheet. "
                f"Available sheets: {available}"
            ) from exc

        log("Writing Master Comparision headers...")
        master_ws.Range("C2:H2").Value = (tuple(HEADERS),)

        # Shift formula row down if imported data would overlap it.
        if formula_target_row > EXISTING_FORMULA_ROW:
            rows_to_insert = formula_target_row - EXISTING_FORMULA_ROW
            log(
                f"Moving formula row down by inserting {rows_to_insert} row(s)..."
            )
            master_ws.Rows(
                f"{EXISTING_FORMULA_ROW}:{formula_target_row - 1}"
            ).Insert(Shift=XL_SHIFT_DOWN)

        log("Writing UW records into Master Comparision...")
        master_ws.Range(
            master_ws.Cells(DEST_START_ROW, DEST_START_COL),
            master_ws.Cells(
                last_data_row, DEST_START_COL + len(HEADERS) - 1
            )
        ).Value = tuple(records)

        # STEP 2: Master Comparision D:G -> Expiring Dump D:G
        log("Preparing Expiring Dump transfer...")
        expiring_records = master_ws.Range(
            f"D{DEST_START_ROW}:G{last_data_row}"
        ).Value

        log("Clearing previous values in Expiring Dump D:G...")
        expiring_last_row = expiring_ws.Cells(
            expiring_ws.Rows.Count, 4
        ).End(XL_UP).Row

        if expiring_last_row >= DEST_START_ROW:
            expiring_ws.Range(
                f"D{DEST_START_ROW}:G{expiring_last_row}"
            ).ClearContents()

        log(
            f"Writing {count} records to Expiring Dump "
            f"D3:G{last_data_row}..."
        )
        expiring_ws.Range(
            f"D{DEST_START_ROW}:G{last_data_row}"
        ).Value = expiring_records

        # STEP 3: Prepare Rates Payroll Column P and look up into Expiring Dump.
        log("Opening Rates Payroll workbook...")
        rates_wb = excel.Workbooks.Open(
            str(RATES_FILE), UpdateLinks=0, ReadOnly=False
        )

        try:
            rates_ws = rates_wb.Worksheets(RATES_SHEET)
        except Exception as exc:
            raise ValueError(
                f"Rates Payroll sheet '{RATES_SHEET}' was not found."
            ) from exc

        log("Reading Rates Payroll records...")
        rates_last_row = rates_ws.Cells(
            rates_ws.Rows.Count, 5
        ).End(XL_UP).Row  # Column E

        if rates_last_row < 2:
            raise ValueError(
                f"No data found in '{RATES_SHEET}' Column E below row 1."
            )

        # Read E:P. Indexes relative to E:P:
        # E=0, H=3, I=4, L=7, M=8.
        rates_data = as_rows(
            rates_ws.Range(
                rates_ws.Cells(2, 5),
                rates_ws.Cells(rates_last_row, 16)
            ).Value2
        )

        p_values = []
        lookup = {}
        kept_count = 0
        skipped_count = 0

        for row in rates_data:
            status = cell_text(row[0])  # Column E
            value_h = row[3]            # Column H
            value_i = row[4]            # Column I
            value_k = row[7]            # Column L
            value_l = row[9]            # Column N

            # Keep only records whose Column E begins with KEYPEO.
            if not status.upper().startswith("KEYPEO"):
                p_values.append((None,))
                skipped_count += 1
                continue

            # Concatenate H + I without a separator and write to Column P.
            key_display = cell_text(value_h) + cell_text(value_i)
            p_values.append((key_display or None,))
            kept_count += 1

            key = make_lookup_key(key_display)
            # Match XLOOKUP behavior: use the first occurrence of a key.
            if key and key not in lookup:
                lookup[key] = (value_k, value_l)

        log("Writing KEYPEO concatenation keys into Rates Payroll Column P...")
        rates_ws.Range(
            rates_ws.Cells(2, 16),
            rates_ws.Cells(rates_last_row, 16)
        ).ClearContents()
        rates_ws.Range(
            rates_ws.Cells(2, 16),
            rates_ws.Cells(rates_last_row, 16)
        ).Value = tuple(p_values)

        log("Matching Expiring Dump Column B against Rates Payroll Column P...")
        template_last_row = expiring_ws.Cells(
            expiring_ws.Rows.Count, 2
        ).End(XL_UP).Row  # Column B

        match_count = 0
        missing_count = 0

        if template_last_row >= DEST_START_ROW:
            template_keys = as_rows(
                expiring_ws.Range(
                    expiring_ws.Cells(DEST_START_ROW, 2),
                    expiring_ws.Cells(template_last_row, 2)
                ).Value2
            )

            output_h = []
            output_k = []

            for row in template_keys:
                key = make_lookup_key(row[0])
                if key and key in lookup:
                    result_k, result_l = lookup[key]
                    output_h.append((result_k,))
                    output_k.append((result_l,))
                    match_count += 1
                else:
                    output_h.append((None,))
                    output_k.append((None,))
                    if key:
                        missing_count += 1

            # Clear old outputs to avoid stale values for unmatched records.
            expiring_ws.Range(
                expiring_ws.Cells(DEST_START_ROW, 8),
                expiring_ws.Cells(template_last_row, 8)
            ).ClearContents()
            expiring_ws.Range(
                expiring_ws.Cells(DEST_START_ROW, 11),
                expiring_ws.Cells(template_last_row, 11)
            ).ClearContents()

            # Rates Payroll L -> Expiring Dump H.
            expiring_ws.Range(
                expiring_ws.Cells(DEST_START_ROW, 8),
                expiring_ws.Cells(template_last_row, 8)
            ).Value = tuple(output_h)

            # Rates Payroll M -> Expiring Dump K.
            expiring_ws.Range(
                expiring_ws.Cells(DEST_START_ROW, 11),
                expiring_ws.Cells(template_last_row, 11)
            ).Value = tuple(output_k)
        else:
            log("No lookup keys found in Expiring Dump Column B from row 3.")

        # Save both edited workbooks.
        log("Saving Rates Payroll workbook...")
        rates_wb.Save()

        log("Saving template workbook...")
        dest_wb.Save()

        # Validate the D:G transfer from Step 2.
        first_master = master_ws.Range("D3:G3").Value
        first_expiring = expiring_ws.Range("D3:G3").Value
        if first_master != first_expiring:
            raise RuntimeError(
                "Validation failed: first row in Expiring Dump "
                "does not match Master Comparision."
            )

        log(f"Validation passed: {count} rows copied to Master Comparision.")
        log(f"Validation passed: {count} rows copied to Expiring Dump D:G.")
        log(f"Rates Payroll rows checked: {len(rates_data)}")
        log(f"Rows kept (Column E begins with KEYPEO): {kept_count}")
        log(f"Rows skipped (Column E does not begin with KEYPEO): {skipped_count}")
        log(f"Unique nonblank lookup keys: {len(lookup)}")
        log(f"Expiring Dump rows checked: {max(0, template_last_row - 2)}")
        log(f"Lookup matches found: {match_count}")
        log(f"Nonblank lookup keys with no match: {missing_count}")
        log("Rates Payroll L -> Expiring Dump H.")
        log("Rates Payroll N -> Expiring Dump K.")
        log("SUCCESS: All three transfer/lookup steps completed.")

    finally:
        if source_wb is not None:
            try:
                source_wb.Close(SaveChanges=False)
            except Exception:
                pass

        if rates_wb is not None:
            try:
                rates_wb.Close(SaveChanges=True)
            except Exception:
                pass

        if dest_wb is not None:
            try:
                dest_wb.Close(SaveChanges=True)
            except Exception:
                pass

        if excel is not None:
            try:
                if old_calculation is not None:
                    excel.Calculation = old_calculation
                if old_screen is not None:
                    excel.ScreenUpdating = old_screen
                if old_events is not None:
                    excel.EnableEvents = old_events
            except Exception:
                pass
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
