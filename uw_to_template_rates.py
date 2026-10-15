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
QUOTE_FILE = BASE_DIR / "Quote" / "Q-279399_Lotus_HR_Inc_Master_v2.xlsx"
TEMPLATE_FILE = (
    BASE_DIR
    / "template"
    / "Lotus HR Inc WC 26-27 Net Rate Comparision (KeyRisk)_Internal.xlsx"
)

SOURCE_SHEET = "Master"
RATES_SHEET = "Lotus Merged Payroll"
DESTINATION_SHEET = "Master Comparision"
EXPIRING_SHEET = "Expiring Dump (KeyRisk)"
QUOTE_MASTER_SHEET = "KeyRisk_Renewal_Quote_Master"
QUOTE_PIVOT_SHEET = "KeyRisk_Renewal_Pivot"

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


def count_nonblank_values(value):
    """Count nonblank cells in a scalar or two-dimensional COM value."""
    return sum(
        1
        for row in as_rows(value)
        for cell in row
        if cell not in (None, "")
    )


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


def count_blank_cells(ws, column, start_row, last_row):
    """Count blank cells in a destination column over the imported row range."""
    if last_row < start_row:
        return 0
    values = as_rows(
        ws.Range(
            ws.Cells(start_row, column),
            ws.Cells(last_row, column),
        ).Value2
    )
    return sum(1 for row in values if row[0] in (None, ""))


def find_quote_source_sheet(workbook):
    """Choose the Quote sheet with the most Classcode rows in Column F."""
    selected = None
    highest_count = -1
    for index in range(1, workbook.Worksheets.Count + 1):
        sheet = workbook.Worksheets(index)
        last_row = sheet.Cells(sheet.Rows.Count, 6).End(XL_UP).Row
        if last_row < 2:
            continue
        values = as_rows(sheet.Range(sheet.Cells(2, 6), sheet.Cells(last_row, 6)).Value2)
        classcode_count = sum(
            1 for row in values if "classcode" in normalize(row[0])
        )
        if classcode_count > highest_count:
            selected = sheet
            highest_count = classcode_count

    if selected is None:
        raise ValueError("No populated worksheet was found in the Quote workbook.")
    log(f"Using Quote source sheet: {selected.Name}")
    return selected


def copy_quote_block_and_refresh_pivot(quote_wb, template_wb):
    """Copy the contiguous Quote block from B2 to J2 and refresh pivots."""
    source_ws = find_quote_source_sheet(quote_wb)
    anchor = source_ws.Cells(2, 2)  # B2
    if anchor.Value in (None, ""):
        raise ValueError(f"Quote source cell {source_ws.Name}!B2 is blank.")

    source_region = anchor.CurrentRegion
    last_row = source_region.Row + source_region.Rows.Count - 1
    last_col = source_region.Column + source_region.Columns.Count - 1
    if last_row < 2 or last_col < 2:
        raise ValueError("No Quote data block was found beginning at B2.")

    source_range = source_ws.Range(
        source_ws.Cells(2, 2),
        source_ws.Cells(last_row, last_col),
    )
    row_count = source_range.Rows.Count
    column_count = source_range.Columns.Count

    try:
        master_ws = template_wb.Worksheets(QUOTE_MASTER_SHEET)
        pivot_ws = template_wb.Worksheets(QUOTE_PIVOT_SHEET)
    except Exception as exc:
        raise ValueError(
            f"Required template sheet missing: {QUOTE_MASTER_SHEET} "
            f"or {QUOTE_PIVOT_SHEET}"
        ) from exc

    target_range = master_ws.Range(
        master_ws.Cells(2, 10),  # J2
        master_ws.Cells(2 + row_count - 1, 10 + column_count - 1),
    )
    source_values = source_range.Value
    target_range.Value = source_values

    # Excel COM may convert equivalent values between types (for example,
    # empty string to None or date representations), so direct tuple equality
    # is too strict. Verify that the same number of populated cells arrived.
    source_nonblank = count_nonblank_values(source_values)
    target_nonblank = count_nonblank_values(target_range.Value)
    log(
        f"Quote copy validation: source nonblank={source_nonblank}, "
        f"destination nonblank={target_nonblank}"
    )
    if target_nonblank != source_nonblank:
        raise RuntimeError(
            "Quote B2 to Quote Master J2 copy validation failed: "
            f"source has {source_nonblank} populated cells, destination has "
            f"{target_nonblank}."
        )

    log(
        f"Copied {source_ws.Name}!B2:{source_ws.Cells(last_row, last_col).Address(False, False)} "
        f"to {QUOTE_MASTER_SHEET}!J2 "
        f"({row_count} rows x {column_count} columns)."
    )

    pivot_tables = pivot_ws.PivotTables()
    pivot_count = pivot_tables.Count
    if pivot_count < 1:
        raise ValueError(f"No pivot table found on {QUOTE_PIVOT_SHEET}.")
    log(
        f"Quote copy validation passed. Refreshing {pivot_count} pivot "
        f"table(s) on {QUOTE_PIVOT_SHEET}..."
    )

    for index in range(1, pivot_count + 1):
        pivot = pivot_tables.Item(index)
        pivot.PivotCache().Refresh()
        result = pivot.RefreshTable()
        if result is False:
            raise RuntimeError(f"Pivot refresh failed: {pivot.Name}")
        log(f"Refreshed pivot table: {pivot.Name}")

    return row_count, column_count, pivot_count


def main():
    for path, label in (
        (SOURCE_FILE, "UW workbook"),
        (RATES_FILE, "Rates Payroll workbook"),
        (QUOTE_FILE, "Quote workbook"),
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
    quote_wb = None
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
        # E=0, H=3, I=4, L=7, N=9.
        rates_data = as_rows(
            rates_ws.Range(
                rates_ws.Cells(2, 5),
                rates_ws.Cells(rates_last_row, 16)
            ).Value2
        )

        p_values = []
        lookup_b_to_l = {}
        lookup_f_to_n = {}
        kept_count = 0
        skipped_count = 0

        for row in rates_data:
            status = cell_text(row[0])   # Column E
            value_h = row[3]             # Column H
            value_i = row[4]             # Column I
            value_l = row[7]             # Column L
            value_n = row[9]             # Column N

            # Keep only records whose Column E begins with KEYPEO.
            if not status.upper().startswith("KEYPEO"):
                p_values.append((None,))
                skipped_count += 1
                continue

            # First lookup key: concatenate H + I without a separator and
            # write it to Column P. Expiring Dump B uses this key to return L.
            key_display = cell_text(value_h) + cell_text(value_i)
            p_values.append((key_display or None,))
            kept_count += 1

            key_b = make_lookup_key(key_display)
            if key_b and key_b not in lookup_b_to_l:
                lookup_b_to_l[key_b] = value_l

            # Second, independent lookup: Expiring Dump F matches Rates H and
            # returns Rates N into Expiring Dump K.
            key_f = make_lookup_key(value_h)
            if key_f and key_f not in lookup_f_to_n:
                lookup_f_to_n[key_f] = value_n

        log("Writing KEYPEO concatenation keys into Rates Payroll Column P...")
        rates_ws.Range(
            rates_ws.Cells(2, 16),
            rates_ws.Cells(rates_last_row, 16)
        ).ClearContents()
        rates_ws.Range(
            rates_ws.Cells(2, 16),
            rates_ws.Cells(rates_last_row, 16)
        ).Value = tuple(p_values)

        log(
            "Matching Expiring Dump B -> Rates P (return L), and "
            "Expiring Dump F -> Rates H (return N)..."
        )
        last_row_b = expiring_ws.Cells(
            expiring_ws.Rows.Count, 2
        ).End(XL_UP).Row
        last_row_f = expiring_ws.Cells(
            expiring_ws.Rows.Count, 6
        ).End(XL_UP).Row
        template_last_row = max(last_row_b, last_row_f)

        match_count_b = 0
        missing_count_b = 0
        match_count_f = 0
        missing_count_f = 0

        if template_last_row >= DEST_START_ROW:
            # Read B:F so each destination uses its own independent key:
            # B (index 0) -> Rates P -> Rates L -> destination H
            # F (index 4) -> Rates H -> Rates N -> destination K
            template_keys = as_rows(
                expiring_ws.Range(
                    expiring_ws.Cells(DEST_START_ROW, 2),
                    expiring_ws.Cells(template_last_row, 6)
                ).Value2
            )

            output_h = []
            output_k = []

            for row in template_keys:
                key_b = make_lookup_key(row[0])
                if key_b and key_b in lookup_b_to_l:
                    output_h.append((lookup_b_to_l[key_b],))
                    match_count_b += 1
                else:
                    output_h.append((None,))
                    if key_b:
                        missing_count_b += 1

                key_f = make_lookup_key(row[4])
                if key_f and key_f in lookup_f_to_n:
                    output_k.append((lookup_f_to_n[key_f],))
                    match_count_f += 1
                else:
                    output_k.append((None,))
                    if key_f:
                        missing_count_f += 1

            # Clear old outputs to avoid stale values for unmatched records.
            expiring_ws.Range(
                expiring_ws.Cells(DEST_START_ROW, 8),
                expiring_ws.Cells(template_last_row, 8)
            ).ClearContents()
            expiring_ws.Range(
                expiring_ws.Cells(DEST_START_ROW, 11),
                expiring_ws.Cells(template_last_row, 11)
            ).ClearContents()

            # Rates Payroll L -> Expiring Dump H (matched B against P).
            expiring_ws.Range(
                expiring_ws.Cells(DEST_START_ROW, 8),
                expiring_ws.Cells(template_last_row, 8)
            ).Value = tuple(output_h)

            # Rates Payroll N -> Expiring Dump K (matched F against H).
            expiring_ws.Range(
                expiring_ws.Cells(DEST_START_ROW, 11),
                expiring_ws.Cells(template_last_row, 11)
            ).Value = tuple(output_k)
        else:
            log("No lookup keys found in Expiring Dump B or F from row 3.")

        # Stage 1 indication: count rates still blank after Rates Payroll
        # lookups. H is Manual Rate; K is All-In Fund Rate.
        # Count blanks only across this run's imported UW records (rows 3 to
        # last_data_row). Formula-filled Column B can extend much farther down
        # and must not inflate the H/K blank summary.
        blank_h_count = count_blank_cells(
            expiring_ws, 8, DEST_START_ROW, last_data_row
        )
        blank_k_count = count_blank_cells(
            expiring_ws, 11, DEST_START_ROW, last_data_row
        )
        log(
            f"BLANK SUMMARY: Expiring Dump H (Manual Rate): "
            f"{blank_h_count}"
        )
        log(
            f"BLANK SUMMARY: Expiring Dump K (All-In Fund Rate): "
            f"{blank_k_count}"
        )
        if blank_h_count:
            log(
                "NEXT STEP REQUIRED: run fill_manual_rates_from_pdf.py "
                "after downloading the policy PDF."
            )
        else:
            log("No PDF Manual Rate fallback is required for Column H.")

        # FINAL STEP: preserve all existing UW/Rates logic, then copy the Quote
        # data block from B2 into Quote Master J2 and refresh renewal pivots.
        log("Opening Quote workbook read-only...")
        quote_wb = excel.Workbooks.Open(
            str(QUOTE_FILE), UpdateLinks=0, ReadOnly=True
        )
        quote_rows, quote_columns, pivot_count = (
            copy_quote_block_and_refresh_pivot(quote_wb, dest_wb)
        )

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
        log(
            "Unique B->P lookup keys (return L): "
            f"{len(lookup_b_to_l)}"
        )
        log(
            "Unique F->H lookup keys (return N): "
            f"{len(lookup_f_to_n)}"
        )
        log(
            f"Expiring Dump lookup rows processed: "
            f"{max(0, template_last_row - 2)}"
        )
        log(f"Current imported rows used for blank summary: {count}")
        log(f"B->P matches found for destination H: {match_count_b}")
        log(f"Nonblank B keys with no P match: {missing_count_b}")
        log(f"F->H matches found for destination K: {match_count_f}")
        log(f"Nonblank F keys with no H match: {missing_count_f}")
        log(f"Blank H Manual Rates after Rates lookup: {blank_h_count}")
        log(f"Blank K All-In Fund Rates after Rates lookup: {blank_k_count}")
        log(
            f"Quote block copied: {quote_rows} rows x {quote_columns} columns"
        )
        log(f"KeyRisk renewal pivot tables refreshed: {pivot_count}")
        log("Expiring B -> Rates P -> Rates L -> Expiring H.")
        log("Expiring F -> Rates H -> Rates N -> Expiring K.")
        log(
            "STAGE 1 SUCCESS: existing UW/Rates logic completed; Quote data "
            "copied and renewal pivot refreshed."
        )

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

        if quote_wb is not None:
            try:
                quote_wb.Close(SaveChanges=False)
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
