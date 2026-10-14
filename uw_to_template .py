from pathlib import Path
from datetime import datetime
import re
import shutil
import sys

try:
    import win32com.client as win32
except ImportError:
    print("ERROR: Install pywin32 using: pip install pywin32", flush=True)
    raise


BASE_DIR = Path(r"C:\Users\Samarth\Documents\Key Risk automation")
PDF_DOWNLOAD_DIR = BASE_DIR / "downloads"

# Leave as None to use the newest PDF in PDF_DOWNLOAD_DIR. To force a specific
# policy PDF, assign a full Path here.
POLICY_PDF_FILE = None

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


def normalize_state(value):
    """Normalize a two-letter state code for PDF/Excel matching."""
    state = cell_text(value).upper()
    return state if re.fullmatch(r"[A-Z]{2}", state) else ""


def normalize_class_code(value):
    """Normalize numeric class codes while tolerating Excel float values."""
    text = cell_text(value)
    digits = re.sub(r"\D", "", text)
    if not digits:
        return ""
    return digits.lstrip("0") or "0"


def parse_rate(value):
    """Convert a PDF table rate cell to a non-negative number."""
    if isinstance(value, (int, float)):
        rate = float(value)
        return rate if rate >= 0 else None

    text = cell_text(value).replace(",", "").replace("$", "")
    match = re.search(r"-?\d+(?:\.\d+)?", text)
    if not match:
        return None
    rate = float(match.group(0))
    return rate if rate >= 0 else None


def add_policy_rate(rates, state, class_code, rate):
    """Add the first valid rate for a normalized state/class pair."""
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
            f"WARNING: PDF contains multiple rates for {state_key} "
            f"class {class_key}: {rates[key]} and {parsed_rate}. "
            "Keeping the first rate."
        )


def find_table_headers(table):
    """Find State, Class Code, and Rate columns in a PDF-extracted table."""
    if not table:
        return None

    # Some PDFs split multi-line headers into two or three physical rows.
    for start in range(min(10, len(table))):
        for depth in range(1, 4):
            rows = table[start:start + depth]
            if len(rows) < depth:
                continue
            width = max((len(row or []) for row in rows), default=0)
            combined = []
            for column in range(width):
                parts = []
                for row in rows:
                    if row and column < len(row) and row[column] is not None:
                        parts.append(cell_text(row[column]))
                combined.append(normalize(" ".join(parts)))

            state_column = next(
                (i for i, value in enumerate(combined) if value == "state"),
                None,
            )
            class_column = next(
                (i for i, value in enumerate(combined) if "classcode" in value),
                None,
            )
            rate_column = next(
                (
                    i
                    for i, value in enumerate(combined)
                    if "rateper100" in value
                    or ("rate" in value and "remuneration" in value)
                ),
                None,
            )
            if None not in (state_column, class_column, rate_column):
                return start + depth, state_column, class_column, rate_column
    return None


def extract_rates_from_tables(page, rates):
    """Extract policy rates from recognized PDF tables on one page."""
    for table in page.extract_tables() or []:
        header = find_table_headers(table)
        if header is None:
            continue
        data_start, state_column, class_column, rate_column = header
        required_column = max(state_column, class_column, rate_column)
        for row in table[data_start:]:
            if not row or len(row) <= required_column:
                continue
            add_policy_rate(
                rates,
                row[state_column],
                row[class_column],
                row[rate_column],
            )


def extract_rates_from_text(page, rates):
    """Fallback for PDFs whose visual table has no extractable cell borders."""
    text = page.extract_text(layout=True) or page.extract_text() or ""
    row_start = re.compile(r"^\s*([A-Z]{2})\s+(\d{3,8})\s+(.+?)\s*$")
    # Typical row ending: annual payroll, rate per $100, annual premium.
    row_end = re.compile(
        r"(?:\$?\(?-?[\d,]+(?:\.\d+)?\)?)\s+"
        r"(\d+(?:\.\d+)?)\s+"
        r"\$?\(?-?[\d,]+(?:\.\d+)?\)?\s*$"
    )

    for line in text.splitlines():
        start_match = row_start.match(line)
        if not start_match:
            continue
        state, class_code, remainder = start_match.groups()
        rate_match = row_end.search(remainder)
        if rate_match:
            add_policy_rate(rates, state, class_code, rate_match.group(1))


def select_policy_pdf():
    """Return the configured PDF or the newest downloaded PDF."""
    if POLICY_PDF_FILE is not None:
        pdf_path = Path(POLICY_PDF_FILE)
        if not pdf_path.exists():
            raise FileNotFoundError(f"Configured policy PDF not found: {pdf_path}")
        return pdf_path

    if not PDF_DOWNLOAD_DIR.exists():
        raise FileNotFoundError(
            f"Policy download folder not found: {PDF_DOWNLOAD_DIR}"
        )
    pdf_files = [
        path
        for path in PDF_DOWNLOAD_DIR.iterdir()
        if path.is_file() and path.suffix.casefold() == ".pdf"
    ]
    if not pdf_files:
        raise FileNotFoundError(
            f"No PDF files found in: {PDF_DOWNLOAD_DIR}"
        )
    return max(pdf_files, key=lambda path: path.stat().st_mtime)


def extract_policy_rates(pdf_path):
    """Build a {(state, class_code): rate} map from a policy PDF."""
    try:
        import pdfplumber
    except ImportError as exc:
        raise RuntimeError(
            "PDF fallback requires pdfplumber. Install dependencies with: "
            "python -m pip install -r requirements.txt"
        ) from exc

    rates = {}
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            extract_rates_from_tables(page, rates)
            extract_rates_from_text(page, rates)
    return rates


def fill_blank_expiring_rates_from_pdf(expiring_ws, last_row):
    """Fill blank H using F=State, G=Class Code and the downloaded PDF."""
    if last_row < DEST_START_ROW:
        return 0, 0

    rows = as_rows(
        expiring_ws.Range(
            expiring_ws.Cells(DEST_START_ROW, 6),
            expiring_ws.Cells(last_row, 8),
        ).Value2
    )
    missing_rows = []
    for offset, row in enumerate(rows):
        state = normalize_state(row[0])
        class_code = normalize_class_code(row[1])
        current_rate = row[2]
        if current_rate in (None, ""):
            missing_rows.append((DEST_START_ROW + offset, state, class_code))

    if not missing_rows:
        log("No blank Expiring Dump H rates require PDF fallback.")
        return 0, 0

    pdf_path = select_policy_pdf()
    log(f"Reading PDF fallback rates from: {pdf_path}")
    policy_rates = extract_policy_rates(pdf_path)
    if not policy_rates:
        raise ValueError(
            f"No State/Class/Rate rows could be extracted from PDF: {pdf_path}"
        )
    log(f"Unique State/Class rates extracted from PDF: {len(policy_rates)}")

    filled = 0
    unmatched = 0
    for excel_row, state, class_code in missing_rows:
        rate = policy_rates.get((state, class_code))
        if rate is None:
            unmatched += 1
            log(
                f"PDF rate not found for Expiring row {excel_row}: "
                f"state={state or '<blank>'}, class={class_code or '<blank>'}"
            )
            continue
        expiring_ws.Cells(excel_row, 8).Value = rate
        filled += 1
        log(
            f"PDF fallback filled Expiring H{excel_row}: "
            f"state={state}, class={class_code}, rate={rate}"
        )

    return filled, unmatched


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

        # PDF fallback for rates that the B -> P lookup did not find.
        # Match Expiring Dump F (State) + G (Class Code) against the newest
        # downloaded policy PDF and write Rate Per $100 Remuneration into H.
        pdf_filled_count, pdf_unmatched_count = (
            fill_blank_expiring_rates_from_pdf(
                expiring_ws,
                template_last_row,
            )
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
        log(f"Expiring Dump rows checked: {max(0, template_last_row - 2)}")
        log(f"B->P matches found for destination H: {match_count_b}")
        log(f"Nonblank B keys with no P match: {missing_count_b}")
        log(f"F->H matches found for destination K: {match_count_f}")
        log(f"Nonblank F keys with no H match: {missing_count_f}")
        log(f"Blank H rates filled from policy PDF: {pdf_filled_count}")
        log(f"Blank H rates still unmatched after PDF: {pdf_unmatched_count}")
        log("Expiring B -> Rates P -> Rates L -> Expiring H.")
        log("Expiring F -> Rates H -> Rates N -> Expiring K.")
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
