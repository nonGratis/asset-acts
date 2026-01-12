import os
from decimal import Decimal
from typing import Dict, Any, Optional

from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

from .config import (
    SERVICE_ACCOUNT_KEYFILE,
    SCOPES,
    ASSETS_SPREADSHEET_ID,
    ASSETS_SHEET_NAME,
    DEPARTMENTS_SPREADSHEET_ID,
    DEPARTMENTS_SHEET_NAME,
    TEMPLATE_PATH,
    COL_NAME,
    COL_INVENTORY_NUMBER,
    COL_UNIT,
    COL_QUANTITY,
    COL_PRICE,
    COL_OWNERS,
    COL_GENERATE_FLAG,
    DEPT_COL_CODE,
    DEPT_COL_STATUS,
    DEPT_COL_POSITION,
    DEPT_COL_FULLNAME,
    DEPT_COL_RECEIVER_POSITION,
    DEPT_COL_RECEIVER_FULLNAME,
    ALLOW_ROUNDING_ADJUST,
    log,
)
from .data_utils import (
    safe_get,
    log_row_data,
    parse_string_number,
    quantize_money,
    parse_owner_token,
    validate_required_fields,
    format_ukrainian_name,
    ProcessingStats,
)


def check_constants() -> None:
    """Validate that all required configuration values and files exist.

    Raises:
        SystemExit: If any critical constants or files are missing
    """
    missing = []
    if not os.path.isfile(SERVICE_ACCOUNT_KEYFILE):
        missing.append(f"SERVICE_ACCOUNT_KEYFILE file not found: {SERVICE_ACCOUNT_KEYFILE}")
    if not os.path.isfile(TEMPLATE_PATH):
        missing.append(f"Template file not found: {TEMPLATE_PATH}")
    for keyname, value in (
        ("ASSETS_SPREADSHEET_ID", ASSETS_SPREADSHEET_ID),
        ("DEPARTMENTS_SPREADSHEET_ID", DEPARTMENTS_SPREADSHEET_ID),
    ):
        if not value:
            missing.append(f"Missing constant {keyname}")
    if missing:
        for m in missing:
            log.error(m)
        raise SystemExit("Missing critical constants; aborting.")


def build_services():
    """Build and return Google API service clients.

    Returns:
        Tuple of (sheets_service, drive_service, docs_service)
    """
    creds = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_KEYFILE, scopes=SCOPES
    )
    sheets = build("sheets", "v4", credentials=creds, cache_discovery=False)
    drive = build("drive", "v3", credentials=creds, cache_discovery=False)
    docs = build("docs", "v1", credentials=creds, cache_discovery=False)
    return sheets, drive, docs


def ensure_file_is_spreadsheet(drive_service, file_id: str, label: str) -> None:
    """Verify that file_id exists and is a Google Spreadsheet.

    Args:
        drive_service: Google Drive API service instance
        file_id: Google Drive file ID to verify
        label: Descriptive label for error messages

    Raises:
        SystemExit: If file doesn't exist or is not a spreadsheet
    """
    try:
        meta = drive_service.files().get(fileId=file_id, fields="id, name, mimeType").execute()
    except HttpError as e:
        log.error(f"Cannot fetch {label} (id={file_id}): {e}")
        raise SystemExit(1)
    mime = meta.get("mimeType", "")
    if mime != "application/vnd.google-apps.spreadsheet":
        log.error(
            f"{label} (id={file_id}) is not a Google Spreadsheet (mimeType={mime}). "
            "Please provide the correct spreadsheet ID."
        )
        raise SystemExit(1)
    log.info(f"{label} found: {meta.get('name', '<untitled>')} (id={file_id})")


def read_sheet_values(sheets_service, spreadsheet_id: str, sheet_name: str):
    """Read all values from a Google Sheet.

    Args:
        sheets_service: Google Sheets API service instance
        spreadsheet_id: ID of the spreadsheet
        sheet_name: Name of the sheet within the spreadsheet

    Returns:
        List of rows, where each row is a list of cell values

    Raises:
        SystemExit: If the file is not a spreadsheet
        HttpError: For other API errors
    """
    rng = f"{sheet_name}"
    try:
        res = sheets_service.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id, range=rng
        ).execute()
        return res.get("values", [])
    except HttpError as e:
        msg = str(e)
        if "This operation is not supported for this document" in msg or "not supported for this document" in msg:
            log.error(
                f"Sheets API: file id {spreadsheet_id} is not a spreadsheet or not supported by Sheets API. "
                "Check that the ID points to a Google Sheet and the service account has access."
            )
            raise SystemExit(1)
        log.error(f"Sheets API error for spreadsheet={spreadsheet_id}, range={rng}: {e}")
        raise


def load_departments(sheets_service) -> Dict[str, Dict[str, str]]:
    """Load department data from Google Sheets.

    Args:
        sheets_service: Google Sheets API service instance

    Returns:
        Dictionary mapping normalized department codes to department info
    """
    vals = read_sheet_values(sheets_service, DEPARTMENTS_SPREADSHEET_ID, DEPARTMENTS_SHEET_NAME)
    depts = {}
    if not vals or len(vals) < 2:
        log.warning("Departments sheet empty or missing rows.")
        return depts
    for i, row in enumerate(vals[1:], start=2):
        code = str(safe_get(row, DEPT_COL_CODE, "")).strip()
        if not code:
            row_data = log_row_data(
                row,
                [
                    ("fullname", DEPT_COL_FULLNAME),
                    ("position", DEPT_COL_POSITION),
                ],
            )
            log.warning(f"Departments row {i} missing code; skipping. Row data: {row_data}")
            continue
        
        fullname = safe_get(row, DEPT_COL_FULLNAME, "")
        receiver_fullname = safe_get(row, DEPT_COL_RECEIVER_FULLNAME, "")
        
        try:
            formatted_name = format_ukrainian_name(fullname) if fullname else ""
        except ValueError as e:
            log.error(f"Departments row {i}: {e}; skipping.")
            continue
        
        try:
            receiver_formatted = format_ukrainian_name(receiver_fullname) if receiver_fullname else ""
        except ValueError as e:
            log.error(f"Departments row {i} receiver name: {e}; skipping.")
            continue
        
        key = code
        depts[key] = {
            "code": safe_get(row, DEPT_COL_CODE, ""),
            "status": safe_get(row, DEPT_COL_STATUS, ""),
            "position": safe_get(row, DEPT_COL_POSITION, ""),
            "fullname": fullname,
            "formatted_name": formatted_name,
            "receiver_position": safe_get(row, DEPT_COL_RECEIVER_POSITION, ""),
            "receiver_fullname": receiver_fullname,
            "receiver_formatted": receiver_formatted,
        }
    return depts


def _parse_owner_tokens_from_string(owners_raw: str) -> list:
    """Parse owner tokens from raw string.

    Args:
        owners_raw: Raw owner string from spreadsheet (may contain newlines and commas)

    Returns:
        List of cleaned owner token strings
    """
    tokens = []
    for line in str(owners_raw).split('\n'):
        for token in line.split(','):
            token = token.strip()
            if token:
                tokens.append(token)
    return tokens


def _validate_explicit_owners(token_infos: list, qty: int, row_index: int, row: list) -> bool:
    if not all(t[2] for t in token_infos):
        row_data = log_row_data(row, [("name", COL_NAME), ("owners", COL_OWNERS)])
        log.error(f"Row {row_index} mixed explicit and implicit owners; skip. Row data: {row_data}")
        return False
    
    total_spec = sum(t[1] for t in token_infos)
    if total_spec != qty:
        name = safe_get(row, COL_NAME, "")
        owners_raw = safe_get(row, COL_OWNERS, "")
        log.error(
            f"Row {row_index} owner sum {total_spec} != qty {qty} | "
            f"name='{name}' owners='{owners_raw}'"
        )
        return False
    
    return True


def _validate_implicit_owner(token_infos: list, row_index: int, row: list) -> bool:
    if len(token_infos) != 1:
        row_data = log_row_data(row, [("name", COL_NAME), ("owners", COL_OWNERS)])
        log.error(f"Row {row_index} ambiguous multiple owners without counts; skip. Row data: {row_data}")
        return False
    return True


def _validate_and_parse_owner_tokens(tokens: list, qty: int, row_index: int, row: list) -> Optional[list]:
    token_infos = []
    any_explicit = False
    
    for tok in tokens:
        base, num, explicit = parse_owner_token(tok)
        if explicit:
            any_explicit = True
        token_infos.append((base, num, explicit))

    if any_explicit:
        if not _validate_explicit_owners(token_infos, qty, row_index, row):
            return None
    else:
        if not _validate_implicit_owner(token_infos, row_index, row):
            return None
        base = token_infos[0][0]
        token_infos[0] = (base, qty, True)

    return token_infos


def _resolve_owners_with_departments(token_infos: list, departments: Dict[str, Dict[str, str]], 
                                      row_index: int, stats: ProcessingStats) -> list:
    owners_for_row = []
    for base, num, _ in token_infos:
        key = base.strip()
        dept = departments.get(key)
        if not dept:
            log.error(f"Row {row_index} owner '{base}' not found")
            stats.skip_owner()
            continue
        owners_for_row.append((key, int(num), dept))
    return owners_for_row


def _calculate_owner_amounts(owners_for_row: list, unit_price: Decimal, total_price: Decimal, 
                             row_index: int) -> list:
    owner_sums = []
    for _, oqty, _ in owners_for_row:
        owner_sums.append(quantize_money(unit_price * Decimal(oqty)))
    
    sum_owner_sums = sum(owner_sums)
    price_q = quantize_money(total_price)
    if sum_owner_sums != price_q:
        diff = price_q - sum_owner_sums
        if ALLOW_ROUNDING_ADJUST:
            owner_sums[-1] = quantize_money(owner_sums[-1] + diff)
            log.warning(f"Row {row_index} rounding adjusted by {diff} on last owner.")
        else:
            log.warning(f"Row {row_index} rounding mismatch {price_q - sum_owner_sums}; continuing.")
    
    return owner_sums


def _extract_asset_row_data(row: list, row_index: int, stats: ProcessingStats) -> Optional[Dict[str, Any]]:
    required_fields = [
        (COL_NAME, "name"),
        (COL_INVENTORY_NUMBER, "inventory_number"),
        (COL_UNIT, "unit"),
        (COL_QUANTITY, "quantity"),
        (COL_PRICE, "price"),
        (COL_OWNERS, "owners"),
    ]
    
    missing_fields = validate_required_fields(row, required_fields)
    if missing_fields:
        row_data = log_row_data(
            row,
            [
                ("name", COL_NAME),
                ("inventory", COL_INVENTORY_NUMBER),
                ("unit", COL_UNIT),
                ("qty", COL_QUANTITY),
                ("price", COL_PRICE),
                ("owners", COL_OWNERS),
            ],
        )
        log.error(
            f"Row {row_index} missing required fields: {', '.join(missing_fields)}; skipping. "
            f"Row data: {row_data}"
        )
        stats.skip_row()
        return None

    try:
        name = safe_get(row, COL_NAME, "")
        invnum = safe_get(row, COL_INVENTORY_NUMBER, "")
        unit = safe_get(row, COL_UNIT, "").lower()
        qty_raw = safe_get(row, COL_QUANTITY, "")
        price_raw = safe_get(row, COL_PRICE, "")
        owners_raw = safe_get(row, COL_OWNERS, "")
        qty = int(parse_string_number(qty_raw))
        price = parse_string_number(price_raw)

        if qty <= 0:
            log.error(f"Row {row_index} has non-positive quantity {qty}; skip row.")
            stats.skip_row()
            return None

        unit_price = quantize_money(price / Decimal(qty))

        return {
            "name": name,
            "invnum": invnum,
            "unit": unit,
            "qty": qty,
            "price": price,
            "unit_price": unit_price,
            "owners_raw": owners_raw,
        }
    except Exception as e:
        log.error(f"Row {row_index} parse error: {e}; skipping row.")
        stats.skip_row()
        return None


def parse_assets(sheets_service, departments: Dict[str, Dict[str, str]]):
    vals = read_sheet_values(sheets_service, ASSETS_SPREADSHEET_ID, ASSETS_SHEET_NAME)
    if not vals or len(vals) < 2:
        log.info("No assets rows found.")
        return {}, ProcessingStats().to_dict()

    stats = ProcessingStats()
    per_owner: Dict[str, Dict[str, Any]] = {}

    for rindex, row in enumerate(vals[1:], start=2):
        if not any(str(cell).strip() for cell in row):
            continue

        gen_flag = safe_get(row, COL_GENERATE_FLAG, "")
        if str(gen_flag).strip().upper() != "TRUE":
            stats.skip_row()
            continue
        
        asset_data = _extract_asset_row_data(row, rindex, stats)
        if not asset_data:
            continue

        tokens = _parse_owner_tokens_from_string(asset_data["owners_raw"])
        if not tokens:
            row_data = log_row_data(
                row,
                [
                    ("name", COL_NAME),
                    ("inventory", COL_INVENTORY_NUMBER),
                    ("qty", COL_QUANTITY),
                    ("price", COL_PRICE),
                ],
            )
            log.error(f"Row {rindex} no owners; skip. Row data: {row_data}")
            stats.skip_row()
            continue

        token_infos = _validate_and_parse_owner_tokens(tokens, asset_data["qty"], rindex, row)
        if not token_infos:
            stats.skip_row()
            continue

        owners_for_row = _resolve_owners_with_departments(token_infos, departments, rindex, stats)
        if not owners_for_row:
            log.info(f"Row {rindex}: all owners skipped; skipping row.")
            stats.skip_row()
            continue

        owner_sums = _calculate_owner_amounts(
            owners_for_row, asset_data["unit_price"], asset_data["price"], rindex
        )

        for (key, oqty, dept), osum in zip(owners_for_row, owner_sums):
            if key not in per_owner:
                per_owner[key] = {"dept": dept, "items": [], "tot_qty": 0, "tot_sum": Decimal("0.00")}
            per_owner[key]["items"].append(
                {
                    "name": asset_data["name"],
                    "inventory": asset_data["invnum"],
                    "unit": asset_data["unit"],
                    "qty": int(oqty),
                    "unit_price": asset_data["unit_price"],
                    "sum": osum,
                    "note": "",
                }
            )
            per_owner[key]["tot_qty"] += int(oqty)
            per_owner[key]["tot_sum"] += osum
            stats.add_item(osum)

        stats.process_row()

    return per_owner, stats.to_dict()
