import re
from decimal import Decimal, ROUND_HALF_UP
from typing import List, Tuple, Optional, Dict, Any

from .config import log


def safe_get(row: list, col: int, default=""):
    """Return 1-based column from row safely.

    Args:
        row: list of cell values as returned by Sheets API
        col: 1-based column index
        default: default value if column not found

    Returns:
        Cell value or default if not found
    """
    if row is None:
        return default
    idx = col - 1
    if idx < 0:
        return default
    if idx >= len(row):
        return default
    val = row[idx]
    return val if val is not None else default


def log_row_data(row: list, columns: List[Tuple[str, int]]) -> str:
    """Create formatted string of row data for logging purposes.

    Args:
        row: list of cell values
        columns: list of tuples (column_name, column_index)

    Returns:
        Formatted string representation of row data
    """
    data = {}
    for name, idx in columns:
        val = str(safe_get(row, idx, "")).strip()
        if val:
            data[name] = val
    return str(data) if data else "{empty row}"


def parse_string_number(s) -> Decimal:
    """
    Args:
        s: String representation of number

    Returns:
        Decimal value

    Raises:
        ValueError: If string cannot be parsed as number
    """
    if s is None:
        log.error("Empty numeric")
        raise ValueError("Empty numeric")
    st = str(s).strip()
    st = st.replace("\xa0", "").replace(" ", "")
    st = st.replace(",", ".")
    if st == "":
        log.error("Empty numeric")
        raise ValueError("Empty numeric")
    try:
        return Decimal(st)
    except Exception as e:
        log.error("Invalid numeric '%s': %s", s, e)
        raise


def quantize_money(d: Decimal) -> Decimal:
    """Round Decimal to 2 decimal places (money format).

    Args:
        d: Decimal value to round

    Returns:
        Decimal rounded to 2 decimal places
    """
    return d.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)


def parse_owner_token(tok: str) -> Tuple[str, Optional[int], bool]:
    """Parse owner token to extract base code, quantity, and explicit flag.

    Handles formats like "DEPT-10" (explicit quantity) or "DEPT" (implicit).

    Args:
        tok: Owner token string

    Returns:
        Tuple of (base_code, quantity, is_explicit)
    """
    tok = tok.strip()
    m = re.match(r"^(.*?)-\s*([0-9]+)\s*$", tok)
    if m:
        return m.group(1).strip(), int(m.group(2)), True
    return tok, None, False


def validate_required_fields(row: list, field_definitions: List[Tuple[int, str]]) -> List[str]:
    missing_fields = []
    for col_idx, field_name in field_definitions:
        value = safe_get(row, col_idx, "")
        if not str(value).strip():
            missing_fields.append(field_name)
    return missing_fields


class ProcessingStats:
    def __init__(self):
        self.rows_processed = 0
        self.rows_skipped = 0
        self.owners_skipped = 0
        self.total_items_in_acts = 0
        self.total_value_generated = Decimal("0.00")

    def skip_row(self) -> None:
        self.rows_skipped += 1

    def process_row(self) -> None:
        self.rows_processed += 1

    def skip_owner(self) -> None:
        self.owners_skipped += 1

    def add_item(self, value: Decimal) -> None:
        self.total_items_in_acts += 1
        self.total_value_generated += value

    def to_dict(self) -> Dict[str, Any]:
        """
        Returns:
            Dictionary with all statistics
        """
        return {
            "rows_processed": self.rows_processed,
            "rows_skipped": self.rows_skipped,
            "owners_skipped": self.owners_skipped,
            "total_items_in_acts": self.total_items_in_acts,
            "total_value_generated": self.total_value_generated,
        }
