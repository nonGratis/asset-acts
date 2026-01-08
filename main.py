import sys

from src.cli import setup_args, apply_args
from src.config import ASSETS_SPREADSHEET_ID, DEPARTMENTS_SPREADSHEET_ID, log
from src.google_api import (
    check_constants,
    build_services,
    ensure_file_is_spreadsheet,
    load_departments,
    parse_assets,
)
from src.document_export import create_act_docs_local
from src.formatters import fmt_number


def main():
    check_constants()
    sheets_svc, drive_svc, _ = build_services()

    try:
        ensure_file_is_spreadsheet(drive_svc, ASSETS_SPREADSHEET_ID, "Assets spreadsheet")
        ensure_file_is_spreadsheet(drive_svc, DEPARTMENTS_SPREADSHEET_ID, "Departments spreadsheet")
    except SystemExit:
        raise

    log.info(f"Start processing assets spreadsheet (ID={ASSETS_SPREADSHEET_ID})")
    departments = load_departments(sheets_svc)
    per_owner, stats = parse_assets(sheets_svc, departments)

    if not per_owner:
        log.info("No valid owners/items found; nothing to generate.")
        return

    created = create_act_docs_local(per_owner, drive_svc)

    log.info(
        f"rows_processed={stats['rows_processed']}, "
        f"rows_skipped={stats['rows_skipped']}, "
        f"owners_skipped={stats['owners_skipped']}, "
        f"acts_generated={len(created)}, "
        f"items_in_acts={stats['total_items_in_acts']}, "
        f"total_value_generated={fmt_number(stats['total_value_generated'])}"
    )


if __name__ == "__main__":
    try:
        args = setup_args()
        apply_args(args)
        main()
    except Exception as exc:
        log.error(f"Unhandled exception: {exc}")
        sys.exit(1)