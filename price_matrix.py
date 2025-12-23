import argparse
from typing import List, Optional

from matrix_export import export_matrix_to_sheet, load_matrix_from_cache, load_sheet_id
from project_config import (
    PROJECT_PATH,
    SHEET_NAME_DEFAULT,
    SHEET_SERVICE_ACCOUNT_PATH,
    SHEET_TITLE_DEFAULT,
    load_currencies,
)


DEFAULT_SERVICE_ACCOUNT = SHEET_SERVICE_ACCOUNT_PATH


def _load_or_init_matrix(limit: int = 0) -> tuple[List[str], List[List[Optional[float]]]]:
    cached = load_matrix_from_cache()
    if cached:
        currencies, matrix = cached
    else:
        currencies = load_currencies(PROJECT_PATH)
        size = len(currencies)
        matrix = [[None for _ in range(size)] for _ in range(size)]
    if limit and limit > 0:
        currencies = currencies[:limit]
        matrix = [row[: len(currencies)] for row in matrix[: len(currencies)]]
    return currencies, matrix


def export_cached_matrix(
    oauth_client_path: str,
    service_account_path: str,
    sheet_id: Optional[str],
    sheet_name: str,
    sheet_title: str,
    apply_format: bool,
    limit: int = 0,
) -> str:
    currencies, matrix = _load_or_init_matrix(limit)
    return export_matrix_to_sheet(
        currencies,
        matrix,
        oauth_client_path,
        service_account_path,
        sheet_id,
        sheet_name,
        sheet_title,
        apply_format,
    )


def main() -> None:
    parser = argparse.ArgumentParser(description="Export the currency price matrix to Google Sheets.")
    parser.add_argument("--limit", type=int, default=0, help="Limit number of currencies (0 = all).")
    parser.add_argument("--gsheet", action="store_true", help="Export to Google Sheets.")
    parser.add_argument("--no-gsheet", action="store_true", help="Disable Google Sheets export.")
    parser.add_argument("--oauth-client", default="", help="Path to OAuth client JSON.")
    parser.add_argument("--service-account", default=DEFAULT_SERVICE_ACCOUNT, help="Path to service account JSON.")
    parser.add_argument("--sheet-id", default="", help="Existing spreadsheet ID.")
    parser.add_argument("--sheet-name", default=SHEET_NAME_DEFAULT)
    parser.add_argument("--sheet-title", default=SHEET_TITLE_DEFAULT)
    parser.add_argument("--no-format", action="store_true", help="Disable sheet formatting.")
    args = parser.parse_args()

    currencies, matrix = _load_or_init_matrix(args.limit)

    if args.no_gsheet:
        raise SystemExit("Google Sheets export disabled.")
    if not args.gsheet:
        args.gsheet = True

    if args.gsheet:
        has_service = bool(args.service_account) and os.path.exists(args.service_account)
        has_oauth = bool(args.oauth_client)
        if not (has_service or has_oauth):
            raise SystemExit("Use --gsheet with --service-account or --oauth-client to export.")
        sheet_id = args.sheet_id or load_sheet_id()
        sheet_id = export_matrix_to_sheet(
            currencies,
            matrix,
            args.oauth_client,
            args.service_account if has_service else "",
            sheet_id,
            args.sheet_name,
            args.sheet_title,
            not args.no_format,
        )
        print(f"Sheet: https://docs.google.com/spreadsheets/d/{sheet_id}/edit")


if __name__ == "__main__":
    main()
