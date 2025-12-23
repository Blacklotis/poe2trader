import argparse
import json
import os
import time
from typing import List, Optional


PROJECT_PATH = "project.json"
DEFAULT_SHEET_TITLE = "Price Matrix"
SHEET_ID_PATH = "price_matrix_sheet.json"
CACHE_PATH = "price_matrix_cache.json"


def _load_currencies(path: str) -> List[str]:
    with open(path, "r", encoding="utf-8") as f:
        data = json.load(f)
    return [str(c) for c in data.get("currencies", [])]


def _matrix_to_values(currencies: List[str], matrix: List[List[Optional[float]]]) -> List[List[str]]:
    values: List[List[str]] = []
    values.append([""] + currencies)
    for i, row in enumerate(matrix):
        out = [currencies[i]]
        for v in row:
            out.append("" if v is None else f"{v:.2f}")
        values.append(out)
    return values


def _matrix_to_strings(matrix: List[List[Optional[float]]]) -> List[List[str]]:
    out: List[List[str]] = []
    for row in matrix:
        out.append(["" if v is None else f"{v:.2f}" for v in row])
    return out


def _a1(row: int, col: int) -> str:
    name = ""
    c = col
    while c > 0:
        c, rem = divmod(c - 1, 26)
        name = chr(65 + rem) + name
    return f"{name}{row}"


def _load_cache() -> Optional[dict]:
    if not os.path.exists(CACHE_PATH):
        return None
    with open(CACHE_PATH, "r", encoding="utf-8") as f:
        return json.load(f)


def _save_cache(currencies: List[str], matrix: List[List[Optional[float]]]) -> None:
    data = {"currencies": currencies, "matrix": _matrix_to_strings(matrix)}
    with open(CACHE_PATH, "w", encoding="utf-8") as f:
        json.dump(data, f)


def _export_google_sheet(
    oauth_client_path: str,
    service_account_path: str,
    values: List[List[str]],
    title: str,
    sheet_name: str,
    sheet_id: Optional[str],
    apply_format: bool,
) -> str:
    try:
        from google.oauth2.credentials import Credentials
        from google.oauth2.service_account import Credentials as ServiceCredentials
        from google_auth_oauthlib.flow import InstalledAppFlow
        from googleapiclient.discovery import build
    except Exception as exc:
        raise SystemExit(
            "Missing Google API packages. Install: pip install google-auth-oauthlib google-api-python-client"
        ) from exc

    scopes = ["https://www.googleapis.com/auth/spreadsheets"]
    creds = None
    if service_account_path:
        creds = ServiceCredentials.from_service_account_file(service_account_path, scopes=scopes)
    else:
        token_path = "token.json"
        if os.path.exists(token_path):
            creds = Credentials.from_authorized_user_file(token_path, scopes=scopes)
        if not creds or not creds.valid:
            flow = InstalledAppFlow.from_client_secrets_file(oauth_client_path, scopes=scopes)
            creds = flow.run_local_server(port=0)
            with open(token_path, "w", encoding="utf-8") as f:
                f.write(creds.to_json())

    service = build("sheets", "v4", credentials=creds)

    if not sheet_id:
        body = {"properties": {"title": title}}
        sheet = service.spreadsheets().create(body=body, fields="spreadsheetId").execute()
        sheet_id = sheet["spreadsheetId"]
    with open(SHEET_ID_PATH, "w", encoding="utf-8") as f:
        json.dump(
            {"spreadsheetId": sheet_id, "url": f"https://docs.google.com/spreadsheets/d/{sheet_id}/edit"},
            f,
        )

    # Fetch sheet ID for formatting and update
    sheet_info = service.spreadsheets().get(
        spreadsheetId=sheet_id,
        ranges=[sheet_name],
        fields="sheets(properties(sheetId,title),conditionalFormats)",
    ).execute()
    sheet_id_num = None
    for s in sheet_info.get("sheets", []):
        props = s.get("properties", {})
        if props.get("title") == sheet_name:
            sheet_id_num = props.get("sheetId")
            break
    if sheet_id_num is None:
        raise SystemExit(f"Sheet name not found: {sheet_name}")

    # Build data with timestamp row + header row
    timestamp = time.strftime("%Y-%m-%d %H:%M:%S")
    header = values[0]
    data_rows = values[1:]
    cached = _load_cache()
    currencies = header[1:]
    matrix_strings = [row[1:] for row in data_rows]

    requests = []
    # Always update timestamp row (A1:B1)
    requests.append(
        {
            "range": "A1:B1",
            "values": [["Last Updated", timestamp]],
        }
    )

    if not cached or cached.get("currencies") != currencies:
        # Full header + data update if currencies changed or no cache.
        requests.append(
            {
                "range": f"A2:{_a1(2, len(header))}",
                "values": [header],
            }
        )
        if data_rows:
            end_col = _a1(2, len(header))[:-1]
            end_cell = _a1(2 + len(data_rows), len(header))
            requests.append(
                {
                    "range": f"A3:{end_cell}",
                    "values": data_rows,
                }
            )
    else:
        # Only update changed cells in data grid.
        old = cached.get("matrix", [])
        for i, row in enumerate(matrix_strings):
            old_row = old[i] if i < len(old) else []
            for j, val in enumerate(row):
                old_val = old_row[j] if j < len(old_row) else ""
                if val == old_val:
                    continue
                row_idx = 3 + i
                col_idx = 2 + j
                requests.append(
                    {
                        "range": _a1(row_idx, col_idx),
                        "values": [[val]],
                    }
                )

    if requests:
        service.spreadsheets().values().batchUpdate(
            spreadsheetId=sheet_id,
            body={"valueInputOption": "RAW", "data": requests},
        ).execute()

    if apply_format:
        # Clear existing conditional formats on the sheet.
        for s in sheet_info.get("sheets", []):
            props = s.get("properties", {})
            if props.get("title") != sheet_name:
                continue
            formats = s.get("conditionalFormats", [])
            if formats:
                delete_reqs = []
                for i in range(len(formats) - 1, -1, -1):
                    delete_reqs.append(
                        {"deleteConditionalFormatRule": {"sheetId": sheet_id_num, "index": i}}
                    )
                service.spreadsheets().batchUpdate(
                    spreadsheetId=sheet_id, body={"requests": delete_reqs}
                ).execute()
            break

        rows = len(matrix_strings)
        cols = len(matrix_strings[0]) if matrix_strings else 0
        if rows and cols:
            requests = [
                {
                    "addConditionalFormatRule": {
                        "rule": {
                            "ranges": [
                                {
                                    "sheetId": sheet_id_num,
                                    "startRowIndex": 2,
                                    "endRowIndex": 2 + rows,
                                    "startColumnIndex": 1,
                                    "endColumnIndex": 1 + cols,
                                }
                            ],
                            "gradientRule": {
                                "minpoint": {
                                    "color": {"red": 1.0, "green": 0.9, "blue": 0.9},
                                    "type": "MIN",
                                },
                                "midpoint": {
                                    "color": {"red": 1.0, "green": 1.0, "blue": 1.0},
                                    "type": "PERCENTILE",
                                    "value": "50",
                                },
                                "maxpoint": {
                                    "color": {"red": 0.8, "green": 1.0, "blue": 0.8},
                                    "type": "MAX",
                                },
                            },
                        },
                        "index": 0,
                    }
                }
            ]
            service.spreadsheets().batchUpdate(
                spreadsheetId=sheet_id, body={"requests": requests}
            ).execute()

    _save_cache(currencies, matrix_strings)

    return sheet_id


def _load_matrix_from_cache() -> Optional[tuple[List[str], List[List[Optional[float]]]]]:
    cached = _load_cache()
    if not cached:
        return None
    currencies = [str(c) for c in cached.get("currencies", [])]
    matrix_str = cached.get("matrix", [])
    matrix: List[List[Optional[float]]] = []
    for row in matrix_str:
        out_row = []
        for v in row:
            if v == "" or v is None:
                out_row.append(None)
            else:
                try:
                    out_row.append(float(v))
                except Exception:
                    out_row.append(None)
        matrix.append(out_row)
    return currencies, matrix


def main() -> None:
    parser = argparse.ArgumentParser(description="Export the currency price matrix to Google Sheets.")
    parser.add_argument("--limit", type=int, default=0, help="Limit number of currencies (0 = all).")
    parser.add_argument("--gsheet", action="store_true", help="Export to Google Sheets.")
    parser.add_argument("--oauth-client", default="", help="Path to OAuth client JSON.")
    parser.add_argument("--service-account", default="", help="Path to service account JSON.")
    parser.add_argument("--sheet-id", default="", help="Existing spreadsheet ID.")
    parser.add_argument("--sheet-name", default="Sheet1")
    parser.add_argument("--sheet-title", default=DEFAULT_SHEET_TITLE)
    parser.add_argument("--no-format", action="store_true", help="Disable sheet formatting.")
    args = parser.parse_args()

    cached = _load_matrix_from_cache()
    if cached:
        currencies, matrix = cached
    else:
        currencies = _load_currencies(PROJECT_PATH)
        size = len(currencies)
        matrix = [[None for _ in range(size)] for _ in range(size)]
    if args.limit and args.limit > 0:
        currencies = currencies[: args.limit]
        matrix = [row[: len(currencies)] for row in matrix[: len(currencies)]]

    if args.gsheet and (args.oauth_client or args.service_account):
        sheet_id = _export_google_sheet(
            args.oauth_client,
            args.service_account,
            _matrix_to_values(currencies, matrix),
            args.sheet_title,
            args.sheet_name,
            args.sheet_id or None,
            not args.no_format,
        )
        print(f"Sheet: https://docs.google.com/spreadsheets/d/{sheet_id}/edit")
    else:
        raise SystemExit("Use --gsheet with --service-account or --oauth-client to export.")


if __name__ == "__main__":
    main()
