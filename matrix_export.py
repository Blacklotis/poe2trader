import json
import os
import time
from typing import List, Optional


BASE_DIR = os.path.dirname(__file__)
WEB_STUFF_DIR = os.path.join(BASE_DIR, "web_stuff")
SHEET_ID_PATH = os.path.join(WEB_STUFF_DIR, "price_matrix_sheet.json")
CACHE_PATH = os.path.join(WEB_STUFF_DIR, "price_matrix_cache.json")
DEFAULT_SHEET_TITLE = "Price Matrix"


def _matrix_to_values(currencies: List[str], matrix: List[List[Optional[float]]]) -> List[List[str]]:
    values: List[List[str]] = []
    values.append(["Currency"] + currencies)
    for i, row in enumerate(matrix):
        out = [currencies[i]]
        for v in row:
            out.append("" if v is None else f"{v:.2f}")
        values.append(out)
    return values


def _matrix_to_strings(matrix: List[List[Optional[float]]]) -> List[List[str]]:
    out: List[List[str]] = []
    for row in matrix:
        row_out = []
        for v in row:
            if v is None:
                row_out.append("")
            elif isinstance(v, (int, float)):
                row_out.append(f"{v:.2f}")
            else:
                row_out.append(str(v))
        out.append(row_out)
    return out


def _a1(row: int, col: int) -> str:
    name = ""
    c = col
    while c > 0:
        c, rem = divmod(c - 1, 26)
        name = chr(65 + rem) + name
    return f"{name}{row}"


def load_sheet_id() -> Optional[str]:
    if not os.path.exists(SHEET_ID_PATH):
        return None
    with open(SHEET_ID_PATH, "r", encoding="utf-8") as f:
        data = json.load(f)
    return str(data.get("spreadsheetId", "")) or None


def load_cache() -> Optional[dict]:
    if not os.path.exists(CACHE_PATH):
        return None
    with open(CACHE_PATH, "r", encoding="utf-8") as f:
        return json.load(f)


def save_cache(currencies: List[str], matrix: List[List[Optional[float]]]) -> None:
    data = {"currencies": currencies, "matrix": _matrix_to_strings(matrix)}
    with open(CACHE_PATH, "w", encoding="utf-8") as f:
        json.dump(data, f)


def load_matrix_from_cache() -> Optional[tuple[List[str], List[List[Optional[float]]]]]:
    cached = load_cache()
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


def export_matrix_to_sheet(
    currencies: List[str],
    matrix: List[List[Optional[float]]],
    oauth_client_path: str,
    service_account_path: str,
    sheet_id: Optional[str],
    sheet_name: str,
    sheet_title: str,
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
        body = {"properties": {"title": sheet_title}}
        sheet = service.spreadsheets().create(body=body, fields="spreadsheetId").execute()
        sheet_id = sheet["spreadsheetId"]
    with open(SHEET_ID_PATH, "w", encoding="utf-8") as f:
        json.dump(
            {"spreadsheetId": sheet_id, "url": f"https://docs.google.com/spreadsheets/d/{sheet_id}/edit"},
            f,
        )

    values = _matrix_to_values(currencies, matrix)
    header = values[0]
    data_rows = values[1:]
    cached = load_cache()
    currencies_cached = header[1:]
    matrix_strings = [row[1:] for row in data_rows]

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

    top_offset_rows = 5
    timestamp = time.strftime("%Y-%m-%d %H:%M:%S")
    requests = [{"range": "A1:B1", "values": [["Last Updated", timestamp]]}]

    if not cached or cached.get("currencies") != currencies_cached:
        requests.append(
            {
                "range": f"A{top_offset_rows + 1}:{_a1(top_offset_rows + 1, len(header))}",
                "values": [header],
            }
        )
        if data_rows:
            end_cell = _a1(top_offset_rows + 1 + len(data_rows), len(header))
            requests.append(
                {
                    "range": f"A{top_offset_rows + 2}:{end_cell}",
                    "values": data_rows,
                }
            )
    else:
        old = cached.get("matrix", [])
        for i, row in enumerate(matrix_strings):
            old_row = old[i] if i < len(old) else []
            for j, val in enumerate(row):
                old_val = old_row[j] if j < len(old_row) else ""
                if val == old_val:
                    continue
                row_idx = top_offset_rows + 2 + i
                col_idx = 2 + j
                requests.append({"range": _a1(row_idx, col_idx), "values": [[val]]})

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
                                    "startRowIndex": top_offset_rows + 1,
                                    "endRowIndex": top_offset_rows + 1 + rows,
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
            header_row = top_offset_rows
            requests.append(
                {
                    "repeatCell": {
                        "range": {
                            "sheetId": sheet_id_num,
                            "startRowIndex": header_row,
                            "endRowIndex": header_row + 1,
                            "startColumnIndex": 0,
                            "endColumnIndex": 1 + cols,
                        },
                        "cell": {
                            "userEnteredFormat": {
                                "backgroundColor": {"red": 0.80, "green": 0.90, "blue": 1.0},
                                "textFormat": {
                                    "foregroundColor": {"red": 0.8, "green": 0.0, "blue": 0.0},
                                    "bold": True,
                                },
                                "horizontalAlignment": "CENTER",
                            }
                        },
                        "fields": "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)",
                    }
                }
            )
            requests.append(
                {
                    "repeatCell": {
                        "range": {
                            "sheetId": sheet_id_num,
                            "startRowIndex": header_row,
                            "endRowIndex": header_row + 1 + rows,
                            "startColumnIndex": 0,
                            "endColumnIndex": 1,
                        },
                        "cell": {
                            "userEnteredFormat": {
                                "backgroundColor": {"red": 0.80, "green": 0.90, "blue": 1.0},
                                "textFormat": {
                                    "foregroundColor": {"red": 0.8, "green": 0.0, "blue": 0.0},
                                    "bold": True,
                                },
                                "horizontalAlignment": "LEFT",
                            }
                        },
                        "fields": "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)",
                    }
                }
            )
            for idx, name in enumerate(header):
                base = max(140, min(320, 8 * len(name) + 40))
                if idx == 0:
                    base = max(base, 220)
                requests.append(
                    {
                        "updateDimensionProperties": {
                            "range": {
                                "sheetId": sheet_id_num,
                                "dimension": "COLUMNS",
                                "startIndex": idx,
                                "endIndex": idx + 1,
                            },
                            "properties": {"pixelSize": base},
                            "fields": "pixelSize",
                        }
                    }
                )
            border_style = {"style": "SOLID_MEDIUM", "width": 2, "color": {"red": 0, "green": 0, "blue": 0}}
            requests.append(
                {
                    "updateBorders": {
                        "range": {
                            "sheetId": sheet_id_num,
                            "startRowIndex": header_row,
                            "endRowIndex": header_row + 1,
                            "startColumnIndex": 0,
                            "endColumnIndex": 1 + cols,
                        },
                        "top": border_style,
                        "bottom": border_style,
                        "left": border_style,
                        "right": border_style,
                        "innerHorizontal": border_style,
                        "innerVertical": border_style,
                    }
                }
            )
            requests.append(
                {
                    "updateBorders": {
                        "range": {
                            "sheetId": sheet_id_num,
                            "startRowIndex": header_row,
                            "endRowIndex": header_row + 1 + rows,
                            "startColumnIndex": 0,
                            "endColumnIndex": 1,
                        },
                        "top": border_style,
                        "bottom": border_style,
                        "left": border_style,
                        "right": border_style,
                        "innerHorizontal": border_style,
                        "innerVertical": border_style,
                    }
                }
            )
            service.spreadsheets().batchUpdate(
                spreadsheetId=sheet_id, body={"requests": requests}
            ).execute()

    save_cache(currencies_cached, matrix_strings)
    return sheet_id
