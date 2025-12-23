import json
import os
import time
from typing import List, Optional

from project_config import SHEET_TITLE_DEFAULT, WEB_STUFF_DIR
SHEET_ID_PATH = os.path.join(WEB_STUFF_DIR, "price_matrix_sheet.json")
DEFAULT_SHEET_TITLE = SHEET_TITLE_DEFAULT
FEEDBACK_HEADERS = ("Correct Divine Orb", "Correct Exalted Orb")
FEEDBACK_START_COL = 4
FEEDBACK_END_COL = 5


def _matrix_to_values(
    col_labels: List[str],
    matrix: List[List[Optional[float]]],
    row_labels: Optional[List[str]] = None,
) -> List[List[str]]:
    if row_labels is None:
        row_labels = col_labels
    values: List[List[str]] = []
    values.append(["Currency"] + col_labels)
    for i, row in enumerate(matrix):
        label = row_labels[i] if i < len(row_labels) else ""
        out = [label]
        for v in row:
            out.append("" if v is None else f"{v:.2f}")
        values.append(out)
    return values


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


def _build_sheet_service(service_account_path: str, oauth_client_path: str):
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

    return build("sheets", "v4", credentials=creds)


def export_matrix_to_sheet(
    currencies: List[str],
    matrix: List[List[Optional[float]]],
    oauth_client_path: str,
    service_account_path: str,
    sheet_id: Optional[str],
    sheet_name: str,
    sheet_title: str,
    apply_format: bool,
    row_labels: Optional[List[str]] = None,
) -> str:
    service = _build_sheet_service(service_account_path, oauth_client_path)

    if not sheet_id:
        body = {"properties": {"title": sheet_title}}
        sheet = service.spreadsheets().create(body=body, fields="spreadsheetId").execute()
        sheet_id = sheet["spreadsheetId"]
    with open(SHEET_ID_PATH, "w", encoding="utf-8") as f:
        json.dump(
            {"spreadsheetId": sheet_id, "url": f"https://docs.google.com/spreadsheets/d/{sheet_id}/edit"},
            f,
        )

    values = _matrix_to_values(currencies, matrix, row_labels=row_labels)
    header = values[0]
    data_rows = values[1:]
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
    requests = [
        {"range": "A1:B1", "values": [["Last Updated", timestamp]]},
        {
            "range": f"A{top_offset_rows + 1}:{_a1(top_offset_rows + 1, len(header))}",
            "values": [header],
        },
    ]
    feedback_header_range = (
        f"{_a1(top_offset_rows + 1, FEEDBACK_START_COL)}:"
        f"{_a1(top_offset_rows + 1, FEEDBACK_END_COL)}"
    )
    requests.append({"range": feedback_header_range, "values": [list(FEEDBACK_HEADERS)]})
    if data_rows:
        end_cell = _a1(top_offset_rows + 1 + len(data_rows), len(header))
        requests.append(
            {
                "range": f"A{top_offset_rows + 2}:{end_cell}",
                "values": data_rows,
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

        rows = len(matrix)
        cols = len(matrix[0]) if matrix else 0
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

    return sheet_id


def read_feedback_columns(
    sheet_id: str,
    sheet_name: str,
    service_account_path: str,
    oauth_client_path: str,
    header_row: int = 6,
    first_data_row: int = 7,
    start_col: int = 4,
    end_col: int = 5,
) -> List[List[str]]:
    service = _build_sheet_service(service_account_path, oauth_client_path)
    start_cell = _a1(first_data_row, start_col)
    end_cell = _a1(1000, end_col)
    range_name = f"{sheet_name}!{start_cell}:{end_cell}"
    resp = service.spreadsheets().values().get(
        spreadsheetId=sheet_id,
        range=range_name,
    ).execute()
    return resp.get("values", [])


def clear_sheet(
    sheet_id: str,
    sheet_name: str,
    service_account_path: str,
    oauth_client_path: str,
) -> None:
    service = _build_sheet_service(service_account_path, oauth_client_path)
    service.spreadsheets().values().clear(
        spreadsheetId=sheet_id,
        range=sheet_name,
        body={},
    ).execute()
