import time
from typing import Optional, Tuple

import os
import cv2
import mss
import numpy as np
import pytesseract
import tkinter as tk

from price_matrix import export_cached_matrix
from project_config import (
    DEBUG_DIR,
    DEBUG_DUMP_DEFAULT,
    DEBUG_DUMP_INTERVAL_SEC_DEFAULT,
    NamedRect,
    PROJECT_PATH,
    SHEET_NAME_DEFAULT,
    SHEET_SERVICE_ACCOUNT_PATH,
    SHEET_TITLE_DEFAULT,
    SHEET_UPDATE_EVERY_DEFAULT,
    WEB_STUFF_DIR,
    load_currencies,
    load_ratio_regions,
    load_triggers,
)
from trades import TradeRunner
from ocr_utils import compute_display, format_decimals, get_ocr_preps, prep_for_ocr, read_ratio_from_image
from matrix_export import export_matrix_to_sheet, load_matrix_from_cache, load_sheet_id

DEBUG_DUMP = DEBUG_DUMP_DEFAULT
DEBUG_DUMP_INTERVAL_SEC = DEBUG_DUMP_INTERVAL_SEC_DEFAULT
SHEET_SERVICE_ACCOUNT = SHEET_SERVICE_ACCOUNT_PATH
SHEET_NAME = SHEET_NAME_DEFAULT
SHEET_TITLE = SHEET_TITLE_DEFAULT
SHEET_UPDATE_EVERY = SHEET_UPDATE_EVERY_DEFAULT
HELP_TEXT = """Usage:
  python main.py [--help]

Per-ratio config fields (project.json -> ratios):
  buy: string (required)
  sell: string (required)
  x, y, w, h: integers (required)
  ocr_mode: "multi" or "gray_only"
  scale: float (OCR resize scale)
  decimal_rule: "" or "tail_zero_two_dp"
  expected_min: float
  expected_max: float
  expected_integer_only: true/false
  debug_output: true/false
  clean_min_area_ratio: float
  clean_left_margin_ratio: float
  clean_left_max_area_ratio: float
  clean_left_max_width: int

Example:
  See config.example.json
"""


def configure_tesseract() -> None:
    env_path = os.environ.get("TESSERACT_CMD", "").strip()
    if env_path:
        pytesseract.pytesseract.tesseract_cmd = env_path
        return

    candidates = [
        r"C:\Program Files\Tesseract-OCR\tesseract.exe",
        r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
    ]
    for p in candidates:
        if os.path.exists(p):
            pytesseract.pytesseract.tesseract_cmd = p
            return




def init_overlay(ratios: Tuple[NamedRect, ...]) -> tk.Tk:
    root = tk.Tk()
    root.overrideredirect(True)
    root.attributes("-topmost", True)
    root.attributes("-transparentcolor", "magenta")
    root.configure(bg="magenta")

    sw = root.winfo_screenwidth()
    sh = root.winfo_screenheight()
    root.geometry(f"{sw}x{sh}+0+0")

    canvas = tk.Canvas(root, bg="magenta", highlightthickness=0)
    canvas.pack(fill="both", expand=True)

    for rr in ratios:
        tx0 = rr.rect.x
        ty0 = rr.rect.y
        tx1 = rr.rect.x + rr.rect.w
        ty1 = rr.rect.y + rr.rect.h
        canvas.create_rectangle(tx0, ty0, tx1, ty1, outline="#ff3b3b", width=3)

    root.update_idletasks()
    root.update()
    return root


def _load_or_init_matrix(currencies: list[str]) -> list[list[Optional[float]]]:
    cached = load_matrix_from_cache()
    if cached:
        cached_currencies, matrix = cached
        if cached_currencies == currencies:
            return matrix
    size = len(currencies)
    return [[None for _ in range(size)] for _ in range(size)]


def _currency_index_map(currencies: list[str]) -> dict[str, int]:
    return {str(name): idx for idx, name in enumerate(currencies)}


def main():
    configure_tesseract()
    ratios = load_ratio_regions(PROJECT_PATH)
    triggers = load_triggers(PROJECT_PATH)
    currencies = load_currencies(PROJECT_PATH)
    matrix = _load_or_init_matrix(currencies)
    currency_map = _currency_index_map(currencies)
    trade_runner = (
        TradeRunner(PROJECT_PATH, delay_sec=1.0)
        if os.path.exists(PROJECT_PATH)
        else None
    )
    trigger_state = {}
    last_config_mtime = os.path.getmtime(PROJECT_PATH) if os.path.exists(PROJECT_PATH) else 0.0
    sheet_id = load_sheet_id()
    sheet_ready = os.path.exists(SHEET_SERVICE_ACCOUNT)
    overlay = init_overlay(ratios)
    if DEBUG_DUMP:
        os.makedirs(DEBUG_DIR, exist_ok=True)
        for name in os.listdir(DEBUG_DIR):
            path = os.path.join(DEBUG_DIR, name)
            if os.path.isfile(path):
                os.remove(path)
    last_dump_ts = time.monotonic()
    dump_index = 0
    with mss.mss() as sct:
        while True:
            if os.path.exists(PROJECT_PATH):
                mtime = os.path.getmtime(PROJECT_PATH)
                if mtime != last_config_mtime:
                    try:
                        ratios = load_ratio_regions(PROJECT_PATH)
                        triggers = load_triggers(PROJECT_PATH)
                        currencies = load_currencies(PROJECT_PATH)
                        matrix = _load_or_init_matrix(currencies)
                        currency_map = _currency_index_map(currencies)
                        last_config_mtime = mtime
                        if overlay is not None:
                            overlay.destroy()
                        overlay = init_overlay(ratios)
                    except Exception:
                        pass
            print(f"Count: {dump_index + 1}")
            for rr in ratios:
                img = np.array(sct.grab(rr.rect.as_mss()))[:, :, :3]
                preps = get_ocr_preps(rr)
                ratio, raw, candidates = read_ratio_from_image(img, preps)
                display = None
                if ratio is not None:
                    display = compute_display(raw, ratio, rr)
                if display is None:
                    for _, c_raw, c_ratio in candidates:
                        if c_ratio is None:
                            continue
                        display = compute_display(c_raw, c_ratio, rr)
                        if display is not None:
                            break
                if display is None:
                    print(f"{rr.label}: ?")
                else:
                    print(f"{rr.label}: {format_decimals(display, 2)}")
                    row = currency_map.get(rr.buy)
                    col = currency_map.get(rr.sell)
                    if row is not None and col is not None:
                        if 0 <= row < len(matrix) and 0 <= col < len(matrix):
                            matrix[row][col] = display
                if display is not None:
                    for trig in triggers:
                        if str(trig.get("ratio")) != rr.key:
                            continue
                        op = str(trig.get("op", ">"))
                        target = float(trig.get("value", 0.0))
                        is_true = display > target if op == ">" else display < target
                        was_true = trigger_state.get(rr.key, False)
                        if is_true and not was_true and trade_runner is not None:
                            trade_runner.run_trade(
                                str(trig.get("trade", "")),
                                overrides=trig.get("overrides"),
                            )
                        trigger_state[rr.key] = is_true
                for label, c_raw, c_ratio in candidates:
                    if not rr.debug_output:
                        continue
                    if c_ratio is None:
                        print(f"  {label}: raw='{c_raw}' value=?")
                    else:
                        adjusted = compute_display(c_raw, c_ratio, rr)
                        if adjusted is None:
                            print(f"  {label}: raw='{c_raw}' value=?")
                        else:
                            print(f"  {label}: raw='{c_raw}' value={format_decimals(adjusted, 2)}")
                if DEBUG_DUMP and (time.monotonic() - last_dump_ts >= DEBUG_DUMP_INTERVAL_SEC):
                    raw_path = os.path.join(DEBUG_DIR, f"{rr.file_tag}_{dump_index}_raw.png")
                    bin_path = os.path.join(DEBUG_DIR, f"{rr.file_tag}_{dump_index}_bin.png")
                    bin_img = prep_for_ocr(img, scale=rr.scale)
                    cv2.imwrite(raw_path, img)
                    cv2.imwrite(bin_path, bin_img)
            print("===========")
            if sheet_ready and (dump_index + 1) % SHEET_UPDATE_EVERY == 0:
                sheet_id = export_matrix_to_sheet(
                    currencies,
                    matrix,
                    oauth_client_path="",
                    service_account_path=SHEET_SERVICE_ACCOUNT,
                    sheet_id=sheet_id,
                    sheet_name=SHEET_NAME,
                    sheet_title=SHEET_TITLE,
                    apply_format=False,
                )
            overlay.update_idletasks()
            overlay.update()
            if DEBUG_DUMP and (time.monotonic() - last_dump_ts >= DEBUG_DUMP_INTERVAL_SEC):
                last_dump_ts = time.monotonic()
                dump_index += 1
            time.sleep(DEBUG_DUMP_INTERVAL_SEC)


if __name__ == "__main__":
    import sys

    if "--help" in sys.argv or "-h" in sys.argv:
        print(HELP_TEXT)
        raise SystemExit(0)
    if "--export-matrix" in sys.argv:
        service_path = os.path.join(WEB_STUFF_DIR, "poe2trader-2e14fc353067.json")
        sheet_id = export_cached_matrix(
            oauth_client_path="",
            service_account_path=service_path,
            sheet_id=None,
            sheet_name=SHEET_NAME,
            sheet_title=SHEET_TITLE,
            apply_format=True,
        )
        print(f"Sheet: https://docs.google.com/spreadsheets/d/{sheet_id}/edit")
        raise SystemExit(0)
    main()
