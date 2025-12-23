import time
from threading import Event
from typing import Optional, Tuple

import os
import cv2
import mss
import numpy as np
import pytesseract
import tkinter as tk
import threading
import ctypes
import ctypes.wintypes
try:
    import keyboard as kb
except Exception:
    kb = None

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
from ocr_utils import (
    compute_display,
    format_decimals,
    get_ocr_preps,
    load_ocr_weights,
    prep_for_ocr,
    read_ratio_from_image,
    save_ocr_weights,
    set_ocr_weights,
)
from matrix_export import clear_sheet, export_matrix_to_sheet, load_sheet_id, read_feedback_columns

DEBUG_DUMP = True
DEBUG_DUMP_INTERVAL_SEC = DEBUG_DUMP_INTERVAL_SEC_DEFAULT
DEBUG_DUMP_MAX_READINGS = None
LOOP_DELAY_SEC = DEBUG_DUMP_INTERVAL_SEC
SHEET_SERVICE_ACCOUNT = SHEET_SERVICE_ACCOUNT_PATH
SHEET_NAME = SHEET_NAME_DEFAULT
SHEET_TITLE = SHEET_TITLE_DEFAULT
SHEET_UPDATE_EVERY = SHEET_UPDATE_EVERY_DEFAULT
TARGET_COLUMNS = ("Divine Orb", "Exalted Orb")
FEEDBACK_HEADER_ROW = 6
FEEDBACK_FIRST_DATA_ROW = 7
FEEDBACK_START_COL = 4
FEEDBACK_END_COL = 5
OCR_WEIGHTS_PATH = os.path.join(WEB_STUFF_DIR, "ocr_weights.json")
AUTOMATION_SETTLE_SEC = 1
SELECT_SELL_TRADE = "select_sell"
SELECT_BUY_TRADE = "select_buy"
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



def init_overlay(ratios: Tuple[NamedRect, ...], on_escape=None) -> tk.Tk:
    root = tk.Tk()
    root.overrideredirect(True)
    root.attributes("-topmost", True)
    root.attributes("-transparentcolor", "magenta")
    root.configure(bg="magenta")
    if on_escape is not None:
        root.bind_all("<Escape>", on_escape)

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


def _cleanup_debug_dir(debug_dir: str) -> None:
    if not os.path.isdir(debug_dir):
        return
    for name in os.listdir(debug_dir):
        path = os.path.join(debug_dir, name)
        if not os.path.isfile(path):
            continue
        try:
            os.remove(path)
        except OSError as exc:
            print(f"Failed to remove {path}: {exc}")


def _safe_tag(value: str) -> str:
    out = []
    for ch in value:
        if ch.isalnum() or ch in ("_", "-"):
            out.append(ch)
        elif ch.isspace():
            out.append("_")
    return "".join(out) or "ratio"


def _start_windows_hotkey(stop_event: Event) -> Optional[tuple[threading.Thread, int]]:
    if os.name != "nt":
        return None
    user32 = ctypes.windll.user32
    kernel32 = ctypes.windll.kernel32
    MOD_CONTROL = 0x0002
    MOD_SHIFT = 0x0004
    VK_Q = 0x51
    WM_HOTKEY = 0x0312
    WM_QUIT = 0x0012

    thread_id_holder = {"id": 0}

    def _hotkey_loop():
        thread_id_holder["id"] = kernel32.GetCurrentThreadId()
        if not user32.RegisterHotKey(None, 1, MOD_CONTROL | MOD_SHIFT, VK_Q):
            err = ctypes.windll.kernel32.GetLastError()
            print(f"Failed to register hotkey Ctrl+Shift+Q (error {err})")
            return
        msg = ctypes.wintypes.MSG()
        while True:
            ret = user32.GetMessageW(ctypes.byref(msg), None, 0, 0)
            if ret == 0:
                break
            if msg.message == WM_HOTKEY:
                stop_event.set()
                break
        user32.UnregisterHotKey(None, 1)

    thread = threading.Thread(target=_hotkey_loop, name="hotkey-listener", daemon=True)
    thread.start()
    return thread, thread_id_holder["id"]


def _parse_ocr_log(path: str) -> dict:
    results = {}
    current_label = None
    if not os.path.exists(path):
        return results
    with open(path, "r", encoding="utf-8") as f:
        for line in f:
            line = line.rstrip("\n")
            if not line.strip():
                continue
            if not line.startswith("  "):
                parts = line.split(":", 1)
                if not parts:
                    continue
                current_label = parts[0].strip()
                results[current_label] = {}
                continue
            if current_label is None:
                continue
            stripped = line.strip()
            if stripped.startswith("raw=") or stripped.startswith("saved:"):
                continue
            if ": raw=" not in stripped:
                continue
            cand_label, rest = stripped.split(": raw=", 1)
            display_val = None
            if " display=" in rest:
                display_str = rest.split(" display=", 1)[1].strip()
                try:
                    display_val = float(display_str)
                except Exception:
                    display_val = None
            results[current_label][cand_label] = display_val
    return results


def _update_ocr_weights_from_feedback(
    log_path: str,
    row_labels: list[str],
    feedback_rows: list[list[str]],
) -> dict:
    weights = load_ocr_weights(OCR_WEIGHTS_PATH)
    candidates_by_label = _parse_ocr_log(log_path)
    for idx, row in enumerate(feedback_rows):
        label = row_labels[idx] if idx < len(row_labels) else None
        if not label:
            continue
        for col_idx, sell_name in enumerate(TARGET_COLUMNS):
            if col_idx >= len(row):
                continue
            raw_val = str(row[col_idx]).strip()
            if not raw_val:
                continue
            try:
                expected = float(raw_val)
            except Exception:
                continue
            pair_label = f"{sell_name} -> {label}"
            candidates = candidates_by_label.get(pair_label, {})
            if not candidates:
                continue
            best_label = None
            best_diff = None
            for cand_label, cand_val in candidates.items():
                if cand_val is None:
                    continue
                diff = abs(cand_val - expected)
                if best_diff is None or diff < best_diff:
                    best_diff = diff
                    best_label = cand_label
            if best_label is None:
                continue
            if best_diff is not None and best_diff <= 0.02:
                weights[best_label] = int(weights.get(best_label, 0)) + 1
    save_ocr_weights(OCR_WEIGHTS_PATH, weights)
    return weights


def _create_overlay(
    ratios: Tuple[NamedRect, ...],
    existing: Optional[tk.Tk],
    on_escape=None,
) -> Optional[tk.Tk]:
    try:
        new_overlay = init_overlay(ratios, on_escape=on_escape)
    except Exception as exc:
        print(f"Failed to init overlay: {exc}")
        return existing
    if existing is not None:
        try:
            existing.destroy()
        except Exception as exc:
            print(f"Failed to destroy overlay: {exc}")
    return new_overlay


def _load_or_init_matrix(rows: list[str], cols: list[str]) -> list[list[Optional[float]]]:
    return [[None for _ in range(len(cols))] for _ in range(len(rows))]


def _index_map(values: list[str]) -> dict[str, int]:
    return {str(name): idx for idx, name in enumerate(values)}


def _build_matrix_layout(currencies: list[str]) -> tuple[list[str], list[str]]:
    target_set = set(TARGET_COLUMNS)
    row_labels = [c for c in currencies if c not in target_set]
    col_labels = list(TARGET_COLUMNS)
    return row_labels, col_labels


def _build_pairs(row_labels: list[str], col_labels: list[str]) -> list[tuple[str, str]]:
    return [(row, col) for col in col_labels for row in row_labels]


def _reload_config_if_changed(
    project_path: str,
    last_mtime: float,
    overlay: Optional[tk.Tk],
    on_escape=None,
) -> tuple[float, Optional[tk.Tk], Optional[tuple]]:
    if not os.path.exists(project_path):
        return last_mtime, overlay, None
    try:
        mtime = os.path.getmtime(project_path)
    except OSError as exc:
        print(f"Failed to stat config: {exc}")
        return last_mtime, overlay, None
    if mtime == last_mtime:
        return last_mtime, overlay, None

    try:
        ratios = load_ratio_regions(project_path)
        triggers = load_triggers(project_path)
        currencies = load_currencies(project_path)
    except Exception as exc:
        print(f"Failed to reload config: {exc}")
        return last_mtime, overlay, None

    row_labels, col_labels = _build_matrix_layout(currencies)
    matrix = _load_or_init_matrix(row_labels, col_labels)
    row_map = _index_map(row_labels)
    col_map = _index_map(col_labels)
    pairs = _build_pairs(row_labels, col_labels)
    overlay = _create_overlay(ratios, overlay, on_escape=on_escape)
    state = (ratios, triggers, currencies, row_labels, col_labels, matrix, row_map, col_map, pairs)
    return mtime, overlay, state


def main():
    configure_tesseract()
    ratios = load_ratio_regions(PROJECT_PATH)
    triggers = load_triggers(PROJECT_PATH)
    currencies = load_currencies(PROJECT_PATH)
    row_labels, col_labels = _build_matrix_layout(currencies)
    matrix = _load_or_init_matrix(row_labels, col_labels)
    row_map = _index_map(row_labels)
    col_map = _index_map(col_labels)
    pairs = _build_pairs(row_labels, col_labels)
    pair_index = 0
    last_sell_item = None
    trade_runner = (
        TradeRunner(PROJECT_PATH, delay_sec=1.0)
        if os.path.exists(PROJECT_PATH)
        else None
    )
    trigger_state = {}
    last_config_mtime = os.path.getmtime(PROJECT_PATH) if os.path.exists(PROJECT_PATH) else 0.0
    sheet_id = load_sheet_id()
    sheet_ready = os.path.exists(SHEET_SERVICE_ACCOUNT)
    if sheet_ready:
        set_ocr_weights(load_ocr_weights(OCR_WEIGHTS_PATH))
    if sheet_ready and sheet_id:
        clear_sheet(
            sheet_id=sheet_id,
            sheet_name=SHEET_NAME,
            service_account_path=SHEET_SERVICE_ACCOUNT,
            oauth_client_path="",
        )
    stop_event = Event()

    def _on_escape(_event=None):
        stop_event.set()

    kb_hook = None
    hotkey_thread = None
    hotkey_thread_id = None
    if os.name == "nt":
        hotkey = _start_windows_hotkey(stop_event)
        if hotkey is not None:
            hotkey_thread, hotkey_thread_id = hotkey
        elif kb is not None:
            def _on_stop_hotkey():
                stop_event.set()

            kb_hook = kb.add_hotkey("ctrl+shift+q", _on_stop_hotkey, suppress=False)

    overlay = init_overlay(ratios, on_escape=_on_escape)
    if DEBUG_DUMP:
        os.makedirs(DEBUG_DIR, exist_ok=True)
        _cleanup_debug_dir(DEBUG_DIR)
    last_dump_ts = time.monotonic()
    dump_index = 0
    reading_count = 0
    debug_log = None
    if DEBUG_DUMP:
        debug_log = open(os.path.join(DEBUG_DIR, "ocr_log.txt"), "w", encoding="utf-8")
    loop_count = 0
    stopped_after_exalted = False
    try:
        with mss.mss() as sct:
            while True:
                if stop_event.is_set():
                    print("Stopped on hotkey.")
                    break
                loop_count += 1
                last_config_mtime, overlay, state = _reload_config_if_changed(
                    PROJECT_PATH,
                    last_config_mtime,
                    overlay,
                    on_escape=_on_escape,
                )
                if state is not None:
                    ratios, triggers, currencies, row_labels, col_labels, matrix, row_map, col_map, pairs = state
                    pair_index = 0

                row_item = None
                col_item = None
                if pairs:
                    row_item, col_item = pairs[pair_index]
                if pairs and last_sell_item == "Exalted Orb" and col_item != last_sell_item:
                    print("Stopped after finishing Exalted Orb prices.")
                    stopped_after_exalted = True
                    break
                if trade_runner is not None and pairs:
                    try:
                        if col_item != last_sell_item:
                            trade_runner.run_trade(
                                SELECT_SELL_TRADE,
                                overrides={"itemName": col_item},
                            )
                            last_sell_item = col_item
                        trade_runner.run_trade(
                            SELECT_BUY_TRADE,
                            overrides={"itemName": row_item},
                        )
                    except Exception as exc:
                        print(f"Trade automation failed: {exc}")
                    if AUTOMATION_SETTLE_SEC > 0:
                        if stop_event.wait(AUTOMATION_SETTLE_SEC):
                            print("Stopped on hotkey.")
                            break
                    pair_index = (pair_index + 1) % len(pairs)

                stop_after_debug = False
                for rr in ratios:
                    if stop_event.is_set():
                        print("Stopped on hotkey.")
                        stop_after_debug = True
                        break
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
                    effective_buy = rr.buy
                    effective_sell = rr.sell
                    if row_item and rr.buy.lower().startswith("market"):
                        effective_buy = row_item
                    if col_item and rr.sell.lower().startswith("market"):
                        effective_sell = col_item
                    label = rr.label
                    if effective_buy or effective_sell:
                        label = f"{effective_sell} -> {effective_buy}".strip()
                    if display is None:
                        print(f"{label}: ?")
                    else:
                        print(f"{label}: {format_decimals(display, 2)}")
                        row = row_map.get(effective_buy)
                        col = col_map.get(effective_sell)
                        if row is not None and col is not None:
                            if 0 <= row < len(matrix) and 0 <= col < len(matrix[row]):
                                matrix[row][col] = display
                                if sheet_ready:
                                    sheet_id = export_matrix_to_sheet(
                                        col_labels,
                                        matrix,
                                        oauth_client_path="",
                                        service_account_path=SHEET_SERVICE_ACCOUNT,
                                        sheet_id=sheet_id,
                                        sheet_name=SHEET_NAME,
                                        sheet_title=SHEET_TITLE,
                                        apply_format=False,
                                        row_labels=row_labels,
                                    )
                    if debug_log is not None:
                        display_text = "?" if display is None else format_decimals(display, 2)
                        debug_log.write(f"{label}: {display_text}\n")
                        debug_log.write(
                            f"  raw='{raw}' ratio={ratio} display={display}\n"
                        )
                        for c_label, c_raw, c_ratio in candidates:
                            c_display = None
                            if c_ratio is not None:
                                c_display = compute_display(c_raw, c_ratio, rr)
                            debug_log.write(
                                f"  {c_label}: raw='{c_raw}' ratio={c_ratio} display={c_display}\n"
                            )
                        debug_log.flush()
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
                    if DEBUG_DUMP and (
                        DEBUG_DUMP_MAX_READINGS is None
                        or reading_count < DEBUG_DUMP_MAX_READINGS
                    ):
                        tag_base = f"{effective_sell}-{effective_buy}".strip("-")
                        file_tag = _safe_tag(tag_base) if tag_base else rr.file_tag
                        raw_path = os.path.join(
                            DEBUG_DIR, f"{file_tag}_{reading_count}_raw.png"
                        )
                        bin_path = os.path.join(
                            DEBUG_DIR, f"{file_tag}_{reading_count}_bin.png"
                        )
                        gray_path = os.path.join(
                            DEBUG_DIR, f"{file_tag}_{reading_count}_gray.png"
                        )
                        bin_img = prep_for_ocr(img, scale=rr.scale)
                        gray_img = preps[0][1](img)
                        cv2.imwrite(raw_path, img)
                        cv2.imwrite(bin_path, bin_img)
                        cv2.imwrite(gray_path, gray_img)
                        if debug_log is not None:
                            debug_log.write(
                                f"  saved: {raw_path}, {bin_path}, {gray_path}\n"
                            )
                            debug_log.flush()
                        reading_count += 1
                        if (
                            DEBUG_DUMP_MAX_READINGS is not None
                            and reading_count >= DEBUG_DUMP_MAX_READINGS
                        ):
                            stop_after_debug = True
                            break
                if sheet_ready and loop_count % SHEET_UPDATE_EVERY == 0:
                    sheet_id = export_matrix_to_sheet(
                        col_labels,
                        matrix,
                        oauth_client_path="",
                        service_account_path=SHEET_SERVICE_ACCOUNT,
                        sheet_id=sheet_id,
                        sheet_name=SHEET_NAME,
                        sheet_title=SHEET_TITLE,
                        apply_format=False,
                        row_labels=row_labels,
                    )
                if overlay is not None:
                    overlay.update_idletasks()
                    overlay.update()
                if stop_after_debug:
                    print(f"Stopped after {DEBUG_DUMP_MAX_READINGS} OCR readings.")
                    break
                if stop_event.is_set():
                    print("Stopped on hotkey.")
                    break
                if DEBUG_DUMP and (time.monotonic() - last_dump_ts >= DEBUG_DUMP_INTERVAL_SEC):
                    last_dump_ts = time.monotonic()
                    dump_index += 1
                sleep_sec = 0.0
                if trade_runner is None or not pairs:
                    sleep_sec = LOOP_DELAY_SEC
                if sleep_sec > 0:
                    if stop_event.wait(sleep_sec):
                        print("Stopped on hotkey.")
                        break
    finally:
        if debug_log is not None:
            debug_log.close()
        if kb is not None and kb_hook is not None:
            kb.remove_hotkey(kb_hook)
        if os.name == "nt" and hotkey_thread_id:
            ctypes.windll.user32.PostThreadMessageW(hotkey_thread_id, 0x0012, 0, 0)
    if stopped_after_exalted and sheet_ready and sheet_id:
        feedback_rows = read_feedback_columns(
            sheet_id=sheet_id,
            sheet_name=SHEET_NAME,
            service_account_path=SHEET_SERVICE_ACCOUNT,
            oauth_client_path="",
            header_row=FEEDBACK_HEADER_ROW,
            first_data_row=FEEDBACK_FIRST_DATA_ROW,
            start_col=FEEDBACK_START_COL,
            end_col=FEEDBACK_END_COL,
        )
        print("Feedback columns (D:E):")
        for idx, row in enumerate(feedback_rows):
            label = row_labels[idx] if idx < len(row_labels) else f"Row {idx + 1}"
            if not row:
                continue
            divine = row[0] if len(row) > 0 else ""
            exalted = row[1] if len(row) > 1 else ""
            if divine or exalted:
                print(f"{label}: Divine={divine} Exalted={exalted}")
        _update_ocr_weights_from_feedback(
            os.path.join(DEBUG_DIR, "ocr_log.txt"),
            row_labels,
            feedback_rows,
        )


if __name__ == "__main__":
    import sys

    if "--help" in sys.argv or "-h" in sys.argv:
        print(HELP_TEXT)
        raise SystemExit(0)
    main()
