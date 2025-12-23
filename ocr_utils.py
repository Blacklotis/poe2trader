import json
import re
from typing import Optional, Tuple

import cv2
import numpy as np
import pytesseract


def prep_for_ocr(img_bgr: np.ndarray, scale: float = 3.0) -> np.ndarray:
    g = cv2.cvtColor(img_bgr, cv2.COLOR_BGR2GRAY)
    if scale != 1.0:
        g = cv2.resize(g, None, fx=scale, fy=scale, interpolation=cv2.INTER_CUBIC)
    g = cv2.GaussianBlur(g, (3, 3), 0)
    _, g = cv2.threshold(g, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
    g = cv2.medianBlur(g, 3)
    return g


def prep_for_ocr_no_denoise(img_bgr: np.ndarray, scale: float = 3.0) -> np.ndarray:
    g = cv2.cvtColor(img_bgr, cv2.COLOR_BGR2GRAY)
    if scale != 1.0:
        g = cv2.resize(g, None, fx=scale, fy=scale, interpolation=cv2.INTER_CUBIC)
    g = cv2.GaussianBlur(g, (3, 3), 0)
    _, g = cv2.threshold(g, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
    return g


def clean_binary(
    bin_img: np.ndarray,
    min_area_ratio: float,
    left_margin_ratio: float,
    left_max_area_ratio: float,
    left_max_width: int,
) -> np.ndarray:
    h, w = bin_img.shape[:2]
    inv = bin_img
    if float(bin_img.mean()) > 127.0:
        inv = 255 - bin_img

    n, _, stats, _ = cv2.connectedComponentsWithStats(inv, connectivity=8)
    if n <= 1:
        return bin_img

    min_area = max(8, int((h * w) * min_area_ratio))
    left_margin = int(w * left_margin_ratio)
    left_max_area = int((h * w) * left_max_area_ratio)

    keep = np.zeros_like(inv)
    for i in range(1, n):
        x, y, bw, bh, area = stats[i]
        if area < min_area:
            continue
        if x + bw <= left_margin and (area < left_max_area or bw <= left_max_width):
            continue
        keep[y : y + bh, x : x + bw] = inv[y : y + bh, x : x + bw]

    if inv is not bin_img:
        return 255 - keep
    return keep


def prep_for_ocr_clean(
    img_bgr: np.ndarray,
    scale: float,
    min_area_ratio: float,
    left_margin_ratio: float,
    left_max_area_ratio: float,
    left_max_width: int,
) -> np.ndarray:
    g = prep_for_ocr_no_denoise(img_bgr, scale=scale)
    return clean_binary(g, min_area_ratio, left_margin_ratio, left_max_area_ratio, left_max_width)


def prep_gray(img_bgr: np.ndarray, scale: float = 3.0) -> np.ndarray:
    g = cv2.cvtColor(img_bgr, cv2.COLOR_BGR2GRAY)
    if scale != 1.0:
        g = cv2.resize(g, None, fx=scale, fy=scale, interpolation=cv2.INTER_CUBIC)
    g = cv2.GaussianBlur(g, (3, 3), 0)
    return g


def prep_gray_no_blur(img_bgr: np.ndarray, scale: float = 3.0) -> np.ndarray:
    g = cv2.cvtColor(img_bgr, cv2.COLOR_BGR2GRAY)
    if scale != 1.0:
        g = cv2.resize(g, None, fx=scale, fy=scale, interpolation=cv2.INTER_CUBIC)
    return g


OCR_WEIGHTS = {}


def set_ocr_weights(weights: dict) -> None:
    global OCR_WEIGHTS
    OCR_WEIGHTS = dict(weights or {})


def load_ocr_weights(path: str) -> dict:
    if not path:
        return {}
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        return dict(data) if isinstance(data, dict) else {}
    except FileNotFoundError:
        return {}
    except Exception:
        return {}


def save_ocr_weights(path: str, weights: dict) -> None:
    if not path:
        return
    with open(path, "w", encoding="utf-8") as f:
        json.dump(weights, f, indent=2, sort_keys=True)


def parse_ratio(text: str) -> Optional[float]:
    s = (text or "").replace(";", ":").replace(",", ".")
    if not s:
        return None
    match = re.search(r"([0-9]+(?:\.[0-9]+)?)\s*[:/xX]\s*([0-9]+(?:\.[0-9]+)?)", s)
    if match:
        try:
            return float(match.group(2))
        except Exception:
            return None
    filtered = []
    dot_used = False
    for ch in s.replace(" ", ""):
        if ch.isdigit():
            filtered.append(ch)
            continue
        if ch == "." and not dot_used:
            filtered.append(ch)
            dot_used = True
    cleaned = "".join(filtered)
    if not cleaned:
        return None
    try:
        return float(cleaned)
    except Exception:
        return None



def format_decimals(value: float, places: int = 2) -> str:
    return f"{value:.{places}f}"


def apply_decimal_rule(raw: str, value: float, rule: str) -> float:
    if rule != "tail_zero_two_dp":
        return value
    raw = raw or ""
    dot_pos = raw.find(".")
    if dot_pos >= 0:
        tail = raw[dot_pos + 1 :]
        if any(ch.isdigit() for ch in tail):
            return value
    digits = "".join(ch for ch in (raw or "") if ch.isdigit())
    if len(digits) < 2:
        return value
    if len(digits) >= 3 and digits[-1] == "0":
        insert = len(digits) - 2
    else:
        insert = len(digits) - 1
    cooked = digits[:insert] + "." + digits[insert:]
    try:
        return float(cooked)
    except Exception:
        return value


def coerce_ratio_merged_one(raw: str, min_v: float, max_v: float) -> Optional[float]:
    digits = "".join(ch for ch in (raw or "") if ch.isdigit())
    if len(digits) < 3:
        return None
    candidates = []
    if digits.endswith("1"):
        candidates.append(digits[:-1])
    if digits.startswith("1"):
        candidates.append(digits[1:])
    for candidate in candidates:
        if not candidate:
            continue
        try:
            value = float(candidate)
        except Exception:
            continue
        if min_v <= value <= max_v:
            return value
    return None


def coerce_full_number(raw: str, value: float) -> float:
    if value >= 1.0:
        return value
    digits = "".join(ch for ch in (raw or "") if ch.isdigit())
    if not digits:
        return value
    try:
        return float(digits)
    except Exception:
        return value


def apply_expected_range(
    raw: str,
    value: float,
    min_v: Optional[float],
    max_v: Optional[float],
    integer_only: bool,
) -> Optional[float]:
    if min_v is None or max_v is None:
        return value
    if min_v <= value <= max_v:
        return value
    digits = "".join(ch for ch in (raw or "") if ch.isdigit())
    if len(digits) < 2:
        return None
    if integer_only:
        for length in range(1, len(digits) + 1):
            for start in range(0, len(digits) - length + 1):
                chunk = digits[start : start + length]
                try:
                    v = float(chunk)
                except Exception:
                    continue
                if min_v <= v <= max_v:
                    return v
        return None
    best = None
    for i in range(1, len(digits)):
        candidate = digits[:i] + "." + digits[i:]
        try:
            v = float(candidate)
        except Exception:
            continue
        if min_v <= v <= max_v:
            best = v
            break
    return best


def _within_range(value: float, min_v: Optional[float], max_v: Optional[float]) -> bool:
    if min_v is None or max_v is None:
        return True
    return min_v <= value <= max_v


def _drop_leading_one(text: str) -> Optional[float]:
    if not text or not text.startswith("1"):
        return None
    candidate = text[1:].strip()
    if not candidate:
        return None
    if candidate.startswith("."):
        candidate = "0" + candidate
    try:
        return float(candidate)
    except Exception:
        return None


def _swap_two_digits(text: str) -> Optional[float]:
    digits = "".join(ch for ch in (text or "") if ch.isdigit())
    if len(digits) != 2:
        return None
    swapped = digits[1] + digits[0]
    try:
        return float(swapped)
    except Exception:
        return None


def _split_ratio_raw(raw: str) -> tuple[str, Optional[str]]:
    if not raw:
        return raw, None
    match = re.search(r"([0-9][0-9\.,]*)\s*[:/xX]\s*([0-9][0-9\.,]*)", raw)
    if match:
        lhs_raw = match.group(1)
        rhs_raw = match.group(2)
        lhs_value = parse_ratio(lhs_raw)
        rhs_value = parse_ratio(rhs_raw)
        if rhs_value == 1 and lhs_value is not None:
            return raw, lhs_raw
        return raw, rhs_raw
    match = re.search(r"1\s+([0-9][0-9\.,]*)", raw)
    if match:
        return raw, match.group(1)
    if ":" not in raw:
        return raw, None
    parts = raw.split(":")
    if len(parts) < 2:
        return raw, None
    rhs = ":".join(parts[1:])
    return raw, rhs.strip()


def compute_display(raw: str, value: float, rr) -> Optional[float]:
    raw_full, rhs_raw = _split_ratio_raw(raw)
    raw_for_rules = rhs_raw or raw_full
    base_value = value
    if rhs_raw:
        rhs_value = parse_ratio(rhs_raw)
        if rhs_value is not None:
            base_value = rhs_value
        lhs_raw = None
        match = re.search(r"([0-9][0-9\.,]*)\s*[:/xX]\s*([0-9][0-9\.,]*)", raw_full)
        if match:
            lhs_raw = match.group(1)
            rhs_match = match.group(2)
            rhs_match_value = parse_ratio(rhs_match)
            if rhs_match_value == 1 and lhs_raw:
                dropped = _drop_leading_one(lhs_raw)
                swapped = _swap_two_digits(lhs_raw)
                candidates = []
                if dropped is not None and _within_range(dropped, rr.expected_min, rr.expected_max):
                    candidates.append(dropped)
                if swapped is not None and _within_range(swapped, rr.expected_min, rr.expected_max):
                    candidates.append(swapped)
                if candidates:
                    if rr.expected_min is not None and rr.expected_max is not None:
                        mid = (rr.expected_min + rr.expected_max) / 2.0
                        candidates.sort(key=lambda v: abs(v - mid))
                    return candidates[0]
    display = coerce_full_number(raw_for_rules, base_value)
    if (
        rr.expected_min is not None
        and rr.expected_max is not None
        and rr.expected_min <= display <= rr.expected_max
    ):
        return display
    if (
        rr.expected_min is not None
        and rr.expected_max is not None
        and not any(ch in (raw_for_rules or "") for ch in (":", "/", "x", "X", "."))
    ):
        merged = coerce_ratio_merged_one(raw_for_rules, rr.expected_min, rr.expected_max)
        if merged is not None:
            return merged
    display = apply_decimal_rule(raw_for_rules, display, rr.decimal_rule)
    return apply_expected_range(
        raw_for_rules,
        display,
        rr.expected_min,
        rr.expected_max,
        rr.expected_integer_only,
    )


def ocr_candidates(
    img_bgr: np.ndarray,
    preps: Tuple[Tuple[str, callable], ...],
) -> Tuple[Tuple[str, str, Optional[float]], ...]:
    cfgs = (
        (
            "psm7",
            "--oem 1 --psm 7 -c tessedit_char_whitelist=0123456789.:/xX "
            "-c classify_bln_numeric_mode=1 -c user_defined_dpi=300",
        ),
        (
            "psm8",
            "--oem 1 --psm 8 -c tessedit_char_whitelist=0123456789.:/xX "
            "-c classify_bln_numeric_mode=1 -c user_defined_dpi=300",
        ),
        (
            "psm6",
            "--oem 1 --psm 6 -c tessedit_char_whitelist=0123456789.:/xX "
            "-c classify_bln_numeric_mode=1 -c user_defined_dpi=300",
        ),
        (
            "psm13",
            "--oem 1 --psm 13 -c tessedit_char_whitelist=0123456789.:/xX "
            "-c classify_bln_numeric_mode=1 -c user_defined_dpi=300",
        ),
    )
    candidates = []
    for label, prep in preps:
        g = prep(img_bgr)
        for cfg_label, cfg in cfgs:
            raw = pytesseract.image_to_string(g, config=cfg) or ""
            ratio = parse_ratio(raw)
            candidates.append((f"{label}:{cfg_label}", raw.strip(), ratio))
    return tuple(candidates)


def read_ratio_from_image(
    img_bgr: np.ndarray,
    preps: Tuple[Tuple[str, callable], ...],
) -> Tuple[Optional[float], str, Tuple[Tuple[str, str, Optional[float]], ...]]:
    candidates = ocr_candidates(img_bgr, preps)
    best_ratio = None
    best_raw = ""
    best_score = (-1, -1, -1, -1, -1)
    for _, candidate, ratio in candidates:
        if ratio is None:
            continue
        digits = "".join(ch for ch in candidate if ch.isdigit())
        has_dot = 1 if "." in candidate else 0
        has_sep = 1 if any(ch in candidate for ch in (":", "/", "x", "X")) else 0
        dot_pos = candidate.find(".")
        digits_after_dot = 0
        if dot_pos >= 0:
            tail = candidate[dot_pos + 1 :]
            digits_after_dot = sum(1 for ch in tail if ch.isdigit())
        dot_score = 1 if (has_dot and digits_after_dot > 0) else 0
        weight = OCR_WEIGHTS.get(_, 0)
        score = (weight, has_sep, dot_score, digits_after_dot, -len(digits), -len(candidate.strip()))
        if score > best_score:
            best_score = score
            best_ratio = ratio
            best_raw = candidate.strip()
    if best_ratio is not None:
        return best_ratio, best_raw, candidates
    return None, "", candidates


def get_ocr_preps(rr) -> Tuple[Tuple[str, callable], ...]:
    def gray():
        return lambda img, scale=rr.scale: prep_gray(img, scale=scale)

    def gray_no_blur():
        return lambda img, scale=rr.scale: prep_gray_no_blur(img, scale=scale)

    def bin_no_denoise():
        return lambda img, scale=rr.scale: prep_for_ocr_no_denoise(img, scale=scale)

    def clean():
        return lambda img, scale=rr.scale: prep_for_ocr_clean(
            img,
            scale=scale,
            min_area_ratio=rr.clean_min_area_ratio,
            left_margin_ratio=rr.clean_left_margin_ratio,
            left_max_area_ratio=rr.clean_left_max_area_ratio,
            left_max_width=rr.clean_left_max_width,
        )

    def bin_denoise():
        return lambda img, scale=rr.scale: prep_for_ocr(img, scale=scale)

    if rr.ocr_mode == "gray_only":
        return (("gray", gray()), ("gray_no_blur", gray_no_blur()))
    return (
        ("gray", gray()),
        ("bin", bin_no_denoise()),
        ("clean", clean()),
        ("bin_denoise", bin_denoise()),
    )
