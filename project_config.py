import json
import os
from dataclasses import dataclass
from typing import Dict, List, Optional, Tuple


BASE_DIR = os.path.dirname(__file__)
PROJECT_PATH = os.path.join(BASE_DIR, "project.json")
WEB_STUFF_DIR = os.path.join(BASE_DIR, "web_stuff")
DEBUG_DIR = os.path.join(BASE_DIR, "debug_dumps")
DEBUG_DUMP_DEFAULT = False
DEBUG_DUMP_INTERVAL_SEC_DEFAULT = 1.0
SHEET_SERVICE_ACCOUNT_PATH = os.path.join(WEB_STUFF_DIR, "poe2trader-2e14fc353067.json")
SHEET_NAME_DEFAULT = "Sheet1"
SHEET_TITLE_DEFAULT = "Price Matrix"
SHEET_UPDATE_EVERY_DEFAULT = 5


@dataclass(frozen=True)
class Rect:
    x: int
    y: int
    w: int
    h: int

    def as_mss(self) -> Dict[str, int]:
        return {"left": self.x, "top": self.y, "width": self.w, "height": self.h}


@dataclass(frozen=True)
class NamedRect:
    buy: str
    sell: str
    key: str
    label: str
    file_tag: str
    rect: Rect
    ocr_mode: str
    scale: float
    decimal_rule: str
    expected_min: Optional[float]
    expected_max: Optional[float]
    expected_integer_only: bool
    debug_output: bool
    clean_min_area_ratio: float
    clean_left_margin_ratio: float
    clean_left_max_area_ratio: float
    clean_left_max_width: int


OCR_MODE_DEFAULT = "multi"
SCALE_DEFAULT = 3.0
DECIMAL_RULE_DEFAULT = ""
EXPECTED_MIN_DEFAULT = None
EXPECTED_MAX_DEFAULT = None
EXPECTED_INTEGER_ONLY_DEFAULT = False
DEBUG_OUTPUT_DEFAULT = True
CLEAN_MIN_AREA_RATIO_DEFAULT = 0.0007
CLEAN_LEFT_MARGIN_RATIO_DEFAULT = 0.20
CLEAN_LEFT_MAX_AREA_RATIO_DEFAULT = 0.0020
CLEAN_LEFT_MAX_WIDTH_DEFAULT = 2

DEFAULTS = {
    "ocr_mode": OCR_MODE_DEFAULT,
    "scale": SCALE_DEFAULT,
    "decimal_rule": DECIMAL_RULE_DEFAULT,
    "expected_min": EXPECTED_MIN_DEFAULT,
    "expected_max": EXPECTED_MAX_DEFAULT,
    "expected_integer_only": EXPECTED_INTEGER_ONLY_DEFAULT,
    "debug_output": DEBUG_OUTPUT_DEFAULT,
    "clean_min_area_ratio": CLEAN_MIN_AREA_RATIO_DEFAULT,
    "clean_left_margin_ratio": CLEAN_LEFT_MARGIN_RATIO_DEFAULT,
    "clean_left_max_area_ratio": CLEAN_LEFT_MAX_AREA_RATIO_DEFAULT,
    "clean_left_max_width": CLEAN_LEFT_MAX_WIDTH_DEFAULT,
}


def load_project(path: str = PROJECT_PATH) -> dict:
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def _safe_tag(value: str) -> str:
    out = []
    for ch in value:
        if ch.isalnum() or ch in ("_", "-"):
            out.append(ch)
        elif ch.isspace():
            out.append("_")
    return "".join(out) or "ratio"


def load_ratio_regions(path: str = PROJECT_PATH) -> Tuple[NamedRect, ...]:
    cfg = load_project(path)
    regions = []
    for r in cfg.get("ratios", []):
        buy = str(r.get("buy", "")).strip()
        sell = str(r.get("sell", "")).strip()
        key = f"{buy}->{sell}".strip()
        label = key
        if buy or sell:
            label = f"{buy} -> {sell}".strip()
        file_tag = _safe_tag(key)
        regions.append(
            NamedRect(
                buy=buy,
                sell=sell,
                key=key,
                label=label,
                file_tag=file_tag,
                rect=Rect(int(r["x"]), int(r["y"]), int(r["w"]), int(r["h"])),
                ocr_mode=str(r.get("ocr_mode", DEFAULTS["ocr_mode"])),
                scale=float(r.get("scale", DEFAULTS["scale"])),
                decimal_rule=str(r.get("decimal_rule", DEFAULTS["decimal_rule"])),
                expected_min=(
                    float(r["expected_min"]) if "expected_min" in r else DEFAULTS["expected_min"]
                ),
                expected_max=(
                    float(r["expected_max"]) if "expected_max" in r else DEFAULTS["expected_max"]
                ),
                expected_integer_only=bool(
                    r.get("expected_integer_only", DEFAULTS["expected_integer_only"])
                ),
                debug_output=bool(r.get("debug_output", DEFAULTS["debug_output"])),
                clean_min_area_ratio=float(
                    r.get("clean_min_area_ratio", DEFAULTS["clean_min_area_ratio"])
                ),
                clean_left_margin_ratio=float(
                    r.get("clean_left_margin_ratio", DEFAULTS["clean_left_margin_ratio"])
                ),
                clean_left_max_area_ratio=float(
                    r.get("clean_left_max_area_ratio", DEFAULTS["clean_left_max_area_ratio"])
                ),
                clean_left_max_width=int(
                    r.get("clean_left_max_width", DEFAULTS["clean_left_max_width"])
                ),
            )
        )
    return tuple(regions)


def load_triggers(path: str = PROJECT_PATH) -> List[dict]:
    cfg = load_project(path)
    out: List[dict] = []
    for t in cfg.get("triggers", []):
        if isinstance(t, dict):
            out.append(t)
    return out


def load_currencies(path: str = PROJECT_PATH) -> List[str]:
    cfg = load_project(path)
    return [str(c) for c in cfg.get("currencies", [])]
