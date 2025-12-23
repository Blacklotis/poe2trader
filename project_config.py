import json
import os
from dataclasses import dataclass
from typing import Dict, List, Optional, Tuple


BASE_DIR = os.path.dirname(__file__)
PROJECT_PATH = os.path.join(BASE_DIR, "project.json")
WEB_STUFF_DIR = os.path.join(BASE_DIR, "web_stuff")


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
    name: str
    rect: Rect
    position: Optional[Tuple[int, int]]
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


DEFAULTS = {
    "position": None,
    "ocr_mode": "multi",
    "scale": 3.0,
    "decimal_rule": "",
    "expected_min": None,
    "expected_max": None,
    "expected_integer_only": False,
    "debug_output": True,
    "clean_min_area_ratio": 0.0007,
    "clean_left_margin_ratio": 0.20,
    "clean_left_max_area_ratio": 0.0020,
    "clean_left_max_width": 2,
}


def load_project(path: str = PROJECT_PATH) -> dict:
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def load_ratio_regions(path: str = PROJECT_PATH) -> Tuple[NamedRect, ...]:
    cfg = load_project(path)
    regions = []
    for r in cfg.get("ratios", []):
        regions.append(
            NamedRect(
                name=str(r["name"]),
                rect=Rect(int(r["x"]), int(r["y"]), int(r["w"]), int(r["h"])),
                position=(
                    (int(r["position"][0]), int(r["position"][1]))
                    if "position" in r
                    else DEFAULTS["position"]
                ),
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
