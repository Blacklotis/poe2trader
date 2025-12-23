import json
import time
from dataclasses import dataclass
from typing import Dict, Optional, Tuple

import os
import cv2
import mss
import numpy as np
import pytesseract
import tkinter as tk


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
    ocr_mode: str
    scale: float
    decimal_rule: str
    use_templates: bool
    clean_min_area_ratio: float
    clean_left_margin_ratio: float
    clean_left_max_area_ratio: float
    clean_left_max_width: int

CONFIG_PATH = os.path.join(os.path.dirname(__file__), "config.json")
REGION = Rect(x=1630, y=500, w=470, h=140)
TEMPLATE_DIR = os.path.join(os.path.dirname(__file__), "characters")
TEMPLATE_SCALE = 2.0
TEMPLATE_INVERT = True
MATCH_THRESHOLD = 0.55
MIN_COMPONENT_AREA_RATIO = 0.01
DEBUG_DUMP = True
DEBUG_EVERY_N = 5
DEBUG_DIR = os.path.join(os.path.dirname(__file__), "debug_dumps")
DEFAULTS = {
    "ocr_mode": "multi",
    "scale": 3.0,
    "decimal_rule": "",
    "use_templates": False,
    "clean_min_area_ratio": 0.0007,
    "clean_left_margin_ratio": 0.20,
    "clean_left_max_area_ratio": 0.0020,
    "clean_left_max_width": 2,
}


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


def load_ratio_regions(path: str) -> Tuple[NamedRect, ...]:
    if not os.path.exists(path):
        return tuple()
    with open(path, "r", encoding="utf-8") as f:
        cfg = json.load(f)
    regions = []
    for r in cfg.get("ratios", []):
        regions.append(
            NamedRect(
                name=str(r["name"]),
                rect=Rect(int(r["x"]), int(r["y"]), int(r["w"]), int(r["h"])),
                ocr_mode=str(r.get("ocr_mode", DEFAULTS["ocr_mode"])),
                scale=float(r.get("scale", DEFAULTS["scale"])),
                decimal_rule=str(r.get("decimal_rule", DEFAULTS["decimal_rule"])),
                use_templates=bool(r.get("use_templates", DEFAULTS["use_templates"])),
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

    n, labels, stats, _ = cv2.connectedComponentsWithStats(inv, connectivity=8)
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


def parse_ratio(text: str) -> Optional[float]:
    s = (text or "").replace(" ", "").replace(";", ":").replace(",", ".")
    if not s:
        return None
    filtered = []
    dot_used = False
    for ch in s:
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
    if "." in (raw or ""):
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


def ocr_candidates(
    img_bgr: np.ndarray,
    preps: Tuple[Tuple[str, callable], ...],
) -> Tuple[Tuple[str, str, Optional[float]], ...]:
    cfg = (
        "--oem 1 --psm 8 -c tessedit_char_whitelist=0123456789. "
        "-c classify_bln_numeric_mode=1 -c user_defined_dpi=300"
    )
    candidates = []
    for label, prep in preps:
        g = prep(img_bgr)
        raw = pytesseract.image_to_string(g, config=cfg) or ""
        ratio = parse_ratio(raw)
        candidates.append((label, raw.strip(), ratio))
    return tuple(candidates)


def read_ratio_from_image(
    img_bgr: np.ndarray,
    preps: Tuple[Tuple[str, callable], ...],
    use_templates: bool,
) -> Tuple[Optional[float], str, Tuple[Tuple[str, str, Optional[float]], ...]]:
    raw = ""
    if use_templates and TEMPLATE_READER is not None:
        raw = TEMPLATE_READER.read_text(img_bgr)
    if raw:
        ratio = parse_ratio(raw)
        return ratio, raw.strip(), tuple()

    candidates = ocr_candidates(img_bgr, preps)
    best_ratio = None
    best_raw = ""
    best_score = (-1, -1, -1)
    for _, candidate, ratio in candidates:
        if ratio is None:
            continue
        digits = "".join(ch for ch in candidate if ch.isdigit())
        has_dot = 1 if "." in candidate else 0
        dot_pos = candidate.find(".")
        digits_after_dot = 0
        if dot_pos >= 0:
            tail = candidate[dot_pos + 1 :]
            digits_after_dot = sum(1 for ch in tail if ch.isdigit())
        score = (has_dot, digits_after_dot, len(digits), len(candidate.strip()))
        if score > best_score:
            best_score = score
            best_ratio = ratio
            best_raw = candidate.strip()
    if best_ratio is not None:
        return best_ratio, best_raw, candidates
    ratio = parse_ratio(raw)
    return ratio, raw.strip(), candidates


def get_ocr_preps(rr: NamedRect) -> Tuple[Tuple[str, callable], ...]:
    def gray():
        return lambda img, scale=rr.scale: prep_gray(img, scale=scale)

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
        return (("gray", gray()),)
    return (
        ("gray", gray()),
        ("bin", bin_no_denoise()),
        ("clean", clean()),
        ("bin_denoise", bin_denoise()),
    )


class FixedFontTemplateReader:
    def __init__(
        self,
        glyph_templates: Dict[str, np.ndarray],
        scale: float = TEMPLATE_SCALE,
        invert: bool = TEMPLATE_INVERT,
        match_threshold: float = MATCH_THRESHOLD,
    ):
        self.scale = float(scale)
        self.invert = bool(invert)
        self.match_threshold = float(match_threshold)
        prepped = {k: self._prep_template(v) for k, v in glyph_templates.items()}
        max_w = max(t.shape[1] for t in prepped.values())
        max_h = max(t.shape[0] for t in prepped.values())
        self.gw, self.gh = int(max_w), int(max_h)
        self.templates = {
            k: cv2.resize(t, (self.gw, self.gh), interpolation=cv2.INTER_NEAREST)
            for k, t in prepped.items()
        }

    def read_text(self, img_bgr: np.ndarray) -> str:
        bin_img = self._prep_crop(img_bgr)
        boxes = self._find_glyph_boxes(bin_img)
        if not boxes:
            return ""
        out = []
        for (x, y, w, h) in boxes:
            glyph = bin_img[y : y + h, x : x + w]
            ch = self._match_glyph(glyph)
            if ch is not None:
                out.append(ch)
        return "".join(out)

    def _prep_crop(self, bgr: np.ndarray) -> np.ndarray:
        g = cv2.cvtColor(bgr, cv2.COLOR_BGR2GRAY) if bgr.ndim == 3 else bgr.copy()
        if self.scale != 1.0:
            g = cv2.resize(g, None, fx=self.scale, fy=self.scale, interpolation=cv2.INTER_CUBIC)
        g = cv2.GaussianBlur(g, (3, 3), 0)
        _, g = cv2.threshold(g, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
        if self.invert:
            g = 255 - g
        # Remove tiny artifacts by filtering small connected components.
        h, w = g.shape[:2]
        inv = 255 - g
        n, _, stats, _ = cv2.connectedComponentsWithStats(inv, connectivity=8)
        if n > 1:
            min_area = max(10, int((h * w) * MIN_COMPONENT_AREA_RATIO))
            mask = np.zeros_like(inv)
            for i in range(1, n):
                x, y, bw, bh, area = stats[i]
                if area >= min_area:
                    mask[y : y + bh, x : x + bw] = inv[y : y + bh, x : x + bw]
            g = 255 - mask
        return g

    def _prep_template(self, img: np.ndarray) -> np.ndarray:
        g = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY) if img.ndim == 3 else img.copy()
        _, g = cv2.threshold(g, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
        if self.invert:
            g = 255 - g
        return g

    def _find_glyph_boxes(self, bin_img: np.ndarray) -> list:
        h, w = bin_img.shape[:2]
        inv = 255 - bin_img
        n, _, stats, _ = cv2.connectedComponentsWithStats(inv, connectivity=8)
        boxes = []
        if n <= 1:
            return boxes

        min_area = max(6, int((h * w) * 0.0002))
        max_area = int((h * w) * 0.20)

        for i in range(1, n):
            x, y, bw, bh, area = stats[i]
            if area < min_area or area > max_area:
                continue
            if bh < max(4, int(h * 0.20)) or bh > int(h * 0.95):
                continue
            if bw < max(2, int(w * 0.01)) or bw > int(w * 0.40):
                continue
            boxes.append((int(x), int(y), int(bw), int(bh)))

        boxes.sort(key=lambda b: b[0])

        merged = []
        for b in boxes:
            if not merged:
                merged.append(b)
                continue
            x, y, bw, bh = b
            px, py, pbw, pbh = merged[-1]
            if x <= px + pbw + 2:
                nx = px
                ny = min(py, y)
                nr = max(px + pbw, x + bw)
                nb = max(py + pbh, y + bh)
                merged[-1] = (nx, ny, nr - nx, nb - ny)
            else:
                merged.append(b)
        return merged

    def _match_glyph(self, glyph_bin: np.ndarray) -> Optional[str]:
        g = self._tight_crop(glyph_bin)
        if g.size == 0:
            return None
        g = cv2.resize(g, (self.gw, self.gh), interpolation=cv2.INTER_NEAREST)

        best_ch = None
        best_score = float("-inf")

        for ch, t in self.templates.items():
            score = self._ncc(g, t)
            if score > best_score:
                best_score = score
                best_ch = ch

        if best_ch is None or best_score < self.match_threshold:
            return None
        return best_ch

    def _tight_crop(self, bin_img: np.ndarray) -> np.ndarray:
        inv = 255 - bin_img
        ys, xs = np.where(inv > 0)
        if len(xs) == 0:
            return bin_img[0:0, 0:0]
        x0, x1 = int(xs.min()), int(xs.max()) + 1
        y0, y1 = int(ys.min()), int(ys.max()) + 1
        return bin_img[y0:y1, x0:x1]

    def _ncc(self, a: np.ndarray, b: np.ndarray) -> float:
        af = a.astype(np.float32)
        bf = b.astype(np.float32)
        am = af.mean()
        bm = bf.mean()
        ad = af - am
        bd = bf - bm
        denom = float(np.sqrt((ad * ad).sum()) * np.sqrt((bd * bd).sum()))
        if denom < 1e-6:
            return -1.0
        return float((ad * bd).sum() / denom)


def load_templates() -> Dict[str, np.ndarray]:
    mapping = {}
    if not os.path.isdir(TEMPLATE_DIR):
        return mapping

    name_map = {"dot": "."}
    for fname in os.listdir(TEMPLATE_DIR):
        if not fname.lower().endswith(".png"):
            continue
        stem = os.path.splitext(fname)[0].lower()
        ch = name_map.get(stem, stem)
        if ch == ":":
            continue
        if len(ch) != 1:
            continue
        path = os.path.join(TEMPLATE_DIR, fname)
        img = cv2.imread(path, cv2.IMREAD_COLOR)
        if img is None:
            continue
        mapping[ch] = img
    return mapping


TEMPLATE_READER = None


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

    x0 = REGION.x
    y0 = REGION.y
    x1 = REGION.x + REGION.w
    y1 = REGION.y + REGION.h
    canvas.create_rectangle(x0, y0, x1, y1, outline="#ffffff", width=4)

    for rr in ratios:
        tx0 = rr.rect.x
        ty0 = rr.rect.y
        tx1 = rr.rect.x + rr.rect.w
        ty1 = rr.rect.y + rr.rect.h
        canvas.create_rectangle(tx0, ty0, tx1, ty1, outline="#ff3b3b", width=3)

    root.update_idletasks()
    root.update()
    return root


def main():
    configure_tesseract()
    ratios = load_ratio_regions(CONFIG_PATH)
    global TEMPLATE_READER
    if any(r.use_templates for r in ratios):
        templates = load_templates()
        if templates:
            TEMPLATE_READER = FixedFontTemplateReader(templates)
    overlay = init_overlay(ratios)
    if DEBUG_DUMP:
        os.makedirs(DEBUG_DIR, exist_ok=True)
        for name in os.listdir(DEBUG_DIR):
            path = os.path.join(DEBUG_DIR, name)
            if os.path.isfile(path):
                os.remove(path)
    tick = 0
    with mss.mss() as sct:
        while True:
            for rr in ratios:
                img = np.array(sct.grab(rr.rect.as_mss()))[:, :, :3]
                preps = get_ocr_preps(rr)
                ratio, raw, candidates = read_ratio_from_image(img, preps, rr.use_templates)
                if ratio is None:
                    print(f"{rr.name}: ?")
                else:
                    display = coerce_full_number(raw, ratio)
                    display = apply_decimal_rule(raw, display, rr.decimal_rule)
                    print(f"{rr.name}: {format_decimals(display, 2)}")
                for label, c_raw, c_ratio in candidates:
                    if c_ratio is None:
                        print(f"  {label}: raw='{c_raw}' value=?")
                    else:
                        adjusted = apply_decimal_rule(c_raw, c_ratio, rr.decimal_rule)
                        print(f"  {label}: raw='{c_raw}' value={format_decimals(adjusted, 2)}")
                if DEBUG_DUMP and (tick % DEBUG_EVERY_N == 0):
                    raw_path = os.path.join(DEBUG_DIR, f"{rr.name}_{tick}_raw.png")
                    bin_path = os.path.join(DEBUG_DIR, f"{rr.name}_{tick}_bin.png")
                    bin_img = prep_for_ocr(img, scale=rr.scale)
                    cv2.imwrite(raw_path, img)
                    cv2.imwrite(bin_path, bin_img)
            print("===========")
            overlay.update_idletasks()
            overlay.update()
            tick += 1
            time.sleep(1.0)


if __name__ == "__main__":
    main()
