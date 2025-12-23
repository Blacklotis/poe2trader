import argparse
import ctypes
import json
import time
from dataclasses import dataclass
from typing import Iterable, List, Optional, Dict, Any


user32 = ctypes.WinDLL("user32", use_last_error=True)
kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)

SM_CXSCREEN = 0
SM_CYSCREEN = 1

INPUT_MOUSE = 0
INPUT_KEYBOARD = 1

MOUSEEVENTF_MOVE = 0x0001
MOUSEEVENTF_ABSOLUTE = 0x8000
MOUSEEVENTF_LEFTDOWN = 0x0002
MOUSEEVENTF_LEFTUP = 0x0004
MOUSEEVENTF_RIGHTDOWN = 0x0008
MOUSEEVENTF_RIGHTUP = 0x0010

KEYEVENTF_KEYUP = 0x0002
KEYEVENTF_UNICODE = 0x0004

VK_SHIFT = 0x10
VK_CONTROL = 0x11
VK_MENU = 0x12  # Alt
VK_RETURN = 0x0D
VK_V = 0x56

CF_UNICODETEXT = 13
GMEM_MOVEABLE = 0x0002
GMEM_ZEROINIT = 0x0040


@dataclass(frozen=True)
class Modifier:
    name: str
    vk: int


MODIFIER_MAP = {
    "shift": Modifier("shift", VK_SHIFT),
    "ctrl": Modifier("ctrl", VK_CONTROL),
    "control": Modifier("control", VK_CONTROL),
    "alt": Modifier("alt", VK_MENU),
}


class MOUSEINPUT(ctypes.Structure):
    _fields_ = [
        ("dx", ctypes.c_long),
        ("dy", ctypes.c_long),
        ("mouseData", ctypes.c_ulong),
        ("dwFlags", ctypes.c_ulong),
        ("time", ctypes.c_ulong),
        ("dwExtraInfo", ctypes.POINTER(ctypes.c_ulong)),
    ]


class KEYBDINPUT(ctypes.Structure):
    _fields_ = [
        ("wVk", ctypes.c_ushort),
        ("wScan", ctypes.c_ushort),
        ("dwFlags", ctypes.c_ulong),
        ("time", ctypes.c_ulong),
        ("dwExtraInfo", ctypes.POINTER(ctypes.c_ulong)),
    ]


class INPUT(ctypes.Structure):
    class _INPUT(ctypes.Union):
        _fields_ = [("mi", MOUSEINPUT), ("ki", KEYBDINPUT)]

    _fields_ = [("type", ctypes.c_ulong), ("union", _INPUT)]


def _send_input(inputs: List[INPUT]) -> None:
    n = len(inputs)
    if n == 0:
        return
    cb = ctypes.sizeof(INPUT)
    sent = user32.SendInput(n, (INPUT * n)(*inputs), cb)
    if sent != n:
        raise ctypes.WinError(ctypes.get_last_error())


def _screen_size() -> tuple[int, int]:
    return user32.GetSystemMetrics(SM_CXSCREEN), user32.GetSystemMetrics(SM_CYSCREEN)


def _normalize_coords(x: int, y: int) -> tuple[int, int]:
    w, h = _screen_size()
    nx = int(x * 65535 / max(1, w - 1))
    ny = int(y * 65535 / max(1, h - 1))
    return nx, ny


def _mouse_input(flags: int, dx: int = 0, dy: int = 0) -> INPUT:
    mi = MOUSEINPUT(dx=dx, dy=dy, mouseData=0, dwFlags=flags, time=0, dwExtraInfo=None)
    return INPUT(type=INPUT_MOUSE, union=INPUT._INPUT(mi=mi))


def _key_input(vk: int, flags: int = 0) -> INPUT:
    ki = KEYBDINPUT(wVk=vk, wScan=0, dwFlags=flags, time=0, dwExtraInfo=None)
    return INPUT(type=INPUT_KEYBOARD, union=INPUT._INPUT(ki=ki))


def _unicode_input(ch: str, keyup: bool = False) -> INPUT:
    flags = KEYEVENTF_UNICODE | (KEYEVENTF_KEYUP if keyup else 0)
    ki = KEYBDINPUT(wVk=0, wScan=ord(ch), dwFlags=flags, time=0, dwExtraInfo=None)
    return INPUT(type=INPUT_KEYBOARD, union=INPUT._INPUT(ki=ki))


def _parse_modifiers(mods: Optional[Iterable[str]]) -> List[Modifier]:
    out: List[Modifier] = []
    for m in mods or []:
        key = str(m).strip().lower()
        if key in MODIFIER_MAP:
            out.append(MODIFIER_MAP[key])
    return out


def _press_modifiers(mods: List[Modifier]) -> None:
    _send_input([_key_input(m.vk, 0) for m in mods])


def _release_modifiers(mods: List[Modifier]) -> None:
    _send_input([_key_input(m.vk, KEYEVENTF_KEYUP) for m in reversed(mods)])


def click(
    x: int,
    y: int,
    button: str = "left",
    modifiers: Optional[Iterable[str]] = None,
) -> None:
    mods = _parse_modifiers(modifiers)
    nx, ny = _normalize_coords(x, y)

    if mods:
        _press_modifiers(mods)

    move = _mouse_input(MOUSEEVENTF_MOVE | MOUSEEVENTF_ABSOLUTE, dx=nx, dy=ny)
    if button == "right":
        down = _mouse_input(MOUSEEVENTF_RIGHTDOWN)
        up = _mouse_input(MOUSEEVENTF_RIGHTUP)
    else:
        down = _mouse_input(MOUSEEVENTF_LEFTDOWN)
        up = _mouse_input(MOUSEEVENTF_LEFTUP)
    _send_input([move, down, up])

    if mods:
        _release_modifiers(mods)


def _set_clipboard_text(text: str) -> None:
    user32.OpenClipboard.argtypes = [ctypes.c_void_p]
    user32.OpenClipboard.restype = ctypes.c_bool
    user32.EmptyClipboard.argtypes = []
    user32.EmptyClipboard.restype = ctypes.c_bool
    user32.SetClipboardData.argtypes = [ctypes.c_uint, ctypes.c_void_p]
    user32.SetClipboardData.restype = ctypes.c_void_p
    user32.CloseClipboard.argtypes = []
    user32.CloseClipboard.restype = ctypes.c_bool

    kernel32.GlobalAlloc.argtypes = [ctypes.c_uint, ctypes.c_size_t]
    kernel32.GlobalAlloc.restype = ctypes.c_void_p
    kernel32.GlobalLock.argtypes = [ctypes.c_void_p]
    kernel32.GlobalLock.restype = ctypes.c_void_p
    kernel32.GlobalUnlock.argtypes = [ctypes.c_void_p]
    kernel32.GlobalUnlock.restype = ctypes.c_bool

    if not user32.OpenClipboard(None):
        raise ctypes.WinError(ctypes.get_last_error())
    try:
        if not user32.EmptyClipboard():
            raise ctypes.WinError(ctypes.get_last_error())
        data = text.encode("utf-16-le") + b"\x00\x00"
        h_mem = kernel32.GlobalAlloc(GMEM_MOVEABLE | GMEM_ZEROINIT, len(data))
        if not h_mem:
            raise ctypes.WinError(ctypes.get_last_error())
        locked = kernel32.GlobalLock(h_mem)
        if not locked:
            raise ctypes.WinError(ctypes.get_last_error())
        try:
            ctypes.memmove(locked, data, len(data))
        finally:
            kernel32.GlobalUnlock(h_mem)
        if not user32.SetClipboardData(CF_UNICODETEXT, h_mem):
            raise ctypes.WinError(ctypes.get_last_error())
    finally:
        user32.CloseClipboard()


def _paste_clipboard() -> None:
    _send_input([_key_input(VK_CONTROL, 0), _key_input(VK_V, 0)])
    _send_input([_key_input(VK_V, KEYEVENTF_KEYUP), _key_input(VK_CONTROL, KEYEVENTF_KEYUP)])


def type_text(
    text: str,
    press_enter: bool = False,
    key_delay: float = 0.0,
    method: str = "unicode",
) -> None:
    method = (method or "unicode").lower()
    if method == "paste":
        _set_clipboard_text(text)
        _paste_clipboard()
    else:
        for ch in text:
            _send_input([_unicode_input(ch, False), _unicode_input(ch, True)])
            if key_delay > 0:
                time.sleep(key_delay)
    if press_enter:
        _send_input([_key_input(VK_RETURN, 0), _key_input(VK_RETURN, KEYEVENTF_KEYUP)])


def _split_mods(raw: str) -> List[str]:
    if not raw:
        return []
    return [s.strip() for s in raw.split(",") if s.strip()]


def _load_trades(path: str) -> Dict[str, Any]:
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def run_trade(path: str, name: str, delay_sec: float) -> None:
    data = _load_trades(path)
    trades = data.get("trades", [])
    trade = None
    for t in trades:
        if t.get("name") == name:
            trade = t
            break
    if trade is None:
        raise SystemExit(f"Trade not found: {name}")

    actions = trade.get("actions", [])
    for i, action in enumerate(actions):
        a_type = action.get("type")
        if a_type == "click":
            click(
                int(action["x"]),
                int(action["y"]),
                button=str(action.get("button", "left")).lower(),
                modifiers=action.get("modifiers", []),
            )
        elif a_type == "type":
            type_text(
                str(action.get("text", "")),
                press_enter=bool(action.get("press_enter", False)),
                key_delay=float(action.get("key_delay_sec", 0.0)),
                method=str(action.get("method", "unicode")),
            )
        elif a_type == "delay":
            time.sleep(float(action.get("seconds", 0.0)))
        else:
            raise SystemExit(f"Unknown action type: {a_type}")
        if delay_sec > 0 and i < len(actions) - 1:
            time.sleep(delay_sec)


def run_trades(path: str, names: List[str], delay_sec: float, repeat: int, interval_sec: float) -> None:
    if repeat < 0:
        repeat = 0
    loop_forever = repeat == 0
    count = 0
    while loop_forever or count < repeat:
        for name in names:
            run_trade(path, name, delay_sec)
        count += 1
        if interval_sec > 0 and (loop_forever or count < repeat):
            time.sleep(interval_sec)


def main() -> None:
    parser = argparse.ArgumentParser(description="Mouse click and typing helper.")
    sub = parser.add_subparsers(dest="cmd", required=True)

    p_click = sub.add_parser("click", help="Click at a screen position.")
    p_click.add_argument("--x", type=int, required=True)
    p_click.add_argument("--y", type=int, required=True)
    p_click.add_argument("--button", choices=["left", "right"], default="left")
    p_click.add_argument("--mods", default="")

    p_type = sub.add_parser("type", help="Type text at current focus.")
    p_type.add_argument("--text", required=True)
    p_type.add_argument("--enter", action="store_true")
    p_type.add_argument("--delay", type=float, default=0.0)
    p_type.add_argument("--method", choices=["unicode", "paste"], default="unicode")

    p_run = sub.add_parser("run-trade", help="Run a trade by name from trades.json.")
    p_run.add_argument("--file", default="trades.json")
    p_run.add_argument("--name", required=True)
    p_run.add_argument("--delay", type=float, default=1.0)
    p_run.add_argument("--repeat", type=int, default=1, help="0 = loop forever")
    p_run.add_argument("--interval", type=float, default=0.0)

    p_runs = sub.add_parser("run-trades", help="Run multiple trades in sequence.")
    p_runs.add_argument("--file", default="trades.json")
    p_runs.add_argument("--names", required=True, help="Comma-separated trade names")
    p_runs.add_argument("--delay", type=float, default=1.0)
    p_runs.add_argument("--repeat", type=int, default=1, help="0 = loop forever")
    p_runs.add_argument("--interval", type=float, default=0.0)

    import sys

    argv = sys.argv[1:]
    known_cmds = {"click", "type", "run-trade"}
    if argv and argv[0] not in known_cmds:
        argv = ["run-trade", "--name", argv[0]] + argv[1:]

    args = parser.parse_args(argv)
    if args.cmd == "click":
        click(args.x, args.y, button=args.button, modifiers=_split_mods(args.mods))
    elif args.cmd == "type":
        type_text(args.text, press_enter=args.enter, key_delay=args.delay, method=args.method)
    elif args.cmd == "run-trade":
        run_trades(args.file, [args.name], args.delay, args.repeat, args.interval)
    elif args.cmd == "run-trades":
        names = [n.strip() for n in args.names.split(",") if n.strip()]
        if not names:
            raise SystemExit("No trade names provided.")
        run_trades(args.file, names, args.delay, args.repeat, args.interval)


if __name__ == "__main__":
    main()
