import time
from dataclasses import dataclass
from typing import Optional

import tkinter as tk


@dataclass(frozen=True)
class Rect:
    x: int
    y: int
    w: int
    h: int


class RectPicker:
    def __init__(self):
        self.root = tk.Tk()
        self.root.withdraw()
        self.root.attributes("-topmost", True)

        self.win = tk.Toplevel(self.root)
        self.win.overrideredirect(True)
        self.win.attributes("-topmost", True)
        self.win.attributes("-alpha", 0.25)

        self.win.configure(cursor="crosshair")

        sw = self.win.winfo_screenwidth()
        sh = self.win.winfo_screenheight()
        self.win.geometry(f"{sw}x{sh}+0+0")

        self.canvas = tk.Canvas(self.win, highlightthickness=0)
        self.canvas.pack(fill="both", expand=True)

        self._start_x: Optional[int] = None
        self._start_y: Optional[int] = None
        self._rect_id: Optional[int] = None
        self.result: Optional[Rect] = None

        self.canvas.bind("<ButtonPress-1>", self._on_down)
        self.canvas.bind("<B1-Motion>", self._on_drag)
        self.canvas.bind("<ButtonRelease-1>", self._on_up)
        self.win.bind("<Escape>", self._on_escape)

    def pick(self) -> Optional[Rect]:
        self.win.deiconify()
        self.root.deiconify()
        self.root.mainloop()
        return self.result

    def _on_down(self, e):
        self._start_x = self.win.winfo_pointerx()
        self._start_y = self.win.winfo_pointery()
        self._rect_id = self.canvas.create_rectangle(
            self._start_x, self._start_y, self._start_x, self._start_y,
            width=2
        )

    def _on_drag(self, e):
        if self._start_x is None or self._start_y is None or self._rect_id is None:
            return
        cx = self.win.winfo_pointerx()
        cy = self.win.winfo_pointery()
        self.canvas.coords(self._rect_id, self._start_x, self._start_y, cx, cy)

    def _on_up(self, e):
        if self._start_x is None or self._start_y is None:
            self._quit()
            return
        end_x = self.win.winfo_pointerx()
        end_y = self.win.winfo_pointery()

        x0 = min(self._start_x, end_x)
        y0 = min(self._start_y, end_y)
        x1 = max(self._start_x, end_x)
        y1 = max(self._start_y, end_y)

        self.result = Rect(x=x0, y=y0, w=x1 - x0, h=y1 - y0)
        print(f"{self.result.x}, {self.result.y}, {self.result.w}, {self.result.h}")
        self._quit()

    def _on_escape(self, e):
        self.result = None
        self._quit()

    def _quit(self):
        try:
            self.win.destroy()
        except Exception:
            pass
        try:
            self.root.quit()
            self.root.destroy()
        except Exception:
            pass


if __name__ == "__main__":
    r = RectPicker().pick()
    if r is None:
        print("Cancelled")
    else:
        print(f"Rect(x={r.x}, y={r.y}, w={r.w}, h={r.h})")
