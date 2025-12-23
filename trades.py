import json
import time
from dataclasses import dataclass
from typing import Any, Dict, Iterable, List, Optional

from input_core import click, type_text


@dataclass(frozen=True)
class TradeAction:
    type: str
    payload: Dict[str, Any]


@dataclass(frozen=True)
class Trade:
    name: str
    actions: List[TradeAction]
    vars: Dict[str, Any]


class TradeRunner:
    def __init__(
        self,
        trades_path: str = "trades.json",
        delay_sec: float = 1.0,
        project_path: Optional[str] = None,
    ):
        self.trades_path = trades_path
        self.project_path = project_path
        self.delay_sec = float(delay_sec)
        self._trades: Dict[str, Trade] = {}
        self.reload()

    def reload(self) -> None:
        data = self._load_source()
        trades = {}
        for t in data.get("trades", []):
            name = str(t.get("name", "")).strip()
            if not name:
                continue
            actions = []
            vars_map = {k: v for k, v in t.items() if k not in ("name", "actions")}
            for a in t.get("actions", []):
                a_type = str(a.get("type", "")).strip().lower()
                if not a_type:
                    continue
                payload = dict(a)
                payload.pop("type", None)
                actions.append(TradeAction(type=a_type, payload=payload))
            trades[name] = Trade(name=name, actions=actions, vars=vars_map)
        self._trades = trades

    def _load_source(self) -> Dict[str, Any]:
        if self.project_path:
            with open(self.project_path, "r", encoding="utf-8") as f:
                return json.load(f)
        with open(self.trades_path, "r", encoding="utf-8") as f:
            return json.load(f)

    def list_trades(self) -> List[str]:
        return sorted(self._trades.keys())

    def run_trade(
        self, name: str, delay_sec: Optional[float] = None, overrides: Optional[Dict[str, Any]] = None
    ) -> None:
        trade = self._trades.get(name)
        if trade is None:
            raise ValueError(f"Trade not found: {name}")
        delay = self.delay_sec if delay_sec is None else float(delay_sec)
        vars_map = dict(trade.vars)
        if overrides:
            vars_map.update(overrides)
        self._run_actions(trade.actions, delay, vars_map)

    def run_trades(
        self,
        names: Iterable[str],
        delay_sec: Optional[float] = None,
        repeat: int = 1,
        interval_sec: float = 0.0,
    ) -> None:
        delay = self.delay_sec if delay_sec is None else float(delay_sec)
        repeat = int(repeat)
        loop_forever = repeat == 0
        count = 0
        names_list = [str(n) for n in names]
        while loop_forever or count < repeat:
            for name in names_list:
                self.run_trade(name, delay_sec=delay)
            count += 1
            if interval_sec > 0 and (loop_forever or count < repeat):
                time.sleep(float(interval_sec))

    def _resolve_text(self, text: str, vars_map: Dict[str, Any]) -> str:
        out = text
        for k, v in vars_map.items():
            out = out.replace("{" + str(k) + "}", str(v))
        return out

    def _run_actions(self, actions: List[TradeAction], delay: float, vars_map: Dict[str, Any]) -> None:
        for i, action in enumerate(actions):
            a_type = action.type
            payload = action.payload
            if a_type == "click":
                click(
                    int(payload["x"]),
                    int(payload["y"]),
                    button=str(payload.get("button", "left")).lower(),
                    modifiers=payload.get("modifiers", []),
                )
            elif a_type == "type":
                type_text(
                    self._resolve_text(str(payload.get("text", "")), vars_map),
                    press_enter=bool(payload.get("press_enter", False)),
                    key_delay=float(payload.get("key_delay_sec", 0.0)),
                    method=str(payload.get("method", "unicode")),
                )
            elif a_type == "delay":
                time.sleep(float(payload.get("seconds", 0.0)))
            else:
                raise ValueError(f"Unknown action type: {a_type}")

            if delay > 0 and i < len(actions) - 1:
                time.sleep(delay)
