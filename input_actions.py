import argparse

from input_core import click, split_mods, type_text
from trades import TradeRunner


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
    known_cmds = {"click", "type", "run-trade", "run-trades"}
    if argv and argv[0] not in known_cmds:
        argv = ["run-trade", "--name", argv[0]] + argv[1:]

    args = parser.parse_args(argv)
    if args.cmd == "click":
        click(args.x, args.y, button=args.button, modifiers=split_mods(args.mods))
    elif args.cmd == "type":
        type_text(args.text, press_enter=args.enter, key_delay=args.delay, method=args.method)
    elif args.cmd == "run-trade":
        runner = TradeRunner(trades_path=args.file)
        runner.run_trades([args.name], delay_sec=args.delay, repeat=args.repeat, interval_sec=args.interval)
    elif args.cmd == "run-trades":
        names = [n.strip() for n in args.names.split(",") if n.strip()]
        if not names:
            raise SystemExit("No trade names provided.")
        runner = TradeRunner(trades_path=args.file)
        runner.run_trades(names, delay_sec=args.delay, repeat=args.repeat, interval_sec=args.interval)


if __name__ == "__main__":
    main()
