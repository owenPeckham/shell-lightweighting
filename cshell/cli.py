"""Command line (THEORY §10.3): python -m cshell run <cfg> [--start S] [--phase P] [--set k=v ...] [--resume DIR]
[--max-hours H]. Further commands (setup, collect, report, bench, faces, init) arrive with Stage 2."""
from __future__ import annotations

import argparse
import json
import os

from . import config


def cmd_run(a):
    from .optimise import Context, Run, run_info
    if a.resume:
        cfg = json.load(open(os.path.join(a.resume, "config_used.json")))
        out = a.resume
    else:
        cfg = config.load(a.config, config.parse_set(a.set))
        out = os.path.join(cfg["output_dir"], cfg["name"])
        os.makedirs(out, exist_ok=True)
        json.dump(cfg, open(os.path.join(out, "config_used.json"), "w"), indent=1)
    starts = [a.start] if a.start else cfg["optimiser"]["starts"]
    run_info(out, dict(command="run", starts=starts, resume=bool(a.resume)))
    ctx = Context(cfg)
    for s in starts:
        r = Run(cfg, s, out, max_hours=a.max_hours).run(ctx, resume=bool(a.resume), only_phase=a.phase)
        print(json.dumps(dict(start=s, **r)))
        if r.get("status") == "paused":
            break


def main(argv=None):
    p = argparse.ArgumentParser(prog="cshell")
    sub = p.add_subparsers(dest="cmd", required=True)
    r = sub.add_parser("run", help="optimise one part")
    r.add_argument("config", nargs="?")
    r.add_argument("--start", choices=["uniform", "neutral", "solid", "vshell"])
    r.add_argument("--phase", type=int)
    r.add_argument("--set", nargs="*", help="overrides, e.g. objective.w_capacity=0.8")
    r.add_argument("--resume", help="run directory to resume")
    r.add_argument("--max-hours", type=float)
    r.set_defaults(fn=cmd_run)
    a = p.parse_args(argv)
    a.fn(a)
