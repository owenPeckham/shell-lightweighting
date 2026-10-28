"""Command line: python -m vshell {run, faces, report, init} ..."""
from __future__ import annotations

import argparse
import json
import os
import sys


def _parse_set(items):
    """--set grid.sdf_mm=0.5 search.max_fe_evaluations=6  ->  nested dict."""
    out = {}
    for it in items or []:
        key, _, val = it.partition("=")
        try:
            val = json.loads(val)
        except json.JSONDecodeError:
            pass
        d = out
        parts = key.split(".")
        for p in parts[:-1]:
            d = d.setdefault(p, {})
        d[parts[-1]] = val
    return out


def cmd_run(a):
    from . import optimise
    dec = optimise.run(a.config, _parse_set(a.set))
    return 0 if dec is not None else 1


def cmd_faces(a):
    from . import geometry
    shape = geometry.load_step(a.step)
    tab = geometry.face_table(shape)
    print(f"{'id':>4} {'type':<15} {'area mm2':>9}  centroid (mm)              detail")
    for r in tab:
        det = ""
        if r["type"] == "Plane":
            det = "n = (" + ", ".join(f"{v:+.2f}" for v in r["normal"]) + ")"
        elif r["type"] == "Cylinder":
            det = f"r = {r['radius']:.3f}, axis = (" + ", ".join(f"{v:+.2f}" for v in r["axis"]) + ")"
        if a.all or r["type"] != "BSplineSurface":
            c = ", ".join(f"{v:8.2f}" for v in r["centroid"])
            print(f"{r['id']:>4} {r['type']:<15} {r['area']:9.1f}  ({c})  {det}")
    if a.png:
        from .report import faces_figure
        fig = faces_figure(shape, tab)
        fig.savefig(a.png, dpi=130, bbox_inches="tight")
        print(f"wrote {a.png}")
    return 0


def cmd_report(a):
    from . import report
    for d in a.run_dirs:
        report.write_report(d, dpi=a.dpi)
    return 0


def cmd_init(a):
    """Template config: the smallest cylinder radius that occurs at least twice (bolt holes) as supports."""
    from collections import Counter
    from . import geometry
    shape = geometry.load_step(a.step)
    tab = geometry.face_table(shape)
    cyl = Counter(round(r["radius"], 2) for r in tab if r["type"] == "Cylinder")
    rep = sorted(r for r, n in cyl.items() if n >= 2) or sorted(cyl)
    sup = {"type": "cylinder", "radius": rep[0], "tol": 0.05} if rep else {"type": "ids", "ids": []}
    cfg = {"name": os.path.splitext(os.path.basename(a.step))[0].replace(" ", "_"),
           "step_file": os.path.abspath(a.step), "output_dir": "runs", "supports": [sup],
           "loads": [{"faces": {"type": "ids", "ids": []}, "force_N": [0, 0, -200]}]}
    out = a.out or (os.path.splitext(a.step)[0] + ".json")
    with open(out, "w") as f:
        json.dump(cfg, f, indent=2)
    print(f"wrote {out}: check 'supports' and fill in the load face ids (python -m vshell faces {a.step})")
    return 0


def main(argv=None):
    p = argparse.ArgumentParser(prog="python -m vshell", description="Constant-mass variable-thickness shelling.")
    sub = p.add_subparsers(dest="cmd", required=True)
    r = sub.add_parser("run", help="optimise a part described by a JSON config")
    r.add_argument("config")
    r.add_argument("--set", nargs="*", help="overrides, e.g. grid.sdf_mm=0.5 search.max_fe_evaluations=6")
    r.set_defaults(func=cmd_run)
    f = sub.add_parser("faces", help="list the faces of a STEP file (to pick supports and loads)")
    f.add_argument("step")
    f.add_argument("--all", action="store_true", help="include BSpline faces")
    f.add_argument("--png", help="also draw the analytic faces with their ids")
    f.set_defaults(func=cmd_faces)
    rp = sub.add_parser("report", help="(re)write report.pdf for run directories")
    rp.add_argument("run_dirs", nargs="+")
    rp.add_argument("--dpi", type=int, default=130)
    rp.set_defaults(func=cmd_report)
    i = sub.add_parser("init", help="write a template config for a STEP file")
    i.add_argument("step")
    i.add_argument("--out")
    i.set_defaults(func=cmd_init)
    a = p.parse_args(argv)
    sys.exit(a.func(a))
