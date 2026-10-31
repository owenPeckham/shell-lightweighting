"""Command line (THEORY §10.3):

    python -m cshell faces <step> [--png faces.png] [--all]     list faces (= dshell faces)
    python -m cshell init  <step> [--out cfg.json]              template config (= dshell init)
    python -m cshell setup <cfg> [--set k=v ...]                geometry, bands, P per phase; DOF and memory plan
    python -m cshell run   <cfg> [--start S] [--phase P] [--set k=v ...] [--resume DIR] [--max-hours H]
    python -m cshell collect <run_dir> [--allow-incomplete] [--no-validate]   merge starts, export, re-score, decide
    python -m cshell report  <run_dir>                          report.pdf (cshell.report.build)
    python -m cshell bench   [B1|B2|B3|all] [--set ...] [--no-collect]       Appendix B benchmarks: run + collect

`run` exit codes: 0 done, 3 paused by --max-hours (resume with --resume <run_dir>). With $CSHELL_RUNDIR_FILE set, the
run directory is written there (cluster/run.sh uses it to chain allocations). A finished start leaves
`done_<start>.json` in the run directory and is skipped on --resume.
"""
from __future__ import annotations

import argparse
import json
import os
import sys

from . import config

BENCH = {"B1": "B1_stub.json", "B1b": "B1b_stub_slender.json","B2": "B2_slender.json", "B3": "B3_column.json"}
EXIT_PAUSED = 3


def cmd_run(a):
    from .optimise import Context, Run, run_info
    if a.resume:
        cfg = config.normalise_starts(json.load(open(os.path.join(a.resume, "config_used.json"))))
        out = a.resume
    else:
        cfg = config.load(a.config, config.parse_set(a.set))
        out = os.path.join(cfg["output_dir"], cfg["name"])
        os.makedirs(out, exist_ok=True)
        json.dump(cfg, open(os.path.join(out, "config_used.json"), "w"), indent=1)
    if os.environ.get("CSHELL_RUNDIR_FILE"):
        with open(os.environ["CSHELL_RUNDIR_FILE"], "w") as f:
            f.write(os.path.abspath(out) + "\n")
    starts = [config.canonical_start(a.start)] if a.start else cfg["optimiser"]["starts"]
    if a.resume:
        starts = [s for s in starts if not os.path.exists(config.start_file(out, "done_{}.json", s))]
    run_info(out, dict(command="run", starts=starts, resume=bool(a.resume)))
    if not starts:
        print("[cshell] all starts already done")
        return 0
    ctx = Context(cfg)
    for s in starts:
        run = Run(cfg, s, out, max_hours=a.max_hours)
        run.ctx = ctx                                    # used by the dshell start (Run.initial_xi)
        r = run.run(ctx, resume=bool(a.resume), only_phase=a.phase)
        print(json.dumps(dict(start=s, **r), default=float))
        if r.get("status") == "paused":
            return EXIT_PAUSED
        if a.phase is None:
            json.dump(dict(start=s, **r), open(os.path.join(out, f"done_{s}.json"), "w"), indent=1, default=float)
    return 0


def cmd_collect(a):
    from .collect import collect
    S = collect(a.run_dir, allow_incomplete=a.allow_incomplete, validate=False if a.no_validate else None)
    print(f"{S['outcome']}: {S['reason']}")
    return 0


def cmd_report(a):
    from . import report
    print(report.build(a.run_dir))
    return 0


def cmd_setup(a):
    """Context and, per phase, the band, P, BCs and the solid's FE active set: prints sizes and a memory plan."""
    import numpy as np
    from scipy import ndimage
    from .optimise import Context, PhaseModel
    cfg = config.load(a.config, config.parse_set(a.set))
    ctx = Context(cfg)
    D = cfg["design"]["outward_extent_initial_mm"]
    print(f"part {cfg['name']}: Γ_h {ctx.surf.n_vertices} vertices, mounts {len(ctx.mount_ids)} faces, "
          f"|F| {np.linalg.norm(ctx.F):.1f} N")
    for ph in range(len(cfg["grid"]["phases"])):
        m = PhaseModel(ctx, ph, D)
        margin = int(np.ceil(cfg["grid"]["active_margin_mm"] / m.h_e))
        mask = ndimage.binary_dilation((np.asarray(m.rho0) > 0).reshape(m.elem_shape), iterations=margin)
        nodes = ndimage.binary_dilation(np.pad(mask, 1), np.ones((2, 2, 2), bool))[1:, 1:, 1:].sum()
        dof = 3 * int(nodes)
        P = getattr(m.disc, "P", None)                   # CPU chain: CSR P; GPU chain: compact band (nbytes)
        geo = (P.data.nbytes + P.indices.nbytes + P.indptr.nbytes + m.band.nbytes) if P is not None else m.disc.nbytes
        # FE estimate from the measured cantilever phase B (6.1 M DOF -> 84 GiB peak incl. factor; STATUS/BENCHMARKS)
        fe_gib = 84.0 * dof / 6.1e6
        print(f"phase {ph}: h_e {m.h_e} h_g {m.h_g}  band {m.band.n / 1e6:.1f} M cells  elements {int(mask.sum())}  "
              f"DOF ≈ {dof / 1e6:.2f} M  geometry {geo / 2**30:.1f} GiB ({m.geometry_device if hasattr(m, 'geometry_device') else 'cpu'})  FE ≈ {fe_gib:.0f} GiB GPU (scaled "
              f"estimate){'  > 95 GiB: does NOT fit one GH200' if fe_gib > 95 else ''}  setup {m.t_setup:.0f} s")
        if hasattr(m, "release"):
            m.release()
        del m
    return 0


def cmd_bench(a):
    from pathlib import Path
    root = Path(__file__).resolve().parents[1] / "examples/cshell/bench"
    names = list(BENCH) if a.which == "all" else [a.which]
    rc = 0
    for n in names:
        cfg_path = str(root / BENCH[n])
        cfg = config.load(cfg_path, config.parse_set(a.set))
        ns = argparse.Namespace(config=cfg_path, set=a.set, resume=None, start=None, phase=None, max_hours=None)
        rc = cmd_run(ns) or rc
        if not a.no_collect:
            cmd_collect(argparse.Namespace(run_dir=os.path.join(cfg["output_dir"], cfg["name"]),
                                           allow_incomplete=False, no_validate=False))
    return rc


def _dshell(argv):
    from dshell import cli as vcli
    return vcli.main(argv)


def main(argv=None):
    argv = sys.argv[1:] if argv is None else list(argv)
    if argv and argv[0] in ("faces", "init"):            # delegate to dshell's identical commands
        return _dshell(argv)
    p = argparse.ArgumentParser(prog="cshell")
    sub = p.add_subparsers(dest="cmd", required=True)
    r = sub.add_parser("run", help="optimise one part")
    r.add_argument("config", nargs="?")
    r.add_argument("--start", choices=["uniform", "neutral", "solid", "dshell", "vshell"],
                   help="one start (default: the config's optimiser.starts); vshell = old name of dshell")
    r.add_argument("--phase", type=int)
    r.add_argument("--set", nargs="*", help="overrides, e.g. objective.w_capacity=0.8")
    r.add_argument("--resume", help="run directory to resume")
    r.add_argument("--max-hours", type=float)
    r.set_defaults(fn=cmd_run)
    c = sub.add_parser("collect", help="merge starts, export, re-score, validate, decide")
    c.add_argument("run_dir")
    c.add_argument("--allow-incomplete", action="store_true", help="use the latest checkpoint if no start finished")
    c.add_argument("--no-validate", action="store_true", help="skip the body-fitted CalculiX check")
    c.set_defaults(fn=cmd_collect)
    rp = sub.add_parser("report", help="report.pdf from a run directory")
    rp.add_argument("run_dir")
    rp.set_defaults(fn=cmd_report)
    s = sub.add_parser("setup", help="build geometry, bands and P; print DOF and memory plan")
    s.add_argument("config")
    s.add_argument("--set", nargs="*")
    s.set_defaults(fn=cmd_setup)
    b = sub.add_parser("bench", help="Appendix B benchmarks (run + collect)")
    b.add_argument("which", nargs="?", default="all", choices=list(BENCH) + ["all"])
    b.add_argument("--set", nargs="*")
    b.add_argument("--no-collect", action="store_true")
    b.set_defaults(fn=cmd_bench)
    sub.add_parser("faces", help="list STEP faces (dshell faces)")
    sub.add_parser("init", help="template config (dshell init)")
    a = p.parse_args(argv)
    return a.fn(a)
