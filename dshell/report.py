"""
Report writer: summary.json/.txt, plots/*.png and report.pdf for a dshell run directory.

Runs at the end of every optimisation and on its own (python -m dshell report <run_dir>) from the
run's files: run_info.json, regions.csv, evaluations.csv, mass_sweep.csv, escape_holes.csv and
arrays.npz. Matplotlib is driven through its object-oriented API with the Agg canvas (no display,
no global backend change). Every figure is built in its own try/except; one that fails is listed in
plots/report_log.txt and the rest are still written.
"""
from __future__ import annotations

import json
import os
import textwrap
import traceback
from timeit import default_timer as timer

import numpy as np

from . import util

C_SOLID = "#4d4d4d"
C_HOLLOW = "#1f77b4"
C_MOUNT = "#9e9e9e"
C_JOINT = "#d62728"
C_ACC = "#ff7f0e"
C_OK = "#2c7fb8"
C_BAD = "#d62728"


def new_figure(width=11.0, height=7.5):
    from matplotlib.figure import Figure
    from matplotlib.backends.backend_agg import FigureCanvasAgg
    fig = Figure(figsize=(width, height))
    FigureCanvasAgg(fig)
    return fig


class Ctx:
    def __init__(self, run_dir, dpi=130):
        self.dir = run_dir
        self.dpi = dpi
        self.plot_dir = os.path.join(run_dir, "plots")
        os.makedirs(self.plot_dir, exist_ok=True)
        with open(os.path.join(run_dir, "run_info.json"), encoding="utf-8") as f:
            self.info = json.load(f)
        self.regions = util.read_csv(os.path.join(run_dir, "regions.csv"))
        self.evals = util.read_csv(os.path.join(run_dir, "evaluations.csv"))
        self.sweep = util.read_csv(os.path.join(run_dir, "mass_sweep.csv"))
        self.holes = util.read_csv(os.path.join(run_dir, "escape_holes.csv"))
        p = os.path.join(run_dir, "arrays.npz")
        self.A = dict(np.load(p)) if os.path.isfile(p) else {}
        self.name = self.info.get("name")
        self.log = []
        self.written = []

    def region(self, rid):
        for r in self.regions:
            if int(r["id"]) == int(rid):
                return r
        return None


# ----------------------------------------------------------------------------------------------
# summary
# ----------------------------------------------------------------------------------------------
def _fmt(v, nd=4):
    if v is None:
        return "-"
    if isinstance(v, bool):
        return "yes" if v else "no"
    if isinstance(v, (int, np.integer)):
        return str(int(v))
    if isinstance(v, (float, np.floating)):
        v = float(v)
        if not np.isfinite(v):
            return "inf" if v > 0 else "nan"
        if v != 0 and (abs(v) < 1e-3 or abs(v) >= 1e6):
            return f"{v:.{nd}e}"
        return f"{v:.{nd}g}"
    if isinstance(v, str) and v in ("inf", "-inf"):
        return v
    if isinstance(v, (list, tuple)):
        return "[" + ", ".join(_fmt(x, nd) for x in v[:14]) + (", ..." if len(v) > 14 else "") + "]"
    return str(v)


def compute_summary(c: Ctx) -> dict:
    I = c.info
    base, best, dec = I.get("baseline") or {}, I.get("best") or {}, I.get("decision") or {}
    cfg = I.get("config", {})
    S = {}
    S["decision"] = {"outcome": "HOLLOW design accepted" if dec.get("accepted") else "original SOLID part kept",
                     "reason": dec.get("reason"), "best design (k solid regions)": dec.get("best_k"),
                     "final mesh": "final_design.stl"}
    g = I.get("geometry", {})
    S["part"] = {"STEP file": cfg.get("step_file"), "volume (B-rep) mm3": g.get("V0_brep_mm3"),
                 "volume (field) mm3": g.get("V0_field_mm3"), "mass g": g.get("mass_g"), "faces": g.get("n_faces"),
                 "analytic faces": g.get("n_analytic"), "support faces": g.get("support_faces"),
                 "load faces": g.get("load_faces"), "force N": g.get("force_N")}
    seg = I.get("segmentation", {})
    S["segmentation"] = {"joints": seg.get("n_joint"), "struts": seg.get("n_strut"), "mounts": seg.get("n_mount"),
                         "rank order (most shear-critical first)": I.get("rank_order")}
    S["baseline (solid)"] = {"stiffness K0 N/mm": base.get("K"), "load capacity P0 N": base.get("P"),
                             "governing mode": base.get("mode"), "lambda strength": base.get("lam_s"),
                             "lambda Euler": base.get("lam_E"), "von Mises q MPa": base.get("sigma_q"),
                             "mount von Mises q MPa": base.get("sigma_mount_q"), "FE DOF": base.get("ndof"),
                             "FE cells": base.get("n_cells"), "CG iterations": base.get("cg_iterations")}
    if best:
        S["best hollow candidate"] = {"solid regions": best.get("solid"), "hollow regions": best.get("hollow"),
                                      "outer offset / strut wall s mm": best.get("s"),
                                      "as-built wall mm": (best.get("s") or 0) - cfg.get("manufacturing", {}).get("wall_bias_mm", 0),
                                      "stiffness K N/mm": best.get("K"), "K/K0": best.get("K_ratio"),
                                      "load capacity P N": best.get("P"), "P/P0": best.get("P_ratio"), "J": best.get("J"),
                                      "governing mode": best.get("mode"), "governing region": best.get("governing_region"),
                                      "lambda strength / Euler / local": [best.get("lam_s"), best.get("lam_E"), best.get("lam_L")],
                                      "mass (field) g": best.get("mass_g"), "enclosed voids": len(best.get("voids") or []),
                                      "filled small voids mm3": best.get("filled_mm3"),
                                      "mesh volume mm3": (best.get("mesh") or {}).get("volume"),
                                      "mesh watertight": (best.get("mesh") or {}).get("watertight")}
    S["search"] = {"hollow designs analysed": len(c.evals),
                   "feasible in mass sweep": sum(1 for r in c.sweep if r.get("feasible") in (True, "True", 1.0)) - 1,
                   "objective": cfg.get("objective")}
    S["grid"] = I.get("grid")
    S["timings s"] = I.get("timings")
    pv = I.get("provenance", {})
    S["provenance"] = {"time (UTC)": pv.get("timestamp_utc"), "host": pv.get("host"), "platform": pv.get("platform"),
                       "python": pv.get("python"), "git": (pv.get("git") or {}).get("describe"),
                       "STEP sha256": (pv.get("step_file") or {}).get("sha256", "")[:16],
                       "packages": ", ".join(f"{k} {v}" for k, v in (pv.get("packages") or {}).items() if v)}
    return S


def summary_text(S: dict, name: str) -> str:
    lines = [f"dshell run summary: {name}", "=" * 78]
    for sec, d in S.items():
        if not d:
            continue
        lines += ["", sec, "-" * len(sec)]
        if isinstance(d, dict):
            for k, v in d.items():
                val = _fmt(v)
                wrapped = textwrap.wrap(val, 46) or [""]
                lines.append(f"  {k:30s} {wrapped[0]}")
                lines += [" " * 33 + w for w in wrapped[1:]]
        else:
            lines.append("  " + _fmt(d))
    return "\n".join(lines) + "\n"


def fig_summary(c: Ctx, S):
    text = summary_text(S, c.name).splitlines()
    figs = []
    per = 80
    for i in range(0, len(text), per):
        fig = new_figure(8.27, 11.69)
        if i == 0:
            acc = c.info.get("decision", {}).get("accepted")
            fig.text(0.04, 0.975, f"{c.name}: {'HOLLOW DESIGN ACCEPTED' if acc else 'ORIGINAL SOLID KEPT'}",
                     fontsize=13, weight="bold", color=C_OK if acc else C_SOLID, va="top")
            y0 = 0.95
        else:
            y0 = 0.98
        fig.text(0.04, y0, "\n".join(text[i:i + per]), family="monospace", fontsize=7.0, va="top", ha="left")
        figs.append((f"00_summary_p{i // per + 1}", fig))
    return figs


# ----------------------------------------------------------------------------------------------
# 3D helpers
# ----------------------------------------------------------------------------------------------
def _equal3d(ax, P):
    c = (P.max(0) + P.min(0)) / 2
    r = (P.max(0) - P.min(0)).max() / 2 * 1.05
    ax.set_xlim(c[0] - r, c[0] + r)
    ax.set_ylim(c[1] - r, c[1] + r)
    ax.set_zlim(c[2] - r, c[2] + r)
    ax.set_xlabel("x", fontsize=7)
    ax.set_ylabel("y", fontsize=7)
    ax.set_zlabel("z", fontsize=7)
    ax.tick_params(labelsize=6)


def _draw_mesh(ax, V, T, face_colour=None, base="0.78"):
    from mpl_toolkits.mplot3d.art3d import Poly3DCollection
    tri = V[T]
    n = np.cross(tri[:, 1] - tri[:, 0], tri[:, 2] - tri[:, 0])
    n /= np.linalg.norm(n, axis=1, keepdims=True) + 1e-12
    shade = 0.45 + 0.55 * np.clip(n @ np.array([0.35, -0.45, 0.82]), 0, 1)
    import matplotlib.colors as mc
    if face_colour is None:
        rgb = np.tile(np.array(mc.to_rgb(base)), (len(T), 1))
    else:
        rgb = np.asarray(face_colour)[:, :3]
    col = np.clip(rgb * shade[:, None], 0, 1)
    pc = Poly3DCollection(tri, facecolors=col, edgecolor="none")
    pc.set_rasterized(True)        # keeps report.pdf small
    ax.add_collection3d(pc)


VIEWS = [(25, -60, "isometric"), (90, -90, "top (+z)"), (0, -90, "front (-y)")]


def faces_figure(shape, table):
    """Analytic faces coloured, with their ids (for picking supports and loads)."""
    from . import geometry
    V, T, F = geometry.tessellate(shape, 0.3)
    fig = new_figure(15, 5.5)
    analytic = np.array([table[i]["type"] != "BSplineSurface" for i in F])
    import matplotlib
    cmap = matplotlib.colormaps["tab20"]
    col = np.array([cmap(f % 20) if a else (0.8, 0.8, 0.8, 1) for f, a in zip(F, analytic)])
    for j, (el, az, title) in enumerate(VIEWS):
        ax = fig.add_subplot(1, 3, j + 1, projection="3d")
        _draw_mesh(ax, V, T, col)
        for r in table:
            if r["type"] != "BSplineSurface" and r["area"] > 1.0:
                ax.text(*r["centroid"], str(r["id"]), fontsize=6, color="k")
        _equal3d(ax, V)
        ax.view_init(el, az)
        ax.set_title(title, fontsize=9)
    fig.suptitle("Analytic faces (coloured) with face ids; freeform faces grey", fontsize=10)
    fig.tight_layout()
    return fig


# ----------------------------------------------------------------------------------------------
# figures
# ----------------------------------------------------------------------------------------------
def fig_geometry(c: Ctx):
    A = c.A
    if "prev_V" not in A:
        return None
    V, T, F = A["prev_V"], A["prev_T"], A["prev_F"]
    g = c.info["geometry"]
    sup = set(g["support_faces"])
    load = {i for L in g["load_faces"] for i in L}
    col = np.array([(0.12, 0.47, 0.71, 1) if f in sup else ((1.0, 0.5, 0.05, 1) if f in load else (0.8, 0.8, 0.8, 1)) for f in F])
    fig = new_figure(15, 5.6)
    Fv = np.asarray(g["force_N"], float)
    lc = np.mean([c.info["faces"][i]["centroid"] for i in load], axis=0) if load else V.mean(0)
    span = np.ptp(V, 0).max()
    for j, (el, az, title) in enumerate(VIEWS):
        ax = fig.add_subplot(1, 3, j + 1, projection="3d")
        _draw_mesh(ax, V, T, col)
        d = -Fv / (np.linalg.norm(Fv) + 1e-12) * 0.25 * span
        ax.quiver(*(lc + d), *(-d), color=C_BAD, lw=2, arrow_length_ratio=0.25)
        _equal3d(ax, V)
        ax.view_init(el, az)
        ax.set_title(title, fontsize=9)
    fig.suptitle(f"{c.name}: supports (blue, all DOF fixed) and load (orange faces, red arrow, |F| = "
                 f"{np.linalg.norm(Fv):.0f} N)", fontsize=11)
    fig.tight_layout()
    return fig


def fig_segmentation(c: Ctx):
    A = c.A
    if "seg_xyz" not in A:
        return None
    P, L = A["seg_xyz"], A["seg_label"]
    sub = np.random.default_rng(0).random(len(P)) < min(1.0, 60000 / max(len(P), 1))
    P, L = P[sub], L[sub]
    import matplotlib
    vir = matplotlib.colormaps["viridis"]
    col, kinds = [], {}
    for l in L:
        r = c.region(l)
        k = r["kind"] if r else "mount"
        if k == "mount":
            col.append((0.62, 0.62, 0.62, 1))
        elif k == "joint":
            col.append((0.84, 0.15, 0.16, 1))
        else:
            col.append(vir(min((r.get("lam") or 0) / 60.0, 1.0)))
    col = np.array(col)
    fig = new_figure(15, 5.8)
    for j, (el, az, title) in enumerate(VIEWS):
        ax = fig.add_subplot(1, 3, j + 1, projection="3d")
        ax.scatter(*P.T, c=col, s=1.2, depthshade=False, rasterized=True)
        for r in c.regions:
            ax.text(r["cx"], r["cy"], r["cz"], str(int(r["id"])), fontsize=6)
        _equal3d(ax, P)
        ax.view_init(el, az)
        ax.set_title(title, fontsize=9)
    s = c.info.get("segmentation", {})
    fig.suptitle(f"Segmentation: {s.get('n_joint')} joints (red), {s.get('n_strut')} struts (colour = slenderness "
                 f"lambda, 0 to 60+), {s.get('n_mount')} mounts (grey, never hollowed); numbers = region ids", fontsize=10)
    fig.tight_layout()
    return fig


def fig_baseline_fe(c: Ctx, tag="base", title="Original solid"):
    A = c.A
    if f"{tag}_cell_xyz" not in A:
        return None
    P, vm, ex = A[f"{tag}_cell_xyz"], A[f"{tag}_cell_vm"], A[f"{tag}_cell_excluded"]
    reg, w = A[f"{tag}_cell_region"], A[f"{tag}_cell_vf"]
    mount_ids = [int(r["id"]) for r in c.regions if r["kind"] == "mount"]
    is_mount = np.isin(reg, mount_ids)
    cap = c.info["config"]["capacity"]
    use = ~ex & (~is_mount | bool(cap["include_mounts"]))
    keep = np.random.default_rng(1).random(len(P)) < min(1.0, 80000 / len(P))
    q = cap["stress_quantile"]
    from .mechanics import weighted_quantile
    vq = weighted_quantile(vm[use], w[use], q) if use.any() else float(vm.max())
    vmax = vq
    fig = new_figure(15, 6.2)
    for j, (el, az, t) in enumerate(VIEWS[:2]):
        ax = fig.add_subplot(1, 3, j + 1, projection="3d")
        sc = ax.scatter(*P[keep].T, c=np.clip(vm[keep], 0, vmax), cmap="inferno", s=1.0, depthshade=False, vmin=0, vmax=vmax,
                        rasterized=True)
        _equal3d(ax, P)
        ax.view_init(el, az)
        ax.set_title(t, fontsize=9)
    fig.colorbar(sc, ax=fig.axes, shrink=0.6, label=f"von Mises [MPa] (clipped at the design-region {q:.3g} quantile)")
    ax = fig.add_subplot(1, 3, 3)
    bins = np.linspace(0, max(float(vm[~ex].max()) if (~ex).any() else 1.0, 1e-6), 80)
    ax.hist(vm[use], bins=bins, weights=w[use], color=C_HOLLOW if tag == "best" else C_SOLID, alpha=0.8,
            label="joints and struts (used for strength)")
    mm = ~ex & is_mount
    if mm.any() and not cap["include_mounts"]:
        ax.hist(vm[mm], bins=bins, weights=w[mm], histtype="step", color=C_MOUNT, lw=1.2, label="mounts (reported only)")
    ax.axvline(vq, color=C_ACC, ls="--", label=f"{q:.3g} quantile = {vq:.1f} MPa")
    ax.axvline(c.info["config"]["material"]["strength_MPa"], color=C_BAD, label="failure stress")
    ax.set_yscale("log")
    ax.set_xlabel("von Mises [MPa] (nominal wall; the strength check adds the as-built factor s/(s-bias))", fontsize=7)
    ax.set_ylabel("volume-weighted count")
    ax.legend(fontsize=7)
    ax.set_title("Stress distribution (cells near supports and loads excluded)", fontsize=9)
    fig.suptitle(f"{title}: von Mises stress at |F| = {np.linalg.norm(c.info['geometry']['force_N']):.0f} N", fontsize=11)
    return fig


def fig_best_fe(c: Ctx):
    return fig_baseline_fe(c, "best", "Best hollow candidate")


def fig_regions(c: Ctx):
    R = [r for r in c.regions if r["kind"] != "mount"]
    if not R:
        return None
    fig = new_figure(13, 6.5)
    ax1, ax2 = fig.subplots(1, 2)
    order = c.info.get("rank_order") or [int(r["id"]) for r in R]
    rr = [c.region(i) for i in order]
    x = np.arange(len(rr))
    cols = [C_JOINT if r["kind"] == "joint" else C_HOLLOW for r in rr]
    ax1.bar(x, [r.get("R") or 0 for r in rr], color=cols)
    ax1.set_xticks(x, [str(int(r["id"])) for r in rr], fontsize=6)
    ax1.set_xlabel("region id, in rank order (made solid first, left)")
    ax1.set_ylabel("R = U_shear / U_bend (original solid)")
    ax1b = ax1.twinx()
    ax1b.plot(x, [100 * (r.get("U_share") or 0) for r in rr], "k.", ms=5)
    ax1b.set_ylabel("share of strain energy [%] (dots)")
    ax1.set_title("Shear-to-bending energy ratio per region (red = joint)", fontsize=9)
    for r in R:
        lam = r.get("lam") or 2.0
        mk = "o" if r["kind"] == "strut" else "s"
        fc = C_HOLLOW if r.get("final") == "hollow" else "white"
        ax2.scatter(lam, r.get("R") or 0, s=30 + 900 * (r.get("U_share") or 0), marker=mk, facecolor=fc,
                    edgecolor=C_JOINT if r["kind"] == "joint" else "k")
        ax2.annotate(str(int(r["id"])), (lam, r.get("R") or 0), fontsize=6, xytext=(3, 3), textcoords="offset points")
    ax2.set_xlabel("slenderness lambda = L / r_g (joints plotted at 2)")
    ax2.set_ylabel("R")
    ax2.set_title("Filled = hollow in the final design; marker area = strain-energy share", fontsize=9)
    fig.tight_layout()
    return fig


def fig_mass_sweep(c: Ctx):
    if not c.sweep:
        return None
    cfg = c.info["config"]["manufacturing"]
    k = np.array([r["k"] for r in c.sweep])
    s = np.array([r["s"] if r["s"] is not None else np.nan for r in c.sweep])
    hv = np.array([r.get("hollow_vol_share") or 0 for r in c.sweep])
    fig = new_figure(11, 5.5)
    ax = fig.add_subplot(1, 1, 1)
    ax.plot(k[:-1], s[:-1], "o-", color=C_HOLLOW, label="outer offset = hollow wall s")
    ax.plot(k[:-1], s[:-1] - cfg["wall_bias_mm"], ":", color=C_HOLLOW, label=f"as built (s - {cfg['wall_bias_mm']} mm)")
    ax.axhline(cfg["t_min_mm"], color=C_BAD, ls="--", label=f"t_min = {cfg['t_min_mm']} mm")
    ev = [r for r in c.evals]
    if ev:
        ax.plot([r["k"] for r in ev], [r["s"] for r in ev], "*", ms=10, color=C_ACC, label="analysed by FE")
    ax.set_xlabel("k = number of candidate regions kept solid (in rank order)")
    ax.set_ylabel("wall thickness [mm]")
    ax2 = ax.twinx()
    ax2.fill_between(k, 0, 100 * hv, color="0.85", alpha=0.6, step="mid")
    ax2.set_ylabel("hollowed share of the part volume [%]", color="0.4")
    ax.set_zorder(ax2.get_zorder() + 1)
    ax.patch.set_visible(False)
    ax.legend(fontsize=8, loc="upper right")
    ax.set_title("Constant-mass wall thickness for each design in the ranked family", fontsize=10)
    return fig


def fig_search(c: Ctx):
    if not c.evals:
        return None
    o = c.info["config"]["objective"]
    ev = sorted(c.evals, key=lambda r: r["k"])
    k = [r["k"] for r in ev]
    fig = new_figure(13, 6.5)
    ax1, ax2 = fig.subplots(1, 2)
    ax1.plot(k, [r["K_ratio"] for r in ev], "o-", color=C_SOLID, label="K / K0 (stiffness)")
    ax1.plot(k, [r["P_ratio"] for r in ev], "s-", color=C_HOLLOW, label="P / P0 (load capacity)")
    ax1.axhline(1.0, color="k", lw=0.7)
    ax1.axhline(o["floor_stiffness"], color=C_SOLID, ls=":", lw=0.8, label="stiffness floor")
    ax1.axhline(o["floor_capacity"], color=C_HOLLOW, ls=":", lw=0.8, label="capacity floor")
    for r in ev:
        ab = {"strength": "S", "Euler buckling": "E", "local buckling": "L"}.get(str(r["mode"]), "?")
        ax1.annotate(ab, (r["k"], r["P_ratio"]), fontsize=7, xytext=(3, -10), textcoords="offset points")
    ax1.set_xlabel("k (regions kept solid)")
    ax1.set_ylabel("ratio to the original solid")
    ax1.legend(fontsize=7)
    ax1.set_title("Performance of the analysed designs (governing mode: S strength, E Euler, L local)", fontsize=9)
    J = [r["J"] for r in ev]
    cols = [C_OK if r["acceptable"] in (True, "True", 1.0) else "0.6" for r in ev]
    ax2.bar(k, J, color=cols)
    ax2.axhline(np.log(1 + o["min_gain"]), color=C_BAD, ls="--", label=f"acceptance: ln(1 + {o['min_gain']})")
    ax2.axhline(0, color="k", lw=0.7)
    ax2.set_xlabel("k")
    ax2.set_ylabel("J = [w_K ln(K/K0) + w_P ln(P/P0)] / (w_K + w_P)")
    dec = c.info.get("decision", {})
    ax2.set_title(f"Objective (blue = acceptable); best k = {dec.get('best_k')}, "
                  f"{'accepted' if dec.get('accepted') else 'rejected -> solid kept'}", fontsize=9)
    ax2.legend(fontsize=7)
    fig.tight_layout()
    return fig


def fig_capacity(c: Ctx):
    base, best = c.info.get("baseline") or {}, c.info.get("best") or {}
    if not base:
        return None
    labels = ["strength", "Euler buckling", "local buckling"]
    keys = ["lam_s", "lam_E", "lam_L"]
    F = np.linalg.norm(c.info["geometry"]["force_N"])

    def vals(d):
        out = []
        for k in keys:
            v = d.get(k)
            v = np.inf if v in (None, "inf") else float(v)
            out.append(v * F)
        return np.array(out)
    b = vals(base)
    fig = new_figure(11, 5.5)
    ax = fig.add_subplot(1, 1, 1)
    x = np.arange(3)
    top = max([v for v in b if np.isfinite(v)] + ([v for v in vals(best) if np.isfinite(v)] if best else []) + [1])
    ax.bar(x - 0.2, np.where(np.isfinite(b), b, 0), 0.4, color=C_SOLID, label="original solid")
    if best:
        h = vals(best)
        ax.bar(x + 0.2, np.where(np.isfinite(h), h, 0), 0.4, color=C_HOLLOW, label="best hollow candidate")
        for xi, v in zip(x, h):
            if not np.isfinite(v):
                ax.text(xi + 0.2, top * 0.02, "n/a", ha="center", fontsize=7)
    for xi, v in zip(x, b):
        if not np.isfinite(v):
            ax.text(xi - 0.2, top * 0.02, "n/a", ha="center", fontsize=7)
    ax.set_xticks(x, labels)
    ax.set_ylabel("load at which the mode is reached [N]")
    ax.set_title("Load capacity by failure mode: the lowest bar of each design governs", fontsize=10)
    ax.legend(fontsize=8)
    return fig


def fig_table(c: Ctx):
    R = c.regions
    cols = ["id", "kind", "rank", "vol_mm3", "r_mm", "L_mm", "lam", "R", "U_share", "final"]
    head = ["id", "kind", "rank", "volume mm3", "r mm", "L mm", "lambda", "R", "U share", "final"]
    rows = []
    for r in R:
        row = []
        for k in cols:
            v = r.get(k)
            if k == "U_share" and v is not None:
                v = f"{100 * v:.1f}%"
            elif isinstance(v, float):
                v = f"{v:.3g}"
            row.append("-" if v in (None, "") else str(v))
        rows.append(row)
    per = 38
    figs = []
    for p in range(0, len(rows), per):
        fig = new_figure(8.27, 11.69)
        ax = fig.add_subplot(1, 1, 1)
        ax.axis("off")
        t = ax.table(cellText=rows[p:p + per], colLabels=head, loc="upper center", cellLoc="center")
        t.auto_set_font_size(False)
        t.set_fontsize(7)
        t.scale(1, 1.25)
        for (i, j), cell in t.get_celld().items():
            if i > 0 and rows[p + i - 1][-1] == "hollow":
                cell.set_facecolor("#dbe9f6")
        ax.set_title("Regions (blue rows are hollow in the final design)", fontsize=10)
        figs.append((f"06_regions_table_p{p // per + 1}", fig))
    return figs


def _slice_panel(ax, img, extent, title, theta=None):
    a = np.asarray(img, float).T
    solid = a <= 0
    ax.imshow(np.where(solid, 0.25, 1.0), origin="lower", extent=extent, cmap="gray", vmin=0, vmax=1, interpolation="nearest")
    if theta is not None:
        th = np.asarray(theta, float).T
        th = np.ma.masked_where(~solid, th)
        ax.imshow(th, origin="lower", extent=extent, cmap="coolwarm", vmin=0, vmax=1, alpha=0.35, interpolation="nearest")
    ax.contour(np.linspace(extent[0], extent[1], a.shape[1]), np.linspace(extent[2], extent[3], a.shape[0]), a, [0],
               colors="k", linewidths=0.5)
    ax.set_title(title, fontsize=8)
    ax.tick_params(labelsize=6)


def fig_sections(c: Ctx):
    A = c.A
    if "orig_slice_x" not in A:
        return None
    lo, h, cen = A["grid_lo"], float(A["grid_h"][0]), A["slice_centre"]
    has_best = "best_slice_x" in A
    fig = new_figure(15, 9 if has_best else 5)
    rows = 2 if has_best else 1
    names = ["x", "y", "z"]
    for j, ax_name in enumerate(names):
        o = A[f"orig_slice_{ax_name}"]
        keep = [i for i in range(3) if i != j]
        ext = [lo[keep[0]] - h / 2, lo[keep[0]] + h * (o.shape[0] - 0.5), lo[keep[1]] - h / 2, lo[keep[1]] + h * (o.shape[1] - 0.5)]
        ax = fig.add_subplot(rows, 3, j + 1)
        _slice_panel(ax, o, ext, f"original: {ax_name} = {cen[j]:.1f} mm")
        ax.set_xlabel("xyz"[keep[0]] + " [mm]", fontsize=7)
        ax.set_ylabel("xyz"[keep[1]] + " [mm]", fontsize=7)
        if has_best:
            ax = fig.add_subplot(rows, 3, j + 4)
            th = A.get(f"theta_slice_{ax_name}")
            _slice_panel(ax, A[f"best_slice_{ax_name}"], ext, f"best candidate: {ax_name} = {cen[j]:.1f} mm "
                         "(tint: red = solid region, blue = hollow)", th)
            ax.set_xlabel("xyz"[keep[0]] + " [mm]", fontsize=7)
            ax.set_ylabel("xyz"[keep[1]] + " [mm]", fontsize=7)
    fig.suptitle("Sections through the largest hollow region of the best candidate (or the largest joint); material grey, void white", fontsize=11)
    fig.tight_layout()
    return fig


def fig_render_final(c: Ctx):
    """The part coloured by region state in the best candidate: hollow blue, solid grey, mounts dark grey."""
    from scipy.spatial import cKDTree
    A = c.A
    best = c.info.get("best") or {}
    if "prev_V" not in A or "seg_xyz" not in A or not best:
        return None
    V, T = A["prev_V"], A["prev_T"]
    cen = V[T].mean(1)
    _, j = cKDTree(A["seg_xyz"]).query(cen)
    lab = A["seg_label"][j]
    hollow = set(int(i) for i in best.get("hollow") or [])
    mounts = {int(r["id"]) for r in c.regions if r["kind"] == "mount"}
    col = np.array([(0.17, 0.50, 0.80, 1) if l in hollow else ((0.35, 0.35, 0.35, 1) if l in mounts else (0.82, 0.82, 0.82, 1))
                    for l in lab])
    fig = new_figure(15, 5.6)
    for k, (el, az, title) in enumerate(VIEWS):
        ax = fig.add_subplot(1, 3, k + 1, projection="3d")
        _draw_mesh(ax, V, T, col)
        _equal3d(ax, V)
        ax.view_init(el, az)
        ax.set_title(title, fontsize=9)
    acc = c.info.get("decision", {}).get("accepted")
    fig.suptitle(f"Best candidate: hollow regions blue (wall {best.get('s', 0):.2f} mm), solid regions light grey, mounts dark grey"
                 f"{'' if acc else '  -  NOT adopted: final_design.stl is the original solid'}", fontsize=10)
    fig.tight_layout()
    return fig


def fig_voids(c: Ctx):
    best = c.info.get("best") or {}
    voids = best.get("voids") or []
    if not voids and not c.holes:
        return None
    fig = new_figure(8.27, 11.69)
    ax = fig.add_subplot(1, 1, 1)
    ax.axis("off")
    lines = [f"Enclosed voids of the best candidate: {len(voids)}"]
    lines += [f"  void {v['label']}: {v['volume_mm3']:.0f} mm3" for v in voids[:40]]
    lines += ["", f"Small voids filled before analysis: {best.get('filled_mm3', 0):.1f} mm3", "",
              "Suggested escape holes (lowest-stress wall points, outward direction n):"]
    for h in c.holes[:60]:
        lines.append(f"  void {int(h['void'])}: ({h['x']:.1f}, {h['y']:.1f}, {h['z']:.1f}) n = ({h['nx']:+.2f}, {h['ny']:+.2f}, "
                     f"{h['nz']:+.2f}), von Mises {h['vm_MPa']:.1f} MPa, d = {h['diameter_mm']} mm")
    lines += ["", "Escape holes are suggestions only: they are not cut into the mesh and their mass",
              "(about pi d^2 t / 4 each) is not included in the constant-mass balance."]
    ax.text(0.0, 1.0, "\n".join(lines), family="monospace", fontsize=7, va="top")
    return fig


def fig_timing(c: Ctx):
    t = c.info.get("timings") or {}
    if not t:
        return None
    keys = [k for k in t if k != "total"]
    fig = new_figure(9, 4.5)
    ax = fig.add_subplot(1, 1, 1)
    ax.barh(keys, [t[k] for k in keys], color=C_HOLLOW)
    ax.set_xlabel("seconds")
    ax.set_title(f"Run time by stage (total {t.get('total', 0):.0f} s)", fontsize=10)
    fig.tight_layout()
    return fig


# ----------------------------------------------------------------------------------------------
# decision robustness (offline: lambda_E scales with 1/K_eff^2 and lambda_s with sigma_f)
# ----------------------------------------------------------------------------------------------
def _f(v):
    if v in (None, "", "inf"):
        return np.inf
    return float(v)


def sensitivity(c: Ctx, K_list=(0.5, 0.6, 0.7, 0.85, 1.0), sf_scale=(0.8, 0.9, 1.0, 1.1, 1.2)):
    """Re-score the analysed designs for other Euler length factors and failure stresses, without new FE.

    Returns rows: K_eff, sigma_f, best k, its K/K0 and P/P0, J, accepted."""
    cfg = c.info["config"]
    K0, sf0 = cfg["capacity"]["K_eff"], cfg["material"]["strength_MPa"]
    o = cfg["objective"]
    base = c.info.get("baseline") or {}
    out = []
    for Ke in K_list:
        for sc in sf_scale:
            def lam(d):
                return min(_f(d.get("lam_s")) * sc, _f(d.get("lam_E")) * (K0 / Ke) ** 2, _f(d.get("lam_L")))
            lb = lam(base)
            best = None
            for r in c.evals:
                rK = float(r["K_ratio"])
                rP = lam(r) / lb
                J = (o["w_stiffness"] * np.log(rK) + o["w_capacity"] * np.log(rP)) / (o["w_stiffness"] + o["w_capacity"])
                ok = J > np.log(1 + o["min_gain"]) and rK >= o["floor_stiffness"] and rP >= o["floor_capacity"]
                key = (ok, J)
                if best is None or key > best[0]:
                    best = (key, dict(K_eff=Ke, sigma_f=sf0 * sc, k=int(r["k"]), K_ratio=rK, P_ratio=rP, J=J, accepted=ok))
            if best:
                out.append(best[1])
    return out


def fig_sensitivity(c: Ctx):
    if not c.evals or not c.info.get("baseline"):
        return None
    K_list, sf_scale = (0.5, 0.6, 0.7, 0.85, 1.0), (0.8, 0.9, 1.0, 1.1, 1.2)
    rows = sensitivity(c, K_list, sf_scale)
    import csv as _csv
    with open(os.path.join(c.dir, "sensitivity.csv"), "w", newline="", encoding="utf-8") as f:
        w = _csv.DictWriter(f, fieldnames=list(rows[0].keys()))
        w.writeheader()
        w.writerows(rows)
    P = np.array([[next(r for r in rows if r["K_eff"] == k and abs(r["sigma_f"] / c.info["config"]["material"]["strength_MPa"] - s) < 1e-9)["P_ratio"]
                   for s in sf_scale] for k in K_list])
    A = np.array([[next(r for r in rows if r["K_eff"] == k and abs(r["sigma_f"] / c.info["config"]["material"]["strength_MPa"] - s) < 1e-9)["accepted"]
                   for s in sf_scale] for k in K_list])
    fig = new_figure(10, 5.6)
    ax = fig.add_subplot(1, 1, 1)
    im = ax.imshow(np.log(P), cmap="RdBu", vmin=-np.log(3), vmax=np.log(3), origin="lower", aspect="auto")
    sf0 = c.info["config"]["material"]["strength_MPa"]
    ax.set_xticks(range(len(sf_scale)), [f"{sf0 * s:.1f}" for s in sf_scale])
    ax.set_yticks(range(len(K_list)), [f"{k:g}" for k in K_list])
    ax.set_xlabel("failure stress sigma_f [MPa]")
    ax.set_ylabel("Euler effective-length factor K_eff")
    for i in range(len(K_list)):
        for j in range(len(sf_scale)):
            ax.text(j, i, f"{'HOLLOW' if A[i, j] else 'solid'}\nP/P0 {P[i, j]:.2f}", ha="center", va="center", fontsize=7,
                    weight="bold" if A[i, j] else "normal")
    cfg = c.info["config"]
    ki = list(K_list).index(cfg["capacity"]["K_eff"]) if cfg["capacity"]["K_eff"] in K_list else None
    if ki is not None:
        ax.add_patch(__import__("matplotlib").patches.Rectangle((len(sf_scale) // 2 - 0.5, ki - 0.5), 1, 1, fill=False, lw=2, ec="k"))
    fig.colorbar(im, ax=ax, label="ln(P/P0) of the selected design")
    ax.set_title("Decision robustness: the analysed designs re-scored for other K_eff and sigma_f (no new FE); "
                 "box = this run", fontsize=9)
    return fig


FIGURES = [
    ("01_geometry_bcs", fig_geometry),
    ("02_segmentation", fig_segmentation),
    ("03_baseline_stress", fig_baseline_fe),
    ("04_regions_R_lambda", fig_regions),
    ("05_mass_sweep", fig_mass_sweep),
    ("06_regions_table", fig_table),
    ("07_search", fig_search),
    ("08_capacity_modes", fig_capacity),
    ("08b_decision_robustness", fig_sensitivity),
    ("09_sections", fig_sections),
    ("10_best_stress", fig_best_fe),
    ("11_design_map", fig_render_final),
    ("12_voids_escape_holes", fig_voids),
    ("13_timing", fig_timing),
]


def write_report(run_dir, dpi=130, verbose=True):
    """summary.json/.txt, plots/*.png and report.pdf for one run directory."""
    from matplotlib.backends.backend_pdf import PdfPages
    t0 = timer()
    c = Ctx(run_dir, dpi)
    log_path = os.path.join(c.plot_dir, "report_log.txt")
    if os.path.exists(log_path):
        os.remove(log_path)
    S = compute_summary(c)
    util.write_json(os.path.join(run_dir, "summary.json"), S)
    with open(os.path.join(run_dir, "summary.txt"), "w", encoding="utf-8") as f:
        f.write(summary_text(S, c.name))
    pdf_path = os.path.join(run_dir, "report.pdf")
    n = 0
    with PdfPages(pdf_path) as pdf:
        for name, fn in [("00_summary", lambda cc: fig_summary(cc, S))] + FIGURES:
            try:
                figs = fn(c)
            except Exception as e:
                c.log.append(f"figure '{name}' failed: {type(e).__name__}: {e}")
                with open(log_path, "a") as f:
                    f.write(f"--- {name} ---\n{traceback.format_exc()}\n")
                continue
            if figs is None:
                continue
            if not isinstance(figs, (list, tuple)):
                figs = [(name, figs)]
            for fname, fig in figs:
                try:
                    fig.savefig(os.path.join(c.plot_dir, f"{fname}.png"), dpi=dpi, bbox_inches="tight")
                    pdf.savefig(fig, bbox_inches="tight", dpi=min(dpi, 150))
                    n += 1
                except Exception as e:
                    c.log.append(f"saving '{fname}' failed: {e}")
        info = pdf.infodict()
        info["Title"] = f"dshell report: {c.name}"
        info["Subject"] = "Constant-mass variable-thickness shelling"
    if c.log:
        with open(log_path, "a") as f:
            f.write("\n".join(c.log) + "\n")
    if verbose:
        print(f"  [report] {c.name}: {n} pages -> {pdf_path} ({timer() - t0:.1f} s)"
              + (f"; {len(c.log)} figure(s) failed, see plots/report_log.txt" if c.log else ""))
    return pdf_path
