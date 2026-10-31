"""cshell report.pdf (THEORY_continuous.md section 9), built from the run directory only.

    python -m cshell report <run_dir>      ->  cshell.report.build(run_dir) -> <run_dir>/report.pdf

Style follows vshell/report.py: text pages first, one figure per page, each figure in its own try/except with failures
in plots/report_log.txt, 3D artists rasterised, PDF well below 5 MB.  A page whose inputs are missing is rendered as a
"not available" page (listing the missing files), so the page count is always 15 and numbering is stable.

Inputs (all optional): summary.json, history.csv, config_used.json, run_info.json, xi_<start>_phase<k>.npy, fields.vtp,
validation.csv, sensitivity.csv, escape_holes.csv, wall_thickness.csv, original.stl, final_design.stl.

summary.json keys read (ASSUMED until cshell/collect.py documents its schema; all optional, any other key is dumped as text):
  name, accepted (bool), reason (str), mass {target, final, error}, comparison [rows {label, K_K0, P_P0, J, governing,
  w_min, w_max}], resolutions [rows {label, J, K_K0, P_P0}], load_cases [rows {case, lam_s, lam_b, softmin}],
  resources {gpu_model, peak_gpu_GiB, gpu_hours}.
sensitivity.csv columns read: kappa_b, sigma_f (percent), and the first of J_star / J / accepted as the value.
fields.vtp point arrays read (case-insensitive aliases): xi, d, a, w, r_MA, d_hi, passive, W_bar, stress.
NOT available from any run-dir file: buckling mode shapes (page 8 renders "not available").
"""
from __future__ import annotations

import json
import os
import traceback
from timeit import default_timer as timer

import numpy as np

from vshell import report as vr
from vshell import util

PAGES = ["summary", "geometry", "design_space", "result_fields", "sections", "convergence", "stress", "buckling",
         "optimality", "capacity_robustness", "resolution", "validation", "voids", "resources", "provenance"]
SOURCE_FILES = ["__init__.py", "band.py", "cli.py", "config.py", "design.py", "export.py", "fe_sparse.py", "fea.py",
                "mma.py", "objective.py", "optimise.py", "report.py", "surface.py", "validate.py", "collect.py"]
REF_FILES = ["analysis.py", "gpu_analysis.py"]
ALIASES = {"xi": ["xi", "xi_a", "xi_d"], "d": ["d", "offset"], "a": ["a", "inset"], "w": ["w", "wall"],
           "r_MA": ["r_ma", "rma"], "d_hi": ["d_hi", "dhi"], "passive": ["passive"],
           "W_bar": ["w_bar", "wbar", "w_mean"], "stress": ["stress", "von_mises", "vm", "sigma_vm"]}


class Run:
    """Lazy, failure-tolerant access to the run directory."""

    def __init__(self, d, dpi=110):
        self.dir, self.dpi, self.log = d, dpi, []
        self.plot_dir = os.path.join(d, "plots")
        os.makedirs(self.plot_dir, exist_ok=True)
        self.summary = self._json("summary.json")
        self.cfg = self._json("config_used.json")
        self.info = self._json("run_info.json")
        self.hist = util.read_csv(self.p("history.csv"))
        self.name = self.summary.get("name") or self.cfg.get("name") or os.path.basename(os.path.normpath(d))
        self._f = None
        self.missing = []

    def p(self, f):
        return os.path.join(self.dir, f)

    def has(self, f):
        return os.path.isfile(self.p(f))

    def _json(self, f):
        try:
            with open(self.p(f), encoding="utf-8") as fh:
                x = json.load(fh)
            return x if isinstance(x, dict) else {}
        except Exception:
            return {}

    def csv(self, f):
        return util.read_csv(self.p(f))

    def fields(self):
        """fields.vtp as (pyvista mesh, {canonical name: array}) or None."""
        if self._f is None:
            self._f = False
            if self.has("fields.vtp"):
                import pyvista as pv
                m = pv.read(self.p("fields.vtp"))
                low = {k.lower(): k for k in list(m.point_data.keys())}
                F = {}
                for k, al in ALIASES.items():
                    for a in al:
                        if a in low:
                            F[k] = np.asarray(m.point_data[low[a]], float)
                            break
                self._f = (m, F)
        return self._f or None

    def field(self, *names):
        f = self.fields()
        if f is None:
            raise NA("fields.vtp")
        miss = [n for n in names if n not in f[1]]
        if miss:
            raise NA("fields.vtp arrays: " + ", ".join(miss))
        return f[0], [f[1][n] for n in names]


class NA(Exception):
    """Raised when an input is missing; the page becomes a 'not available' page."""


def _need(*conds_msgs):
    miss = [m for ok, m in conds_msgs if not ok]
    if miss:
        raise NA(", ".join(miss))


def _na_page(title, why):
    fig = vr.new_figure(8.27, 11.69)
    fig.text(0.5, 0.6, f"{title}\n\nnot available", ha="center", fontsize=14, weight="bold", color="0.4")
    fig.text(0.5, 0.52, f"missing input: {why}", ha="center", fontsize=9, color="0.4")
    return fig


def _text_page(title, lines, banner=None, colour="k"):
    fig = vr.new_figure(8.27, 11.69)
    y = 0.975
    if banner:
        fig.text(0.04, y, banner, fontsize=13, weight="bold", color=colour, va="top")
        y -= 0.03
    fig.text(0.04, y, title, fontsize=10, weight="bold", va="top")
    fig.text(0.04, y - 0.025, "\n".join(lines[:85]), family="monospace", fontsize=6.8, va="top")
    return fig


def _table(rows, cols, w=16):
    f = lambda v: "" if v is None else (f"{v:.4g}" if isinstance(v, float) else str(v))
    out = ["  ".join(f"{c[:w]:>{w}s}" for c in cols)]
    out += ["  ".join(f"{f(r.get(c))[:w]:>{w}s}" for c in cols) for r in rows]
    return out


def _col(rows, k):
    return np.array([np.nan if r.get(k) is None or isinstance(r.get(k), str) else r[k] for r in rows], float)


def _surf(ax, m, vals, title, cmap="viridis", view=(25, -60), vmin=None, vmax=None):
    """Rasterised triangle-surface colour map on a 3D axis (vertex values averaged per triangle)."""
    from mpl_toolkits.mplot3d.art3d import Poly3DCollection
    from matplotlib import cm, colors
    s = m.extract_surface(algorithm="dataset_surface").triangulate()
    V, T = np.asarray(s.points), np.asarray(s.faces).reshape(-1, 4)[:, 1:]
    ok = np.isfinite(vals)
    lo = np.nanmin(vals[ok]) if vmin is None and ok.any() else vmin
    hi = np.nanmax(vals[ok]) if vmax is None and ok.any() else vmax
    c = np.nanmean(vals[T], axis=1) if ok.any() else np.zeros(len(T))
    nrm = colors.Normalize(lo, hi if hi > lo else lo + 1)
    step = max(1, len(T) // 60000)
    col = Poly3DCollection(V[T[::step]], rasterized=True, linewidths=0)
    col.set_facecolor(__import__("matplotlib").colormaps[cmap](nrm(c[::step])))
    ax.add_collection3d(col)
    vr._equal3d(ax, V)
    ax.view_init(*view)
    ax.set_title(title, fontsize=8)
    ax.set_axis_off()
    return cm.ScalarMappable(nrm, cmap)


def _three_views(R, names, titles, cmaps=None, suptitle=""):
    m, vals = R.field(*names)
    fig = vr.new_figure(15, 4.6 * len(names))
    n = len(names)
    for i, (v, t) in enumerate(zip(vals, titles)):
        for j, (el, az, _) in enumerate(vr.VIEWS):
            ax = fig.add_subplot(n, 3, 3 * i + j + 1, projection="3d")
            sm = _surf(ax, m, v, f"{t}  ({vr.VIEWS[j][2]})", (cmaps or ["viridis"] * n)[i], (el, az))
        fig.colorbar(sm, ax=fig.axes[3 * i:3 * i + 3], shrink=0.7, label=t)
    fig.suptitle(suptitle, fontsize=11)
    return fig


# ---------------------------------------------------------------------------------------------------------------------
# pages
# ---------------------------------------------------------------------------------------------------------------------
def p_summary(R):
    S = R.summary
    _need((bool(S), "summary.json"))
    acc = S.get("accepted", S.get("decision", {}).get("accepted") if isinstance(S.get("decision"), dict) else None)
    reason = S.get("reason") or (S.get("decision") or {}).get("reason", "") if S else ""
    banner = {True: "HOLLOW DESIGN ACCEPTED", False: "ORIGINAL SOLID KEPT"}.get(acc, "DECISION UNKNOWN")
    lines = []
    for ln in str(reason).split("\n"):
        lines += vr.textwrap.wrap(ln, 110) or [""]
    lines += ["", "Mass check: " + vr._fmt(S.get("mass")), "", "Comparison (solid, uniform shell, vshell best, cshell starts, export)"]
    comp = S.get("comparison") or []
    if comp:
        lines += _table(comp, ["label", "K_K0", "P_P0", "J", "governing", "w_min", "w_max"], 13)
    shown = {"name", "accepted", "reason", "mass", "comparison", "decision"}
    lines += ["", "Run settings / other summary keys"]
    for k, v in S.items():
        if k not in shown:
            lines += vr.textwrap.wrap(f"  {k}: {vr._fmt(v)}", 120)
    o = R.cfg.get("optimiser", {})
    if o:
        lines += ["", "optimiser: " + vr._fmt(o), "grid: " + vr._fmt(R.cfg.get("grid"))]
    return _text_page(f"{R.name}: summary", lines, banner, vr.C_OK if acc else vr.C_SOLID)


def p_geometry(R):
    """Original geometry with loads (config) and supports; reuses vshell.report.fig_geometry when its arrays exist."""
    if R.has("arrays.npz") and R.info.get("geometry"):
        c = vr.Ctx.__new__(vr.Ctx)
        c.A, c.info, c.name = dict(np.load(R.p("arrays.npz"))), R.info, R.name
        f = vr.fig_geometry(c)
        if f is not None:
            return f
    _need((R.has("original.stl"), "original.stl"))
    import trimesh
    tm = trimesh.load(R.p("original.stl"), force="mesh")
    V, T = np.asarray(tm.vertices), np.asarray(tm.faces)
    fig = vr.new_figure(15, 5.6)
    loads = R.cfg.get("loads", [])
    F = np.array(loads[0]["force_N"], float) if loads and "force_N" in loads[0] else None
    for j, (el, az, t) in enumerate(vr.VIEWS):
        ax = fig.add_subplot(1, 3, j + 1, projection="3d")
        vr._draw_mesh(ax, V, T)
        if F is not None:
            span = np.ptp(V, 0).max()
            d = -F / (np.linalg.norm(F) + 1e-12) * 0.25 * span
            ax.quiver(*(V.mean(0) + d), *(-d), color=vr.C_BAD, lw=2, arrow_length_ratio=0.25)
        vr._equal3d(ax, V)
        ax.view_init(el, az)
        ax.set_title(t, fontsize=9)
    sup = R.cfg.get("supports", [])
    fig.suptitle(f"{R.name}: original geometry; loads {[l.get('force_N') for l in loads]} N (arrow, indicative placement); "
                 f"supports {[s.get('type') for s in sup]} (face selection from config only; arrays.npz absent)", fontsize=9)
    return fig


def p_design_space(R):
    return _three_views(R, ["r_MA", "d_hi", "passive"], ["r_MA [mm]", "d_hi [mm]", "passive"],
                        ["viridis", "viridis", "gray_r"], f"{R.name}: design space on Gamma_h")


def p_result_fields(R):
    m, (w, d, a) = R.field("w", "d", "a")
    fig = _three_views(R, ["w", "d", "a"], ["wall w [mm]", "offset d [mm]", "inset a [mm]"], None,
                       f"{R.name}: result fields")
    hist = None
    if R.has("wall_thickness.csv"):
        rows = R.csv("wall_thickness.csv")
        ks = [k for k in rows[0] if k.lower() not in ("x", "y", "z", "id", "vertex")] if rows else []
        hist = {k: _col(rows, k) for k in ks}
    fig.subplots_adjust(bottom=0.14)
    ax = fig.add_axes([0.06, 0.02, 0.25, 0.07])
    ax.hist(w[np.isfinite(w)], bins=60, alpha=0.6, label="w (fields.vtp, nominal)")
    for k, v in (hist or {}).items():
        v = v[np.isfinite(v)]
        if v.size:
            ax.hist(v, bins=60, histtype="step", label=f"{k} (wall_thickness.csv)")
    ax.set_xlabel("w [mm]", fontsize=6)
    ax.tick_params(labelsize=6)
    ax.legend(fontsize=5)
    ax.set_title("w histogram", fontsize=7)
    return fig


def p_sections(R):
    _need((R.has("original.stl"), "original.stl"), (R.has("final_design.stl"), "final_design.stl"))
    import trimesh
    A = trimesh.load(R.p("original.stl"), force="mesh")
    B = trimesh.load(R.p("final_design.stl"), force="mesh")
    pts = {"largest joint (max w)": None, "most-changed strut (max |d|)": None}
    f = R.fields()
    if f is not None and "w" in f[1]:
        pts["largest joint (max w)"] = np.asarray(f[0].points)[np.nanargmax(f[1]["w"])]
    if f is not None and "d" in f[1]:
        pts["most-changed strut (max |d|)"] = np.asarray(f[0].points)[np.nanargmax(np.abs(f[1]["d"]))]
    pts = {k: (v if v is not None else A.centroid) for k, v in pts.items()}
    fig = vr.new_figure(15, 9)
    for r, (lab, c0) in enumerate(pts.items()):
        for j, ax_i in enumerate(range(3)):
            ax = fig.add_subplot(2, 3, 3 * r + j + 1)
            n = np.eye(3)[ax_i]
            for mesh, col, nm in ((A, "0.5", "original"), (B, "tab:red", "final")):
                try:
                    sec = mesh.section(plane_origin=c0, plane_normal=n)
                    if sec is None:
                        continue
                    ij = [i for i in range(3) if i != ax_i]
                    for e in sec.entities:
                        q = sec.vertices[e.points][:, ij]
                        ax.plot(q[:, 0], q[:, 1], color=col, lw=0.8, label=nm)
                except Exception as e:
                    R.log.append(f"section {lab}/{ax_i}: {e}")
            ax.set_aspect("equal")
            ax.set_title(f"{lab}: cut normal {'xyz'[ax_i]}", fontsize=8)
            h, l = ax.get_legend_handles_labels()
            if h:
                ax.legend(dict(zip(l, h)).values(), dict(zip(l, h)).keys(), fontsize=6)
    fig.suptitle("sections: original (grey) vs final (red)")
    return fig


def p_convergence(R):
    _need((bool(R.hist), "history.csv"))
    keys = [("J", "J"), ("C_C0", "C/C0"), ("lam_s_ratio", "lam_s/lam_s0"), ("lam_b_ratio", "lam_b/lam_b0"),
            ("mass_err", "V/V0 - 1"), ("G_s", "G_s"), ("max_dxi", "max|dxi|")]
    starts = list(dict.fromkeys(r["start"] for r in R.hist))
    fig = vr.new_figure(15, 9)
    for i, (k, lab) in enumerate(keys):
        ax = fig.add_subplot(3, 3, i + 1)
        for s in starts:
            rows = [r for r in R.hist if r["start"] == s]
            y = _col(rows, k)
            x = np.arange(len(rows))
            if np.isfinite(y).any():
                ax.plot(x, y, lw=0.9, label=str(s))
            ph = _col(rows, "phase")
            for b in np.where(np.diff(ph) != 0)[0] + 1 if len(ph) > 1 else []:
                ax.axvline(b, color="0.7", lw=0.5)
        if k in ("G_s", "max_dxi"):
            ax.set_yscale("log")
        ax.set_title(lab, fontsize=9)
        ax.set_xlabel("iteration (phase boundaries grey)", fontsize=7)
    fig.axes[0].legend(fontsize=6)
    fig.tight_layout()
    return fig


def p_stress(R):
    return _three_views(R, ["stress"], ["von Mises [MPa]"], ["inferno"],
                        f"{R.name}: material von Mises on final design "
                        "(solid-baseline stress not stored in the run dir)")


def p_buckling(R):
    raise NA("buckling mode shapes (no run-dir file stores eigenvectors)")


def p_optimality(R):
    m, (Wb,) = R.field("W_bar")
    lam = None
    for r in R.hist[::-1]:
        if r.get("beta") is not None:
            break
    fig = vr.new_figure(15, 5.5)
    ax = fig.add_subplot(1, 2, 1)
    ok = np.isfinite(Wb)
    f = R.fields()[1]
    x = f.get("w", np.arange(len(Wb)))
    ax.scatter(x[ok], Wb[ok], s=1, rasterized=True)
    Lam = R.summary.get("Lambda") or R.summary.get("lambda")
    if isinstance(Lam, (int, float)):
        ax.axhline(Lam / 2, color="r", label="Lambda/2")
        ax.legend()
    ax.set_xlabel("w [mm]" if "w" in f else "vertex")
    ax.set_ylabel("W_bar")
    ax.set_title("W_bar on free surface (Lambda/2 shown if in summary.json)", fontsize=9)
    ax3 = fig.add_subplot(1, 2, 2, projection="3d")
    sm = _surf(ax3, m, Wb, "W_bar", "magma")
    fig.colorbar(sm, ax=ax3, shrink=0.7)
    return fig


def p_capacity(R):
    lc = R.summary.get("load_cases")
    sens = R.csv("sensitivity.csv")
    _need((bool(lc) or bool(sens), "summary.json:load_cases and sensitivity.csv"))
    figs = []
    fig = vr.new_figure(11, 8)
    if lc:
        ax = fig.add_subplot(2, 1, 1)
        x = np.arange(len(lc))
        for i, k in enumerate(["lam_s", "lam_b", "softmin"]):
            ax.bar(x + 0.27 * (i - 1), _col(lc, k), 0.27, label=k)
        ax.set_xticks(x)
        ax.set_xticklabels([str(r.get("case")) for r in lc], fontsize=7)
        ax.legend(fontsize=7)
        ax.set_title("capacity per load case", fontsize=9)
    if sens:
        kc = [k for k in sens[0] if k.lower().startswith("kappa")]
        sc = [k for k in sens[0] if k.lower().startswith("sigma")]
        vc = [k for k in ("J_star", "J", "accepted") if k in sens[0]]
        if kc and sc and vc:
            kb, sf, v = _col(sens, kc[0]), _col(sens, sc[0]), _col(sens, vc[0])
            ks, ss = np.unique(kb[np.isfinite(kb)]), np.unique(sf[np.isfinite(sf)])
            G = np.full((len(ss), len(ks)), np.nan)
            for a, b, c in zip(kb, sf, v):
                if np.isfinite(a) and np.isfinite(b):
                    G[np.searchsorted(ss, b), np.searchsorted(ks, a)] = c
            ax = fig.add_subplot(2, 1, 2)
            im = ax.imshow(G, origin="lower", aspect="auto")
            ax.set_xticks(range(len(ks)))
            ax.set_xticklabels(ks)
            ax.set_yticks(range(len(ss)))
            ax.set_yticklabels(ss)
            ax.set_xlabel("kappa_b")
            ax.set_ylabel("sigma_f change [%]")
            fig.colorbar(im, ax=ax, label=vc[0])
            ax.set_title("decision re-scored (no new FE: lam_b ~ kappa_b, lam_s ~ sigma_f)", fontsize=9)
    fig.tight_layout()
    return fig


def p_resolution(R):
    rows = R.summary.get("resolutions")
    _need((bool(rows), "summary.json:resolutions"))
    fig = vr.new_figure(8.27, 11.69)
    ax = fig.add_subplot(2, 1, 1)
    x = np.arange(len(rows))
    ax.plot(x, _col(rows, "J"), "o-")
    ax.set_xticks(x)
    ax.set_xticklabels([str(r.get("label")) for r in rows], rotation=30, fontsize=7)
    ax.set_ylabel("J")
    ax2 = fig.add_subplot(2, 1, 2)
    ax2.plot(x, _col(rows, "K_K0"), "s-", label="K/K0")
    ax2.plot(x, _col(rows, "P_P0"), "^-", label="P/P0")
    ax2.legend()
    ax2.set_xticks(x)
    ax2.set_xticklabels([str(r.get("label")) for r in rows], rotation=30, fontsize=7)
    fig.tight_layout()
    return fig


def p_validation(R):
    rows = R.csv("validation.csv")
    _need((bool(rows), "validation.csv"))
    cols = list(rows[0])
    return _text_page("Body-fitted against voxel validation", _table(rows, cols, max(8, 120 // len(cols))))


def p_voids(R):
    holes = R.csv("escape_holes.csv")
    wt = R.csv("wall_thickness.csv")
    _need((bool(holes) or bool(wt), "escape_holes.csv, wall_thickness.csv"))
    lines = [f"escape holes: {len(holes)}"]
    if holes:
        lines += _table(holes[:60], list(holes[0]), max(8, 120 // len(holes[0])))
    fig = _text_page("Voids and escape holes", lines)
    if wt:
        ks = [k for k in wt[0] if k.lower() not in ("x", "y", "z", "id", "vertex")]
        if ks:
            ax = fig.add_axes([0.1, 0.08, 0.8, 0.2])
            for k in ks[:3]:
                v = _col(wt, k)
                ax.hist(v[np.isfinite(v)], bins=50, histtype="step", label=k)
            ax.legend(fontsize=7)
            ax.set_title("minimum-wall data (wall_thickness.csv)", fontsize=8)
    return fig


def p_resources(R):
    _need((bool(R.hist), "history.csv"))
    ks = ["t_geometry", "t_fe", "t_mma", "t_restore"]
    starts = list(dict.fromkeys(r["start"] for r in R.hist))
    fig = vr.new_figure(11, 8)
    ax = fig.add_subplot(2, 1, 1)
    bot = None
    tot = 0.0
    for r_ in R.hist:
        tot += sum((r_.get(k) or 0.0) for k in ks if isinstance(r_.get(k), (int, float)))
    y = {k: np.nan_to_num(_col(R.hist, k)) for k in ks}
    x = np.arange(len(R.hist))
    bot = np.zeros(len(x))
    for k in ks:
        ax.bar(x, y[k], bottom=bot, width=1.0, label=k)
        bot += y[k]
    ax.legend(fontsize=7)
    ax.set_xlabel("iteration (all starts concatenated)")
    ax.set_ylabel("s")
    ax2 = fig.add_subplot(2, 1, 2)
    ax2.plot(x, _col(R.hist, "gpu_mem_GiB"))
    ax2.set_ylabel("GPU mem [GiB]")
    ax2.set_xlabel("iteration")
    gpu = (R.summary.get("resources") or {})
    hours = gpu.get("gpu_hours")
    fig.suptitle(f"starts {starts}; summed iteration time {tot / 3600:.2f} h"
                 f"; peak mem {np.nanmax(_col(R.hist, 'gpu_mem_GiB')):.1f} GiB"
                 + (f"; GPU {gpu.get('gpu_model')}; GPU-hours {hours}" if gpu else "")
                 + "; jobs " + ",".join(str(a.get("slurm_job")) for a in R.info.get("allocations", [])), fontsize=9)
    return fig


def p_provenance(R):
    import socket
    root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    pkg = os.path.join(root, "cshell")
    prov = util.provenance(R.cfg.get("step_file"), R.p("config_used.json") if R.has("config_used.json") else None)
    git = prov["git"]
    src = {f"cshell/{f}": util.sha256_of(os.path.join(pkg, f)) for f in SOURCE_FILES if os.path.isfile(os.path.join(pkg, f))}
    for f in REF_FILES:
        pth = os.path.join(root, "continuous", "reference", f)
        src[f"continuous/reference/{f}"] = util.sha256_of(pth) if os.path.isfile(pth) else None
    pk = dict(prov["packages"])
    from importlib.metadata import version
    for n in ("cupy", "pyvista", "scipy", "cupy-cuda12x"):
        try:
            pk[n] = version(n)
        except Exception:
            pass
    cuda = "n/a"
    try:
        import cupy as cp
        cuda = f"driver {cp.cuda.runtime.driverGetVersion()}, runtime {cp.cuda.runtime.runtimeGetVersion()}"
    except Exception as e:
        cuda = f"unavailable ({type(e).__name__})"
    allocs = R.info.get("allocations", [])
    lines = [f"git commit  {git['commit']}   dirty: {git['dirty']}   (state when this report was built)",
             f"report built on host {socket.gethostname()} at {prov['timestamp_utc']}; CUDA {cuda}",
             "run hosts / jobs (run_info.json): " + "; ".join(f"{a.get('host')} job {a.get('slurm_job')}" for a in allocs),
             "STEP: " + vr._fmt(prov.get("step_file")), "config_used.json: " + vr._fmt(prov.get("config_file")),
             "", "source SHA-256:"] + [f"  {k:42s} {v}" for k, v in src.items()] + ["", "packages:"] + \
            [f"  {k:20s} {v}" for k, v in pk.items() if v]
    return _text_page("Provenance", lines)


PAGE_FUNCS = [("01_summary", p_summary), ("02_geometry_bcs", p_geometry), ("03_design_space", p_design_space),
              ("04_result_fields", p_result_fields), ("05_sections", p_sections), ("06_convergence", p_convergence),
              ("07_stress", p_stress), ("08_buckling", p_buckling), ("09_optimality", p_optimality),
              ("10_capacity_robustness", p_capacity), ("11_resolution", p_resolution), ("12_validation", p_validation),
              ("13_voids", p_voids), ("14_resources", p_resources), ("15_provenance", p_provenance)]


def build(run_dir: str, dpi: int = 110, verbose: bool = True) -> str:
    """Write <run_dir>/report.pdf (15 pages) and plots/*.png; return the path."""
    from matplotlib.backends.backend_pdf import PdfPages
    t0 = timer()
    R = Run(run_dir, dpi)
    log_path = os.path.join(R.plot_dir, "report_log.txt")
    if os.path.exists(log_path):
        os.remove(log_path)
    pdf_path = os.path.join(run_dir, "report.pdf")
    n = 0
    with PdfPages(pdf_path) as pdf:
        for (name, fn), title in zip(PAGE_FUNCS, PAGES):
            try:
                figs = fn(R)
            except NA as e:
                figs = _na_page(f"{name}: {title}", str(e))
            except Exception as e:
                R.log.append(f"figure '{name}' failed: {type(e).__name__}: {e}")
                with open(log_path, "a") as f:
                    f.write(f"--- {name} ---\n{traceback.format_exc()}\n")
                figs = _na_page(f"{name}: {title}", f"error {type(e).__name__}: {e} (see plots/report_log.txt)")
            if not isinstance(figs, (list, tuple)):
                figs = [(name, figs)]
            for fname, fig in figs:
                try:
                    fig.savefig(os.path.join(R.plot_dir, f"{fname}.png"), dpi=dpi, bbox_inches="tight")
                    pdf.savefig(fig, bbox_inches="tight", dpi=min(dpi, 150))
                    n += 1
                except Exception as e:
                    R.log.append(f"saving '{fname}' failed: {e}")
        pdf.infodict()["Title"] = f"cshell report: {R.name}"
    if R.log:
        with open(log_path, "a") as f:
            f.write("\n".join(R.log) + "\n")
    if verbose:
        print(f"  [report] {R.name}: {n} pages -> {pdf_path} ({timer() - t0:.1f} s)"
              + (f"; {len(R.log)} issue(s), see plots/report_log.txt" if R.log else ""))
    return pdf_path
