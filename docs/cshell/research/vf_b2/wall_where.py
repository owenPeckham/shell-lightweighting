"""Where are the B2 export wall samples below t_min (= t_min - b_v as built)? usage: python wall_where.py RUN_DIR"""
import json
import os
import sys

sys.path.insert(0, "/home/b6bl/owenpeckham.b6bl/shell-lightweighting/.claude/worktrees/cshell-stage2")
import numpy as np
import trimesh
from scipy.spatial import cKDTree
from cshell import collect as co, config as ccfg, export as cex
from cshell.optimise import Context


def main():
    rd = os.path.abspath(sys.argv[1])
    cfg = ccfg.load(json.load(open(os.path.join(rd, "config_used.json"))))
    t_min, b_v = cfg["manufacturing"]["t_min_mm"], cfg["manufacturing"]["wall_bias_mm"]
    rows = co.start_table(rd, cfg)
    best = co.pick_best(rows)
    ctx = Context(cfg, log=lambda *_: None)
    xi = np.load(os.path.join(rd, best["xi_file"]))
    vf = cex.vertex_fields(ctx, xi, best["D_comp"], float(cfg["grid"]["phases"][-1]["h_g"]))
    w_v = vf["a"] + vf["d"]
    act = ~vf["passive"]
    closed = vf["xt_a"] >= cex.ExportConfig.closed_thr
    print("design vertex wall a+d (active, open): min %.4f p1 %.4f median %.4f; n active open %d; t_min %.2f b_v %.2f" % (
        w_v[act & ~closed].min(), np.quantile(w_v[act & ~closed], 0.01), np.median(w_v[act & ~closed]), (act & ~closed).sum(), t_min, b_v))
    print("fraction of active open vertices with a+d < t_min + 0.01: %.3f ; < t_min: %.3f" % (
        (w_v[act & ~closed] < t_min + 0.01).mean(), (w_v[act & ~closed] < t_min - 1e-6).mean()))
    tm = trimesh.load(os.path.join(rd, "best_candidate.stl"), process=True)
    Vm, Fm = np.asarray(tm.vertices, float), np.asarray(tm.faces, np.int64)
    rng = np.random.default_rng(0)
    w = cex.wall_thickness(Vm, Fm, 20000, rng)
    opp = w["cos"] < -0.7
    P, N, t = w["P"][opp], w["N"][opp], w["t"][opp]
    low = t < t_min
    print("samples %d, below t_min %d (%.3f), min %.4f" % (len(t), low.sum(), low.mean(), t.min()))
    print("t quantiles 1,5,25,50:", np.round(np.quantile(t, [0.01, 0.05, 0.25, 0.5]), 3))
    # location along the axis
    xb = np.linspace(P[:, 0].min(), P[:, 0].max(), 9)
    print("x bins:", np.round(xb, 1))
    print("low per bin :", np.histogram(P[low, 0], xb)[0].tolist())
    print("all per bin :", np.histogram(P[:, 0], xb)[0].tolist())
    print("low-sample t quantiles (min, 25,50,75,max):", np.round(np.quantile(t[low], [0, .25, .5, .75, 1]), 3))
    # outer or inner surface: radial distance from the axis (B2 rod along x)
    r = np.linalg.norm(P[:, 1:3] - P[:, 1:3].mean(0), axis=1)
    for nm, sel in (("low", low), ("ok", ~low)):
        print(nm, "radius quantiles (5,50,95):", np.round(np.quantile(r[sel], [.05, .5, .95]), 2))
    # design wall at the nearest surface vertex
    tree = cKDTree(ctx.surf.V)
    dist, iv = tree.query(P)
    print("dist to nearest design vertex median %.3f" % np.median(dist))
    wd = w_v[iv]
    print("measured t - design (a+d) at nearest vertex: all median %.3f ; low median %.3f p5 %.3f p95 %.3f" % (
        np.median(t - wd), np.median((t - wd)[low]), *np.quantile((t - wd)[low], [.05, .95])))
    print("design a+d under the LOW samples: min %.3f median %.3f ; fraction of low samples with design wall >= t_min: %.3f" % (
        wd[low].min(), np.median(wd[low]), (wd[low] >= t_min - 1e-6).mean()))
    print("passive vertices under low samples: %.3f ; closed: %.3f" % (vf["passive"][iv][low].mean(), closed[iv][low].mean()))
    # holes
    H = np.genfromtxt(os.path.join(rd, "escape_holes.csv"), delimiter=",", names=True)
    hp = np.c_[np.atleast_1d(H["x"]), np.atleast_1d(H["y"]), np.atleast_1d(H["z"])]
    dh = np.linalg.norm(P[low][:, None] - hp[None], axis=2).min(1)
    print("low samples within 3 mm of a hole centre: %d ; nearest-hole distance quantiles" % (dh < 3).sum(), np.round(np.quantile(dh, [0, .25, .5]), 1))
    # curvature check: wall along the normal of a concentric annulus is exact, so compare with the mean of two neighbours
    # (a thin ring of local minima would indicate a marching-cubes dent)
    sub = low & (t > t_min - 0.1)
    print("share of low samples within 0.1 mm of t_min: %.3f" % (sub.sum() / max(low.sum(), 1)))


if __name__ == "__main__":
    main()
