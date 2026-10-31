"""Where does the export closure rule (P2) still add volume over the optimiser on the MBB smoke best (+0.505 %, CF_smoke)?
Cells where every positive-weight vertex is closed use a = A_CLOSED in the export, a = P a (≈ a^hi) in the optimiser.
Splits the extra material Σ(χ_rule − χ_opt) h³ by the class of the cell's vertices (mount, thin-passive, active closed)
and by depth −φ against the interpolated r_MA and a^hi.  usage: python closure_diag.py <run_dir>  (GPU job)"""
import json
import os
import sys

sys.path.insert(0, os.getcwd())
import cupy as cp
import numpy as np
from cshell import collect as C, config as ccfg, export as cex
from cshell.design import passive_thin
from cshell.optimise import Context

rd = os.path.abspath(sys.argv[1])
cfg = ccfg.load(json.load(open(os.path.join(rd, "config_used.json"))))
cfg["grid"]["geometry_device"] = "gpu"
ctx = Context(cfg, log=lambda *_: None)
best = C.pick_best(C.start_table(rd, cfg))
xi = np.load(os.path.join(rd, best["xi_file"]))
h_g = float(cfg["grid"]["phases"][-1]["h_g"])
vf = cex.vertex_fields(ctx, xi, best["D_comp"], h_g)
m = C.make_model(ctx, C.resolutions(cfg)[0], best["D_comp"], print)
D = m.disc
closed = vf["xt_a"] >= cex.ExportConfig.closed_thr
a_hi = ctx.r_ma + cfg["design"]["closure_gamma"] * h_g
thin = passive_thin(a_hi, cfg["manufacturing"]["t_min_mm"], cfg["manufacturing"]["void_min_mm"])
cls = np.where(ctx.mount, 0, np.where(thin, 1, 2))          # 0 mount, 1 thin passive, 2 active
print("closed vertices: mount", int((closed & (cls == 0)).sum()), "thin", int((closed & (cls == 1)).sum()),
      "active", int((closed & (cls == 2)).sum()), "of", len(closed), flush=True)
Th = cp.asarray(ctx.surf.T, cp.int32)
dv = {k: cp.asarray(v, cp.float64) for k, v in dict(d=vf["d"], a=vf["a"], cl=closed.astype(float), r=ctx.r_ma,
                                                     ahi=a_hi, c0=(cls == 0).astype(float),
                                                     c1=(cls == 1).astype(float)).items()}
eps = D.eps


def H(s):
    z = cp.clip(s / eps, -1.0, 1.0)
    return 0.5 + z * (15.0 / 16.0 - z * z * (5.0 / 8.0 - z * z * 3.0 / 16.0))


acc = {}
ch = 1 << 25
for s in range(0, D.n, ch):
    sl = slice(s, min(D.n, s + ch))
    phi = D.phi[sl].astype(cp.float64)
    v = Th[D.face[sl]]
    b = cp.maximum(D.bary[sl].reshape(-1, 3).astype(cp.float64), 0.0)
    w = b / b.sum(1, keepdims=True)

    def P(x):
        return (w * x[v]).sum(1)
    dc, ac, cc = P(dv["d"]), P(dv["a"]), P(dv["cl"])
    allc = cc >= 1 - 1e-9
    chi_o = H(-cp.maximum(phi - dc, -phi - ac))
    chi_r = H(-cp.maximum(phi - dc, -phi - cp.where(allc, cex.A_CLOSED, ac)))
    ex = chi_r - chi_o
    depth = -phi
    rr, ah = P(dv["r"]), P(dv["ahi"])
    c0, c1 = P(dv["c0"]), P(dv["c1"])
    k = cp.where(c0 > 0.5, 0, cp.where(c1 > 0.5, 1, 2))
    bins = cp.where(depth <= ah, 0, cp.where(depth <= rr * 1.5, 1, cp.where(depth <= rr * 3, 2, 3)))
    for kk in range(3):
        for bb in range(4):
            sel = (k == kk) & (bins == bb)
            acc[(kk, bb)] = acc.get((kk, bb), 0.0) + float(ex[sel].sum())
V0 = float(m.V0)
tot = sum(acc.values()) * D.h_g ** 3
print(f"extra volume of the rule {tot:.1f} mm3 = {tot / V0:+.3%} of V0", flush=True)
names = ["mount", "thin passive", "active"]
bn = ["depth <= P a_hi", "a_hi < depth <= 1.5 P r_MA", "1.5 < depth/r_MA <= 3", "depth > 3 P r_MA"]
for kk in range(3):
    for bb in range(4):
        x = acc[(kk, bb)] * D.h_g ** 3
        if abs(x) > 0.5:
            print(f"  {names[kk]:<13} {bn[bb]:<28} {x:9.1f} mm3 ({x / V0:+.3%})", flush=True)
