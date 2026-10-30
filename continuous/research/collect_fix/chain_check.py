"""P5 acceptance on the MBB smoke (h_g 0.1): the GPU re-scoring chain (DeviceDesignMap.rescore) against the CPU chain
(host P / φ) on the optimisation grid, for the solid and the exported uniform shell (CF_smoke): ρ, V and J.
usage: python chain_check.py <run_dir>   (GPU job, ~90 GiB host for the CPU-chain model)"""
import copy
import json
import os
import sys

sys.path.insert(0, os.getcwd())
import numpy as np
from cshell import collect as C, config as ccfg, export as cex
from cshell.optimise import Context

rd = os.path.abspath(sys.argv[1])
cfg = ccfg.load(json.load(open(os.path.join(rd, "config_used.json"))))
ctx = Context(cfg, log=lambda *_: None)
best = C.pick_best(C.start_table(rd, cfg))
D_comp = best["D_comp"]
h_g = float(cfg["grid"]["phases"][-1]["h_g"])
b_v, b_o = cfg["manufacturing"]["wall_bias_mm"], cfg["manufacturing"]["outer_bias_mm"]
ex = json.load(open(os.path.join(rd, "summary.json")))["export"]
fu = cex.uniform_shell_fields(ctx, ex["uniform_shell_t_mm"], h_g)
f = C.design_fields(fu["d"], fu["a"], fu["closed"], ctx.r_ma, b_v)
c0 = np.round(0.5 * (ctx.V.min(0) + ctx.V.max(0)), 3)
out = {}
for dev in ("gpu", "cpu"):
    c2 = copy.copy(ctx)
    c2.cfg = copy.deepcopy(cfg)
    c2.cfg["grid"]["geometry_device"] = dev
    m = C.make_model(c2, C.resolutions(cfg)[0], D_comp, lambda *_: None)
    soup = C.stl_soup(os.path.join(rd, "uniform_shell.stl"), c0)
    phiE = C.mesh_phi(m, soup)
    del soup
    rho_s, V_s = C.cell_rescore(m, None, None, b_o, solid=True)
    rho_u, V_u = C.cell_rescore(m, phiE, f, b_o)
    del phiE
    if dev == "gpu":
        m.disc.release()
    st = C.scoring_settings(m)
    r0 = C.raw_responses(m, rho_s, st)
    ru = C.raw_responses(m, rho_u, st)
    out[dev] = dict(rho_s=rho_s, rho_u=rho_u, V_s=V_s, V_u=V_u, J=C.score(ru, r0, st)["J"], C0=r0["C"], Cu=ru["C"])
    print(dev, {k: v for k, v in out[dev].items() if not k.startswith("rho")}, f"rss {C.rss_gib():.1f} GiB", flush=True)
    C.free_model(m)
    m = None
a, b = out["gpu"], out["cpu"]
print(f"max|drho| solid {np.abs(a['rho_s'] - b['rho_s']).max():.2e} uniform shell {np.abs(a['rho_u'] - b['rho_u']).max():.2e}")
print(f"V rel diff solid {a['V_s'] / b['V_s'] - 1:+.2e} uniform shell {a['V_u'] / b['V_u'] - 1:+.2e}")
print(f"C rel diff solid {a['C0'] / b['C0'] - 1:+.2e} uniform shell {a['Cu'] / b['Cu'] - 1:+.2e}")
print(f"J gpu {a['J']:.12f} cpu {b['J']:.12f} diff {a['J'] - b['J']:+.2e}  (acceptance |dJ| <= 1e-10)")
