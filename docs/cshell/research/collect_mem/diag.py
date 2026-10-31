"""Collect re-scoring diagnosis on the SMOKE_mbb run (diagnosis only; no cshell/ edits).

Usage (GPU job, from the repo root):  python docs/cshell/research/collect_mem/diag.py <run_dir> <out_dir> [--shift]
  1. Build the last-phase PhaseModel on the optimisation grid with geometry_device cpu (collect's choice) and gpu.
     For the run's best xi and the uniform-start xi: nominal V/V0-1, as-built rho, C, lam1, sigma_q on each chain;
     max |rho_cpu - rho_gpu|.
  2. CPU model: uniform shell built analytically on the band, psi = max(phi - t, -phi) (= export.baselines' definition,
     no STL), with/without the as-built increment and with the mount cells kept solid; load-node isolation stats
     (load nodes whose 8 adjacent elements all have rho < thr; share of the total force on them); PCG iterations.
  3. --shift: the same uniform shell and best on the 1/2 h_e shifted grid (collect.make_model) to reproduce the PCG
     stall.
Writes diag.json in out_dir.
"""
import copy, json, os, sys, time, traceback
import numpy as np

sys.path.insert(0, os.getcwd())
from cshell import collect as C, config as ccfg                       # noqa: E402
from cshell.optimise import Context, PhaseModel                       # noqa: E402

run_dir, out_dir = sys.argv[1], sys.argv[2]
do_shift = "--shift" in sys.argv
os.makedirs(out_dir, exist_ok=True)
OUT = {}


def rss():
    return int(open("/proc/self/status").read().split("VmRSS:")[1].split()[0]) / 2 ** 20   # GiB


def dump():
    json.dump(OUT, open(os.path.join(out_dir, "diag.json"), "w"), indent=1, default=float)


def log(*a):
    print(f"[diag {time.strftime('%H:%M:%S')} rss {rss():.1f} GiB]", *a, flush=True)


cfg = ccfg.load(json.load(open(os.path.join(run_dir, "config_used.json"))))
ctx = Context(cfg, log=log)
nv = ctx.surf.n_vertices
D_comp = cfg["design"]["outward_extent_initial_mm"]
rows = C.start_table(run_dir, cfg)
best = C.pick_best(rows)
D_comp = best["D_comp"]
xi_best = np.load(os.path.join(run_dir, best["xi_file"]))
xi_unif = np.concatenate([np.full(nv, 0.5), np.zeros(nv)])
log("best start", best["start"], "D_comp", D_comp)
ex = json.load(open(os.path.join(run_dir, "export_baselines", "uniform_shell_export.json")))
t_shell = 2.0497 - 0.0006844820191808103
for k in ("t_final_mm",):
    if k in ex:
        t_shell = ex[k]


def model_for(dev, r=None):
    c = copy.deepcopy(cfg)
    c["grid"]["geometry_device"] = dev
    c2 = copy.copy(ctx); c2.cfg = c
    ph = len(c["grid"]["phases"]) - 1
    if r is None:
        return PhaseModel(c2, ph, D_comp, log)
    c2.cfg = c
    with C.grid_shift(r["shift"]):
        return PhaseModel(c2, ph, D_comp, log)


def node_max_rho(model, rho):
    e = np.zeros(tuple(n + 2 for n in model.elem_shape), np.float32)
    e[1:-1, 1:-1, 1:-1] = rho.reshape(model.elem_shape)
    ns = model.node_shape
    out = np.zeros(ns, np.float32)
    for a in (0, 1):
        for b in (0, 1):
            for c in (0, 1):
                np.maximum(out, e[a:a + ns[0], b:b + ns[1], c:c + ns[2]], out=out)
    return out


def bc_stats(model, rho):
    nm = node_max_rho(model, rho).ravel()
    L = np.abs(model.load.reshape(-1, 3)).sum(1)
    Fx = model.fixed.reshape(-1, 3).any(1)
    ln = L > 0
    s = {}
    for thr in (1e-3, 0.05, 0.2):
        iso = ln & (nm < thr)
        s[f"load_nodes_rho<{thr}"] = int(iso.sum())
        s[f"load_share_rho<{thr}"] = float(L[iso].sum() / L.sum())
        s[f"fixed_nodes_with_rho>={thr}"] = int((Fx & (nm >= thr)).sum())
    s["n_load_nodes"], s["n_fixed_nodes"] = int(ln.sum()), int(Fx.sum())
    return s


def respond(model, rho, st, tag):
    t0 = time.time()
    try:
        r = C.raw_responses(model, rho, st)
        r["pcg"] = list(getattr(model.fe, "stats_solves", []))[:3]
    except Exception as e:
        r = dict(error=repr(e)[:300])
    r["t"] = time.time() - t0
    log(tag, {k: r.get(k) for k in ("C", "lam1", "sigma_q", "error", "pcg")})
    return r


def score(r, r0, st):
    if "error" in r:
        return None
    return C.score(r, r0, st)


# ---- 1. CPU vs GPU chain on the optimisation grid ------------------------------------------------------------------
rho_ref = {}
for dev in ("cpu", "gpu"):
    t0 = time.time()
    m = model_for(dev)
    OUT[f"{dev}_build"] = dict(seconds=time.time() - t0, rss_GiB=rss(), V0=float(C._h(m.V0)), bc=m.bc_info)
    st = C.scoring_settings(m)
    r0 = respond(m, C._h(m.rho0).astype(np.float64), st, f"{dev} solid")
    D = {"solid": dict(raw=r0)}
    for nm, xi in (("best", xi_best), ("uniform_start", xi_unif)):
        x = m.dm.xp.asarray(xi)
        V = m.dm.volume(m.dm.forward(x[:nv], x[nv:], as_built=False))
        rho = C._h(m.dm.rho(m.dm.forward(x[:nv], x[nv:], as_built=True))).astype(np.float64)
        if dev == "cpu":
            rho_ref[nm] = rho
            dr = None
        else:
            dr = float(np.abs(rho - rho_ref[nm]).max())
        raw = respond(m, rho, st, f"{dev} {nm}")
        D[nm] = dict(V_rel=V / float(C._h(m.V0)) - 1, raw=raw, score=score(raw, r0, st), max_drho_vs_cpu=dr,
                     bc=bc_stats(m, rho))
        log(dev, nm, D[nm]["V_rel"], D[nm]["score"], dr)
    OUT[f"{dev}_designs"] = D
    dump()
    if dev == "gpu":
        C._release(m); del m
        break
    # ---- 2. uniform shell on the band (CPU model, collect's CellData) -------------------------------------------
    cd = C.CellData(m, ctx)
    OUT["celldata_rss_GiB"] = rss()
    b_v, b_o = cfg["manufacturing"]["wall_bias_mm"], cfg["manufacturing"]["outer_bias_mm"]
    psiU = np.maximum(cd.phi - t_shell, -cd.phi)
    Dlt = C.asbuilt_increment(cd, np.full(nv, t_shell), np.zeros(nv), ctx.r_ma, b_v, b_o)
    mount_c = (cd.P @ ctx.mount.astype(float)) > 0.5
    U = {}
    for nm, psi in (("asbuilt", psiU + Dlt), ("nominal", psiU),
                    ("asbuilt_mount_solid", np.where(mount_c, cd.phi + b_o, psiU + Dlt))):
        chi, rho = C.chi_rho(cd, psi)
        Vn = float(C.chi_rho(cd, psiU)[0].sum()) * cd.h_g ** 3 + cd.V_extra
        raw = respond(m, rho, st, f"uniform shell {nm}")
        U[nm] = dict(V_rel_nominal=Vn / float(C._h(m.V0)) - 1, raw=raw, score=score(raw, r0, st), bc=bc_stats(m, rho),
                     V_extra=cd.V_extra)
        log("uniform", nm, U[nm]["score"], U[nm]["bc"])
        del chi, rho
    OUT["uniform_shell_band"] = U
    OUT["t_shell"] = t_shell
    dump()
    # ---- 2b. best design: nominal volume with the export's closure rule (xt_a >= 0.98 -> a = A_CLOSED via P) -------
    from cshell import export as cex
    vf = cex.vertex_fields(ctx, xi_best, D_comp, float(cfg["grid"]["phases"][-1]["h_g"]))
    V0f = float(C._h(m.V0))
    B = {}
    a_cl = np.where(vf["xt_a"] >= cex.ExportConfig.closed_thr, cex.A_CLOSED, vf["a"])
    for nm, a_v in (("vf_nominal", vf["a"]), ("closed_thr", a_cl)):
        psi = np.maximum(cd.phi - cd.P @ vf["d"], -cd.phi - cd.P @ a_v)
        chi, _ = C.chi_rho(cd, psi)
        B[nm] = float(chi.sum()) * cd.h_g ** 3 + cd.V_extra
        B[nm + "_rel"] = B[nm] / V0f - 1
        del chi, psi
    act = ~vf["passive"]
    B["n_closed_thr_active"] = int((act & (vf["xt_a"] >= 0.98)).sum())
    B["n_open_active"] = int((act & (vf["xt_a"] < 0.98)).sum())
    B["n_xt_a_0.5_0.98"] = int((act & (vf["xt_a"] >= 0.5) & (vf["xt_a"] < 0.98)).sum())
    void_w = 2 * (ctx.r_ma - vf["a"])
    B["n_open_void_lt_vmin"] = int((act & (vf["xt_a"] < 0.98) & (void_w < cfg["manufacturing"]["void_min_mm"])).sum())
    B["area_frac_open_void_lt_vmin"] = float((ctx.surf.A * (act & (vf["xt_a"] < 0.98) & (void_w < 3.0))).sum()
                                             / (ctx.surf.A * act).sum())
    OUT["best_volume_decomp"] = B
    log("best volume decomposition", B)
    dump()
    del cd, psiU, Dlt, mount_c
    C._release(m); del m

# ---- 3. shifted grid -----------------------------------------------------------------------------------------------
if do_shift:
    r = C.resolutions(cfg)[1]
    m = C.make_model(ctx, r, D_comp, log)
    st = C.scoring_settings(m)
    cd = C.CellData(m, ctx)
    r0 = respond(m, C._h(m.rho0).astype(np.float64), st, "shift solid")
    b_v, b_o = cfg["manufacturing"]["wall_bias_mm"], cfg["manufacturing"]["outer_bias_mm"]
    psiU = np.maximum(cd.phi - t_shell, -cd.phi)
    Dlt = C.asbuilt_increment(cd, np.full(nv, t_shell), np.zeros(nv), ctx.r_ma, b_v, b_o)
    S = {}
    _, rho = C.chi_rho(cd, psiU + Dlt)
    S["uniform_asbuilt"] = dict(raw=respond(m, rho, st, "shift uniform"), bc=bc_stats(m, rho))
    x = m.dm.xp.asarray(xi_best)
    rho = C._h(m.dm.rho(m.dm.forward(x[:nv], x[nv:], as_built=True))).astype(np.float64)
    S["best_param"] = dict(raw=respond(m, rho, st, "shift best"), bc=bc_stats(m, rho))
    for k in S:
        S[k]["score"] = score(S[k]["raw"], r0, st)
    OUT["shift"] = S
    dump()
log("done")
