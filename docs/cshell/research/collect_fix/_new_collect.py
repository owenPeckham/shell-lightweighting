def export_verdict(ex: dict | None, vol_tol: float = 2e-3, G_s: float | None = None):
    """(ok, reason) of an export summary (PLAN_collect_fixes P3, D2): mass correction bracketed and within δ_fail,
    watertight, |V_STL/V₀ − 1| ≤ vol_tol. A failed bracket after a large void_min opening fill is reported as the
    optimiser's non-convergence (§8.2/§8.4 behave as specified), not as a re-scoring error."""
    if ex is None:
        return True, ""
    why = []
    mc, mesh, cl = ex.get("mass_correction", {}), ex.get("mesh", {}), ex.get("cleanup", {}) or {}
    V0 = ex.get("V0_mm3") or float("nan")
    if not mc.get("ok", False):
        bf = mc.get("bracket_failed") or {}
        why.append(f"mass correction did not bracket (V/V0-1 before correction {mc.get('rel_before', float('nan')):+.2%}"
                   + (f", limit ±{bf['limit']:g} mm" if "limit" in bf else "") + ")")
        fill = cl.get("opening_removed_mm3")
        if fill and fill / V0 > 0.01:
            why.append(f"the void_min opening filled {fill / V0:+.1%} of V0: the design is not converged in the "
                       f"sliver measure" + (f" (final G_s = {G_s:.3g})" if G_s is not None and np.isfinite(G_s) else "")
                       + " - optimiser non-convergence, not a re-scoring error")
    elif mc.get("fail", False):
        why.append(f"mass correction |delta| = {abs(mc.get('delta', float('nan'))):.3f} mm above delta_fail")
    if not mesh.get("watertight_manifold", False):
        why.append(f"mesh not watertight ({mesh.get('bodies')} bodies, {mesh.get('duplicate_directed_edges')} "
                   "duplicated directed edges)")
    rel = mesh.get("volume_rel")
    if rel is None or not np.isfinite(rel) or abs(rel) > vol_tol:
        why.append(f"V_STL/V0 - 1 = {rel if rel is not None else float('nan'):+.2%} (tolerance {vol_tol:.1%})")
    ok = bool(ex.get("ok", False)) and not why
    if not ok and not why:
        why.append("export checks failed")
    return ok, "; ".join(why)


def collect(run_dir: str, allow_incomplete: bool = False, validate: bool | None = None, export_kw: dict | None = None,
            log=None, geometry_device: str | None = None) -> dict:
    """Merge the starts of `run_dir`, export, re-score, validate and decide (module docstring). Returns summary.
    geometry_device: "cpu" | "gpu" | "auto" for the re-scoring cell maps (default: config grid.geometry_device)."""
    from . import export as cex
    from .optimise import Context, PeakGPU
    run_dir = os.path.abspath(run_dir)
    logf = open(os.path.join(run_dir, "collect_log.txt"), "a")

    def _log(msg):
        print(msg, flush=True)
        logf.write(msg + "\n"); logf.flush()
    log = log or _log
    T0 = time.time()
    tim, mem = {}, {}
    pk_gpu = PeakGPU()
    try:
        pk_gpu.start()
    except Exception:
        pk_gpu = None

    def memo(stage):
        mem[stage] = round(rss_gib(), 2)
        log(f"[collect] host RSS {mem[stage]:.1f} GiB after {stage}")
    cfg = json.load(open(os.path.join(run_dir, "config_used.json")))
    from . import config as ccfg
    cfg = ccfg.load(cfg)                                  # re-apply defaults (keys added after the run)
    if geometry_device is not None:
        cfg["grid"]["geometry_device"] = geometry_device
    o, m = cfg["objective"], cfg["manufacturing"]
    crit = dict(g_min=o["min_gain"], ln_1_g_min=math.log(1 + o["min_gain"]), f_K=o["floor_stiffness"],
                f_P=o["floor_capacity"])
    files = {}
    ekw = dict(export_kw or {})

    # 1. starts ----------------------------------------------------------------------------------------------------
    t0 = time.time()
    rows = start_table(run_dir, cfg)
    best = pick_best(rows)
    if best is None and allow_incomplete:
        cand = [r for r in rows if r["status"] != "missing" and np.isfinite(r["J_opt"])]
        if cand:
            best = max(cand, key=lambda r: r["J_opt"])
            log(f"[collect] WARNING no complete feasible start; --allow-incomplete uses the latest checkpoint of "
                f"{best['start']} (status {best['status']})")
    tim["starts"] = time.time() - t0
    log("[collect] starts: " + "; ".join(f"{r['start']} {r['status']} J {r['J_opt']:+.4f} feasible {r['feasible']}"
                                         for r in rows))
    t0 = time.time()
    ctx = Context(cfg, log=log)
    tim["context"] = time.time() - t0
    memo("context")
    nv = ctx.surf.n_vertices
    xi = None
    if best is not None:
        if best["xi_file"]:
            xi = np.load(os.path.join(run_dir, best["xi_file"]))
        else:
            xi = np.asarray(latest_checkpoint(run_dir, best["start"])["xi"])
        if len(xi) != 2 * nv:
            raise ValueError(f"xi of {best['start']} has {len(xi)} entries, Γ_h has {nv} vertices (2N expected)")
    D_comp = best["D_comp"] if best else cfg["design"]["outward_extent_initial_mm"]
    h_g_last = float(cfg["grid"]["phases"][-1]["h_g"])
    for r in rows:                                        # wall range per start (vertex fields, final-phase map)
        r["w_min"] = r["w_max"] = r["closed_area_frac"] = None
        p = os.path.join(run_dir, r["xi_file"]) if r["xi_file"] else None
        if p and os.path.exists(p):
            f = cex.vertex_fields(ctx, np.load(p), r["D_comp"], h_g_last)
            act = ~f["passive"]
            op = act & (f["a"] < ctx.r_ma)
            w = (f["a"] + f["d"])[op]
            r["w_min"], r["w_max"] = (float(w.min()), float(w.max())) if w.size else (None, None)
            A = ctx.surf.A * act
            r["closed_area_frac"] = float((A * (f["a"] >= ctx.r_ma)).sum() / A.sum())
    b_v, b_o = m["wall_bias_mm"], m["outer_bias_mm"]
    vf = cex.vertex_fields(ctx, xi, D_comp, h_g_last) if xi is not None else None
    f_best = None
    if vf is not None:
        f_best = design_fields(vf["d"], vf["a"], vf["xt_a"] >= cex.ExportConfig.closed_thr, ctx.r_ma, b_v)

    # 2. parametric FE of the best design on the optimisation grid; uniform-shell t on its band ----------------------
    res_list = resolutions(cfg)
    t0 = time.time()
    model = make_model(ctx, res_list[0], D_comp, log)
    memo("model opt (parametric)")
    st = scoring_settings(model)
    V0_band = float(_h(model.V0))
    rule, rho_p = {}, None
    if xi is not None:                                    # geometry first (band on the device), then the FE
        dm = model.dm
        x = dm.xp.asarray(xi)
        V_opt = dm.volume(dm.forward(x[:nv], x[nv:], as_built=False))
        _, V_rule = cell_rescore(model, None, f_best, 0.0, want_rho=False)
        rule = dict(V_opt_rel=V_opt / V0_band - 1.0, V_rule_rel=V_rule / V0_band - 1.0, rule_vs_opt=V_rule / V_opt - 1.0)
        log(f"[collect] best ({best['start']}) band volume: optimiser V/V0-1 {rule['V_opt_rel']:+.3e}, export closure "
            f"rule {rule['V_rule_rel']:+.3e} (rule/optimiser - 1 = {rule['rule_vs_opt']:+.3e})")
        rho_p = _h(dm.rho(dm.forward(x[:nv], x[nv:], as_built=True))).astype(np.float64)
    # uniform shell (§7.4 S1, D1: mount faces solid): t with V_nominal(t) = V0 on this band
    has_mount = bool(np.asarray(ctx.mount).any())

    def shell_dV(t):
        fu = cex.uniform_shell_fields(ctx, t, h_g_last)
        return cell_rescore(model, None, design_fields(fu["d"], fu["a"], fu["closed"], ctx.r_ma, b_v), 0.0,
                            want_rho=False)[1] / V0_band - 1.0
    from scipy.optimize import brentq
    t_band = None
    try:
        t_band = brentq(shell_dV, 0.02, D_comp, xtol=1e-5)
        log(f"[collect] uniform shell{' (mounts solid)' if has_mount else ''}: t = {t_band:.4f} mm for constant mass "
            "on the optimisation band")
    except ValueError as e:
        log(f"[collect] uniform shell: no t in [0.02, {D_comp}] mm restores the mass on the band ({e})")
        if has_mount:
            raise RuntimeError(f"uniform shell: no t ≤ D_comp = {D_comp} mm restores the solid's mass") from e
    if model.geometry_device == "gpu":
        model.disc.release()                             # the FE gets the device memory (as fe_eval does)
    raw_solid0 = raw_responses(model, model.rho0, st, keep_fields=True)
    raw_param = None
    if rho_p is not None:
        raw_param = raw_responses(model, rho_p, st, keep_fields=True)
        rho_p = None
        log(f"[collect] parametric best ({best['start']}) on the optimisation grid: "
            f"{score(raw_param, raw_solid0, st)}")
    g_opt = _grid_info(model)
    tim["param_fe"] = time.time() - t0
    free_model(model)
    model = None
    memo("release model opt (parametric)")

    # 3. export (the PhaseModel is released: the forked export workers do not inherit it) ----------------------------
    t0 = time.time()
    bdir = os.path.join(run_dir, "export_baselines")
    base_ex = cex.export_baselines(ctx, bdir, t_coarse=t_band, h_g=h_g_last, log=log, **ekw)
    for nm in ("original", "uniform_shell"):
        shutil.copyfile(os.path.join(bdir, f"{nm}.stl"), os.path.join(run_dir, f"{nm}.stl"))
        files[f"{nm}.stl"] = f"{nm}.stl"
    t_shell = float(base_ex["uniform_shell"]["t_final_mm"])
    if t_shell > D_comp:
        raise RuntimeError(f"uniform shell t = {t_shell:.2f} mm exceeds the band's outward extent {D_comp} mm")
    tim["export_baselines"] = time.time() - t0
    trim_host_memory()
    memo("export baselines")
    best_ex, export_ok, export_why = None, True, ""
    if xi is not None:
        t0 = time.time()
        stress_fn = DenseLookup(g_opt, raw_param["vm_dense"])
        best_ex = cex.export_design(ctx, xi, run_dir, D_comp, stress=stress_fn, name="best_candidate", h_g=h_g_last,
                                    log=log, **ekw)
        files["best_candidate.stl"] = "best_candidate.stl"
        tim["export_best"] = time.time() - t0
        trim_host_memory()
        memo("export best")
        hist_b = [h_ for h_ in read_history(run_dir) if h_["start"] == best["start"]]
        G_s = _f(hist_b[-1].get("G_s")) if hist_b else None
        export_ok, export_why = export_verdict(best_ex, G_s=G_s)
        if not export_ok:
            log(f"[collect] best_candidate export FAILED its checks: {export_why}. The STL is kept for inspection and "
                "not re-scored (D2); the original solid is kept.")

    # 4. re-score ----------------------------------------------------------------------------------------------------
    stl = {"uniform_shell": os.path.join(run_dir, "uniform_shell.stl")}
    fields = {"uniform_shell": None}
    fu = cex.uniform_shell_fields(ctx, t_shell, h_g_last)
    fields["uniform_shell"] = design_fields(fu["d"], fu["a"], fu["closed"], ctx.r_ma, b_v)
    if best_ex is not None and export_ok:
        stl["best"] = os.path.join(run_dir, "best_candidate.stl")
        fields["best"] = f_best
    c0 = np.round(0.5 * (ctx.V.min(0) + ctx.V.max(0)), 3)            # band.c0 (Warp frame)
    res_rows, raws = [], {}
    fields_extra, modes = {}, {}
    for r in res_list:
        t0 = time.time()
        model = make_model(ctx, r, D_comp, log)
        if not np.allclose(model.band.c0, c0):
            raise RuntimeError("band centring offset differs from the mesh frame used for the STL soup")
        memo(f"model {r['res']}")
        st_r = scoring_settings(model)
        bcs = DesignBCs(model, ctx)
        V0 = float(_h(model.V0))
        rho_chk, _ = cell_rescore(model, None, None, b_o, solid=True)
        if not np.allclose(rho_chk, _h(model.rho0), rtol=0, atol=1e-12):
            raise RuntimeError("re-scoring ρ of the solid differs from the optimiser's baseline ρ0")
        del rho_chk
        rhos = {"solid": _h(model.rho0).astype(np.float64)}
        vols = {"solid": V0}
        for name, path in stl.items():
            soup = stl_soup(path, c0)
            phiE = mesh_phi(model, soup)
            del soup
            rhos[name], vols[name] = cell_rescore(model, phiE, fields[name], b_o)
            del phiE
            trim_host_memory()
        if model.geometry_device == "gpu":
            model.disc.release()                         # the FE gets the device memory (as fe_eval does)
        designs, bcinfo = {}, {}
        for name, rho in rhos.items():
            rho2, fixed, load, info = bcs.adjust(rho)
            bcinfo[name] = info
            if info["islands"] or info["fixed_dropped"] or any(info["load_nodes_dropped"]):
                log(f"[collect] {r['res']}: {name}: BCs on material: {info}")
            bcs.apply(fixed, load)
            keep = r["res"] == "opt" and name in ("solid", "best")
            try:
                designs[name] = raw_responses(model, rho2, st_r, keep_fields=keep)
            except RuntimeError as e:                        # e.g. PCG not converged: recorded, the row fails
                log(f"[collect] {r['res']}: {name}: FE FAILED: {e!r}")
                designs[name] = dict(C=float("nan"), sigma_pn=float("nan"), sigma_q=float("nan"),
                                     mu_tilde=float("nan"), lam1=float("nan"), dof=0, seconds=0.0, error=repr(e)[:300])
            fe = model.fe
            designs[name]["pcg"] = list(getattr(fe, "stats_solves", []) or [])[-3:] if fe is not None else []
            del rho2
        bcs.apply(bcs.fixed0, bcs.load0)
        raw0 = designs["solid"]
        if r["res"] == "opt":
            log(f"[collect] solid re-score vs the parametric pass on the same grid: C {raw0['C'] / raw_solid0['C'] - 1:+.2e}")
        for name, raw in designs.items():
            sc = score(raw, raw0, st_r) if np.isfinite(raw["C"]) else dict(
                J=float("nan"), K_K0=float("nan"), P_P0=float("nan"), lam_s=float("nan"), lam_b=float("nan"),
                lam=float("nan"), governing=None)
            bi = bcinfo[name]
            row = dict(label=f"{r['res']}: {name}", res=r["res"], design=name, h_e=r["h_e"], h_g=r["h_g"],
                       shift_mm=r["shift"], **sc, C_Nmm=raw["C"], K_N_mm=float(ctx.F @ ctx.F / raw["C"]),
                       lam1=raw["lam1"], sigma_pn=raw["sigma_pn"], sigma_q=raw["sigma_q"], V_mm3=vols[name],
                       mass_err=vols[name] / V0 - 1.0, dof=raw["dof"], seconds=raw["seconds"],
                       error=raw.get("error"), pcg=raw.get("pcg"),
                       bc_load_share_redistributed=bi["load_share_redistributed"],
                       bc_load_nodes_dropped=bi["load_nodes_dropped"], bc_fixed_dropped=bi["fixed_dropped"],
                       bc_isolated_load_nodes=bi["isolated_load_nodes"], islands=bi["islands"],
                       island_mm3=bi["island_mm3"])
            res_rows.append(row)
            log(f"[collect] {row['label']:<24} J {sc['J']:+.4f} K/K0 {sc['K_K0']:.4f} P/P0 {sc['P_P0']:.4f} "
                f"{sc['governing']} V/V0-1 {row['mass_err']:+.2e} load moved {bi['load_share_redistributed']:.1%}")
        raws[r["res"]] = (designs, st_r)
        if r["res"] == "opt":
            surfX = ctx.surf.V
            Ns = ctx.surf.projected_normals() if hasattr(ctx.surf, "projected_normals") else ctx.surf.N
            dd = vf["d"] if vf is not None else np.zeros(nv)
            pts_d = surfX + (dd - 0.5 * model.h_e)[:, None] * Ns
            pts_s = surfX - 0.5 * model.h_e * Ns
            fields_extra = dict(stress_solid=DenseLookup(model, raw0["vm_dense"])(pts_s))
            # design fields: the re-scored export when it passed its checks, else the parametric best (same grid)
            rb, gb = (designs["best"], model) if "best" in designs and "vm_dense" in designs["best"] else (raw_param, g_opt)
            if rb is not None:
                fields_extra["stress"] = DenseLookup(gb, rb["vm_dense"])(pts_d)
                fields_extra["W_bar"] = DenseLookup(gb, rb["W_dense"])(pts_d)
            modes = dict(solid_mag=mode_magnitudes(model, raw0, pts_s), solid_lambda=1.0 / raw0["mu"][raw0["mu"] > 0]
                         if raw0.get("mu") is not None and raw0["mu"].size else np.zeros(0))
            if rb is not None and rb.get("mu") is not None:
                modes.update(design_mag=mode_magnitudes(gb, rb, pts_d),
                             design_lambda=1.0 / rb["mu"][rb["mu"] > 0] if rb["mu"].size else np.zeros(0))
            for dsg in designs.values():                  # drop big arrays
                for k in ("vm_dense", "W_dense", "modes", "node_dense"):
                    dsg.pop(k, None)
        tim[f"rescore_{r['res']}"] = time.time() - t0
        rb = gb = None
        free_model(model)
        model = bcs = None
        trim_host_memory()
        memo(f"release model {r['res']}")
    raw_param = raw_solid0 = None

    # 5. body-fitted validation ---------------------------------------------------------------------------------
    from . import validate as cval
    do_val = cfg["validation"].get("body_fitted", True) if validate is None else validate
    val = dict(ran=False, why="disabled" if not do_val else ("export failed" if not export_ok else "no design"))
    blocking = val_failed = False
    if do_val and "best" in stl:
        try:
            t0 = time.time()
            vd = os.path.join(run_dir, "validation")
            h_val = _get(cfg, "validation.h_mm", 0.5)       # coarse level; the finer is h/sqrt(2) (>= 2 tets per wall)
            kw = dict(h=h_val, h_bc=res_list[0]["h_e"], n_modes=4, nproc=int(_get(cfg, "validation.nproc", min(
                32, len(os.sched_getaffinity(0))))), want_stress=True, E=cfg["material"]["E_MPa"], nu=cfg["material"]["nu"],
                fallback=None, timeout=_get(cfg, "validation.timeout_s", None))
            bf, dec, errs, used = {}, {}, {}, {}
            VE, TE, _ = load_mesh(stl["best"])
            # Same mesher for solid and design (ratios cancel mesher bias): isotropic surface remesh to h, gmsh volume
            # inside that triangulation, quality gate; solved on the GPU (cuDSS, C3D10-identical). Two levels
            # (h, h/sqrt(2)) give the mesh-convergence estimate; the finer level is compared. No tetgen fallback:
            # its sliver mesh moved lambda_b by -46 % on B2 (research/b2_bodyfit).
            kw.update(engine=_get(cfg, "validation.engine", "gpu"))
            for nm, (Vs, Ts) in (("solid", (ctx.V, ctx.T)), ("design", (VE, TE))):
                levels = [h_val, h_val / 2 ** 0.5]
                dec[nm] = dict(remesh="isotropic", levels=levels)
                try:
                    lv = []
                    for i_l, hh in enumerate(levels):
                        # the coarse level is reported (gate numbers included) but only the finer one must pass the gate
                        lv.append(cval.validate_design(Vs, Ts, ctx.V, ctx.T, ctx.TF, ctx.support, ctx.loads, vd,
                                                       f"{nm}_h{hh:.3f}", backend="keep", enforce_gate=(i_l == len(levels) - 1),
                                                       **dict(kw, h=hh)))
                    bf[nm] = lv[-1]
                    dec[nm]["level_change"] = dict(C=lv[1].compliance / lv[0].compliance - 1,
                                                   lam=lv[1].lambda_b / lv[0].lambda_b - 1)
                    dec[nm]["converged_C"] = bool(abs(dec[nm]["level_change"]["C"]) <= 0.02)
                    dec[nm]["gate"] = [x.info.get("mesh_quality") for x in lv]
                    dec[nm]["lambda_unexcl"] = [x.info.get("lam_unexcl") for x in lv]
                    dec[nm]["seconds"] = [dict(mesh=x.seconds_mesh, solve=x.seconds_ccx) for x in lv]
                    used[nm] = "keep"
                    if not dec[nm]["converged_C"]:
                        errs[f"{nm}/convergence"] = f"compliance changes {dec[nm]['level_change']['C']:+.2%} between levels"
                        log(f"[collect] WARNING {nm}: {errs[f'{nm}/convergence']} (> 2 %)")
                except Exception as e:                        # recorded; a failed check blocks acceptance below
                    errs[f"{nm}/keep"] = repr(e)[:600]
                    log(f"[collect] body-fitted {nm} FAILED: {e!r}")
            del VE, TE
            if len(bf) < 2:
                raise _ValidationFailed(errs)
            d_opt = raws["opt"][0]
            vox = {k: dict(compliance=d_opt[n]["C"], lambda_b=d_opt[n]["lam1"], sigma_q=d_opt[n]["sigma_q"])
                   for k, n in (("solid", "solid"), ("design", "best"))}
            bfd = {k: dict(compliance=v.compliance, lambda_b=v.lambda_b, sigma_q=v.sigma_q) for k, v in bf.items()}
            vrows = cval.compare(vox, bfd)
            cval.write_validation_csv(vrows, os.path.join(run_dir, "validation.csv"))
            files["validation.csv"] = "validation.csv"
            blocking = any(r["blocking"] and not r["passed"] for r in vrows)
            for r in vrows:
                if r.get("under"):
                    log(f"[collect] body-fitted: voxel under-predicts {r['metric']} by {-r['over']:.1%} "
                        "(conservative; logged, not blocking, §8.6 D3)")
            val = dict(ran=True, engine=sorted({v.engine for v in bf.values()}), blocking_miss=blocking, surface=dec,
                       errors=errs, mesher=used,
                       n_rows=len(vrows), n_failed=sum(not r["passed"] for r in vrows), file="validation.csv",
                       under_predicted=[r["metric"] for r in vrows if r.get("under")],
                       seconds_mesh={k: v.seconds_mesh for k, v in bf.items()},
                       seconds_ccx={k: v.seconds_ccx for k, v in bf.items()})
            tim["validation"] = time.time() - t0
            log(f"[collect] body-fitted: {val}")
        except _ValidationFailed as e:
            val_failed = True
            val = dict(ran=False, why="body-fitted check failed (mesher/solver)", errors=e.args[0], blocking_miss=True)
            tim["validation"] = time.time() - t0
            log(f"[collect] body-fitted check FAILED: {e.args[0]}")

    # 6. decision ----------------------------------------------------------------------------------------------------
    checks = [dict(res=r["res"], J=r["J"], K_K0=r["K_K0"], P_P0=r["P_P0"]) for r in res_rows if r["design"] == "best"]
    extra = ["body-fitted check could not be run (see validation.errors)"] if val_failed else []
    extra += [f"{r['res']}: FE of the best export failed ({r['error']})" for r in res_rows
              if r["design"] == "best" and r.get("error")]
    acc, reason, checks = decide(checks, crit, export_ok, blocking, extra, export_why=export_why)
    final = os.path.join(run_dir, "final_design.stl")
    shutil.copyfile(os.path.join(run_dir, "best_candidate.stl" if acc else "original.stl"), final)
    files["final_design.stl"] = "final_design.stl"
    log(f"[collect] DECISION: {reason}")

    # sensitivity ------------------------------------------------------------------------------------------------
    sens = []
    for kb in KAPPAS:
        for sp in SIGMA_PCT:
            js, ks, ps, jopt = [], [], [], float("nan")
            for rk, (designs, st_r) in raws.items():
                if "best" not in designs or not np.isfinite(designs["best"]["C"]):
                    continue
                sc = score(designs["best"], designs["solid"], st_r, st_r.sigma_f * (1 + sp / 100), kb)
                js.append(sc["J"]); ks.append(sc["K_K0"]); ps.append(sc["P_P0"])
                if rk == "opt":
                    jopt = sc["J"]
            ok = bool(js) and decide([dict(res="s", J=a, K_K0=b, P_P0=c) for a, b, c in zip(js, ks, ps)], crit,
                                     export_ok, blocking, extra)[0]
            sens.append(dict(kappa_b=kb, sigma_f=sp, sigma_f_MPa=cfg["material"]["strength_MPa"] * (1 + sp / 100),
                             J_star=jopt, J_min=min(js) if js else float("nan"), K_K0_min=min(ks) if ks else float("nan"),
                             P_P0_min=min(ps) if ps else float("nan"), accepted=int(ok)))
    _write_csv(os.path.join(run_dir, "sensitivity.csv"), sens)
    files["sensitivity.csv"] = "sensitivity.csv"

    # fields.vtp, buckling_modes.npz, bc_faces.npz -----------------------------------------------------------------
    s = ctx.surf
    pdata = {}
    if vf is not None:
        op = vf["a"] < ctx.r_ma
        pdata = dict(xi=xi[nv:], xi_d=xi[:nv], xi_a=xi[nv:], xit_a=vf["xt_a"], d=vf["d"], a=vf["a"],
                     w=np.where(op, vf["a"] + vf["d"], np.nan), r_MA=ctx.r_ma, d_hi=vf["d_hi"],
                     passive=vf["passive"].astype(bool))
    else:
        pdata = dict(r_MA=ctx.r_ma)
    pdata.update(fields_extra)
    write_vtp(os.path.join(run_dir, "fields.vtp"), s.V, s.T, pdata)
    files["fields.vtp"] = "fields.vtp"
    np.savez_compressed(os.path.join(run_dir, "buckling_modes.npz"), V=s.V.astype(np.float32), T=s.T.astype(np.int32),
                        **{k: np.asarray(v, np.float32) for k, v in modes.items()})
    files["buckling_modes.npz"] = "buckling_modes.npz"
    load_id = np.full(len(ctx.T), -1, np.int32)
    for i, L in enumerate(ctx.loads):
        load_id[np.isin(ctx.TF, L["faces"])] = i
    np.savez_compressed(os.path.join(run_dir, "bc_faces.npz"), V=ctx.V.astype(np.float32), T=ctx.T.astype(np.int32),
                        support=np.isin(ctx.TF, ctx.support), load=load_id)
    files["bc_faces.npz"] = "bc_faces.npz"

    # summary ------------------------------------------------------------------------------------------------------
    rho_g = cfg["material"]["density_g_cm3"] * 1e-3
    bmesh = (best_ex or {}).get("mesh", {})
    V0x = (best_ex or base_ex["original"]).get("V0_mm3")
    vfin = bmesh.get("volume") if acc else base_ex["original"].get("mesh", {}).get("volume")
    mass = dict(target=V0x * rho_g if V0x else None, final=vfin * rho_g if vfin else None,
                error=bmesh.get("volume_rel"), V0_mm3=V0x, density_g_mm3=rho_g,
                mass_matched=bool(export_ok) if best_ex is not None else None)
    d_opt = {r["design"]: r for r in res_rows if r["res"] == "opt"}
    comparison = [dict(label="solid", source="rescored", J=0.0, K_K0=1.0, P_P0=1.0,
                       governing=d_opt["solid"]["governing"], w_min=None, w_max=None),
                  dict(label="uniform shell", source="rescored", **{k: d_opt["uniform_shell"][k] for k in
                                                                    ("J", "K_K0", "P_P0", "governing")},
                       w_min=t_shell, w_max=t_shell)]
    vrow = None
    try:
        vs = load_vshell(vshell_run_dir(cfg))
        b = vs["best"]
        vrow = dict(run_dir=str(vshell_run_dir(cfg)), J=b.get("J"), K_K0=b.get("K/K0"), P_P0=b.get("P/P0"),
                    accepted=vs["accepted"], s_mm=vs["s"])
        comparison.append(dict(label="vshell best", source="vshell", J=b.get("J"), K_K0=b.get("K/K0"),
                               P_P0=b.get("P/P0"), governing=b.get("governing mode"),
                               w_min=b.get("as-built wall mm"), w_max=b.get("outer offset / strut wall s mm")))
    except (FileNotFoundError, KeyError) as e:
        log(f"[collect] vshell row not available: {e}")
    for r in rows:
        comparison.append(dict(label=f"cshell {r['start']}", source="optimiser", J=r["J_opt"], K_K0=r["K_K0_opt"],
                               P_P0=r["P_P0_opt"], governing=None, w_min=r["w_min"], w_max=r["w_max"]))
    if "best" in d_opt:
        bw = (best_ex or {}).get("wall", {})
        comparison.append(dict(label="cshell export", source="rescored", **{k: d_opt["best"][k] for k in
                                                                            ("J", "K_K0", "P_P0", "governing")},
                               w_min=bw.get("min_nominal"), w_max=next((r["w_max"] for r in rows if best and
                                                                        r["start"] == best["start"]), None)))
    elif best_ex is not None:
        comparison.append(dict(label="cshell export", source="not mass-matched", J=None, K_K0=None, P_P0=None,
                               governing=None, w_min=None, w_max=None))
    load_cases = [dict(case=f"1 ({n})", lam_s=d_opt[k]["lam_s"], lam_b=d_opt[k]["lam_b"], softmin=d_opt[k]["lam"])
                  for k, n in (("solid", "solid"), ("best", "cshell")) if k in d_opt]
    Lam = None
    if vf is not None and "W_bar" in pdata:
        sel = ~vf["passive"] & (vf["a"] < ctx.r_ma) & np.isfinite(pdata["W_bar"]) & (pdata["W_bar"] > 0)
        if sel.any():
            w_, a_ = pdata["W_bar"][sel], ctx.surf.A[sel]
            o_ = np.argsort(w_)
            c_ = np.cumsum(a_[o_])
            Lam = float(w_[o_][np.searchsorted(c_, 0.5 * c_[-1])])
    hist = read_history(run_dir)
    peak = max([_f(h_["gpu_mem_GiB"]) for h_ in hist if np.isfinite(_f(h_["gpu_mem_GiB"]))] or [float("nan")])
    gpu_model = None
    try:
        import cupy as cp
        gpu_model = cp.cuda.runtime.getDeviceProperties(0)["name"].decode()
    except Exception:
        pass
    tim["total"] = time.time() - T0
    import resource
    mem["peak_rss_GiB"] = resource.getrusage(resource.RUSAGE_SELF).ru_maxrss / 2 ** 20
    mem["peak_rss_children_GiB"] = resource.getrusage(resource.RUSAGE_CHILDREN).ru_maxrss / 2 ** 20
    peak_collect_gpu = pk_gpu.stop() if pk_gpu is not None else float("nan")
    ex_sum = None
    if best_ex is not None:
        ex_sum = {k: v for k, v in best_ex.items() if k not in ("timing_s",)}
    S = dict(schema=SCHEMA, name=cfg["name"], accepted=acc,
             outcome="cshell design accepted" if acc else "original SOLID part kept", reason=reason, criteria=crit,
             mass=mass,
             best=None if best is None else dict(start=best["start"], J_opt=best["J_opt"], D_comp=D_comp,
                                                 xi_file=best["xi_file"], phase=len(cfg["grid"]["phases"]) - 1),
             starts=rows, comparison=comparison, resolutions=res_rows, checks=checks, load_cases=load_cases,
             Lambda=Lam,
             export=dict(best=ex_sum, best_ok=export_ok if best_ex is not None else None,
                         best_reject_reason=export_why or None, closure_rule_band=rule or None,
                         uniform_shell_t_mm=t_shell, uniform_shell_t_band_mm=t_band,
                         uniform_shell_mounts_solid=bool(base_ex["uniform_shell"].get("mounts_solid", False)),
                         delta_mm=(best_ex or {}).get("mass_correction", {}).get("delta"),
                         voids=(best_ex or {}).get("voids"), holes=((best_ex or {}).get("voids") or {}).get("holes"),
                         wall=(best_ex or {}).get("wall")),
             validation=val,
             sensitivity=dict(file="sensitivity.csv", accepted_fraction=float(np.mean([r["accepted"] for r in sens])),
                              kappa_b=list(KAPPAS), sigma_f_pct=list(SIGMA_PCT)),
             vshell=vrow,
             resources=dict(gpu_model=gpu_model, peak_gpu_GiB=peak, peak_gpu_collect_GiB=peak_collect_gpu,
                            host_rss_GiB=mem,
                            gpu_hours=float(np.nansum([r["hours"] for r in rows])) + tim["total"] / 3600,
                            collect_s=tim["total"]),
             timings_s=tim, files=files)
    S["files"]["summary.json"] = "summary.json"
    S["files"]["summary.txt"] = "summary.txt"
    S = _jsonable(S)
    json.dump(S, open(os.path.join(run_dir, "summary.json"), "w"), indent=1)
    with open(os.path.join(run_dir, "summary.txt"), "w") as f:
        f.write(summary_text(S))
    log(f"[collect] done in {tim['total']:.0f} s (peak host RSS {mem['peak_rss_GiB']:.1f} GiB, export workers "
        f"{mem['peak_rss_children_GiB']:.1f} GiB, peak GPU {peak_collect_gpu:.1f} GiB)")
    return S


