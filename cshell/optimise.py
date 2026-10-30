"""Phases, starts, the MMA loop, mass restoration, checkpoints and history (THEORY §7, §5.4).

One `Run` per (config, start). Each continuation phase builds a `PhaseModel` (band, P, discretisation, design space,
FE boundary conditions, solid baseline) on its own grids; ξ lives on Γ_h and carries over unchanged (§7.3).
The FE engine is the active-set cuDSS engine (fe_sparse): the active set is every element with ρ_e > 0 dilated by
`grid.active_margin_mm`; it is rebuilt (new sparsity pattern) whenever the design puts material outside it, so the
gradient stays exact (elements outside carry ρ = 0 and H′ = 0).
"""
from __future__ import annotations

import csv
import glob
import json
import os
import socket
import time
from dataclasses import dataclass

import numpy as np
from scipy import ndimage

from . import config as ccfg
from . import fea
from .band import Band, P_from_compact
from .design import DesignMap, DesignSpace, DeviceDesignMap, DeviceDisc, Discretisation, passive_thin
from .fe_sparse import ActiveGrid, SparseFE
from .mma import MMA, ProjectedLBFGSB, restore_mass
from .objective import Settings, evaluate
from .surface import DesignSurface, medial_depth, mount_face_ids, mount_vertices, outward_clearance

HIST_COLS = ["start", "phase", "iter", "J", "C_C0", "lam_s_ratio", "lam_b_ratio", "lam_ratio", "governing",
             "mass_err", "G_s", "max_dxi", "max_dxi_raw", "move", "n_restore", "eig_iters", "eig_converged", "active_dof",
             "t_geometry", "t_fe", "t_mma", "t_restore", "gpu_mem_GiB", "c_norm", "sliver_w", "stress_p", "beta",
             "t_pull", "t_mmasub", "t_upload", "gpu_peak_GiB", "fe_engine"]    # t_mma = t_pull + t_mmasub (+ bookkeeping)


# ---------------------------------------------------------------------------------------------------------------------
class Context:
    """Phase-independent data: geometry, faces, BCs, design surface and its per-vertex geometry (THEORY §3.4-3.5)."""

    def __init__(self, cfg: dict, log=print):
        from vshell import geometry
        self.cfg, self.log = cfg, log
        t0 = time.time()
        shape = geometry.load_step(cfg["step_file"])
        self.faces = geometry.face_table(shape)
        self.support = sorted({i for s in cfg["supports"] for i in geometry.select_faces(self.faces, s)})
        self.loads = [dict(faces=geometry.select_faces(self.faces, L["faces"]), force=np.asarray(L["force_N"], float))
                      for L in cfg["loads"]]
        if not self.support or any(not L["faces"] for L in self.loads):
            raise ValueError("support or load selector matched no faces")
        self.F = np.sum([L["force"] for L in self.loads], 0)
        extra = sorted({i for s in cfg["mounts"]["extra_faces"] for i in geometry.select_faces(self.faces, s)})
        forced = sorted(set(self.support) | {i for L in self.loads for i in L["faces"]} | set(extra))
        analytic = [f["id"] for f in self.faces if f["type"] in geometry.ANALYTIC]
        self.V, self.T, self.TF = geometry.tessellate(shape, cfg["grid"]["tessellation_mm"])
        d = cfg["design"]
        self.surf = DesignSurface.build(self.V, self.T, self.TF, ell=d["surface_edge_mm"])
        s = self.surf
        self.mount_ids = mount_face_ids(analytic, forced, s.V0, s.T0, s.TF0, cfg["mounts"]["auto_analytic"],
                                        cfg["mounts"]["min_blob_fraction"])
        self.mount = mount_vertices(s.V, s.V0, s.T0, s.TF0, self.mount_ids, cfg["mounts"]["dilate_mm"])
        h_fine = min(p["h_g"] for p in cfg["grid"]["phases"])
        N = s.projected_normals() if hasattr(s, "projected_normals") else s.N
        self.r_ma = medial_depth(s.V, N, s.query0, h_g=h_fine)
        self.r_out = outward_clearance(s.V, N, s.query0)
        self.keepout = np.full(s.n_vertices, np.inf)          # no keep-outs needed for Study 1 (D1 of §14)
        if cfg["design"]["keep_out"]:
            raise NotImplementedError("keep_out bodies: surface.keepout_distance wiring is a Stage 1 TODO")
        self.info = dict(n_vertices=s.n_vertices, hausdorff=s.info.get("hausdorff"), mount_faces=self.mount_ids,
                         setup_s=time.time() - t0)
        log(f"[cshell] context: {s.n_vertices} vertices, mounts {len(self.mount_ids)} faces, "
            f"r_MA {np.median(self.r_ma):.2f} mm median, {time.time() - t0:.0f} s")


# ---------------------------------------------------------------------------------------------------------------------
class PhaseModel:
    """Grids, P, design space, BCs and solid baseline of one continuation phase (THEORY §7.3)."""

    def __init__(self, ctx: Context, phase: int, D_comp: float, log=print):
        cfg = ctx.cfg
        self.ctx, self.phase, self.D_comp, self.log = ctx, phase, float(D_comp), log
        ph = cfg["grid"]["phases"][phase]
        self.h_e, self.h_g = float(ph["h_e"]), float(ph["h_g"])
        self.k = int(round(self.h_e / self.h_g))
        t0 = time.time()
        band = Band.build(ctx.V, ctx.T, self.h_g, D_comp)
        self.band = band
        face, bary = band.P_compact(ctx.surf.V, ctx.surf.T)
        ijk = band.ijk()
        eshape = tuple(int(np.ceil(n / self.k)) for n in band.grid.shape)
        eijk = ijk // self.k
        del ijk
        elem = (eijk[:, 0] * eshape[1] + eijk[:, 1]) * eshape[2] + eijk[:, 2]
        del eijk
        self.elem_shape, self.node_shape = eshape, tuple(n + 1 for n in eshape)
        self.origin = band.grid.origin                         # FE node (0,0,0) = band-grid corner
        rho_min = cfg["grid"]["rho_min"]
        n_elem, V_extra = int(np.prod(eshape)), band.n_deep * self.h_g ** 3
        self.geometry_device = geometry_device(cfg["grid"].get("geometry_device", "auto"))
        self.geometry_offload = bool(cfg["grid"].get("geometry_offload", True))
        if self.geometry_device == "gpu":                      # compact band on the device (24 B/cell)
            disc = DeviceDisc(band.phi, face, bary, ctx.surf.T, elem.astype(np.int32), ctx.surf.n_vertices, n_elem,
                              self.k, self.h_g, rho_min, V_extra, log=log)
        else:
            P = P_from_compact(face, bary, np.ascontiguousarray(ctx.surf.T, np.int32), ctx.surf.n_vertices)
            disc = Discretisation(phi=band.phi.astype(np.float64), P=P, elem=elem, n_elem=n_elem, k=self.k,
                                  h_g=self.h_g, rho_min=rho_min, V_extra=V_extra)
        del face, bary, elem
        m = cfg["manufacturing"]
        d = cfg["design"]
        a_hi = ctx.r_ma + d["closure_gamma"] * self.h_g
        d_hi = np.minimum.reduce([ctx.r_out - d["gap_mm"], ctx.keepout - d["keep_out_clearance_mm"],
                                  np.full(ctx.surf.n_vertices, D_comp)]).clip(0.0)
        if not d["allow_outward"]:
            d_hi[:] = 0.0
        self.d_hi_by_comp = d_hi >= D_comp - 1e-9
        thin = passive_thin(a_hi, m["t_min_mm"], m["void_min_mm"])
        passive = ctx.mount | thin
        self.passive_parts = dict(mount=int(ctx.mount.sum()), thin=int(thin.sum()))
        self.r_f = ccfg.phase_value(cfg["filter"]["radius_mm"], phase)
        ds = DesignSpace(r_ma=ctx.r_ma, a_hi=a_hi, d_hi=d_hi, area=ctx.surf.A, passive=passive,
                         filt=ctx.surf.filter(self.r_f), t_min=m["t_min_mm"], v_min=m["void_min_mm"],
                         b_v=m["wall_bias_mm"], b_o=m["outer_bias_mm"], soft_w=d["inset_soft_mm"])
        self.ds, self.disc = ds, disc
        self.dm = DeviceDesignMap(ds, disc) if self.geometry_device == "gpu" else DesignMap(ds, disc)
        # solid: ψ = φ (nominal), ψ_ab = φ + b_o (THEORY §3.8, §3.9)
        self.V0 = self.dm.solid_volume()
        self.rho0 = self.dm.solid_rho(m["outer_bias_mm"])
        # BCs on the dense node box (THEORY §4.3)
        self.fixed, self.load, self.bc_info = fea.build_bcs(self.origin, self.h_e, self.node_shape, ctx.V, ctx.T, ctx.TF,
                                                            ctx.support, ctx.loads)
        self.stress_mask_dense = self._stress_mask()
        self.ksig_mask_dense = self._ksig_mask()
        self.t_setup = time.time() - t0
        self.grid_fe = None
        self.fe = None
        self.fe_engine = None
        log(f"[cshell] phase {phase}: h_e {self.h_e} h_g {self.h_g}, band {band.n / 1e6:.1f} M cells, element box "
            f"{eshape}, passive {self.passive_parts}, V0 {self.V0:.1f} mm³, {self.t_setup:.0f} s")
        if self.geometry_device == "gpu":
            log(f"[cshell]   geometry chain on the GPU: compact band {disc.nbytes / 2**30:.2f} GiB"
                + (", released from the device during each FE evaluation" if self.geometry_offload else ""))
        else:
            log("[cshell]   geometry chain on the CPU (NumPy/SciPy)")

    def _ksig_mask(self) -> np.ndarray | None:
        """1 where an element keeps its geometric stiffness, 0 within exclude_bc_mm of a clamped node (Owen 2026-10-29:
        the clamp idealisation leaves a singular concave corner where the design meets a support, and a local mode
        there is mesh-dependent in both voxel and body-fitted models). `capacity.ksigma_exclude_bc` = false disables."""
        cap = self.ctx.cfg["capacity"]
        if not cap.get("ksigma_exclude_bc", True):
            return None
        r = int(np.ceil(cap["exclude_bc_mm"] / self.h_e))
        near = ndimage.binary_dilation(self.fixed.any(-1), iterations=r)[:-1, :-1, :-1]
        return (~near).ravel().astype(np.float64)

    def _stress_mask(self) -> np.ndarray:
        """m_e = 0 within exclude_bc_mm of a fixed or loaded node and in mounts unless include_mounts (THEORY §4.4(b))."""
        cap = self.ctx.cfg["capacity"]
        bc = np.zeros(self.node_shape, bool)
        bc |= self.fixed.any(-1)
        bc |= (self.load != 0).any(-1)
        r = int(np.ceil(cap["exclude_bc_mm"] / self.h_e))
        near = ndimage.binary_dilation(bc, iterations=r)[:-1, :-1, :-1]
        m = ~near
        if not cap["include_mounts"]:
            frac = self.dm.elem_mean(self.ctx.mount.astype(float))
            m &= frac.reshape(self.elem_shape) < 0.5
        return m.ravel()

    # -- FE active set -------------------------------------------------------------------------------------------
    def ensure_active(self, rho: np.ndarray) -> bool:
        """(Re)build the active set if material lies outside it. Returns True when rebuilt."""
        if self.grid_fe is not None:
            outside = rho.copy()
            outside[self.grid_fe.elem_dense] = 0.0
            if not np.any(outside > 0):
                return False
        margin = int(np.ceil(self.ctx.cfg["grid"]["active_margin_mm"] / self.h_e))
        mask = ndimage.binary_dilation((rho > 0).reshape(self.elem_shape), iterations=margin)
        if self.fe is not None:
            self.fe.free()
            self.fe = None
        free_device_memory()          # old factor and cached blocks go back before the new plan sizes itself
        self.grid_fe, self.fe, self.fe_engine = self._build_engine(mask, rho)
        if self.ksig_mask_dense is not None:
            self.fe.ksig_mask = self.fe.xp.asarray(self.grid_fe.gather_elems(self.ksig_mask_dense))
        st = self.fe.stats
        extra = (f", coarse {st['coarse_ndof']} DOF ({st['coarse_precision']})" if self.fe_engine == "twogrid"
                 else f", {st.get('precision', '')}{' hybrid' if st.get('hybrid') else ''}")
        self.log(f"[cshell]   active set rebuilt: {self.grid_fe.Ne} elements, {self.grid_fe.ndof} DOF, engine "
                 f"{self.fe_engine}{extra}, {_gpu_free():.1f} GiB device memory free")
        return True

    def _build_engine(self, mask: np.ndarray, rho: np.ndarray):
        """FE engine for the active set `mask` (config `fe.engine`). "auto": the direct cuDSS fp64 factor when the
        problem is small (≤ fe.auto_direct_max_dof) and cuDSS's plan-time estimate fits in device memory with margin
        (checked by planning and factorising the current design: fe_sparse switches to hybrid/fp32 otherwise, and
        then the two-grid engine is used instead); else the matrix-free two-grid PCG engine (fe_twogrid)."""
        from .fe_twogrid import LeanGrid, TwoGridFE
        c = self.ctx.cfg
        fc = c["fe"]
        E, nu = c["material"]["E_MPa"], c["material"]["nu"]
        engine = {"sparse": "direct"}.get(fc.get("engine", "auto"), fc.get("engine", "auto"))
        lean = None
        if engine == "auto":
            lean = LeanGrid(mask)
            engine = "direct" if lean.ndof <= int(fc.get("auto_direct_max_dof", 2_000_000)) else "twogrid"
            if engine == "direct":
                lean = None
                g = ActiveGrid(mask)
                fe = SparseFE(g, self.fixed, self.load, self.h_e, E, nu, backend=fc["backend"], precision="auto")
                if fe.backend != "gpu":                             # "cpu" / "cpu_direct": SciPy direct solve
                    return g, fe, "direct"
                fe.factor(self.disc.rho_min + (1 - self.disc.rho_min) * g.gather_elems(rho))
                if fe.stats.get("precision") == "fp64" and not fe.stats.get("hybrid"):
                    return g, fe, "direct"
                self.log(f"[cshell]   fe auto: fp64 cuDSS factor does not fit in device memory "
                         f"({fe.stats.get('factor_estimate_GiB', float('nan')):.1f} GiB): two-grid engine instead")
                fe.free()
                del fe, g
                free_device_memory()
                engine = "twogrid"
        if engine == "direct":
            g = ActiveGrid(mask)
            return g, SparseFE(g, self.fixed, self.load, self.h_e, E, nu, backend=fc["backend"],
                               precision=fc.get("precision", "auto")), "direct"
        tg = fc.get("twogrid", {})
        g = lean if lean is not None else LeanGrid(mask)
        fe = TwoGridFE(g, self.fixed, self.load, self.h_e, E, nu, backend=fc["backend"],
                       coarse_precision=tg.get("coarse_precision", "auto"), degree=tg.get("degree", 5),
                       smoothing_range=tg.get("smoothing_range", 30.0),
                       tol=tg.get("tol", 1e-10), krylov_tol=tg.get("krylov_tol", 1e-2),
                       maxiter=tg.get("maxiter", 100), nrhs=tg.get("nrhs", 7))
        return g, fe, "twogrid"

    def settings(self, frac: float) -> Settings:
        c = self.ctx.cfg
        cap, o = c["capacity"], c["objective"]
        return Settings(w_K=o["w_stiffness"], w_P=o["w_capacity"], sigma_f=c["material"]["strength_MPa"],
                        knockdown=cap["knockdown"], beta=ccfg.phase_value(cap["softmin_beta"], self.phase, frac),
                        stress_p=ccfg.phase_value(cap["stress_p"], self.phase, frac), stress_q=cap["stress_q"],
                        quantile=cap["stress_quantile"], buckling=cap["buckling"] and o["w_capacity"] > 0,
                        num_modes=cap["buckling_modes"], buckling_p=cap["buckling_p"],
                        min_density=cap["min_density_geometric"], eig_tol=cap["eig_tol"],
                        rho_min=c["grid"]["rho_min"])

    def fe_eval(self, rho: np.ndarray, st: Settings, base, c_norm: float, prev_modes=None):
        """FE responses for element densities ρ (dense element box). With the GPU chain and `grid.geometry_offload`
        the compact band leaves the device for the FE evaluation, so the factor sees the same free memory as with
        the CPU chain (same fp64/fp32 choice); `t_upload` records the copy back."""
        off = self.geometry_device == "gpu" and self.geometry_offload
        if off:
            self.disc.release()
        try:
            return self._fe_eval(rho, st, base, c_norm, prev_modes)
        finally:
            if off:
                self.t_upload = self.disc.upload()

    def release(self):
        """Free the FE factor and the device band (before the next PhaseModel is built)."""
        if self.fe is not None:
            self.fe.free()
            self.fe = None
        if self.geometry_device == "gpu":
            self.disc.release()
        free_device_memory()

    def _fe_eval(self, rho: np.ndarray, st: Settings, base, c_norm: float, prev_modes=None):
        if self.ensure_active(rho):
            prev_modes = None                                # new DOF numbering: warm-start modes no longer apply
        g = self.grid_fe
        r = g.gather_elems(rho)
        rho_hat = self.disc.rho_min + (1 - self.disc.rho_min) * r
        m = g.gather_elems(self.stress_mask_dense).astype(float)
        if st.w_P == 0:                                     # compliance only (Stage 1): skip stress and buckling
            self.fe.factor(rho_hat)
            pr = self.fe.primal()
            dC = self.fe.dC_drho_hat(pr)
            J = 0.0 if base is None else float(np.log(base["C0"] / pr.compliance))
            dJ = -dC / pr.compliance
            return dict(J=J, dJ=dJ, C=pr.compliance, lam_s=np.nan, mu_tilde=0.0, lam=np.nan, lam_b=np.inf,
                        governing="stiffness", eig_iters=0, eig_converged=True, quantile=np.nan, sigma_pn=np.nan,
                        modes=None)
        ev = evaluate(self.fe, rho_hat, r, m, st, base, c_norm, prev_modes)
        bk = ev.parts["buckling"]
        return dict(J=ev.J, dJ=ev.dJ, C=ev.C, lam_s=ev.lam_s, mu_tilde=ev.mu_tilde, lam=ev.lam, lam_b=ev.lam_b,
                    governing=ev.governing, eig_iters=bk.iterations if bk else 0,
                    eig_converged=bk.converged if bk else True, quantile=ev.parts["stress"].quantile,
                    sigma_pn=ev.sigma_pn, modes=bk.modes if bk else None)

    def baseline(self, st: Settings) -> dict:
        """Solid responses through the same pipeline (as-built, b_o) and the initial normalisation c (THEORY §4.4(b))."""
        b0 = self.fe_eval(self.rho0, st, None, 1.0)
        c = b0["quantile"] / b0["sigma_pn"] if np.isfinite(b0["quantile"]) else 1.0
        b = self.fe_eval(self.rho0, st, None, c)
        return dict(C0=b["C"], lam0=b["lam"], sigma_pn0=b["sigma_pn"], mu0=b["mu_tilde"], c=c, lam_s0=b["lam_s"],
                    lam_b0=b["lam_b"], governing0=b["governing"])


# ---------------------------------------------------------------------------------------------------------------------
@dataclass
class RunState:
    start: str
    phase: int
    it: int
    xi: np.ndarray
    c_norm: float
    s_J: float | None
    opt_state: dict | None
    D_comp: float
    hours_used: float


class Run:
    """One start of one part: all phases, checkpoints every iteration (THEORY §7.6)."""

    def __init__(self, cfg: dict, start: str, out_dir: str, max_hours: float | None = None, log=None):
        self.cfg, self.start, self.out = cfg, start, out_dir
        os.makedirs(os.path.join(out_dir, "checkpoints"), exist_ok=True)
        self.max_hours = max_hours
        self.t_begin = time.time()
        self._logf = open(os.path.join(out_dir, "run_log.txt"), "a")
        self.log = log or self._log

    def _log(self, msg):
        print(msg, flush=True)
        self._logf.write(msg + "\n"); self._logf.flush()

    # -- starts (THEORY §7.4) ------------------------------------------------------------------------------------
    def initial_xi(self, nv: int) -> np.ndarray:
        if self.start == "uniform":
            return np.concatenate([np.full(nv, 0.5), np.zeros(nv)])
        if self.start == "neutral":
            return np.concatenate([np.full(nv, 0.5), np.full(nv, 0.5)])
        if self.start == "solid":
            return np.concatenate([np.zeros(nv), np.ones(nv)])
        if self.start == "vshell":                       # S2: map vshell's best design (cshell.collect, §7.4)
            from .collect import vshell_start_xi
            ctx = getattr(self, "ctx", None)             # set by the CLI; otherwise build the context once
            if ctx is None:
                self.log("[cshell] vshell start: no context attached to Run, building one")
                ctx = self.ctx = Context(self.cfg, log=self.log)
            xi = vshell_start_xi(ctx, self.cfg, log=self.log)
            if len(xi) != 2 * nv:
                raise ValueError(f"vshell start has {len(xi)} entries, expected {2 * nv}")
            return xi
        raise NotImplementedError(f"start {self.start!r}")

    # -- checkpoints ---------------------------------------------------------------------------------------------
    def _ckpt_path(self, rs: RunState):
        return os.path.join(self.out, "checkpoints", f"{self.start}_{rs.phase}_{rs.it:04d}.npz")

    def save(self, rs: RunState):
        flat = {f"opt_{k}": v for k, v in (rs.opt_state or {}).items()}
        np.savez(self._ckpt_path(rs), xi=rs.xi, phase=rs.phase, it=rs.it, c_norm=rs.c_norm,
                 s_J=np.nan if rs.s_J is None else rs.s_J, D_comp=rs.D_comp, hours_used=rs.hours_used,
                 rng=np.random.default_rng(self.cfg["optimiser"]["seed"]).bit_generator.state["state"]["state"], **flat)
        files = sorted(glob.glob(os.path.join(self.out, "checkpoints", f"{self.start}_*.npz")))
        for f in files[:-3]:
            os.remove(f)

    def load_latest(self) -> RunState | None:
        files = sorted(glob.glob(os.path.join(self.out, "checkpoints", f"{self.start}_*.npz")))
        if not files:
            return None
        z = np.load(files[-1])
        opt = {k[4:]: z[k] for k in z.files if k.startswith("opt_")}
        sJ = float(z["s_J"])
        return RunState(self.start, int(z["phase"]), int(z["it"]), z["xi"], float(z["c_norm"]),
                        None if np.isnan(sJ) else sJ, opt or None, float(z["D_comp"]), float(z["hours_used"]))

    def _history(self, row: dict):
        """Append to this start's own history_<start>.csv (starts run as concurrent jobs in one run dir), then rebuild
        the merged history.csv atomically."""
        path = os.path.join(self.out, f"history_{self.start}.csv")
        new = not os.path.exists(path)
        cols = HIST_COLS
        if not new:                                      # keep an existing file's columns (runs resumed across versions)
            with open(path, newline="") as f:
                cols = next(csv.reader(f), None) or HIST_COLS
        with open(path, "a", newline="") as f:
            w = csv.DictWriter(f, fieldnames=cols, extrasaction="ignore")
            if new:
                w.writeheader()
            w.writerow(row)
        merge_history(self.out)

    # -- main loop -----------------------------------------------------------------------------------------------
    def run(self, ctx: Context, resume: bool = False, only_phase: int | None = None) -> dict:
        cfg, opt = self.cfg, self.cfg["optimiser"]
        nv = ctx.surf.n_vertices
        rs = self.load_latest() if resume else None
        if rs is None:
            rs = RunState(self.start, 0, 0, self.initial_xi(nv), 1.0, None, None,
                          cfg["design"]["outward_extent_initial_mm"], 0.0)
        else:
            self.log(f"[cshell] resuming {self.start} at phase {rs.phase} iteration {rs.it}")
        hours0 = rs.hours_used
        phases = range(len(cfg["grid"]["phases"])) if only_phase is None else [only_phase]
        model = None
        result = {}
        for ph in phases:
            if ph < rs.phase:
                continue
            pending_d = None
            while True:                                          # restarts after an outward-extent extension
                if model is not None:                            # free the previous band/factor before building anew
                    model.release()
                model = dm = disc = None
                model = PhaseModel(ctx, ph, rs.D_comp, self.log)
                if pending_d is not None:                        # §3.5 step 2: keep d unchanged on the new d^hi
                    tgt = np.clip(pending_d / np.maximum(model.ds.d_hi, 1e-12), 0.0, 1.0)
                    xd = np.clip(model.ds.filt.invert(tgt), 0.0, 1.0)
                    rs.xi[:nv] = xd
                    pending_d = None
                dm, disc = model.dm, model.disc
                act = np.concatenate([~model.ds.passive, ~model.ds.passive])
                n_act = int(act.sum())
                st0 = model.settings(0.0)
                base = model.baseline(st0)
                self.log(f"[cshell] phase {ph} solid: C0 {base['C0']:.4f} N·mm, K0 {ctx.F @ ctx.F / base['C0']:.2f} N/mm, "
                         f"λ_s0 {base['lam_s0']}, λ_b0 {base['lam_b0']}, c {base['c']:.3f}")
                c_norm = base["c"] if rs.it == 0 or ph != rs.phase else rs.c_norm

                def vol(x):
                    return dm.volume(dm.forward(x[:nv], x[nv:], as_built=False))

                def vol_grad(x):
                    gd, ga = dm.volume_grad(dm.forward(x[:nv], x[nv:], as_built=False))
                    return np.concatenate([gd, ga])

                xi = rs.xi.copy()
                # passive vertices carry their imposed values in ξ itself, so the filter blends them into active
                # neighbours consistently (THEORY §3.4); otherwise a start value bleeds through F and caps V.
                xi[:nv][model.ds.passive] = 0.0
                xi[nv:][model.ds.passive] = 1.0
                xi, eta, nres = restore_mass(xi, vol, vol_grad, model.V0, act, tol=opt["mass_tol"])
                extended = False
                Cls = MMA if opt["method"] == "mma" else ProjectedLBFGSB
                optimiser = Cls(n_act, 2, move=opt["move"], con_scale=1.0)
                if rs.opt_state is not None and ph == rs.phase and rs.it > 0:
                    optimiser.load_state_dict(rs.opt_state)
                it0 = rs.it if ph == rs.phase else 0
                max_it = ccfg.phase_value(opt["max_iter"], ph)
                s_J = rs.s_J if ph == rs.phase else None
                tau = 5e-4
                calm, J_prev, prev_modes = 0, None, None
                for it in range(it0, max_it):
                    frac = it / max(max_it - 1, 1)
                    st = model.settings(frac)
                    w_s = ccfg.phase_value(opt["sliver_weight"], ph, frac)
                    t0 = time.time()
                    sab = dm.forward(xi[:nv], xi[nv:], as_built=True)
                    rho = dm.rho(sab)
                    t_geo = time.time() - t0
                    t0 = time.time()
                    base_c = dict(base)
                    if st.w_P > 0:                                   # λ0 with the current c (strength ∝ 1/c)
                        lam_s0 = cfg["material"]["strength_MPa"] / (c_norm * base["sigma_pn0"])
                        from .objective import capacity
                        base_c["lam0"] = capacity(lam_s0, base["mu0"], st.knockdown, st.beta)[0]
                    peak = PeakGPU().start()
                    fe = model.fe_eval(rho, st, base_c, c_norm, prev_modes)
                    t_fe = time.time() - t0
                    gpu_peak = peak.stop()
                    if not np.isfinite(fe["J"]):
                        raise FloatingPointError(f"non-finite objective at phase {ph} iteration {it}")
                    t0 = time.time()
                    cx = dm.cxp                                  # element-box array module (CuPy on the GPU chain)
                    g_rho = cx.zeros(disc.n_elem)
                    g_rho[cx.asarray(model.grid_fe.elem_dense)] = cx.asarray(fe["dJ"])
                    gJd, gJa = dm.pull_rho_hat(sab, g_rho)
                    del g_rho
                    snom = dm.forward(xi[:nv], xi[nv:], as_built=False)
                    V = dm.volume(snom)
                    gVd, gVa = dm.volume_grad(snom)
                    Gs, _ = dm.sliver(snom)
                    gGd, gGa = dm.sliver_grad(snom)
                    gJ = np.concatenate([gJd, gJa])[act]
                    gG = np.concatenate([gGd, gGa])[act]
                    gV = np.concatenate([gVd, gVa])[act] / model.V0
                    t_pull = time.time() - t0
                    if s_J is None:
                        s_J = 10.0 / max(np.abs(gJ).max() * opt["move"] * 10, 1e-12)   # |Δf0| ~ 10 per full move
                    f0 = s_J * (-fe["J"] + w_s * Gs)
                    df0 = s_J * (-gJ + w_s * gG)
                    e = V / model.V0 - 1.0
                    g = np.array([e - tau, -e - tau]) / tau
                    dg = np.vstack([gV, -gV]) / tau
                    x_act = xi[act]
                    t1 = time.time()
                    x_new = optimiser.step(x_act, f0, df0, g, dg)
                    t_mmasub = time.time() - t1
                    xi_new = xi.copy()
                    xi_new[act] = x_new
                    t_mma = time.time() - t0
                    t0 = time.time()
                    xi_new, eta, nres = restore_mass(xi_new, vol, vol_grad, model.V0, act, tol=opt["mass_tol"])
                    t_res = time.time() - t0
                    dxi_raw = float(np.abs(xi_new - xi).max())
                    filt = model.ds.filt                         # convergence on the physical (filtered) design: raw ξ
                    dxi = float(max(np.abs(filt.apply(xi_new[:nv] - xi[:nv])).max(),   # 2-cycles in the filter's
                                    np.abs(filt.apply(xi_new[nv:] - xi[nv:])).max()))  # near-null space are harmless
                    xi = xi_new
                    mass_err = vol(xi) / model.V0 - 1.0
                    prev_modes = fe["modes"]
                    if st.w_P > 0 and np.isfinite(fe["quantile"]):            # Le et al. (2010) normalisation update
                        c_norm = float(np.sqrt(c_norm) * np.sqrt(fe["quantile"] / fe["sigma_pn"]))
                    if hasattr(optimiser, "note_objective"):
                        optimiser.note_objective(fe["J"])
                    hours = hours0 + (time.time() - self.t_begin) / 3600
                    row = dict(start=self.start, phase=ph, iter=it, J=fe["J"], C_C0=fe["C"] / base["C0"],
                               lam_s_ratio=fe["lam_s"] / base["lam_s0"] if st.w_P > 0 else np.nan,
                               lam_b_ratio=fe["lam_b"] / base["lam_b0"] if st.w_P > 0 and np.isfinite(base["lam_b0"]) else np.nan,
                               lam_ratio=fe["lam"] / base_c["lam0"] if st.w_P > 0 else np.nan, governing=fe["governing"],
                               mass_err=mass_err, G_s=Gs, max_dxi=dxi, max_dxi_raw=dxi_raw,
                               move=float(np.ravel(getattr(getattr(optimiser, "state", None), "move", opt["move"]))[0]),
                               n_restore=nres, eig_iters=fe["eig_iters"], eig_converged=fe["eig_converged"],
                               active_dof=model.grid_fe.ndof, fe_engine=model.fe_engine, t_geometry=t_geo, t_fe=t_fe, t_mma=t_mma, t_restore=t_res,
                               gpu_mem_GiB=_gpu_mem(), gpu_peak_GiB=gpu_peak, t_pull=t_pull, t_mmasub=t_mmasub,
                               t_upload=getattr(model, "t_upload", 0.0), c_norm=c_norm, sliver_w=w_s, stress_p=st.stress_p, beta=st.beta)
                    self._history(row)
                    self.log(f"[{self.start}] ph {ph} it {it:3d}  J {fe['J']:+.5f}  C/C0 {row['C_C0']:.4f}  "
                             f"V/V0-1 {mass_err:+.1e}  Gs {Gs:.3f}  |dξ̃| {dxi:.4f} (raw {dxi_raw:.2f})  FE {t_fe:.1f}s geo {t_geo:.1f}s")
                    if abs(mass_err) > opt["mass_tol"] * 1.0001:
                        raise RuntimeError(f"R1 violated after restoration: V/V0 - 1 = {mass_err:.2e}")
                    rs = RunState(self.start, ph, it + 1, xi, c_norm, s_J,
                                  optimiser.state_dict() if hasattr(optimiser, "state_dict") else None, rs.D_comp, hours)
                    self.save(rs)
                    if J_prev is not None and dxi < opt["tol_x"] and abs(fe["J"] - J_prev) < opt["tol_f"]:
                        calm += 1
                    else:
                        calm = 0
                    J_prev = fe["J"]
                    if calm >= 5:
                        self.log(f"[cshell] phase {ph} converged at iteration {it}")
                        break
                    if self.max_hours and (time.time() - self.t_begin) / 3600 >= self.max_hours:
                        self.log(f"[cshell] --max-hours reached; checkpoint written, resume with --resume")
                        return dict(status="paused", phase=ph, iteration=it + 1)
                    if self._check_extension(model, xi, nv):
                        extended = True
                        break
                if extended:                                     # §3.5: double D_comp, rebuild, re-map, restore, reset MMA
                    _, _, pending_d, _, _, _ = model.dm.vertex_fields(xi[:nv], xi[nv:])
                    model.release()
                    D_new = 2.0 * rs.D_comp
                    self.log(f"[cshell] outward extent extended {rs.D_comp:.1f} -> {D_new:.1f} mm at phase {ph} "
                             f"iteration {rs.it}")
                    rs = RunState(self.start, ph, rs.it, xi, c_norm, s_J, None, D_new, rs.hours_used)
                    continue
                break
            stats = phase_stats(model, xi, nv)
            self.log(f"[cshell] phase {ph} design: " + json.dumps(stats))
            result[f"phase_{ph}"] = dict(J=J_prev, xi=xi, **stats)
            np.save(os.path.join(self.out, f"xi_{self.start}_phase{ph}.npy"), xi)
            rs = RunState(self.start, ph + 1, 0, xi, c_norm, None, None, rs.D_comp, rs.hours_used)
            model.release()
        return dict(status="done", **{k: {kk: vv for kk, vv in v.items() if kk != "xi"} for k, v in result.items()})

    def _check_extension(self, model: PhaseModel, xi, nv) -> bool:
        """THEORY §3.5: True if an active vertex whose d^hi is set by D_comp has d > 0.8 D_comp."""
        _, _, d, _, _, _ = model.dm.vertex_fields(xi[:nv], xi[nv:])
        hit = model.d_hi_by_comp & ~model.ds.passive & (d > 0.8 * model.D_comp)
        if hit.any():
            self.log(f"[cshell] {int(hit.sum())} vertices reach 0.8 D_comp = {0.8 * model.D_comp:.2f} mm (§3.5)")
        return bool(hit.any())


def phase_stats(model: PhaseModel, xi, nv) -> dict:
    """Area-weighted design statistics on the active vertices: closed share (a ≥ r_MA), ξ̃ᵃ ≥ 0.98 share,
    wall w = a + d where open, outward offset d (THEORY §3.8, §8.2)."""
    _, xt_a, d, a, _, _ = model.dm.vertex_fields(xi[:nv], xi[nv:])
    act = ~model.ds.passive
    A = model.ctx.surf.A * act
    At = float(A.sum())
    closed = a >= model.ctx.r_ma
    opened = act & ~closed
    w = (a + d)[opened]
    return dict(active_vertices=int(act.sum()), closed_area_frac=float((A * closed).sum() / At),
                snap_closed_area_frac=float((A * (xt_a >= 0.98)).sum() / At),
                d_max=float(d[act].max()), d_mean=float((A * d).sum() / At),
                wall_open_min=float(w.min()) if w.size else None, wall_open_max=float(w.max()) if w.size else None)


def geometry_device(choice: str) -> str:
    """`grid.geometry_device`: "auto" (GPU when CuPy sees a CUDA device), "cpu" or "gpu" (error without CUDA)."""
    if choice not in ("auto", "cpu", "gpu"):
        raise ValueError(f"grid.geometry_device must be auto, cpu or gpu, not {choice!r}")
    if choice == "cpu":
        return "cpu"
    from .fe_sparse import _cuda_ok
    if _cuda_ok():
        return "gpu"
    if choice == "gpu":
        raise RuntimeError("grid.geometry_device = 'gpu' but CuPy finds no CUDA device")
    return "cpu"


class PeakGPU:
    """Peak device memory in use (driver view, includes cuDSS) between start() and stop(), sampled every 20 ms."""

    def __init__(self, dt: float = 0.02):
        import threading
        self.dt, self.peak, self._stop, self._th = dt, float("nan"), threading.Event(), None
        self._threading = threading

    def start(self):
        try:
            import cupy as cp
            cp.cuda.runtime.memGetInfo()
        except Exception:
            return self
        self.peak = _gpu_mem()
        self._stop.clear()

        def loop():
            while not self._stop.wait(self.dt):
                self.peak = max(self.peak, _gpu_mem())
        self._th = self._threading.Thread(target=loop, daemon=True)
        self._th.start()
        return self

    def stop(self) -> float:
        if self._th is not None:
            self._stop.set()
            self._th.join()
            self._th = None
            self.peak = max(self.peak, _gpu_mem())
        return self.peak


def free_device_memory() -> None:
    """Collect unreachable objects (the nvmath/cuDSS solver of a freed SparseFE sits in reference cycles and keeps its
    factor alive until the cyclic GC runs) and return CuPy's cached blocks to the driver."""
    import gc
    gc.collect()
    try:
        import cupy as cp
        cp.get_default_memory_pool().free_all_blocks()
    except Exception:                                        # no CuPy, or CuPy without a CUDA device (CPU nodes)
        pass


def _gpu_free() -> float:
    try:
        import cupy as cp
        return cp.cuda.runtime.memGetInfo()[0] / 2 ** 30
    except Exception:
        return float("nan")


def _gpu_mem() -> float:
    try:
        import cupy as cp
        free, total = cp.cuda.runtime.memGetInfo()
        return (total - free) / 2 ** 30
    except Exception:
        return float("nan")


def merge_history(out_dir: str) -> None:
    """history.csv = all history_<start>.csv (columns unioned, rows in start then file order), written to a temporary
    file and renamed, so concurrent writers never interleave and readers never see a partial file."""
    files = sorted(glob.glob(os.path.join(out_dir, "history_*.csv")))
    if not files:
        return
    rows, cols = [], list(HIST_COLS)
    for fn in files:
        with open(fn, newline="") as f:
            r = csv.DictReader(f)
            cols += [c for c in (r.fieldnames or []) if c not in cols]
            rows += list(r)
    tmp = os.path.join(out_dir, f".history.csv.{os.getpid()}")
    with open(tmp, "w", newline="") as f:
        w = csv.DictWriter(f, fieldnames=cols, extrasaction="ignore")
        w.writeheader()
        w.writerows(rows)
    os.replace(tmp, os.path.join(out_dir, "history.csv"))


def run_info(out_dir: str, extra: dict):
    """Record one allocation in run_info.d/<host>_<job>_<pid>.json and rebuild run_info.json from all of them (atomic
    rename; concurrent jobs of one run never lose each other's entries)."""
    d = os.path.join(out_dir, "run_info.d")
    old = os.path.join(out_dir, "run_info.json")
    if not os.path.isdir(d):
        os.makedirs(d, exist_ok=True)
        if os.path.exists(old):                          # allocations recorded before run_info.d existed
            json.dump(json.load(open(old)).get("allocations", []), open(os.path.join(d, "_legacy.json"), "w"))
    rec = dict(host=socket.gethostname(), slurm_job=os.environ.get("SLURM_JOB_ID"),
               time=time.strftime("%Y-%m-%d %H:%M:%S"), **extra)
    json.dump(rec, open(os.path.join(d, f"{rec['host']}_{rec['slurm_job']}_{os.getpid()}.json"), "w"), indent=1)
    leg = os.path.join(d, "_legacy.json")
    allocs = json.load(open(leg)) if os.path.exists(leg) else []
    for fn in sorted(glob.glob(os.path.join(d, "[!_]*.json")), key=os.path.getmtime):
        allocs.append(json.load(open(fn)))
    tmp = os.path.join(out_dir, f".run_info.json.{os.getpid()}")
    json.dump(dict(allocations=allocs), open(tmp, "w"), indent=1)
    os.replace(tmp, old)
