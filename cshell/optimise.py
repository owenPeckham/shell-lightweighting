"""Phases, starts, the MMA loop, mass restoration, checkpoints and history (THEORY §7, §5.4).

One `Run` per (config, start). Each continuation phase builds a `PhaseModel` (band, P, discretisation, design space,
FE boundary conditions, solid baseline) on its own grids; ξ lives on Γ_h and carries over unchanged (§7.3).
The FE engine is the active-set cuDSS engine (fe_sparse): the active set is every element with ρ_e > 0 dilated by
`grid.active_margin_mm`; it is rebuilt (new sparsity pattern) whenever the design puts material outside it, so the
gradient stays exact (elements outside carry ρ = 0 and H′ = 0).

PLAN_single_run (Owen 2026-10-31): one start from the solid (§1); ξ̃ = clip(F² ξ), r_MA smoothed, d tapered towards
passive regions (§2); compliance objective with buckling and strength as MMA constraints relative to the solid, the
stress exclusion zone r_ex = max(exclude_bc_mm, 2 r_MA) around constrained and loaded faces, and the growth of the
hollow-allowed set H = {R_v ≤ R_k} through the quantiles of the shear/flexure ratio R_v (§4). Contracts C1/C2 (§7):
`evaluate`, `vertex_fields`, `load_design` and the Run outputs (history.csv, xi_<start>_phase<k>.npy,
design_<start>_phase<k>.npz, fields.vtp, bc_faces.npz).
"""
from __future__ import annotations

import csv
import glob
import math
import resource
from contextlib import contextmanager
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
from .design import (DesignMap, DesignSpace, DeviceDesignMap, DeviceDisc, Discretisation, design_vertex_fields,
                     passive_thin, solid_start_xi_a, taper_weight)
from .fe_sparse import ActiveGrid, SparseFE
from .mma import MMA, ProjectedLBFGSB, restore_mass
from .objective import (ElementStressSampler, Settings, evaluate_compliance, shear_flexure_ratio, wall_energies)
from .objective import evaluate as evaluate_weighted
from .surface import (DesignSurface, RepeatedFilter, medial_depth, mount_face_ids, mount_vertices, outward_clearance,
                      surface_distance)

# C2 history columns first (PLAN_single_run §7), then diagnostics and timings (t_mma = t_pull + t_mmasub + bookkeeping)
HIST_COLS = ["start", "it", "phase", "growth_step", "R_k", "J", "K_K0", "P_P0", "lamB_ratio", "sig_ratio", "mass_err",
             "closed_frac", "hollow_allowed_frac", "governing", "lam_ratio", "G_s", "max_dxi", "max_dxi_raw", "move",
             "n_restore", "eig_iters", "eig_converged", "active_dof", "fe_engine", "t_geometry", "t_fe", "t_R", "t_mma",
             "t_pull", "t_mmasub", "t_restore", "t_upload", "t_iter", "gpu_mem_GiB", "gpu_peak_GiB", "host_peak_GiB",
             "c_norm", "sliver_w", "stress_p", "beta"]


# ---------------------------------------------------------------------------------------------------------------------
class Context:
    """Phase-independent data: geometry, faces, BCs, design surface and its per-vertex geometry (THEORY §3.4-3.5)."""

    def __init__(self, cfg: dict, log=print):
        from dshell import geometry
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
        self.N = N
        self.r_ma_raw = medial_depth(s.V, N, s.query0, h_g=h_fine)
        r_sm = float(d.get("r_ma_smooth_mm", 0.0) or 0.0)       # PLAN_single_run §2: smooth r_MA once (~1 mm)
        self.r_ma = s.filter(r_sm).apply(self.r_ma_raw) if r_sm > 0 else self.r_ma_raw.copy()
        self.r_out = outward_clearance(s.V, N, s.query0)
        self.keepout = np.full(s.n_vertices, np.inf)          # no keep-outs needed for Study 1 (D1 of §14)
        if cfg["design"]["keep_out"]:
            raise NotImplementedError("keep_out bodies: surface.keepout_distance wiring is a Stage 1 TODO")
        self.info = dict(n_vertices=s.n_vertices, hausdorff=s.info.get("hausdorff"), mount_faces=self.mount_ids,
                         setup_s=time.time() - t0)
        log(f"[cshell] context: {s.n_vertices} vertices, mounts {len(self.mount_ids)} faces, "
            f"r_MA {np.median(self.r_ma):.2f} mm median (smoothed {r_sm} mm: |Δ| p99 "
            f"{np.percentile(np.abs(self.r_ma - self.r_ma_raw), 99):.3f} mm), {time.time() - t0:.0f} s")

    def section_depth(self) -> np.ndarray:
        """True section depth per Γ_h vertex (§8 item 3): the largest raw r_MA within design.section_radius_mm
        (default 3 mm; covers the cells near convex edges, whose closest point is an edge vertex of small r_MA but
        whose depth is the section's), smoothed with the r_MA filter (design.r_ma_smooth_mm), never below r_MA."""
        if getattr(self, "_r_sec", None) is None:
            from scipy.spatial import cKDTree
            d = self.cfg["design"]
            R = float(d.get("section_radius_mm", 3.0))
            X = self.surf.V
            tree = cKDTree(X)
            S = tree.sparse_distance_matrix(tree, R, output_type="coo_matrix")
            rmax = self.r_ma_raw.copy()
            np.maximum.at(rmax, S.row, self.r_ma_raw[S.col])
            del S, tree
            r_sm = float(d.get("r_ma_smooth_mm", 0.0) or 0.0)
            rs = self.surf.filter(r_sm).apply(rmax) if r_sm > 0 else rmax
            self._r_sec = np.maximum(rs, self.r_ma)
        return self._r_sec

    def write_bc_faces(self, out_dir: str) -> str:
        """bc_faces.npz (C2): V, T (input tessellation), support (bool per triangle), load (load index per triangle,
        −1 none), support_faces (B-rep face ids), load_faces (n_loads × max, −1 padded), force_N (n_loads × 3)."""
        load_id = np.full(len(self.T), -1, np.int32)
        for i, L in enumerate(self.loads):
            load_id[np.isin(self.TF, L["faces"])] = i
        nmax = max(len(L["faces"]) for L in self.loads)
        lf = np.full((len(self.loads), nmax), -1, np.int64)
        for i, L in enumerate(self.loads):
            lf[i, :len(L["faces"])] = L["faces"]
        path = os.path.join(out_dir, "bc_faces.npz")
        np.savez_compressed(path, V=self.V.astype(np.float32), T=self.T.astype(np.int32),
                            support=np.isin(self.TF, self.support), load=load_id,
                            support_faces=np.asarray(self.support, np.int64), load_faces=lf,
                            force_N=np.array([L["force"] for L in self.loads], float))
        return path


# ---------------------------------------------------------------------------------------------------------------------
class PhaseModel:
    """Grids, P, design space, BCs and solid baseline of one continuation phase (THEORY §7.3)."""

    def __init__(self, ctx: Context, phase: int, D_comp: float, log=print, hollow=None, shift=None):
        cfg = ctx.cfg
        self.ctx, self.phase, self.D_comp, self.log = ctx, phase, float(D_comp), log
        ph = cfg["grid"]["phases"][phase]
        self.h_e, self.h_g = float(ph["h_e"]), float(ph["h_g"])
        self.k = int(round(self.h_e / self.h_g))
        t0 = time.time()
        with grid_shift(shift):
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
        ds, info = design_space(ctx, phase, D_comp, hollow)
        self.d_hi_by_comp, self.passive_parts, self.r_f = info["d_hi_by_comp"], info["passive_parts"], info["r_f"]
        self._dist_passive = info["dist_passive"]
        self.hollow = info["hollow"]
        self.ds, self.disc = ds, disc
        self.dm = DeviceDesignMap(ds, disc) if self.geometry_device == "gpu" else DesignMap(ds, disc)
        # solid: ψ = φ (nominal), ψ_ab = φ + b_o (THEORY §3.8, §3.9)
        self.V0 = self.dm.solid_volume()
        self.rho0 = self.dm.solid_rho(m["outer_bias_mm"])
        nv = ctx.surf.n_vertices                      # closed design (ξ̃ᵃ = 1): spurious void where a^hi < local depth
        self.closed_deficit = self.dm.volume(self.dm.forward(np.zeros(nv), np.ones(nv), as_built=False)) / self.V0 - 1
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
            f"{eshape}, passive {self.passive_parts}, V0 {self.V0:.1f} mm³, closed-design V/V0-1 "
            f"{self.closed_deficit:+.2e}, stress mask {self.stress_mask_info}, {self.t_setup:.0f} s")
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
        """m_e = 0 in the exclusion zone (PLAN_single_run §4: within r_ex = max(exclude_bc_mm, f·r_MA at that face)
        of any constrained or loaded face, f = capacity.exclude_rma_factor = 2, about one section depth) and in mounts
        unless include_mounts (THEORY §4.4(b))."""
        from .fea import nodes_near_faces
        ctx, cap = self.ctx, self.ctx.cfg["capacity"]
        faces = sorted(set(ctx.support) | {f for L in ctx.loads for f in L["faces"]})
        r_ex = exclusion_radii(ctx.r_ma, ctx.surf.face_id, faces, cap["exclude_bc_mm"],
                               cap.get("exclude_rma_factor", 2.0))
        src = [(nodes_near_faces(self.origin, self.h_e, self.node_shape, ctx.V, ctx.T, ctx.TF, [f]), r_ex[f])
               for f in faces]
        near = exclusion_mask(self.node_shape, self.h_e, src)
        m = ~near
        if not cap["include_mounts"]:
            frac = self.dm.elem_mean(self.ctx.mount.astype(float))
            m &= frac.reshape(self.elem_shape) < 0.5
        self.r_ex = r_ex
        fid = np.asarray(ctx.surf.face_id)
        self.stress_mask_info = dict(r_ex_mm={int(f): round(float(r_ex[f]), 2) for f in faces},
                                     n_vertices={int(f): int((fid == f).sum()) for f in faces},
                                     excluded=round(float(near.mean()), 4))
        return m.ravel()

    # -- hollow-allowed set (PLAN_single_run §4) --------------------------------------------------------------------
    def set_hollow(self, hollow) -> None:
        """H (bool per vertex, None = everywhere): ξ̃ᵃ ← 1 − τ_a(1 − ξ̃ᵃ), τ_a = 0 outside H and on passive vertices,
        rising to 1 over r_f inside H (hard 0/1 with design.taper = false)."""
        self.hollow = None if hollow is None else np.asarray(hollow, bool).copy()
        self.ds.taper_a = taper_a(self.ctx, self.ds.passive, self.hollow, self.r_f, self._dist_passive)

    def active_mask(self) -> np.ndarray:
        """MMA variables: ξᵈ on non-passive vertices, ξᵃ on non-passive vertices in H."""
        act_d = ~self.ds.passive
        act_a = act_d if self.hollow is None else act_d & self.hollow
        return np.concatenate([act_d, act_a])

    def shear_flexure(self, fe_res: dict, xi, n_samples: int = 16) -> np.ndarray:
        """R_v at the design xi from the FE solution in `fe_res` (PLAN_single_run §4): through-thickness energies from
        the outer surface (offset d) to the medial axis (to the void surface a where open), smoothed with the design filter F (once) as
        R_v = F U_shear / F U_flex (smoothing the energies, not the ratio, keeps one near-zero U_flex local)."""
        ctx, fe = self.ctx, self.fe
        nv = ctx.surf.n_vertices
        sig = fe.centre_stress(fe_res["primal"].u, solid=True)
        smp = ElementStressSampler(sig, fe_res["rho_active"], self.grid_fe.elem_dense, self.elem_shape, self.origin,
                                   self.h_e, xp=fe.xp)
        _, _, d, a, _, _ = self.dm.vertex_fields(xi[:nv], xi[nv:])
        mat = ctx.cfg["material"]                        # material only: outer surface to the void surface or r_MA
        Uf, Us = wall_energies(smp, ctx.surf.V, ctx.N, d, np.minimum(a, ctx.r_ma), mat["E_MPa"], mat["nu"],
                               n_samples, xp=fe.xp)
        del smp, sig
        F = self.ds.filt.base if isinstance(self.ds.filt, RepeatedFilter) else self.ds.filt
        return shear_flexure_ratio(np.maximum(F.apply(Uf), 0.0), np.maximum(F.apply(Us), 0.0))

    def surface_von_mises(self, fe_res: dict, d) -> np.ndarray:
        """Material von Mises stress (MPa) at each Γ_h vertex, sampled half an element inside the outer surface
        x + (d − h_e/2) n from the element-centre stresses of `fe_res` (ρ-weighted interpolation)."""
        from .objective import von_mises
        ctx, fe = self.ctx, self.fe
        xp = fe.xp
        sig = fe.centre_stress(fe_res["primal"].u, solid=True)
        smp = ElementStressSampler(sig, fe_res["rho_active"], self.grid_fe.elem_dense, self.elem_shape, self.origin,
                                   self.h_e, xp=xp)
        n = ctx.N / np.maximum(np.linalg.norm(ctx.N, axis=1, keepdims=True), 1e-300)
        P = ctx.surf.V + (np.asarray(d, float) - 0.5 * self.h_e)[:, None] * n
        s, _ = smp(xp.asarray(P))
        vm, _ = von_mises(xp, s)
        return np.asarray(vm.get() if hasattr(vm, "get") else vm, float)

    def solid_surface_von_mises(self) -> np.ndarray:
        """`surface_von_mises` of the solid (ρ0, compliance solve only; may rebuild the active set)."""
        import dataclasses
        st = dataclasses.replace(self.settings(1.0), objective="compliance", constraints=(), buckling=False)
        fe0 = self.fe_eval(self.rho0, st, None)
        return self.surface_von_mises(fe0, np.zeros(self.ctx.surf.n_vertices))

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
        engine = {"sparse": "direct"}.get(fc.get("engine", "twogrid"), fc.get("engine", "twogrid"))
        if fc["backend"] == "cpu_direct":                        # explicit SciPy direct solve (CPU tests)
            engine = "direct"
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
                       maxiter=tg.get("maxiter", 100), nrhs=tg.get("nrhs", 7),
                       deterministic=tg.get("deterministic"))
        return g, fe, "twogrid"

    def settings(self, frac: float) -> Settings:
        c = self.ctx.cfg
        cap, o = c["capacity"], c["objective"]
        typ = o.get("type", "compliance")
        cons = tuple(o.get("constraints", ())) if typ == "compliance" else ()
        if not cap["buckling"]:
            cons = tuple(x for x in cons if x != "buckling")
        if typ == "weighted" and o.get("floors_as_constraints", False):   # §8: K/K0 ≥ f_K, P/P0 ≥ f_P (MMA)
            cons = ("floor_K",) + (("floor_P",) if o["w_capacity"] > 0 else ())
        return Settings(w_K=o["w_stiffness"], w_P=o["w_capacity"], sigma_f=c["material"]["strength_MPa"],
                        knockdown=cap["knockdown"], beta=ccfg.phase_value(cap["softmin_beta"], self.phase, frac),
                        stress_p=ccfg.phase_value(cap["stress_p"], self.phase, frac), stress_q=cap["stress_q"],
                        quantile=cap["stress_quantile"],
                        buckling=cap["buckling"] and (o["w_capacity"] > 0 if typ == "weighted" else True),
                        num_modes=cap["buckling_modes"], buckling_p=cap["buckling_p"],
                        min_density=cap["min_density_geometric"], eig_tol=cap["eig_tol"],
                        rho_min=c["grid"]["rho_min"], objective=typ, constraints=cons)

    def fe_eval(self, rho: np.ndarray, st: Settings, base, c_norm: float = 1.0, prev_modes=None, want_all=False):
        """FE responses for element densities ρ (dense element box). With the GPU chain and `grid.geometry_offload`
        the compact band leaves the device for the FE evaluation, so the factor sees the same free memory as with
        the CPU chain (same fp64/fp32 choice); `t_upload` records the copy back."""
        off = self.geometry_device == "gpu" and self.geometry_offload
        if off:
            self.disc.release()
        try:
            return self._fe_eval(rho, st, base, c_norm, prev_modes, want_all)
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

    def _fe_eval(self, rho: np.ndarray, st: Settings, base, c_norm: float, prev_modes=None, want_all=False):
        if self.ensure_active(rho):
            prev_modes = None                                # new DOF numbering: warm-start modes no longer apply
        g = self.grid_fe
        r = g.gather_elems(rho)
        rho_hat = self.disc.rho_min + (1 - self.disc.rho_min) * r
        m = g.gather_elems(self.stress_mask_dense).astype(float)
        if st.objective == "compliance":                   # PLAN_single_run §4
            ev = evaluate_compliance(self.fe, rho_hat, r, m, st, base, prev_modes,
                                     want_stress=True if want_all else None,
                                     want_buckling=st.buckling if want_all else None)
            bk = ev.parts["buckling"]
            out = dict(J=ev.J, dJ=ev.dJ, C=ev.C, g=ev.g, dg=ev.dg, sigma_pn=ev.sigma_pn, mu_tilde=ev.mu_tilde,
                       governing=ev.governing, eig_iters=bk.iterations if bk else 0,
                       eig_converged=bk.converged if bk else True, modes=bk.modes if bk else None,
                       primal=ev.parts["primal"], rho_active=r, lam=np.nan, lam_s=np.nan, quantile=np.nan,
                       lam_b=st.knockdown / ev.mu_tilde if ev.mu_tilde > 0 else np.inf,
                       lamB_ratio=np.nan, sig_ratio=np.nan)
            if base is not None:
                if bk is not None:
                    out["lamB_ratio"] = (base["mu0"] / ev.mu_tilde if ev.mu_tilde > 0 else np.inf) \
                        if base.get("mu0", 0) > 0 else np.inf
                if np.isfinite(ev.sigma_pn) and np.isfinite(base.get("sigma_pn0", np.nan)):
                    out["sig_ratio"] = ev.sigma_pn / base["sigma_pn0"]
            return out
        if st.w_P == 0:                                     # compliance only (Stage 1): skip stress and buckling
            self.fe.factor(rho_hat)
            pr = self.fe.primal()
            dC = self.fe.dC_drho_hat(pr)
            J = 0.0 if base is None else float(np.log(base["C0"] / pr.compliance))
            dJ = -dC / pr.compliance
            out = dict(J=J, dJ=dJ, C=pr.compliance, lam_s=np.nan, mu_tilde=0.0, lam=np.nan, lam_b=np.inf,
                       governing="stiffness", eig_iters=0, eig_converged=True, quantile=np.nan, sigma_pn=np.nan,
                       modes=None, primal=pr, rho_active=r, g={}, dg={}, P_P0=np.nan, lamB_ratio=np.nan,
                       sig_ratio=np.nan, K_K0=np.nan if base is None else base["C0"] / pr.compliance)
            if base is not None and "floor_K" in st.constraints:
                out["g"]["floor_K"] = float(np.log(pr.compliance / base["C0"])
                                            + np.log(self.ctx.cfg["objective"]["floor_stiffness"]))
                out["dg"]["floor_K"] = dC / pr.compliance
            return out
        ev = evaluate_weighted(self.fe, rho_hat, r, m, st, base, c_norm, prev_modes)
        bk = ev.parts["buckling"]
        out = dict(J=ev.J, dJ=ev.dJ, C=ev.C, lam_s=ev.lam_s, mu_tilde=ev.mu_tilde, lam=ev.lam, lam_b=ev.lam_b,
                   governing=ev.governing, eig_iters=bk.iterations if bk else 0,
                   eig_converged=bk.converged if bk else True, quantile=ev.parts["stress"].quantile,
                   sigma_pn=ev.sigma_pn, modes=bk.modes if bk else None, primal=ev.parts["primal"], rho_active=r,
                   g={}, dg={}, K_K0=np.nan, P_P0=np.nan, lamB_ratio=np.nan, sig_ratio=np.nan)
        if base is not None:                               # §8: floors K/K0 ≥ f_K, P/P0 ≥ f_P as g ≤ 0
            o = self.ctx.cfg["objective"]
            out["K_K0"] = base["C0"] / ev.C
            out["P_P0"] = ev.lam / base["lam0"]
            if np.isfinite(base.get("lam_b0", np.nan)) and np.isfinite(ev.lam_b):
                out["lamB_ratio"] = ev.lam_b / base["lam_b0"]
            if np.isfinite(base.get("lam_s0", np.nan)):
                out["sig_ratio"] = base["lam_s0"] / ev.lam_s * (base.get("c", 1.0) / c_norm)   # σ_PN/σ_PN0 (same c)
            if "floor_K" in st.constraints:
                out["g"]["floor_K"] = float(np.log(ev.C / base["C0"]) + np.log(o["floor_stiffness"]))
                out["dg"]["floor_K"] = ev.parts["dC"] / ev.C
            if "floor_P" in st.constraints:
                out["g"]["floor_P"] = float(np.log(o["floor_capacity"]) - np.log(ev.lam / base["lam0"]))
                out["dg"]["floor_P"] = -ev.parts["dlnlam"]
        return out

    def baseline(self, st: Settings, want_all: bool = False) -> dict:
        """Solid responses through the same pipeline (as-built, b_o) and the initial normalisation c (THEORY §4.4(b)).
        Compliance objective (PLAN_single_run §4): C0, σ_PN0 (p-norm outside the exclusion zones, c = 1) and μ̃0
        (λ_B0 = knockdown/μ̃0) of the constraints in `st` (all with want_all)."""
        if st.objective == "compliance":
            b = self.fe_eval(self.rho0, st, None, 1.0, want_all=want_all)
            return dict(C0=b["C"], sigma_pn0=b["sigma_pn"], mu0=b["mu_tilde"], lam_b0=b["lam_b"], c=1.0,
                        governing0="solid", lam0=np.nan, lam_s0=np.nan)
        b0 = self.fe_eval(self.rho0, st, None, 1.0)
        c = b0["quantile"] / b0["sigma_pn"] if np.isfinite(b0["quantile"]) else 1.0
        b = self.fe_eval(self.rho0, st, None, c)
        return dict(C0=b["C"], lam0=b["lam"], sigma_pn0=b["sigma_pn"], mu0=b["mu_tilde"], c=c, lam_s0=b["lam_s"],
                    lam_b0=b["lam_b"], governing0=b["governing"])


# ---------------------------------------------------------------------------------------------------------------------
# Design space, tapers, exclusion zone, grid shift (PLAN_single_run §2, §4; shared by PhaseModel, evaluate, export)
# ---------------------------------------------------------------------------------------------------------------------
def design_space(ctx: Context, phase: int, D_comp: float, hollow=None, h_g: float | None = None):
    """DesignSpace of continuation phase `phase` (no band needed): a^hi = r_MA + γ_c h_g (smoothed r_MA), d^hi from
    the outward clearance and D_comp, passive = mounts ∪ thin, ξ̃ = clip(Fⁿ ξ) (filter.passes), τ_d tapering d to 0
    over r_f towards passive vertices, τ_a closing everything outside H (and passive) with the same taper, closure
    projection at design.closure_xi. Returns (ds, info)."""
    cfg = ctx.cfg
    m, d = cfg["manufacturing"], cfg["design"]
    h_g = float(cfg["grid"]["phases"][phase]["h_g"]) if h_g is None else float(h_g)
    nv = ctx.surf.n_vertices
    a_hi = ctx.r_ma + d["closure_gamma"] * h_g
    d_hi = np.minimum.reduce([ctx.r_out - d["gap_mm"], ctx.keepout - d["keep_out_clearance_mm"],
                              np.full(nv, D_comp)]).clip(0.0)
    if not d["allow_outward"]:
        d_hi[:] = 0.0
    v_min = ccfg.void_min_eff(cfg)                 # §8 item 5: ≥ hole_d + clearance when escape holes are on
    thin = passive_thin(a_hi, m["t_min_mm"], v_min)
    passive = ctx.mount | thin
    r_f = ccfg.phase_value(cfg["filter"]["radius_mm"], phase)
    taper = bool(d.get("taper", False))
    dist_p = surface_distance(ctx.surf.V, ctx.surf.T, passive) if taper else None
    H = None if hollow is None else np.asarray(hollow, bool).copy()
    ds = DesignSpace(r_ma=ctx.r_ma, a_hi=a_hi, d_hi=d_hi, area=ctx.surf.A, passive=passive,
                     filt=ctx.surf.filter(r_f, int(cfg["filter"].get("passes", 1))), t_min=m["t_min_mm"],
                     v_min=v_min, b_v=m["wall_bias_mm"], b_o=m["outer_bias_mm"],
                     soft_w=d["inset_soft_mm"], closure_xi=closure_xi(cfg),
                     closure_band=float(d.get("closure_band", 0.1)))
    if d.get("closed_raw_depth", True) and getattr(ctx, "r_ma_raw", None) is not None:
        # §8 item 3: closed vertices reach the true section depth (the largest raw r_MA within section_radius_mm,
        # smoothed: `Context.section_depth`), so the closed design has no void anywhere and the void surface in the
        # closure band blends between two smooth fields
        ds.a_cl = np.maximum(a_hi, ctx.section_depth() + d["closure_gamma"] * h_g)
    if d.get("void_min_open", True):
        # sr-fix 2: open voids are at least v_min across (a ≤ r_MA − v_min/2), as the export's opening keeps them;
        # thinner voids exist only in the closure band, so the optimiser never counts mass the export fills back
        ds.a_top = np.minimum(ctx.r_ma - 0.5 * v_min, a_hi)
    r_c = float(d.get("closure_skirt_mm", 0.0) or 0.0)
    if r_c > 0 and ds.closure_xi is not None:
        # sr-v4: closure spreads over the surface (design.closure_skirt), so the void ends in a mesh-resolved taper
        # instead of a saw-tooth wall along the triangle borders of the closure band (export fins)
        ds.closure_skirt = ctx.surf.filter(r_c)
    ds.taper_d = taper_weight(dist_p, r_f) if taper else None
    ds.taper_a = taper_a(ctx, passive, H, r_f, dist_p)
    return ds, dict(d_hi_by_comp=d_hi >= D_comp - 1e-9, passive_parts=dict(mount=int(ctx.mount.sum()),
                    thin=int(thin.sum())), r_f=r_f, dist_passive=dist_p, hollow=H)


def closure_xi(cfg: dict) -> float | None:
    """ξ̃ᵃ at and above which a vertex is closed, in the optimiser and the export alike: export.closed_thr (falls back to
    design.closure_xi)."""
    v = (cfg.get("export") or {}).get("closed_thr", cfg["design"].get("closure_xi"))
    return None if v is None else float(v)


def taper_a(ctx: Context, passive, hollow, r_f: float, dist_passive=None):
    """τ_a: 0 on passive vertices and outside H, S(dist/r_f) inside (design.taper), else hard 0/1; None when H is None
    and there is no taper (the Stage 2 hard passive override)."""
    taper = bool(ctx.cfg["design"].get("taper", False))
    if hollow is None:
        return taper_weight(dist_passive, r_f) if taper else None
    closed = np.asarray(passive, bool) | ~np.asarray(hollow, bool)
    if not taper:
        return (~closed).astype(float)
    return taper_weight(surface_distance(ctx.surf.V, ctx.surf.T, closed), r_f)


def exclusion_radii(r_ma, face_id, faces, exclude_bc_mm: float, factor: float = 2.0) -> dict:
    """r_ex per B-rep face = max(exclude_bc_mm, factor · median r_MA of the design vertices on that face)."""
    out = {}
    for f in faces:
        sel = np.asarray(face_id) == f
        out[f] = max(float(exclude_bc_mm), float(factor) * (float(np.median(r_ma[sel])) if sel.any() else 0.0))
    return out


def exclusion_mask(node_shape, h: float, sources) -> np.ndarray:
    """Elements (node_shape − 1 box) with any corner node within Euclidean distance r of a source node, for
    `sources` = [(flat node ids, r), ...]. One exact EDT per distinct r, on the sources' bounding box padded by r."""
    ns = tuple(int(n) for n in node_shape)
    near = np.zeros(ns, bool)
    groups: dict = {}
    for nodes, r in sources:
        if len(nodes):
            groups.setdefault(round(float(r), 9), []).append(np.asarray(nodes, np.int64))
    for r, lst in groups.items():
        ijk = np.stack(np.unravel_index(np.concatenate(lst), ns), 1)
        pad = int(np.ceil(r / h)) + 1
        lo = np.maximum(ijk.min(0) - pad, 0)
        hi = np.minimum(ijk.max(0) + pad + 1, ns)
        sub = np.ones(tuple(hi - lo), bool)
        loc = ijk - lo
        sub[loc[:, 0], loc[:, 1], loc[:, 2]] = False
        near[lo[0]:hi[0], lo[1]:hi[1], lo[2]:hi[2]] |= ndimage.distance_transform_edt(sub, sampling=h) <= r + 1e-9
    e = np.zeros(tuple(n - 1 for n in ns), bool)
    for c in range(8):
        a, b, k = (c >> 2) & 1, (c >> 1) & 1, c & 1
        e |= near[a:a + ns[0] - 1, b:b + ns[1] - 1, k:k + ns[2] - 1]
    return e


@contextmanager
def grid_shift(shift):
    """Shift the band/FE grid origin by −s (s ≥ 0 per axis; scalar or (sx, sy, sz)) while a PhaseModel is built; the
    geometry stays put (as collect.grid_shift, which shifts all three axes by one scalar)."""
    from . import band as bandmod
    if shift is None or not np.any(np.asarray(shift, float)):
        yield
        return
    s = np.broadcast_to(np.asarray(shift, float), (3,)).copy()
    if np.any(s < 0):
        raise ValueError("grid_shift: components must be >= 0")
    orig = bandmod.make_grid

    def mg(V, h, D_comp):
        g = orig(V, h, D_comp)
        n = int(math.ceil(float(s.max()) / h)) + 1
        return bandmod.Grid(g.origin - s, g.h, tuple(int(x) + n for x in g.shape), g.n_pad)

    bandmod.make_grid = mg
    try:
        yield
    finally:
        bandmod.make_grid = orig


def weighted_quantile(x, w, q: float) -> float:
    """Quantile of x with weights w (area-weighted over Γ_h), q ∈ [0, 1]; q = 1 is max(x)."""
    x, w = np.asarray(x, float), np.asarray(w, float)
    o = np.argsort(x)
    c = np.cumsum(w[o])
    return float(x[o][min(int(np.searchsorted(c, q * c[-1] - 1e-12 * c[-1])), len(x) - 1)])


# ---------------------------------------------------------------------------------------------------------------------
# C2: evaluate a design, vertex fields, load a finished design (PLAN_single_run §7)
# ---------------------------------------------------------------------------------------------------------------------
def vertex_fields(ctx: Context, xi, phase: int = -1, D_comp: float | None = None, hollow=None,
                  h_g: float | None = None) -> dict:
    """Per-vertex fields of design ξ on Γ_h with the optimiser's map (phase `phase`): d, a, ξ̃ᵈ, ξ̃ᵃ (after taper, the
    export's closure test is ξ̃ᵃ ≥ design.closure_xi), ξ̂ᵃ, closed (a ≥ r_MA), w (= a + d where open, NaN closed),
    d_hi, a_hi, r_ma, passive, hollow_allowed, tau_d, tau_a. `hollow` is the run's H (from `load_design`); None means
    hollowing allowed everywhere, which is NOT the solid start's design when H was grown."""
    cfg = ctx.cfg
    ph = phase % len(cfg["grid"]["phases"])
    D = float(cfg["design"]["outward_extent_initial_mm"] if D_comp is None else D_comp)
    ds, info = design_space(ctx, ph, D, hollow, h_g=h_g)
    nv = ctx.surf.n_vertices
    xi = np.asarray(xi, float)
    f = design_vertex_fields(np, ds, xi[:nv], xi[nv:])
    closed = f["a"] >= ctx.r_ma
    return dict(d=f["d"], a=f["a"], xt_d=f["xt_d"], xt_a=f["xt_a"], xh_a=f["xh_a"], closed=closed,
                w=np.where(closed, np.nan, f["a"] + f["d"]), d_hi=ds.d_hi, a_hi=ds.a_hi, r_ma=ctx.r_ma,
                passive=ds.passive, hollow_allowed=np.ones(nv, bool) if hollow is None else np.asarray(hollow, bool),
                tau_d=ds.tau_d(), tau_a=ds.tau_a(), closure_xi=ds.closure_xi)


def load_design(run_dir: str, start: str = "solid", phase: int | None = None) -> dict:
    """The design of the last (or the given) finished phase: {xi, hollow (None = everywhere), D_comp, phase, R}."""
    pat = ccfg.start_file(run_dir, "design_{}_phase*.npz", start)
    files = sorted(glob.glob(pat))
    if phase is not None:
        files = [f for f in files if f.endswith(f"phase{phase}.npz")]
    if not files:
        raise FileNotFoundError(f"no design_{start}_phase*.npz in {run_dir}")
    z = np.load(files[-1])
    H = z["hollow"]
    return dict(xi=z["xi"], hollow=H.astype(bool) if H.size else None, D_comp=float(z["D_comp"]),
                phase=int(z["phase"]), R=z["R"] if z["R"].size else None)


def evaluate(ctx: Context, xi, phase: int, shift=None, D_comp: float | None = None, hollow=None, log=None) -> dict:
    """Score design ξ with the optimiser's model of `phase` (C2), on a grid shifted by `shift` (see grid_shift):
    J = ln(K/K0) = ln(C0/C), K_K0, lamB_ratio = λ_B/λ_B0 = μ̃0/μ̃, sig_ratio = σ_PN/σ_PN0 (both always evaluated),
    mass_err = V/V0 − 1, feasible (every constraint g ≤ objective.feasibility_tol), governing, plus C, C0, σ, μ̃ and V.
    The solid baseline is recomputed on the same grid. Pass the run's `hollow` (see `vertex_fields`)."""
    import dataclasses
    cfg = ctx.cfg
    log = log or getattr(ctx, "log", print)
    ph = phase % len(cfg["grid"]["phases"])
    D = float(cfg["design"]["outward_extent_initial_mm"] if D_comp is None else D_comp)
    model = PhaseModel(ctx, ph, D, log, hollow=hollow, shift=shift)
    try:
        o = cfg["objective"]
        if o.get("type", "compliance") == "weighted":     # §8: the run's own objective, J = w_K ln K/K0 + w_P ln P/P0
            st = model.settings(1.0)
            base = model.baseline(st)
            nv = ctx.surf.n_vertices
            xi = np.asarray(xi, float)
            sab = model.dm.forward(xi[:nv], xi[nv:], as_built=True)
            fe = model.fe_eval(model.dm.rho(sab), st, base, base["c"])
            V = model.dm.volume(model.dm.forward(xi[:nv], xi[nv:], as_built=False))
            tol = float(o.get("feasibility_tol", 1e-3))
            K, P = base["C0"] / fe["C"], fe["P_P0"]
            floors = dict(K=bool(np.log(K) >= np.log(o["floor_stiffness"]) - tol),
                          P=bool(st.w_P == 0 or np.log(P) >= np.log(o["floor_capacity"]) - tol))
            return dict(objective="weighted", J=fe["J"], J_K=float(np.log(K)), K_K0=K, P_P0=P,
                        lamB_ratio=fe["lamB_ratio"], sig_ratio=fe["sig_ratio"], mass_err=V / model.V0 - 1.0,
                        feasible=bool(all(floors.values())), floors=floors,
                        floor_K=o["floor_stiffness"], floor_P=o["floor_capacity"], governing=fe["governing"],
                        g=dict(fe["g"]), C=fe["C"], C0=base["C0"], lam=fe["lam"], lam0=base["lam0"], V=V,
                        V0=model.V0, h_e=model.h_e, origin=np.asarray(model.origin).tolist(), phase=ph)
        cons = ("buckling", "strength") if cfg["capacity"]["buckling"] else ("strength",)
        st = dataclasses.replace(model.settings(1.0), objective="compliance", constraints=cons,
                                 buckling=bool(cfg["capacity"]["buckling"]))
        base = model.baseline(st, want_all=True)
        nv = ctx.surf.n_vertices
        xi = np.asarray(xi, float)
        sab = model.dm.forward(xi[:nv], xi[nv:], as_built=True)
        fe = model.fe_eval(model.dm.rho(sab), st, base, want_all=True)
        V = model.dm.volume(model.dm.forward(xi[:nv], xi[nv:], as_built=False))
        tol = float(cfg["objective"].get("feasibility_tol", 1e-3))
        feasible = all(v <= tol for v in fe["g"].values())
        return dict(objective="compliance", J_K=fe["J"], P_P0=None,
                    J=fe["J"], K_K0=base["C0"] / fe["C"], lamB_ratio=fe["lamB_ratio"], sig_ratio=fe["sig_ratio"],
                    mass_err=V / model.V0 - 1.0, feasible=bool(feasible),
                    governing=fe["governing"], g=dict(fe["g"]), C=fe["C"], C0=base["C0"], sigma_pn=fe["sigma_pn"],
                    sigma_pn0=base["sigma_pn0"], mu_tilde=fe["mu_tilde"], mu0=base["mu0"], V=V, V0=model.V0,
                    h_e=model.h_e, origin=np.asarray(model.origin).tolist(), phase=ph)
    finally:
        model.release()


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
    hollow: np.ndarray | None = None      # hollow-allowed set H (None: everywhere)
    growth: dict | None = None            # growth-phase state (new_growth)


GROWTH_ARRAYS = ("best_xi", "best_H")
GROWTH_SCALARS = ("step", "R_k", "done", "in_step", "J_prev_it", "low", "best_J", "best_k")


def new_growth() -> dict:
    """Growth state: next boundary `step`, R_k, J at each boundary J_b, MMA iterations of each step n_it, iterations in
    the current step, J of the previous iteration, consecutive low-gain steps, best boundary (J, ξ, H, k), done."""
    return dict(step=0, R_k=np.nan, J_b=[], n_it=[], done=False, in_step=0, J_prev_it=None, low=0, best_J=None,
                best_xi=None, best_H=None, best_k=-1)


class Run:
    """One start of one part: all phases, checkpoints every iteration (THEORY §7.6). The solid start grows the
    hollow-allowed set in phase growth.phase (PLAN_single_run §4) and keeps it frozen afterwards."""

    def __init__(self, cfg: dict, start: str, out_dir: str, max_hours: float | None = None, log=None):
        self.cfg, self.start, self.out = cfg, ccfg.canonical_start(start), out_dir
        os.makedirs(os.path.join(out_dir, "checkpoints"), exist_ok=True)
        self.max_hours = max_hours
        self.t_begin = time.time()
        self._logf = open(os.path.join(out_dir, "run_log.txt"), "a")
        self.log = log or self._log
        self.last_R = None

    def _log(self, msg):
        print(msg, flush=True)
        self._logf.write(msg + "\n"); self._logf.flush()

    # -- starts (THEORY §7.4; PLAN_single_run §1) ------------------------------------------------------------------
    def initial_xi(self, nv: int) -> np.ndarray:
        if self.start == "uniform":
            return np.concatenate([np.full(nv, 0.5), np.zeros(nv)])
        if self.start == "neutral":
            return np.concatenate([np.full(nv, 0.5), np.full(nv, 0.5)])
        if self.start == "solid":                        # ξᵃ replaced by solid_start_xi_a (a = r_MA) in run()
            return np.concatenate([np.zeros(nv), np.ones(nv)])
        if self.start == "dshell":                       # S2: map dshell's best design (cshell.collect, §7.4)
            from .collect import dshell_start_xi
            ctx = getattr(self, "ctx", None)             # set by the CLI; otherwise build the context once
            if ctx is None:
                self.log("[cshell] dshell start: no context attached to Run, building one")
                ctx = self.ctx = Context(self.cfg, log=self.log)
            xi = dshell_start_xi(ctx, self.cfg, log=self.log)
            if len(xi) != 2 * nv:
                raise ValueError(f"dshell start has {len(xi)} entries, expected {2 * nv}")
            return xi
        raise NotImplementedError(f"start {self.start!r}")

    def _grows(self, ph: int) -> bool:
        g = self.cfg.get("growth", {})
        return self.start == "solid" and bool(g.get("enabled", True)) and ph == int(g.get("phase", 0))

    # -- checkpoints ---------------------------------------------------------------------------------------------
    def _ckpt_path(self, rs: RunState):
        return os.path.join(self.out, "checkpoints", f"{self.start}_{rs.phase}_{rs.it:04d}.npz")

    def save(self, rs: RunState):
        flat = {f"opt_{k}": v for k, v in (rs.opt_state or {}).items()}
        gr = rs.growth or {}
        none = np.zeros(0)
        np.savez(self._ckpt_path(rs), xi=rs.xi, phase=rs.phase, it=rs.it, c_norm=rs.c_norm,
                 s_J=np.nan if rs.s_J is None else rs.s_J, D_comp=rs.D_comp, hours_used=rs.hours_used,
                 rng=np.random.default_rng(self.cfg["optimiser"]["seed"]).bit_generator.state["state"]["state"],
                 hollow=none if rs.hollow is None else rs.hollow, has_growth=bool(rs.growth),
                 g_J=np.asarray(gr.get("J_b", []), float), g_n_it=np.asarray(gr.get("n_it", []), np.int64),
                 **{f"g_{k}": np.asarray(np.nan if gr.get(k) is None else gr.get(k), float) for k in GROWTH_SCALARS
                    if gr},
                 **{f"g_{k}": none if gr.get(k) is None else gr[k] for k in GROWTH_ARRAYS if gr}, **flat)
        files = sorted(glob.glob(os.path.join(self.out, "checkpoints", f"{self.start}_*.npz")))
        for f in files[:-3]:
            os.remove(f)

    def load_latest(self) -> RunState | None:
        files = []
        for name in ccfg.start_names(self.start):       # canonical first; legacy (vshell_*) only if no new ones
            files = files or sorted(glob.glob(os.path.join(self.out, "checkpoints", f"{name}_*.npz")))
        if not files:
            return None
        z = np.load(files[-1])
        opt = {k[4:]: z[k] for k in z.files if k.startswith("opt_")}
        sJ = float(z["s_J"])
        hollow = growth = None
        if "hollow" in z.files and z["hollow"].size:
            hollow = z["hollow"].astype(bool)
        if "has_growth" in z.files and bool(z["has_growth"]) and "g_in_step" in z.files:
            growth = new_growth()
            for k in GROWTH_SCALARS:
                v = float(z[f"g_{k}"])
                growth[k] = None if np.isnan(v) and k in ("J_prev_it", "best_J") else v
            for k in ("step", "in_step", "low", "best_k"):
                growth[k] = int(growth[k])
            growth["done"] = bool(growth["done"])
            growth["J_b"], growth["n_it"] = [float(j) for j in z["g_J"]], [int(n) for n in z["g_n_it"]]
            growth["best_xi"] = z["g_best_xi"] if z["g_best_xi"].size else None
            growth["best_H"] = z["g_best_H"].astype(bool) if z["g_best_H"].size else None
        return RunState(self.start, int(z["phase"]), int(z["it"]), z["xi"], float(z["c_norm"]),
                        None if np.isnan(sJ) else sJ, opt or None, float(z["D_comp"]), float(z["hours_used"]),
                        hollow, growth)

    def _history(self, row: dict):
        """Append to this start's own history_<start>.csv, then rebuild the merged history.csv atomically."""
        path = ccfg.start_file(self.out, "history_{}.csv", self.start)   # legacy history_vshell.csv on resume
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

    # -- one evaluation ------------------------------------------------------------------------------------------
    def _eval(self, model: "PhaseModel", xi, st, base, c_norm, prev_modes):
        nv = model.ctx.surf.n_vertices
        t0 = time.time()
        sab = model.dm.forward(xi[:nv], xi[nv:], as_built=True)
        rho = model.dm.rho(sab)
        t_geo = time.time() - t0
        base_c = dict(base)
        if st.objective == "weighted" and st.w_P > 0:            # λ0 with the current c (strength ∝ 1/c)
            from .objective import capacity
            lam_s0 = self.cfg["material"]["strength_MPa"] / (c_norm * base["sigma_pn0"])
            base_c["lam0"] = capacity(lam_s0, base["mu0"], st.knockdown, st.beta)[0]
        t0 = time.time()
        peak = PeakGPU().start()
        fe = model.fe_eval(rho, st, base_c, c_norm, prev_modes)
        t_fe = time.time() - t0
        gpu_peak = peak.stop()
        if not np.isfinite(fe["J"]):
            raise FloatingPointError(f"non-finite objective at phase {model.phase}")
        fe.update(t_geometry=t_geo, t_fe=t_fe, gpu_peak=gpu_peak, base_c=base_c)
        return fe, sab

    # -- main loop -----------------------------------------------------------------------------------------------
    def run(self, ctx: Context, resume: bool = False, only_phase: int | None = None) -> dict:
        cfg, opt = self.cfg, self.cfg["optimiser"]
        gcfg = cfg.get("growth", {})
        ocfg = cfg["objective"]
        con_scale = float(ocfg.get("constraint_scale", 0.01))
        nv = ctx.surf.n_vertices
        rs = self.load_latest() if resume else None
        fresh = rs is None
        if fresh:
            rs = RunState(self.start, 0, 0, self.initial_xi(nv), 1.0, None, None,
                          cfg["design"]["outward_extent_initial_mm"], 0.0)
        else:
            self.log(f"[cshell] resuming {self.start} at phase {rs.phase} iteration {rs.it}")
        ctx.write_bc_faces(self.out)
        hours0 = rs.hours_used
        phases = range(len(cfg["grid"]["phases"])) if only_phase is None else [only_phase]
        model = None
        result = {}
        solid_init = not (fresh and self.start == "solid")
        quant = [float(q) for q in gcfg.get("quantiles", [1.0])]
        max_ips = int(gcfg.get("max_iters_per_step", gcfg.get("iters_per_step", 25)))
        min_ips = min(int(gcfg.get("min_iters_per_step", 2)), max_ips)
        step_tol = float(gcfg.get("step_tol", 1e-4))
        g_tol = float(gcfg.get("tol", 1e-3))
        patience = gcfg.get("patience")                  # §8: None = every quantile runs, then the best boundary
        patience = len(quant) + 1 if patience is None else int(patience)
        A = ctx.surf.A
        for ph in phases:
            if ph < rs.phase:
                continue
            grow = self._grows(ph)
            if grow and rs.growth is None:
                rs.growth = new_growth()
                rs.hollow = np.zeros(nv, bool)
            pending_d = None
            last = None
            while True:                                          # restarts after an outward-extent extension
                if model is not None:                            # free the previous band/factor before building anew
                    model.release()
                model = dm = disc = None
                model = PhaseModel(ctx, ph, rs.D_comp, self.log, hollow=rs.hollow)
                if not solid_init:                               # §1: ξᵈ = 0, a = r_MA (geometrically solid)
                    rs.xi[:nv] = 0.0
                    rs.xi[nv:] = solid_start_xi_a(model.ds)
                    solid_init = True
                if pending_d is not None:                        # §3.5 step 2: keep d unchanged on the new d^hi
                    tgt = np.clip(pending_d / np.maximum(model.ds.d_hi, 1e-12), 0.0, 1.0)
                    xd = np.clip(model.ds.filt.invert(tgt), 0.0, 1.0)
                    rs.xi[:nv] = xd
                    pending_d = None
                dm, disc = model.dm, model.disc
                st0 = model.settings(0.0)
                base = model.baseline(st0)
                if st0.objective == "compliance":
                    self.log(f"[cshell] phase {ph} solid: C0 {base['C0']:.4f} N·mm, K0 {ctx.F @ ctx.F / base['C0']:.2f} "
                             f"N/mm, σ_PN0 {base['sigma_pn0']:.3f} MPa, λ_B0 {base['lam_b0']:.4g}, constraints "
                             f"{list(st0.constraints)}")
                else:
                    self.log(f"[cshell] phase {ph} solid: C0 {base['C0']:.4f} N·mm, K0 {ctx.F @ ctx.F / base['C0']:.2f} "
                             f"N/mm, λ_s0 {base['lam_s0']}, λ_b0 {base['lam_b0']}, c {base['c']:.3f}")
                c_norm = base["c"] if rs.it == 0 or ph != rs.phase else rs.c_norm
                cons = list(st0.constraints)

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
                act = model.active_mask()
                xi, eta, nres = restore_mass(xi, vol, vol_grad, model.V0, act, tol=opt["mass_tol"])
                extended = False
                Cls = MMA if opt["method"] == "mma" else ProjectedLBFGSB

                def new_optimiser(n):
                    return Cls(n, 2 + len(cons), move=opt["move"], con_scale=1.0)
                optimiser = new_optimiser(int(act.sum()))
                if rs.opt_state is not None and ph == rs.phase and rs.it > 0:
                    sizes = {np.size(v) for k, v in rs.opt_state.items()
                             if k.split("_", 1)[-1] in ("xold1", "xold2", "low", "upp")}
                    if sizes <= {int(act.sum())}:
                        optimiser.load_state_dict(rs.opt_state)
                    else:                                        # active set changed since the checkpoint
                        self.log("[cshell] optimiser state does not match the active set; asymptotes reset")
                it0 = rs.it if ph == rs.phase else 0
                max_it = len(quant) * max_ips + 1 if grow else ccfg.phase_value(opt["max_iter"], ph)
                s_J = rs.s_J if ph == rs.phase else None
                tau = 5e-4
                calm, J_prev, prev_modes = 0, None, None
                for it in range(it0, max_it):
                    t_it = time.time()
                    frac = it / max(max_it - 1, 1)
                    st = model.settings(frac)
                    w_s = ccfg.phase_value(opt["sliver_weight"], ph, frac)
                    fe, sab = self._eval(model, xi, st, base, c_norm, prev_modes)
                    last = (fe, xi.copy())
                    t_R = 0.0
                    gr = rs.growth if grow else None
                    step_done = False
                    if gr is not None and not gr["done"]:
                        n_in, Jp = gr["in_step"], gr["J_prev_it"]
                        step_done = (gr["step"] == 0 and n_in == 0) or n_in >= max_ips or (
                            n_in >= min_ips and Jp is not None and abs(fe["J"] - Jp) < step_tol)
                    if step_done:
                        # ---- growth boundary k (sr-fix 1: steps run to in-step convergence; patience; best) -------
                        t0 = time.time()
                        k = gr["step"]
                        gr["J_b"].append(float(fe["J"]))
                        gr["n_it"].append(int(gr["in_step"]))
                        stop, why = k >= len(quant), "all quantiles used"
                        if k >= 1:
                            gain = gr["J_b"][k] - gr["J_b"][k - 1]
                            gr["low"] = gr["low"] + 1 if gain < g_tol else 0
                            self.log(f"[cshell] growth step {k - 1} done after {gr['in_step']} iterations: ΔJ "
                                     f"{gain:+.2e} (low-gain steps in a row {gr['low']}/{patience})")
                            if gr["low"] >= patience:
                                stop, why = True, f"{patience} consecutive steps raised J by < tol {g_tol:g}"
                        if gr["best_J"] is None or fe["J"] > gr["best_J"]:
                            gr["best_J"], gr["best_xi"], gr["best_H"], gr["best_k"] = (
                                float(fe["J"]), xi.copy(), rs.hollow.copy(), k)
                        if stop:
                            gr["done"] = True
                            if gr["best_k"] != k:                # revert to the best boundary design
                                xi, rs.hollow = gr["best_xi"].copy(), gr["best_H"].copy()
                                model.set_hollow(rs.hollow)
                                fe, sab = self._eval(model, xi, st, base, c_norm, None)
                                last = (fe, xi.copy())
                                why += f"; reverted to boundary {gr['best_k']} (J {gr['best_J']:+.5f})"
                            self.log(f"[cshell] growth stops at boundary {k}: {why}; J by boundary "
                                     f"{[round(j, 5) for j in gr['J_b']]}, iterations per step {gr['n_it'][1:]}")
                            rs = RunState(self.start, ph, it, xi, c_norm, s_J, None, rs.D_comp, rs.hours_used,
                                          rs.hollow, gr)
                            self.save(rs)
                            break
                        gr["in_step"], gr["J_prev_it"] = 0, None
                        R = model.shear_flexure(fe, xi, int(gcfg.get("samples", 16)))
                        self.last_R = R
                        pool = ~model.ds.passive
                        R_k = weighted_quantile(R[pool], A[pool], quant[k])
                        rs.hollow = rs.hollow | (pool & (R <= R_k))
                        model.set_hollow(rs.hollow)
                        gr["step"], gr["R_k"] = k + 1, R_k
                        act = model.active_mask()
                        xi, eta, nres = restore_mass(xi, vol, vol_grad, model.V0, act, tol=opt["mass_tol"])
                        optimiser = new_optimiser(int(act.sum()))
                        s_J = None
                        prev_modes = fe["modes"]
                        t_R = time.time() - t0
                        hf = float((A * rs.hollow).sum() / (A * pool).sum())
                        self.log(f"[cshell] growth step {k}: q {quant[k]:.2f}, R_k {R_k:.4g}, H {hf:.3f} of the "
                                 f"active area, J at boundary {fe['J']:+.5f}, R {t_R:.1f} s")
                        fe, sab = self._eval(model, xi, st, base, c_norm, prev_modes)
                        last = (fe, xi.copy())
                    t0 = time.time()
                    cx = dm.cxp                                  # element-box array module (CuPy on the GPU chain)
                    edense = cx.asarray(model.grid_fe.elem_dense)

                    def pull(v):
                        g_rho = cx.zeros(disc.n_elem)
                        g_rho[edense] = cx.asarray(v)
                        gd, ga = dm.pull_rho_hat(sab, g_rho)
                        return np.concatenate([gd, ga])[act]
                    gJ = pull(fe["dJ"])
                    g_con, dg_con = [], []
                    for name in cons:
                        gv = fe["g"].get(name, -np.inf)
                        if np.isfinite(gv):
                            g_con.append(gv / con_scale)
                            dg_con.append(pull(fe["dg"][name]) / con_scale)
                        else:
                            g_con.append(-1.0)
                            dg_con.append(np.zeros(int(act.sum())))
                    snom = dm.forward(xi[:nv], xi[nv:], as_built=False)
                    V = dm.volume(snom)
                    gVd, gVa = dm.volume_grad(snom)
                    Gs, _ = dm.sliver(snom)
                    gGd, gGa = dm.sliver_grad(snom)
                    gG = np.concatenate([gGd, gGa])[act]
                    gV = np.concatenate([gVd, gVa])[act] / model.V0
                    t_pull = time.time() - t0
                    if s_J is None:
                        s_J = 10.0 / max(np.abs(gJ).max() * opt["move"] * 10, 1e-12)   # |Δf0| ~ 10 per full move
                    f0 = s_J * (-fe["J"] + w_s * Gs)
                    df0 = s_J * (-gJ + w_s * gG)
                    e = V / model.V0 - 1.0
                    g = np.array([e - tau, -e - tau] + g_con) / np.array([tau, tau] + [1.0] * len(cons))
                    dg = np.vstack([gV / tau, -gV / tau] + dg_con)
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
                    filt = model.ds.filt                         # convergence on the physical (filtered) design
                    dxi = float(max(np.abs(filt.apply(xi_new[:nv] - xi[:nv])).max(),
                                    np.abs(filt.apply(xi_new[nv:] - xi[nv:])).max()))
                    # design statistics of the evaluated design (consistent with J): closed a ≥ r_MA, area of H
                    pool = ~model.ds.passive
                    Ap = float((A * pool).sum())
                    closed_frac = float((A * pool * (sab.a >= ctx.r_ma)).sum() / Ap)
                    hollow_frac = 1.0 if rs.hollow is None else float((A * (pool & rs.hollow)).sum() / Ap)
                    xi = xi_new
                    mass_err = vol(xi) / model.V0 - 1.0
                    prev_modes = fe["modes"]
                    if st.objective == "weighted" and st.w_P > 0 and np.isfinite(fe["quantile"]):
                        c_norm = float(np.sqrt(c_norm) * np.sqrt(fe["quantile"] / fe["sigma_pn"]))   # Le et al. (2010)
                    if hasattr(optimiser, "note_objective"):
                        optimiser.note_objective(fe["J"])
                    hours = hours0 + (time.time() - self.t_begin) / 3600
                    row = dict(start=self.start, it=it, phase=ph, growth_step=(gr["step"] - 1) if gr else -1,
                               R_k=gr["R_k"] if gr else np.nan, J=fe["J"], K_K0=base["C0"] / fe["C"],
                               P_P0=fe.get("P_P0", np.nan), lamB_ratio=fe.get("lamB_ratio", np.nan),
                               sig_ratio=fe.get("sig_ratio", np.nan), mass_err=mass_err, closed_frac=closed_frac,
                               hollow_allowed_frac=hollow_frac, governing=fe["governing"],
                               lam_ratio=fe["lam"] / fe["base_c"]["lam0"] if st.objective == "weighted" and st.w_P > 0
                               else np.nan, G_s=Gs, max_dxi=dxi, max_dxi_raw=dxi_raw,
                               move=float(np.ravel(getattr(getattr(optimiser, "state", None), "move", opt["move"]))[0]),
                               n_restore=nres, eig_iters=fe["eig_iters"], eig_converged=fe["eig_converged"],
                               active_dof=model.grid_fe.ndof, fe_engine=model.fe_engine, t_geometry=fe["t_geometry"],
                               t_fe=fe["t_fe"], t_R=t_R, t_mma=t_mma, t_pull=t_pull, t_mmasub=t_mmasub,
                               t_restore=t_res, t_upload=getattr(model, "t_upload", 0.0), t_iter=time.time() - t_it,
                               gpu_mem_GiB=_gpu_mem(), gpu_peak_GiB=fe["gpu_peak"], host_peak_GiB=host_peak_gib(),
                               c_norm=c_norm, sliver_w=w_s, stress_p=st.stress_p, beta=st.beta)
                    self._history(row)
                    self.log(f"[{self.start}] ph {ph} it {it:3d}  J {fe['J']:+.5f}  K/K0 {row['K_K0']:.4f}  P/P0 {row['P_P0']:.4f}  λB/λB0 "
                             f"{row['lamB_ratio']:.4f}  σ/σ0 {row['sig_ratio']:.4f}  V/V0-1 {mass_err:+.1e}  closed "
                             f"{closed_frac:.3f}  H {hollow_frac:.3f}  |dξ̃| {dxi:.4f}  FE {fe['t_fe']:.1f}s "
                             f"it {row['t_iter']:.1f}s")
                    if abs(mass_err) > opt["mass_tol"] * 1.0001:
                        raise RuntimeError(f"R1 violated after restoration: V/V0 - 1 = {mass_err:.2e}")
                    if gr is not None and not gr["done"]:
                        gr["in_step"] += 1
                        gr["J_prev_it"] = float(fe["J"])
                    rs = RunState(self.start, ph, it + 1, xi, c_norm, s_J,
                                  optimiser.state_dict() if hasattr(optimiser, "state_dict") else None, rs.D_comp,
                                  hours, rs.hollow, gr if gr is not None else rs.growth)
                    self.save(rs)
                    if not grow:
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
                    rs = RunState(self.start, ph, rs.it, xi, c_norm, s_J, None, D_new, rs.hours_used, rs.hollow,
                                  rs.growth)
                    continue
                break
            stats = phase_stats(model, xi, nv)
            surf_vm = {}
            if last is not None and model.fe is not None:        # R_v and σ_vM at the last FE evaluation (fields.vtp)
                try:
                    self.last_R = model.shear_flexure(last[0], last[1], int(gcfg.get("samples", 16)))
                    d_last = model.dm.vertex_fields(last[1][:nv], last[1][nv:])[2]
                    surf_vm["stress"] = model.surface_von_mises(last[0], d_last)
                    surf_vm["stress_solid"] = model.solid_surface_von_mises()
                except Exception as e:                           # diagnostics only
                    self.log(f"[cshell] R_v / surface stress for fields.vtp failed: {e}")
            last = None
            stats["growth"] = None if rs.growth is None else dict(
                steps=int(rs.growth["step"]), R_k=float(rs.growth["R_k"]), J_b=[float(j) for j in rs.growth["J_b"]],
                iters=[int(n) for n in rs.growth["n_it"][1:]], best_boundary=int(rs.growth["best_k"]),
                best_J=rs.growth["best_J"])
            self.log(f"[cshell] phase {ph} design: " + json.dumps(stats))
            result[f"phase_{ph}"] = dict(J=J_prev, xi=xi, **stats)
            np.save(os.path.join(self.out, f"xi_{self.start}_phase{ph}.npy"), xi)
            np.savez(os.path.join(self.out, f"design_{self.start}_phase{ph}.npz"), xi=xi, phase=ph, D_comp=rs.D_comp,
                     hollow=np.zeros(0, bool) if rs.hollow is None else rs.hollow,
                     R=np.zeros(0) if self.last_R is None else self.last_R)
            self.write_fields(model, xi, rs.hollow, surf_vm)
            rs = RunState(self.start, ph + 1, 0, xi, c_norm, None, None, rs.D_comp, rs.hours_used, rs.hollow,
                          rs.growth)
            model.release()
        return dict(status="done", **{k: {kk: vv for kk, vv in v.items() if kk != "xi"} for k, v in result.items()})

    def write_fields(self, model: "PhaseModel", xi, hollow, extra: dict | None = None) -> str:
        """fields.vtp (C2) on Γ_h: d, a, w (= a + d where open, NaN closed), r_MA, R, closed, hollow_allowed, passive
        (0/1), stress and stress_solid (surface von Mises, MPa, final design and solid), plus xi_d, xi_a, xit_a,
        r_MA_raw."""
        ctx = model.ctx
        nv = ctx.surf.n_vertices
        xt_d, xt_a, d, a, _, _ = model.dm.vertex_fields(xi[:nv], xi[nv:])
        closed = a >= ctx.r_ma
        pdata = dict(d=d, a=a, w=np.where(closed, np.nan, a + d), r_MA=ctx.r_ma,
                     R=np.full(nv, np.nan) if self.last_R is None else self.last_R, closed=closed,
                     hollow_allowed=np.ones(nv, bool) if hollow is None else np.asarray(hollow, bool),
                     xi_d=xi[:nv], xi_a=xi[nv:], xit_a=xt_a, r_MA_raw=ctx.r_ma_raw,
                     passive=np.asarray(model.ds.passive, bool))
        for k in ("stress", "stress_solid"):
            pdata[k] = (extra or {}).get(k, np.full(nv, np.nan))
        path = os.path.join(self.out, "fields.vtp")
        write_vtp(path, ctx.surf.V, ctx.surf.T, pdata)
        return path

    def _check_extension(self, model: PhaseModel, xi, nv) -> bool:
        """THEORY §3.5: True if an active vertex whose d^hi is set by D_comp has d > 0.8 D_comp."""
        _, _, d, _, _, _ = model.dm.vertex_fields(xi[:nv], xi[nv:])
        hit = model.d_hi_by_comp & ~model.ds.passive & (d > 0.8 * model.D_comp)
        if hit.any():
            self.log(f"[cshell] {int(hit.sum())} vertices reach 0.8 D_comp = {0.8 * model.D_comp:.2f} mm (§3.5)")
        return bool(hit.any())


def host_peak_gib() -> float:
    """Peak resident host memory of this process (GiB, Linux ru_maxrss in KiB)."""
    return resource.getrusage(resource.RUSAGE_SELF).ru_maxrss / 2 ** 20


def write_vtp(path, V: np.ndarray, T: np.ndarray, point_data: dict) -> None:
    """ASCII VTK XML PolyData (ParaView); bool arrays as UInt8 (same format as collect.write_vtp)."""
    def arr(a):
        return " ".join(f"{x:.7g}" for x in np.asarray(a, float).ravel())
    tmp = path + f".tmp{os.getpid()}"
    with open(tmp, "w") as f:
        f.write('<?xml version="1.0"?>\n<VTKFile type="PolyData" version="0.1" byte_order="LittleEndian">\n<PolyData>\n')
        f.write(f'<Piece NumberOfPoints="{len(V)}" NumberOfPolys="{len(T)}">\n<PointData>\n')
        for k, v in point_data.items():
            v = np.asarray(v)
            typ = "UInt8" if v.dtype == bool else "Float64"
            f.write(f'<DataArray type="{typ}" Name="{k}" format="ascii">\n')
            f.write((" ".join(str(int(x)) for x in v) if v.dtype == bool else arr(v)) + "\n</DataArray>\n")
        f.write('</PointData>\n<Points>\n<DataArray type="Float64" NumberOfComponents="3" format="ascii">\n')
        f.write(arr(V) + "\n</DataArray>\n</Points>\n<Polys>\n")
        f.write('<DataArray type="Int64" Name="connectivity" format="ascii">\n'
                + " ".join(map(str, np.asarray(T, np.int64).ravel())) + "\n</DataArray>\n")
        f.write('<DataArray type="Int64" Name="offsets" format="ascii">\n'
                + " ".join(map(str, 3 * np.arange(1, len(T) + 1))) + "\n</DataArray>\n</Polys>\n</Piece>\n</PolyData>\n</VTKFile>\n")
    os.replace(tmp, path)


def phase_stats(model: PhaseModel, xi, nv) -> dict:
    """Area-weighted design statistics on the active vertices: closed share (a ≥ r_MA), ξ̃ᵃ ≥ closure_xi share,
    hollow-allowed share, wall w = a + d where open, outward offset d (THEORY §3.8, §8.2)."""
    _, xt_a, d, a, _, _ = model.dm.vertex_fields(xi[:nv], xi[nv:])
    act = ~model.ds.passive
    A = model.ctx.surf.A * act
    At = float(A.sum())
    closed = a >= model.ctx.r_ma
    opened = act & ~closed
    w = (a + d)[opened]
    xc = model.ds.closure_xi or 0.98
    H = np.ones(nv, bool) if model.hollow is None else model.hollow
    return dict(active_vertices=int(act.sum()), closed_area_frac=float((A * closed).sum() / At),
                snap_closed_area_frac=float((A * (xt_a >= xc)).sum() / At),
                hollow_allowed_area_frac=float((A * H).sum() / At),
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
            for row in r:                                # start column "vshell" in runs from before the rename
                if row.get("start"):
                    row["start"] = ccfg.canonical_start(row["start"])
                rows.append(row)
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
