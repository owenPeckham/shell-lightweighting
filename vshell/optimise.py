"""
The optimisation run: STEP in, constant-mass variable-thickness design (or the original solid) out.

Stages
  1. geometry      STEP -> face table, supports/loads, tessellation, volume V0
  2. field         signed-distance grid of the original part
  3. segment       mounts, joints, struts (N control regions)
  4. baseline      FE of the original solid: stiffness K0, load capacity P0, shear/bending ratio R_i
  5. rank          candidate regions ordered by how shear-critical they are (R_i, highest first)
  6. mass sweep    design k = the k most shear-critical regions solid, the rest hollow; the outer offset
                   s_k that restores the original mass; feasibility s_k >= t_min
  7. search        FE of feasible designs (budgeted coarse sweep + local refinement over k)
  8. decide        J = w_K ln(K/K0) + w_P ln(P/P0); accept the best hollow design only if it gains
                   (J > ln(1 + min_gain)) without dropping K or P below their floors; else keep the solid
  9. outputs       final mesh, CSV/JSON logs, arrays for the report, report.pdf
"""
from __future__ import annotations

import os
import shutil
import time
import traceback

import numpy as np

from . import config as config_mod
from . import fe, field, geometry, mechanics, segment, util


class Run:
    def __init__(self, cfg: dict, config_file: str | None = None):
        self.cfg = cfg
        self.config_file = config_file
        self.out = os.path.join(cfg["output_dir"], cfg["name"])
        os.makedirs(self.out, exist_ok=True)
        self.v = bool(cfg.get("verbose", True))
        self.timings = {}
        self.evaluations = []
        self.notes = []

    def log(self, msg):
        if self.v:
            print(msg, flush=True)
        with open(os.path.join(self.out, "run_log.txt"), "a", encoding="utf-8") as f:
            f.write(msg + "\n")

    def _t(self, key, t0):
        self.timings[key] = self.timings.get(key, 0.0) + time.time() - t0

    # ------------------------------------------------------------------------------------------
    def setup(self):
        cfg = self.cfg
        t0 = time.time()
        self.log(f"[vshell] {cfg['name']}: reading {cfg['step_file']}")
        self.shape = geometry.load_step(cfg["step_file"])
        self.faces = geometry.face_table(self.shape)
        self.V0_brep = geometry.volume(self.shape)
        self.support_faces = sorted({i for s in cfg["supports"] for i in geometry.select_faces(self.faces, s)})
        self.loads = []
        for L in cfg["loads"]:
            ids = geometry.select_faces(self.faces, L["faces"])
            if not ids:
                raise ValueError(f"load selector matched no faces: {L['faces']}")
            self.loads.append(dict(faces=ids, force=np.asarray(L["force_N"], float)))
        if not self.support_faces:
            raise ValueError("support selectors matched no faces")
        self.F = np.sum([L["force"] for L in self.loads], axis=0)
        extra = sorted({i for s in cfg["mounts"]["extra_faces"] for i in geometry.select_faces(self.faces, s)})
        self.forced_mount_faces = sorted(set(self.support_faces) | {i for L in self.loads for i in L["faces"]} | set(extra))
        self.analytic_faces = [f["id"] for f in self.faces if f["type"] in geometry.ANALYTIC]
        self.V, self.T, self.TF = geometry.tessellate(self.shape, cfg["grid"]["tessellation_mm"])
        self.log(f"  {len(self.faces)} faces ({len(self.analytic_faces)} analytic), V0 = {self.V0_brep:.1f} mm3 "
                 f"({self.V0_brep * cfg['material']['density_g_cm3'] / 1000:.2f} g); supports on faces {self.support_faces}; "
                 f"loads on {[L['faces'] for L in self.loads]}, F = {self.F.tolist()} N")
        self._t("geometry", t0)
        t0 = time.time()
        g = cfg["grid"]
        smax = cfg["design"]["max_offset_mm"]
        pad = max(g["pad_mm"], smax + 2 * g["sdf_mm"])          # the largest offset must fit inside the grid ...
        band = max(g["band_mm"], smax + g["sdf_mm"])             # ... and inside the exact-distance band
        self.grid = field.PartGrid(self.V, self.T, h=g["sdf_mm"], k_fe=g["fe_factor"], h_seg=g["seg_mm"],
                                   band=band, pad=pad, verbose=self.v)
        self.V0_field = self.grid.solid_volume()
        self.grid.prepare_active(cfg["design"]["max_offset_mm"])
        self._t("sdf", t0)
        t0 = time.time()
        self.seg = segment.segment(self.grid, self.V, self.T, self.TF, self.analytic_faces, self.forced_mount_faces, cfg)
        info = self.seg.info
        self.log(f"  segmentation: {info['n_joint']} joints, {info['n_strut']} struts, {info['n_mount']} mounts "
                 f"({time.time() - t0:.0f} s)")
        self._t("segmentation", t0)
        n = len(self.seg.regions)
        self.rin = np.zeros(n + 1, np.float32)
        for r in self.seg.regions:
            self.rin[r["id"]] = r["rin_mm"]
        self.mount_ids = set(self.seg.ids("mount"))
        self.candidates = [r["id"] for r in self.seg.regions if r["kind"] != "mount"]
        from scipy import ndimage as _ndi
        self.mount_mask_c = np.isin(self.seg.label_ext_c, list(self.mount_ids))
        self.mount_zone_c = (_ndi.distance_transform_edt(~self.mount_mask_c) * self.grid.h_seg
                             <= cfg["design"]["transition_mm"]) if self.mount_ids else None

    # ------------------------------------------------------------------------------------------
    def build_design(self, solid_ids, clean_voids=True):
        """Fields, offset s and psi (active subset) for a set of solid candidate regions."""
        cfg = self.cfg
        solid = set(solid_ids) | self.mount_ids
        tab = np.zeros(len(self.seg.regions) + 1, np.float32)
        for i in solid:
            tab[i] = 1.0
        hollow = [i for i in self.candidates if i not in solid]
        if not hollow:
            return dict(solid=sorted(solid), hollow=[], s=0.0, psi_a=None, voids=[], filled_mm3=0.0, feasible=True,
                        original=True)
        theta_a, rin_a = self.grid.design_fields_active(self.seg.label_ext_c, tab, self.rin, cfg["design"]["transition_mm"],
                                                        self.mount_zone_c, self.mount_mask_c if self.mount_ids else None)
        s, nev = self.grid.solve_offset_active(theta_a, rin_a, self.V0_field)
        out = dict(solid=sorted(solid), hollow=hollow, s=s, n_volume_evals=nev, voids=[], filled_mm3=0.0,
                   original=False, theta_a=theta_a, rin_a=rin_a)
        if not np.isfinite(s):
            out.update(psi_a=None, feasible=False, reason="no offset restores the mass")
            return out
        psi_a = self.grid.psi_active(theta_a, rin_a, s)
        forced_a = None
        if clean_voids:
            for _ in range(2):
                full = self.grid.full_psi(psi_a)
                fill, voids, _ = self.grid.small_voids(full, cfg["manufacturing"]["void_min_mm3"])
                if not fill.any():
                    break
                fa = fill.ravel()[self.grid.active]
                forced_a = fa if forced_a is None else (forced_a | fa)
                out["filled_mm3"] = float(forced_a.sum() * self.grid.h ** 3)
                s, _ = self.grid.solve_offset_active(theta_a, rin_a, self.V0_field, forced_a=forced_a)
                if not np.isfinite(s):
                    break
                psi_a = self.grid.psi_active(theta_a, rin_a, s, forced_a)
            full = self.grid.full_psi(psi_a)
            _, voids, _ = self.grid.small_voids(full, 0.0)
            out["voids"] = voids
            del full
        out.update(s=s, psi_a=psi_a, forced_a=forced_a,
                   feasible=bool(np.isfinite(s) and s >= cfg["manufacturing"]["t_min_mm"]))
        if not out["feasible"]:
            out["reason"] = f"wall {s:.3f} mm < t_min {cfg['manufacturing']['t_min_mm']} mm"
        return out

    def psi_full(self, design):
        return self.grid.phi.copy() if design.get("original") else self.grid.full_psi(design["psi_a"])

    # ------------------------------------------------------------------------------------------
    def analyse(self, design, label):
        """FE + metrics of one design."""
        cfg = self.cfg
        t0 = time.time()
        psi = self.psi_full(design)
        vf, origin, hfe = self.grid.fe_fractions(psi)
        volume = self.grid.volume_of(psi)
        del psi
        mesh = fe.VoxelMesh(vf, origin, hfe, vf_min=cfg["fe"]["vf_min"])
        del vf
        import igl
        tol2 = (0.75 * hfe) ** 2
        Vf = np.ascontiguousarray(self.V, np.float64)

        def nodes_near(face_ids):
            T = np.ascontiguousarray(self.T[np.isin(self.TF, face_ids)], np.int64)
            d2 = igl.point_mesh_squared_distance(np.ascontiguousarray(mesh.X, np.float64), Vf, T)[0]
            return np.where(d2 <= tol2)[0]
        fixed = nodes_near(self.support_faces)
        if len(fixed) == 0:
            raise RuntimeError("no FE nodes found on the support faces; refine grid.sdf_mm")
        load_nodes, f_sets = [], []
        for L in self.loads:
            ln = nodes_near(L["faces"])
            if len(ln) == 0:
                raise RuntimeError(f"no FE nodes found on load faces {L['faces']}")
            load_nodes.append(ln)
            f_sets.append(L["force"])
        # several loads: solve with the summed nodal force vector
        all_load = np.unique(np.concatenate(load_nodes))
        if len(self.loads) == 1:
            sol = fe.solve(mesh, cfg["material"]["E_MPa"], cfg["material"]["nu"], fixed, load_nodes[0], f_sets[0],
                           tol=cfg["fe"]["tol"], maxiter=cfg["fe"]["maxiter"], verbose=self.v)
        else:
            sol = self._solve_multi(mesh, fixed, load_nodes, f_sets)
        cell_region = self.grid.fe_cell_labels(self.seg.label_ext_c, mesh.cell_ijk)
        met = mechanics.evaluate(mesh, sol, cell_region, self.seg, cfg, self.F, mesh.X[fixed], mesh.X[all_load], design)
        met.update(volume_field=volume, mass_g=volume * cfg["material"]["density_g_cm3"] / 1000,
                   n_cells=mesh.n_cells, ndof=sol["ndof"], cg_iterations=sol["iterations"], cg_rel_residual=sol["rel_residual"],
                   converged=sol["converged"], dropped_cells=mesh.dropped_cells, n_fixed_nodes=len(fixed),
                   n_load_nodes=len(all_load), fe_time=time.time() - t0, label=label)
        self._t("fe", t0)
        return met

    def _solve_multi(self, mesh, fixed, load_nodes, f_sets):
        """Superpose several load patches by building one nodal force vector."""
        import pyamg
        cfg = self.cfg
        K = fe.assemble(mesh, cfg["material"]["E_MPa"], cfg["material"]["nu"])
        ndof = K.shape[0]
        f = np.zeros(ndof)
        for ln, F in zip(load_nodes, f_sets):
            for d in range(3):
                f[3 * ln + d] += F[d] / len(ln)
        fx = np.zeros(ndof, bool)
        for d in range(3):
            fx[3 * fixed + d] = True
        free = np.where(~fx)[0]
        Kff = K[free][:, free].tocsr()
        del K
        ml = pyamg.smoothed_aggregation_solver(Kff, B=fe.rigid_body_modes(mesh.X)[free], symmetry="hermitian", max_coarse=1000)
        res = []
        uf = ml.solve(f[free], tol=cfg["fe"]["tol"], accel="cg", maxiter=cfg["fe"]["maxiter"], residuals=res)
        u = np.zeros(ndof)
        u[free] = uf
        rel = res[-1] / res[0] if res and res[0] > 0 else 0.0
        return dict(u=u.reshape(-1, 3), compliance=float(f @ u), iterations=len(res) - 1, rel_residual=float(rel),
                    ndof=ndof, converged=bool(rel <= cfg["fe"]["tol"] * 10))

    # ------------------------------------------------------------------------------------------
    def objective(self, met):
        o = self.cfg["objective"]
        rK = self.base["compliance"] / met["compliance"]      # = K/K0 for the same load set (robust if loads cancel)
        rP = met["lam"] / self.base["lam"]                    # = P/P0
        J = o["w_stiffness"] * np.log(rK) + o["w_capacity"] * np.log(rP)
        J /= (o["w_stiffness"] + o["w_capacity"])
        ok = (J > np.log(1 + o["min_gain"])) and rK >= o["floor_stiffness"] and rP >= o["floor_capacity"]
        return float(rK), float(rP), float(J), bool(ok)

    @staticmethod
    def _sel(v):
        """Selection order: acceptable designs first, then the highest J."""
        return (bool(v[0]["acceptable"]), float(v[0]["J"]))

    def run(self):
        t_all = time.time()
        cfg = self.cfg
        open(os.path.join(self.out, "run_log.txt"), "w", encoding="utf-8").close()
        self.prov = util.provenance(cfg["step_file"], self.config_file)
        util.write_json(os.path.join(self.out, "config_used.json"), cfg)
        self.setup()
        # ---- baseline --------------------------------------------------------------------------
        self.log("  baseline: FE of the original solid")
        base_design = dict(solid=sorted(self.mount_ids | set(self.candidates)), hollow=[], s=0.0, original=True,
                           feasible=True, voids=[], filled_mm3=0.0)
        self.base = self.analyse(base_design, "solid")
        if not self.base["converged"] or not np.isfinite(self.base["compliance"]):
            raise RuntimeError("baseline FE did not converge; check supports (is the part fully restrained?)")
        self.log(f"  baseline: K0 = {self.base['K']:.2f} N/mm, P0 = {self.base['P']:.1f} N ({self.base['mode']}), "
                 f"mass {self.base['mass_g']:.2f} g")
        Rmap = {r["id"]: r for r in self.base["per_region"]}
        for r in self.seg.regions:
            pr = Rmap.get(r["id"], {})
            r["R"] = pr.get("R")
            r["U_share"] = pr.get("U", 0.0) / max(sum(x.get("U", 0.0) for x in self.base["per_region"]), 1e-30)
        # ---- rank -------------------------------------------------------------------------------
        if cfg["search"]["rank_by"] == "R":
            key = lambda i: -(self.seg.by_id()[i].get("R") or 0.0)
        else:
            key = lambda i: (self.seg.by_id()[i].get("lam") or 0.0)
        self.order = sorted(self.candidates, key=key)
        Nc = len(self.order)
        # ---- mass sweep -------------------------------------------------------------------------
        t0 = time.time()
        self.sweep = []
        for k in range(Nc):
            d = self.build_design(self.order[:k], clean_voids=False)
            self.sweep.append(dict(k=k, n_hollow=Nc - k, s=d["s"], feasible=d["feasible"],
                                   hollow_vol_share=sum(self.seg.by_id()[i]["vol_mm3"] for i in d["hollow"]) / self.V0_field))
        self.sweep.append(dict(k=Nc, n_hollow=0, s=0.0, feasible=True, hollow_vol_share=0.0))
        self._t("mass_sweep", t0)
        feas = [r["k"] for r in self.sweep[:-1] if r["feasible"]]
        self.log(f"  mass sweep over k = 0..{Nc}: {len(feas)} feasible hollow designs "
                 f"(s from {self.sweep[0]['s']:.3f} mm at k = 0); {time.time() - t0:.0f} s")
        # ---- search -----------------------------------------------------------------------------
        budget = int(cfg["search"]["max_fe_evaluations"])
        evaluated = {}

        def run_k(k):
            if k in evaluated or budget - len(evaluated) <= 0:
                return
            d = self.build_design(self.order[:k], clean_voids=True)
            if not d["feasible"]:
                evaluated[k] = None
                self.log(f"  k = {k}: infeasible after void clean-up ({d.get('reason')})")
                return
            label = f"k{k:02d}"
            try:
                met = self.analyse(d, label)
            except Exception as e:
                self.log(f"  k = {k}: FE failed: {e}")
                self.notes.append(f"k={k}: FE failed: {traceback.format_exc()}")
                evaluated[k] = None
                return
            if not met["converged"] or not np.isfinite(met["compliance"]) or met["compliance"] <= 0:
                self.log(f"  k = {k}: FE did not converge (rel. residual {met['cg_rel_residual']:.1e}); design skipped")
                self.notes.append(f"k={k}: FE not converged")
                evaluated[k] = None
                return
            rK, rP, J, ok = self.objective(met)
            if not np.isfinite(J):
                self.log(f"  k = {k}: non-finite objective; design skipped")
                evaluated[k] = None
                return
            row = dict(k=k, n_solid=k, n_hollow=len(d["hollow"]), s=d["s"], filled_mm3=d["filled_mm3"],
                       n_voids=len(d["voids"]), mass_g=met["mass_g"], K=met["K"], P=met["P"], K_ratio=rK, P_ratio=rP,
                       J=J, acceptable=ok, lam_s=met["lam_s"], lam_E=met["lam_E"], lam_L=met["lam_L"], mode=met["mode"],
                       governing_region=met["governing_region"], sigma_q=met["sigma_q"], fe_time=met["fe_time"],
                       ndof=met["ndof"], cg_iterations=met["cg_iterations"], converged=met["converged"],
                       solid_regions=d["solid"], hollow_regions=d["hollow"])
            self.evaluations.append(row)
            # keep the large arrays only for the best design so far
            prev = [v for v in evaluated.values() if v is not None]
            best_prev = max(prev, key=self._sel) if prev else None
            if best_prev is not None and (ok, J) > self._sel(best_prev):
                for key_ in ("theta_a", "rin_a", "psi_a", "forced_a"):
                    best_prev[1][key_] = None
                best_prev[2]["cell"] = None
            elif best_prev is not None:
                for key_ in ("theta_a", "rin_a", "psi_a", "forced_a"):
                    d[key_] = None
                met["cell"] = None
            evaluated[k] = (row, d, met)
            self.log(f"  k = {k:2d}: s = {d['s']:.3f} mm, K/K0 = {rK:.3f}, P/P0 = {rP:.3f} ({met['mode']}), "
                     f"J = {J:+.4f}{'  acceptable' if ok else ''}")

        if feas:
            if len(feas) <= budget:
                plan = feas
            else:
                m = max(budget - 3, 2)
                plan = sorted({feas[int(round(i))] for i in np.linspace(0, len(feas) - 1, m)})
            for k in plan:
                run_k(k)
            # local refinement around the best design
            while budget - len(evaluated) > 0:
                done = [k for k, v in evaluated.items() if v is not None]
                if not done:
                    break
                best_k = max(done, key=lambda k: self._sel(evaluated[k]))
                pos = feas.index(best_k)
                nbrs = [feas[p] for p in (pos - 1, pos + 1) if 0 <= p < len(feas) and feas[p] not in evaluated]
                if not nbrs:
                    # widen: nearest unevaluated feasible k on either side
                    left = [k for k in feas if k < best_k and k not in evaluated]
                    right = [k for k in feas if k > best_k and k not in evaluated]
                    nbrs = ([left[-1]] if left else []) + ([right[0]] if right else [])
                if not nbrs:
                    break
                for k in nbrs:
                    run_k(k)
        else:
            self.log("  no hollow design satisfies the wall-thickness limit at constant mass")
        # ---- decide -----------------------------------------------------------------------------
        cands = [v for v in evaluated.values() if v is not None]
        best = max(cands, key=self._sel) if cands else None
        accepted = best is not None and best[0]["acceptable"]
        self.decision = dict(accepted=bool(accepted), best_k=best[0]["k"] if best else None,
                             best_J=best[0]["J"] if best else None,
                             reason=self._reason(best, accepted))
        self.log(f"  decision: {'HOLLOW design k = %d accepted' % best[0]['k'] if accepted else 'keep the original SOLID part'}"
                 f" -- {self.decision['reason']}")
        # ---- outputs ----------------------------------------------------------------------------
        t0 = time.time()
        self.write_outputs(best, accepted)
        self._t("outputs", t0)
        self.timings["total"] = time.time() - t_all
        self.write_info(best, accepted)
        if cfg["report"]["enabled"]:
            try:
                from . import report
                t0 = time.time()
                report.write_report(self.out, dpi=cfg["report"]["dpi"], verbose=self.v)
                self.timings["report"] = time.time() - t0
                self.write_info(best, accepted)
            except Exception:
                self.log("  report failed:\n" + traceback.format_exc())
        self.log(f"[vshell] done in {time.time() - t_all:.0f} s -> {self.out}")
        return self.decision

    def _reason(self, best, accepted):
        o = self.cfg["objective"]
        if best is None:
            return "no feasible hollow design could be analysed"
        r = best[0]
        if accepted:
            return (f"best hollow design changes stiffness by {100 * (r['K_ratio'] - 1):+.1f}% and load capacity by "
                    f"{100 * (r['P_ratio'] - 1):+.1f}% at constant mass (J = {r['J']:+.4f})")
        why = []
        if r["J"] <= np.log(1 + o["min_gain"]):
            why.append(f"combined gain J = {r['J']:+.4f} is not above ln(1 + {o['min_gain']})")
        if r["K_ratio"] < o["floor_stiffness"]:
            why.append(f"stiffness would fall to {r['K_ratio']:.3f} of the solid's")
        if r["P_ratio"] < o["floor_capacity"]:
            why.append(f"load capacity would fall to {r['P_ratio']:.3f} of the solid's")
        return "shelling does not improve the part: " + "; ".join(why)

    # ------------------------------------------------------------------------------------------
    def write_outputs(self, best, accepted):
        import trimesh
        out = self.out
        orig = trimesh.Trimesh(self.V, self.T, process=False)
        orig.export(os.path.join(out, "original.stl"))
        arrays = {}
        if best is not None:
            row, d, met = best
            psi = self.psi_full(d)
            m = self.grid.export_mesh(psi, os.path.join(out, "best_candidate.stl"))
            self.best_mesh_info = dict(volume=float(m.volume), area=float(m.area), n_faces=int(len(m.faces)),
                                       watertight=bool(m.is_watertight), bodies=int(len(m.split(only_watertight=False))))
            self.slice_centre = self._slice_centre(d)
            arrays.update(self._slices(psi, "best"))
            arrays.update(self._theta_slices(d))
            self.escape = self._escape_holes(psi, d, met)
            del psi
            for k_, v_ in (met["cell"] or {}).items():
                arrays[f"best_cell_{k_}"] = v_
        else:
            self.best_mesh_info, self.escape = None, []
        final = os.path.join(out, "final_design.stl")
        if accepted:
            shutil.copyfile(os.path.join(out, "best_candidate.stl"), final)
        else:
            shutil.copyfile(os.path.join(out, "original.stl"), final)
        if not hasattr(self, "slice_centre"):
            self.slice_centre = self._slice_centre(None)
        arrays.update(self._slices(self.grid.phi, "orig"))
        for k_, v_ in self.base["cell"].items():
            arrays[f"base_cell_{k_}"] = v_
        # previews for the report
        from OCP.BRepTools import BRepTools
        BRepTools.Clean_s(self.shape)                    # drop the fine triangulation so the preview is coarse
        Vp, Tp, Fp = geometry.tessellate(self.shape, 0.3)
        arrays.update(prev_V=Vp.astype(np.float32), prev_T=Tp.astype(np.int32), prev_F=Fp.astype(np.int32))
        idx = np.argwhere(self.seg.label_c > 0)
        arrays.update(seg_xyz=(self.grid.lo + self.grid.h_seg * idx).astype(np.float32),
                      seg_label=self.seg.label_c[tuple(idx.T)].astype(np.int16),
                      skel_pts=self.seg.skel_pts.astype(np.float32), skel_region=self.seg.skel_region.astype(np.int16))
        np.savez_compressed(os.path.join(out, "arrays.npz"), **arrays)

    def _slice_centre(self, d):
        """Centroid of the largest hollow region of the best design, else of the largest joint."""
        R = self.seg.by_id()
        hol = [R[i] for i in (d or {}).get("hollow", [])] if d else []
        if hol:
            return np.array(max(hol, key=lambda r: r["vol_mm3"])["centroid"])
        joints = [r for r in self.seg.regions if r["kind"] == "joint"]
        if joints:
            return np.array(max(joints, key=lambda r: r["vol_mm3"])["centroid"])
        return np.array(self.V.mean(0))

    def _slices(self, psi, tag):
        """Three orthogonal slices through self.slice_centre."""
        g = self.grid
        cen = self.slice_centre
        idx = np.clip(np.round((cen - g.lo) / g.h).astype(int), 0, np.array(g.shape) - 1)
        out = {f"{tag}_slice_x": psi[idx[0], :, :].astype(np.float16), f"{tag}_slice_y": psi[:, idx[1], :].astype(np.float16),
               f"{tag}_slice_z": psi[:, :, idx[2]].astype(np.float16), "slice_centre": cen, "grid_lo": g.lo,
               "grid_h": np.array([g.h])}
        return out

    def _theta_slices(self, d):
        if d.get("original") or d.get("theta_a") is None:
            return {}
        g = self.grid
        th = np.zeros(int(np.prod(g.shape)), np.float32)
        th[g.active] = d["theta_a"]
        th = th.reshape(g.shape)
        cen = self.slice_centre
        idx = np.clip(np.round((cen - g.lo) / g.h).astype(int), 0, np.array(g.shape) - 1)
        return {"theta_slice_x": th[idx[0]].astype(np.float16), "theta_slice_y": th[:, idx[1]].astype(np.float16),
                "theta_slice_z": th[:, :, idx[2]].astype(np.float16)}

    def _escape_holes(self, psi, d, met):
        """Suggested escape-hole positions: per enclosed void, wall points where the FE stress is lowest."""
        from scipy import ndimage
        from scipy.spatial import cKDTree
        if d.get("original") or not d.get("voids"):
            return []
        man = self.cfg["manufacturing"]
        g = self.grid
        void = psi > 0
        lab, _ = ndimage.label(void)
        tree = cKDTree(met["cell"]["xyz"])
        out = []
        for v in d["voids"]:
            m = lab == v["label"]
            wall = m & ndimage.binary_dilation(~void, iterations=1)
            pts = np.argwhere(wall)
            if len(pts) == 0:
                continue
            if len(pts) > 20000:
                pts = pts[np.random.default_rng(0).choice(len(pts), 20000, replace=False)]
            xyz = g.lo + g.h * pts
            # outward direction: gradient of phi
            P = np.clip(pts, 1, np.array(g.shape) - 2)
            grad = np.stack([g.phi[tuple((P + e).T)] - g.phi[tuple((P - e).T)] for e in np.eye(3, dtype=int)], -1)
            dist_out = -g.phi[tuple(pts.T)] + d["s"]       # material between void wall and the outside
            _, j = tree.query(xyz)
            vm = met["cell"]["vm"][j]
            score = vm + 1e3 * (dist_out > d["s"] + 2 * g.h)   # prefer plain wall (thickness ~ s)
            chosen = []
            for i in np.argsort(score):
                if all(np.linalg.norm(xyz[i] - xyz[c]) > 15.0 for c in chosen):
                    chosen.append(i)
                if len(chosen) >= man["escape_holes_per_void"]:
                    break
            for i in chosen:
                n = grad[i] / (np.linalg.norm(grad[i]) + 1e-12)
                out.append(dict(void=v["label"], void_volume_mm3=v["volume_mm3"], x=float(xyz[i, 0]), y=float(xyz[i, 1]),
                                z=float(xyz[i, 2]), nx=float(n[0]), ny=float(n[1]), nz=float(n[2]), vm_MPa=float(vm[i]),
                                diameter_mm=man["escape_hole_diameter_mm"]))
        return out

    # ------------------------------------------------------------------------------------------
    def write_info(self, best, accepted):
        out = self.out
        base = {k: v for k, v in self.base.items() if k not in ("cell", "per_region")}
        regions = []
        best_hollow = set(best[1]["hollow"]) if best else set()
        for r in self.seg.regions:
            row = {k: v for k, v in r.items() if k not in ("tangent", "centroid")}
            row["rank"] = self.order.index(r["id"]) + 1 if r["id"] in self.order else None
            row["best_candidate"] = "hollow" if r["id"] in best_hollow else "solid"
            row["final"] = ("hollow" if (accepted and r["id"] in best_hollow) else "solid")
            row["cx"], row["cy"], row["cz"] = r["centroid"]
            regions.append(row)
        util.write_csv(os.path.join(out, "regions.csv"), regions)
        util.write_csv(os.path.join(out, "evaluations.csv"), self.evaluations)
        util.write_csv(os.path.join(out, "mass_sweep.csv"), self.sweep)
        util.write_csv(os.path.join(out, "escape_holes.csv"), self.escape or [])
        best_info = None
        if best:
            row, d, met = best
            best_info = {k: v for k, v in met.items() if k not in ("cell",)}
            best_info["per_region"] = met["per_region"]
            best_info.update(s=d["s"], voids=d["voids"], filled_mm3=d["filled_mm3"], solid=d["solid"], hollow=d["hollow"],
                             mesh=self.best_mesh_info, K_ratio=row["K_ratio"], P_ratio=row["P_ratio"], J=row["J"])
        info = dict(name=self.cfg["name"], provenance=self.prov, config=self.cfg,
                    geometry=dict(V0_brep_mm3=self.V0_brep, V0_field_mm3=self.V0_field,
                                  mass_g=self.V0_brep * self.cfg["material"]["density_g_cm3"] / 1000,
                                  n_faces=len(self.faces), n_analytic=len(self.analytic_faces),
                                  support_faces=self.support_faces, load_faces=[L["faces"] for L in self.loads],
                                  force_N=self.F, mount_faces_forced=self.forced_mount_faces,
                                  n_triangles=int(len(self.T)), bbox=[self.V.min(0), self.V.max(0)]),
                    faces=self.faces,
                    grid=dict(sdf_mm=self.grid.h, shape=self.grid.shape, fe_mm=self.grid.h * self.grid.k,
                              seg_mm=self.grid.h_seg, n_exact=self.grid.n_band, n_active=len(self.grid.active)),
                    segmentation=self.seg.info, rank_order=self.order, baseline=base,
                    baseline_per_region=self.base["per_region"], best=best_info, decision=self.decision,
                    escape_holes=self.escape, timings=self.timings, notes=self.notes)
        util.write_json(os.path.join(out, "run_info.json"), info)


def run(config_path_or_dict, overrides=None):
    cfg = config_mod.load(config_path_or_dict, overrides)
    r = Run(cfg, config_path_or_dict if isinstance(config_path_or_dict, str) else None)
    return r.run()
