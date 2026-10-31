> Renamed 2026-10-31: vshell → dshell (discrete optimiser, `dshell/`), continuous/ → docs/cshell/, shell_lightweighting/ → ushell/.

# PLAN: export (§8.1–8.5) and body-fitted validation (§8.6), built in parallel (agreed plan: STATUS.md §2.4 item 4)

Repo `/home/b6bl/owenpeckham.b6bl/shell-lightweighting`, branch `cshell` (checked out). **Do not commit, switch branch or
push**; touch only your task's files. Spec: `docs/cshell/THEORY_continuous.md` v1.2 (note §14 D6: inset bound
a ≥ a_lo(d), and §4.2 v1.2 note on the FE engine). Common rules (env, sbatch, tests, style) as in
`docs/cshell/plans/PLAN_parallel_modules.md` "Common rules". Env:
`/projects/b6bl/owenpeckham.b6bl/home_data/shell-lightweighting/envs/cshell` (has trimesh, scikit-image, pyvista,
tetgen, gmsh/python-gmsh, ccx (CalculiX 2.20 binary at `$CSHELL_ENV/bin/ccx`), scikit-fem, warp, igl).

Existing interfaces to use (read the module docstrings, not whole files):
- `cshell/design.py`: `DesignMap.vertex_fields(xi_d, xi_a)` → (ξ̃ᵈ, ξ̃ᵃ, d, a, a_lo, da_lo) per vertex of Γ_h; ψ(x) =
  max(φ(x) − d(π(x)), −φ(x) − a(π(x))) with π the closest point on Γ₀ and d, a interpolated on Γ_h (§3.2, §3.6).
- `cshell/band.py`: `Band.build(V, T, h_g, D_comp)` (φ, closest triangle, barycentrics on a sparse band; Warp GPU or
  igl), `Band.P(Vh, Th)` (cells → Γ_h vertices). `cshell/surface.py`: `DesignSurface` (Γ_h: V, T, A, N, query0).
- `cshell/optimise.py`: `Context(cfg)` (geometry, faces, BCs, Γ_h, r_MA), `PhaseModel(ctx, phase, D_comp)` (band, P,
  DesignMap `dm`, design space `ds`, FE BCs, `fe_eval`). Test configs: `examples/cshell/bench/B{1,2,3}_*.json` (rods,
  r0 = 5 mm), parts `examples/cshell/{cantilever,l_bracket,mbb}.json`.
- `cshell/fea.py`: `build_bcs`, `nodes_near_faces` (vshell's 0.75 h rule).

## Task E: `cshell/export.py` (§8.1–8.5)
Input: ctx, the final ξ, and the stress field (optional, for hole siting). Output files in a run directory.
1. Nominal ψ on an export grid h_x = 0.05 mm (configurable), **tiled** so the full grid is never held (band of the
   export grid only: |φ| region that can be material, i.e. φ ≤ max d + h_x); reuse Band/P machinery per tile or a
   direct closest-point query on Γ₀ then Γ_h per tile.
2. Clean-up (§8.2): vertices with ξ̃ᵃ ≥ 0.98 treated as exactly closed (a = ∞); morphological opening of the void mask
   with a ball of diameter v_min (3 mm) — implement with EDT on the tile set with halo ≥ v_min (scipy.ndimage
   distance transforms), check the rule (re-opening changes ≤ one cell); fill enclosed voids < 30 mm³; keep the largest
   material component; report volumes filled.
3. Escape holes (§8.3): per enclosed void component, candidate wall sites (nominal wall ≥ t_min + b_v, > 3 mm from
   mounts, lowest-30 % stress if a stress field is given, else lowest-30 % depth), choose 2 maximising separation
   along the void's principal axis, cut cylinders D = 3 mm along the local outward normal through the wall
   (ψ ← max(ψ, D/2 − ‖(x − p) × n‖) capped to the wall); write `escape_holes.csv`.
4. Mass correction (§8.4): δ by Brent with d_i ← clip(d_i + δ, 0, d^hi_i) on active vertices so V = V₀ on the export
   grid; warn > 0.1 mm, fail > 0.3 mm.
5. Mesh (§8.5): marching cubes (skimage) on ψ per tile, stitched watertight (or a single sparse-volume MC), trimesh
   checks (watertight, manifold), `final_design.stl`, |V_mesh/V₀ − 1| ≤ 2e-3 (R1); wall-thickness histogram by inward
   ray casting (nominal and as-built = − b_v on void walls); void count and holes per void.
6. Also `original.stl` and `uniform_shell.stl` (ψ = max(φ − t, −φ), t solved for constant mass, §7.4 S1 note).
Tests (`tests/cshell/test_export.py`, CPU, small): sphere/rod with a known analytic shell → volume, wall thickness and
void count; opening rule on a synthetic sliver (sliver filled, ≥ v_min void kept with rounded ends); escape holes cut
and the void connected to the exterior afterwards; mass correction restores V within 2e-3 on the mesh. A GPU/CPU job
exporting the B2 rod from a hollow ξ (ξᵈ = 0.5, ξᵃ = 0) end to end, timing reported. Then the cantilever solid +
uniform shell export timing (h_x = 0.05) to `docs/cshell/research/export/bench.md`.

## Task V: `cshell/validate.py` (§8.6)
1. Tetrahedral mesh of an STL (TetGen via `tetgen` python, or gmsh) with quadratic tets (C3D10), ≥ 2 elements through
   the thinnest wall (size from the wall histogram or a given min wall), mesh-size control near thin walls.
2. CalculiX deck writer and driver (`ccx`): linear static step + `*BUCKLE` (n modes) with the same BCs as cshell —
   all DOF clamped on nodes within 0.75·h (use h = local mesh size, or a tolerance) of the support-face triangles,
   total load shared equally over nodes on the load faces. Faces are given as triangle sets of the original
   tessellation (`ctx.V, ctx.T, ctx.TF`, face ids `ctx.support`, `ctx.loads[i]["faces"]`); mounts never move, so the
   original face triangles locate the BC nodes on any exported design.
3. Parse results (`.dat`/`.frd`): compliance f·u, first buckling factors, von Mises stress at nodes/elements (0.995
   volume-weighted quantile with the same exclusion-zone rule as §4.4(b), 2 mm from BC nodes); write `validation.csv`
   comparing voxel vs body-fitted (tolerances §8.6: compliance ≤ 5 %, first buckling factor ≤ 10 %, stress quantile
   ≤ 15 %, ratios within 5 %).
4. Fallback if ccx fails: scikit-fem P2 tets for compliance (log the fallback; no silent fallback).
Tests (`tests/cshell/test_validate.py`): cantilever rod (r = 5, L = 40, tip shear) vs Timoshenko within 2 %;
clamped–free column buckling vs π²EI/(4L²) within 3 %; deck/parse round-trip. Then **V3-b**: tube R_m = 10, t = 1.5,
L = 60, axial compression, ccx `*BUCKLE` on C3D10 (mesh-converged, ≥ 3 element layers through t), and compare with
cshell's sparse engine (`cshell.fe_sparse.buckling`, build the voxel tube like
`docs/cshell/research/stage0/tubes.py::euler_rod` does for a rod) at h_e = 0.8 and 0.4 mm (0.2 if it fits; the engine
falls back to cuDSS hybrid memory when the factor does not fit — request `--mem` ≥ 200G for that). Record in
`docs/cshell/research/validate/v3b.md`.

## Return (each, under 250 words)
Files written; test results with job ids; measured timings; deviations from the spec and why; open issues.
