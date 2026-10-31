# Body-fitted check at production scale (2026-10-31)

Scripts (run from the repo root, never on the login node; `gpu_py.sh <script> [args]` is the GPU job wrapper):
- `diag.py <run_dir> solid|design A<h>,B<h>,G<h>,T<h>`: surface options for the remesh + gmsh chain (no solve).
  A = marching cubes (MC) of the SDF at s = h into gmsh directly; B = MC at h/3 → pymeshlab isotropic remesh → gmsh;
  G = gmsh reparametrisation of the MC surface; T = B with the thickness-aware remesh (`remesh_surface(thin=2, fine=s)`).
- `stuff_check.py <run_dir> solid|design [solve]`: isosurface-stuffing mesh at both collect levels, gate, deviation.
- `factor_estimate.py`, `solve_check.py`: cuDSS memory / full-chain solve on a production model.
- `summ.sh <logs>`: one JSON line per diag result. `submit_diag.sh`, `submit_stuff.sh`: the job sets below.

## Root cause
1. The surface step, not the input size: the pymeshlab isotropic remesh to edge h creates self-intersections wherever
   the wall or the gap is thinner than ~2h (knife-edge rims, sub-mm walls: the production designs have 10 % of their
   area there, `thin_faces_in` 0.16-0.76 M of 2.6-7.3 M MC faces); the delete + close-holes repair then breaks
   watertightness, and gmsh finds overlapping facets. Shown with a clean, watertight MC input at h/3 (B, T):
   designs still fail (L-bracket 21-26 repairs, MBB 34-56, Cantilever 50-75; gmsh "overlapping facets" after repair),
   solids pass (no repairs) except the 2.5 h edge criterion by 0.1-0.6 % (2.52-2.62 h).
   A (MC at h, no remesh): min dihedral 0.16°, 0.7 % of tets < 5°, gmsh segfault at h/√2 (B2). G: gmsh
   "Wrong topology of boundary mesh for parametrization" (B2).
2. Behind it, the solver: the cuDSS direct factor does not scale to the production meshes at h and h/√2 (below).

## Mesher: isosurface stuffing (validate.stuff_tets, collect `validation.mesher` = "stuff", default)
BCC lattice of spacing L = c·h (c = √2: tet density of gmsh at MeshSizeMax h), φ = SDF of the model surface (Warp,
winding-number sign), warp nodes onto cut points closer than α|e| (Labelle-Shewchuk α), clip lattice tets by the
linear zero set, split prisms / pyramids by the min-index diagonal rule (conforming). No surface remesh, no Delaunay.
Unit tests: sphere shell (conforming, closed boundary, dihedral > 5°, edge ≤ 1.5 L, O(L²) volume), CPU job 7883510.

Gate (unchanged thresholds) and deviation of the meshed boundary from the exported STL (design) / context surface
(solid), jobs 7883410-14 (B2, designs), 7883586-88 (solids):

| part | model | h | tets | min dihedral ° | frac < 5° | layers p01 | edge/h | V/V_ref−1 | Hausdorff mm | 99.9 % from STL | 99.9 % to STL | mean from STL |
|---|---|---|---|---|---|---|---|---|---|---|---|---|
| MBB | solid | 0.354 | 8.78 M | 9.46 | 0 | 10.9 | 1.77 | −1.0e-3 | 0.261 | 0.199 | 0.117 | 0.0084 |
| MBB | design | 0.354 | 9.45 M | 3.19 | 4e-7 | 3.10 | 2.05 | +9.4e-4 | 1.951 | 0.968 | 0.138 | 0.0190 |
| Cantilever | solid | 0.354 | 6.29 M | 9.75 | 0 | 10.9 | 1.76 | −1.4e-3 | 0.345 | 0.197 | 0.109 | 0.0080 |
| Cantilever | design | 0.354 | 6.79 M | 5.42 | 0 | 3.06 | 1.95 | +2.6e-4 | 1.954 | 0.791 | 0.130 | 0.0160 |
| L-bracket | solid | 0.354 | 4.31 M | 10.4 | 0 | 8.70 | 1.76 | −1.0e-3 | 0.246 | 0.215 | 0.105 | 0.0098 |
| L-bracket | design | 0.354 | 4.47 M | 5.58 | 0 | 3.03 | 1.96 | −1.9e-4 | 1.962 | 0.670 | 0.127 | 0.0138 |
| B2 | solid | 0.354 | 0.65 M | 12.5 | 0 | 21.9 | 1.75 | −1.2e-3 | 0.231 | 0.231 | 0.093 | 0.135* |
| B2 | design | 0.354 | 0.70 M | 8.38 | 0 | 3.13 | 1.83 | −2.1e-4 | 0.288 | 0.132 | 0.086 | 0.0073 |

All pass the gate at both levels (h = 0.5: min dihedral 3.6-13°, layers p01 2.13-2.22 for the designs).
Volume within 0.5 % everywhere (max |ΔV/V| 2.7e-3, solid at h = 0.5). Max deviation ≤ h/4 is NOT met: the designs'
Hausdorff is ~2 mm and 0.1 % of the export vertices lie 0.67-1.23 mm from the meshed boundary — knife-edge rims and
walls thinner than the lattice spacing (0.5 mm at h/√2) are not resolved; solids 0.25-0.48 mm (sharp CAD edges cut
at lattice resolution). (*B2 solid "from STL" is from the 248-face STEP tessellation, vertices only at corners.)
The same holds for any mesh of edge ~h: the MC surface at h/3 already had Hausdorff 0.75-1.2 mm on the designs.

## Solver: blocker
B2 design, h/√2: 3.17 M DOF, cuDSS factor 51 GiB (peak device), solve 237 s (job 7883411). Production fine meshes
are 4.3-9.4 M tets ≈ 19-42 M DOF: nested-dissection fill ~n^(4/3) → ~0.3-1.5 TB factor; even the coarse level
(1.6-3.5 M tets, 7-16 M DOF) is 0.15-0.45 TB. Measured: the GPU assembly of K (`P2Model` pattern: 900 int64 entries
per tet plus a global unique) runs out of the 96 GB device at 4.5 M tets (L-bracket design, h/√2: OOM at 73 GB
allocated, job 7883508); MBB design h = 0.5 (3.5 M tets) OOM at 51 GB (job 7883507); cuDSS hybrid-mode planning on
the L-bracket coarse level: std::bad_alloc (7883508). Hybrid memory mode was added to `bodyfit_gpu.Cholesky`
(automatic when the factor does not fit); it helps only up to the host memory (≤ 460 GB), i.e. not here.
Four GPUs (384 GB) would not hold the fine-level factors either.

Options (need a decision): (a) iterative solver for the body-fitted model (chunked assembly + AMG-preconditioned
CG on the GPU; buckling by block Krylov with inner PCG solves - hours per model, to be verified against cuDSS on B2
to 1e-6); (b) a coarser body-fitted mesh for the production parts (e.g. only the h = 0.5 level, still ~0.15-0.45 TB
direct); (c) solve a reduced region / substructure. None is in PLAN_validation_fixes (cuDSS, one GPU).

## B2 with the stuffing chain (collect job 7883832, copy runs at home_data/shell-lightweighting/bodyfit_scale/CF_B2)
Accepted (J +1.1572). Body-fitted vs the gmsh chain (export_watertight/CF_B2): solid C 78.313 vs 78.165 (+0.19 %),
λ_b 22.318 vs 22.313 (+0.02 %); design C 23.594 vs 23.620 (−0.11 %), λ_b 43.182 vs 43.132 (+0.12 %); K/K0 3.3192 vs
3.3093 (+0.30 %), Pb/Pb0 1.9348 vs 1.9331 (+0.09 %), Ps/Ps0 2.2096 vs 2.2115 (−0.09 %). Level change C −0.12 % / +0.08 %.
λ_b without K_σ exclusion (h/√2): solid 22.285, design 26.630. Wall 11:45, peak host RSS 68 GiB (cuDSS hybrid mode
switched on for the solid at h/√2: 64.9 GiB factor vs 66.8 GiB free; the stuffing mesh has ~12 % more tets than gmsh).
Tests: CPU 7883583 (test_validate + test_collect, all pass), GPU 7883584 (all 39 pass), new-unit subset 7883510.
