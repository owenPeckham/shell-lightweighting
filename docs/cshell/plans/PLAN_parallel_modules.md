> Renamed 2026-10-31: vshell → dshell (discrete optimiser, `dshell/`), continuous/ → docs/cshell/, shell_lightweighting/ → ushell/.

# PLAN: cshell modules built in parallel with the FE engine (agreed plan: docs/cshell/STATUS.md §2.3)

Repo `/home/b6bl/owenpeckham.b6bl/shell-lightweighting`, branch `cshell` (already checked out; **do not commit, do not switch branch, do not push** — the orchestrator commits). Spec: `docs/cshell/THEORY_continuous.md` (cite § in docstrings). Owen's decision D1 (STATUS.md §4): the inset lower bound is a ≥ (t_min − d)⁺ smoothed; it affects `design.py`, not these modules.

Common rules
- Env: `source cshell/cluster/env.sh` (conda env at `/projects/b6bl/owenpeckham.b6bl/home_data/shell-lightweighting/envs/cshell`: numpy, scipy, numba-cuda, cupy, warp-lang, libigl, trimesh, scikit-image, mmapy, build123d/OCP, pytest, pyvista, gmsh). Packages must be aarch64; if one is missing, `pip install` it into that env and add it to `docs/cshell/environment.yml`.
- Login node: edit only. Run tests via `sbatch cshell/cluster/cpu_test.sh <pytest args>` or `sbatch cshell/cluster/gpu_test.sh pytest <args>` / `gpu_test.sh script.py`. Logs go to `logs/`. Respect your slot budget.
- Units mm, N, MPa. Type hints; docstrings give the equation and §. Float64 for reductions. Match the style of `cshell/fea.py`.
- Reuse vshell (`dshell/geometry.py`: `load_step`, `face_table`, `select_faces`, `tessellate` → (V, T, TF); `dshell/segment.py` for mount detection; `dshell/field.py` for SDF ideas). Do not change vshell; if you must, run `python tests/dshell/test_basic.py`.
- Tests in `tests/cshell/test_<module>.py`, CPU by default, GPU tests `@pytest.mark.gpu` (fixtures in `tests/cshell/conftest.py`). Only touch the files listed for your task.

## Task A: `cshell/surface.py` (§3.4, §3.5)
- `DesignSurface` built from (V, T, TF): isotropic remesh to target edge ℓ_Γ (default 0.75 mm). gpytoolbox has no aarch64 wheel: use pymeshlab if it installs on aarch64, else implement Botsch–Kobbelt (split/collapse/flip/tangential relax with back-projection onto the input mesh). Check Hausdorff distance to the input ≤ 0.02 mm (assert/report). Keep B-rep face id per vertex (nearest input triangle). Preserve sharp feature edges between B-rep faces if feasible (needed for the Hausdorff bound).
- Lumped (Voronoi or barycentric) mass `M` (diag), cotangent `L` (PSD sign), vertex areas A_i, normals n_i (area-weighted, outward).
- `HelmholtzFilter(M, L, r_f)`: one sparse factorisation; `apply(x)` = (M + r_f² L)⁻¹ M x, `adjoint(g)` = M (M + r_f² L)⁻¹ g, `invert(x̃)` = M⁻¹(M + r_f² L) x̃ (§3.5 step 2). Tests V1-d: maximum principle (outputs in [min x, max x] for x ∈ [0,1]), xᵀM F y = yᵀ M F x to 1e-12, adjoint dot-product test.
- `medial_depth(...)`: r_MA per vertex by marching inward in h_g/2 steps with **exact** point-to-mesh distance (igl, or warp `mesh_query_point`) until D(t) < t − δ (δ = 1e-3), then bisection to 0.01 mm (§3.5). Ray-chord upper bound as a sanity check (report violations). Tests V1-b: cylinder r = 5 (and a 10 × 20 box: inradius 5) within ±0.1 h_g, h_g = 0.1.
- `outward_clearance(...)`: r_out = g/2 from a ray along +n hitting the mesh (∞ if none). `keepout_distance(...)` from keep-out SDFs (primitives sphere/box/cylinder + STL), bound by first point with φ_K < c_K. Tests on two parallel cylinders and a slab with a sphere keep-out.
- `passive_vertices(...)`: mount zone = within `dilate_mm` of mount/BC faces (analytic faces per `vshell.segment` logic + support/load faces + extra faces; reuse vshell's code paths, keep only the mount part); and thin zone a^{hi} ≤ a^{lo} + v_min/2 — expose the thin rule as a function of given arrays so design.py can apply the D1 variant.
- Performance target: cantilever (`examples/parts/Solid Cantilever.step`, ≈ 6·10⁴ vertices) end to end < 5 min on one CPU node; report timings.

## Task B: `cshell/mma.py` (§7.1, §7.2) and B4
- Wrapper around `mmapy` (`mmasub`; GCMMA optional) with state object (xold1, xold2, low, upp, iteration), move limit, asyinit/asyincr/asydecr 0.5/1.2/0.7, a0=1, a_i=0, c_i=1e3, d_i=1, constraint scaling, and `reset_asymptotes()`; state serialisable to/from a dict of numpy arrays (checkpoints, §7.6), bit-exact resume.
- `restore_mass(xi, volume_fn, grad_fn, V0, active, tol=1e-4)`: ξ ← clip(ξ⁺ + η g/‖g‖∞) with g = ∇V on active entries, Brent on η (§7.2); returns ξ, η, number of volume evaluations; raises if no bracket. Generic: `volume_fn` is any callable (the real one comes later from design.py). Tests V4-7 with a synthetic monotone volume (e.g. V = Σ w_i s(Fξ)_i with a smooth monotone s and positive filter) after a random MMA-like step: |V/V₀ − 1| < 1e-4 in ≤ 20 evaluations.
- V5-B4: reproduce Svanberg's two toy problems from "MMA and GCMMA – two methods for nonlinear optimization" (2007) with their published optima (and the mmapy repository examples); tests assert optimum to the published digits.
- Optimiser comparison hook (Owen's request): `ProjectedLBFGSB` alternative with the same `step()` interface (scipy L-BFGS-B on the box, mass held by `restore_mass`), so B2/B3 can compare later. Test on a toy box-constrained problem.

## Task C: `cshell/band.py` (§3.1, §3.6)
- Sparse band on the bounding grid of spacing h_g: active cells = {−r_MA,max − h_g ≤ φ ≤ D_comp + h_g} (in practice interior + outer shell). Per cell: flat index (int64), φ (float32, mm, negative inside, exact distance signed by winding number), closest-triangle id (int32), barycentrics (3 × float32). GPU via NVIDIA Warp (`wp.Mesh`, `mesh_query_point_sign_winding_number` or `mesh_query_point` + winding); CPU fallback via igl for tests. Find active cells without evaluating the whole box: coarse-to-fine (evaluate at a coarse grid, keep blocks whose |φ| bound intersects the band). Tiled so peak memory stays bounded.
- `Band.P(surface)`: sparse matrix (n_cells × N_v, 3 nnz/row) mapping closest point on the fine tessellation onto the design surface Γ_h (closest triangle of Γ_h + barycentrics). Interface: takes Γ_h vertices/triangles as arrays (Task A builds Γ_h; don't import it — use a test remesh/fixture).
- `Band.extend(D_new)` (§3.5 D_comp doubling): add the new shell cells, keep existing ones.
- Cache to `.npz` keyed by STEP SHA-256 + grid parameters (h_g, D_comp, tessellation tol) in `$CSHELL_CACHE` (default `.cache_cshell/`, a symlink to projects storage — create `/projects/b6bl/owenpeckham.b6bl/home_data/shell-lightweighting/.cache_cshell` and symlink it if absent).
- Volume helper V = h_g³ Σ χ over band (+ interior not in band) in float64; V₀ of the solid with the quintic Heaviside H_ε, ε = h_g (§3.7, §3.10).
- Tests (CPU, small): φ against analytic sphere/cylinder/box SDF (|err| < 1e-4 mm); V1-a volumes of a solid cylinder and a tube via ψ to < 2e-3 at h_g = 0.1 (use primitives meshed finely; tube ψ = max(φ − t_out, −φ − a) style per §3.2 is fine with constant fields); P rows sum to 1 and reproduce a linear field exactly. GPU test: Warp result equals igl on a sample. Benchmark (GPU job): band build for the cantilever at h_g = 0.2 and 0.1 (cells, time, memory); write the numbers into `docs/cshell/research/band_bench.md`.

## Return (each task, under 250 words)
Files written; test results with job ids; measured timings; deviations from the spec and why; open issues.
