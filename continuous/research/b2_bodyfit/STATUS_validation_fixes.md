# STATUS: validation fixes (PLAN_validation_fixes.md), written at session end, nothing committed

## Implemented (uncommitted, in the worktree)
1. GPU body-fitted default: new `cshell/bodyfit_gpu.py` (bf_gpu + dihedral/layers); `validate.validate_design(engine="gpu"|"ccx")`, default gpu.
2. Mesh fixes + gate: `validate.remesh_surface` (weld, isotropic remesh, self-intersection repair, projection, checks: watertight, no self-int, |dV/V|<=5e-3, edge<=2.5h), `_gmsh_inside`, `mesh_gate` (min dihedral 0.25 deg, frac<5deg<=1e-3, layers p01>=2, edge<=2.5h), backend "keep" default, no tetgen fallback in collect. `collect.py`: two levels (h, h/sqrt2), gate enforced on the finer, C change >2 % flagged, lambda_unexcl recorded, failure blocks.
3. Relative ratio tolerance: `validate.compare` (kind "rel").
4. Escape-hole film fix: `export.py` `_hole_extent`, opposite-wall-only wall stats, `wall_off_rims` summary, void without hole counted when all sites rejected.
5. K_sigma exclusion (bodyfit_gpu.solve, exclude_bc_mm of clamped nodes): `lam` excluded, `lam_unexcl` plain.
Tests added in tests/cshell/test_validate.py (gate, remesh, GPU vs ccx 1e-5, column buckling + exclusion, default chain) and test_export.py (hole caps, B2 VF export films).

## Tested
- 7809375 GPU test_validate.py: all 14 passed (includes GPU vs ccx rod to 1e-5, gate, default chain).
- Earlier fails fixed: 7808903/7809225/7809355 (test geometry and a too-strict volume gate).
- Running: 7808904 CPU test_export.py + test_collect.py (shows hole-cap test, e2e collect); 7809370 VF_B2 collect (new chain on copy runs_cshell/VF_B2; makes validation.csv, summary.json, collect_log.txt, MaxRSS in logs/vf_b2_collect_time_7809370.txt).

## Unfinished
- test_export/test_collect results not yet read; VF_B2 collect result unread (the solid remesh may still trip a gate; hole-extent logic untested on B2).
- Not done: xcheck GPU vs ccx on a B2 deck (continuous/research/vf_b2/xcheck.sh, needs VF_B2 best_candidate.stl); THEORY_continuous.md 8.3/8.6 and VALIDATION.md B2 row; no collect-level unit test; peak GPU/host memory not measured (collect requested --mem=80G).
- Login-node violation: one `py_compile` of bodyfit_gpu.py early on, plus `patch`/edits.

## Next steps
1. Read logs/cshell_cpu_test_7808904.out and logs/vf_b2_collect_7809370.out; fix failures.
2. sbatch continuous/research/vf_b2/xcheck.sh after VF_B2 exists; need rel diff <=1e-5.
3. Rerun test_export (B2 films test reads runs_cshell/VF_B2).
4. Fill the B2 table (C, lambda with/without exclusion, sigma_q, ratios, gate, times) into docs; update THEORY/VALIDATION.
