> Renamed 2026-10-31: vshell → dshell (discrete optimiser, `dshell/`), continuous/ → docs/cshell/, shell_lightweighting/ → ushell/.

# PLAN: body-fitted validation fixes (agreed with Owen 2026-10-29)

Worktree `/home/b6bl/owenpeckham.b6bl/shell-lightweighting/.claude/worktrees/cshell-stage2` (= REPO), branch
`cshell-stage2`. Do not commit, switch branch or push. Common rules as in `docs/cshell/plans/PLAN_parallel_modules.md`;
env `source cshell/cluster/env.sh` (submit from REPO; absolute test paths). **Nothing runs on the login node**
(no tests, imports or `python -c`); everything via sbatch. Minimum GPU **and host** memory per job, sized from a small
timed run; report peak GPU and host RSS. Runs under `runs_cshell/VF_*`.

Background: `docs/cshell/research/b2_bodyfit/diagnosis.md` (read first), draft diffs `export.diff`, `validate.diff`,
`collect.diff` (untested), GPU body-fitted prototype `bf_gpu.py`, mesh repair `mesh2.py` (if present) and the job logs
there. The B2 run dir `runs_cshell/C_bench/B2_slender` is the test case.

Owen's decisions:
1. **GPU body-fitted solver is the default** for the §8.6 check: C3D10 + cuDSS (from `bf_gpu.py`), same BCs and load
   model as the ccx deck (clamped support-face nodes, area-weighted traction on load faces), compliance, buckling
   (first n modes) and the 0.995 stress quantile with the 2 mm BC exclusion. One GPU; multi-GPU is allowed for
   validation solves only if one does not fit (log it). CalculiX stays available (`validation.engine`: "gpu" | "ccx")
   and is used as a cross-check in a test.
2. **Mesh fixes + quality gate:** gmsh remesh of the exported surface (with the repair/projection step) and volume
   meshing for solid and design alike, **no tetgen fallback**; gate: ≥ 2 tets through the thinnest wall, no tets below
   a minimum dihedral angle (choose and justify, e.g. 5°), max edge bounded; two mesh levels (report both, use the
   finer, flag non-convergence > 2 % in C). A failed gate is a logged validation failure that blocks acceptance.
3. **Relative ratio tolerance:** |r_bf / r_vox − 1| ≤ 0.05 for K/K0 and P/P0 (§8.6); per-quantity tolerances unchanged.
4. **Escape-hole film fix (export.py):** cut length chosen so both ends of each hole clear the walls (no sub-t_min films
   at hole ends); wall-thickness statistics count only true opposite-wall hits. Test: no wall < t_min − b_v in the
   B2 export apart from flagged hole rims, and holes still connect each void to the exterior.
5. **K_σ exclusion near clamps (body-fitted side):** elements within `capacity.exclude_bc_mm` (2 mm) of clamped nodes
   carry no geometric stiffness in the body-fitted buckling (the voxel side is done by the orchestrator in
   `fe_sparse.buckling` with the same rule). Report λ with and without the exclusion on B2.

Files you own: `cshell/validate.py`, `cshell/export.py`, `cshell/collect.py` (validation step and tolerance only),
new `cshell/bodyfit_gpu.py`, `tests/cshell/test_validate.py`, `tests/cshell/test_export.py`, `tests/cshell/test_collect.py`.
Do **not** touch `cshell/fe_sparse.py`, `cshell/fe_twogrid.py`, `cshell/optimise.py`, `cshell/objective.py`,
`cshell/config.py` (another developer is building the two-grid engine; list any config keys you need and read them
with defaults).

Acceptance: tests pass (CPU and GPU jobs); GPU body-fitted vs ccx on one B2 deck to ≤ 1e-5; re-run `collect` on
`runs_cshell/C_bench/B2_slender` (copy it to `runs_cshell/VF_B2` first) with the new chain: voxel vs body-fitted table
(C, λ_b with and without the exclusion, σ_q, ratios), mesh-gate numbers, time per model, decision. Update
`docs/cshell/THEORY_continuous.md` §8.3, §8.6 (relative tolerance, GPU engine, K_σ exclusion, quality gate) and
`docs/cshell/VALIDATION.md` (B2 body-fitted row).

Slots: 3 (check `squeue -u $USER -h | wc -l` < 8 before each submission). Return under 250 words.
