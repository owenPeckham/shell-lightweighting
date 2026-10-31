# PLAN: repository restructure (agreed by Owen 2026-10-31)

Goal: restructure the shell-lightweighting repo to an intuitive layout, rename two packages, add an INDEX.md, and
commit the continuous optimiser's (cshell) results for the three example parts, so the repo can be cloned on a laptop
and everything found.

Branch `restructure` (from main 0d60809); commits in logical steps; push `restructure` to origin; never push to main,
never force-push.

## Target layout (`git mv` so history follows)
```
README.md            overview of the three optimisers + pointers (rewrite briefly; keep author/description)
INDEX.md             NEW: map of the repo - every top-level dir and key files, where results/reports are for each
                     optimiser and part, how to run each, where docs are
ushell/              was shell_lightweighting/  (Fusion 360 add-in, uniform shell)
dshell/              was vshell/  (discrete-region non-uniform shell optimiser)
  README.md          was README_vshell.md
  requirements.txt   was requirements_vshell.txt
cshell/              continuous non-uniform shell optimiser (package unchanged)
  README.md          was README_cshell.md
  cluster/           was continuous/cluster/  (Slurm scripts)
examples/
  parts/             was examples/step/  (shared STEP files)
  dshell/            was examples/*.json + examples/runs/  (dshell configs + runs/<part>/ old results)
  cshell/            cshell configs (cantilever/mbb/l_bracket.json), bench/, runs/<part>/ (NEW results)
docs/
  cshell/            was continuous/  (THEORY, STATUS, VALIDATION, BENCHMARKS, HANDOVER, plans/, research/,
                     reference/, PROMPT files)
feasibility/         unchanged
tests/
  dshell/            was tests/*.py (the vshell tests)
  cshell/            unchanged
```

## Renames (everywhere)
- `vshell` -> `dshell`: package dir, `python -m dshell`, all imports (incl. cshell's S2 start which builds its initial
  design from the discrete optimiser), CLI help, configs, tests, READMEs/INDEX, `.gitignore`.
- The cshell start `vshell` (config `optimiser.starts`, `--start vshell`, history_vshell.csv, xi_vshell_phase*.npy,
  checkpoints/vshell_*, done_vshell.json, run_info.d, report/summary labels) becomes `dshell`. `vshell` stays accepted
  as an ALIAS: config/CLI start "vshell" maps to "dshell", and reading an existing run directory (resume, collect,
  report, history merge, checkpoint lookup, done_<start>.json, xi files, summary rows) works with old `vshell` names.
  A test covers the alias on a small fake run dir.
- `shell_lightweighting` -> `ushell` (Fusion add-in folder; `ShellOptimisation.py/.manifest` unchanged; INDEX/README
  say how to install it).
- Historical docs: no history rewrite; one rename note at the top of each top-level doc; path references a reader
  would follow are updated. Living docs use the new names fully.

## Path fixes
All old paths (`git grep`), relative `step_file` paths in example/bench JSON, `Path(__file__).parents[N]` roots in
moved scripts/tests, Slurm scripts (`cshell/cluster/env.sh`), `.gitignore` (`runs_cshell` stays ignored;
`examples/*/runs/` not ignored), configs' `output_dir` keeps resolving to the repo-root `runs_cshell`.

## cshell results -> examples/cshell/runs/<part>/ (MBB, Cantilever, L-bracket)
From `/projects/b6bl/owenpeckham.b6bl/home_data/shell-lightweighting/runs_cshell/FULL/<part>/` (read-only): reports,
summaries, logs, configs, CSVs, histories, npz, plots/, xi_<start>_phase{0,1}.npy, fields.vtp; vshell-named files
renamed to dshell. No checkpoints, no full STLs. Decimated preview STLs (~500k triangles, binary, <= ~30 MB,
volume change <= 0.5 %, watertightness reported) via sbatch CPU jobs. examples/cshell/runs/README.md describes files,
where the full data is, how to regenerate.

## Verify (acceptance)
- `git grep -w vshell` and old-path greps: only historical/alias/rename-note hits.
- Full test suites via sbatch (CPU, GPU in 3 jobs, dshell script tests) all pass (baseline GPU 162, CPU 124).
- `python -m dshell --help`, `python -m cshell --help`, `python -m cshell setup examples/cshell/cantilever.json` in jobs.
- No added file > 50 MB; report `git count-objects -vH` delta.

Rules: nothing runs on the login node; minimum resources; 12 Slurm slots.
