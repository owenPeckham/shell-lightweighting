# cshell lightweight results (LW1, 2026-11-09)

Default objective `objective.type = "lightweight"` (`docs/cshell/plans/PLAN_lightweight.md`): minimise V/V0 subject to
K/K0 ≥ 1.02 and P/P0 ≥ 1.02 (MMA), no mass restoration; accepted iff K/K0 ≥ 1 and P/P0 ≥ 1 on the optimisation and the
shifted grid. One Slurm job per part (`cshell/cluster/run.sh`, 1 GH200, optimise + finalise). Full run directories on
Isambard-AI: `/projects/b6bl/owenpeckham.b6bl/home_data/shell-lightweighting/runs_cshell/LW1/<part>/` (STL, fields,
checkpoints). Configs: `examples/cshell/<part>.json` with `output_dir` overridden.

| Part | V/V0 (saving) | K/K0 opt / shifted | P/P0 opt / shifted | governing | walls min / median / max mm | decision | job (time) |
| --- | --- | --- | --- | --- | --- | --- | --- |
| Cantilever | 0.976 (2.4 %) | 1.020 / 1.020 | 1.034 / 1.034 | buckling | 1.44 / 5.00 / 41.5 | accepted | 8241139 (35 min) |
| L-bracket | 0.961 (3.9 %) | 1.020 / 1.020 | 1.020 / 1.021 | strength | 1.50 / 3.44 / 42.2 | accepted | 8241144 (29 min) |
| MBB | 0.851 (14.9 %) | 1.020 / 1.019 | 1.025 / 1.025 | buckling | 1.50 / 2.26 / 25.8 | accepted | 8241141 (38 min) |

Mass saving is the exported STL against the solid (the export δ targets V_opt). Constant-mass baseline (round 4,
weighted objective): `examples/cshell/runs/<part>/`.

Files per part: `summary.txt`, `summary.json`, `history.csv`, `config_used.json`, `report.pdf`, `plots/*.png`.

A first attempt (`runs_cshell/LW1_try1`) used the weighted path's objective scaling (f0 normalised by the largest
per-variable gradient); MMA's elastic slack then made constraint violation cheap and MBB stalled at K/K0 0.990 (solid
kept). Fixed by f0 = (V/V0)/constraint_scale (commit 5d0375e).
