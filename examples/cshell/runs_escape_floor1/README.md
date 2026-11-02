# cshell results: escape holes on, all floors 1.00

Single-run cshell (one start from the solid, optimise + finalise in one Slurm job, 2026-11-02) of the three example
parts with the round-4 defaults except:

```
--set export.escape_holes=true objective.floor_stiffness=1.0 objective.floor_capacity=1.0 \
      objective.floor_strength=1.0 objective.floor_buckling=1.0
```

i.e. K/K0 ≥ 1, P/P0 ≥ 1, σ_PN0/σ_PN ≥ 1 (strength) and λ_B/λ_B0 ≥ 1 (buckling), each relative to the solid part,
all as MMA constraints. `floor_strength` / `floor_buckling` were added for this run (see `cshell/README.md`): P is a
soft-min of strength and buckling, so a P floor alone binds only the governing mode. Full runs on Isambard-AI:
`/projects/b6bl/owenpeckham.b6bl/home_data/shell-lightweighting/runs_cshell/SR_ESC1/<part>/`. Round-4 results with the
0.85 floors and no escape holes: `../runs/`. File layout as in `../runs/README.md`, plus `escape_holes.csv` (hole
positions and axes) and `history_solid.csv`.

## Results (shifted grid)

| Part | Decision | J | K/K0 | P/P0 | λ_B/λ_B0 | σ_PN/σ_PN0 | governing | escape holes (voids) |
| --- | --- | --- | --- | --- | --- | --- | --- | --- |
| MBB | accepted | +0.4926 | 1.111 | 2.411 | 2.449 | 0.859 | buckling | 4 (2) |
| Cantilever | **solid kept** | +0.4487 | 0.999 | 2.456 | 2.462 | 1.001 | buckling | 2 (1) |
| L-bracket | accepted | +0.0731 | 1.099 | 1.053 | 1.074 | 0.950 | strength | 4 (2) |

Round 4 (floors 0.85, no holes), for comparison: J = +0.4954 / +0.4781 / +0.0700.

**Cantilever:** the optimum sits on the stiffness and strength floors (K/K0 0.9993, σ_PN/σ_PN0 1.001 on the
optimisation grid, feasible within `objective.feasibility_tol` = 1e-3 in log). On the shifted re-score grid K/K0 =
0.9988 (ln = −1.18e-3), 0.02 % beyond the tolerance, so finalise keeps the original solid. The design itself is
exported (`fields.vtp`, previews). A run with a small margin on the floors (e.g. 1.005) would likely be accepted.

**L-bracket:** the minimum wall 0.49 mm is at an escape-hole rim; off the rims it is 1.45 mm.

## Preview STLs

`final_design_preview.stl`: 500 k-face quadric decimation of the full export (`previews.json`); `original_preview.stl`
is copied from `../runs/` (same original solid).

| Part | Full triangles (final) | Full volume (mm³) | ΔV preview vs full | Shells (full / preview) | Watertight (full / preview) |
| --- | --- | --- | --- | --- | --- |
| MBB | 4,753,836 | 85,216.7 | +0.004 % | 1 / 1 | yes / yes |
| Cantilever | 3,360,836 | 60,940.8 | -0.001 % | 1 / 1 | yes / yes |
| L-bracket | 1,599,944 | 41,947.5 | -0.001 % | 1 / 1 | yes / yes |

With escape holes every enclosed void is opened to the outside, so each part is a single shell.
