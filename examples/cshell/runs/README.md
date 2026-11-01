# cshell results for the three example parts

Copies of the round-4 validation runs (2026-11-01; one start from the solid, optimise + finalise in one Slurm job) from
Isambard-AI `/projects/b6bl/owenpeckham.b6bl/home_data/shell-lightweighting/runs_cshell/SR_VAL4/<part>/`, one folder
per part: `MBB/`, `Cantilever/`, `L-bracket/` (configs: `examples/cshell/<part>.json`). Results table and known issues:
`docs/cshell/STATUS.md` §5. The Stage 2 copies (3 starts, collect stage) are replaced; they remain in git history.

## Files in each `<part>/`

| File | Contents |
| --- | --- |
| `report.pdf`, `plots/*.png` | The report and its figures, one per page |
| `summary.txt`, `summary.json` | Decision; J, K/K0, P/P0 on the optimisation and shifted grids; mass check; export checks (shells, voids, fins removed); timing and peak memory |
| `config_used.json` | Full configuration of the run (defaults filled in) |
| `run_log.txt`, `finalise_log.txt` | Optimiser log and finalise log |
| `history.csv` | Optimiser history per iteration (J, K, P, mass error, timings) |
| `xi_solid_phase0.npy`, `xi_solid_phase1.npy`, `design_solid_phase*.npz` | Design vector ξ per design-surface vertex and the design fields at the end of each continuation phase |
| `done_solid.json` | Per-phase result record (J, growth steps, wall range) |
| `fields.vtp` | Design surface with thickness, offsets, stresses (open in ParaView) |
| `bc_faces.npz` | Surface mesh (`V`, `T`) with support and load masks (for plotting) |
| `final_design_export.json`, `wall_thickness.csv` | Export record (tiles, clean-up, volumes, mass correction) and wall-thickness histogram |
| `final_design_preview.stl`, `original_preview.stl` | **Decimated previews** (500 k triangles, binary, 25.0 MB each) |
| `previews.json` | Source path, full and preview triangle counts, volumes, watertightness and shell counts |

## Preview STLs

Quadric edge collapse (pymeshlab, topology and boundary preserving, target 500 k faces). `final_design` is from
the SR_VAL4 run (multi-shell: sealed voids are separate shells, preserved in the preview). `original_preview.stl` is
kept from the earlier previews: the original solid is the same part (source in `previews.json`).

| Part | Full triangles (final) | Full volume (mm³) | ΔV preview vs full | Shells (full / preview) | Watertight (full / preview) |
| --- | --- | --- | --- | --- | --- |
| MBB | 4,747,668 | 85,217.5 | +0.003 % | 3 / 3 | yes / yes |
| Cantilever | 3,582,528 | 60,940.7 | +0.003 % | 2 / 2 | yes / yes |
| L-bracket | 1,562,200 | 41,947.4 | −0.001 % | 5 / 5 | yes / yes |

Original solids: 38,570,616 / 29,631,268 / 18,947,100 triangles; preview ΔV −0.007 / −0.005 / −0.003 %.

## Not in the repository (Isambard-AI only)

At `.../runs_cshell/SR_VAL4/<part>/`: the full-resolution `final_design.stl` (78–237 MB; GitHub's limit is 100 MB per
file), `checkpoints/`, `run_info.json`, `run_info.d/`. Original STLs: `.../runs_cshell/FULL/<part>/original.stl`.

## Regenerating

- Previews, from the repository root on Isambard-AI (CPU job, 4 cores, 24 GB, ≈ 1–2 min per part):
  `sbatch --mem=24G --cpus-per-task=4 examples/cshell/runs/make_previews.sh <run dir> <part>`
  (`make_previews.py <run dir> <out dir> [--faces N] [--names final_design original]` does the work; other entries in an
  existing `previews.json` are kept).
- The runs: `cshell/cluster/launch_part.sh examples/cshell/<part>.json`; finalise again with
  `sbatch cshell/cluster/export.sh <run dir>` (see `cshell/README.md`).
