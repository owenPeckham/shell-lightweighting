# cshell results for the three example parts

Copies of the Stage 2 full runs (2026-10-30/31; 3 starts per part, collect jobs 7888538–40) from Isambard-AI
`/projects/b6bl/owenpeckham.b6bl/home_data/shell-lightweighting/runs_cshell/FULL/<part>/`, one folder per part:
`MBB/`, `Cantilever/`, `L-bracket/` (configs: `examples/cshell/<part>.json`). Summary of the decisions:
`docs/cshell/STATUS.md` §5; comparison with solid, uniform shell, dshell and the manuscript: `docs/cshell/VALIDATION.md` V6.

## Files in each `<part>/`

| File | Contents |
| --- | --- |
| `report.pdf` | The report: decision, comparison table, geometry/BCs, fields, sections, convergence, stress, buckling, optimality, robustness, resolution, voids, resources, provenance |
| `plots/*.png` | The report's figures, one per page |
| `summary.txt`, `summary.json` | Decision and reason; J, K/K0, P/P0 per resolution; mass check; comparison rows (solid, uniform shell, dshell best, cshell per start, cshell export); validation status; sensitivity (schema `cshell.summary/1`) |
| `config_used.json` | Full configuration of the run (defaults filled in) |
| `run_log.txt`, `collect_log.txt` | Optimiser and collect logs (raw; they use the old start name `vshell`) |
| `history.csv`, `history_<start>.csv` | Optimiser history per iteration (J, C/C0, λ ratios, mass error, timings) for the starts `uniform`, `dshell`, `neutral` |
| `xi_<start>_phase0.npy`, `xi_<start>_phase1.npy` | Design vector ξ = [ξᵈ, ξᵃ] (per design-surface vertex) at the end of each continuation phase |
| `fields.vtp` | Design surface with thickness, offsets, stresses (open in ParaView) |
| `buckling_modes.npz` | Surface mesh (`V`, `T`) with the first buckling mode magnitude and eigenvalue of the solid and of the design |
| `bc_faces.npz` | Surface mesh (`V`, `T`) with support and load masks (for plotting) |
| `best_candidate_export.json` | Export record: tiles, clean-up (voids, islands), escape holes per void, volumes, mass correction |
| `wall_thickness.csv` | Wall-thickness histogram of the exported design (nominal and as-built) |
| `escape_holes.csv` | Escape (powder-removal) holes: void, position, normal, local wall, score, void volume |
| `sensitivity.csv` | Decision re-scored for other knock-down / failure-stress assumptions |
| `final_design_preview.stl`, `original_preview.stl`, `uniform_shell_preview.stl` | **Decimated previews** (~0.5 M triangles, binary) of the final design, the original solid and the uniform-shell baseline |
| `previews.json` | Triangle counts, volumes and watertightness of the full and preview STLs |

The start that begins from the discrete optimiser's design was called `vshell` when these runs were made. In this copy
its files are renamed (`history_dshell.csv`, `xi_dshell_phase*.npy`) and the start label is `dshell` in
`summary.json`, `summary.txt`, `config_used.json` and the history CSVs; the logs, `report.pdf` and the plots keep
`vshell`. cshell reads either name (start alias), so `python -m cshell report <part dir>` works on this copy.

## Preview STLs

Quadric edge collapse (pymeshlab `meshing_decimation_quadric_edge_collapse`, topology and boundary preserving,
target 500 k faces) of the full STLs. Volume change relative to the full-resolution mesh:

| Part | File | Full triangles | Full volume (mm³) | Preview triangles | ΔV preview vs full | Watertight (full / preview) |
| --- | --- | --- | --- | --- | --- | --- |
| MBB | final_design | 82,898,852 | 85,238.0 | 500,000 | +0.030 % | yes / yes |
| MBB | original | 38,570,616 | 85,228.8 | 500,000 | −0.007 % | yes / yes |
| MBB | uniform_shell | 86,044,292 | 85,230.5 | 500,000 | −0.014 % | yes / yes |
| Cantilever | final_design | 60,711,700 | 60,944.1 | 500,000 | +0.025 % | yes / yes |
| Cantilever | original | 29,631,268 | 60,939.2 | 500,000 | −0.005 % | yes / yes |
| Cantilever | uniform_shell | 62,536,188 | 60,940.6 | 500,000 | −0.010 % | yes / yes |
| L-bracket | final_design | 29,098,532 | 41,955.9 | 500,000 | +0.001 % | yes / yes |
| L-bracket | original | 18,947,100 | 41,946.3 | 500,000 | −0.003 % | yes / yes |
| L-bracket | uniform_shell | 35,351,688 | 41,946.9 | 500,000 | −0.006 % | yes / yes |

Each preview is 25.0 MB. Watertight here means no boundary and no non-manifold edges (pymeshlab topological
measures). Jobs 7891439–41: 27–72 min and 29–72 GB peak RSS per part.

## Not in the repository (on Isambard-AI only)

At `/projects/b6bl/owenpeckham.b6bl/home_data/shell-lightweighting/runs_cshell/FULL/<part>/`: the full-resolution
STLs `final_design.stl`, `best_candidate.stl`, `original.stl`, `uniform_shell.stl` (1–4 GB each; GitHub's limit is
100 MB per file), `export_baselines/`, `checkpoints/`, `done_<start>.json`, `run_info.json`, `run_info.d/`.

## Regenerating

- The previews, from the repository root on Isambard-AI (CPU job; MBB took 72 min and 72 GB peak, so request ~100G and 1.5 h):
  `sbatch --mem=160G examples/cshell/runs/make_previews.sh <full run dir> <part>`
  (`make_previews.py <full run dir> <out dir> [--faces N] [--max-dvol 0.005]` does the work).
- The runs: `cshell/cluster/launch_part.sh examples/cshell/<part>.json` (three starts on one GPU each, then collect
  and report; see `cshell/README.md`); a report alone: `python -m cshell report <run dir>` in a GPU job.
