# Export timings (Task E, `cshell/export.py`, THEORY §8.1-8.5)

Job 7796052 (`bench_all.sh`: one node, 72 Grace cores, `--mem=100G`, CPU only, 8 min 18 s in total); tests job 7796042
(12 passed). Times in seconds; `A` = geometry pass (closest-point queries, codes, shell table), `B` = clean-up (opening by a
3 mm ball, filling, islands, opening-rule re-check), `C` = escape holes, `D` = mass correction (Brent on the cached shell),
`E` = marching cubes + weld + checks + STL, `W` = wall-thickness ray casting. RSS = parent peak (GB, `getrusage`; Slurm's
MaxRSS sums forked workers and is not meaningful). All meshes: watertight, edge-manifold, consistently oriented (own
checks; trimesh cross-check below 2 M faces), |V_mesh/V0 − 1| well below R1 (2e-3). Tile 128^3, block 16.

| case | h_x (mm) | faces | A | B | C | D | E | W | total | RSS | delta (mm) | V_mesh/V0 − 1 | bodies |
|---|---|---|---|---|---|---|---|---|---|---|---|---|---|
| B2 hollow (xi_d = 0.5, xi_a = 0) | 0.1 | 1.62 M | 1.0 | 4.2 | 0.6 | 1.0 | 4.5 | 1.3 | 12.5 | 3.0 | 0.0248 | −1.8e-4 | 1 |
| B2 hollow | 0.05 | 6.45 M | 4.6 | 19.4 | 1.3 | 5.6 | 14.0 | 2.4 | 47.3 | 3.3 | 0.0261 | −1.1e-4 | 1 |
| B2 original | 0.05 | 2.69 M | 1.0 | – | – | 0.6 | 4.5 | 2.1 | 8.1 | 3.3 | 0 | −2.2e-5 | 1 |
| B2 uniform shell (t = 1.9226 mm) | 0.05 | 6.63 M | 1.7 | – | – | 1.8 | 13.9 | 2.0 | 19.5 | 3.9 | −0.0006 | −3.3e-5 | 2 |
| cantilever original (solid) | 0.1 | 7.41 M | 10.0 | – | – | 1.1 | 16.5 | 2.7 | 30.2 | 2.6 | 0 | −1.6e-4 | 1 |
| cantilever uniform shell (t = 1.8613 mm) | 0.1 | 17.79 M | 12.7 | – | – | 3.3 | 21.0 | 2.8 | 39.7 | 6.7 | −0.0005 | 1.4e-5 | 2 |
| **cantilever original (solid)** | **0.05** | 29.63 M | 58.1 | – | – | 4.3 | 38.3 | 4.9 | **105.6** | 12.3 | 0 | −3.4e-5 | 1 |
| **cantilever uniform shell** | **0.05** | 71.19 M | 71.5 | – | – | 14.4 | 97.5 | 6.3 | **189.8** | 30.9 | −0.0006 | 9.9e-6 | 2 |

* The two cantilever rows at h_x = 0.05 are the production setting (3328 × 2944 × 2944 grid, 1.0·10^9 closest-point queries
  for the shell, 71 M-triangle STL of 3.6 GB). End-to-end for the whole benchmark: B2 (context 5 s + design + both baselines)
  80 s, cantilever 304 s.
* "bodies" counts connected shells of the mesh: the uniform shell has the outer skin plus the enclosed original volume (2).
  The B2 hollow design has 1 body: the void is joined to the exterior by the two escape holes (surface of genus 1).
* B2 hollow: V0 = 6272.5 mm^3; raw void 5908.9 mm^3, opening removed 62.7 mm^3 (tapering ends), nothing filled, one
  material component; re-opening changed 0 cells; delta = +0.026 mm (< 0.1 warn); 0.8 % of wall samples are below
  t_min = 1.5 mm (p1 = 1.59 mm nominal, 1.44 mm as-built), the thin ones are at the hole rims and the passive-zone
  transition; 2 holes (D = 3 mm) 74 mm apart along the rod axis (x = 3.0 and 77.0 mm).
* Cost structure: the first pass and the opening dominate for designs with voids (the opening EDT blocks are 250^3 at
  h_x = 0.05; capped at 24 workers for memory); meshing and the global weld dominate for solids. Block pruning (far-field and deep interior, rigorous for
  phi-only designs) reduces the 2.9·10^10-cell grid of the cantilever to 4.1·10^8 closest-point queries (solid) and 9.9·10^8
  (uniform shell).
* Not timed: a void-bearing design of the cantilever at h_x = 0.05 (needs a full `Context`, ~5 min, and a converged xi).
  Its A and B stages scale with the number of tiles that contain void; expect a few times the B2 hollow cost per mm^2 of wall.

Reproduce: `sbatch continuous/research/export/bench_all.sh` (or `bench.sh CASE H_X`); outputs in
`$SCRATCH/cshell_export_bench_final/<case>_h<h>/` (`*_export.json` summaries with per-stage timings, STLs,
`escape_holes.csv`, `wall_thickness.csv`).
