# Watertight export (2026-10-30/31)

Scripts (run from the repo root, never on the login node):
- `cpu_py.sh <script.py> [args]`: CPU job wrapper (`sbatch [--mem --time --job-name] ...`), log next to it.
- `diag_stl.py <stl> [h]`: welds an exported STL by exact position, counts duplicated directed edges, locates them on
  the export lattice (how many coordinates integral, near a tile face?), lists bodies with signed volume and bbox.
- `mc_ties.py`, `mc_resid.py`: skimage 0.26 Lewiner MC on random 48³ fields: duplicated directed edges with/without a
  ψ shift and with ambiguous faces removed.

## Diagnosis (production collects at 36f2b5a)
`diag_stl.py` on best_candidate.stl (jobs 7876958 L-bracket, 7876959 MBB, 7876960 Cantilever):

| part | faces | dup directed edges | on tile faces | extra bodies | extra bodies are |
|---|---|---|---|---|---|
| L-bracket | 29.1 M | 130 | 22/91 vertices (no tile pattern) | 5 | enclosed voids −0.0014…−0.040 mm³, near (−10, −6.4, 126) |
| MBB | 83.0 M | 1072 | - | 32 | enclosed voids −0.0013…−0.56 mm³, 0.05-0.1 mm thick slits at x = ±125, ±5 |
| Cantilever | 60.8 M | 74 | 0/62 | 0 | - |

Every duplicated directed edge is an edge whose undirected edge carries four triangles (two each way) lying on one
cube face with four crossings: an *ambiguous face* resolved differently by the two cubes sharing it. Some faces use
MC33 cube-centre vertices (0 integral coordinates), which the old weld key also assigned to a lattice edge.

Causes (mc_ties 7877003, mc_resid 7877011):
1. **Decider ties**: ±3h code-only cells and clipped ψ give ac = bd on a face; binary ±B fields: 13998 dup edges,
   0 after any 1e-4 h shift.
2. **skimage Lewiner case 10 bug**: tie-free gaussian fields still give 12 dup edges in 40 × 48³ (two adjacent case-10
   cubes, signs [1,0,0,1,1,0,0,1]); 0 once ambiguous faces are removed.
3. **Extra bodies**: sub-cell enclosed void slits (≤ 0.56 mm³ ≪ void_min_mm3 = 30) that the grid clean-up cannot see.

## Fix (cshell/export.py)
- `resolve_ambiguous_faces`: before MC, every material node on an ambiguous lattice face is moved to the void side
  (ψ ← +1e-3 h), iterated to a fixed point (6-connectivity of material, as `ndimage.label` in the clean-up). Evaluated
  per tile on a block with a halo of `MC_HALO` = 32 nodes (`_psi_block(..., halo)`), so neighbouring tiles agree on shared
  nodes. A halo of 4 was not enough (L-bracket uniform shell: cascades of 11 iterations, 6 missing reverse edges).
- Weld keys: vertices not on a lattice edge (MC33 centre vertices) get unique keys.
- `drop_material_islands(..., void_fill=void_min_mm3, near_mount=…)`: enclosed voids below void_min_mm3 are filled
  on the mesh (§8.2) and reported (`voids_filled_mm3`); dropped material islands touching a mount face are counted
  (`islands_on_mount`) and logged.
- Summary `mesh.ambiguous_faces`: nodes flipped (upper bound h³ each), max iterations, tiles not converged / halo exceeded.
