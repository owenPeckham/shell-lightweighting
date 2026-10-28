# Variable-thickness shelling: feasibility prototype

Prototype scripts from the feasibility study for upgrading the constant-mass shelling add-in to variable, region-wise
wall thickness (up to solid at joints) at constant mass. The full write-up, with results and recommendation, is the
"Variable-Thickness Shell Optimisation: Feasibility Study" doc.

Nothing here touches `shell_lightweighting/` or Fusion. Everything runs headless in Python on the STEP files.

## Setup

```
pip install build123d gmsh trimesh scikit-image scipy libigl skan matplotlib
# CalculiX: `ccx` on PATH (Linux: apt install calculix-ccx; Windows: bConverged / PrePoMax builds ship ccx.exe)
set SHELL_STEP_DIR=<folder with the six STEP files, spaces in names replaced by underscores (Solid_L-bracket.step, ...)>
set SHELL_OUT_DIR=<output folder>     (defaults: ../step and ../out)
```

Tested with Python 3.13, build123d 0.13.0 (OCP 8.0.1), gmsh 4.15.2, CalculiX 2.21, on 2 cores and 7 GB RAM.

## Pipeline (run in this order)

| Step | Script | What it does | Time per part (measured) |
| --- | --- | --- | --- |
| 1 | `segment.py Solid_L-bracket` | Mounts from analytic (plane/cylinder) faces; skeleton split into joint and strut regions with slenderness | 13-15 s |
| 2 | `fea_ratio.py Solid_L-bracket 2.5` | One solid C3D10 CalculiX run (bolt holes fixed, 200 N on pad); per-region R = U_shear/U_bend | 37-53 s |
| 3 | `run_designs.py Solid_L-bracket 1.700 [grid_mm]` | Signed-distance field; uniform and region-wise designs at constant mass; STL export | 4-7 min for 5 designs |
| - | `plot_seg.py`, `plot_sections.py` | Figures | |

Diagnostics used in the study: `inspect_step.py` (topology, volumes), `fusion_shell_check.py` (confirms Fusion's
outside shell is an exact +t offset), `occ_offset_test.py` / `occ_offset_diag.py` / `occ_simple_check.py`
(OpenCascade B-rep offset failures on these GD solids), `list_faces.py` (bolt-hole and pad faces).

## Key numbers (results/)

| Part | Fusion t (mm) | Uniform, auto mounts | Joints solid, original size | Joints solid and offset outward |
| --- | --- | --- | --- | --- |
| Cantilever | 1.756 | 1.794 | 1.760 | 1.620 |
| L-bracket | 1.700 | 1.743 | 1.371 | 0.964 |
| MBB | 2.100 | 2.132 | 1.818 | 1.332 |

## Known limitations

- Load directions and pads in `fea_ratio.py` are my reading of Figure 3 (-z on the top pad face); the L-bracket
  stiffness (0.049 kN/mm) sits above the paper's FE bounds, so check that case's load set-up.
- Region blending leaves a few small isolated voids and multi-shell STLs; add a minimum void size filter and escape holes.
- The 0.4-0.5 mm dense grid needs about 2-3 GB; a narrow-band / sparse grid (e.g. OpenVDB) is the production route.
