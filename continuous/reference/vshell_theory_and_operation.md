# vshell: Variable-Thickness Shelling Toolchain — Theory and Operation

Oct 28, 2026 · @Owen

## Overview

vshell takes a solid STEP part, its supports and its loads, and returns a mesh with the same mass. Each region of that mesh is either solid or hollow, with the hollow walls sized automatically. If no hollow design improves stiffness or load capacity, the original solid comes back unchanged with a report saying why.

- **Input:** one STEP solid plus a small JSON file. The JSON names the support faces (all degrees of freedom fixed, the minimum-stiffness bolt-hole case), the load faces with a force vector, and optional settings.
- **Output:** `final_design.stl` (the accepted hollow design, or a copy of the original), `best_candidate.stl` (the best hollow design analysed, whether accepted or not), CSV and JSON logs, and `report.pdf`.
- **Hard constraint:** the mass of every design equals the original's, to within the grid's volume error (about 0.1%).
- **Decision rule:** a hollow design is adopted only if it raises a combined score of stiffness and load capacity by at least 2%, and neither metric falls below 85% of the solid's. Otherwise the walls stay solid.
- **Runs headless:** Python 3.10+, no Fusion licence and no external FE solver; each Study 1 part took 15–25 minutes on a 2-core, 7 GB machine.

The package lives in your repo under `vshell/`, with example configs for the three Study 1 parts in `examples/`.

## Pipeline at a glance

&#91;embedded content: vshell pipeline · one run\]

One FE analysis of the solid sets the baseline and the ranking. Only the search stage is iterative, and the run ends either with the best hollow design or with the original part plus the reason.

## Geometry and the signed-distance field

The STEP solid is read once with OpenCascade (OCP), then everything downstream works on a signed-distance field, because OpenCascade's offset failed on all three Study 1 parts (feasibility study).

**Faces and boundary conditions** (`geometry.py`). Each B-rep face gets an id, type, area and centroid, plus an outward normal for planes and radius and axis for cylinders. Supports and loads are chosen with face selectors in the config:

| Selector | Example | Picks |
| --- | --- | --- |
| `ids` | `{"type": "ids", "ids": [9, 16, 24]}` | Those faces (ids from `python -m vshell faces part.step --png faces.png`) |
| `cylinder` | `{"type": "cylinder", "radius": 4.26, "tol": 0.05}` | Every cylinder of that radius: bolt holes |
| `plane` | `{"type": "plane", "normal": [0,0,1], "box": {"x": [-15, 20]}, "extreme": true}` | Planar faces facing +z with centroid in the box, keeping only the outermost: a load pad's top |

The solid is tessellated to a 0.02 mm chordal tolerance and coincident vertices are merged, which gives a watertight triangle mesh (V, T) with the B-rep face id of every triangle.

**Signed distance** (`field.py`, class `PartGrid`). The field is sampled on a regular grid of spacing h (default 0.4 mm, padded 4 mm):

```latex
\varphi(x) = \pm \min_{y \in \partial\Omega_0} \lVert x - y \rVert, \qquad \varphi < 0 \text{ inside}
```

1. A coarse occupancy grid (1 mm) is classified by the generalised winding number (libigl `fast_winding_number`), which is robust to small mesh defects.
2. Euclidean distance transforms inside and outside give an approximate φ on the coarse grid, copied to the fine grid by nearest index.
3. Within a band |φ| < 4 mm (plus 1.5 coarse cells), φ is replaced by the exact distance to the triangle mesh, signed by winding number (libigl `signed_distance`). This band holds every surface the design can move.

The volume of any implicit solid {ψ ≤ 0} is the sum of cell fractions, with the cell-centred sample treated as a linear ramp across the cell:

```latex
V[\psi] = h^3 \sum_{\text{cells}} \mathrm{clip}\!\left(\tfrac12 - \tfrac{\psi}{h},\, 0,\, 1\right)
```

For the original part (ψ = φ) this lands within 0.06% of the B-rep volume on all three parts. The optimiser holds mass to this field volume V₀, so the grid's own bias cancels.

## Segmentation into N control regions

The part is split automatically into mounts (never changed), joints and struts on a 1 mm grid (`segment.py`). Each joint or strut becomes one design variable.

**Mounts.** Fusion's generative design builds its generated geometry from freeform BSpline faces, while preserve regions (bolt lugs, load pads) keep planes and cylinders. So for every interior voxel the code looks up the nearest boundary triangle and its B-rep face type. Connected blobs of voxels nearest an analytic face become mounts. Blobs below 1% of the part volume are ignored (stray planar caps), and every voxel within 1.5 mm of a support or load face is a mount, however small. Mounts are dilated by 1.5 mm. This replaces the manual split-and-recombine step, and `mounts.extra_faces` adds more by face selector.

**Skeleton.** The occupancy grid is thinned to a curve skeleton (scikit-image `skeletonize`, Lee's 3D method) and converted to a branch graph (skan). A skeleton voxel with three or more neighbours is a junction. Junctions are clustered when their balls overlap:

```latex
\lVert p_a - p_b \rVert < k\,(r_a + r_b), \qquad r = \text{depth below the surface}, \quad k = 1.6
```

Each cluster's balls of radius k·r form one **joint** region. What remains of each branch is a **strut** region. Every interior voxel then takes the label of its nearest seed (Euclidean distance transform with indices), and mounts overwrite their voxels. Joint and strut regions under 0.2% of the volume are merged into their neighbours (mounts never are), and the labels are extended outside the part by nearest region so the outer offset has a region everywhere.

**Per-region quantities**, used by the design and the metrics:

| Symbol | Meaning | How it is computed |
| --- | --- | --- |
| r | typical radius | median skeleton depth (struts); mean junction depth (joints) |
| R\_in | inset that closes the void | max depth in the region + h/2 + 0.5 mm |
| L | strut length | arc length of its skeleton path |
| λ | slenderness | L / r\_g with r\_g = r/2, the radius of gyration of a solid circle |
| t(x) | local tangent | first principal axis of skeleton points within 3 mm |

On the three Study 1 parts this gives 9 (cantilever), 15 (L-bracket) and 28 (MBB) candidate regions, plus 4–8 mounts. The report's segmentation page shows the regions with their ids, so a doubtful split can be checked at a glance.

## Design parameterisation and the constant-mass constraint

A design is a choice of which candidate regions stay solid. One scalar, the outer offset s, then restores the original mass exactly, and s is also the wall thickness of every hollow region.

**Per-region solidity.** θᵢ = 1 keeps region i solid at its original size; θᵢ = 0 hollows it. Mounts always have θ = 1. A Gaussian blend G\_ℓ (σ = ℓ/2, transition length ℓ = 3 mm) turns the piecewise-constant tables into smooth fields:

```latex
\theta(x) = \big(G_\ell * \theta_{i(x)}\big)(x), \qquad R_{in}(x) = \big(G_\ell * R_{in,\,i(x)}\big)(x)
```

**Geometry of a design.** The void wall is moved inward by a(x) and the outer skin outward by d(x):

```latex
a(x) = \theta(x)\, R_{in}(x), \qquad d(x) = s\,\big(1 - \theta(x)\big), \qquad \psi(x) = \max\big(\varphi - d,\; -\varphi - a\big)
```

```latex
\Omega(s) = \{\, x : \psi(x) \le 0 \,\} = \{\, x : -a(x) \le \varphi(x) \le d(x) \,\}
```

- **Hollow region (θ = 0):** a = 0 and d = s, so the material is the band 0 ≤ φ ≤ s. The original solid becomes the void, and the wall of thickness s sits outside it, exactly as in the Fusion add-in.
- **Solid region (θ = 1):** d = 0 and a = R\_in exceeds every depth, so the material is φ ≤ 0, the original geometry.
- **Between (0 < θ < 1, within ℓ of a seam):** the void tapers shut and the outer skin blends back to the original surface, with no step. Next to a mount, the mount is grown by ℓ and counted as solid before the blur, and θ = 1 is imposed inside it afterwards. That keeps the whole taper in the candidate region, so bolt holes and pads are never moved.

**Mass constraint.** Volume grows monotonically with s, since d only increases. So the constraint has at most one root, found by Brent's method on \[0, s\_max\] (s\_max = 4 mm; the grid padding and exact-distance band are grown automatically to contain it):

```latex
V(s) = h^3 \sum \mathrm{clip}\!\left(\tfrac12 - \tfrac{\psi(x;\,s)}{h},\,0,\,1\right) = V_0
```

Only cells with φ < s\_max + h can change state, so each V(s) evaluation runs on that active subset (about 10% of the grid). One design's mass solve takes 1–4 s after the two Gaussian filters.

**Feasibility and clean-up**, applied in this order:

1. **Minimum wall:** s ≥ t\_min (1.5 mm nominal, the bottom of MJF's structural range); otherwise the design is infeasible and never analysed.
2. **Small voids:** enclosed voids under 30 mm³ (typically left where two blends meet) are filled, and s is re-solved with them filled.
3. **Escape holes:** each remaining enclosed void is listed, with suggested hole positions.

**The design family.** The candidates are ranked by how shear-dominated they are in the original solid (R, next sections). Design k keeps the k most shear-critical regions solid and hollows the rest. So k = 0 is the uniform shell of the current add-in, and k = N is the original part. This one ordered family (N + 1 designs) is what the optimiser searches.

## Finite-element model

Every design, including the solid baseline, is analysed by the same linear-elastic voxel model (`fe.py`), so the design-to-design ratios are free of meshing noise. It matches CalculiX C3D8I on an identical mesh to 4×10⁻⁵ and the paper's Abaqus tube stiffness ratios to within 4%.

**Mesh.** FE cells are cubes of side h\_FE = 2h (0.8 mm), each a 2×2×2 block of the distance grid. A cell's solid fraction vf is the mean of its eight fine-cell fractions. Cells with vf < 0.05 are dropped, and only the largest face-connected component is kept, so floating debris cannot make the stiffness matrix singular.

**Element.** An 8-node hexahedron with the three Wilson–Taylor incompatible bubble modes per direction, the formulation of CalculiX/Abaqus C3D8I. Trilinear elements lock in bending; the bubbles remove that. On a cube the internal modes are condensed once, with 2×2×2 Gauss integration:

```latex
K_e = K_{cc} - K_{ci} K_{ii}^{-1} K_{ic}, \qquad K_{xy} = \int_{\Omega_e} B_x^{\mathsf T} D\, B_y \, dV, \qquad u_{i} \propto (1-\xi^2),\,(1-\eta^2),\,(1-\zeta^2)
```

Every cell shares this 24×24 matrix, scaled by E·vf·h. Partially filled surface cells use the iso-strain (Voigt) mixture E\_e = E·vf: their axial and bending stiffness is right in the mean, and the surface error is second order.

**Boundary conditions.** Nodes within 0.75·h\_FE of a support face are clamped in all three directions, the paper's minimum-stiffness bolt-hole case. The load is shared equally by the nodes within the same distance of the load faces.

**Solver.** The global stiffness matrix is assembled in chunks of 20,000 cells and the fixed DOF are eliminated. K\_ff u = f is then solved by conjugate gradients, preconditioned with smoothed-aggregation algebraic multigrid (pyamg). The six rigid-body modes serve as near-null space, and the tolerance is a 10⁻⁸ relative residual. A 430,000-DOF model takes about 50 s and 2 GB, against 74 s and 4.6 GB for CalculiX's direct solver on the same mesh.

**Stresses.** Strain at the cell centre is exact for this element, because the bubble gradients vanish at ξ = 0:

```latex
\varepsilon_e = B(0)\, u_e, \qquad \sigma^{\mathrm{mat}}_e = D(E)\, \varepsilon_e, \qquad \sigma^{\mathrm{hom}}_e = \mathrm{vf}_e\, \sigma^{\mathrm{mat}}_e
```

The material stress σ^mat feeds the strength checks: in an iso-strain mixture the solid fraction carries Eε. The homogenised stress σ^hom integrates to section forces, so it feeds the axial-force and energy sums.

**Verification** (`tests/`):

| Check | Result |
| --- | --- |
| Same 108,906-cell hollow L-bracket mesh vs CalculiX 2.21 C3D8I | compliance 452.0332 vs 452.0331 N·mm; relative L2 displacement difference 4.2×10⁻⁵ |
| Study 2 PA11 tubes A–E, 3-point bend, h\_FE = 0.8 mm | stiffness 0.0086, 0.0105, 0.0160, 0.0245, 0.0367 kN/mm vs Abaqus 0.0086, 0.0104, 0.0158, 0.0245, 0.0351 |
| Same tubes, stiffness ratio to rod A | 1.21, 1.85, 2.84, 4.25 vs Abaqus 1.21, 1.84, 2.85, 4.08 |
| 1.7 mm wall tube, h\_FE = 0.8 / 0.6 / 0.4 / 0.3 mm | within −0.7% / +5.2% / −6.2% of the 0.3 mm result |

The last row sets the model's resolution limit. Walls of 1.5–1.8 mm carry about ±6% grid noise in absolute stiffness, mostly from how the support lines fall on the grid. That noise is shared by every design of a part, but a predicted change below about 5% should be read as no change.

## Performance metrics: the physics

Each design gets one stiffness and one load capacity. Capacity is the lowest of three linear-elastic limits (strength, Euler buckling of a compressed strut, local buckling of a hollow wall), which are the failure mechanisms Study 1 identified (`mechanics.py`).

**Stiffness.** The secant slope along the load:

```latex
K = \frac{\lVert F \rVert^2}{C}, \qquad C = f^{\mathsf T} u = 2U
```

**Strength.** The peak cell stress depends on the mesh, so a high quantile of the von Mises material stress is used instead (q = 0.995, weighted by solid volume). It is taken over the joints and struts, excluding cells within 2 mm of supports or loads. Mounts are never changed, so their stress is reported separately. In hollow regions stresses are scaled by s/(s − δ) for the as-built wall, δ = 0.2 mm (manuscript Table 5):

```latex
\lambda_s = \frac{\sigma_f}{Q_{0.995}\big[\,\sigma_{vm}\, s/(s-\delta)\,\big]}, \qquad \sigma_f = 36.6\ \mathrm{MPa}
```

σ\_f is back-calculated from Study 2's solid PA11 rod A (168.4 N at a 288 mm span, σ = FLc/4I). The 52 MPa datasheet tensile strength over-predicts every Study 2 tube by 30–57%; 36.6 MPa predicts tubes B, C and D within 2–9%, and over-predicts E by 11% (thin-wall local collapse).

**Euler buckling of struts.** The axial force of strut i is its homogenised axial stress integrated over the region and divided by its length, with σ\_tt = tᵀσt along the strut's own skeleton tangent t. For the second moment of area, each cell is assigned to its nearest skeleton station j of that strut. Station centroids are averaged over ±3 mm to give a smooth centroid line c̄, and each cell's offset r⊥ is measured from that line, perpendicular to t. The sums are divided by L\_c, the length the strut's cells actually cover, not the skeleton path length, which can run into a mount. For hollow struts, I\_min is scaled by (s − δ)/s for the as-built wall. On the L-bracket this reproduces exact mesh sections to within 5–8% (struts 7 and 8: 210 and 256 mm⁴, against 220 and 270–285 mm⁴):

```latex
N_i = \frac{1}{L_c} \sum_{e \in i} \sigma^{\mathrm{hom}}_{tt,e}\, h^3, \qquad M_i = \frac{1}{L_c}\sum_{e \in i} \mathrm{vf}_e h^3\, r_\perp r_\perp^{\mathsf T}, \quad r_\perp = (I - tt^{\mathsf T})(x_e - \bar c), \qquad I_{\min,i} = \mathrm{eig}_2(M_i)
```

```latex
P_{cr,i} = \frac{\pi^2 E I_{\min}}{(K_{\mathrm{eff}} L_m)^2}, \quad L_m = L_c + 2kr_i, \quad \lambda_E = \min_{i:\,N_i<0} \frac{P_{cr,i}}{\lvert N_i \rvert}
```

K\_eff = 0.7 (fixed–pinned) by default, since GD struts end in stiff joints or mounts. Hollowing at constant mass raises I\_min about threefold for a round strut, which is where the cantilever's +61% came from.

**Local buckling of hollow walls**, as-built thickness t = s − δ. Struts are treated as thin cylinders with the NASA SP-8007 knock-down, joints as plates of width b = 2r + s:

```latex
\sigma_{cr}^{\mathrm{tube}} = \gamma\,\frac{E\,t}{r_m\sqrt{3(1-\nu^2)}}, \quad \gamma = 1 - 0.901\left(1 - e^{-\sqrt{r_m/t}/16}\right), \qquad \sigma_{cr}^{\mathrm{plate}} = \frac{4\pi^2 E}{12(1-\nu^2)}\left(\frac{t}{b}\right)^2
```

The demand is the 0.995 quantile of the compressive principal stress in that region, so λ\_L = min over hollow regions of σ\_cr / σ\_c.

**Capacity:**

```latex
\lambda = \min(\lambda_s,\, \lambda_E,\, \lambda_L), \qquad P = \lambda\, \lVert F \rVert
```

The governing mode and region are recorded for every design and shown in the report.

**Shear-to-bending ratio**, used only to rank regions, from the original solid's stresses. With n₁, n₂ normal to t:

```latex
R_i = \frac{U^{s}_i}{U^{b}_i} = \frac{\sum_{e\in i} (\tau_{tn_1}^2 + \tau_{tn_2}^2)/(2G)\; \mathrm{vf}_e h^3}{\sum_{e\in i} \sigma_{tt}^2/(2E)\; \mathrm{vf}_e h^3}
```

A high R means the region carries load mainly by transverse shear, which needs material near the neutral axis (finding F5). Those regions are made solid first.

## Optimisation algorithm and acceptance rule

The optimiser searches the ranked family of designs with a fixed FE budget, scores each design against the solid, and adopts the best one only if it clearly improves the part.

**Objective.** For a design with stiffness K and capacity P, against the solid's K₀ and P₀ (weights default to w\_K = w\_P = 0.5):

```latex
J = \frac{w_K \ln(K/K_0) + w_P \ln(P/P_0)}{w_K + w_P}
```

This is the log of a weighted geometric mean of the two ratios. J = 0 is the solid, and a gain in one metric can offset a smaller loss in the other.

**Acceptance.** The best design is adopted only if all three conditions hold:

```latex
J > \ln(1 + g_{\min}), \qquad K/K_0 \ge f_K, \qquad P/P_0 \ge f_P \qquad (g_{\min} = 0.02,\ f_K = f_P = 0.85)
```

Otherwise `final_design.stl` is the original part and the report gives the reason, for example "combined gain J = −0.031 is not above ln(1.02)". This is the stay-solid behaviour you asked for: hollowing has to earn its place.

**Search** (`optimise.Run.run`):

1. **Baseline:** FE of the original gives K₀, P₀ and Rᵢ for every region.
2. **Rank:** candidates sorted by Rᵢ, highest first (or by slenderness, lowest first, with `search.rank_by = "lambda"`).
3. **Mass sweep:** for k = 0…N−1, solve s\_k for constant mass (no FE). Designs with s\_k < t\_min are discarded.
4. **Coarse pass:** FE of evenly spaced feasible k (budget B − 3, at least 2), always including the most hollow feasible design.
5. **Refinement:** FE of the unevaluated neighbours of the best k so far, until the budget B (default 10) runs out.
6. **Decision**, then outputs and the report. Designs are ranked acceptable-first, then by J, so an acceptable design always beats a higher-J one that breaks a floor. Ratios are taken as C₀/C and λ/λ₀, which stay defined when several loads partly cancel. Any design whose CG solve did not converge, or whose objective is not finite, is skipped and logged.

Why a ranked one-parameter family rather than all 2ᴺ solid/hollow combinations: each FE run costs 40–90 s, so a 25-region part could not be searched exhaustively. The ranking encodes the physics (shear-critical regions benefit least from hollowing), so the family runs from the add-in's uniform shell (k = 0) to the original solid (k = N) through the designs most likely to win. The family is not guaranteed to be unimodal (the cantilever's best design sits at k = 7 of 9, past a plateau), which is why the coarse pass spans the whole feasible range before refining.

## Code map, configuration and command line

The package is ten plain-Python modules with no compiled code of its own. Each stage can be called on its own, and the whole run is one function, `vshell.optimise.run(config)`.

| Module | Main objects | Role |
| --- | --- | --- |
| `config.py` | `DEFAULTS`, `load()`, `validate()` | JSON config merged onto defaults; face-selector syntax |
| `geometry.py` | `load_step`, `face_table`, `select_faces`, `tessellate` | OpenCascade (OCP) STEP import, face data, watertight triangulation |
| `field.py` | `PartGrid` | Signed-distance grid, design fields θ and R\_in, ψ, volume, Brent mass solve, void clean-up, FE cell fractions, marching-cubes export |
| `segment.py` | `segment()`, `Segmentation` | Mounts, skeleton, joints and struts, per-region r, R\_in, L, λ, tangents |
| `fe.py` | `unit_element_stiffness`, `VoxelMesh`, `assemble`, `solve`, `cell_stresses` | Incompatible-mode hexahedra, AMG-CG solver, centre stresses |
| `mechanics.py` | `evaluate()` | Stiffness, strength, Euler and local buckling, R per region |
| `optimise.py` | `Run`, `run()` | Pipeline, ranked search, acceptance, outputs |
| `report.py` | `write_report()`, `faces_figure()` | summary.json/.txt, plots/\*.png, report.pdf |
| `util.py` | `provenance()`, `write_csv`, `write_json` | Git hash, package versions, file hashes, logging helpers |
| `cli.py` | `main()` | Command line |

**Command line:**

```bash
python -m vshell faces "part.step" --png faces.png    # list faces, draw analytic faces with their ids
python -m vshell init "part.step"                     # template JSON (bolt-hole cylinders pre-selected)
python -m vshell run examples/l_bracket.json          # optimise, write outputs and report.pdf
python -m vshell run cfg.json --set grid.sdf_mm=0.5 search.max_fe_evaluations=6
python -m vshell report runs/L-bracket                # rebuild the report from saved files
```

**Configuration keys you are most likely to change** (all keys and defaults are in `config.py`):

| Key | Default | Effect |
| --- | --- | --- |
| `supports`, `loads` | required | Face selectors and force vector (N) |
| `manufacturing.t_min_mm` | 1.5 | Thinnest nominal wall allowed |
| `manufacturing.wall_bias_mm` | 0.2 | As-built wall deficit used in the strength and local-buckling checks |
| `material.strength_MPa` | 36.6 | Failure stress for the strength limit |
| `capacity.K_eff` | 0.7 | Euler effective-length factor |
| `objective.w_stiffness` / `w_capacity` | 0.5 / 0.5 | Weights in J; set one to 0 to optimise a single metric |
| `objective.min_gain`, `floor_*` | 0.02, 0.85 | Acceptance rule |
| `search.max_fe_evaluations` | 10 | FE budget for hollow designs |
| `grid.sdf_mm`, `grid.fe_factor` | 0.4, 2 | Resolution: FE cell = 0.8 mm; halve sdf\_mm for walls under 1.2 mm (memory ×8) |
| `design.transition_mm` | 3.0 | Length of the solid-to-hollow taper |

**Installation** (Windows, macOS or Linux, Python 3.10+):

```bash
pip install build123d numpy scipy scikit-image skan libigl pyamg trimesh matplotlib
```

build123d brings the OCP OpenCascade bindings. No CalculiX, Fusion or GPU is needed; CalculiX was used only to verify the element.

## Report and output files

Every run writes one folder, `<output_dir>/<name>/`, holding the meshes, machine-readable logs and a self-contained `report.pdf`. The report can be rebuilt from those files at any time with `python -m vshell report <folder>`. It follows the pattern of `baseline_optimiser.py`'s post-processing: summary text pages, then one page per figure. Each figure is built in its own try/except, and failures are listed in `plots/report_log.txt`.

| File | Contents |
| --- | --- |
| `final_design.stl` | The part to print: the accepted hollow design, or a copy of the original |
| `best_candidate.stl` | Best hollow design analysed, kept even when rejected |
| `original.stl` | Tessellated input, for comparison |
| `report.pdf`, `plots/*.png` | The report, and each page as a PNG |
| `summary.json`, `summary.txt` | Every headline number, machine- and human-readable |
| `run_info.json` | Config used, provenance (time, host, git commit, package versions, SHA-256 of code, config and STEP), face table, grid, baseline and best metrics, decision, timings |
| `regions.csv` | Per region: kind, rank, volume, r, R\_in, L, λ, R, strain-energy share, best-candidate and final state |
| `mass_sweep.csv` | Wall s and feasibility for every k |
| `evaluations.csv` | Per analysed design: k, s, mass, K, P, both ratios, J, λ\_s, λ\_E, λ\_L, governing mode and region, FE size, CG iterations, time |
| `escape_holes.csv` | Suggested hole positions and directions per enclosed void |
| `arrays.npz` | Cell stresses, slices and previews used by the figures |
| `config_used.json`, `run_log.txt` | Exact settings and the console log |

Report pages, in order:

1. **Summary** (1–2 pages): decision banner and reason; part, segmentation, baseline, best candidate, search, grid, timings, provenance.
2. **Geometry and boundary conditions**: three views, support faces blue, load faces orange, force arrow.
3. **Segmentation**: joints, struts coloured by slenderness, mounts, region ids.
4. **Baseline stress**: von Mises in two views, and a histogram against the failure stress.
5. **Regions**: R per region in rank order with strain-energy share; λ against R, marking the final state.
6. **Mass sweep**: constant-mass wall s (nominal and as built) against k, with t\_min and the analysed designs.
7. **Region table**: every region's numbers; hollow rows shaded.
8. **Search**: K/K₀ and P/P₀ against k with governing modes; J per design against the acceptance line.
9. **Capacity by mode**: strength, Euler and local-buckling loads, solid against best candidate. It is followed by \*\*Decision robustness\*\*: the analysed designs re-scored, without new FE, for K\_eff from 0.5 to 1.0 and σ\_f from −20% to +20%. λ\_E scales exactly with 1/K\_eff² and λ\_s with σ\_f, so this shows at once whether the decision would flip under plausible parameter error. The grid is also written to sensitivity.csv.
10. **Sections**: three orthogonal cuts through the largest joint, original and best candidate, tinted by solidity.
11. **Best candidate stress**: as page 4.
12. **Best candidate render**: outer surface, three views.
13. **Voids and escape holes**: void inventory, filled volume, suggested holes.
14. **Timing**: seconds per stage.

## Results on the Study 1 parts and validation

vshell hollows only the cantilever. There it hollows just the two compression struts and predicts 2.6× the solid's load capacity at −0.8% stiffness. The L-bracket and MBB stay solid because no hollow design raised their capacity. Against Study 1's measured capacity changes, those decisions point the right way for the cantilever and the L-bracket, but not for the MBB.

| Part | Candidate regions | Feasible / analysed | Solid K₀ | Solid P₀ (mode) | Best design | K/K₀ | P/P₀ | Decision | Run time | Peak RAM |
| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |
| Cantilever | 9 | 9 / 9 | 257 N/mm | 431 N (Euler) | 2 compression struts hollow, wall 1.91 mm | 0.992 | 2.63 | **Hollow adopted** (J = +0.48) | 19 min | 3.4 GB |
| L-bracket | 15 | 15 / 10 | 48.3 N/mm | 540 N (strength) | 14 of 15 regions hollow, wall 1.81 mm | 1.068 | 0.944 | Solid kept: J = +0.004, below 0.020 | 15 min | 2.4 GB |
| MBB | 28 | 28 / 10 | 473 N/mm | 2742 N (strength) | 1 region hollow (k = 27) | 1.000 | 1.000 | Solid kept: J = +0.0003 | 23 min | 4.5 GB |

All runs used the default config at 200 N; run times include the report. Reports are in `examples/runs/<part>/report.pdf`.

&#91;image: Cantilever best design: compression struts hollow (blue), tension struts and joints solid\]

The cantilever result is the selective hollowing that the manuscript's Section 6.2 recommends, found without being told. The two lower struts are in compression and limit the solid by Euler buckling (λ\_E = 2.16 at 200 N). Hollowing them at constant mass roughly triples their I\_min, so capacity rises until strength takes over. The upper struts are in tension and gain nothing, so they stay solid, as do the joints.

&#91;image: L-bracket search: stiffness rises with hollowing but load capacity falls in every design, so the solid is kept\]

&#91;image: L-bracket decision robustness: solid kept unless K\_eff = 1.0 and the failure stress is at or above 36.6 MPa\]

The decision-robustness page puts the decisions in context. The cantilever is hollow in all 25 combinations of K\_eff (0.5–1.0) and σ\_f (±20%), and the MBB is solid in all 25. The L-bracket flips to hollow only with pinned struts (K\_eff = 1.0) and σ\_f ≥ 36.6 MPa, so it is the borderline case.

**Validation against Study 1.** Design k = 0 is the uniform shell, the counterpart of Study 1's hollow parts. vshell's walls were 1.83, 1.81 and 2.20 mm against Fusion's 1.756, 1.700 and 2.100 mm, because the automatic mounts differ slightly from the manual split.

| Part | Stiffness change: vshell / Abaqus / measured | Capacity change: vshell / measured | Capacity sign |
| --- | --- | --- | --- |
| Cantilever | +1.2% / −11.6% / +3.5% (n.s.) | +148% / +61.1% | agrees; magnitude about 2.4× |
| L-bracket | +6.9% / +4.0% / −7.0% (n.s.) | −5.6% / −11.2% | agrees |
| MBB | −6.0% / −8.2% / −14.9% (n.s.) | −11.1% / +14.8% | disagrees |

| Part | Solid stiffness: vshell / measured / Abaqus bounds | Solid capacity: vshell / measured |
| --- | --- | --- |
| Cantilever | 257 / 193 / 163–268 N/mm | 431 / 513 N (−16%) |
| L-bracket | 48 / 31 / 25–38 N/mm | 540 / 453 N (+19%) |
| MBB | 473 / 602 / 364–899 N/mm | 2742 / 2436 N (+13%) |

What this says about the model:

- **Stiffness** changes stay within ±11% for every design, about the size of the FE grid noise. That matches the paper's finding of no significant stiffness change, so in practice load capacity decides.
- **Solid capacities** land within −16% to +19% of the tests, using one constant calibrated elsewhere (σ\_f, from the Study 2 rod) and the textbook K\_eff = 0.7.
- **Buckling-governed parts are captured.** The cantilever gain comes from the mechanism the paper identifies. Its size is over-predicted (+148% against +61%), partly because the solid is under-predicted by 16% and partly because the hollow test part likely failed by a mode the linear model does not see.
- **Thin-joint losses are captured.** On the L-bracket, the 0.2 mm as-built deficit raises the stress quantile in the hollow joints and capacity drops, as in the test.
- **Ductile strength gains are missed.** The model calls the MBB strength-limited (Euler margin 40×), so hollowing can only add stress, yet the test gained 14.8%. PA11 is ductile (35–50% elongation), so the likely mechanism is plastic redistribution, in which a hollow section's larger plastic modulus raises the collapse load. A peak-elastic criterion cannot see this.

In practice: trust a hollowing decision when buckling governs. Read a keep-solid decision under strength as conservative. Adding a plastic section-capacity check is the first extension to make (next section).

## Assumptions, limitations and extensions

The model is linear-elastic, isotropic and single-load-case, so its capacity numbers are proxies built from the mechanisms Study 1 identified. Treat them as a ranking of designs, not as certified failure loads.

**Assumptions to keep in mind:**

- **Material:** isotropic PA11, E = 1.76 GPa (Study 2 suggests the printed modulus is lower; ratios are unaffected). The failure stress of 36.6 MPa is calibrated on one solid rod.
- **Boundary conditions:** bolt-hole walls fully clamped, the paper's minimum-stiffness case. Its bounds spanned up to 147–151% for the MBB, so absolute stiffness carries that uncertainty; ratios between designs carry much less.
- **Strength** uses a 99.5% volume quantile, not the peak. Sharp stress concentrations finer than 0.8 mm are not resolved and are not meant to govern.
- **Euler buckling** treats each strut as an isolated column of length L + 2kr with K\_eff = 0.7. A strut cut in two by a spurious junction gets two shorter columns, which is unconservative. The frame's global sway modes are not checked.
- **Local buckling** uses classical elastic formulas with a knock-down. Study 1's L-bracket losses look more like local collapse of thin joint walls, which the strength term with the as-built wall factor captures better than elastic buckling.
- **As-built walls:** the 0.2 mm deficit enters capacity, but stiffness uses the nominal geometry.
- **Escape holes** are suggested, not cut, and their mass (about πd²t/4 each, roughly 13 mm³ for 3 mm holes in 1.8 mm walls) is not in the balance.
- **Search** covers one ranked family of N + 1 designs, not every solid/hollow combination, and each hollow region gets the same wall s.

**Natural extensions, roughly in order of value:**

1. **Plastic section capacity**, which fixes the MBB miss. Group cells by their nearest skeleton station, form N, V and M from σ^hom, and check (|N|/N\_p + |M|/M\_p)² + (|V|/V\_p)² ≤ 1, with N\_p = σ\_y A, M\_p = σ\_y Z\_p and V\_p = σ\_y A\_v/√3. Z\_p and A\_v come straight from the voxels, so hollow sections get credit for their larger plastic modulus and hollow joints pay for their smaller shear area.
2. **Nonlinear capacity:** a geometrically nonlinear or linear-buckling eigen-analysis of the final 2–3 designs. pyamg plus scipy's LOBPCG can do the eigen-solve on the same voxel mesh.
3. **Per-region wall thickness:** let each hollow region carry its own inset aᵢ (thicker walls where needed) and optimise with the boundary sensitivities from the feasibility study (MMA, about 20–40 variables).
4. **Several load cases:** run each case and take the worst ratio, or a weighted J.
5. **Escape holes in the mesh:** cut them by setting ψ > 0 inside small cylinders at the suggested positions, then re-solve s so the holes are inside the mass balance.
6. **Calibration:** once tests are possible, fit σ\_f and K\_eff on two or three solid parts and check the prediction on their hollow twins.
7. **Speed and memory:** a sparse narrow-band grid (OpenVDB) and a GPU conjugate-gradient solver should cut a 20-minute, 4.5 GB run to minutes and a fraction of the memory.
