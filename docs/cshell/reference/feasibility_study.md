> Renamed 2026-10-31: vshell → dshell (discrete optimiser, `dshell/`), continuous/ → docs/cshell/, shell_lightweighting/ → ushell/.

# Variable-Thickness Shell Optimisation: Feasibility Study

Oct 28, 2026 · @Owen

**Recommendation: build \[B\], a headless Python tool that varies wall thickness on a signed-distance field, using OpenCascade only to read the STEP. Drive it with one solid FEA per part (shear-to-bending energy ratio) plus slenderness, and keep FEA on hollow designs for validation.** Patch the current add-in's optimiser now, since that is a small fix that removes wasted CAD rebuilds.

- **Not \[A\].** Fusion's shell takes one thickness per feature, the API exposes no simulation, and each iteration took \~124 s in the example log. Per-region thickness would mean splitting and recombining bodies, with steps at every seam.
- **Not a B-rep version of \[B\].** OpenCascade's offset failed on all three GD solids (exceptions, or a "valid" part with a fifth of the volume). CadQuery, build123d and pythonOCC share that kernel.
- **The prototype works.** Mounting lugs and pads were found automatically from their analytic faces, replacing the manual split. The uniform case landed within 1.5–2.5% of Study 1's thicknesses. Constant-mass variable designs took 22–44 s each.
- **FEA beats geometry for the decision, but only once.** A 39–53 s solid run separates shear-dominated regions (L-bracket joints and short struts, R ≈ 0.5–0.65) from bending ones (cantilever struts, R 0.03–0.05). Slenderness alone cannot, for struts with λ 15–25.
- **The binding constraint is mass.** Solid joints at constant mass thin the L-bracket struts to 0.96–1.37 mm, below MJF's 1.5 mm structural minimum. Whether mass stays fixed is the first decision to make.

## Baseline from Study 1

Uniform constant-mass shelling left stiffness statistically unchanged (+3.5% to −14.9%) but moved load capacity by +61.1%, +14.8% and −11.2%, so a variable-thickness tool should target load capacity at joints, not stiffness alone.

| Artefact | Target wall (mm) | Measured wall (mm) | Stiffness change, measured | Stiffness change, FE (max bound) | Load capacity change |
| --- | --- | --- | --- | --- | --- |
| Cantilever | 1.756 | 1.555 | +3.5% (n.s.) | −11.6% | **+61.1%** |
| L-bracket | 1.700 | 1.486 | −7.0% (n.s.) | +4.0% | **−11.2%** |
| MBB | 2.100 | 1.893 | −14.9% (n.s.) | −8.2% | **+14.8%** |

Source: main.pdf, Tables 2, 3 and 5. Hollow parts were also 1.8–4.1% lighter than solid because of the escape holes.

What the manuscript says the next tool must handle:

- **F5 – region type matters.** Long struts gain from hollowing; short joints and cross-section transitions lose, because shear there needs material near the neutral axis.
- **F3 – local shell buckling at joints** drove the L-bracket loss, made worse by walls \~0.2 mm thinner than designed.
- **F7 – thin-wall FE over-predicts stiffness** (up to +31.6% numerically for Al tube E; +55% offset for PA11 even when solid). Any FEA-in-the-loop optimiser leans on its least accurate predictions.
- **Section 6.2** proposes two drivers for selective hollowing: a geometric measure (local strut slenderness) or a mechanical one (local ratio of shear to bending strain energy).

Written as objective inputs for N control regions i = 1…N, those two drivers become:

```latex
\lambda_i = \frac{L_i}{r_{g,i}}, \qquad r_{g,i} = \sqrt{I_i / A_i}
```

```latex
R_i = \frac{U^{\mathrm{shear}}_i}{U^{\mathrm{bend}}_i} = \frac{\int_{\Omega_i} \frac{\tau_{tn_1}^2 + \tau_{tn_2}^2}{2G}\, dV}{\int_{\Omega_i} \frac{\sigma_{tt}^2}{2E}\, dV}
```

Here t is the local skeleton (strut axis) tangent, so σ\_tt carries axial and bending stress and τ\_tn the transverse shear. A region with high λ and low R should be hollowed; one with low λ or high R should stay solid or thicken. Both are computed on the original solid, so neither needs thin-wall FEA.

## Current tool and geometry review

The add-in (ShellOptimisation.py, 785 lines) wraps one Fusion outside-shell feature in a hand-written 1-D Nelder–Mead loop. It works, but bugs waste about half of the CAD rebuilds, and nothing in it can vary thickness by region.

Loop as written: select one body → ShellFeature with `outsideThickness = t` (RoundedOffset) → if Fusion returns a surface, stitch at 0.1 / 1 / 10 mm, patch as fallback, then combine-cut → sum the masses of all bodies in the active component → objective (m − m₀)² → delete the features and repeat.

| Finding | Where | Effect |
| --- | --- | --- |
| Expansion, contraction and shrink steps evaluate `reflected_thickness` instead of the new point | lines 499, 510, 518, 522, 530 | Objective values are stored against the wrong thickness; the example log repeats 1.56 mm for iterations 1–6 and 1.54375 mm for 7–10 |
| An extra full shell rebuild every iteration, only for logging | line 538 | One more CAD regeneration per iteration; the example log took 1,485 s for 12 iterations (\~124 s each) |
| Nelder–Mead on a monotonic 1-D function | whole loop | m(t) rises monotonically with t, so a bracketed secant or Brent solve needs about 4–6 shells, not 12+ iterations of 2–3 shells |
| Mass = sum over every body in the active component | `weighComponent`, line 52 | Split-off mounting bodies are counted, which is what you want; kept tool bodies in the surface path (`isKeepToolBodies = True`) may be double-counted — worth checking |
| One body, one thickness | `createShellFeature` | ShellFeatureInput exposes only a single inside and outside thickness ([API reference](https://help.autodesk.com/cloudhelp/ENU/Fusion-360-API/files/ShellFeatureInput.htm)); no per-face override |

**What the STEP files show.** I parsed all six with OpenCascade and sampled 40,000 points on each shelled surface against its solid.

| Topology | Faces (solid → shelled) | Freeform BSpline faces in solid | Solid volume (mm³) | Shelled volume (mm³) | Change | Shelled skin at exactly +t |
| --- | --- | --- | --- | --- | --- | --- |
| Cantilever | 81 → 193 | 42 (bicubic, C1) | 61,023 | 59,712 | −2.1% | 44.6% |
| L-bracket | 127 → 333 | 93 (bicubic, C1) | 41,985 | 40,556 | −3.4% | 37.6% |
| MBB | 192 → 432 | 116 (bicubic, C1) | 85,377 | 82,141 | −3.8% | 48.4% |

- Fusion's outside shell is an exact outward offset: the outer skin sits at +t to within 0.05 mm, and the void wall is the original surface.
- Each shelled file is one connected shell, because the escape holes join the void to the outside. The modelled volume losses (−2.1%, −3.4%, −3.8%) match the weighed losses (−1.8%, −3.2%, −4.1%).
- The generated geometry is all bicubic BSpline; preserve regions (bolt lugs, load pads) are planes and cylinders. That split is what lets the mounting regions be found automatically (see Prototype results).

So the current tool is a special case of a simple field description: material where −a ≤ φ ≤ d, with φ the signed distance to the original surface, a = 0 and d = t everywhere.

## Platform comparison: \[A\] Fusion add-in vs \[B\] Python tool

&#91;B\] wins, but not as a B-rep offset tool: OpenCascade's offset failed on all three GD solids. The robust form of \[B\] uses OpenCascade only to read the STEP and classify faces, then builds the variable shell on a signed-distance field.

| Criterion | \[A\] Fusion API add-in | \[B1\] Python B-rep (OCCT via pythonOCC, CadQuery or build123d) | \[B2\] Python implicit field (OCCT for import only) |
| --- | --- | --- | --- |
| Variable thickness | ShellFeature takes one inside and one outside thickness. Per region means split body → one shell per region → combine, leaving steps at every seam | `BRepOffset_MakeOffset.SetOffsetOnFace` allows per-face offsets in principle | Any smooth thickness field; a region goes solid when its inset exceeds the local inradius |
| Robustness on these GD solids (tested) | Works (Study 1), with stitch/patch fallbacks already coded for surface results | Join offset threw on cantilever and MBB and silently returned a 20%-volume part for the L-bracket; simple offset closed but left 28–36% of the skin un-offset | Solid volume within 0.05% of the B-rep; uniform case gives 1.743 mm vs Fusion's 1.700 mm on the L-bracket |
| Cost of one mass evaluation | \~124 s per iteration (example log) | 0.3–7 s per offset, when it works | 5–6 s on a 34M-cell grid, after 2–3 min one-off setup |
| FEA coupling | No Simulation API found in the Fusion API reference; results would need STEP export to an external solver anyway | Direct (gmsh + CalculiX) | Direct; the same grid also gives a voxel FE mesh |
| Automation | Needs an open Fusion session, single-threaded, no batch runs | Headless, scriptable, parallel | Headless, scriptable, parallel |
| Output | Editable B-rep in the Fusion timeline | B-rep STEP | Triangle mesh (STL/3MF): fine for MJF, not an editable B-rep |
| Licence | Fusion subscription | Open source (OCCT LGPL) | Open source |

On the three OCCT front ends:

- **They share one kernel**, so all three inherit the offset failures above. Choosing between them is about ergonomics, not capability.
- **build123d** (I used 0.13.0 with OCP 8.0.1) installs with pip and still lets you call raw OCP classes, which is how every test here was written. Recommended for the STEP-in, face-classification and bolt-hole-detection layer.
- **CadQuery** shares OCP with build123d and has the CQ-editor you already have installed; equally fine.
- **pythonOCC** exposes the most of OCCT but is built from source or installed from conda ([install notes](https://github.com/tpaviot/pythonocc-core/blob/master/INSTALL.md)); no advantage here.

The Fusion route stays useful as a front end: the GD outcome, mounting-face picks and final inspection all live there. It is a poor place for the optimisation loop itself.

## Decision engine: FEA vs heuristics

Use a hybrid: geometry segments the part, one solid FEA per part (39–53 s here) sets the thickness allocation, and FEA on the hollow design is kept for final validation. Iterating thin-wall FEA inside the loop costs about two orders of magnitude more per step and is least accurate exactly where it decides (F7).

| Engine | Drives tᵢ from | Cost (measured on these parts) | Convergence | Efficacy and risk |
| --- | --- | --- | --- | --- |
| Geometric heuristic | Skeleton slenderness λᵢ = Lᵢ / r\_g,ᵢ | 13–15 s per part, no solver | Deterministic, one pass | Blind to load path: struts with λ 15–25 can be shear- or bending-dominated (chart below) |
| One-shot solid FEA | Rᵢ = U\_shear / U\_bend and strain-energy share per region, on the original solid | +39–53 s: gmsh C3D10 mesh 14–23 s, CalculiX solve 23–34 s, 283k–403k DOF | Deterministic, one pass | Reads the real load path. Uses solid-state stresses, so it misses redistribution after hollowing; linear only |
| FEA in the loop | Compliance or buckling load factor, gradient update (OC or MMA) over N regions | Each iteration needs a thin-wall mesh with ≥ 3 elements through a 1–2 mm wall (paper's rule). Element size drops from 2.5 to \~0.5 mm, so element count grows \~10²× to millions; 30–100 iterations | Needs sensitivities and move limits; mesh changes each step add noise | Best in principle. Runs on its least accurate predictions (F7) and well past the 128K-node Ansys Student limit ([source](https://ansys.synopsys.com/en-in/academic/students/ansys-student)) |

&#91;image: Slenderness vs shear-to-bending ratio for every region of the three parts\]

What the one-shot FEA showed (200 N, bolt-hole walls fixed, PA11 E = 1.76 GPa):

- **Cantilever:** 67% of strain energy sits in four struts with λ ≈ 86 and R = 0.03–0.05. Pure bending/axial: ideal for hollowing, consistent with its +61.1% load capacity.
- **L-bracket:** two short struts (λ 17.5 and 23.4) have R ≈ 0.6, as do most joints (R 0.52–0.89). This is the shear-critical topology that lost 11.2%.
- **MBB:** struts at similar λ (16–21) have R of only 0.03–0.11. Slenderness alone would have treated them like the L-bracket struts.
- **Stiffness check:** cantilever 0.262 kN/mm and MBB 0.478 kN/mm sit inside the paper's FE bounds; L-bracket 0.049 kN/mm is above them, so my load direction or pad there probably differs from Study 1. The ratios are relative, so this does not change the ranking.

Cheap physics constraints belong beside either engine, not inside FEA: a minimum wall of 1.5 mm plus the 0.2 mm MJF bias; local shell buckling of a tube wall, σ\_cr ≈ η·E·t / (r·√(3(1−ν²))) with a knock-down η; and Euler buckling of compression struts, P\_cr = π²EI / (KL)².

## Prototype results

The implicit pipeline reproduced the current tool's uniform thicknesses to within 2.5%, found the mounting regions automatically, and produced constant-mass variable-thickness parts for all three topologies. The catch: making joints solid at constant mass thins the struts, badly so for the L-bracket.

Where this ran: this session can read your shell-lightweighting folder but has no shell on your PC, and the Fusion API only runs inside Fusion. So I reviewed the add-in by reading it and ran every test on the STEP files in a cloud Python environment (OCP 8.0.1, gmsh 4.15, CalculiX 2.21, libigl, scikit-image).

**1. B-rep offset (OCCT) on the GD solids, offset 1.0 and 1.7 mm**

| Method | Cantilever | L-bracket | MBB |
| --- | --- | --- | --- |
| Offset by join, arc joints (what Fusion's shell resembles) | Exception | "Valid" solid with 20–25% of the original volume | Exception |
| Offset by join, intersection joints | Not done | Not done | Not done |
| Simple offset | Closed, but 28% of the skin not offset; OCC and mesh volumes disagree by 11% | Same pattern, 36% not offset | Closed, unchecked |
| Per-face offset (`SetOffsetOnFace`) | Exception | Same wrong result as above | Not run |

**2. Automatic segmentation into N regions (13–15 s per part).** Interior voxels whose nearest boundary face is analytic (plane or cylinder) become mounts, kept solid. A 1 mm voxel skeleton splits the rest into joints and struts.

&#91;image: Automatic segmentation: joints red, struts coloured by slenderness, mounts grey\]

Result: cantilever 8 joints and 11 struts, L-bracket 9 and 15, MBB 14 and 22. The mounts match the lugs and load pads you currently split by hand, which is why the uniform case lands within 1.5–2.5% of Fusion's thicknesses below.

**3. Constant-mass designs on a signed-distance field (0.4–0.5 mm grid).** Each design solves the one shared outward offset that restores the original mass, using 6–7 field evaluations (22–44 s).

| Part | Fusion t (mm) | Uniform, auto mounts | Joints kept solid at original size | Joints and stubby struts (λ < 15) kept original | Joints solid and offset outward |
| --- | --- | --- | --- | --- | --- |
| Cantilever | 1.756 | 1.794 | 1.760 | 1.765 | 1.620 |
| L-bracket | 1.700 | 1.743 | 1.371 | 1.475 | **0.964** |
| MBB | 2.100 | 2.132 | 1.818 | 1.848 | 1.332 |

Values are the resulting strut wall in mm. Field volume matched the B-rep to within 0.06%; the exported meshes are 0.1–0.2% under the original volume.

&#91;image: L-bracket sections through the largest joint, all at constant mass\]

What the numbers mean:

- **Solid joints cost the struts.** On the L-bracket, solid outward-grown joints leave 0.96 mm struts, under MJF's 1.5–3.0 mm structural range before the −0.2 mm print bias. Keeping joints at their original size leaves 1.37 mm.
- **The cantilever barely notices** (1.76 mm), because its joints hold only 11% of the shellable volume.
- **Artefacts to fix before production.** Blending a solid region into a hollow one leaves a few small isolated voids (visible in the sections) and the meshes come out as several shells. Both need a minimum-void-size filter and escape-hole placement.

## Proposed architecture

Two smooth fields over the original part, built from N per-region values, carry everything: a(x) moves the void wall inward and d(x) moves the outer skin outward. Mass is then a cheap scalar function of the N values, so it is enforced exactly by a 1-D solve after every update.

&#91;embedded content: proposed pipeline · 10 steps, one optional loop\]

Every step after reading the STEP runs headless in Python; Fusion is only the source of the GD outcome and the place to inspect the result. Times are from the prototype runs.

**Parameterisation.** φ(x) is the signed distance to the original solid (negative inside). Regions i = 1…N come from the segmentation; mounts are fixed with a = ∞ and d = 0.

```latex
a(x) = \sum_{i=1}^{N} w_i(x)\, a_i, \qquad d(x) = \sum_{i=1}^{N} w_i(x)\, d_i, \qquad w_i = G_\ell * \mathbf{1}_{\Omega_i}
```

```latex
\Omega(a,d) = \{\, x : -a(x) \le \varphi(x) \le d(x) \,\}, \qquad \text{wall } w(x) = a(x) + d(x)
```

G\_ℓ is a Gaussian with transition length ℓ ≈ 2–3 local radii, which turns region seams into tapers. A region is solid once aᵢ reaches its inradius rᵢ. The current tool is the special case aᵢ = 0, dᵢ = t.

**Constraints.**

```latex
M(a,d) = \rho \int H\big(-\psi\big)\, dV = M_0, \qquad \psi = \max(\varphi - d,\; -\varphi - a)
```

```latex
w(x) \ge t_{\min} + \delta_{\mathrm{bias}} \;(1.5 + 0.2\ \mathrm{mm, MJF}), \qquad \text{void size} \ge v_{\min}, \qquad \text{escape holes per void}
```

**Engine 1, heuristic allocation (one pass, recommended first).** Map each region's FEA ratio and slenderness to a solidity θᵢ, then solve the shared offset s for mass:

```latex
\theta_i = \mathrm{clip}\!\left(\frac{R_i - R_{lo}}{R_{hi} - R_{lo}},\,0,\,1\right) \vee \mathbf{1}[\lambda_i < \lambda_{\min}], \qquad a_i = \theta_i r_i, \qquad d_i = (1-\theta_i)\, s
```

Starting values from the data above: R\_lo ≈ 0.15, R\_hi ≈ 0.45, λ\_min ≈ 8. If the solved wall falls under t\_min, drop θ on the joints with the lowest strain-energy share until it fits, and report the trade.

**Engine 2, gradient loop (later, for compliance or buckling).** With only N ≈ 20–40 variables, MMA or SLSQP is enough. Mass and compliance sensitivities are boundary integrals on the void wall Γ\_v and outer skin Γ\_o, with W = σ:ε from the current design's FEA:

```latex
\frac{\partial M}{\partial a_i} = \rho \int_{\Gamma_v} w_i\, dA, \quad \frac{\partial M}{\partial d_i} = \rho \int_{\Gamma_o} w_i\, dA, \quad \frac{\partial C}{\partial a_i} = -\int_{\Gamma_v} w_i\, W\, dA, \quad \frac{\partial C}{\partial d_i} = -\int_{\Gamma_o} w_i\, W\, dA
```

**Quick win for the existing add-in.** For a uniform outward offset below the local curvature radius, added volume is a cubic in t (Steiner formula), with the linear term equal to the shelled area. Two shell evaluations plus a fit give t, so the Nelder–Mead loop can go entirely.

## Roadmap, risks and open questions

Four phases, the first three closed by a measurable gate; the Phase 2 print-and-test decides whether the gradient engine is worth building. The prototype scripts, results and figures are in your repo under `feasibility/`, with a README giving run order and timings.

&#91;embedded content: roadmap · 4 phases, 3 gates\]

**Risks**

- **Mesh, not B-rep, output.** Fine for MJF and for import into Fusion as a mesh body; an editable B-rep would need manual conversion.
- **Grid cost.** A dense 0.4 mm grid used 2–3 GB for one part. Thinner walls or bigger parts need a sparse narrow-band grid such as OpenVDB.
- **Segmentation on new topologies.** Skeletons can produce spurious junctions; keep an override (face picks or region edits) for unusual parts.
- **Solid-state stresses.** Rᵢ comes from the unhollowed part, so it misses load redistribution. Phase 2's hollow FEA checks this.
- **Boundary conditions.** The paper's stiffness bounds spanned 147–151% for the MBB; Rᵢ is a ratio and less sensitive, but regions next to mounts still depend on the fixing model.
- **Linear proxies for a nonlinear target.** Load capacity came from buckling and local failure; linear buckling only approximates that.

**Open questions for you**

- [ ] Is constant mass a hard constraint? Solid joints at constant mass leave L-bracket struts at 0.96–1.37 mm. "Equal-or-better load capacity at no more than the original mass" would give the optimiser room.
- [ ] Is a mesh result acceptable for the next study, or must it stay an editable Fusion body?
- [ ] Which load case and fixing model should drive Rᵢ: the paper's minimum-stiffness bolt-hole case as used here, or the kinematic washer case?
- [ ] Should R\_lo and R\_hi be calibrated against a physical test (for example two L-bracket variants) before they are fixed?

## Sources

- main.pdf: Peckham, Hicks and Goudswaard, *Hollow Promises?* (manuscript dated 23 October 2026), Tables 2–5, Sections 6.1–6.2
- shell-lightweighting repo: `ushell/ShellOptimisation.py` and the example log
- [Fusion API: ShellFeatureInput](https://help.autodesk.com/cloudhelp/ENU/Fusion-360-API/files/ShellFeatureInput.htm) and [ShellFeature.setThicknesses](https://help.autodesk.com/cloudhelp/ENU/Fusion-360-API/files/ShellFeature_setThicknesses.htm)
- [Autodesk forum: simulation from the API](https://forums.autodesk.com/t5/fusion-api-and-scripts-forum/how-to-run-simulation-from-api-and-extract-found-parameters/m-p/8841078) (2019, newest answer found)
- [Ansys Student problem-size limits](https://ansys.synopsys.com/en-in/academic/students/ansys-student)
- [pythonOCC install notes](https://github.com/tpaviot/pythonocc-core/blob/master/INSTALL.md)
