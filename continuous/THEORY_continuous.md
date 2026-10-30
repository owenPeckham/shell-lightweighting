# cshell: continuous variable-thickness shelling at constant mass

## Theory and operation (implementation specification, v1.2, 28 October 2026; v1.1 adds Owen's decisions of §14; v1.2 records the implementation deviations D1, D2, D3, D6–D7 of §14 — see `continuous/STATUS.md` §4)

Author: Owen Peckham (DMF Lab, University of Bristol), drafted with Claude. Audience: the Claude Code instance that will implement `cshell` on the GPU cluster, and anyone who later maintains it.

This document specifies the **continuous** successor to `vshell`. It covers:

- the geometry representation;
- the mechanics;
- the optimisation problem and its sensitivities;
- the optimiser;
- the outputs and report;
- the software layout;
- the verification plan and the staged roadmap.

It is self-contained, but it assumes you have read these files in the repository:

| File | Why |
| --- | --- |
| `continuous/reference/main.pdf` | The manuscript (Peckham, Hicks, Goudswaard, *Hollow Promises?*): the physical evidence this tool must respect. Tables 2–5 and Sections 6.1–6.2 matter most. |
| `continuous/reference/vshell_theory_and_operation.md` | Theory of the discrete predecessor `vshell` (per-region solid/hollow, one wall offset). Its signed-distance machinery, FE validation and report are reused. |
| `continuous/reference/feasibility_study.md` | Why the implicit-field approach was chosen over B-rep offsets (OpenCascade offsets failed on all three generative-design parts). |
| `continuous/reference/gpu_analysis.py`, `analysis.py` | Owen's GPU finite-element stack: matrix-free PCG with geometric multigrid, linear buckling with adjoint sensitivities. `cshell` must use it for all field analyses (Section 4, Appendix A). |
| `continuous/reference/interfaces.py` | Minimal stub so `analysis.py` imports. The original module was not supplied. |
| `continuous/reference/baseline_optimiser.py` | Owen's report-writing conventions (PDF post-processing), which `vshell/report.py` already follows. |
| `vshell/` (repository root) | Working discrete toolchain: STEP import, face selectors, SDF grid, mounts, voxel FE, report. Reuse, don't rewrite. |
| `examples/*.json`, `examples/step/*.step`, `examples/runs/*/` | The three Study 1 parts, their configs, and the vshell results to compare against. |

Conventions: **must** marks a requirement and **should** a strong recommendation. Section numbers (§) are stable, so code docstrings should cite them (e.g. `# THEORY §6.3`).

---

## 1. Background and goal

### 1.1 Physical evidence (manuscript)

Study 1 shelled three Fusion generative-design (GD) parts uniformly at constant mass, using an outward offset with escape holes. The parts were MJF PA11 (E = 1.76 GPa, ν = 0.42, ρ = 1.04 g/cm³), each loaded at 200 N in the FE model. Measured changes against the solid parts:

| Part | Designed wall (mm) | As-built wall (mm) | Stiffness change | Load-capacity change |
| --- | --- | --- | --- | --- |
| Cantilever | 1.756 | 1.555 | +3.5 % (n.s.) | **+61.1 %** |
| L-bracket | 1.700 | 1.486 | −7.0 % (n.s.) | **−11.2 %** |
| MBB | 2.100 | 1.893 | −14.9 % (n.s.) | **+14.8 %** |

Findings `cshell` must respect:

- **F5. Region type matters.** Long struts gain from hollowing. Short joints and cross-section transitions lose, because transverse shear needs material near the neutral axis.
- **F3. Thin joint walls fail locally.** Local shell buckling or collapse of thin joint walls drove the L-bracket loss. Walls printed about 0.2 mm thinner than designed made it worse (Table 5).
- **F7. Thin-wall FE over-predicts stiffness.** Any FE-in-the-loop optimiser leans hardest on its least accurate predictions, so the robust (as-built) formulation of §3.9 and body-fitted validation of §8.6 are not optional.
- **Section 6.2** recommends selective hollowing driven by a mechanical measure. vshell implements that discretely; `cshell` makes it continuous.

Study 2 (PA11 and aluminium tubes in 3-point bending) supplies the material calibration: σ_f = 36.6 MPa, back-calculated from the solid PA11 rod. See vshell theory, "Strength".

### 1.2 What vshell does and where it stops

vshell segments the part into mounts, joints and struts. Each candidate region is either solid (θ = 1) or hollow (θ = 0). One scalar outer offset s, solved by Brent's method, restores the mass, and s is also every hollow region's wall. It searches a ranked one-parameter family, at most 10 FE solves at 0.8 mm cells, and accepts a hollow design only if it clearly beats the solid.

| Part | Solid K₀ (N/mm) | Solid P₀ (N), mode | vshell best | K/K₀ | P/P₀ | Decision |
| --- | --- | --- | --- | --- | --- | --- |
| Cantilever | 257 | 431, Euler | 2 compression struts hollow, wall 1.91 mm | 0.992 | 2.63 | hollow |
| L-bracket | 48.3 | 540, strength | 14/15 regions hollow, wall 1.81 mm | 1.068 | 0.944 | solid kept |
| MBB | 473 | 2742, strength | 1 region hollow | 1.000 | 1.000 | solid kept |

Its limitations, which `cshell` removes:

1. **One wall thickness for all hollow regions, and an all-or-nothing state per region.** No thickness variation within a region; no thicker walls where stress is high.
2. **Segmentation dependence.** Skeleton junctions can be spurious, and the design space is only as good as the regions.
3. **Heuristic capacity.** Euler buckling of isolated struts (with an assumed K_eff), NASA and plate formulas for local buckling, and a stress quantile. No real eigen-buckling analysis, and no frame sway modes.
4. **Coarse FE.** 0.8 mm cells put only about 2 cells through a 1.8 mm wall; thin-wall stiffness noise is about ±6 %.
5. **No gradients.** It cannot exploit the information in the FE fields.

### 1.3 Goal

`cshell` must take a solid STEP part, its supports and its loads, and return a part of **exactly the same mass**. Its wall thickness varies **continuously over the surface**, built both into the original volume (void wall moved inward) and out of it (outer skin moved outward). The thickness is chosen by **gradient-based optimisation on GPU finite-element fields** to maximise stiffness and load capacity (strength and real eigen-buckling). Wherever hollowing does not pay, the optimiser closes the void locally (solid joints) or everywhere (the solid part is returned with the reason).

Hard requirements:

| ID | Requirement |
| --- | --- |
| R1 | **Constant mass.** Every iterate the optimiser reports satisfies \|V/V₀ − 1\| ≤ 1·10⁻⁴ on the geometry grid (v1.2: was 10⁻³, unified with §7.2), and the exported mesh satisfies \|V/V₀ − 1\| ≤ 2·10⁻³ after the final correction (§8.4). V₀ is the original part's volume on the same grid. |
| R2 | **Stay solid unless hollowing earns its place.** Use vshell's acceptance rule (§5.4), checked at two FE resolutions. Otherwise output the original part and the reason. |
| R3 | **Minimum-stiffness bolt-hole boundary condition.** All DOF of the support faces are clamped (agreed with Owen). Other BC models are optional extras. |
| R4 | **Mounts (bolt lugs, load pads, any user-listed preserve faces) are never changed.** |
| R5 | **Manufacturable for MJF.** Minimum nominal wall t_min = 1.5 mm, as-built bias 0.2 mm, every void at least v_min = 3 mm across, with solid always allowed instead (no slivers), escape holes for every enclosed void, user keep-out volumes respected. |
| R6 | **Every gradient is verified** against finite differences in the test suite (§6.6). |
| R7 | **Every run writes a PDF report and full provenance**, in the style of vshell and `baseline_optimiser.py` (§9). |
| R8 | **Runs headless on the cluster with GPU,** and the test suite runs on a CPU-only node at small sizes. `analysis.py`'s `cpu_direct` engine already does this (Appendix A.4). |

Owen cannot run physical tests at present, so validation is numerical (§11) plus comparison with the manuscript's measured trends.

---

## 2. Notation and units

Internal units are **mm, N, MPa (N/mm²), mm³, N·mm**, as in vshell. The adapter to `gpu_analysis`/`analysis` converts to **SI (m, Pa, N, J)** and back (§4.3).

| Symbol | Meaning |
| --- | --- |
| Ω₀, Γ₀ = ∂Ω₀ | original solid and its boundary |
| φ(x) | signed distance to Γ₀ (mm), **negative inside** |
| V₀ | volume of Ω₀ measured on the geometry grid (the mass target; M₀ = ρ V₀) |
| Γ_h, p_i, A_i | design surface mesh (remeshed Γ₀), vertex i, lumped vertex area |
| ξᵈ, ξᵃ ∈ [0,1]^{N_v} | design variables per vertex (outer-offset and inset fields) |
| F | Helmholtz (Laplace–Beltrami) filter on Γ_h |
| d(p), a(p) | outward offset of the outer skin; inward inset of the void wall, both measured from Γ₀ along the normal (mm) |
| w = a + d | local wall thickness |
| r_MA(p) | inward distance from p to the medial axis (local inradius, "depth") |
| r_out(p) | half the outward clearance from p to another part of Ω₀ |
| ψ(x) | design level-set: material ⇔ ψ ≤ 0 |
| π(x) | closest point on Γ₀; P is the matrix interpolating vertex values at π(x) |
| H_ε | smoothed Heaviside of half-width ε |
| h_g, h_e = k h_g | geometry-grid spacing; FE element size (k ∈ {2, 4}) |
| χ_c, ρ_e | material fraction of geometry cell c; of FE element e |
| ρ̂_e = ρ_min + (1 − ρ_min) ρ_e | stiffness density handed to the FE (E_e = ρ̂_e E) |
| b_v, b_o | as-built bias on the void side and on the outer side (mm) |
| C_k, K_k = \|F_k\|²/C_k | compliance and stiffness of load case k |
| σ_PN | p-norm von Mises stress; λ_s = σ_f/σ_PN the strength load factor |
| μ_j = 1/λ_j, μ̃ | inverse buckling load factors and their p-norm aggregate |
| λ_b = κ_b/μ̃ | buckling load factor with knock-down κ_b |
| λ | capacity load factor, soft-min of λ_s and λ_b |
| J | objective, log-ratio score against the solid (§5.1) |

---

## 3. Geometry representation

### 3.1 Signed distance

The STEP solid is read with OpenCascade (`vshell.geometry.load_step`) and tessellated watertight at 0.02 mm chordal tolerance (`vshell.geometry.tessellate`). The result is (V, T) plus the B-rep face id of every triangle. φ is the exact signed distance to that triangle mesh, signed by generalised winding number. vshell's `PartGrid` does this on a dense grid with an exact band. `cshell` needs φ only in a **band**:

$$\mathcal B = \{x : -\max_p r_{MA}(p) - h_g \le \varphi(x) \le D_{comp} + h_g\}$$

In practice this is the whole interior plus an outer shell of thickness D_comp. D_comp is the **computational** outward extent: 4 mm initially, extended automatically as in §3.5. It is **not a design limit**: there is no ceiling on outward growth (§14). It should be stored sparsely as a list of active cells (flat indices into the bounding grid) with φ, the closest-triangle id and barycentric coordinates per cell. Section 4.2 has the sizes.

Closest-point queries for about 10⁸ points should run on the GPU, using NVIDIA Warp `wp.mesh_query_point_sign_winding_number` or an equivalent BVH kernel. They may run on CPU with libigl (`igl.signed_distance` returns closest points and face ids) parallelised over chunks. They are computed **once per part and resolution** and cached to disk (`.npz`, keyed by STEP SHA-256 and grid parameters).

### 3.2 Two offset fields and the material set

Each point of Γ₀ carries two non-negative lengths:

- d(p), the **outward offset** of the outer skin;
- a(p), the **inward inset** of the void wall.

Both are measured along the normal from the original surface. The designed part is

$$\Omega = \{x : -a(\pi(x)) \le \varphi(x) \le d(\pi(x))\}, \qquad \psi(x) = \max\big(\varphi(x) - d(\pi(x)),\; -\varphi(x) - a(\pi(x))\big), \qquad \Omega = \{\psi \le 0\}.$$

Cross-section of one strut of original radius r = r_MA:

```
   outside        outer skin         void wall              void                void wall   outer skin
  ----------|----------------|====================|  . . . . . . . . . . .  |=============|-----------
         phi = d            phi = 0             phi = -a          (phi = -r at the axis)
            <-- d -->        <------ a ------->
            <------------ wall w = a + d ------->
```

- **Uniform shell of the current add-in:** a = 0, d = t. The original solid becomes the void.
- **Solid at original size:** d = 0 and a ≥ r_MA. The void wall meets itself at the medial axis and the void closes.
- **Solid and grown:** d > 0 and a ≥ r_MA.
- **Hollow:** a < r_MA. The void is the set of points deeper than a below Γ₀, and its thickness across a strut is about 2(r_MA − a).

Only d matters outside Ω₀ (φ > 0, where the first branch of ψ is active), and only a matters inside (φ < 0). So the outward extension of d and the inward extension of a never interact.

### 3.3 Why both fields are needed

With only an outward offset d (the add-in), any void forces d > 0 to restore mass, and no region can be solid. With only an inset a, any void loses mass that cannot be recovered. Constant mass with *locally* solid regions needs both:

- material removed inside (a < r_MA) is put back outside (d > 0) where it helps most;
- the mass constraint couples the two globally, not locally.

So a joint can close (a ≥ r_MA, d = 0) while a strut elsewhere hollows and grows (a = t_min, d = 2 mm). Within one strut, both fields can vary along its length and around its circumference.

### 3.4 Design surface, variables and filter

The design lives on a **design surface mesh** Γ_h, independent of every analysis grid. Γ_h is an isotropic remesh of the tessellation with target edge ℓ_Γ = 0.75 mm (configurable 0.5–1.0). Use gpytoolbox `remesh_botsch`, pymeshlab isotropic remeshing, or similar. Check that its Hausdorff distance to the tessellation is ≤ 0.02 mm, and keep the B-rep face id of each vertex (nearest triangle of the tessellation). The cantilever, with about 2.5·10⁴ mm² of surface, gives about 6·10⁴ vertices at 0.75 mm.

Design variables are two vertex fields ξᵈ, ξᵃ ∈ [0,1]^{N_v}. They are smoothed with the Helmholtz (PDE) filter on the surface (Lazarov & Sigmund 2011; the surface version is the Vertex Morphing idea of Hojjat et al. 2014):

$$(\mathbf M + r_f^2\,\mathbf L)\,\tilde\xi = \mathbf M\,\xi \quad\Longrightarrow\quad \tilde\xi = F\xi,\quad F = (\mathbf M + r_f^2 \mathbf L)^{-1}\mathbf M$$

Here **L** is the cotangent Laplace–Beltrami stiffness (`igl.cotmatrix`, sign so that it is positive semidefinite) and **M** the lumped (Voronoi) mass matrix. The filter radius r_f defaults to 3.0 mm, with continuation 4 → 3 → 2.5 mm (§7.3). Factorise (M + r_f² L) once per r_f (CHOLMOD or `scipy.sparse.linalg.factorized`). With a lumped M and non-negative cotangent weights the filter obeys a discrete maximum principle, so ξ̃ ∈ [0,1]. Clip anyway and check. The adjoint is Fᵀ = M (M + r_f² L)⁻¹, so apply it with the same factorisation.

Physical fields, per vertex:

$$d_i = d^{hi}_i\,\tilde\xi^d_i, \qquad a_i = a^{lo}(d_i) + \big(a^{hi}_i - a^{lo}(d_i)\big)\,\tilde\xi^a_i \quad\text{(v1.2: } a^{lo} \text{ depends on } d\text{, §14 D1)}$$

| Bound | Value | Meaning |
| --- | --- | --- |
| a^{lo}(d_i) | w_s·softplus((t_min − d_i)/w_s), w_s = 0.1 mm (v1.2, Owen's D1) | smoothed (t_min − d)⁺: w = a + d ≥ t_min always, and the wall may lie wholly outside Γ₀ once d ≥ t_min (the add-in's a = 0 shell is in the design space) |
| a^{hi}_i | r_MA,i + γ_c ε, with γ_c = 0.75 | void just closed, §3.8 |
| d^{hi}_i | min(r_out,i − c_gap, keep-out distance − c_K, D_comp) | §3.5. Only the clearance terms are design limits. |

Here c_gap defaults to 1 mm and c_K to 0.5 mm. ξ̃ᵃ = 1 is "locally solid", like vshell's θ. ξ̃ᵈ = 0 is "no outward growth".

**Passive vertices.** Vertices within `mounts.dilate_mm` (1.5 mm) of a mount or BC face are fixed at ξᵃ = 1, ξᵈ = 0. vshell's `segment.py` computes the mount zones from analytic B-rep faces and support and load faces; reuse it, keeping only the mount part of the segmentation. So are vertices where a^{hi} ≤ max(t_min, v_min/2) + 0.05 mm (v1.2, with the D1 bound: a void ≥ v_min needs r_MA ≥ v_min/2 + a with a → 0 once d ≥ t_min), where the part is too thin to host a void wider than v_min. The passive values are imposed **after** filtering, and their gradient is zero. The filter still lets them blend smoothly into active neighbours, because active vertices see the passive neighbours' ξ through F.

### 3.5 Local depth, clearance and keep-outs

All three are per-vertex quantities computed once from φ and the mesh.

- **r_MA(p_i), local inradius.** Along the inward normal, D(t) = −φ(p_i − t n_i) equals t until the medial axis is reached, then falls below it (on a slab, D = 2r − t beyond the axis). Evaluate D with **exact point-to-mesh distance queries, not grid interpolation**:
  1. March in steps of h_g/2 to bracket the first t with D(t) < t − δ, δ = 10⁻³ mm.
  2. **Bisect** the bracket to 0.01 mm. The detection bias is about δ/2, which is negligible.

  A coarse threshold such as δ = h_g/4 with h_g/2 steps overestimates r_MA by 0.1–0.6 h_g. That pushes the axis cells out of the closure window of §3.8 (H′ = 0) and freezes closed regions, so the bisection is required, not optional. The half-chord to the opposite surface (ray cast) is an upper bound, used only as a sanity check. Test: on a cylinder or slab of half-thickness r the result must be r ± 0.1 h_g.
- **r_out(p_i), exterior clearance.** Cast a ray along +n_i. If it hits Γ₀ at distance g, r_out = g/2; otherwise ∞. This stops two members from fusing when both grow outward, which would change the part's topology.
- **Keep-out volumes,** given as STEP/STL bodies or primitives in the config. Each has its own SDF φ_K. March along +n_i to the first point where φ_K < c_K, the clearance (default 0.5 mm); that distance bounds d^{hi}_i. None is needed for the Study 1 parts.
- **No ceiling on outward growth (Owen's decision).** Clearance to other members (r_out − c_gap) and to keep-out bodies are the only design limits on d. Mass conservation limits growth in practice: material added outside must be removed inside. The analysis domain still needs a finite extent, D_comp, which starts at 4 mm. After each iteration, check for any active vertex whose d^{hi} is set by D_comp (not by clearance) and which has d_i > 0.8 D_comp. If one is found:
  1. Double D_comp and rebuild the band, P and the FE grid. Check memory first (§4.2).
  2. Re-map ξᵈ so that d is unchanged: target ξ̃ᵈ = d/d^{hi}_new, then ξᵈ = M⁻¹(M + r_f² L) ξ̃ᵈ, clipped to [0, 1]. This inverts the filter.
  3. Restore mass (§7.2) and reset the MMA asymptotes.
  4. Log the extension. The report shows the final D_comp and the d distribution.
- If `design.allow_outward = false` then d^{hi} ≡ 0. Constant mass then admits only the solid. The tool should say so and return the solid at once.

### 3.6 Extension from the surface into the volume

For every geometry cell centre x in the band, π(x) is the closest point on the fine tessellation (from §3.1). That point is then located on Γ_h (closest triangle t′ and barycentric β′), which gives a fixed sparse interpolation matrix

$$d(x_c) = \sum_{j \in t'(c)} \beta'_{cj}\, d_j, \qquad \text{i.e. } \mathbf d_{grid} = \mathbf P\,\mathbf d, \quad \mathbf a_{grid} = \mathbf P\,\mathbf a,$$

with three non-zeros per row. Γ₀ never moves, so **P is computed once**. All design dependence is in the vertex values. The gradient pull-back is Pᵀ, a sparse transpose product.

Points beyond the medial axis project onto the opposite face. The void wall there uses that face's a, which is the intended behaviour: each face of a strut has its own wall. Where π jumps (on the medial axis) both sides give depth r_MA. So the void boundary is continuous wherever a < r_MA on both sides, and closes cleanly when either reaches a^{hi}.

### 3.7 From ψ to material fractions

On the geometry grid (h_g = 0.1 mm production, 0.2 mm coarse), each cell's material fraction is a smoothed Heaviside of the level-set value at its centre:

$$\chi_c = H_\varepsilon(-\psi_c), \qquad H_\varepsilon(s) = \begin{cases} 0 & s \le -\varepsilon\\ \tfrac12 + \tfrac{15}{16}\big(\tfrac{s}{\varepsilon}\big) - \tfrac58\big(\tfrac{s}{\varepsilon}\big)^3 + \tfrac{3}{16}\big(\tfrac{s}{\varepsilon}\big)^5 & |s| < \varepsilon \\ 1 & s \ge \varepsilon\end{cases}$$

with ε = h_g. This C² quintic has ∫H′ = 1, so volume moves at exactly the surface area per unit normal displacement. vshell's linear ramp clip(½ − ψ/h, 0, 1) is the C⁰ special case. It may be used for volume bookkeeping, but its kinks add noise to gradients, so the optimiser should use the quintic. The smoothed Heaviside is used only at a scale of one geometry cell; it is not a design-relaxation device.

Each FE element is a k³ block of geometry cells (k = h_e/h_g: 4 at h_e = 0.4 mm, h_g = 0.1 mm):

$$\rho_e = \frac{1}{k^3}\sum_{c \in e} \chi_c, \qquad \hat\rho_e = \rho_{min} + (1-\rho_{min})\,\rho_e, \quad \rho_{min} = 10^{-6}.$$

This is the iso-strain (Voigt) ersatz material used and validated in vshell. Within a cut element it is right for axial and bending stiffness in the mean, with a second-order surface error. With q1_im elements it reproduced CalculiX C3D8I to 4·10⁻⁵ on identical meshes, and the Study 2 tube stiffnesses within 4 % of Abaqus at h_e = 0.8 mm. h_e = 0.4 mm puts 3.75 elements through a 1.5 mm wall.

### 3.8 Void closure, and why it can reopen

With exact geometry, the derivative of every response with respect to a is **zero** once the void has closed (a ≥ r_MA). Opening a void is a topological change, so a closed region would stay closed for the rest of the run: the level-set nucleation problem.

The upper bound a^{hi} = r_MA + γ_c ε is chosen so that the medial-axis cells of a fully "solid" vertex sit inside the Heaviside ramp at ψ ≈ −0.75ε. There:

- χ = H(0.75) ≈ 0.984, a 1.6 % deficit in a one-cell layer, about 0.4 % in the FE elements containing it;
- H′ ≈ 0.19 of its peak.

So the region is solid for all practical purposes, but the gradient stays alive. That gradient is the energy density at the medial axis, i.e. on the neutral axis. It approximates the topological sensitivity of opening a void there:

- **shear-dominated joints** have high energy density at the axis and stay closed;
- **bending-dominated struts** have low energy density there and are pushed to open. So are compression struts once the buckling term is active: under compliance alone an axially loaded strut, in tension or compression, has uniform energy across its section and is neutral.

This is finding F5, emerging from the gradient rather than a heuristic.

Do not raise γ_c above about 0.9: H′ → 0 and closed regions freeze. Do not lower it below about 0.6: the axis layer weakens visibly. Unit tests must check that:

1. the all-solid design ξᵃ ≡ 1, ξᵈ ≡ 0 reproduces V₀ within 5·10⁻⁴;
2. its compliance is within 0.5 % of the true solid's.

Errors in r_MA of about h_g/4 move the axis cells within the ramp, which is harmless. The **solid baseline** used for decisions is always evaluated directly from ψ = φ, not through the fields, so it is exact. The export (§8) fills closed regions exactly.

An optional **reopening probe**, at the start of each continuation phase: for closed vertices whose medial-axis energy density is below the current Lagrange multiplier of the mass constraint (§5.3), set ξᵃ to 0.9. This nudges them back into the ramp so the gradient can decide.

**Minimum void size, and the solid option.** v_min = 3.0 mm (Owen's decision, §14). A void, wherever one exists, must be at least 3 mm across. Sliver voids (0 < 2(r_MA − a) < v_min) cannot be de-powdered and act as cracks. But "no void" is always allowed: **the optimiser must remain free to close any region to solid**, and to open one, on its own judgement. The rule is therefore an either/or at every point (void ≥ v_min, or solid), never a minimum void that must exist.

How each piece of the formulation respects this:

1. **Bounds.** At every active vertex the bounds always include the closed state (a up to a^{hi} > r_MA) and the open state. No bound or constraint may ever force a region to stay open.
2. **Gradient at closure.** The closure window of §3.8 keeps the gradient alive at the closed state, so a closed region can reopen and an open one can close.
3. **Sliver penalty.** It snaps each vertex towards the nearer admissible state (§5.2). It is zero both in the closed state and for voids ≥ v_min, so it never favours staying open.
4. **Continuation.** The penalty is off in phase A and ramps up afterwards (§7.3). Early on the optimiser can move regions between open and closed freely; the penalty only tidies up the intermediate zones at the end.
5. **Void ends.** The tapering end of a void, where an open region meets a closed one, necessarily passes through thicknesses below v_min. That is acceptable during optimisation. At export the void is morphologically opened with a ball of diameter v_min (§8.2), which turns every tapering tip into a rounded end of radius ≥ v_min/2 and fills anything thinner. The mass is then corrected (§8.4).

The exported part therefore satisfies the rule exactly: every void is a union of v_min balls, and every other region is solid.

A consequence: wherever a^{hi} ≤ t_min + v_min/2 = 3.0 mm (§3.4 passive rule), no void of at least v_min fits inside a wall of at least t_min. Roughly, members thinner than about 6 mm stay solid automatically. The report must show this passive area.

### 3.9 As-built geometry (robust formulation)

MJF walls came out about 0.2 mm thinner than designed (Table 5). Walls adjacent to a powder-filled void are the ones that lose material. The **as-built** design used for every mechanical response is

$$a^{ab}_i = a_i - b_v\,S\!\Big(\frac{r_{MA,i} - a_i}{b_v}\Big), \qquad \psi^{ab} = \max\big(\varphi - d + b_o,\; -\varphi - a^{ab}\big),$$

where S(z) = smoothstep(clip(z, 0, 1)) = 3z² − 2z³ and b_v = 0.2 mm.

- **Open void:** the wall loses b_v.
- **Closed void (a ≥ r_MA):** S = 0 and nothing changes, so no void reappears inside solid regions. Key the ramp on r_MA, not a^{hi}: the window a^{hi} − r_MA = γ_c ε (0.075–0.15 mm) is narrower than b_v, so keying on a^{hi} would erode a nominally closed void open again.
- **b_o** is an optional outer-surface bias, default 0. If set non-zero it also applies to the solid baseline, for a like-for-like comparison.

The **mass constraint uses the nominal design** (that is what the CAD defines). The report gives the as-built mass estimate too. This is a one-sided robust formulation (Wang, Lazarov & Sigmund 2011), since erosion only weakens the part. A two-design (nominal plus eroded) worst case is an option for later.

### 3.10 Volume and mass

$$V(\xi) = h_g^3 \sum_{c \in \mathcal B} \chi_c \;+\; V_{\text{interior, not in }\mathcal B}$$

The second term is zero when the band covers the interior. Sum in float64. V₀ is the same sum for ψ = φ. Mass is ρV, with ρ = 1.04·10⁻³ g/mm³.

---

## 4. Mechanics model

### 4.1 Finite-element discretisation

- **Grid.** Linear elasticity on a regular hexahedral grid of spacing h_e. The grid covers the bounding box of the band (Ω₀ grown by D_comp plus a one-element pad). It grows if D_comp is extended (§3.5).
- **Material.** Isotropic PA11: E = 1760 MPa, ν = 0.42. Element e has stiffness density ρ̂_e (§3.7), and `analysis.ConjugateGradient._volume_mesh` scales both Lamé parameters by the field it is given. So **pass ρ̂ (not ρ) as the field**; it must already contain the floor ρ_min.
- **Element.** `element_type = "q1_im"`, the Wilson–Taylor incompatible-modes hexahedron, which is the CalculiX/Abaqus C3D8I formulation that vshell validated. Plain trilinear elements lock in bending of thin walls and must not be used for production.
- **Solver.** `preconditioner = "gmg"` and `engine/pcg_engine = "auto"`. `gpu_analysis.probe()` picks the engine for the hardware.

> **v1.2 — FE engine (Owen, 28 October 2026; STATUS.md §2.2).** Production FE uses `cshell/fe_sparse.py`, not the dense
> box: only the **active set** of elements (ρ_e > 0, dilated by `grid.active_margin_mm`) is assembled, with the same
> q1_im element matrix as the reference (`gpu_analysis.element_khat`), and factorised by **cuDSS** sparse Cholesky
> (fp64). One factorisation per design serves the primal, the stress adjoint (§6.3) and the buckling eigen-solve and
> its adjoint (§6.4), so **buckling runs at the design resolution (0.4 mm)**: the 0.8 mm buckling grid and its chained
> sensitivity (item 2 below) are no longer used. Dropping the void elements changes K by O(ρ_min) and no design
> derivative (∂χ/∂ψ = 0 away from the surface); the active set is rebuilt whenever the design puts material outside it.
> The reference `ConjugateGradient` (via `cshell.fea.ShellCG`) remains the test reference. Measured costs:
> `continuous/BENCHMARKS.md`. A single fp64 factor holds ≈ 7 M DOF on a 96 GB GH200 (≈ 9 kB/DOF), so 0.2 mm runs
> need cuDSS hybrid memory or an iterative solver (STATUS K12).

### 4.2 Resolution, problem size and memory

The parts are sparse in their bounding boxes: material fills 1.7–4 % of the box. The table gives DOF counts for the dense box (what `gpu_analysis` solves today) and for the material alone (what a narrow-band sparse grid would solve). Box sizes are from the tessellated parts plus 4 mm padding each side.

| Part | V₀ (mm³) | Box (mm) | Dense DOF h_e = 0.8 / 0.4 / 0.2 mm | Material-only DOF 0.4 / 0.2 / 0.1 mm |
| --- | --- | --- | --- | --- |
| Cantilever | 60,942 | 168 × 148 × 148 | 22 M / 173 M / 1.38 G | 2.9 M / 23 M / 183 M |
| L-bracket | 41,948 | 132 × 188 × 86 | 13 M / 100 M / 0.80 G | 2.0 M / 16 M / 126 M |
| MBB | 85,233 | 308 × 118 × 58 | 12 M / 99 M / 0.79 G | 4.0 M / 32 M / 256 M |

Consequences:

1. **Stages 1–2 run on the dense grid** with h_e = 0.4 mm production and 0.8 mm coarse (continuation phase A). At 0.4 mm a 1.5 mm wall spans 3.75 elements, twice vshell's resolution.
   - **Memory model.** `gpu_analysis.elastic_memory("gpu_gmg", nx, ny, nz)` budgets GMG-PCG at 210 B/DOF × 1.15 safety (206 B/DOF measured on a GH200, the hardware Owen's code was calibrated on). With hierarchy reuse (`"gpu_gmg_lean"`) it is 150 B/DOF. `pcg_jax_memory` models only the JAX engine.
   - **What fits.** Owen's cluster has NVIDIA GH200s. Query the HBM size with `nvidia-smi`: 96 GB or 144 GB, depending on the variant. `gpu_analysis` treats 0.85 of free VRAM as usable.
     - The cantilever at 0.4 mm (174 M DOF) needs about 42 GB and fits either variant.
     - At 0.3 mm (411 M DOF) it needs about 99 GB, which fits only the 144 GB variant.
     - Extending D_comp (§3.5) enlarges the box, so re-check.
   - **Check before every phase.** Run the check and record the result.
2. **Buckling at 0.4 mm dense does not fit.** `gpu_analysis.buckling_memory(engine, nx, ny, nz, n_free, num_modes)` for `gmg_lobpcg32` counts:
   - 1536 B per element of element data (88 GB alone for the cantilever's 57 M elements at 0.4 mm);
   - LOBPCG blocks of 12 × (m + 4) vectors.

   At 0.8 mm the estimate is about 40 GB, which fits. So run buckling on the 0.8 mm grid of the same design (ρ̂ block-averaged 2³ → 1) and chain the sensitivity back through the averaging (§6.4). Query `buckling_memory` before every run, and report which grid and engine buckling used.
3. **Stage 3 builds a narrow-band block-sparse FE grid** (active 8³-element bricks only), to reach h_e = 0.2 mm (and 0.1 mm for validation). That is a 20–60× DOF saving. See Liu et al. 2018 (SPGrid narrow-band TO, about 10⁹ voxels on one workstation) and Aage et al. 2017 (giga-voxel). Coarse levels can be built by brick agglomeration, keeping the matrix-free element kernels.
4. **Crop the box** to {φ < d^{hi} + h_e} in every case, and never let the optimiser touch elements outside it.

Measure the cost per iteration (one primal solve per load case, one stress adjoint per load case, buckling) early in Stage 0 and record it in `continuous/BENCHMARKS.md`. Estimate run times from that, not from guesses.

**As built (2026-10-30; supersedes items 1–3 for the FE).** The FE runs on the **active set** only (elements within `active_margin_mm` of material), not the dense box, at h_e = 0.8 mm (phase A) and 0.4 mm (phase B): 0.9–1.3 M DOF and 6.1–8.7 M DOF for the hollow designs. Default engine `fe.engine = "twogrid"` (`cshell/fe_twogrid.py`) in every phase, on one GPU: matrix-free q1_im K at h_e (no assembled fine matrix), Galerkin coarse operator at 2h_e factorised by cuDSS (fp32), Chebyshev–Jacobi smoothing, fp64 block PCG to 10⁻¹⁰ relative residual; buckling (§4.4) by block-Krylov on K⁻¹G with the same PCG as K⁻¹, so buckling runs at h_e = 0.4 mm on the design grid, without the 0.8 mm averaging of item 2. Measured on MBB phase B (8.4 M DOF): 25–50 s per iteration, 26 GiB GPU, ≈ 34 GB host; agreement with the direct cuDSS fp64 engine (`fe.engine = "direct"`, kept for validation): dJ ≤ 5·10⁻¹⁰ without buckling, J ≤ 6·10⁻⁸ over run histories; the full dJ differs by ~10⁻⁴ through eigenvector truncation at eig_tol = 10⁻³ in either engine. Numbers: `continuous/BENCHMARKS.md` "Two-grid engine".

### 4.3 Boundary conditions, loads and the analysis adapter

`cshell` defines its BCs exactly as vshell does (§1.3 R3):

- **Supports.** Every FE node within 0.75 h_e of a support-face triangle is clamped in x, y and z.
- **Loads.** Each load's total force is shared equally among the nodes within 0.75 h_e of its load faces.
- **Overlap.** Loads on clamped DOF are removed with a warning.

Supports and load faces are mount faces, so they never move (§3.4) and the BCs are design-independent.

`analysis.ConjugateGradient` takes **element-wise** CSG arrays: `csg_domain_data = {"fixed_nodes": bool (nx_e, ny_e, nz_e, 3), "load_vectors": float (nx_e, ny_e, nz_e, 3)}`. `_volume_mesh` smears them to the 8 element corners, with the load divided by 8. That clamps a one-element slab rather than a node set, which is slightly stiffer than vshell.

`cshell` should instead **subclass** `ConjugateGradient` (`cshell.fea.ShellCG`) and override `_volume_mesh(rho)` to use node-level arrays: a fixed-node mask and a nodal load vector. Requirements:

- **Return the same 11-tuple** `(nx, ny, nz, hx, hy, hz, E_elem, nu_elem, b, fixed_mask, rho_flat)` with the same dtypes:
  - `fixed_mask` is **int32 with one entry per DOF** (length 3·N_nodes, flattened from an (n_x, n_y, n_z, 3) array). The method's docstring wrongly says "(Nnodes,)".
  - `b` is float32, length 3·N_nodes.
  - `E_elem` and `nu_elem` are float32 per element.
- **Keep passing element-shaped `csg_domain_data`** (zeros of shape (n_x,e, n_y,e, n_z,e, 3) are fine) to `__init__`. It raises without one, and `_apply_elastic_engine` sizes the grid from `csg_data["fixed_nodes"].shape`.
- **Test every engine.** `_run_analysis_cpu` also calls `_volume_mesh`, so test both the GPU and `cpu_direct` paths through the override.

`run_analysis` and `run_buckling` read the `_last_*` state, so they then work unchanged.

Keep one `ShellCG` instance per load case. Each holds its own warm-start state (`warm_start_cg=True`, `warm_start_state` dict).

Conventions to respect (Appendix A):

- **Ordering.** C-order arrays, so node id = i·(n_y n_z) + j·n_z + k for node counts (n_x, n_y, n_z). DOFs are interleaved [u_x, u_y, u_z] per node.
- **Voigt order.** Stresses come as **(xx, yy, zz, yz, xz, xy)** with engineering shear strain. vshell uses (xx, yy, zz, xy, yz, zx), so convert explicitly and test it.
- **Units.** `youngs_modulus` in Pa (1.76·10⁹), `voxel_size_mm` = h_e, forces in N. Compliance comes back in J = N·m; multiply by 10³ for N·mm. `voxel_sed` comes back in J/m³; multiply by 10⁻⁶ for N·mm/mm³ (MPa). Element energy is U_e = sed_e·h_e³.

### 4.4 Responses

All responses are evaluated on the **as-built** design ψ^{ab} (§3.9), for each load case k.

**(a) Compliance and stiffness.**

$$C_k = \mathbf f_k^{\mathsf T}\mathbf u_k, \qquad K_k = \lVert F_k\rVert^2 / C_k$$

The ratio to the solid, C₀/C, is the stiffness ratio. It stays defined when loads partly cancel, as in vshell.

**(b) Strength.** Material stress in element e is σ^{mat}_e = D⁰ ε_e (unscaled constitutive matrix). `gpu_analysis.compute_element_stress` returns ρ̂-scaled element-averaged stresses, so either divide by ρ̂_e or recompute from u. For q1_im the incompatible-mode gradients vanish at the element centre, so the **centre** strain B(0)u_e is exact. Confirm which point `compute_element_stress` uses and be consistent with the adjoint (§6.3). With s_e the von Mises value of σ^{mat}_e:

$$\sigma_{PN} = \Big(\frac{\sum_e v_e\, s_e^{\,p}}{\sum_e v_e}\Big)^{1/p}, \qquad v_e = \rho_e^{\,q}\, m_e$$

- p = 8 (continuation 6 → 8 → 10).
- q = 2 relaxes the weight of partially filled surface elements, whose strain is unreliable.
- m_e = 0 within `exclude_bc_mm` = 2 mm of any fixed or loaded node, and in mounts unless `include_mounts`. Mounts never change; their stress is reported separately, as in vshell.

The p-norm is a smooth stand-in for the 0.995 volume-weighted quantile that vshell used as "the" stress. It is calibrated to that quantile by the adaptive normalisation of Le et al. (2010):

$$\lambda_s = \frac{\sigma_f}{c\,\sigma_{PN}}, \qquad c^{(n)} = \Big(c^{(n-1)}\Big)^{1/2}\Big(\frac{Q_{0.995}\big(s \mid \rho_e \ge 0.5,\, m_e = 1\big)}{\sigma_{PN}}\Big)^{1/2}_{\text{iterate } n-1}$$

c is held fixed while gradients are evaluated. σ_f = 36.6 MPa, from vshell's calibration on the Study 2 solid rod.

**(c) Linear buckling.** Call `ShellCG.run_buckling(...)` right after `run_analysis` for the same load case, with:

| Argument | Value |
| --- | --- |
| `num_modes` | 6 |
| `aggregation_p` | 8 |
| `prev_modes` | modes of the previous iterate, for continuity |
| `min_density_for_geometric_stiffness` | 0.1, which suppresses spurious modes in void and nearly void elements |
| `engine` | default `"auto"` |
| `lobpcg_tol` | 10⁻³ during optimisation (eigenvalue error ~10⁻⁶); 10⁻⁶, or the `"classic"` engine, for finite-difference tests |

It returns `BucklingResults`: `mu_values` (μ = 1/λ, most critical first), `mu_tilde` (the p-norm aggregate, ≥ max μ) and `sensitivity` = ∂μ̃/∂(field), on the voxel grid.

$$\lambda_{b,k} = \kappa_b / \tilde\mu_k$$

κ_b = 0.85 is a single imperfection knock-down, configurable. The report must show the decision's sensitivity to κ_b over 0.6–1.0 (§9).

**Geometric stiffness excluded next to clamps (v1.3, Owen 2026-10-29).** Elements within `capacity.exclude_bc_mm`
(2 mm) of a clamped node carry no K_σ (their ramp r and r′ are multiplied by a constant 0/1 mask, so ∂μ̃/∂ρ̂ stays the
exact derivative of the model solved). Reason: where the design flares outward next to a support, the all-DOF clamp
(R3) leaves a concave corner with a singular stress, and a local mode there is mesh-dependent in both the voxel and
the body-fitted model (B2: body-fitted λ₁ 23.9 → 19.3 from h = 0.5 to 0.35 mm, `research/b2_bodyfit/diagnosis.md`).
The same rule as the stress exclusion zone (§4.4(b)); the body-fitted check (§8.6) applies it identically. It can hide
genuine local buckling within 2 mm of a support; `capacity.ksigma_exclude_bc = false` turns it off.

The eigen-analysis captures, in one model:

- global Euler modes of struts;
- frame sway modes, which vshell could not see;
- **local shell modes of thin walls** (F3), provided the wall is resolved. 0.4 mm is the minimum and 0.2 mm is preferred. Check the convergence of λ_b with h_e on the benchmark tubes (§11).

Optionally, classify each retained mode as global or local. A local mode puts more than 50 % of its modal strain energy in elements within 3 t_min of a void wall and has short-wavelength cross-sectional distortion. Local modes can get the NASA SP-8007 knock-down, γ = 1 − 0.901(1 − e^{−√(r/t)/16}), and global modes κ_b = 1. Ship the uniform κ_b first.

**(d) Capacity.**

$$\frac{1}{\lambda} = \Big(\sum_k \big(\lambda_{s,k}^{-\beta} + (\tilde\mu_k/\kappa_b)^{\beta}\big)\Big)^{1/\beta}, \qquad \beta = 12$$

Write the buckling term as (μ̃/κ_b)^β, never as λ_b^{−β} = (κ_b/μ̃)^{−β}. `run_buckling` returns μ̃ = 0 in two cases, and the inverse form stays finite in both:

- **No compressive mode exists,** e.g. tension-only loading. The term is 0.
- **Every mode was filtered out.** The buckling load is then unknown, yet `eigensolver_converged` is still True. Detect it as `n_candidates > 0 and mu_values.size == 0`, log it, and retry with more modes or a lower `min_energy_fraction`.

λ_b = κ_b/μ̃ is used only for reporting (∞ when μ̃ = 0).

This is the soft-min of all strength and buckling load factors over all load cases. The report also gives the hard min and the governing mode, load case and location: the peak element for strength, the peak modal amplitude for buckling. P = λ‖F‖ is the predicted capacity in newtons, for a single load case.

**(e) Load cases.** Only the **single** Study 1 load case is required (Owen's decision, §14). The sums over k in this document then have one term. Keep `loads` a list in the config, as vshell does, so several loads applied together still work. Do not build multi-case machinery (case weights ω_k, per-case instances) unless asked.

### 4.5 What the model does not capture

The model is linear elastic and small-strain, and its capacity values are proxies. Treat them as a ranking of designs, not certified failure loads.

- **Ductility.** PA11 elongates 35–50 % before failure. vshell's capacity sign disagreed with the MBB test (predicted −11 %, measured +14.8 %), most likely through plastic redistribution, in which a hollow section's larger plastic modulus raises the collapse load. A peak-elastic criterion cannot see that.
- **Optional Stage 3 item: plastic limit analysis.** Lower-bound limit analysis by second-order cone programming on the voxel mesh. Alternatively, run the plastic section check from the vshell theory doc ("extensions" item 1) on the final design as a post-check.
- **F7 (thin-wall FE over-prediction).** This is addressed by resolution, the as-built bias, and the body-fitted validation of §8.6. Absolute stiffness carries the boundary-condition uncertainty that the manuscript bounded at 147–151 % for the MBB. Ratios between designs carry much less.

---

## 5. The optimisation problem

### 5.1 Objective

For design variables ξ = (ξᵈ, ξᵃ):

$$\max_{\xi}\; J(\xi) = \frac{w_K \ln\!\big(C_{0}/C(\xi)\big) \;+\; w_P \ln\!\big(\lambda(\xi)/\lambda_0\big)}{w_K + w_P}, \qquad w_K = w_P = 0.5.$$

(v1.2: single load case, so no case weights ω_k; §14 D3.)

- C_{0,k} and λ₀ belong to the original solid, evaluated by the same pipeline (as-built, b_o).
- J = 0 is the solid.
- J is the log of a weighted geometric mean of the stiffness and capacity ratios, identical to vshell's J, so the results compare directly.
- **w_K and w_P are user-tunable hyper-parameters.** They are set in the config (`objective.w_stiffness`, `objective.w_capacity`) or on the command line (`--set objective.w_capacity=0.8`), and default to equal. Only their ratio matters.
  - No code change may be needed to change them.
  - They are recorded in `summary.json` and on the report's first page.
  - The acceptance rule (§5.4) uses the same weights, so the decision is consistent with what was optimised.
  - Setting w_K = 0 or w_P = 0 optimises a single metric.
- Optional: `python -m cshell sweep <cfg> --weights 0.2,0.5,0.8` runs the same part at several weight ratios and plots K/K₀ against P/P₀, the trade-off curve.

The MMA objective is f₀ = s_J(−J + w_s G_s), with s_J chosen so that |f₀| is about 10 at the first iterate. G_s is the sliver penalty (§5.2).

### 5.2 Constraints

| # | Constraint | Form | Notes |
| --- | --- | --- | --- |
| g₁ | Mass equality | V(ξ)/V₀ − 1 = 0 | Nominal geometry. Given to MMA as ±(V/V₀ − 1) − τ_V ≤ 0 with τ_V = 5·10⁻⁴, **and** enforced to 10⁻⁴ by the restoration step of §7.2 after every update. |
| g₂ (penalty, not a hard constraint) | Void either ≥ v_min or closed | Adds w_s·G_s to the minimised objective, with G_s(ξ) = (1/A_Γ) Σ_i A_i B(z_i) | z_i = (r_MA,i − a_i)/(v_min/2). B(z) = 16 z²(1 − z)² for 0 < z < 1, else 0. B is **zero both when closed (z ≤ 0) and when the void is ≥ v_min (z ≥ 1)**, so it pushes intermediate vertices to the nearer admissible state. It never makes the solid state costly. w_s continuation: 0 in phase A, then 0 → 0.05 → 0.1 (§7.3). A hard constraint is deliberately not used: it would trap regions in whichever state they reached first and block natural closure. The final guarantee comes from the export opening (§8.2). |
| — | Minimum wall | a ≥ t_min | Box bound via a^{lo} (§3.4). Since d ≥ 0, w ≥ t_min nominal and w^{ab} ≥ t_min − b_v. |
| — | Envelope, keep-outs, no fusing | 0 ≤ d ≤ d^{hi} | Box bound (§3.5). |
| — | Mounts, thin zones | passive vertices | §3.4. |
| g₃, g₄ (optional) | Floors | ln(K/K₀) ≥ ln f_K, ln(P/P₀) ≥ ln f_P | Off by default; the floors are acceptance tests (§5.4). Turn on to steer the optimiser away from designs the rule would reject. |

### 5.3 Optimality conditions and what they mean physically

For compliance alone, with the mass multiplier Λ, take the Lagrangian L = C + Λ(V − V₀). The discrete identity is

$$\partial C/\partial\hat\rho_e = -2\,U^0_e,$$

where U⁰_e = ½ u_eᵀ K_e⁰ u_e is the energy the element would store if fully solid at its current strain. ∂V/∂ρ_e = h_e³. So, at a free vertex i:

$$\text{outer skin:}\quad \overline W_{o,i} = \frac{\Lambda}{2(1-\rho_{min})}, \qquad \text{void wall:}\quad \overline W_{v,i} = \frac{\Lambda}{2(1-\rho_{min})}$$

W̄ is the solid-material strain-energy density averaged over the part of the moving surface that vertex i controls, weighted by H′ and the filter. The conditions at the bounds:

- **d_i = 0:** W̄_{o,i} ≤ Λ/2. The skin there is not worth growing.
- **a_i = t_min:** W̄_{v,i} ≤ Λ/2. The wall would thin further if it were allowed.
- **closed void:** W̄ at the medial axis ≥ Λ/2. The core carries enough energy to stay solid.

**At the optimum, every free moving surface carries the same energy density.** This is a boundary analogue of the fully stressed design, and it explains the expected outcomes:

- **Struts in bending.** Energy peaks at the outer fibres and is low near the axis. a goes to t_min and d grows where the moment is largest, so walls thicken toward high-moment ends.
- **Shear-dominated joints.** Energy is high at mid-depth (transverse shear peaks on the neutral axis). The inner wall moves inward until the void closes, so the joint stays solid (F5).
- **Axially loaded struts (tension or compression).** Energy is uniform across the section, so under compliance neither surface gains. For compression struts the buckling term decides: hollowing at constant mass roughly doubles or triples I_min (vshell's cantilever result), so λ_b drives them open (benchmark B3). Tension struts stay marginal, decided by the strength term and the filter.

With capacity terms, W is replaced by the corresponding adjoint energy densities, but the structure is unchanged. Use the conditions as a **diagnostic**: plot W̄ against Λ/2 on the final surface (§9, figure 8). A converged design shows a flat plateau on the free vertices.

### 5.4 Acceptance and the stay-solid rule

After optimisation (all starts, §7.4), take the best design ξ* that satisfies the constraints and score it on the export geometry (§8, holes cut, mass corrected). Adopt it only if all of the following hold:

$$J^* > \ln(1 + g_{min}), \qquad K/K_0 \ge f_K, \qquad P/P_0 \ge f_P \qquad (g_{min} = 0.02,\; f_K = f_P = 0.85)$$

These are vshell's defaults. They must also hold in a **resolution check** that guards against a gain that is discretisation noise (vshell's ±6 % thin-wall noise):

- **Grid shift.** Re-analyse at the same h_e with the grid origin shifted by ½h_e in x, y and z. This costs nothing extra in memory.
- **Coarser grid.** Re-analyse at h_e = 0.5 mm (h_g = 0.125, k = 4).
- **Finer grid.** Re-analyse at 0.3 mm (k = 3 is allowed for this check) or 0.2 mm, as soon as memory or the Stage 3 sparse grid allows. A dense 0.3 mm cantilever needs about 99 GB (§4.2).

Report every resolution that was run.

Otherwise `final_design.stl` is the original part and the report states the reason, e.g. "best continuous design J = +0.012 ≤ ln 1.02". The report always lists, side by side:

1. the solid;
2. the uniform constant-mass shell (add-in equivalent);
3. vshell's best discrete design;
4. the cshell result.

---

## 6. Sensitivities

### 6.1 Principle and chain

Differentiate the **discrete** model actually computed (discretise-then-differentiate), never a continuous shape-derivative formula. The gradients are then exact up to solver tolerance and float32 rounding, and testable by finite differences.

Forward map:

$$\xi \xrightarrow{F} \tilde\xi \xrightarrow{\text{bounds}} (\mathbf d, \mathbf a) \xrightarrow{\text{as-built}} (\mathbf d, \mathbf a^{ab}) \xrightarrow{\mathbf P} (d_c, a_c) \xrightarrow{\max} \psi_c \xrightarrow{H_\varepsilon} \chi_c \xrightarrow{\text{block mean}} \rho_e \to \hat\rho_e \xrightarrow{\text{FE}} C_k,\ \sigma_{PN,k},\ \tilde\mu_k \to \lambda \to J$$

The volume follows the same chain on the **nominal** a, with no as-built step.

Reverse mode, given g_e = ∂R/∂ρ̂_e from the FE layer:

| Step | Derivative |
| --- | --- |
| ρ̂ → ρ | ∂R/∂ρ_e = (1 − ρ_min) g_e |
| ρ → χ | ∂R/∂χ_c = ∂R/∂ρ_{e(c)} / k³ |
| χ → ψ | ∂χ_c/∂ψ_c = −H′_ε(−ψ_c) |
| ψ → (d, a) | ∂ψ_c/∂d_c = −1 if the outer branch is active (φ_c − d_c + b_o ≥ −φ_c − a^{ab}_c), else 0; ∂ψ_c/∂a^{ab}_c = −1 if the inner branch is active, else 0. Ties are measure-zero; send the gradient to the outer branch. |
| grid → vertex | ∂R/∂**d** = **P**ᵀ ∂R/∂**d**_grid; same for a^{ab} |
| as-built | ∂a^{ab}_i/∂a_i = 1 + S′(z_i), z_i = (r_MA,i − a_i)/b_v. This is ≥ 1, so a^{ab} is monotone in a. |
| bounds | ∂d_i/∂ξ̃ᵈ_i = d^{hi}_i; ∂a_i/∂ξ̃ᵃ_i = a^{hi}_i − a^{lo}(d_i); ∂a_i/∂d_i = a^{lo}′(d_i)(1 − ξ̃ᵃ_i) (D1); zero on passive vertices |
| filter | ∂R/∂ξ = Fᵀ ∂R/∂ξ̃ = M (M + r_f² L)⁻¹ ∂R/∂ξ̃ |

Implement the geometry map on the GPU (CuPy, or JAX with `jax.vjp` for the geometry part only; DLPack interoperates with numba-cuda arrays). Each reverse step needs its own unit test (§6.6).

### 6.2 Compliance

$$\frac{\partial C_k}{\partial \hat\rho_e} = -\frac{2\,U_e}{\hat\rho_e}, \qquad U_e = \text{sed}_e\, h_e^3 \;\text{(from `run_analysis`)}$$

The loads are design-independent because they sit on mounts. A quick check on the CPU path (30×4×4 elements, random ρ ∈ [0.5, 1], q1_im) gave central finite differences of −5.2074·10⁻⁷ against −2 sed V/ρ = −5.2093·10⁻⁷ (same units).

### 6.3 Stress p-norm adjoint (one extra solve per load case)

With s_e(u_e) the von Mises value of D⁰ B̄ u_e (B̄ the centre strain matrix):

$$\frac{\partial\sigma_{PN}}{\partial \mathbf u} = \frac{\sigma_{PN}^{1-p}}{\sum v}\sum_e v_e\, s_e^{p-1}\,\frac{\partial s_e}{\partial \mathbf u_e}, \qquad \frac{\partial s}{\partial \boldsymbol\sigma} = \frac{1}{2s}\big[2\sigma_x - \sigma_y - \sigma_z,\; 2\sigma_y - \sigma_x - \sigma_z,\; 2\sigma_z - \sigma_x - \sigma_y,\; 6\tau_{yz},\; 6\tau_{xz},\; 6\tau_{xy}\big]$$

The Voigt order is (xx, yy, zz, yz, xz, xy). Then:

1. **Adjoint solve.** K **η** = −∂σ_PN/∂**u**, with fixed DOF zero as in the primal. Call `matrix_free_pcg` with exactly the primal's operator and a relative stopping rule, because its defaults differ from `ConjugateGradient`'s and would solve a different K:
   - the same `E_elem`, `nu_elem` and fixed mask;
   - `element_type="q1_im"`, `nu=ν`, `preconditioner="gmg"`;
   - `tol=0.0` and `rel_tol=1e-8`. The default `tol` is an absolute residual in newtons, meaningless for this right-hand side.

   Warm-start from the previous iterate's η. `matrix_free_pcg`, `compute_element_stress` and `compute_element_strain_energy` are **CUDA-only**. The CPU test suite (R8) needs a host path: assemble K with the q1_im element on the host (`ConjugateGradient._cg_assemble_stiffness`, or vshell's `fe.py`) and solve the adjoint with the same factorisation `_run_analysis_cpu` uses.
2. **Total derivative.**

$$\frac{d\sigma_{PN}}{d\hat\rho_e} = \frac{\partial\sigma_{PN}}{\partial v_e}\frac{\partial v_e}{\partial\hat\rho_e} + \boldsymbol\eta_e^{\mathsf T}\mathbf K_e^0\,\mathbf u_e, \qquad \frac{\partial\sigma_{PN}}{\partial v_e} = \frac{\sigma_{PN}^{1-p}}{p\sum v}\big(s_e^{p} - \sigma_{PN}^{p}\big), \qquad \frac{\partial v_e}{\partial\hat\rho_e} = \frac{q\,\rho_e^{q-1} m_e}{1-\rho_{min}}$$

The bilinear element term **η**_eᵀK_e⁰**u**_e can come from `compute_element_strain_energy` by polarisation, ¼[E(u+η) − E(u−η)]·(2/ρ̂_e), with care over the ½ and the ρ̂ scaling. Better, write a small bilinear kernel next to `_element_energy_kernel_mat`.

Then ∂ln λ_s/∂ρ̂ = −∂ln σ_PN/∂ρ̂, with c held fixed.

### 6.4 Buckling

`BucklingResults.sensitivity` is ∂μ̃/∂ρ̂ on the voxel grid passed to `run_analysis`. The geometric-stiffness ramp term ("Term C") is included. So:

$$\frac{\partial \big(\tilde\mu_k/\kappa_b\big)^{\beta}}{\partial\hat\rho_e} = \frac{\beta}{\kappa_b}\Big(\frac{\tilde\mu_k}{\kappa_b}\Big)^{\beta-1}\frac{\partial\tilde\mu_k}{\partial\hat\rho_e} \qquad (\text{zero when } \tilde\mu_k = 0)$$

If buckling runs on the 0.8 mm grid (§4.2 item 2), its sensitivity is spread back to the 2³ fine elements of each coarse one: the adjoint of block averaging is division by 8.

**Important: the reference buckling sensitivity is exact only for `q1`/`q1_gauss`, not for `q1_im`.** The eigenproblem is assembled with the q1_im matrix (`E_e·Khat`). But Term A (−μ φᵀ ∂K/∂ρ φ) and the explicit part of the adjoint term use the trilinear Gauss solid stiffness `ElementData.K0 = λ_s Klam + μ_s Kmu` (`analysis.py` ≈ l. 5018, 6013, 6047; the `ElementData` docstring says so). So for q1_im, ∂μ̃/∂ρ̂ is not the derivative of the model that was solved.

A directional check on the CPU path makes this concrete. Set-up: `engine="classic"`, 30×4×5 column, random ρ̂ ∈ [0.5, 1], mixed axial and transverse tip load, 4 modes, p = 8, smooth perturbation δρ̂ = 0.05 sin(πx)(1 + 0.3 noise), central differences at h = 1, 0.3, 0.1:

| Element type | Relative error | Behaviour |
| --- | --- | --- |
| `q1` | 3·10⁻³, 3·10⁻⁴, 9·10⁻⁵ | converges as O(h²) |
| `q1_im` | 5.4 %, 5.7 %, 5.8 % | constant bias |

Element-wise finite differences with steps of 10⁻³–10⁻² are too noisy to show this; the default `lobpcg_tol = 10⁻³` makes them worse.

`cshell` must therefore correct the two terms for q1_im **inside cshell**. Subclass or wrap the buckling engine, replacing the Gauss solid stiffness with E_s·Khat, E_s = 2μ_s(1 + ν), wherever ∂K/∂ρ̂ appears: Term A and the explicit adjoint term. The geometric-stiffness terms built from compatible-B stresses are consistent as they stand. Do not edit the reference file.

Then the directional test above must reach < 10⁻³ for q1_im, as it does for q1. (v1.2: done — `cshell.fea.ShellCG.run_buckling` applies the fix, directional error 5.3·10⁻² → 1.2·10⁻⁴; the production engine `fe_sparse.buckling` uses K⁰_e = E·K̂ throughout and passes the same test at 3.4·10⁻⁵.) Record the issue, with this reproducer, in `continuous/OPEN_QUESTIONS.md` for Owen. Until the fix passes, buckling gradients are approximate (about 6 % bias, correct direction). MMA tolerates that, but only accept a step if the evaluated J, not the predicted one, improves (GCMMA-style conservativeness).

Test buckling sensitivities with **directional derivatives along smooth perturbations** of size about 10⁻² in ρ̂ (signal far above eigen-solver noise), using `engine="classic"` or a tight `lobpcg_tol`.

### 6.5 Capacity, objective and constraints

With T = Σ_k (λ_{s,k}^{−β} + (μ̃_k/κ_b)^β), so that ln λ = −(1/β) ln T:

$$\frac{\partial\ln\lambda}{\partial\hat\rho} = -\frac{1}{\beta T}\sum_k\Big(-\beta\,\lambda_{s,k}^{-\beta}\,\frac{\partial\ln\lambda_{s,k}}{\partial\hat\rho} + \frac{\partial(\tilde\mu_k/\kappa_b)^{\beta}}{\partial\hat\rho}\Big), \qquad \frac{\partial J}{\partial\hat\rho} = \frac{-w_K\sum_k\omega_k\,\partial\ln C_k/\partial\hat\rho \;+\; w_P\,\partial\ln\lambda/\partial\hat\rho}{w_K+w_P}$$

$$\frac{\partial V}{\partial\psi_c} = -h_g^3\,H'_\varepsilon(-\psi_c) \;\;(\text{nominal geometry}), \qquad \frac{\partial G_s}{\partial a_i} = -\frac{A_i}{A_\Gamma}\,\frac{B'(z_i)}{v_{min}/2}$$

### 6.6 Verification protocol (gate for every stage)

All of these run in `pytest`, CPU-only at small sizes. GPU variants are marked `@pytest.mark.gpu`.

1. **Geometry chain.** Directional derivative of V and of a synthetic linear functional Σ g_e ρ̂_e against central finite differences, along δξ = F·(random). Float64 path, relative error < 10⁻⁶. Run a Taylor test: the first-order remainder must fall as O(h²) over at least two decades.
2. **Filter.** F is symmetric with respect to M, so for random x, y: xᵀ M F y = yᵀ M F x to 10⁻¹². Its adjoint must pass the dot-product test.
3. **Compliance.** Element-wise central finite differences on a 30×4×4 column at tight PCG tolerance: < 10⁻³ relative.
4. **Stress adjoint.** Directional derivative on a 40×6×6 notched bar. Float64 CPU assembly is allowed for the reference; < 10⁻³ relative.
5. **Buckling.** As in §6.4: < 10⁻³ on the directional derivative for q1 (reference code) and for q1_im (after cshell's correction).
6. **End to end.** dJ/dξ along a smooth direction on a small tube cantilever (h_e = 0.8 mm), including the as-built step, filter and mass terms: < 2 %.
7. **Mass restoration** (§7.2). After a random MMA-like step, \|V/V₀ − 1\| < 10⁻⁴ in ≤ 20 volume evaluations.

---

## 7. The optimiser

### 7.1 MMA

Use the Method of Moving Asymptotes (Svanberg 1987), on the CPU in float64. There are n = 2·(active vertices), about 10⁵, design variables and m ≤ 5 constraints. The MMA subproblem is O(n) per inner iteration, which is negligible next to the FE.

Use `mmapy` (GCMMA-MMA-Python, `pip install mmapy`) or port Svanberg's 2007 note faithfully. First test it on Svanberg's toy problems (§11 V5-B4).

| Setting | Value |
| --- | --- |
| Move limit | 0.1 in ξ, halved on oscillation (§7.5) |
| asyinit, asyincr, asydecr | 0.5, 1.2, 0.7 |
| a₀, aᵢ, cᵢ, dᵢ | 1, 0, 10³, 1 |

Scale every constraint to O(1).

### 7.2 Mass restoration after every update

MMA linearises the mass constraint, so its iterates drift by up to τ_V. To meet R1 exactly, every MMA update ξ⁺ is followed by a 1-D restoration that uses only the geometry, with no FE:

$$\xi \leftarrow \operatorname{clip}_{[0,1]}\big(\xi^+ + \eta\,\mathbf g/\lVert\mathbf g\rVert_\infty\big), \qquad \mathbf g = \nabla_\xi V(\xi^+) \text{ on active vertices}, \qquad V(\xi) = V_0 \text{ solved for } \eta \text{ by Brent}.$$

V is monotone in η along g, even with clipping, so the root is unique. Tolerance is 10⁻⁴ relative. Each V evaluation costs milliseconds on the GPU, and typically 5–10 are needed.

Every reported iterate is then an honest constant-mass design that could be exported and compared. The MMA state (asymptotes, previous iterates) uses the restored ξ.

### 7.3 Continuation schedule

| Phase | h_e (mm) | h_g (mm) | r_f (mm) | p (stress) | β (soft-min) | w_s (sliver) | Buckling grid | Max iterations |
| --- | --- | --- | --- | --- | --- | --- | --- | --- |
| A | 0.8 | 0.2 | 4.0 | 6 | 8 | 0 | 0.8 | 60 |
| B | 0.4 | 0.1 | 3.0 | 8 | 12 | 0 → 0.05, linear over the phase | 0.8 or 0.4 (if it fits) | 80 |
| C (Stage 3, sparse) | 0.2 | 0.1 | 2.5 | 10 | 16 | 0.1 | 0.4 or 0.2 | 40 |

With w_s = 0 in phase A, regions open and close on mechanics alone. Check the history (§7.6) for whether regions are still changing state when w_s is switched on. If many are, lengthen phase A rather than raise w_s.

Between phases:

- ξ carries over unchanged, because the design lives on Γ_h, not on a grid.
- Rebuild P and the grids, re-measure V₀ on the new grid (constant mass is relative to each grid's own measure of the solid, as in vshell), and re-run the restoration.
- Reset the MMA asymptotes, re-estimate the stress normalisation c, and re-evaluate the solid baseline (C₀, λ₀) at the new resolution.

### 7.4 Starting designs (multi-start)

The problem is non-convex: void closure is close to a topological switch. Run these starts as independent runs (one GPU each), and keep the best feasible result by J. Report all of them.

| Start | Definition |
| --- | --- |
| S1 uniform | ξᵃ = 0 (a = t_min everywhere), ξᵈ uniform, restored to mass. The closest representable uniform shell. The add-in's own shell (a = 0, d = t) is evaluated separately for the comparison table. Its mount (load/support) faces are kept solid, as in the optimiser (Owen 2026-10-30, D1): ψ = max(φ − Pd, −φ − Pa) with d = t, a = 0 on free vertices and d = 0, closed (a = a^hi) on mount vertices; t is solved for constant mass on the optimisation band and refined by the export's mass correction. (The pure shell ψ = max(φ − t, −φ) put the load face on the inner skin; with the as-built bias its load nodes lost their material on grid-aligned faces: MBB K/K₀ 0.054 against 0.67–0.76.) Every re-scored design applies its BCs only at nodes that touch material (an adjacent element with ρ ≥ 10⁻³); a load's share on dropped nodes is redistributed equally over its remaining nodes and reported, and material components not connected to a clamped node are removed before the FE. |
| S2 vshell | Run `vshell` (or load its saved arrays) and map its best design: ξᵃ_i = θ of the region containing the vertex's inward point at depth r_MA/2, and ξᵈ = s(1 − θ)/d^{hi}. Restore. |
| S3 neutral | ξᵃ = 0.5, ξᵈ restored. |
| S4 energy-guided (optional) | From the solid FE, set ξᵃ_i = clip(W_axis,i / W_surf,i, 0, 1): high axis energy (shear) gives closed, low gives open. |

### 7.5 Convergence and safeguards

A phase stops when all of these hold:

- max|Δξ| < 0.005 and |ΔJ| < 10⁻⁴ for 5 consecutive iterations;
- all constraints are satisfied;

or when it reaches its iteration cap.

Safeguards:

- **Oscillation.** If ΔJ changes sign in 6 of the last 8 iterations, halve the move limit.
- **Failed solve.** A non-converged PCG or eigen-solve (`cg_info`, `eigensolver_converged`) is never accepted silently. Retry once with a cold start and tighter settings; otherwise halve the step and log it.
- **Divergent objective.** A NaN or infinite objective stops the run with a diagnostic dump.

### 7.6 Checkpoints and logging

After every iteration, write `checkpoints/<start>_<phase>_<iter>.npz`, keeping the last three. It holds ξ, the MMA state, c, the previous buckling modes, warm-start displacements, the iteration counter and the RNG state. `--resume <run_dir>` continues from the newest checkpoint. The cluster runs SLURM with a **24 h wall-time limit**, and a full multi-phase run may not fit in one allocation, so:

- `--max-hours H` (e.g. 23.5) stops cleanly after the iteration in progress and writes a checkpoint;
- a resumed run must reproduce the history of an uninterrupted one;
- `run_info.json` lists every allocation the run used, with the SLURM job id if present.

One CSV row per iteration (`history.csv`): start, phase, iteration, J, C/C₀ per case, λ_s/λ_s0, λ_b/λ_b0, λ/λ₀, governing mode/case, V/V₀ − 1, G_s, max|Δξ|, move limit, PCG iterations per solve, eigensolver iterations and convergence flag, time per component (geometry, FE, adjoint, buckling, MMA, restoration), and peak GPU memory.

---

## 8. Post-processing and export

### 8.1 Export geometry

The **nominal** design (not as-built) is exported. ψ is evaluated on an export grid of h_x = 0.05 mm, processed in tiles so the full grid is never held at once.

### 8.2 Clean-up

1. Vertices with ξ̃ᵃ ≥ 0.98 are treated as exactly closed (a = ∞), which removes the axis layer of §3.8.
2. Apply a morphological opening to the void mask with a ball of diameter v_min = 3 mm. Every remaining void is then a union of such balls: tapering void tips become rounded ends, and slivers become solid. Also fill enclosed voids under 30 mm³ (vshell's `void_min_mm3`). Report the volume filled; it is restored by the mass correction (§8.4). Check the rule explicitly: re-opening the cleaned void must change it by no more than one export cell.
3. Keep the largest connected material component.

### 8.3 Escape holes

For each enclosed void component (void cells not connected to the exterior):

1. Collect candidate wall sites on the void's surface where:
   - the nominal wall is ≥ t_min + b_v;
   - the site is more than 3 mm from any mount;
   - the final-FE material von Mises stress is in the lowest 30 % of that void's wall.
2. Choose `escape_holes_per_void` (2) sites that maximise their separation along the void's principal axis.
3. Cut a cylinder of diameter `escape_hole_diameter_mm` (3 mm) along the local outward normal, from the void through the wall: ψ ← max(ψ, D/2 − ‖(x − p) × n‖), capped to the wall.
4. The cut length is chosen so both ends of the hole clear the walls (no sub-t_min film at the hole end); wall-thickness statistics count only true opposite-wall hits, and the hole rims are reported separately (`wall_off_rims`).
5. Log each hole in `escape_holes.csv` (position, normal, wall, stress).

Stage 3 may bring holes inside the loop as fixed features once the void topology settles.

### 8.4 Final mass correction

Find δ by Brent so that V = V₀ on the export grid, applying d_i ← clip(d_i + δ, 0, d^{hi}_i) to the active vertices. Report δ: warn above 0.1 mm, fail above 0.3 mm.

### 8.5 Meshing and checks

1. Marching cubes (skimage) on ψ at h_x, to a watertight, manifold mesh (trimesh checks). Taubin smoothing and decimation are allowed only with a volume-preservation check.
2. Write `final_design.stl` (3MF optional).
3. Check \|V_mesh/V₀ − 1\| ≤ 2·10⁻³ (R1).
4. Measure the wall-thickness distribution by inward ray casting (histogram and minimum, nominal and as-built).
5. Count the voids and confirm each has at least one hole.

### 8.6 Final analyses and body-fitted validation

1. **Voxel FE** of the exported geometry at the optimisation resolution and the next finer one (§5.4). All responses, for both the solid and the final design.
2. **Body-fitted check** (`validate.py`, `bodyfit_gpu.py`; v1.3).
   - Surface: the exported STL is remeshed (pymeshlab isotropic remesh to edge h, quadric pre-decimation of marching-cubes input, midpoint subdivision of very coarse input, repair of self-intersections, projection onto the input). Checks: watertight, no self-intersection, |ΔV/V| ≤ 5·10⁻³, longest edge ≤ 2.5 h. The remesh runs in a spawned process (a second call in one process silently skipped the remeshing on B2). Solid and design go through the same chain, so mesher bias cancels in the ratios.
   - Volume: gmsh inside that triangulation (Delaunay + Netgen/Relocate3D rounds, variants tried until the minimum dihedral is ≥ 1°), no TetGen fallback (its slivers moved λ_b by −46 % on B2). Quadratic tets (C3D10).
   - Levels h = 0.5 mm and h/√2; the finer is compared, and a change in C above 2 % between levels is flagged.
   - Gate on the finer level (failure blocks acceptance): ≥ 2 tets through the thinnest wall (1 % quantile), minimum dihedral ≥ 0.25°, fraction of tets below 5° ≤ 10⁻³, longest edge ≤ 2.5 h.
   - Solve: default `validation.engine = "gpu"`: the same C3D10 model on cuDSS (static, block-Krylov buckling, 0.995 stress quantile with the 2 mm BC exclusion). It matches CalculiX on the B2 design deck (h = 0.5): C 8·10⁻⁹, σ_q 6·10⁻⁸, λ₁ 2·10⁻⁶ (CalculiX prints 7 digits); modes 2–4 differ by 5·10⁻⁴ to 9·10⁻² (CalculiX's eigen-solver; not investigated). `"ccx"` remains available.
   - Buckling uses the K_σ exclusion within `capacity.exclude_bc_mm` of clamped nodes (same rule as the voxel side, §4.4); λ without exclusion is recorded as `lam_unexcl`.
   - Tolerances: compliance ≤ 5 %, λ_b ≤ 10 %, 0.995 stress quantile ≤ 15 % (each relative). The ratios K/K₀ and P/P₀ are checked one-sided (Owen 2026-10-30, D3): a ratio blocks only when the voxel model over-predicts it, r_vox/r_bf − 1 > 5 %. A voxel under-prediction is conservative for the decision and is logged (`under` in validation.csv), not blocking. Evidence: B2 P_b/P_b0 voxel 1.79–1.81 at h_e 0.2–0.4 against the mesh-converged body-fitted 1.932 (−0.14 % between levels; `research/vf_b2/voxel_conv.py`). A blocking ratio miss, or a failed gate, blocks acceptance.
   - Export gate before re-scoring (D2): when the best export fails its checks (mass correction not bracketed or |δ| > δ_fail, not watertight, or |V_STL/V₀ − 1| > 2·10⁻³), the STL is written for inspection, it is not re-scored or validated, and the original solid is kept with reason "export failed: …". A large void_min opening fill (§8.2) on an unconverged design is reported as optimiser non-convergence, not a re-scoring error.

---

## 9. Report (`report.pdf`)

The report follows `vshell/report.py`, which itself follows `baseline_optimiser.py`'s post-processing:

- text summary pages first, then one figure per page;
- each figure built in its own `try/except`, with failures listed in `plots/report_log.txt`;
- 3D artists rasterised so the PDF stays under about 5 MB;
- rebuildable from the run folder with `python -m cshell report <run_dir>`.

Pages:

1. **Summary.** Decision banner and reason. Mass check. Comparison table: solid, uniform shell, vshell best, cshell for each start, final export; with K/K₀, P/P₀, J, governing mode and wall range. Run settings.
2. **Geometry and boundary conditions** (reuse `vshell.report.fig_geometry`).
3. **Design space.** r_MA, d^{hi} (showing keep-outs and clearance) and passive vertices, as surface colour maps.
4. **Result fields.** Wall w, offset d and inset a on the surface, three views each. Closed (solid) area highlighted. Histogram of w, nominal and as-built.
5. **Sections.** Three orthogonal cuts through the largest joint and the most-changed strut, original against final.
6. **Convergence.** J, C/C₀, λ_s, λ_b, V/V₀ − 1, G_s and max|Δξ| against iteration. Phases shaded, all starts overlaid.
7. **Stress.** Material von Mises, solid against final, two views. Histogram against σ_f.
8. **Buckling.** First three mode shapes, solid against final, with λ and global/local tags.
9. **Optimality diagnostic.** W̄ on the free surfaces against Λ/2 (§5.3): a scatter and a surface map.
10. **Capacity and decision robustness.** λ_s, λ_b and soft-min per load case. Decision re-scored over κ_b ∈ {0.6, 0.7, 0.85, 1.0} and σ_f ∈ {−20 %, …, +20 %}, which needs no new FE because λ_b ∝ κ_b and λ_s ∝ σ_f. Also written to `sensitivity.csv`.
11. **Resolution check.** J and the ratios at every resolution of §5.4 (grid shift, coarser, and finer if run).
12. **Validation.** Body-fitted against voxel table.
13. **Voids and escape holes.** Inventory, holes, filled volume, minimum-wall map.
14. **Resources.** GPU model, memory peaks, per-iteration time breakdown, total GPU-hours.
15. **Provenance.**
    - git commit and dirty flag;
    - SHA-256 of the STEP, the config and every source file, including `gpu_analysis.py` and `analysis.py`;
    - package versions;
    - CUDA driver and runtime;
    - hostname and scheduler job id.

---

## 10. Software architecture and operation

### 10.1 Layout (repository root)

```
cshell/
  __init__.py, __main__.py, cli.py
  config.py        defaults + JSON load/validate (extends vshell.config; same supports/loads/material keys)
  surface.py       design mesh Γ_h (remesh), mass/cotangent matrices, Helmholtz filter F and Fᵀ,
                   r_MA, r_out, keep-out bounds, passive vertices                                   (§3.4–3.5)
  band.py          sparse band grid: φ, closest points, P (fine cell → vertex interpolation), cache  (§3.1, 3.6)
  design.py        ξ → d, a, a_ab → ψ → χ → ρ, ρ̂; V; reverse-mode pull-backs; sliver measure G_s     (§3.7–3.10, 6.1)
  fea.py           ShellCG(analysis.ConjugateGradient) with node-level BCs; compliance, stress p-norm
                   and adjoint, buckling wrapper; unit conversions; per-load-case instances          (§4, 6.2–6.4)
  objective.py     J, soft-min, normalisation c, gradients                                           (§5, 6.5)
  mma.py           MMA wrapper (mmapy) + restoration                                                 (§7.1–7.2)
  optimise.py      phases, continuation, multi-start, checkpoints, history, decision                 (§7.3–7.6, 5.4)
  export.py        tiled ψ, clean-up, escape holes, mass correction, marching cubes, checks          (§8.1–8.5)
  validate.py      tetrahedral meshing + CalculiX driver + comparison                                (§8.6)
  report.py        PDF report                                                                        (§9)
  benchmarks.py    analytic benchmarks of Appendix B
tests/cshell/      pytest suite (§6.6, §11); CPU-only unless marked @gpu
examples/cshell/   configs for the three Study 1 parts and the benchmarks
continuous/        this document, the prompt, reference/, environment.yml,
                   STATUS.md (stage log), BENCHMARKS.md (measured costs), OPEN_QUESTIONS.md
```

- **Import the reference stack** from `continuous/reference/` through `cshell/fea.py`, which puts that folder on `sys.path`. Do not copy or edit those files, so Owen's later fixes propagate. A genuine bug in them goes in `continuous/OPEN_QUESTIONS.md` with a minimal reproducer, worked around in `cshell`.
- **Reuse vshell:** `geometry.py` (STEP, face table, selectors, tessellation), the mount detection part of `segment.py`, `field.PartGrid` ideas for SDF and volume, `util.py` (provenance, CSV/JSON), and `report.py` helpers. vshell keeps working unchanged; `cshell` imports from it.

### 10.2 Configuration (JSON; every key defaulted except `step_file`, `supports`, `loads`)

```json
{
  "name": "Cantilever", "step_file": "step/Solid Cantilever.step", "output_dir": "runs_cshell",
  "supports": [{"type": "cylinder", "radius": 5.0, "tol": 0.05}],
  "loads": [{"faces": {"type": "plane", "normal": [0,0,1], "box": {"x": [130,170]}, "extreme": true},
             "force_N": [0,0,-200], "weight": 1.0}],
  "material":      {"E_MPa": 1760, "nu": 0.42, "density_g_cm3": 1.04, "strength_MPa": 36.6},
  "manufacturing": {"t_min_mm": 1.5, "wall_bias_mm": 0.2, "outer_bias_mm": 0.0, "void_min_mm": 3.0,
                    "void_min_mm3": 30, "escape_holes_per_void": 2, "escape_hole_diameter_mm": 3.0},
  "mounts":        {"auto_analytic": true, "extra_faces": [], "min_blob_fraction": 0.01, "dilate_mm": 1.5},
  "design":        {"allow_outward": true, "outward_extent_initial_mm": 4.0, "gap_mm": 1.0, "keep_out": [],
                    "keep_out_clearance_mm": 0.5, "surface_edge_mm": 0.75, "closure_gamma": 0.75},
  "grid":          {"phases": [{"h_e": 0.8, "h_g": 0.2}, {"h_e": 0.4, "h_g": 0.1}], "export_h_mm": 0.05,
                    "rho_min": 1e-6, "tessellation_mm": 0.02},
  "fe":            {"element_type": "q1_im", "preconditioner": "gmg", "engine": "auto",
                    "pcg_rel_tol": 1e-6, "pcg_max_iters": 10000, "warm_start": true},
  "capacity":      {"stress_p": [6, 8, 10], "stress_q": 2, "stress_quantile": 0.995, "exclude_bc_mm": 2.0,
                    "include_mounts": false, "buckling": true, "buckling_modes": 6, "buckling_p": 8,
                    "knockdown": 0.85, "softmin_beta": [8, 12, 16], "lobpcg_tol": 1e-3},
  "objective":     {"w_stiffness": 0.5, "w_capacity": 0.5, "min_gain": 0.02,
                    "floor_stiffness": 0.85, "floor_capacity": 0.85, "floors_as_constraints": false},
  "filter":        {"radius_mm": [4.0, 3.0, 2.5]},
  "optimiser":     {"starts": ["uniform", "vshell", "neutral"], "max_iter": [60, 80, 40], "move": 0.1,
                    "tol_x": 0.005, "tol_f": 1e-4, "mass_tol": 1e-4, "sliver_weight": [0.0, [0.0, 0.05], 0.1]},
  "validation":    {"body_fitted": true, "second_resolution": true},
  "report":        {"enabled": true, "dpi": 130}
}
```

Phase-indexed lists (`stress_p`, `softmin_beta`, `radius_mm`, `max_iter`, `sliver_weight`) give one value per continuation phase. A two-element entry is a linear ramp over that phase.

### 10.3 Command line

```bash
python -m cshell faces "examples/step/Solid L-bracket.step" --png faces.png     # same as vshell
python -m cshell init  "part.step"                                               # template config
python -m cshell setup examples/cshell/cantilever.json                           # geometry, band, P, caches; prints DOF + memory plan
python -m cshell run   examples/cshell/cantilever.json [--start uniform] [--phase B] [--set key=value ...]
python -m cshell run   ... --resume runs_cshell/Cantilever
python -m cshell collect runs_cshell/Cantilever                                  # merge starts, decide, export, validate
python -m cshell report  runs_cshell/Cantilever
python -m cshell bench   [B1|B2|B3|all]                                          # Appendix B benchmarks
```

### 10.4 Execution environment and dependencies

**Facts.**

- **Hardware.** NVIDIA GH200 GPUs (aarch64 Grace host). Check that every dependency has an aarch64 build; `gpu_analysis` already copes with platforms that lack pypardiso.
- **Scheduler.** SLURM with a 24 h wall-time limit, which is why `--max-hours` and `--resume` exist (§7.6).
- **Working practice.** How to work on the cluster (jobs, nodes, scratch, version control) is governed by Owen's own instructions for Claude Code. This spec does not repeat them.

**Dependencies.** Listed in `continuous/environment.yml`:

- Python ≥ 3.10, numpy, scipy, numba (CUDA), cupy;
- build123d (OCP), libigl, trimesh, scikit-image, gpytoolbox or pymeshlab;
- warp-lang, mmapy, matplotlib, pytest, calculix.

Optional: jax[cuda], nvmath-python (cuDSS), scikit-sparse, pyvista, tetgen, gmsh.

### 10.5 Outputs (`<output_dir>/<name>/`)

| File | Contents |
| --- | --- |
| `final_design.stl` | the accepted design, or a copy of the original |
| `best_candidate.stl` | best design, kept even when rejected |
| `original.stl` | the input, tessellated |
| `uniform_shell.stl`, `vshell_best.stl` | comparison designs |
| `fields.vtp` | Γ_h with ξ, d, a, w, r_MA, d^{hi}, W̄, stress, for ParaView |
| `voxels_<phase>.vtr` | optional, size-capped |
| `history.csv`, `evaluations.csv`, `escape_holes.csv`, `sensitivity.csv`, `validation.csv` | logs and results |
| `summary.json/.txt`, `run_info.json`, `config_used.json`, `run_log.txt` | summary and provenance |
| `report.pdf`, `plots/*.png` | report |
| `checkpoints/` | restart state |

### 10.6 Coding standards

- Type hints throughout. Docstrings give the equations and cite this document's §.
- No silent fallbacks: every engine or resolution fallback is logged and shown in the report.
- Float64 for all reductions (volume, J, MMA). Float32 is allowed for GPU fields.
- Deterministic seeds. GPU reductions may be non-deterministic, so state the tolerance.
- Every public function has a unit test or is covered by an integration test.
- Commit at each passing gate with a message stating which tests passed.
- Keep `continuous/STATUS.md` current: done, next, blockers, measured numbers.

---

## 11. Verification and validation plan

| ID | Test | Pass criterion |
| --- | --- | --- |
| V1-a | Geometry: volume of a solid cylinder and a tube (analytic), through ψ | \|ΔV/V\| < 2·10⁻³ at h_g = 0.1 mm |
| V1-b | r_MA on a cylinder (r = 5) and a 10 × 20 mm box | ±0.1 h_g |
| V1-c | All-solid via fields against ψ = φ (§3.8) | V within 5·10⁻⁴; C within 0.5 % |
| V1-d | Filter: maximum principle, M-symmetry, adjoint dot test | exact to round-off |
| V2-a | Solid baselines of the three parts against vshell (h_e = 0.8, same BCs) | K₀ within 2 % |
| V2-b | Study 2 PA11 tubes A–E in 3-point bending against Abaqus (`tests/validate_tubes.py` targets 0.0086, 0.0104, 0.0158, 0.0245, 0.0351 kN/mm) at h_e = 0.8 and 0.4 | within 4 % |
| V2-c | Thin-wall convergence, 1.7 mm-wall tube, h_e = 0.8 → 0.2 | monotone; 0.4 within 3 % of 0.2 |
| V3-a | Euler: clamped–free solid rod r = 5, L = 100 mm, P_cr = π²EI/(4L²) = 213 N | within 5 % at h_e = 0.4 |
| V3-b | Local shell buckling: tube R_m = 10, t = 1.5, L = 60, axial, against ccx `*BUCKLE` on C3D10 | within 10 % at h_e = 0.2–0.4 |
| V4 | Sensitivities §6.6 (1)–(7) | as stated there |
| V5-B1 | Stub cantilever rod (L/D = 0.5): hollowing loses stiffness | **starting from the all-hollow S1 design**, the optimiser closes the void by itself (ξ̃ᵃ → 1 over ≥ 95 % of the active area) and returns solid (J ≤ ln 1.02). This is the natural-closure test. |
| V5-B1b | Rod with a slender part (L/D = 8) joined to a stub (L/D = 0.5) by a fillet | from S1: stub closes and slender part stays hollow; void ends are rounded and ≥ 3 mm after export (§8.2). |
| V5-B2 | Slender cantilever rod (L/D = 8) | hollow; J ≥ best uniform shell; wall/offset larger toward the root |
| V5-B3 | Clamped–free column (L = 100, r₀ = 5) | hollow, a → t_min; λ_b ≥ best uniform tube |
| V5-B4 | MMA on Svanberg's toy problems | reference optima |
| V6 | Study 1 parts, full pipeline | §11.1 |
| V7 | Re-run identical config | final J within 10⁻³ (relaxed from 10⁻⁴ by Owen, 2026-10-30: GPU atomic-add noise ~10⁻¹⁶ is amplified by MMA; bitwise reruns need the opt-in `fe.twogrid.deterministic = "full"` mode plus a deterministic geometry chain, not built) |

### 11.1 What to expect on the Study 1 parts (and how to report it)

- **Cantilever.** Expect the two compression struts hollowed, with walls thickest where the moment is largest. Joints and tension members stay mostly solid. Expect J ≥ vshell's +0.48 (vshell: K/K₀ 0.992, P/P₀ 2.63). The real eigen-analysis may lower the vshell capacity estimate, because frame modes are now visible. Report the change and its mechanism.
- **L-bracket.** Expect joints closed. Thin struts may hollow or stay solid, and the decision is likely marginal; vshell's borderline was J = +0.004. Do not tune anything to force a hollow result.
- **MBB.** A linear model is expected to keep it solid, which conflicts with the +14.8 % measured capacity (ductile redistribution, §4.5). Say so plainly.

In every case, tabulate the predicted K/K₀ and P/P₀ against the manuscript's measured changes, with sign agreement (as in vshell's theory doc).

---

## 12. Roadmap and gates

| Stage | Content | Exit gate |
| --- | --- | --- |
| **0. Set-up** | Environment (§10.4). Import the reference stack (`interfaces.py` stub). `gpu_analysis.probe()`, `verify_host_kernel_consistency()`. `ShellCG` adapter with node BCs and unit conversion. Memory/time benchmarks into `BENCHMARKS.md`. | V2-a and V2-b pass. Measured per-solve cost at h_e = 0.8 and 0.4 for the three parts. |
| **1. Geometry, mass, compliance** | Γ_h, filter, r_MA, r_out, keep-outs, band and P, ψ/χ/ρ chain with pull-backs, V, restoration, MMA. Compliance-only objective (w_P = 0). Starts S1 and S3. Phases A and B. | V1, V4 (1–3, 6, 7), V5-B1, B2, B4 pass. Compliance-only runs of the three parts converge with \|V/V₀ − 1\| ≤ 10⁻⁴ at every iterate. |
| **2. Capacity, robustness, manufacture, decision** | Stress p-norm and adjoint, buckling and chain, as-built, soft-min J, sliver penalty, S2 start, export (clean-up, holes, mass correction), two-resolution decision, body-fitted validation, full report. | V2-c, V3, V4 (4, 5), V5-B3, V6, V7 pass. Reports for all three parts. |
| **3. Scale and extensions** (prioritise with Owen) | (a) Narrow-band sparse FE to 0.2/0.1 mm. (b) Multi-GPU. (c) Escape holes in the loop. (d) Global/local mode classification with separate knock-downs. (e) Plastic limit-load capacity (fixes the MBB miss). (f) Maximum-stiffness (kinematic washer) BC bound and worst case. (g) Interior topology freedom: ribs or lattice inside voids (coating/infill TO, Clausen et al. 2015; Wu et al. 2018). (h) Fusion round-trip of the mesh. | Per item. |

At each gate: tests green, commit, `STATUS.md` updated, and a short summary for Owen in `continuous/STATUS.md`.

---

## 13. Risks and mitigations

| Risk | Mitigation |
| --- | --- |
| Local optima; closed voids freeze | Multi-start (§7.4). γ_c keeps closure gradients alive (§3.8). Optional reopening probe. Continuation from a large filter radius. |
| Voxel noise on thin walls drives the optimiser | Quintic Heaviside at the cell scale. Filter radius ≥ 2.5 mm. Two-resolution acceptance. Body-fitted validation. |
| float32 limits finite-difference checks | Directional derivatives with smooth, larger perturbations. Float64 CPU references at small size. |
| GPU memory at 0.4 mm dense (buckling) | `buckling_memory()` pre-check; buckling on the 0.8 mm grid with chained sensitivity; Stage 3 sparse grid. |
| Spurious buckling modes in void elements | `min_density_for_geometric_stiffness` = 0.1 (built in). Inspect the mode shapes in the report. |
| Mode switching and degenerate clusters | p-norm aggregate, `prev_modes`, whole-cluster retention (built into `analysis.py`). |
| Stress singularities at clamps and loads | Exclusion zones; p-norm calibrated to the 0.995 quantile. |
| Linear model misses ductile gains (MBB) | Stated in the report. Stage 3 plastic limit analysis. |
| Outward growth fuses neighbouring members or enters keep-outs | r_out and keep-out bounds on d (§3.5). Report the d^{hi} map. |
| Medial-axis estimate errors | Cross-checked by ray chord. Unit-tested on primitives. V1-c guards the solid case. |
| Mass drift at export (meshing, holes, clean-up) | Final δ correction with limits (§8.4); R1 check on the mesh. |
| Reference code expects absent modules | `interfaces.py` stub. Note anything else in `OPEN_QUESTIONS.md`. |

---

## 14. Decisions and remaining open questions

**Decided by Owen (28 October 2026):**

| # | Topic | Decision | Where |
| --- | --- | --- | --- |
| D1 | Outward growth | Allowed, **no ceiling**. Keep clearance from other members and from keep-out bodies. No keep-out is needed for the Study 1 parts (ignore the L-bracket's GD obstacle). | §3.1, §3.4, §3.5 |
| D2 | Minimum void thickness | v_min = **3 mm**. The optimiser must still be free to close any region to solid: each point is either a void ≥ 3 mm or solid. | §3.8, §5.2, §7.3, §8.2, V5-B1/B1b |
| D3 | Load cases | **Single** load case only (the Study 1 set-up) | §4.4(e) |
| D4 | Objective weights | Equal by default; **user-tunable hyper-parameters** | §5.1 |
| D5 | Cluster | **GH200** GPUs, **SLURM**, **24 h** wall-time limit. Working practice is covered by Owen's own instructions. | §4.2, §7.6, §10.4 |
| D6 | Inset lower bound (v1.2) | a ≥ a^{lo}(d) = smoothed (t_min − d)⁺ instead of a ≥ t_min, so walls can lie wholly outside Γ₀. | §3.4, §6.1 |
| D7 | FE engine (v1.2) | Research and build the fast FE first; production engine = active-set q1_im + cuDSS Cholesky, buckling on the same factor at 0.4 mm. | §4.1–4.2, §6.3–6.4 |

**Still open.** Implement the default, record the question in `continuous/OPEN_QUESTIONS.md`, and do not block:

1. **Buckling knock-down.** Default κ_b = 0.85.
2. **Escape holes.** Diameter 3 mm and 2 per void (vshell defaults).
3. **Plastic capacity** (Stage 3e). Worth doing before a physical test campaign?
4. **Reference-code issues** found during implementation, starting with the q1_im buckling sensitivity of §6.4.

---

## Appendix A. Reference GPU-FE stack: what is there and what was checked

### A.1 `gpu_analysis.py` (numba-cuda, optional JAX)

| Function | Notes |
| --- | --- |
| `matrix_free_pcg(nx, ny, nz, hx, hy, hz, E_elem_h, nu_elem_h, b_h, fixed_mask_h, max_iters, tol, rel_tol, verbose, u0, …, element_type, nu, preconditioner, gmg_levels, gmg_smoother_steps, gmg_coarse_max_nodes, gmg_reuse_max_rel_change, pcg_engine, pcg_check_every)` | Returns (u, history, cg_info). nx, ny, nz are **node** counts. Convergence is ‖r‖ ≤ max(tol, rel_tol‖b‖). Loads on fixed DOF are stripped. Warm start has an automatic cold-restart guard. |
| `compute_element_strain_energy(nx, ny, nz, hx, hy, hz, E_elem_h, nu_elem_h, u_h, element_type, nu)` | Verify its normalisation (½uᵀKu or uᵀKu; total or density) with a unit test before use. |
| `compute_element_stress(nx, ny, nz, hx, hy, hz, E_elem_h, nu_elem_h, u_h)` | Returns (N_e, 6) element-averaged Voigt stresses (xx, yy, zz, yz, xz, xy), scaled by E_elem. |
| `element_khat(element_type, h, nu)`, `hex_reference_data`, `hex_dof_map`, `lame_from_E_nu` | Element matrices and helpers. `ELEMENT_TYPES = ('q1_gauss', 'q1', 'q1_im')`. |
| `GMGPreconditioner`, `GMGFastPreconditioner` | Galerkin coarse operators, Chebyshev smoother. |
| `JaxSolver`, `pcg(...)` | Alternative engines. |
| `probe()`, `elastic_choice()`, `pcg_choice()`, `pcg_jax_memory()`, `buckling_memory()`, `buckling_defaults()` | Engine selection and memory estimates. |
| `verify_host_kernel_consistency()` | Offline consistency check between the host and CUDA paths. |

### A.2 `analysis.py`

- **`ConjugateGradient(iter_name, printouts ∈ {'Silent','Minimal','Normal','Verbose'}, data_dir, youngs_modulus [Pa], poisson_ratio, voxel_size_mm, pcg_max_iters, pcg_tol [N], pcg_rel_tol, save_meshes, indent, csg_domain_data, warm_start_cg, warm_start_max_rho_change, warm_start_probe_iters, warm_start_min_reduction, warm_start_state, element_type='q1', preconditioner='gmg', gmg_levels, gmg_smoother_steps, gmg_coarse_max_nodes, gmg_reuse_max_rel_change, analysis_refinement=1, engine='auto', pcg_engine='auto', pcg_check_every)`.**
- **`run_analysis(field)`** takes `field`, a 3D array (n_x,e, n_y,e, n_z,e) of the stiffness density that scales λ and μ. It applies **no floor**: pass ρ̂. It returns `(voxel_sed [J/m³], mesh_volume [m³, count of ρ ≥ 0.5])` and sets `.compliance` [J].
- **`run_buckling(num_modes=6, aggregation_p=8.0, prev_modes=None, min_energy_fraction=0.5, mode_continuity_warning_threshold=0.7, min_density_for_geometric_stiffness=0.1, verify_against_gpu=False, eigensolver_tol=1e-12, engine=None, lobpcg_tol=None)`.** It must follow `run_analysis` and needs `analysis_refinement == 1`. It returns `BucklingResults`.
- **`csg_domain_data`** keys are `fixed_nodes` (bool, element-wise, (n_x,e, n_y,e, n_z,e, 3)) and `load_vectors` (float N, element-wise, smeared /8 to the corners).
- **`Regular` and `Conformal`** are CPU (scikit-fem) back ends; `Conformal` needs TetGen. Useful for small cross-checks only.
- **`validate_buckling_sensitivity_regular(...)`** at the bottom of the module is a finite-difference harness worth copying for §6.6.

### A.3 Referenced but not supplied

These are referenced only in docstrings and comments, except `interfaces.py`, which is now stubbed:

- `interfaces.py` (only `BucklingResults` is imported, and that is now stubbed);
- `docs/theory.md`, `docs/README.md`, `buckling_implementation_plan.md`;
- `pcls_optimiser.py`, `problem_parser.py`, `optimise.py`;
- `element_matrices.py`;
- `utils/cpu_backend.py`, `utils/validate_buckling_fd.py`.

If behaviour depends on them, ask Owen (OPEN_QUESTIONS.md).

### A.4 Checks run while writing this document

These ran on a CPU-only cloud VM: Python 3.13, numba 0.68, no CUDA.

1. With the `interfaces.py` stub, both modules import. `analysis.ConjugateGradient` selects `cpu_direct` for elastic solves and `cpu_cholesky_lobpcg` for buckling when no GPU is present. That CPU path is what the CPU test suite should use.
2. A 40×4×4-element (1 mm) clamped cantilever, q1_im, 10 N tip load, gave C = 0.0500 J. Euler–Bernoulli gives 0.0568 J at L = 40 mm, but about 0.0507 J at the ≈ 38.5 mm lever arm left by the one-element clamped slab and the load smeared over the last element. Shear adds about 1 %. So the FE agrees with the effective-length beam to about 2 %.
3. Compliance sensitivity: central finite differences matched −2·sed·V/ρ̂ to 4·10⁻⁴ relative.
4. Buckling sensitivity, directional finite differences (§6.4): exact to 9·10⁻⁵ with `q1`, and a 5.8 % bias with `q1_im` because Term A and the adjoint term use the Gauss stiffness. cshell must correct this.

---

## Appendix B. Analytic benchmarks

**B1/B2: hollowing a cantilever rod at constant mass (Timoshenko).**

- **Section.** Solid rod r₀ = 5 mm, PA11, tip load. Hollow design with a = t_min = 1.5 mm (R_i = 3.5) and mass kept: R_o = √(r₀² + R_i²) = 6.103 mm, I/I₀ = 1.98.
- **Shear coefficient** (Cowper 1966), κ = 6(1+ν)(1+m²)² / [(7+6ν)(1+m²)² + (20+12ν)m²] with m = R_i/R_o: 0.895 (solid) and 0.601 (hollow).
- **Compliance.** δ/P = L³/(3EI) + L/(κGA).

| L (mm) | L/D | C_hollow / C_solid |
| --- | --- | --- |
| 5 | 0.5 | 1.198 (solid better) |
| 10 | 1 | 0.872 |
| 20 | 2 | 0.633 |
| 40 | 4 | 0.540 |
| 80 | 8 | 0.514 |

The stiffness crossover is at L/D ≈ 0.77. A plain round rod loses stiffness to hollowing only when it is very stubby. Joint losses in real parts therefore come from capacity: thin walls carrying transverse shear (τ_max = 2V/A for a thin tube against 4V/3A for the solid, so 1.5× higher), as-built thinning, and local collapse. B1 (L/D = 0.5) must stay solid on stiffness alone. Add a capacity-weighted variant with a short, deep, shear-loaded block, and record what the optimiser does; there is no closed-form answer for it.

**B3: clamped–free column.** r₀ = 5 mm, L = 100 mm, axial compression.

- Solid: P_cr = π²EI₀/(4L²) = 213 N, against 2870 N to reach σ_f. Buckling governs.
- Uniform hollow at a = 1.5: I = 1.98 I₀, so P_cr ≈ 422 N.
- The continuous optimum should reach λ_b ≥ the best uniform tube found by a brute-force sweep over a with the same FE.
- For the optimal clamped–free column, Keller (1960) shows tapering toward the free end raises P_cr by up to a third at equal volume for geometrically similar sections. Expect some taper in d. Report the gain over uniform rather than enforce a value.

**B4: MMA.** Svanberg's 2007 note gives two small test problems with published optima; reproduce them.

---

## Appendix C. References

- Peckham O., Hicks B., Goudswaard M. *Hollow Promises?* Manuscript, 23 October 2026 (`continuous/reference/main.pdf`).
- Allaire G., Jouve F., Toader A.-M. Structural optimization using sensitivity analysis and a level-set method. *J. Comput. Phys.* 194 (2004) 363–393.
- Allaire G., Jouve F., Michailidis G. Thickness control in structural optimization via a level set method. *Struct. Multidisc. Optim.* 53 (2016) 1349–1382.
- Lazarov B.S., Sigmund O. Filters in topology optimization based on Helmholtz-type differential equations. *Int. J. Numer. Meth. Eng.* 86 (2011) 765–781.
- Hojjat M., Stavropoulou E., Bletzinger K.-U. The Vertex Morphing method for node-based shape optimization. *Comput. Methods Appl. Mech. Eng.* 268 (2014) 494–513.
- Wang F., Lazarov B.S., Sigmund O. On projection methods, convergence and robust formulations in topology optimization. *Struct. Multidisc. Optim.* 43 (2011) 767–784.
- Svanberg K. The method of moving asymptotes: a new method for structural optimization. *Int. J. Numer. Meth. Eng.* 24 (1987) 359–373; and Svanberg K. *MMA and GCMMA: two methods for nonlinear optimization* (KTH note, 2007).
- Le C., Norato J., Bruns T., Ha C., Tortorelli D. Stress-based topology optimization for continua. *Struct. Multidisc. Optim.* 41 (2010) 605–620.
- Ferrari F., Sigmund O. Revisiting topology optimization with buckling constraints. *Struct. Multidisc. Optim.* 59 (2019) 1401–1415.
- Aage N., Andreassen E., Lazarov B.S., Sigmund O. Giga-voxel computational morphogenesis for structural design. *Nature* 550 (2017) 84–86.
- Liu H., Hu Y., Zhu B., Matusik W., Sifakis E. Narrow-band topology optimization on a sparsely populated grid. *ACM Trans. Graph.* 37(6) (2018) 251.
- Clausen A., Aage N., Sigmund O. Topology optimization of coated structures and material interface problems. *Comput. Methods Appl. Mech. Eng.* 290 (2015) 524–541.
- Wu J., Aage N., Westermann R., Sigmund O. Infill optimization for additive manufacturing: approaching bone-like porous structures. *IEEE Trans. Vis. Comput. Graph.* 24 (2018) 1127–1140.
- Keller J.B. The shape of the strongest column. *Arch. Ration. Mech. Anal.* 5 (1960) 275–285.
- Cowper G.R. The shear coefficient in Timoshenko's beam theory. *J. Appl. Mech.* 33 (1966) 335–340.
- NASA SP-8007. Buckling of thin-walled circular cylinders (1968; revised 2020).
