# Investigation: assumptions, inputs, units

Code: branch worktree-cshell-solid-pockmarks @ 9781a24, library read-only. Units mm, N, MPa, N·mm.
Material (all runs): E 1760 MPa, ν 0.42, ρ 1.04 g/cm³, strength 36.6 MPa (except the "wbuck" runs, see below).

## Decomposition (decompose.py, decompose.sh)
- Inputs: runs_cshell/SR_VAL2/<part> (config_used.json, design_solid_phase{0,1}.npz); dshell segmentation of the
  same STEP from examples/dshell/runs/<part>/arrays.npz (seg_xyz/seg_label on a 1 mm grid, skeleton points and
  regions.csv). The two models share the STEP frame (bounding boxes agree to < 1 mm, logged). Elements are labelled by
  the nearest segmentation voxel (p99 distance 1.8 mm).
- FE: the optimiser's own PhaseModel, phase 1 grid (h_e 0.4 mm, two-grid PCG, q1_im), as-built ρ. Solid = ρ0 of the
  run; final = the SR_VAL2 phase-1 design with its hollow-allowed set H.
- Section decomposition: per skeleton station (Voronoi of the region's skeleton points, ~1 mm slabs), weighted LS fit
  σ_tt = a + b·y + c·z (weights ρ̂h³, centred section coordinates). Axial = a (sign: tension/compression), bending =
  linear part, residual = rest; shear = τ_t1² + τ_t2² (transverse shear + torsion). Energies use σ²/2E and τ²/2G of the
  material centre stress scaled by ρ̂; "other" = centre-stress energy − parts (transverse normal stresses, in-section
  shear, Poisson coupling). Meaningful for struts; joints/mounts are reported but the beam split is not physical there.
- Centre-stress energy vs exact ½uᵀKu: 0.995–0.997 (all parts, both designs).
- Buckling participation: retained modes of the optimiser's eigen-solve (6 modes, k_σ ramp at ρ̂ 0.1, k_σ masked
  within exclude_bc_mm of clamps). Per mode, modal strain-energy share φ_eᵀK_eφ_e (sums to 1, checked) and the
  destabilising term c_e = −detJ Σ_g σ_g·Γ_g(φ) (Σ c_e = μ_j; p-norm of μ_j reproduces the optimiser's μ̃ to 1e-12).
- Wall-normal chord split (solid only): σ sampled along x − s n, s ∈ [0, 2 r_MA] (24 points), in-plane stresses
  fitted linear in s (mean = membrane, slope = bending); compared with the optimiser's R_v (half depth, lumped).
- q_entry = area-weighted quantile of the phase-0 solid R_v, i.e. the growth step at which the vertex enters H.

## Runs (runs_cshell/INV/<name>, launched with cshell/cluster/launch_part.sh, single --set)
- cantilever_allq: current formulation, growth.patience = 99 (all 10 quantiles).
- weighted_<part>: objective.type = weighted (Stage 2 / dshell-like: J = ½ ln K/K0 + ½ ln λ/λ0, λ = soft-min of
  strength p-norm and global buckling, knockdown 0.85). In this mode the optimiser drops the MMA constraints (only mass
  is constrained); floors (0.85) are NOT enforced by the optimiser (`floors_as_constraints` is not read by optimise.py).
- cantilever_wbuck50, lbracket_wbuck50: weighted with material.strength_MPa = 1e9, so λ ≈ λ_B (buckling only in the
  objective). Deviation from the brief's (a): strength is not a constraint in weighted mode (cannot be by config);
  σ_PN/σ_PN0 is reported afterwards instead.
- Buckling margin λ_B ≥ 1.5 λ_B0: not supported by config (no margin key); not run.
- finalise decides on J = ln K/K0 and feasibility with respect to the solid for every run (including weighted ones), so
  "solid kept" for weighted runs means K/K0 < 1.02 or a constraint violated, not that the weighted J failed.

## Not checked
Mesh convergence of the decomposition (one grid, h_e 0.4 mm); sensitivity to the segmentation; nonlinear
(geometric/material) buckling; body-fitted validation of the INV designs.

## Incident
A first launch passed several `--set` flags; cli.py's `--set` (nargs="*", no append) keeps only the last, so those
six jobs ran wrong configs into runs_cshell/{Cantilever,L-bracket,MBB}. They were killed; two continuations
(7927554, 7927555) that run.sh's TERM trap submitted could not be cancelled from this session. Those three directories
are scratch and can be deleted.
