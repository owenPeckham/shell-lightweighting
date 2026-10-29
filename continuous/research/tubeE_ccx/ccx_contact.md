# Tube E / A: CalculiX reproduction of the Abaqus roller-contact 3-point bend

Script: tube_ccx.py (deck + ccx driver + parser), run.sh (sbatch, CPU), collect.py / collect.sh -> summary.json.
Jobs: 7800452 (E L1 timing), 7800523 (E K/NLGEOM/L2, A L2), 7808290 (faceting, patch-width, A L1), 7808617 (collect). Peak RSS 24.7 GB (A L2 contact); L1 runs need 2-6 GB.
Set-up and assumptions: ASSUMPTIONS.md. k in kN/mm. Rod-A runs in runs/A_L1_*_badmesh used a broken butterfly mesh
(core/ring not merged; area -3 %) and are void.

## k (kN/mm)
| | Abaqus | ccx our BCs (patch 4 mm, node-plane supports) | ccx rollers, small strain | ccx rollers, NLGEOM | voxel |
|---|---|---|---|---|---|
| E | 0.0351 | 0.03687 (L1) / 0.03685 (L2) | 0.03631 (L1) / 0.03629 (L2) | 0.03592 (L1) | 0.0368 |
| A | 0.0086 | 0.008738 (L1) / 0.008743 (L2) | 0.00911 (L1) / 0.00911 (L2) | 0.00883 (L1) | 0.00876 |

Contact k is the secant F/d of the load roller at F = 100 N, interpolated between the time points (the curve is in result.json).
Checks:
- Penalty K 1e5 -> 1e6 changes E by +0.2 %.
- Mesh L1 -> L2: E contact changes 0.06 %, E patch 0.05 %, A patch 0.06 %.
- Energy W is within 1.2 % of F d / 2. That is consistent with a stiffening contact curve.
- The total support reaction equals the roller load.

## Findings
1. Same BCs, same answer. ccx C3D8I with the voxel's patch and node-plane BCs reproduces the voxel k to within 0.15 % for E and 0.2 % for A, at both meshes. So the voxel thin-wall discretisation is not the source of the gap.
2. Roller contact is not the whole story. For E, rigid R25/R10 rollers lower k by 1.5 % (small strain), and NLGEOM lowers it by a further 1.1 %, for −2.5 % in total relative to our BCs. For solid A, rollers *raise* k: +4.2 % small strain, +1.0 % NLGEOM. Most of that comes from the support contact point migrating inward and the reaction tilting at A's 12 mm deflection.
3. With the full physical model (rollers + NLGEOM), Abaqus sits **uniformly** below ccx: E −2.3 %, A −2.6 %. The wall-thickness trend in the voxel–Abaqus gap is therefore explained by the roller-contact physics. In thin walls, local ovalisation/indentation under a ~2 × 1.2 mm contact softens the tube; in the solid rod, contact migration stiffens it. The residual is a size-independent ~2.5 %. It is not explained by anything in the published set-up. A plausible cause is coarse circumferential faceting in the Abaqus mesh: on our patch model, 24 elements round the circumference gives −1.7 % and 32 gives −1.3 %. That is the analytic polygon-I deficit, and it is equal for A and E. This explanation is untested.
4. Contact footprint at 100 N on the load roller (E, L2): half-length 1.0–1.25 mm along the tube and half-width about 0.6 mm round it. That is a ~2.2 × 1.2 mm ellipse; Hertz for a solid section gives a ≈ 0.85 mm. No patch width on our BCs reproduces the contact k. The ccx L2 patch gives 0.03670 / 0.03673 / 0.03678 at pw 0.5 / 1 / 2 mm, because the patch's transverse extent (±3.5 mm of arc) and the node-plane supports dominate. The roller effect is concentrated ring loading, and an axial patch width cannot mimic it. The voxel pw sweep (pw_voxel.py / run_voxel.sh) was therefore not run.
5. Small load: the secant k at 3–10 N is 3.4–5.7 % below the 100 N value for E. This is Hertzian contact (δ ∝ F^(2/3)), so "k" for roller tests depends on the load level.

## For the optimiser
Study 1 parts are loaded through bolt/pad faces (clamped supports, load shared over the faces). These match cshell's BC model, not rollers. The tube discrepancy comes from roller-contact physics, which those parts do not have, plus a uniform ~2.5 % Abaqus offset. Neither implies a change to cshell's FE or BC model. Voxel thin-wall stiffness (2.5 mm walls) agrees with a body-fitted C3D8I to 0.15 %. Do not use Study 2 tube k as a < 3 % validation target unless rollers + NLGEOM are modelled.
