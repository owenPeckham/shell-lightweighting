# Dimples and hollowness in Stage 2 cshell designs: diagnosis (2026-10-31, MBB)

Scripts: this folder. Outputs and renders: `/projects/b6bl/owenpeckham.b6bl/home_data/shell-lightweighting/diag_pockmarks/`
(`dimples/patch_<k>*.png`, `dimples2/variants_<k>.png`, `*_report.json`, `hollow/`). GPU use 0.34 GPU-h (jobs 7900803, 7900910, 7901210).

## Defect 1: dimples (outer and void surfaces) come from the fields, not the export

- **Export is faithful:** the outer STL matches the model P(d+δ) to ≤ 6 µm (p99). Detrended roughness is the same for the STL, export stage A and P d (41.8 vs 41.9 µm in one patch).
- **Dimple geometry:** spacing 0.43–0.65 mm, close to the Γ_h edge length of 0.57 mm. Depth: median about 30 µm outside, up to 0.1 mm; median 0.07–0.38 mm on the void surface.
- **Outer surface:**
  - The 2-D Helmholtz filter has a log-singular kernel, so a one-vertex spike comes out 2.04 times its 1-ring (1.08 with the filter applied twice).
  - Raw-ξ outliers therefore survive as cones of about 0.10 mm (× d_hi = 8 mm).
  - P is linear on each Γ_h triangle, so the surface shows facets with creases along Γ_h edges.
- **Void surface:**
  - a_hi = r_MA + γ h_g uses the per-vertex r_MA unfiltered, and r_MA is very rough (interior high-pass RMS 0.36 mm).
  - Secondary: the stage-B opening sets removed void to the constant −h/2, which leaves terraces one grid cell high where the void is thinner than 3 mm.
- **Passive borders:** d is forced to 0 after the filter, leaving cliffs of full height d with sawtooth rims.
- **Ruled out:** ambiguous-face flips, the mass shift δ, and the a_used bias. The igl sign check failed inside igl, but residuals ≤ 6 µm argue against sign errors.

Field-only measurements of the candidate fixes:

| Fix | Effect |
|---|---|
| Filter twice (F²) | High-pass RMS of d 14.3 → 4.7 µm |
| Smooth r_MA (about 1 mm) | High-pass RMS of a 22.4 → 6.2 µm |
| Both | High-pass RMS of a 3.7 µm; worst void patch 0.42 → 0.15 mm RMS |

## Defect 2: hollowness is what the objective prefers, not a bug or a barrier

- **Re-run of the vshell phase 0** (matches the log):
  - Closed share is 91.6 % at the start, 31 % by iteration 4 and 7.7 % by iteration 9.
  - Capacity is 99.8 % buckling-governed. In the solid core the marginal value per unit volume is 0.62 against a multiplier of 1.37, so every core vertex favours moving material into the outward offset d.
  - The sliver weight is 0 in phase 0, so no barrier is involved.
- **Closure is fragile:** a vertex closes only when ξ̃ᵃ ≳ 0.97–0.98. Final closed share is 0.087 % by a ≥ r_MA and 0 % by the export rule.
- **Forced closed regions, re-optimised in the phase-0 model** (free result J = 0.490):

| Closed region | J |
|---|---|
| Thickest 10 % (r_MA) | 0.405 |
| Thickest 25 % (r_MA) | 0.299 |
| Within 5 mm of mounts | 0.482 |

## Limitations

- Only the MBB was analysed.
- The fixes were not run end to end.
