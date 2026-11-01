# Why the Cantilever and L-bracket results contradicted Owen's predictions (2026-11-01)

Owen predicted:
- the L-bracket stays mostly solid (tension and shear);
- the Cantilever's lower compression struts go hollow, as dshell does, and its upper struts go hollow too.

Scripts are in this folder (`decompose.py`, `decompose.sh`, `ASSUMPTIONS.md`). Outputs are in `/projects/b6bl/owenpeckham.b6bl/home_data/shell-lightweighting/runs_cshell/INV/`:
- `decomp/<case>/` for the member decomposition;
- `cantilever_allq/`, `weighted_<part>/`, `cantilever_wbuck50/` and `lbracket_wbuck50/` for the runs.

All results are on one grid (h_e 0.4 mm).

## Member energy decomposition (solid)

Energy columns are each member's own shares; mode-1 share is the member's part of the destabilising term in buckling mode 1.

| Part / member | Share of part energy | Axial + | Axial − | Bending | Shear | Mode-1 share |
|---|---|---|---|---|---|---|
| Cantilever upper struts 5, 6 | 0.13 each | 0.92 | 0 | 0.04 | 0.05 | ≈ 0 |
| Cantilever lower struts 7, 8 | 0.20 each | 0 | 0.97 | 0.01 | 0.03 | 0.45 each |
| L-bracket arms 8, 9 | 0.11 each | 0.80 | 0 | 0.15 | 0.05 | 0 |
| L-bracket strut 7 | 0.07 | 0 | 0.89 | 0.06 | 0.05 | 0.80 |
| L-bracket light struts and joints | 0.10 total | | | | | 0 |

- **Cantilever:** a truss. 67 % of its energy is axial and 8 % bending.
- **MBB:** the chords are 80–98 % axial. The zero-force members carry under 0.1 % of the energy.

## Hypotheses

- **H1: supported.** Compliance at constant mass rewards only bending (EI); hollowing leaves EA unchanged.
  - λ_B ≥ λ_B0 was active on every SR_VAL2 iteration, so extra buckling margin earns nothing.
  - What compliance does gain comes from moving mass between members (emptying low-energy members) and from the outward offset.
- **H2: mostly not supported.** The L-bracket arms are tension-dominated (15 % bending). Its gain comes mainly from emptying its light members and joints.
- **H3: supported for the order, but the order hardly matters.** R lumps axial with bending, so compression struts enter the allowed set early. Being allowed is not being hollowed: the Cantilever's lower struts were 72 % allowed but 24 % open. dshell's R lumps axial in the same way.

## Runs

| Run | Decision | K/K0 | λ_B/λ_B0 | σ_PN/σ_PN0 | Closed / hollow | Cantilever lower struts open |
|---|---|---|---|---|---|---|
| SR_VAL2 compliance, Cantilever (patience 2) | solid kept | 1.017 | 1.00 | 0.94 | 84 / 11 % | 0.24 |
| Compliance, Cantilever, all quantiles | **accepted**, J 0.059 | 1.061 | 1.08 | 0.99 | – | 0.11–0.30 |
| Weighted (dshell-style), Cantilever | solid kept\* (weighted J 0.496) | 0.983 | 2.76 | 1.10 | 2 / 66 % | **1.00** |
| Weighted, L-bracket | accepted | 1.105 | 1.11 | 0.94 | 37 / 32 % | – |
| Weighted, MBB | accepted | 1.113 | 2.44 | 0.88 | 2 / 77 % | – |

\* finalise judges every objective by the compliance rule (ln K/K0 > ln 1.02, σ_PN ≤ σ_PN0). By dshell's rule the weighted Cantilever passes: J 0.496 against dshell's 0.479 (K −0.8 %, P +163 %).

## Defects found

- finalise decides on the compliance rule whatever the objective.
- summary.txt mixes the weighted J and ln K/K0.
- `floors_as_constraints` is never read, so the weighted mode has only the mass constraint.
- The buckling-margin constraint is not implemented.
  - It would go in `objective.evaluate_compliance` (g_buckling += ln m_B) and in `optimise.evaluate`.
  - Estimated K/K0 at m_B = 1.5: about 1.00–1.04 (Cantilever) and 1.08–1.12 (L-bracket).
- Local shell buckling cannot govern: walls span 4–5 elements and a hand estimate gives λ_L ≈ 50–100. Not checked by refinement.

## Conclusions

- Owen's Cantilever prediction (and dshell's result) is reproduced only when buckling capacity is rewarded, through the weighted objective or a buckling margin.
- No objective tested keeps the L-bracket mostly solid: emptying its light members raises both its stiffness and its strength.
- Stiffness cost of the weighted objective against the best compliance design:

| Part | Cost in K | Buckling gain |
|---|---|---|
| Cantilever | 7 % | ×2.76 |
| L-bracket | 3 % | ×1.11 |
| MBB | 12 % | ×2.44 |
