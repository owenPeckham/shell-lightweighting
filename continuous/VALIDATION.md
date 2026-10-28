# cshell: verification and validation results (THEORY §6.6, §11)

Status per test ID. Engine: active-set q1_im + cuDSS (`cshell/fe_sparse.py`) unless stated. Job ids in parentheses;
logs in `logs/`, scripts in `continuous/research/` and `tests/cshell/`.

## V1 geometry

| ID | Test | Result | Pass |
| --- | --- | --- | --- |
| V1-a | Solid cylinder / tube volume through ψ at h_g = 0.1 | < 2e-3 (`test_band.py`); sphere via the design chain < 2e-3 (`test_design.py`) | ✅ |
| V1-b | r_MA on cylinder r = 5, 10 × 20 box | ±0.1 h_g (`test_surface.py`, job 7795055) | ✅ |
| V1-c | All-solid via fields vs ψ = φ | volume < 5e-4 (`test_design.py`); compliance half: pending | ◐ |
| V1-d | Filter maximum principle, M-symmetry, adjoint | round-off (`test_surface.py`); max principle [−0.01, 1.01] on the cantilever Γ_h (1 % negative cotangent edges) | ✅ |

## V2 FE validation

**V2-a: solid baselines vs vshell, h_e = 0.8 mm, same BCs** (`research/stage0/bench_solid.py`, job 7794574)

| Part | K₀ cshell (N/mm) | K₀ vshell (N/mm) | Δ | K₀ at 0.4 mm |
| --- | --- | --- | --- | --- |
| Cantilever | 257.8 | 256.8 | +0.4 % | 252.3 |
| L-bracket | 48.60 | 48.3 | +0.6 % | 47.90 |
| MBB | 473.5 | 473 | +0.1 % | 467.2 |

Pass (criterion 2 %). The dense reference solver reads 2–7 % stiffer (OPEN_QUESTIONS B5).

**V2-b: Study 2 PA11 tubes, 3-point bending vs Abaqus** (`research/stage0/tubes.py`, job 7795448; quarter model,
checked equal to the full model; load pin as a 4 mm patch, supports one node plane)

| Tube | wall (mm) | Abaqus (kN/mm) | h = 0.8: k, Δ | h = 0.4: k, Δ | ratio to A (0.4) vs Abaqus ratio |
| --- | --- | --- | --- | --- | --- |
| A | rod | 0.0086 | 0.00878, +2.1 % | 0.00876, +1.8 % | 1.000 / 1.000 |
| B | 5.4 | 0.0104 | 0.01066, +2.5 % | 0.01063, +2.3 % | 1.215 / 1.209 (+0.5 %) |
| C | 4.0 | 0.0158 | 0.01609, +1.8 % | 0.01624, +2.8 % | 1.855 / 1.837 (+1.0 %) |
| D | 3.1 | 0.0245 | 0.02485, +1.4 % | 0.02537, +3.6 % | 2.898 / 2.849 (+1.7 %) |
| E | 2.5 | 0.0351 | 0.03697, +5.3 % | 0.03673, +4.6 % | 4.195 / 4.081 (+2.8 %) |

A–D pass at both resolutions; **E misses the 4 % absolute gate** (+4.6 % at 0.4 mm) but its stiffness ratio is within
2.8 %. vshell's own E result was +4.5 % (OPEN_QUESTIONS Q4); Abaqus itself is 12 % below Euler–Bernoulli at E and
over-predicts the measured tubes by 35–60 % (Q6). Earlier runs with a one-node-line load gave E −11 % at 0.4 mm
(singular line load); with a 3 mm-wide clamped support patch the beam was partially clamped (+41 %): both modelling
errors, now fixed.

**V2-c: thin-wall convergence (1.7 mm wall, OD 25)**: k = 0.02669 (0.8 mm) → 0.02627 (0.4 mm), −1.6 %, monotone.
0.2 mm does not fit as one fp64 cuDSS factor even in hybrid memory mode (6.4 M DOF; STATUS K12): **pending**.

## V3 buckling

| ID | Test | Result | Pass |
| --- | --- | --- | --- |
| V3-a | Clamped–free rod r = 5, L = 100: P_cr = 213.2 N | 0.4 mm: 215.7 N (+1.2 %); 0.8 mm: +2.4 % (job 7794919) | ✅ |
| V3-b | Tube R_m = 10, t = 1.5, L = 60 vs ccx C3D10 | in progress (validate agent) | … |

## V4 sensitivities (§6.6)

| Item | Test | Result | Pass |
| --- | --- | --- | --- |
| 1 | Geometry chain, directional + Taylor | < 1e-6, O(h²) remainder, nominal and as-built (`test_design.py`) | ✅ |
| 2 | Filter | `test_surface.py` | ✅ |
| 3 | Compliance, element-wise FD | < 1e-4 float64 (`test_fe_sparse.py`, `test_fea.py` host path) | ✅ |
| 4 | Stress p-norm adjoint, notched bar | < 1e-3 (`test_objective.py`) | ✅ |
| 5 | Buckling directional | sparse engine q1_im 3.4e-5; reference + cshell fix 1.2e-4 (unfixed 5.3e-2) (`test_buckling.py`) | ✅ |
| 6 | End to end dJ/dξ, B3 rod at 0.8 mm | compliance-only 3.3e-5, stress 8.2e-5, full (with buckling) 7.6e-3 at h = 0.125 (< 2 %) (job 7795349) | ✅ |
| 7 | Mass restoration | < 1e-4 in ≤ 20 evaluations (`test_mma.py`) | ✅ |

## V5 benchmarks

| ID | Test | Result | Pass |
| --- | --- | --- | --- |
| B1 | Stub rod L/D = 0.5 from all-hollow S1 | J −0.011 → −0.0003 ≤ ln 1.02, G_s → 0, converged at it 44 (job 7795325); **closed area 100 %** of the active area, d_max 0.05 mm (job 7795487). Caveat: only 189 of 1109 vertices are free (1.5 mm mount zones at both ends of a 5 mm rod plus the thin rule). | ✅ |
| B2 | Slender rod L/D = 8 | J = 1.178, C/C₀ = 0.308, hollow everywhere (closed 0.1 %), D_comp extended 4 → 8 mm once (§3.5) (job 7795581). **Tapered**: wall 4.4 mm at the root → 1.5 mm mid-span (1.8 at the load pad), outer radius 6.8 → 5.1 mm (`research/bench/wall_profile.py`, job 7795953). Constant-mass uniform shell, analytic: C/C₀ ≈ 0.35 (wall wholly outside, D1) / 0.514 (a ≥ t_min). FE uniform-shell sweep: pending (export). | ✅ |
| B3 | Column | pending | … |
| B4 | MMA on Svanberg's problems | reference optima (`test_mma.py`) | ✅ |
