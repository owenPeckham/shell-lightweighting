# Manuscript numbers for cshell validation and reporting

Sources: M = continuous/reference/main.pdf ("Hollow Promises?", PDF page = printed page; 41 pp.), V = continuous/reference/vshell_theory_and_operation.md, T = continuous/THEORY_continuous.md, tests/validate_tubes.py + validate_tubes_results.json (J). Stiffness in kN/mm, load capacity in kN unless stated. n = 5 replicates per type in every table (M p.12, p.14).

## 1. Material and process (M p.11, p.14)
| Quantity | Value | Source |
| --- | --- | --- |
| PA11 (MJF) E, nu, rho | 1.76 GPa, 0.42, 1.04 g/cm3 | M p.11 (3.1.4) |
| Al 6063 E, nu | 68.9 GPa, 0.33 | M p.14 (3.2.2) |
| MJF wall range | printable >= 0.8 mm, structural 1.5-3.0 mm | M p.11 |
| Escape holes | 3 mm dia; cantilever 8, L-bracket 8, MBB 16; placed in low-stress regions | M p.11-12 |
| Print direction | Study 1: solid and hollow of one topology share build direction; tubes axis parallel to build direction | M p.11, p.15 |

## 2. Study 1 set-up (M pp.7-12)
| Item | Value | Source |
| --- | --- | --- |
| GD problems | cantilever, L-bracket, MBB (MBB mirrored to full beam); objective max stiffness, safety factor 1.00 at 0.2 kN | M p.7 |
| GD mass targets | cantilever 60 g, L-bracket 45 g, MBB 45 g (as built solid means 60.9, 43.1, 87.3 g, Table 4; MBB is the mirrored full beam) | M p.7, p.19 |
| Shelling | outward constant-mass shell, Nelder-Mead on thickness, mass tol +-1% | M p.8 |
| FE | Abaqus C3D10, refinement to 1% max von Mises tol; hollow mesh size < 1/3 min wall (Kim, Jang, Choi); load 0.2 kN; same constrained DOF as in GD | M p.9 |
| BC bounds | Min stiffness (A): linear continuum constraints on bolt-hole inside walls only. Max stiffness (B): kinematic coupling on hole walls plus region under washers | M p.9-10, Fig 6 |
| Test | Instron 8872, 5 kN cell; cantilever and L-bracket in thick Al jigs, MBB bolted to base plate; 10 mm radius ball-nose probe (cantilever, L-bracket); MBB 10 mm x 120 mm "fin" onto two pressure pads; 0.1 mm/s to failure or 25 mm | M p.12 |
| Load capacity definition | peak load of the load-deflection curve (Fig 9 up to 25 mm); not stated explicitly, unverified | M p.18 |

## 3. Study 1 results
### Table 2 (M p.16): Abaqus bounds
| Artefact | Min K | Max K | Max % change solid to hollow |
| --- | --- | --- | --- |
| Solid cantilever | 0.163 | 0.268 | |
| Hollow cantilever | 0.156 | 0.237 | -11.6 |
| Solid L-bracket | 0.025 | 0.038 | |
| Hollow L-bracket | 0.026 | 0.037 | +4.0 |
| Solid MBB | 0.364 | 0.899 | |
| Hollow MBB | 0.334 | 0.839 | -8.2 |

Bound width MBB 0.899/0.364 = +147% (text says 147-151%, p.26/29).

### Table 3 (M p.17): measured, n = 5
| Artefact | Mean K | K SD | K % | Mean Pcap | Pcap SD | Pcap % |
| --- | --- | --- | --- | --- | --- | --- |
| Solid cantilever | 0.193 | 0.021 | +3.5 | 0.513 | 0.013 | +61.1 |
| Hollow cantilever | 0.200 | 0.015 | | 0.826 | 0.032 | |
| Solid L-bracket | 0.031 | 0.003 | -7.0 | 0.453 | 0.021 | -11.2 |
| Hollow L-bracket | 0.029 | 0.000 | | 0.403 | 0.010 | |
| Solid MBB | 0.602 | 0.047 | -14.9 | 2.436 | 0.053 | +14.8 |
| Hollow MBB | 0.512 | 0.098 | | 2.796 | 0.045 | |

All measured solid and hollow K lie inside the Abaqus bounds (M p.16, p.29).

### Significance (Welch t-test, Hedges g) (M pp.16-17, 19)
| Result | Statistic | Significant? |
| --- | --- | --- |
| Cantilever K +3.5% | t(7.27) = -0.60, p = 0.569, g = -0.34 | no |
| L-bracket K -7.0% | t(4.23) = 1.80, p = 0.142, g = 1.03 | no |
| MBB K -14.9% | t(5.72) = 1.85, p = 0.117, g = 1.05 | no |
| Cantilever Pcap +61.1% | t(5.18) = -20.15, p = 4.06e-6, g = -11.51 | yes |
| MBB Pcap +14.8% | t(7.81) = -11.63, p = 3.31e-6, g = -6.64 | yes |
| L-bracket Pcap -11.2% | t(5.66) = 4.86, p = 0.00329, g = 2.78 | yes |
| Mass loss cantilever -1.8% | t(7.87) = 4.97, p = 0.00114, g = 2.84 | yes |
| Mass loss L-bracket -3.2% | t(4.47) = 2.89, p = 0.0391, g = 1.65 | yes |
| Mass loss MBB -4.1% | t(5.32) = 12.82, p = 3.36e-5, g = 7.33 | yes |

Direction of K change correctly predicted by Abaqus only for MBB (M p.17).

### Table 4 (M p.19): masses (g), n = 5
| | Cantilever S / H | L-bracket S / H | MBB S / H |
| --- | --- | --- | --- |
| Mean | 60.9 / 59.9 | 43.1 / 41.7 | 87.3 / 83.7 |
| SD | 0.32 / 0.36 | 0.25 / 1.04 | 0.24 / 0.58 |
| Mean % change | -1.8 | -3.2 | -4.1 |

### Table 5 (M p.20): wall thickness, six microscope readings per section
| Artefact | Target (mm) | Mean (mm) | SD | Bias (mm) |
| --- | --- | --- | --- | --- |
| Hollow cantilever | 1.756 | 1.555 | 0.039 | -0.201 |
| Hollow L-bracket | 1.700 | 1.486 | 0.046 | -0.214 |
| Hollow MBB | 2.100 | 1.893 | 0.045 | -0.207 |

Relative deficit (derived): 11.4%, 12.6%, 9.9%.

## 4. Study 2 (tubes), set-up (M pp.13-15)
| Item | Value | Source |
| --- | --- | --- |
| Sizes OD / ID (mm) | A 15.0/0; B 15.8/5.0; C 18.0/10.0; D 21.2/15.0; E 25.0/20.0 (wall 7.5 solid, 5.4, 4.0, 3.1, 2.5) | M p.14 Table 1; J |
| Area | 176.3 +- 0.5 mm2; tol +-0.1 mm | M p.14 |
| Span / length | 288 mm (= 16 x 18 mm, ISO 178, size C); length 345.6 mm (10% overhang each end) | M p.14 |
| Rollers | supports radius 10 mm, load pin radius 25 mm; 0.1 mm/s | M p.15 |
| FE | Abaqus C3D8I, >= 4 elements through wall; rigid rollers, frictionless contact; support ref nodes fixed 6 DOF, load roller free in Z only; stabilisation nodes pinned X,Y; load 0.1 kN | M pp.14-15 |
| Analytical | Euler-Bernoulli 48EI/L^3 (J k_EB matches Table 6 to 4 digits); Timoshenko not used | M p.15; J |

### Table 6 (M p.22): predictions
| Mat | Size | E-B K | E-B % | Abaqus K | Abaqus % |
| --- | --- | --- | --- | --- | --- |
| Al 6063 | A | 0.3440 | | 0.3388 | |
| | B | 0.4193 | +21.9 | 0.4099 | +21.0 |
| | C | 0.6455 | +87.7 | 0.6233 | +84.0 |
| | D | 1.0287 | +199.0 | 0.9676 | +185.6 |
| | E | 1.5673 | +355.6 | 1.3884 | +309.7 |
| PA11 | A | 0.0088 | | 0.0086 | |
| | B | 0.0107 | +21.6 | 0.0104 | +20.8 |
| | C | 0.0165 | +87.5 | 0.0158 | +83.4 |
| | D | 0.0263 | +198.9 | 0.0245 | +184.6 |
| | E | 0.0400 | +354.6 | 0.0351 | +308.4 |

E-B vs Abaqus discrepancy A to E: Al 1.5, 2.3, 3.4, 5.9, 11.4%; PA11 2.3, 2.9, 4.4, 7.0, 12.2% (M p.21).

### Table 7 (M p.23): measured, n = 5
| Mat | Size | Mean K | K % | Mean Pcap | Pcap % |
| --- | --- | --- | --- | --- | --- |
| Al 6063 | A | 0.3003 | | 2.4667 | |
| | B | 0.3678 | +22.5 | 2.6252 | +6.4 |
| | C | 0.5358 | +78.5 | 2.7167 | +10.1 |
| | D | 0.7694 | +156.2 | 3.4437 | +39.6 |
| | E | 1.0549 | +251.3 | 4.1262 | +67.3 |
| PA11 | A | 0.0054 | | 0.1684 | |
| | B | 0.0065 | +19.1 | 0.1901 | +12.9 |
| | C | 0.0100 | +84.1 | 0.2576 | +53.0 |
| | D | 0.0182 | +234.4 | 0.3905 | +131.9 |
| | E | 0.0257 | +372.2 | 0.4160 | +147.0 |

Overprediction of measured K (M pp.22-23): Al E-B A..E +14.6, +14.0, +20.5, +33.7, +48.6%; Al Abaqus +12.9, +11.4, +16.3, +25.8, +31.6%. PA11 E-B +61.9 to +65.4% (A-C), +44.7% (D), +55.9% (E); PA11 Abaqus +57.6 to +60.5% (A-C), +34.7% (D), +36.8% (E). No wall-thickness bias found in tubes (M p.23). Derived: PA11 Abaqus/measured = 0.0086/0.0054 = 1.59 at A, so measured effective E is about 1.1 GPa (derived, unverified; M p.29 says modulus or fixture compliance).

Derived bending stress at peak, sigma = F L c / (4 I), PA11: A 36.6, B 35.7, C 35.8, D 40.1, E 33.1 MPa (my calc from Table 7; A reproduces V's sigma_f = 36.6 MPa).

## 5. Findings (M)
| ID | Statement | Page |
| --- | --- | --- |
| F1 | Stiffness changes from hollowing not significant (+3.5, -7.0, -14.9%) | p.16 |
| F2 | Pcap significantly up for cantilever (+61.1%) and MBB (+14.8%); MBB decouples K and Pcap | p.17 |
| F3 | L-bracket Pcap significantly down (-11.2%) | p.17 |
| F4 | Tubes: K and Pcap rise monotonically with I (PA11 +372.2% K, +147.0% Pcap; Al +251.3%, +67.3%) | p.21 |
| F5 | Stiffness inconsistency from local geometry at joints and cross-section transitions (shear near neutral axis lost) | p.26 |
| F6 | Tube Pcap gain lags K gain (local shell buckling); inverse of joint-dominated Study 1 decoupling | p.28 |
| F7 | Models over-predict tube K, worse for thin walls; point at modulus (E = 1.76 GPa too high) or fixture compliance | p.29 |

6.1.2 (p.27): cantilever/MBB gains attributed to compression struts under global Euler buckling; L-bracket loss to local shell buckling at short multi-axial joints, worsened by thinner walls ("up to 13.6% reduced", see discrepancy 7). Neither study modelled plasticity or buckling load capacity (p.30).

### Section 6.2 recommendations (M pp.30-32; no F-numbers)
1. Selective hollowing: struts yes, joints and transitions no; discriminate by local slenderness or local shear to bending strain-energy ratio (refers to F1, F5, F2, F3).
2. Optimiser-nucleated hollow voids at topology stage (level-set cannot nucleate holes).
3. Correct or characterise thin-wall FE error (e.g. wall-thickness dependent modulus correction) before FE-in-the-loop hollowing (F7).
4. Use manufactured wall thicknesses, not nominal (F3).
5. Restricting hollowing to a subset also reduces fine-mesh cost.

## 6. vshell doc numbers (V)
| Quantity | Value | V location |
| --- | --- | --- |
| CalculiX 2.21 C3D8I vs voxel FE, 108,906-cell hollow L-bracket | compliance 452.0332 vs 452.0331 N.mm; rel. L2 disp diff 4.2e-5 | Verification table |
| Tube K, vshell h_FE = 0.8 (A..E) | 0.0086, 0.0105, 0.0160, 0.0245, 0.0367 (J: 0.008632, 0.010454, 0.015963, 0.024501, 0.036680) | same |
| Abaqus (J ABAQUS dict = Table 6 numerical) | 0.0086, 0.0104, 0.0158, 0.0245, 0.0351 | same |
| Ratios to rod A: vshell / Abaqus / E-B | B 1.211/1.209/1.219; C 1.849/1.837/1.876; D 2.838/2.849/2.990; E 4.249/4.081/4.556 (J) | same |
| vshell vs Abaqus, absolute (derived) | A +0.4%, B +0.5%, C +1.0%, D 0.0%, E +4.5% | derived from J |
| vshell vs Abaqus, ratio (derived) | B +0.1%, C +0.7%, D -0.4%, E +4.1% | derived from J |
| Tube FE set-up | line-contact strips 1 cell wide, F = 100 N, E 1760 MPa, nu 0.42, span 288, length 345.6; absolute K differs from roller-contact Abaqus, ratios are the comparison | J docstring |
| Voxel cells / DOF / time (h = 0.8) | 133k-152k cells, 0.45-0.60 M DOF, 32-52 s | J |
| Thin-wall noise, 1.7 mm wall, h = 0.8/0.6/0.4/0.3 | within -0.7%, +5.2%, -6.2% of the 0.3 mm result; about +-6% grid noise on walls 1.5-1.8 mm; changes < ~5% read as no change | V Verification table, line 156 |
| sigma_f | 36.6 MPa, from solid PA11 rod A, 168.4 N at 288 mm span, sigma = F L c/4I; datasheet 52 MPa over-predicts tubes by 30-57% (my check: A +42, B +46, C +45, D +30, E +57%); with 36.6: B +2.5, C +2.2, D -8.7, E +10.7% (V: "B, C, D within 2-9%, E +11%") | V Strength |
| Quantile and as-built | q = 0.995 von Mises; stress scaled s/(s - delta), delta = 0.2 mm; K_eff = 0.7 | V |
| vshell walls (k = 0) | 1.83, 1.81, 2.20 mm vs Fusion 1.756, 1.700, 2.100 | V |

### vshell vs Study 1 (V, "Validation against Study 1")
| Part | K change vshell / Abaqus / measured | Pcap change vshell / measured | Pcap sign | K sign (derived, not in V) |
| --- | --- | --- | --- | --- |
| Cantilever | +1.2 / -11.6 / +3.5 (n.s.) | +148 / +61.1 | agrees (2.4x) | vshell agrees, Abaqus disagrees |
| L-bracket | +6.9 / +4.0 / -7.0 (n.s.) | -5.6 / -11.2 | agrees | vshell and Abaqus both disagree |
| MBB | -6.0 / -8.2 / -14.9 (n.s.) | -11.1 / +14.8 | disagrees | agrees (all n.s.) |

| Part | Solid K vshell / measured / Abaqus bounds (N/mm) | Solid Pcap vshell / measured; hollow |
| --- | --- | --- |
| Cantilever | 257 / 193 / 163-268 | 431 / 513 N (-16%) |
| L-bracket | 48 / 31 / 25-38 | 540 / 453 N (+19%) |
| MBB | 473 / 602 / 364-899 | 2742 / 2436 N (+13%) |

vshell solid K for L-bracket (48) lies outside Abaqus bounds (25-38) and the measured (31).
Decision-robustness (V): cantilever hollow in 25/25 of K_eff 0.5-1.0 x sigma_f +-20%; MBB solid 25/25; L-bracket flips hollow only at K_eff = 1.0 and sigma_f >= 36.6.

## 7. Discrepancies and flags
1. **V2-b tolerance (T s11).** T asks tube K within 4% of Abaqus at h_e 0.8 and 0.4. J at h = 0.8 gives size E +4.5% absolute (+4.1% ratio), so V2-b fails at the strict reading for E; V's "within 4%" refers to ratios and rounds E down. 0.4 mm is not yet in J (only h = 0.8 stored). Also J uses line contacts (stiffer-load path) vs paper roller contact (radius 10/25 mm); absolute agreement is partly fortuitous (paper's own E-B vs Abaqus gap is 12.2% at E).
2. **V2-c (T) vs V.** T requires 0.4 mm within 3% of 0.2 mm, monotone. V reports -6.2% at 0.4 and +5.2% at 0.6 against 0.3 mm (non-monotone, +-6% noise). The T criterion is tighter than the measured vshell noise; expect failure for a thin wall at this resolution.
3. **Manuscript K vs measured for tubes.** Study 2 Abaqus over-predicts measured PA11 K by +35 to +60% (E-B +45 to +65%). Matching Abaqus in V2-b validates the FE, not the physical stiffness. Ratio K(E)/K(A): Abaqus 4.08, measured 4.72 (Table 7, +372%), so even ratios differ from experiment by about 14%; the vshell doc statement "paper's Abaqus ratios within 4%" is correct but silent on this.
4. **Capacity sign table (T s11.1, V).** T s1.1/s11.1 and V agree with M Table 3 on all measured values. T s11.1 expects cantilever J >= +0.48 "vshell K/K0 0.992, P/P0 2.63", consistent with V (+1.2% and +148% differ: K 1.012 vs 0.992, P 2.48 vs 2.63). Mismatch between T s11.1 numbers (K/K0 0.992, P/P0 2.63) and V's table (+1.2%, +148%), unless T quotes the later best design (k > 0) not k = 0. Unverified which; needs a check against vshell outputs.
5. **T s1.1 vs M.** All wall and change values (Table 3, 5) and E, nu, rho, 200 N match. T lists Study 1 load as 200 N "in the FE model"; M also loaded GD at 0.2 kN and tests to failure. T s1.1 omits the load capacity significance for MBB (p = 3.3e-6, significant) and the stiffness bounds; fine. T cites only F3, F5, F7; F2 (MBB decoupling) is relevant to the MBB expectation in s11.1.
6. **F-number ordering.** F1-F4 are in Sections 4-5, F5-F7 in 6.1; M Section 6.2 has none. T's "Section 6.2 recommends selective hollowing driven by a mechanical measure" matches p.30.
7. **Manuscript internal:** "up to 13.6% reduced" (p.27) does not match Table 5 (derived 9.9-12.6% of target; 14.4% of measured L-bracket wall 0.214/1.486). Table 5 gives 0.2 mm; V uses delta = 0.2 mm, consistent.
8. **Mass targets vs as built:** GD targets 60/45/45 g, measured solid 60.9/43.1/87.3 g; the MBB 45 g target applies to the half model before mirroring (M p.7-8). Mass of hollow is 1.8-4.1% below solid (escape holes), so a constant-mass FE vs as-built comparison is biased low by that amount (T s3.9 handles as-built wall, not hole mass loss; unverified).
9. **Study 1 BC detail not in manuscript:** which DOF were constrained, exact load directions and bolt-hole counts are not stated in text (Fig 3 only). cshell BC adapter must take these from the STEP/Fusion setup, not from the paper.
10. **Tube measured load capacity is not FE-comparable:** no capacity model in M for any study (p.30); only V's sigma_f calibration. sigma_f calibration uses the rod A, so rod A prediction is exact by construction (0.0% in my check).
