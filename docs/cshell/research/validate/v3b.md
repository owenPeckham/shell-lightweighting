# V3-b: local shell buckling, tube R_m = 10, t = 1.5, L = 60 (THEORY §11)

Set-up: E = 1760 MPa, nu = 0.42. Clamped at one end (all DOF); axial compressive reference load P = 1000 N shared over
the far end face, which is free laterally. lambda = buckling load factor for P = 1000 N.
Body-fitted: `cshell/validate.py` (gmsh linear tets graded by wall thickness, quadratic C3D10, ccx 2.23 `*BUCKLE`, SPOOLES),
exact-ish tube STL (720 sections), "layers" = target linear tets through t (quadratic: twice as many intervals).
Voxel: `cshell.fe_sparse.buckling` (q1_im, partial-volume density from 6-8x sub-sampled section fractions, tol 1e-6).
Scripts: `v3b_run.py`, `v3b_ccx.sh`, `v3b_voxel.sh`; raw data `v3b_ccx_area.json`, `v3b_ccx_nodal.json`, `v3b_voxel.json`, `v3b_voxel_modes.json`.

## Body-fitted ccx, consistent (uniform-traction) end load (job 7796155), 72 CPUs

| layers | h (mm) | DOF | lambda_1,2 | lambda_3,4 | lambda_5,6 | C (N mm) | ccx time |
|---|---|---|---|---|---|---|---|
| 2 | 0.75 | 435k | 3.950 | 4.556 | 5.929 | 358.16 | 52 s |
| 3 | 0.50 | 1.20M | 3.944 | 4.556 | 5.911 | 358.20 | 230 s |
| 4 | 0.375 | 2.44M | 3.941 | 4.555 | 5.902 | 358.22 | 846 s |

Mesh-converged: lambda_1 changes by 0.2 % (2 -> 3 layers) and 0.08 % (3 -> 4); reference lambda_1 = 3.941 (3 layers: 0.07 % from the finest).
Layers 5 (est. 4.7M DOF) was OOM-killed at 160 GB (job 7795559) and was not repeated.

## Voxel (cshell sparse engine, GPU, job 7796117)

| h_e (mm) | DOF | lambda_1 | vs ccx 3.941 | lambda_3 | C (N mm) | time |
|---|---|---|---|---|---|---|
| 0.8 | 77k | 4.687 | +18.9 % | 5.110 | 416.0 | 6 s |
| 0.4 | 429k | 4.136 | +4.9 % | 4.631 | 374.9 | 20 s |
| 0.2 | 2.77M | 4.000 | +1.5 % | 4.582 | 362.9 | 160 s |

Observed order of the lambda_1 error: 0.746 -> 0.195 -> 0.059 (ratios 3.8, 3.3), about 1.9 (second order expected for q1_im).
Voxel lambda_1 and lambda_2 split (4.000, 4.089 at 0.2) by grid anisotropy; the ccx pair is degenerate. The second group
(4.58 vs 4.555) agrees to 0.6 %; compliance at 0.2 agrees to 1.3 % (the BCs and elasticity are consistent).

**Criterion "within 10 % at h_e = 0.2-0.4": PASS (+1.5 % at 0.2, +4.9 % at 0.4).** At h_e = 0.8 (1.9 elements through t) it is not met.

Mode character (voxel 0.4, job 7796752, modal strain energy in axial slabs of 10 mm, slab 0 = clamped end):
modes 1-2 carry 70 % of the energy in the loaded free-end slab (local free-edge shell mode); modes 3-4 are spread along the length.
The classical Lorenz value would be lambda = 15.8 (sigma_cl = E t / (R sqrt(3(1-nu^2))) = 168 MPa), so the governing mode here is
an end-localised one, below Lorenz; both codes see it.

## Finding: equal nodal load share is not valid on C3D10 (job 7795559)

With the spec's "load shared equally over the nodes on the load faces", applied to all C3D10 nodes (corner and mid-edge alike),
the end traction is not uniform and the first factor depended on the mesh by up to 2x with no trend:

| layers | 2 | 3 | 4 |
|---|---|---|---|
| lambda_1 (equal nodal share) | 1.511 | 2.412 | 1.857 |
| lambda_1 (consistent A/3 on mid-edge nodes) | 3.950 | 3.944 | 3.941 |

(compliance changed by only 4 %, 374 -> 361 vs 358.2). `validate_design` therefore defaults to `load_distribution="area"`;
`"nodal"` is kept as an option. The voxel engine keeps the equal share (on a regular grid that is a uniform traction).
