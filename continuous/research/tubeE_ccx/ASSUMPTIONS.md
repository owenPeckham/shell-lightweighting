# ASSUMPTIONS: CalculiX reproduction of the Abaqus tube 3-point bend (Study 2)

Units: mm, N, MPa; k reported in kN/mm.

## Inputs (manuscript continuous/reference/main.pdf pp. 14-15, Table 1, Table 6)
- PA11: E = 1760 MPa, nu = 0.42, linear isotropic.
- Tubes (OD/ID): A 15/0, E 25/20 (also B-D defined in tube_ccx.py). Span 288, length 345.6 (10 % overhang each side).
- Rollers: support pin radius 10 mm, loading pin radius 25 mm (p. 15, Sec. 3.2.4; Abaqus "rollers of the appropriate dimensions").
- Load 0.1 kN; Abaqus k = F / (load-roller displacement).
- Abaqus (Table 6): A 0.0086, E 0.0351 kN/mm.

## Model (tube_ccx.py)
- Quarter model: symmetry ux = 0 on x = 0 (midspan) and uy = 0 on y = 0 (vertical mid-plane). Valid: geometry, load and
  frictionless BCs are symmetric; Abaqus's stabilisation nodes (pinned X, Y) are implied by the symmetry planes.
- Elements: C3D8I (as Abaqus), structured. Tubes: radial sweep (nr through wall, nt over the half circumference,
  graded 3:1 finer at top/bottom). Rod A: butterfly (square core c = R/2 + O-ring). Axial: uniform hc within +-6 mm of
  load and support, linear growth (rate 0.12) to hmax elsewhere. L1: nr 4, nt 64, hc 0.5, hmax 3; L2: nr 8, nt 128, hc 0.25, hmax 1.5.
- Rollers: one layer of C3D8 on a cylinder sector, every node prescribed (= rigid body; support fixed in all DOF, load
  roller ux = uy = 0, uz = -d). Faceted master surface (facet 0.25 / 0.15 mm, sagitta < 1e-3 mm).
- Contact: CalculiX surface-to-surface, frictionless, linear penalty K = 1e5 N/mm^3 (sensitivity run at 1e6).
  Abaqus default is "hard" contact (Lagrange/penalty); the penalty overclosure is reported and must be << delta.
- Displacement control of the load roller (d = 3.0 mm E, 12.5 mm A); F from the roller reaction; k = F / d at
  time points 0.01 ... 1.0, interpolated to F = 100 N. Abaqus used force control: same F-delta curve.
- Tube rigid z mode before contact closes: SPRINGA 1e-3 N/mm to ground at midspan bottom (force < 3e-3 N vs 100 N).
- Small strain (no NLGEOM) by default, NLGEOM run as a check. Contact is nonlinear in both.
- "Patch" BCs = cshell voxel BCs (continuous/research/tubeE/tubes_pw.py): uniform vertical pressure on outer faces with
  x <= 2 and z >= zmax - 0.5 (4 mm wide); supports = nodes on plane x = 144 with z <= zmin + 0.5, uz = 0;
  k = F^2 / (2 W) (W = strain energy), exactly the voxel definition.

## Limitations
- Rigid rollers are discretised (faceted) not analytical; ccx penalty, not Abaqus hard contact.
- Abaqus mesh (unknown beyond ">= 4 through wall, 1 % von Mises convergence") not reproduced element-for-element.
