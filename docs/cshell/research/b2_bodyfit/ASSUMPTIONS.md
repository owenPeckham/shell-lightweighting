# B2 body-fitted diagnosis: inputs and assumptions

- Units: mm, N, MPa. E = 1760 MPa, nu = 0.42 (cshell.validate defaults, as in the decks).
- Run: runs_cshell/C_bench/B2_slender (best_candidate.stl = exported design, 5.78 M faces; escape_holes.csv;
  validation/{solid,design}.inp = the decks collect ran; bc_faces.npz = original CAD tessellation and face masks).
- BCs: as collect: clamp all nodes within 0.75 x 0.4 mm of the x = 0 face; 20 N in -z on the x = 80 face as an
  area-weighted (consistent P2) traction; buckling = linear eigen-buckling about the static stress state, 4-point rule.
- GPU solver (bf_gpu.py): straight-edged quadratic tets in C3D10 node order, 4-point rule for K, stress and K_sigma,
  cuDSS Cholesky (fp64), block Krylov on K^-1 G with K-orthonormal Rayleigh-Ritz. Verified against ccx 2.23 on the
  identical decks (see diagnosis.md).
- Corrected meshes (mesh2.py): isotropic surface remesh (pymeshlab) to edge h, gmsh Delaunay + Netgen optimisation
  inside that surface; h = 0.7, 0.5, 0.35 mm for design and solid.
- Not done: no re-export with the export.py patch (needs the optimiser context; the patch is untested code); the films
  at the escape holes are below the remesh resolution and are smoothed away in the corrected meshes instead.
- Voxel values are taken from validation.csv (opt grid, h_e = 0.4 mm); the voxel model was not re-run here.
