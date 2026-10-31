# Tube E gap (job 7800269; tubes_pw.py, tubeE.json). k in kN/mm, quarter model, E = 0.0351 Abaqus, EB = 0.04004
h sweep (4 mm patch): 0.8 -> 0.03697; 0.4 -> 0.03673; 0.2 -> 0.03677 (6.9 M DOF). Voxel limit ~0.0368 = +4.8 % vs Abaqus.
Patch width at h=0.4: 2/4/6/10 mm -> 0.03658/0.03673/0.03680/0.03695 (0.2 mm: 2 mm 0.03668, 4 mm 0.03677).
V2-c (1.7 mm wall): 0.8 0.02669; 0.4 0.02627; 0.2 0.02626 (5.2 M DOF, fp32-preconditioned fp64 CG, 17 s).
Not done: ccx C3D10 cross-check.
