"""
Mesh-convergence check for walls as thin as the optimiser produces (1.5-1.8 mm): a 15 mm OD tube with a
1.7 mm wall in 3-point bending (span 120 mm), analysed at FE cell sizes h. Reports stiffness relative to
the finest grid.   python tests/validate_thin_wall.py 0.8 0.6 0.4
"""
import sys, os
import numpy as np
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", ".."))
import validate_tubes as vt

vt.SIZES = {"T": (15.0, 15.0 - 2 * 1.7)}
vt.ABAQUS = {"T": float("nan")}
vt.SPAN, vt.LENGTH = 120.0, 144.0
hs = [float(a) for a in sys.argv[1:]] or [0.8, 0.6, 0.4]
res = [vt.run("T", h) for h in hs]
ref = res[-1]["k_dshell"]
for r in res:
    print(f"h = {r['h']:.2f} mm ({r['wall'] / r['h']:.1f} cells through the wall): k = {r['k_dshell']:.5f} kN/mm, "
          f"{100 * (r['k_dshell'] / ref - 1):+.1f}% vs h = {hs[-1]}; Euler-Bernoulli {r['k_EB']:.5f}; cells {r['cells']}")
