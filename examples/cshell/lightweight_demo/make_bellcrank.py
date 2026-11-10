"""Lightweighting demo, non-beam part: thick bell crank (L-shaped lever). Pivot boss (R 16, bore r 6) at the origin,
arm 1 along +x to a tip boss (R 11, bore r 4), arm 2 along +y to a tip boss (R 11, bore r 3.5); arms W x T section,
inner corner filleted. Pivot bore clamped; tip-1 bore loaded -y, tip-2 bore loaded +x (both bend their arm in plane,
summed as one case). Distinct bore radii let the cylinder selectors pick each bore. Writes STEP + cshell config.
Usage: python make_bellcrank.py W T L1 L2 F [R_pivot r_fillet]   (run via sbatch: run_bellcrank.sh)."""
import json
import os
import sys
from build123d import Axis, Box, Cylinder, export_step, fillet

os.chdir(os.path.dirname(os.path.abspath(__file__)))
W, T, L1, L2, F = (float(x) for x in sys.argv[1:6])
R_p = float(sys.argv[6]) if len(sys.argv) > 6 else 16.0      # pivot boss radius
r_fil = float(sys.argv[7]) if len(sys.argv) > 7 else 8.0     # inner-corner fillet radius
r_p, R_t, r_1, r_2 = 6.0, 11.0, 4.0, 3.5
name = f"BellCrank_{int(W)}x{int(T)}" + (f"_P{int(R_p)}F{int(r_fil)}" if (R_p, r_fil) != (16.0, 8.0) else "")

body = (Box(L1, W, T).translate((L1 / 2, 0, 0)) + Box(W, L2, T).translate((0, L2 / 2, 0))
        + Cylinder(R_p, T) + Cylinder(R_t, T).translate((L1, 0, 0)) + Cylinder(R_t, T).translate((0, L2, 0)))
# inner corner of the L: the z-parallel edge at (W/2, W/2)
corner = [e for e in body.edges().filter_by(Axis.Z)
          if abs(e.center().X - W / 2) < 1e-6 and abs(e.center().Y - W / 2) < 1e-6]
if len(corner) == 1:
    body = fillet(corner, r_fil)
else:
    print("inner-corner edge not found (", len(corner), "): no fillet")
body = (body - Cylinder(r_p, T + 2) - Cylinder(r_1, T + 2).translate((L1, 0, 0))
        - Cylinder(r_2, T + 2).translate((0, L2, 0)))
export_step(body, f"{name}.step")
cfg = {"name": name, "step_file": f"{name}.step",
       "output_dir": "/projects/b6bl/owenpeckham.b6bl/home_data/shell-lightweighting/runs_cshell/LW_DEMO",
       "supports": [{"type": "cylinder", "radius": r_p, "tol": 0.05}],
       "loads": [{"faces": {"type": "cylinder", "radius": r_1, "tol": 0.05}, "force_N": [0, -F, 0]},
                 {"faces": {"type": "cylinder", "radius": r_2, "tol": 0.05}, "force_N": [F, 0, 0]}],
       # a CAD solid has only analytic faces: auto_analytic would make every face a mount (nothing designable);
       # the bores (support/load faces) are the mounts
       "mounts": {"auto_analytic": False}}
json.dump(cfg, open(f"{name}.json", "w"), indent=1)
print(name, "volume mm3", body.volume, "solids", len(body.solids()))
