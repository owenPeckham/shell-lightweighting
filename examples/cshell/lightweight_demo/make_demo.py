"""Lightweighting demo part: stocky solid rectangular beam (b_y wide, h_z deep, length L along +x) clamped at x = 0,
tip load -z (bending about y) on the x = L face. Writes STEP + cshell config (lightweight default objective).
Usage: python make_demo.py b_y h_z L F   (run via sbatch: examples/cshell/lightweight_demo/run_demo.sh)."""
import json
import os
import sys
from build123d import Box, export_step

os.chdir(os.path.dirname(os.path.abspath(__file__)))
by, hz, L, F = (float(x) for x in sys.argv[1:5])
name = f"Beam_{int(by)}x{int(hz)}x{int(L)}"
shape = Box(L, by, hz).translate((L / 2, 0, 0))
export_step(shape, f"{name}.step")
cfg = {"name": name, "step_file": f"{name}.step",
       "output_dir": "/projects/b6bl/owenpeckham.b6bl/home_data/shell-lightweighting/runs_cshell/LW_DEMO",
       "supports": [{"type": "plane", "normal": [-1, 0, 0], "extreme": True}],
       "loads": [{"faces": {"type": "plane", "normal": [1, 0, 0], "extreme": True}, "force_N": [0, 0, -F]}],
       "mounts": {"auto_analytic": False}}
json.dump(cfg, open(f"{name}.json", "w"), indent=1)
print(name, "volume mm3", shape.volume)
