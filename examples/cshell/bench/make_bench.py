"""Appendix B benchmark parts: solid rods r0 = 5 mm along +x (B1 stub L/D = 0.5, B2 slender L/D = 8, B3 column
L = 100 axial). Clamped at x = 0 (support face), load on the x = L face. Writes STEP + cshell configs."""
import json
from build123d import Axis, Cylinder, export_step

R0 = 5.0
CASES = {"B1_stub": (5.0, [0, 0, -200.0]), "B2_slender": (80.0, [0, 0, -20.0]), "B3_column": (100.0, [-50.0, 0, 0])}
for name, (L, F) in CASES.items():
    rod = Cylinder(R0, L).rotate(Axis.Y, 90).translate((L / 2, 0, 0))
    export_step(rod, f"{name}.step")
    cfg = {"name": name, "step_file": f"{name}.step", "output_dir": "../../../runs_cshell/bench",
           "supports": [{"type": "plane", "normal": [-1, 0, 0], "extreme": True}],
           "loads": [{"faces": {"type": "plane", "normal": [1, 0, 0], "extreme": True}, "force_N": F}],
           "mounts": {"auto_analytic": False}}
    if name != "B3_column":
        cfg["objective"] = {"w_stiffness": 1.0, "w_capacity": 0.0}
    json.dump(cfg, open(f"{name}.json", "w"), indent=1)
    print(name, L)
