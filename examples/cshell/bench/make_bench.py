"""Appendix B benchmark parts: solid rods r0 = 5 mm along +x (B1 stub L/D = 0.5, B2 slender L/D = 8, B3 column
L = 100 axial). Clamped at x = 0 (support face), load on the x = L face. Writes STEP + cshell configs.
B1b (V5-B1b): slender rod (r0 = 5, L = 80, L/D = 8) clamped at x = 0, joined at x = 80 by a fillet (r = 3 mm) to a
stub (R = 10, L = 10, L/D = 0.5) that carries the tip load on its end face (B2 load and BCs).
Usage: python make_bench.py [names...] (default: B1, B2, B3; B1b only when named); run it via sbatch."""
import json
import os
import sys
from build123d import Axis, Cylinder, GeomType, export_step, fillet

os.chdir(os.path.dirname(os.path.abspath(__file__)))
R0 = 5.0
CASES = {"B1_stub": (5.0, [0, 0, -200.0]), "B2_slender": (80.0, [0, 0, -20.0]), "B3_column": (100.0, [-50.0, 0, 0])}


def b1b(root=False):
    """root=False: stub at the load end (B1b_stub_slender); root=True: stub at the clamped end (B1b_root_stub)."""
    L_s, R_t, L_t, r_f = 80.0, 10.0, 10.0, 3.0
    x_s, x_t, x_j = (L_t, 0.0, L_t) if root else (0.0, L_s, L_s)
    slender = Cylinder(R0, L_s).rotate(Axis.Y, 90).translate((x_s + L_s / 2, 0, 0))
    stub = Cylinder(R_t, L_t).rotate(Axis.Y, 90).translate((x_t + L_t / 2, 0, 0))
    body = slender + stub
    edge = [e for e in body.edges().filter_by(GeomType.CIRCLE)
            if abs(e.center().X - x_j) < 1e-6 and abs(e.radius - R0) < 1e-6]
    assert len(edge) == 1, len(edge)
    return fillet(edge, r_f), L_s + L_t


def write(name, shape, L, F):
    export_step(shape, f"{name}.step")
    cfg = {"name": name, "step_file": f"{name}.step", "output_dir": "../../../runs_cshell/bench",
           "supports": [{"type": "plane", "normal": [-1, 0, 0], "extreme": True}],
           "loads": [{"faces": {"type": "plane", "normal": [1, 0, 0], "extreme": True}, "force_N": F}],
           "mounts": {"auto_analytic": False}}
    if name != "B3_column":
        cfg["objective"] = {"w_stiffness": 1.0, "w_capacity": 0.0}
    json.dump(cfg, open(f"{name}.json", "w"), indent=1)
    print(name, L, "volume", shape.volume)


if __name__ == "__main__":
    want = sys.argv[1:] or list(CASES)
    for name, (L, F) in CASES.items():
        if name in want or name.split("_")[0] in want:
            write(name, Cylinder(R0, L).rotate(Axis.Y, 90).translate((L / 2, 0, 0)), L, F)
    if "B1b" in want:
        shape, L = b1b()
        write("B1b_stub_slender", shape, L, [0, 0, -20.0])
    if "B1b_root" in want:
        shape, L = b1b(root=True)
        write("B1b_root_stub", shape, L, [0, 0, -20.0])
