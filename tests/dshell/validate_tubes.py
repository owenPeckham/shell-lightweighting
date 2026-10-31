"""
Thin-wall check of the voxel FE against Study 2 of the manuscript (Table 6, Nylon PA11):
3-point bending of constant-area tubes, span 288 mm, total length 345.6 mm, E = 1760 MPa, nu = 0.42.
Supports and load are line contacts (all nodes on a 1-cell-wide strip), so absolute stiffness differs a
little from the paper's roller-contact Abaqus model; the stiffness RATIOS to the solid rod are compared.

    python tests/validate_tubes.py [h_fe_mm ...]
"""
import sys, os, time, json
import numpy as np
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))
from vshell import fe

E, NU = 1760.0, 0.42
SIZES = {"A": (15.0, 0.0), "B": (15.8, 5.0), "C": (18.0, 10.0), "D": (21.2, 15.0), "E": (25.0, 20.0)}
ABAQUS = {"A": 0.0086, "B": 0.0104, "C": 0.0158, "D": 0.0245, "E": 0.0351}     # kN/mm, manuscript Table 6
SPAN, LENGTH, F = 288.0, 345.6, 100.0


def tube_fraction(od, idd, h, sub=4):
    """Cell volume fractions of a tube along x, by supersampling each cell's y-z section."""
    ro, ri = od / 2, idd / 2
    ny = int(np.ceil((ro + h) * 2 / h)) + 1
    nx = int(np.ceil(LENGTH / h))
    yc = (np.arange(ny) - (ny - 1) / 2) * h
    o = (np.arange(sub) + 0.5) / sub - 0.5
    Y = (yc[:, None] + h * o[None, :]).ravel()
    r = np.sqrt(Y[:, None] ** 2 + Y[None, :] ** 2)
    inside = ((r <= ro) & (r >= ri)).reshape(ny, sub, ny, sub).mean((1, 3))
    vf = np.repeat(inside[None, :, :], nx, axis=0)
    origin = np.array([-LENGTH / 2, yc[0] - h / 2, yc[0] - h / 2])
    return vf, origin


def run(size, h):
    od, idd = SIZES[size]
    vf, origin = tube_fraction(od, idd, h)
    mesh = fe.VoxelMesh(vf, origin, h, vf_min=0.05)
    X = mesh.X
    zmin, zmax = X[:, 2].min(), X[:, 2].max()
    strip = lambda x0, z0: np.where((np.abs(X[:, 0] - x0) <= h / 2 + 1e-9) & (np.abs(X[:, 2] - z0) <= 1e-9))[0]
    sl, sr, top = strip(-SPAN / 2, zmin), strip(SPAN / 2, zmin), strip(0.0, zmax)
    if len(top) == 0:   # odd cell count: use the nearest node line
        xs = X[np.argmin(np.abs(X[:, 0])), 0]
        top = np.where((np.abs(X[:, 0] - xs) < 1e-9) & (np.abs(X[:, 2] - zmax) < 1e-9))[0]
    dofs = list(3 * sl + 2) + list(3 * sr + 2)                     # rollers: uz = 0
    dofs += list(3 * sl + 0)                                        # one end pinned in x
    mid_y = lambda nodes: nodes[np.argmin(np.abs(X[nodes, 1]))]
    dofs += [3 * mid_y(sl) + 1, 3 * mid_y(sr) + 1]                  # stop rigid rotation/translation in y
    t0 = time.time()
    sol = fe.solve(mesh, E, NU, None, top, np.array([0, 0, -F]), fixed_dofs=np.array(dofs), tol=1e-9, maxiter=3000)
    k = F ** 2 / sol["compliance"] / 1000.0    # kN/mm
    A = np.pi * ((od / 2) ** 2 - (idd / 2) ** 2)
    I = np.pi / 64 * (od ** 4 - idd ** 4)
    k_eb = 48 * E * I / SPAN ** 3 / 1000.0
    return dict(size=size, h=h, wall=(od - idd) / 2, cells=mesh.n_cells, dof=sol["ndof"], k_vshell=k, k_EB=k_eb,
                k_abaqus=ABAQUS[size], time=time.time() - t0, area=A, area_vox=float(vf.sum() * h * h / vf.shape[0]))


if __name__ == "__main__":
    hs = [float(a) for a in sys.argv[1:]] or [0.8]
    rows = []
    for h in hs:
        for s in SIZES:
            r = run(s, h)
            rows.append(r)
            print(f"h={h} size {s}: wall {r['wall']:.1f} mm, cells {r['cells']}, k = {r['k_vshell']:.5f} kN/mm "
                  f"(Euler-Bernoulli {r['k_EB']:.5f}, Abaqus {r['k_abaqus']:.4f}), {r['time']:.0f} s", flush=True)
        base = [r for r in rows if r["h"] == h and r["size"] == "A"][0]
        for r in rows:
            if r["h"] == h:
                r["ratio_vshell"] = r["k_vshell"] / base["k_vshell"]
                r["ratio_abaqus"] = r["k_abaqus"] / ABAQUS["A"]
                r["ratio_EB"] = r["k_EB"] / base["k_EB"]
    for r in rows:
        print(f"h={r['h']} {r['size']}: ratio to rod A  vshell {r['ratio_vshell']:.3f} | Abaqus {r['ratio_abaqus']:.3f} | E-B {r['ratio_EB']:.3f}")
    with open(os.path.join(os.path.dirname(__file__), "validate_tubes_results.json"), "w") as f:
        json.dump(rows, f, indent=1)
