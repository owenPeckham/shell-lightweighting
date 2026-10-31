"""
Fast unit tests (about 1 minute):   python -m pytest tests/dshell/test_basic.py -q    (or run this file directly)
"""
import os
import sys

import numpy as np

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", ".."))
from dshell import fe  # noqa: E402

HERE = os.path.dirname(__file__)


def test_element_rigid_modes_and_symmetry():
    K = fe.unit_element_stiffness(0.3)
    assert np.allclose(K, K.T, atol=1e-12)
    ev = np.linalg.eigvalsh(K)
    assert np.sum(np.abs(ev) < 1e-10) == 6, "exactly six rigid-body modes"
    assert ev.min() > -1e-10


def test_patch_uniform_strain():
    """A linear displacement field gives zero nodal force at interior nodes (patch test)."""
    vf = np.ones((4, 4, 4))
    m = fe.VoxelMesh(vf, (0, 0, 0), 1.0)
    K = fe.assemble(m, 1000.0, 0.3)
    eps = np.array([[1e-3, 2e-4, 0], [2e-4, -5e-4, 1e-4], [0, 1e-4, 3e-4]])
    u = (m.X @ eps.T).ravel()
    f = (K @ u).reshape(-1, 3)
    interior = np.all((m.node_ijk > 0) & (m.node_ijk < 4), axis=1)
    assert np.abs(f[interior]).max() < 1e-9 * np.abs(f).max() + 1e-12
    # centre strain recovers the field exactly
    e, s, _ = fe.cell_stresses(m, u.reshape(-1, 3), 1000.0, 0.3)
    assert np.allclose(e[:, 0], 1e-3) and np.allclose(e[:, 3], 4e-4)


def test_cantilever_beam_against_timoshenko():
    """Square 4 x 4 mm, 60 mm long cantilever, tip load: within 4% of Timoshenko beam theory."""
    h, b, L, E, nu, F = 0.5, 4.0, 60.0, 1000.0, 0.3, 1.0
    vf = np.ones((int(L / h), int(b / h), int(b / h)))
    m = fe.VoxelMesh(vf, (0, 0, 0), h)
    fixed = np.where(np.isclose(m.X[:, 0], 0))[0]
    tip = np.where(np.isclose(m.X[:, 0], L))[0]
    sol = fe.solve(m, E, nu, fixed, tip, np.array([0, 0, -F]), tol=1e-10)
    I, A, G = b ** 4 / 12, b * b, E / (2 * (1 + nu))
    delta = F * L ** 3 / (3 * E * I) + F * L / (5 / 6 * G * A)
    assert abs(sol["compliance"] / F - delta) / delta < 0.04


def test_face_selectors_on_example():
    p = os.path.join(HERE, "..", "..", "examples", "parts", "Solid L-bracket.step")
    if not os.path.isfile(p):
        return
    from dshell import geometry
    tab = geometry.face_table(geometry.load_step(p))
    assert geometry.select_faces(tab, {"type": "cylinder", "radius": 4.26, "tol": 0.05}) == [9, 16, 24]
    pad = geometry.select_faces(tab, {"type": "plane", "normal": [0, 0, 1], "box": {"x": [-15, 20]}, "extreme": True})
    assert pad == [37]


def test_config_defaults():
    from dshell import config
    cfg = config.load({"step_file": "x.step", "supports": [{"type": "ids", "ids": [1]}],
                       "loads": [{"faces": {"type": "ids", "ids": [2]}, "force_N": [0, 0, -1]}]})
    assert cfg["material"]["E_MPa"] == 1760.0 and cfg["grid"]["fe_factor"] == 2 and cfg["name"] == "x"



def test_euler_section_inertia():
    """Strut section inertia from the voxels: solid round bar r = 4 mm -> I = pi r^4 / 4 within 3%."""
    from dshell import mechanics, config
    from dshell.segment import Segmentation
    h, r, L = 0.4, 4.0, 40.0
    n = int(L / h)
    m_ = int(np.ceil(2 * (r + h) / h)) + 1
    yc = (np.arange(m_) - (m_ - 1) / 2) * h
    sub = 4
    o = (np.arange(sub) + 0.5) / sub - 0.5
    Y = (yc[:, None] + h * o[None, :]).ravel()
    inside = (np.sqrt(Y[:, None] ** 2 + Y[None, :] ** 2) <= r).reshape(m_, sub, m_, sub).mean((1, 3))
    vf = np.repeat(inside[None], n, 0)
    mesh = fe.VoxelMesh(vf, (0.0, yc[0] - h / 2, yc[0] - h / 2), h)
    xs = np.arange(0.5, L, 1.0)
    pts = np.c_[xs, np.zeros_like(xs), np.zeros_like(xs)]
    seg = Segmentation(None, None, [dict(id=1, kind="strut", r_mm=r, L_mm=L, lam=L / (r / 2))], pts,
                       np.tile([1.0, 0, 0], (len(xs), 1)), np.ones(len(xs), int))
    cfg = config.load({"step_file": "x.step", "supports": [{"type": "ids", "ids": [0]}],
                       "loads": [{"faces": {"type": "ids", "ids": [0]}, "force_N": [-1, 0, 0]}]})
    sol = dict(u=np.zeros((mesh.n_nodes, 3)), compliance=1.0)
    met = mechanics.evaluate(mesh, sol, np.ones(mesh.n_cells, int), seg, cfg, np.array([-1.0, 0, 0]),
                             np.zeros((1, 3)) - 100, np.zeros((1, 3)) - 100, dict(hollow=[], s=0.0))
    I = met["per_region"][0]["I_min"]
    assert abs(I / (np.pi * r ** 4 / 4) - 1) < 0.03, I


if __name__ == "__main__":
    for name, fn in list(globals().items()):
        if name.startswith("test_"):
            fn()
            print("ok ", name)
