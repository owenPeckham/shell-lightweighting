"""optimise helpers (PLAN_single_run §4, §7): exclusion zone r_ex = max(exclude_bc_mm, 2 r_MA), grid shift, weighted
quantile, config defaults (C1)."""
import numpy as np

from cshell.optimise import exclusion_mask, exclusion_radii, grid_shift, weighted_quantile


def test_exclusion_radii_rule():
    r_ma = np.array([1.0, 1.0, 1.2, 3.0, 3.0, 3.0])
    face_id = np.array([0, 0, 0, 1, 1, 1])
    r = exclusion_radii(r_ma, face_id, [0, 1, 2], exclude_bc_mm=2.0, factor=2.0)
    assert r == {0: 2.0, 1: 6.0, 2: 2.0}                     # max(2 mm, 2 × median r_MA); no vertices → 2 mm


def test_exclusion_mask_matches_brute_force():
    """An element is excluded iff one of its corner nodes lies within r of a source node of radius r."""
    ns, h = (21, 17, 13), 0.5
    src = [(np.array([np.ravel_multi_index((5, 8, 6), ns)]), 2.0),
           (np.array([np.ravel_multi_index((15, 8, 6), ns), np.ravel_multi_index((16, 8, 6), ns)]), 3.0)]
    got = exclusion_mask(ns, h, src)
    assert got.shape == tuple(n - 1 for n in ns)
    I = np.stack(np.meshgrid(*(np.arange(n - 1) for n in ns), indexing="ij"), -1)
    want = np.zeros(got.shape, bool)
    for nodes, r in src:
        for nd in nodes:
            p = np.array(np.unravel_index(nd, ns))
            dmin = np.full(got.shape, np.inf)
            for c in range(8):
                o = np.array([(c >> 2) & 1, (c >> 1) & 1, c & 1])
                dmin = np.minimum(dmin, h * np.linalg.norm(I + o - p, axis=-1))
            want |= dmin <= r + 1e-9
    assert np.array_equal(got, want)
    assert got.sum() > 0 and not got.all()
    # the larger radius reaches further: an element 2.5 mm (in y) from source 2 is excluded, from source 1 it is not
    assert got[15, 8 + 5, 6] and not got[5, 8 + 5, 6]


def test_grid_shift_tuple():
    from cshell import band
    V = np.array([[0.0, 0, 0], [10, 5, 3]])
    g0 = band.make_grid(V, 0.2, 4.0)
    with grid_shift((0.1, 0.0, 0.3)):
        g1 = band.make_grid(V, 0.2, 4.0)
    g2 = band.make_grid(V, 0.2, 4.0)
    assert np.allclose(np.asarray(g1.origin), np.asarray(g0.origin) - [0.1, 0.0, 0.3])
    assert all(a > b for a, b in zip(g1.shape, g0.shape)) and np.allclose(g2.origin, g0.origin)
    with grid_shift(None):
        assert np.allclose(band.make_grid(V, 0.2, 4.0).origin, g0.origin)


def test_weighted_quantile():
    x = np.array([4.0, 1.0, 3.0, 2.0])
    w = np.ones(4)
    assert weighted_quantile(x, w, 0.1) == 1.0 and weighted_quantile(x, w, 0.5) == 2.0
    assert weighted_quantile(x, w, 1.0) == 4.0
    assert weighted_quantile(x, np.array([0.0, 10.0, 0.0, 0.0]), 0.9) == 1.0


def test_config_defaults_c1():
    from cshell import config
    D = config.DEFAULTS
    O = D["objective"]                                     # PLAN_lightweight: lightweight default (weighted kept)
    assert O["type"] == "lightweight" and O["lw_margin_stiffness"] == O["lw_margin_capacity"] == 1.02
    assert O["constraints"] == [] and O["floors_as_constraints"]
    assert O["floor_stiffness"] == 0.85 and O["floor_capacity"] == 0.85 and O["w_stiffness"] == O["w_capacity"] == 0.5
    assert D["optimiser"]["starts"] == ["solid"]
    assert D["growth"]["tol"] == 1e-3 and D["growth"]["patience"] is None          # every quantile runs
    assert D["growth"]["step_tol"] == 1e-4 and D["growth"]["max_iters_per_step"] == 25
    assert D["growth"]["quantiles"] == [0.1, 0.2, 0.3, 0.4, 0.5, 0.6, 0.7, 0.8, 0.9, 1.0]
    assert D["capacity"]["exclude_bc_mm"] == 2.0
    assert D["export"] == {"h_mm": 0.2, "device": "gpu", "closed_thr": 0.9, "escape_holes": False}
    assert D["filter"]["passes"] == 2 and D["design"]["closure_xi"] == 0.9 and D["design"]["closure_band"] == 0.1


def test_void_min_eff():
    """§8 item 5: holes off → void_min_mm; holes on → max(void_min_mm, hole_d + clearance) (3 + 0.5 = 3.5 mm)."""
    import copy
    from cshell import config
    c = copy.deepcopy(config.DEFAULTS)
    assert config.void_min_eff(c) == 3.0
    c["export"]["escape_holes"] = True
    assert config.void_min_eff(c) == 3.5
    c["manufacturing"]["void_min_mm"] = 4.0
    assert config.void_min_eff(c) == 4.0


def test_settings_floor_constraints():
    """floors_as_constraints is read: weighted runs carry floor_K and floor_P as MMA constraints, compliance runs
    keep their own list."""
    from types import SimpleNamespace
    from cshell import config
    from cshell.optimise import PhaseModel
    import copy
    for typ, fl, want in (("weighted", True, ("floor_K", "floor_P")), ("weighted", False, ()),
                          ("compliance", True, ())):
        c = copy.deepcopy(config.DEFAULTS)
        c["objective"].update(type=typ, floors_as_constraints=fl)
        m = SimpleNamespace(ctx=SimpleNamespace(cfg=c), phase=0)
        st = PhaseModel.settings(m, 0.0)
        assert st.objective == typ and st.constraints == want, (typ, fl, st.constraints)
