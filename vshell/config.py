"""
Run configuration: defaults, JSON loading and validation.

A run is described by one JSON file. Every key has a default except `step_file`, `supports` and
`loads`. Units: mm, N, MPa, g/cm^3.

Face selectors (used by `supports`, `loads[].faces` and `mounts.extra_faces`), any of:
    {"type": "ids", "ids": [9, 16, 24]}                      face indices from `python -m vshell faces`
    {"type": "cylinder", "radius": 4.26, "tol": 0.05}         every cylindrical face of that radius
    {"type": "plane", "normal": [0, 0, 1], "tol_deg": 2,      planar faces facing `normal`, optionally
     "box": {"x": [-15, 99]}, "extreme": true}                 inside a centroid box, and only those at the
                                                              extreme position along `normal`
"""
from __future__ import annotations

import copy
import json
import os

DEFAULTS = {
    "name": None,                      # run name (default: STEP file stem)
    "step_file": None,
    "output_dir": "runs",
    "material": {
        "E_MPa": 1760.0,               # PA11 (manuscript, Section 3.1.4)
        "nu": 0.42,
        "density_g_cm3": 1.04,
        "strength_MPa": 36.6,          # flexural failure stress of MJF PA11, back-calculated from the solid
                                       # 15 mm rod of Study 2 (168.4 N, span 288 mm); the datasheet's tensile
                                       # strength (52 MPa, ASTM D638) overpredicts the Study 2 tubes by ~40%
    },
    "manufacturing": {
        "t_min_mm": 1.5,               # minimum nominal wall (MJF structural range 1.5-3.0 mm)
        "wall_bias_mm": 0.2,           # as-built walls ~0.2 mm thinner than designed (manuscript Table 5)
        "void_min_mm3": 30.0,          # enclosed voids smaller than this are filled
        "escape_holes_per_void": 2,
        "escape_hole_diameter_mm": 3.0,
    },
    "supports": [],                    # face selectors, all DOF fixed (minimum-stiffness bolt-hole case)
    "loads": [],                       # [{"faces": selector, "force_N": [fx, fy, fz]}]
    "mounts": {
        "auto_analytic": True,         # analytic (plane/cylinder/...) faces mark preserve regions -> kept solid
        "extra_faces": [],             # selectors for further faces whose material must stay solid
        "min_blob_fraction": 0.01,     # analytic blobs smaller than this share of volume are ignored
        "dilate_mm": 1.5,
    },
    "grid": {
        "sdf_mm": 0.4,                 # signed-distance grid
        "fe_factor": 2,                # FE cell = fe_factor * sdf_mm
        "seg_mm": 1.0,                 # segmentation / skeleton grid
        "band_mm": 4.0,                # exact signed distance within this band of the surface
        "pad_mm": 4.0,
        "tessellation_mm": 0.02,       # B-rep tessellation chordal tolerance
    },
    "segmentation": {
        "joint_k": 1.6,                # joint ball radius = joint_k * local radius
    },
    "design": {
        "transition_mm": 3.0,          # solid-to-hollow taper length (Gaussian blend)
        "max_offset_mm": 4.0,          # largest outer offset searched; grid padding and exact band grow to fit
    },
    "capacity": {
        "stress_quantile": 0.995,      # volume-weighted quantile used instead of the (mesh-dependent) peak
        "exclude_bc_mm": 2.0,          # cells this close to fixed or loaded nodes are not used for strength
        "include_mounts": False,       # mounts are never changed; their stress is reported separately
        "K_eff": 0.7,                  # Euler effective-length factor (0.7 = fixed-pinned; GD struts end in
                                       # rigid joints or mounts; 1.0 = pinned-pinned, conservative)
        "local_buckling": True,
    },
    "objective": {
        "w_stiffness": 0.5,            # J = (w_k ln(K/K0) + w_p ln(P/P0)) / (w_k + w_p)
        "w_capacity": 0.5,
        "min_gain": 0.02,              # accept a hollow design only if J > ln(1 + min_gain) ...
        "floor_stiffness": 0.85,       # ... and K/K0 >= floor and P/P0 >= floor
        "floor_capacity": 0.85,
    },
    "search": {
        "max_fe_evaluations": 10,      # hollow designs analysed by FE (the solid baseline is extra)
        "rank_by": "R",                # order in which regions are made solid: 'R' (shear/bending) or 'lambda'
    },
    "fe": {
        "vf_min": 0.05,
        "tol": 1e-8,
        "maxiter": 800,
    },
    "report": {
        "enabled": True,
        "dpi": 130,
    },
    "verbose": True,
}


def _merge(base: dict, over: dict) -> dict:
    out = copy.deepcopy(base)
    for k, v in (over or {}).items():
        if isinstance(v, dict) and isinstance(out.get(k), dict):
            out[k] = _merge(out[k], v)
        else:
            out[k] = copy.deepcopy(v)
    return out


def load(path_or_dict, overrides: dict | None = None) -> dict:
    """Merge a JSON file (or dict) and optional overrides onto DEFAULTS, then validate."""
    if isinstance(path_or_dict, dict):
        user, base_dir = path_or_dict, os.getcwd()
    else:
        with open(path_or_dict) as f:
            user = json.load(f)
        base_dir = os.path.dirname(os.path.abspath(path_or_dict))
    cfg = _merge(DEFAULTS, user)
    if overrides:
        cfg = _merge(cfg, overrides)
    if not cfg["step_file"]:
        raise ValueError("config: 'step_file' is required")
    if not os.path.isabs(cfg["step_file"]):
        cfg["step_file"] = os.path.normpath(os.path.join(base_dir, cfg["step_file"]))
    if not os.path.isabs(cfg["output_dir"]):
        cfg["output_dir"] = os.path.normpath(os.path.join(base_dir, cfg["output_dir"]))
    if not cfg["name"]:
        cfg["name"] = os.path.splitext(os.path.basename(cfg["step_file"]))[0].replace(" ", "_")
    validate(cfg)
    return cfg


def validate(cfg: dict) -> None:
    if not cfg["supports"]:
        raise ValueError("config: at least one 'supports' face selector is required")
    if not cfg["loads"]:
        raise ValueError("config: at least one entry in 'loads' is required")
    for L in cfg["loads"]:
        if "faces" not in L or "force_N" not in L or len(L["force_N"]) != 3:
            raise ValueError("config: each load needs 'faces' (selector) and 'force_N' [fx, fy, fz]")
    g = cfg["grid"]
    if int(g["fe_factor"]) < 1:
        raise ValueError("grid.fe_factor must be >= 1")
    o = cfg["objective"]
    if o["w_stiffness"] < 0 or o["w_capacity"] < 0 or (o["w_stiffness"] + o["w_capacity"]) == 0:
        raise ValueError("objective weights must be >= 0 and not both zero")
    if cfg["search"]["rank_by"] not in ("R", "lambda"):
        raise ValueError("search.rank_by must be 'R' or 'lambda'")
