"""cshell run configuration: defaults, JSON loading and validation (THEORY §10.2).

Same `supports`, `loads`, `material`, `mounts` keys and face selectors as vshell (see vshell/config.py). Every key has a
default except `step_file`, `supports` and `loads`. Phase-indexed lists (`stress_p`, `softmin_beta`, `radius_mm`,
`max_iter`, `sliver_weight`, `grid.phases`) give one value per continuation phase; a two-element entry is a linear
ramp over that phase. Units: mm, N, MPa, g/cm³.
"""
from __future__ import annotations

import copy
import json
import os

from vshell.config import _merge

DEFAULTS = {
    "name": None,
    "step_file": None,
    "output_dir": "runs_cshell",
    "supports": [],                    # face selectors; all DOF clamped (R3, minimum-stiffness bolt-hole case)
    "loads": [],                       # [{"faces": selector, "force_N": [fx, fy, fz]}]; summed as one case (D3)
    "material": {"E_MPa": 1760.0, "nu": 0.42, "density_g_cm3": 1.04, "strength_MPa": 36.6},
    "manufacturing": {"t_min_mm": 1.5, "wall_bias_mm": 0.2, "outer_bias_mm": 0.0, "void_min_mm": 3.0,
                      "void_min_mm3": 30.0, "escape_holes_per_void": 2, "escape_hole_diameter_mm": 3.0},
    "mounts": {"auto_analytic": True, "extra_faces": [], "min_blob_fraction": 0.01, "dilate_mm": 1.5},
    "design": {"allow_outward": True, "outward_extent_initial_mm": 4.0, "gap_mm": 1.0, "keep_out": [],
               "keep_out_clearance_mm": 0.5, "surface_edge_mm": 0.75, "closure_gamma": 0.75,
               "inset_soft_mm": 0.1},                            # D1: width of the smoothed (t_min − d)⁺ bound
    "grid": {"phases": [{"h_e": 0.8, "h_g": 0.2}, {"h_e": 0.4, "h_g": 0.1}], "export_h_mm": 0.05,
             "rho_min": 1e-6, "tessellation_mm": 0.02, "active_margin_mm": 0.5,    # FE active-set dilation
             "geometry_device": "auto",     # band chain ψ → χ → ρ, V, pull-backs: "auto" (GPU if CUDA) | "cpu" | "gpu"
             "geometry_offload": True},     # GPU chain: release the band from the device during each FE evaluation
    "fe": {"engine": "auto",      # "auto" (direct fp64 if it fits and DOF <= auto_direct_max_dof, else twogrid)
                                  # | "direct" (fe_sparse, cuDSS) | "twogrid" (fe_twogrid, matrix-free two-grid PCG)
           "auto_direct_max_dof": 2_000_000,
           "element_type": "q1_im", "backend": "auto",
           "precision": "auto",                             # direct cuDSS factor: "auto" | "fp64" | "fp32" (fe_sparse)
           "twogrid": {"coarse_precision": "auto",           # fp32 coarse factor, fp64 if fp32 breaks down
                       "degree": 5, "smoothing_range": 30.0,      # Chebyshev(Jacobi) smoother on [λmax/r, λmax]
                       "tol": 1e-10, "krylov_tol": 1e-2,          # PCG: exact solves / buckling Krylov expansion
                       "maxiter": 100, "nrhs": 7}},               # PCG block width (work memory ∝ nrhs)
    "capacity": {"stress_p": [6, 8, 10], "stress_q": 2, "stress_quantile": 0.995, "exclude_bc_mm": 2.0,
                 "include_mounts": False, "buckling": True, "buckling_modes": 6, "buckling_p": 8,
                 "knockdown": 0.85, "softmin_beta": [8, 12, 16], "eig_tol": 1e-3, "min_density_geometric": 0.1},
    "objective": {"w_stiffness": 0.5, "w_capacity": 0.5, "min_gain": 0.02, "floor_stiffness": 0.85,
                  "floor_capacity": 0.85, "floors_as_constraints": False},
    "filter": {"radius_mm": [4.0, 3.0, 2.5]},
    "optimiser": {"method": "mma", "starts": ["uniform", "vshell", "neutral"], "max_iter": [60, 80, 40],
                  "move": 0.1, "tol_x": 0.005, "tol_f": 1e-4, "mass_tol": 1e-4,
                  "sliver_weight": [0.0, [0.0, 0.05], 0.1], "seed": 0},
    "validation": {"body_fitted": True, "second_resolution": True},
    "report": {"enabled": True, "dpi": 130},
}


def load(path_or_dict, overrides: dict | None = None) -> dict:
    """Merge a JSON file (or dict) and optional overrides onto DEFAULTS, resolve paths, validate."""
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
    for key in ("step_file", "output_dir"):
        if not os.path.isabs(cfg[key]):
            cfg[key] = os.path.normpath(os.path.join(base_dir, cfg[key]))
    if not cfg["name"]:
        cfg["name"] = os.path.splitext(os.path.basename(cfg["step_file"]))[0].replace(" ", "_")
    validate(cfg)
    return cfg


def parse_set(items) -> dict:
    """['objective.w_capacity=0.8', 'grid.phases=[...]'] → nested dict; values parsed as JSON where possible."""
    out: dict = {}
    for it in items or []:
        key, _, val = it.partition("=")
        try:
            v = json.loads(val)
        except json.JSONDecodeError:
            v = val
        d = out
        parts = key.split(".")
        for p in parts[:-1]:
            d = d.setdefault(p, {})
        d[parts[-1]] = v
    return out


def phase_value(entry, phase: int, frac: float = 0.0):
    """Value of a phase-indexed entry for continuation phase `phase` at fraction `frac` ∈ [0, 1] of that phase."""
    if not isinstance(entry, (list, tuple)):
        return entry
    v = entry[min(phase, len(entry) - 1)]
    if isinstance(v, (list, tuple)):
        return v[0] + (v[1] - v[0]) * frac
    return v


def validate(cfg: dict) -> None:
    if not cfg["supports"]:
        raise ValueError("config: at least one 'supports' face selector is required")
    if not cfg["loads"]:
        raise ValueError("config: at least one entry in 'loads' is required")
    for L in cfg["loads"]:
        if "faces" not in L or "force_N" not in L or len(L["force_N"]) != 3:
            raise ValueError("config: each load needs 'faces' (selector) and 'force_N' [fx, fy, fz]")
    o = cfg["objective"]
    if o["w_stiffness"] < 0 or o["w_capacity"] < 0 or (o["w_stiffness"] + o["w_capacity"]) == 0:
        raise ValueError("objective weights must be >= 0 and not both zero")
    g = cfg["design"]["closure_gamma"]
    if not 0.6 <= g <= 0.9:
        raise ValueError("design.closure_gamma must lie in [0.6, 0.9] (THEORY §3.8)")
    for ph in cfg["grid"]["phases"]:
        k = ph["h_e"] / ph["h_g"]
        if abs(k - round(k)) > 1e-9 or round(k) < 1:
            raise ValueError(f"grid phase {ph}: h_e must be an integer multiple of h_g")
    if cfg["optimiser"]["method"] not in ("mma", "lbfgsb"):
        raise ValueError("optimiser.method must be 'mma' or 'lbfgsb'")
    if cfg["fe"]["engine"] not in ("auto", "direct", "sparse", "twogrid"):
        raise ValueError("fe.engine must be 'auto', 'direct' ('sparse') or 'twogrid'")
    if cfg["fe"]["element_type"] != "q1_im":
        print("config: warning: element_type != q1_im is not a production setting (THEORY §4.1)")


def copy_cfg(cfg: dict) -> dict:
    return copy.deepcopy(cfg)
