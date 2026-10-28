"""
Minimal stand-in for the `interfaces.py` module of the original PCLS code base.

`analysis.py` imports only `BucklingResults` from it. The original module (with `AnalysisInterface`,
`BucklingAnalysisInterface`, ...) was not supplied; this stub reproduces the one container that
`analysis.py` constructs, with the fields it passes (collected from every `BucklingResults(...)` call).

Fields (mu = 1/lambda, the inverse buckling load factor):
    mu_values             (m,)   retained inverse eigenvalues, largest (most critical) first
    mode_shapes           (Ndofs, m) mode shapes, full DOF vectors (fixed DOFs zero)
    mu_tilde              float  p-norm aggregate (sum mu_i^p)^(1/p) >= max mu_i
    sensitivity           array  d(mu_tilde)/d(field) on the voxel grid passed to run_analysis
    mac_matrix            (m, m) or None  modal assurance criterion against prev_modes
    solver_iters          int
    n_candidates          int
    eigensolver_converged bool
    subspace_continuity   float or None
    order                 array or None  mode ordering against prev_modes
"""
from __future__ import annotations

from dataclasses import dataclass
from typing import Any, Optional

import numpy as np


@dataclass
class BucklingResults:
    mu_values: np.ndarray
    mode_shapes: np.ndarray
    mu_tilde: float
    sensitivity: np.ndarray
    mac_matrix: Optional[np.ndarray] = None
    solver_iters: int = 0
    n_candidates: int = 0
    eigensolver_converged: bool = True
    subspace_continuity: Optional[float] = None
    order: Optional[Any] = None

    @property
    def lambda_values(self) -> np.ndarray:
        """Buckling load factors 1/mu for the positive mu (compressive modes)."""
        mu = np.asarray(self.mu_values, float)
        return np.where(mu > 0, 1.0 / np.where(mu > 0, mu, 1.0), np.inf)
