"""MMA wrapper, constant-mass restoration and a projected L-BFGS alternative (THEORY_continuous §7.1, §7.2, §7.5, §7.6).

`MMA` wraps `mmapy.mmasub` (Svanberg 2007 note; float64, CPU) with the §7.1 settings, serialisable state
(§7.6, bit-exact resume) and the §7.5 oscillation safeguard. `restore_mass` is the §7.2 one-dimensional
restoration  xi <- clip(xi+ + eta g/||g||_inf),  V(xi) = V0 solved for eta by Brent.
`ProjectedLBFGSB` has the same `step()` interface so optimisers can be compared (objective gradient only;
mass is held by `restore_mass` after each step).
"""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Callable, Dict, Optional

import numpy as np
from scipy.optimize import brentq


def _col(a, n=None) -> np.ndarray:
    a = np.asarray(a, dtype=np.float64)
    return a.reshape(-1, 1) if n is None else a.reshape(n, 1)


@dataclass
class MMAState:
    """Optimiser memory: xold1, xold2, asymptotes, iteration counter, move limit, recent dJ (§7.6)."""
    xold1: np.ndarray
    xold2: np.ndarray
    low: np.ndarray
    upp: np.ndarray
    iteration: int = 0
    move: float = 0.1
    dJ_hist: list = field(default_factory=list)
    J_prev: float = float("nan")

    def to_dict(self) -> Dict[str, np.ndarray]:
        return {"xold1": self.xold1.ravel().copy(), "xold2": self.xold2.ravel().copy(),
                "low": self.low.ravel().copy(), "upp": self.upp.ravel().copy(),
                "iteration": np.array(self.iteration), "move": np.array(self.move),
                "dJ_hist": np.array(self.dJ_hist, dtype=np.float64), "J_prev": np.array(self.J_prev)}

    @classmethod
    def from_dict(cls, d: Dict[str, np.ndarray]) -> "MMAState":
        return cls(_col(d["xold1"]), _col(d["xold2"]), _col(d["low"]), _col(d["upp"]),
                   int(d["iteration"]), float(d["move"]), [float(v) for v in np.atleast_1d(d["dJ_hist"])],
                   float(d["J_prev"]))


class MMA:
    """Method of Moving Asymptotes, constraints g_i(x) <= 0 (§7.1).

    Parameters: n variables, m constraints, bounds, `move` (fraction of xmax-xmin), asyinit/asyincr/asydecr,
    a0, a_i, c_i, d_i as §7.1 table, `con_scale` (length m; constraints/gradients are multiplied by it so
    they are O(1)).
    """

    def __init__(self, n: int, m: int, xmin=0.0, xmax=1.0, move: float = 0.1, asyinit: float = 0.5,
                 asyincr: float = 1.2, asydecr: float = 0.7, a0: float = 1.0, a=0.0, c=1e3, d=1.0,
                 con_scale=1.0):
        self.n, self.m = n, m
        self.xmin = _col(np.broadcast_to(xmin, (n,)))
        self.xmax = _col(np.broadcast_to(xmax, (n,)))
        self.asyinit, self.asyincr, self.asydecr = asyinit, asyincr, asydecr
        self.a0 = float(a0)
        self.a = _col(np.broadcast_to(a, (m,)))
        self.c = _col(np.broadcast_to(c, (m,)))
        self.d = _col(np.broadcast_to(d, (m,)))
        self.con_scale = _col(np.broadcast_to(con_scale, (m,)))
        self.move0 = move
        self.state: Optional[MMAState] = None

    def _init_state(self, x: np.ndarray) -> None:
        self.state = MMAState(x.copy(), x.copy(), self.xmin.copy(), self.xmax.copy(), 0, self.move0)

    def reset_asymptotes(self) -> None:
        """Restart asymptotes (e.g. at a continuation-phase change): next call behaves like iteration 1."""
        if self.state is not None:
            s = self.state
            s.iteration = 0
            s.low, s.upp = self.xmin.copy(), self.xmax.copy()
            s.dJ_hist, s.J_prev = [], float("nan")

    def step(self, x, f0: float, df0, g=None, dg=None) -> np.ndarray:
        """One MMA update from x with objective f0, gradient df0, constraints g (m,), Jacobian dg (m, n)."""
        from mmapy import mmasub
        x = _col(x, self.n)
        if self.state is None:
            self._init_state(x)
        s = self.state
        g = np.zeros((0, 1)) if self.m == 0 else _col(g, self.m) * self.con_scale
        dg = np.zeros((0, self.n)) if self.m == 0 else np.asarray(dg, dtype=np.float64).reshape(self.m, self.n) * self.con_scale
        s.iteration += 1
        xmma, _, _, _, _, _, _, _, _, low, upp = mmasub(
            self.m, self.n, s.iteration, x, self.xmin, self.xmax, s.xold1, s.xold2, float(f0), _col(df0, self.n),
            g, dg, s.low, s.upp, self.a0, self.a, self.c, self.d, s.move, self.asyinit, self.asydecr, self.asyincr)
        s.low, s.upp = low, upp
        s.xold2, s.xold1 = s.xold1, x.copy()
        return xmma.ravel().copy()

    def note_objective(self, J: float) -> bool:
        """§7.5 oscillation safeguard: halve the move limit if dJ changes sign in >= 6 of the last 8 iterations.
        Returns True if halved."""
        s = self.state
        if np.isfinite(s.J_prev):
            s.dJ_hist.append(float(J) - s.J_prev)
            s.dJ_hist = s.dJ_hist[-9:]
        s.J_prev = float(J)
        h = np.sign(s.dJ_hist)
        if len(h) >= 9 and np.sum(h[1:] * h[:-1] < 0) >= 6:
            s.move *= 0.5
            s.dJ_hist = []
            return True
        return False

    def state_dict(self) -> Dict[str, np.ndarray]:
        d = self.state.to_dict() if self.state is not None else {}
        return {f"mma_{k}": v for k, v in d.items()}

    def load_state_dict(self, d: Dict[str, np.ndarray]) -> None:
        self.state = MMAState.from_dict({k[4:]: v for k, v in d.items() if k.startswith("mma_")})


def restore_mass(xi, volume_fn: Callable[[np.ndarray], float], grad_fn: Callable[[np.ndarray], np.ndarray],
                 V0: float, active: Optional[np.ndarray] = None, tol: float = 1e-4, max_bracket: int = 30):
    """Constant-mass restoration (§7.2): xi <- clip_[0,1](xi+ + eta g/||g||_inf), g = grad V(xi+) on active entries.

    Solves |V/V0 - 1| < tol for eta by Brent. Returns (xi, eta, n_volume_evaluations).
    Raises RuntimeError if no bracket exists (V saturates by clipping before reaching V0).
    """
    xi0 = np.asarray(xi, dtype=np.float64)
    act = np.ones(xi0.shape, bool) if active is None else np.asarray(active, bool)
    gv = np.zeros_like(xi0)
    gv[act] = np.asarray(grad_fn(xi0), dtype=np.float64)[act]
    gmax = np.abs(gv).max()
    if gmax == 0.0:
        raise RuntimeError("restore_mass: zero volume gradient on the active set")
    dirn = gv / gmax
    nev = 0

    def F(eta: float) -> float:
        nonlocal nev
        nev += 1
        return volume_fn(np.clip(xi0 + eta * dirn, 0.0, 1.0)) / V0 - 1.0

    class _Done(Exception):
        pass

    best = {}

    def Fw(eta):
        f = F(eta)
        best["eta"], best["f"] = eta, f
        if abs(f) < tol:
            raise _Done
        return f

    try:
        f0 = Fw(0.0)
        sgn = -1.0 if f0 > 0 else 1.0     # V increases with eta
        lo, flo = 0.0, f0
        step, hi, fhi = 0.05, None, None
        for _ in range(max_bracket):
            e = sgn * step
            fe = Fw(e)
            if fe * f0 < 0:
                hi, fhi = e, fe
                break
            lo, flo = e, fe
            step *= 2.0
        if hi is None:
            raise RuntimeError("restore_mass: no bracket for the volume root (V saturated by clipping)")
        eta = brentq(Fw, min(lo, hi), max(lo, hi), xtol=1e-14, rtol=1e-14, maxiter=100)
        best["eta"] = eta
    except _Done:
        pass
    eta = best["eta"]
    return np.clip(xi0 + eta * dirn, 0.0, 1.0), eta, nev


class ProjectedLBFGSB:
    """Box-projected L-BFGS step with the `MMA.step()` interface (objective gradient only; if `dg` (the
    mass-constraint gradient, shape (1, n)) is given the gradient is projected onto its tangent; other constraints
    ignored; mass held by `restore_mass` afterwards). Free set = variables not on a bound with the descent direction
    pointing outside; L-BFGS two-loop direction on the free set; max-norm step limited by `move`; projection onto
    [xmin, xmax]; curvature pairs kept only if s.y > 0.
    """

    def __init__(self, n: int, m: int = 0, xmin=0.0, xmax=1.0, move: float = 0.1, memory: int = 10, **_):
        self.n = n
        self.xmin = np.broadcast_to(np.asarray(xmin, float), (n,)).copy()
        self.xmax = np.broadcast_to(np.asarray(xmax, float), (n,)).copy()
        self.move, self.memory = move, memory
        self.S, self.Y = [], []
        self._x = self._g = None
        self.iteration = 0

    def reset_asymptotes(self) -> None:
        self.S, self.Y, self._x, self._g = [], [], None, None

    def step(self, x, f0: float, df0, g=None, dg=None) -> np.ndarray:
        x = np.asarray(x, float).ravel()
        gr = np.asarray(df0, float).ravel()
        free = ~(((x <= self.xmin) & (gr > 0)) | ((x >= self.xmax) & (gr < 0)))
        if dg is not None and np.size(dg):
            # project the gradient onto the tangent of the (single, linearised) mass constraint on the free set
            a = np.where(free, np.asarray(dg, float).reshape(-1, self.n)[0], 0.0)
            if a @ a > 0:
                gr = gr - np.where(free, (np.where(free, gr, 0.0) @ a) / (a @ a) * a, 0.0)
            free = ~(((x <= self.xmin) & (gr > 0)) | ((x >= self.xmax) & (gr < 0)))
        self.iteration += 1
        if self._x is not None:
            s, y = x - self._x, gr - self._g
            if s @ y > 1e-12 * np.linalg.norm(s) * np.linalg.norm(y):
                self.S.append(s); self.Y.append(y)
                self.S, self.Y = self.S[-self.memory:], self.Y[-self.memory:]
        self._x, self._g = x.copy(), gr.copy()
        q = np.where(free, gr, 0.0)
        al = []
        for s, y in zip(reversed(self.S), reversed(self.Y)):
            s_, y_ = np.where(free, s, 0.0), np.where(free, y, 0.0)
            rho = 1.0 / (s_ @ y_) if s_ @ y_ > 0 else 0.0
            a_ = rho * (s_ @ q)
            q = q - a_ * y_
            al.append((a_, rho, s_, y_))
        if self.S:
            gam = (self.S[-1] @ self.Y[-1]) / (self.Y[-1] @ self.Y[-1])
        else:
            gam = self.move / max(np.abs(q).max(), 1e-300)      # first step: steepest descent of length `move`
        r = gam * q
        for a_, rho, s_, y_ in reversed(al):
            b = rho * (y_ @ r)
            r = r + s_ * (a_ - b)
        dmax = np.abs(r).max()
        if dmax > self.move:
            r *= self.move / dmax
        return np.clip(x - r, self.xmin, self.xmax)

    def state_dict(self) -> Dict[str, np.ndarray]:
        return {"lb_S": np.array(self.S).reshape(len(self.S), -1), "lb_Y": np.array(self.Y).reshape(len(self.Y), -1),
                "lb_x": np.zeros(0) if self._x is None else self._x, "lb_g": np.zeros(0) if self._g is None else self._g,
                "lb_iteration": np.array(self.iteration)}

    def load_state_dict(self, d: Dict[str, np.ndarray]) -> None:
        self.S, self.Y = list(d["lb_S"]), list(d["lb_Y"])
        self._x = d["lb_x"] if d["lb_x"].size else None
        self._g = d["lb_g"] if d["lb_g"].size else None
        self.iteration = int(d["lb_iteration"])
