"""Responses, capacity soft-min and the objective J with gradients w.r.t. ρ̂ on the active elements.

THEORY §4.4, §5.1, §6.2-6.5. All responses use the as-built ρ̂ and one factorisation of K (cshell.fe_sparse):
primal, stress adjoint and buckling (+ its adjoint) share it. Single load case (Owen's D3): no case weights ω_k.
"""
from __future__ import annotations

from dataclasses import dataclass, field

import numpy as np

from .fe_sparse import Primal, SparseFE, _host, _scatter_add, buckling


def von_mises(xp, s):
    """s: (..., 6) Voigt (xx, yy, zz, yz, xz, xy) → von Mises and ∂s/∂σ (THEORY §6.3)."""
    sx, sy, sz, tyz, txz, txy = (s[..., i] for i in range(6))
    vm = xp.sqrt(xp.maximum(sx * sx + sy * sy + sz * sz - sx * sy - sy * sz - sz * sx
                            + 3.0 * (tyz * tyz + txz * txz + txy * txy), 1e-300))
    ds = xp.stack([2 * sx - sy - sz, 2 * sy - sx - sz, 2 * sz - sx - sy, 6 * tyz, 6 * txz, 6 * txy], -1) / (2 * vm[..., None])
    return vm, ds


@dataclass
class StressPN:
    value: float              # σ_PN (MPa)
    grad: np.ndarray          # dσ_PN/dρ̂_e per active element (host)
    vm: np.ndarray            # material von Mises per active element (host)
    quantile: float           # Q_q(s | ρ_e ≥ 0.5, m_e = 1), the normalisation target (THEORY §4.4(b))
    adjoint_s: float = 0.0


def stress_pnorm(fe: SparseFE, primal: Primal, rho, m, p: float = 8.0, q: float = 2.0, quantile: float = 0.995,
                 rho_min: float = 1e-6) -> StressPN:
    """σ_PN = (Σ v_e s_e^p / Σ v_e)^{1/p}, v_e = ρ_e^q m_e, s_e = von Mises of the material centre stress D⁰B(0)u_e.

    Adjoint (THEORY §6.3): K η = −∂σ_PN/∂u with the same factor; dσ_PN/dρ̂_e = ∂σ_PN/∂v_e ∂v_e/∂ρ̂_e + η_eᵀK⁰_e u_e.
    `rho` (ρ_e, not ρ̂) and `m` (0/1 weights) are per active element."""
    import time
    xp = fe.xp
    rho = xp.asarray(rho, dtype=xp.float64)
    m = xp.asarray(m, dtype=xp.float64)
    sig = fe.centre_stress(primal.u, solid=True)                        # (Ne, 6)
    s, dsds = von_mises(xp, sig)
    v = rho ** q * m
    Sv = float(v.sum())
    if Sv <= 0:
        raise ValueError("stress p-norm: no element carries weight")
    smax = float((s * (m > 0)).max())
    r = s / smax
    A = float((v * r ** p).sum()) / Sv                                  # (σ_PN/smax)^p
    sig_pn = smax * A ** (1.0 / p)
    # ∂σ_PN/∂u: element vectors σ^{1−p}/Σv · v_e s_e^{p−1} (∂s/∂σ)ᵀ D B(0)
    coef = (sig_pn / smax) ** (1.0 - p) / Sv * v * r ** (p - 1.0)          # σ^{1−p} s^{p−1} = (σ/smax)^{1−p} r^{p−1}
    DB = fe.E * xp.asarray(_unit_DB(fe))                                 # (6, 24)
    ge = (coef[:, None] * dsds) @ DB                                      # (Ne, 24)
    t0 = time.time()
    rhs = xp.zeros(fe.g.ndof)
    _scatter_add(xp, rhs, fe._dofs.reshape(-1), ge.reshape(-1))
    eta = fe.solve(-rhs)
    t_adj = time.time() - t0
    dv = (sig_pn / p / Sv) * ((r ** p) / A - 1.0)          # = σ^{1−p}(s^p − σ^p)/(pΣv), overflow-safe
    dvdrho = q * rho ** (q - 1.0) * m / (1.0 - rho_min)
    grad = dv * dvdrho + fe.E * xp.einsum("ei,ij,ej->e", eta[fe._dofs], fe._Khat, primal.u[fe._dofs])
    s_h, rho_h, m_h = _host(s), _host(rho), _host(m)
    sel = (rho_h >= 0.5) & (m_h > 0)
    Q = float(np.quantile(s_h[sel], quantile)) if sel.any() else float("nan")
    return StressPN(value=float(sig_pn), grad=_host(grad), vm=s_h, quantile=Q, adjoint_s=t_adj)


def _unit_DB(fe: SparseFE) -> np.ndarray:
    from .fe_sparse import centre_B, unit_D
    return unit_D(fe.nu) @ centre_B(fe.h)


@dataclass
class Settings:
    """Objective settings for one continuation phase (THEORY §5.1, §4.4, §10.2)."""
    w_K: float = 0.5
    w_P: float = 0.5
    sigma_f: float = 36.6
    knockdown: float = 0.85
    beta: float = 12.0
    stress_p: float = 8.0
    stress_q: float = 2.0
    quantile: float = 0.995
    buckling: bool = True
    num_modes: int = 6
    buckling_p: float = 8.0
    min_density: float = 0.1
    eig_tol: float = 1e-6
    rho_min: float = 1e-6


@dataclass
class Evaluation:
    J: float
    dJ: np.ndarray                     # ∂J/∂ρ̂ per active element
    C: float
    sigma_pn: float
    lam_s: float
    mu_tilde: float
    lam_b: float                       # κ_b/μ̃ (∞ if μ̃ = 0)
    lam: float                         # soft-min capacity factor
    governing: str
    parts: dict = field(default_factory=dict)


def capacity(lam_s: float, mu_tilde: float, kappa: float, beta: float):
    """ln λ = −(1/β) ln T, T = λ_s^{−β} + (μ̃/κ)^β (THEORY §4.4(d)), via log-sum-exp; returns (λ, weights) where the
    weights are the shares of the two terms in T (they multiply the per-term log derivatives)."""
    terms = [-beta * np.log(lam_s)]
    if mu_tilde > 0:
        terms.append(beta * np.log(mu_tilde / kappa))
    t = np.array(terms)
    lse = t.max() + np.log(np.exp(t - t.max()).sum())
    w = np.exp(t - lse)
    return float(np.exp(-lse / beta)), w


def evaluate(fe: SparseFE, rho_hat_active, rho_active, m_active, st: Settings, base: dict | None = None,
             c_norm: float = 1.0, prev_modes=None) -> Evaluation:
    """Factor, solve, responses, J and ∂J/∂ρ̂ (THEORY §6.5). `base` holds the solid's C0 and λ0 (J = 0 there); with
    base None, J = 0 and only the responses are returned (used to evaluate the baseline itself)."""
    fe.factor(rho_hat_active)
    pr = fe.primal()
    C = pr.compliance
    dC = fe.dC_drho_hat(pr)
    sp = stress_pnorm(fe, pr, rho_active, m_active, p=st.stress_p, q=st.stress_q, quantile=st.quantile,
                      rho_min=st.rho_min)
    lam_s = st.sigma_f / (c_norm * sp.value)
    mu_t, dmu, bk = 0.0, None, None
    if st.buckling:
        bk = buckling(fe, pr, num_modes=st.num_modes, p=st.buckling_p, min_density=st.min_density, tol=st.eig_tol,
                      X0=prev_modes)
        mu_t, dmu = bk.mu_tilde, bk.dmu_tilde
    lam, w = capacity(lam_s, mu_t, st.knockdown, st.beta)
    # d ln λ = w_s d ln λ_s − w_b d ln μ̃  (T-shares w); d ln λ_s = −dσ/σ
    dlnlam = -w[0] * sp.grad / sp.value
    if mu_t > 0:
        dlnlam = dlnlam - w[1] * dmu / mu_t
    governing = "strength" if (mu_t == 0 or w[0] >= w[1]) else "buckling"
    lam_b = st.knockdown / mu_t if mu_t > 0 else float("inf")
    if base is None:
        J, dJ = 0.0, np.zeros_like(dC)
    else:
        wsum = st.w_K + st.w_P
        J = (st.w_K * np.log(base["C0"] / C) + st.w_P * np.log(lam / base["lam0"])) / wsum
        dJ = (-st.w_K * dC / C + st.w_P * dlnlam) / wsum
    parts = dict(stress=sp, buckling=bk, primal=pr, dC=dC, dlnlam=dlnlam, shares=w.tolist())
    return Evaluation(J=float(J), dJ=dJ, C=C, sigma_pn=sp.value, lam_s=lam_s, mu_tilde=mu_t, lam_b=lam_b, lam=lam,
                      governing=governing, parts=parts)
