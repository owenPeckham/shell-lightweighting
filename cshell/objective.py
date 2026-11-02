"""Responses, capacity soft-min and the objective J with gradients w.r.t. ρ̂ on the active elements.

PLAN_single_run §4 (Owen 2026-10-31): the default objective is compliance at constant mass, J = ln(C0/C) = ln(K/K0),
with buckling and strength as constraints relative to the solid, ln(μ̃/μ̃0) = ln(λ_B0/λ_B) ≤ 0 and ln(σ_PN/σ_PN0) ≤ 0
(`evaluate_compliance`). The weighted stiffness/soft-min capacity objective (`evaluate`) stays as objective.type
"weighted". `shear_flexure_ratio` is the skeleton-free R_v that drives the growth of the hollow-allowed set.

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
    fe._scatter_dofs(rhs, ge)  # fixed-order scatter (two-grid: colours)
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
    eig_tol: float = 1e-3                 # dual-norm residual; eigenvalue error ≈ 1e-6 (THEORY §4.4(c))
    rho_min: float = 1e-6
    objective: str = "weighted"           # "weighted" (Stage 2) | "compliance" (PLAN_single_run §4)
    constraints: tuple = ()               # subset of ("buckling", "strength") for objective "compliance"


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
    parts = dict(stress=sp, buckling=bk, primal=pr, dC=dC, dlnlam=dlnlam, shares=w.tolist(),
                 dlnsig=sp.grad / sp.value, dlnmu=dmu / mu_t if mu_t > 0 else None)   # per-mode floors (optimise)
    return Evaluation(J=float(J), dJ=dJ, C=C, sigma_pn=sp.value, lam_s=lam_s, mu_tilde=mu_t, lam_b=lam_b, lam=lam,
                      governing=governing, parts=parts)


# ---------------------------------------------------------------------------------------------------------------------
# PLAN_single_run §4: compliance objective, buckling and strength constraints relative to the solid
# ---------------------------------------------------------------------------------------------------------------------
@dataclass
class ConstrainedEvaluation:
    J: float                           # ln(C0/C) (0 without base)
    dJ: np.ndarray                     # ∂J/∂ρ̂ per active element
    C: float
    sigma_pn: float                    # NaN when strength is not evaluated
    mu_tilde: float                    # 0 when buckling is not evaluated or no positive μ
    g: dict                            # {"buckling": ln(μ̃/μ̃0), "strength": ln(σ_PN/σ_PN0)} (only with base)
    dg: dict                           # their ∂/∂ρ̂ per active element
    governing: str                     # the constraint with the largest g (or "stiffness")
    parts: dict = field(default_factory=dict)


def evaluate_compliance(fe: SparseFE, rho_hat_active, rho_active, m_active, st: Settings, base: dict | None = None,
                        prev_modes=None, want_stress: bool | None = None, want_buckling: bool | None = None
                        ) -> ConstrainedEvaluation:
    """Factor, solve, C and dC; σ_PN (p-norm outside the exclusion zones, c = 1) and μ̃ (block-Krylov) when the
    constraint is on (or forced by want_*). With `base` = {C0, sigma_pn0, mu0} the objective J = ln(C0/C) and the
    constraints g_i = ln(r_i) ≤ 0 with their gradients. A constraint whose solid value is undefined (μ̃0 = 0: no
    compressive mode) is reported as g = −∞ (inactive)."""
    fe.factor(rho_hat_active)
    pr = fe.primal()
    C = pr.compliance
    dC = fe.dC_drho_hat(pr)
    do_s = ("strength" in st.constraints) if want_stress is None else want_stress
    do_b = (("buckling" in st.constraints) and st.buckling) if want_buckling is None else want_buckling
    sp = bk = None
    sig, mu = float("nan"), 0.0
    if do_s and not float(fe.xp.asarray(m_active).sum()) > 0:     # exclusion zones cover the part: no strength term
        do_s = False
    if do_s:
        sp = stress_pnorm(fe, pr, rho_active, m_active, p=st.stress_p, q=st.stress_q, quantile=st.quantile,
                          rho_min=st.rho_min)
        sig = sp.value
    if do_b:
        bk = buckling(fe, pr, num_modes=st.num_modes, p=st.buckling_p, min_density=st.min_density, tol=st.eig_tol,
                      X0=prev_modes)
        mu = bk.mu_tilde
    g, dg = {}, {}
    if base is None:
        J, dJ = 0.0, np.zeros_like(dC)
    else:
        J, dJ = float(np.log(base["C0"] / C)), -dC / C
        if sp is not None and "strength" in st.constraints:
            g["strength"] = float(np.log(sig / base["sigma_pn0"]))
            dg["strength"] = sp.grad / sig
        elif "strength" in st.constraints:                    # no element outside the exclusion zones: inactive
            g["strength"], dg["strength"] = -np.inf, np.zeros_like(dC)
        if bk is not None and "buckling" in st.constraints:
            if base.get("mu0", 0.0) > 0 and mu > 0:
                g["buckling"] = float(np.log(mu / base["mu0"]))
                dg["buckling"] = bk.dmu_tilde / mu
            else:
                g["buckling"], dg["buckling"] = -np.inf, np.zeros_like(dC)
    governing = max(g, key=g.get) if g and max(g.values()) > -np.inf else "stiffness"
    return ConstrainedEvaluation(J=J, dJ=dJ, C=C, sigma_pn=sig, mu_tilde=mu, g=g, dg=dg, governing=governing,
                                 parts=dict(stress=sp, buckling=bk, primal=pr, dC=dC))


# ---------------------------------------------------------------------------------------------------------------------
# PLAN_single_run §4: shear/flexure ratio R_v (dshell's R without a skeleton)
# ---------------------------------------------------------------------------------------------------------------------
def wall_frames(N: np.ndarray):
    """Unit tangents (t1, t2) with (t1, t2, n) right-handed; t1 ⟂ the coordinate axis least aligned with n."""
    n = N / np.maximum(np.linalg.norm(N, axis=1, keepdims=True), 1e-300)
    ax = np.eye(3)[np.argmin(np.abs(n), axis=1)]
    t1 = np.cross(n, ax)
    t1 /= np.maximum(np.linalg.norm(t1, axis=1, keepdims=True), 1e-300)
    return t1, np.cross(n, t1), n


def wall_energies(sampler, X, N, d_out, r_in, E: float, nu: float, n_samples: int = 16, chunk: int = 1 << 18,
                  xp=np):
    """Through-thickness energies per surface point (PLAN_single_run §4):
        U_flex  = ∫ [(σ_11² + σ_22²)/2E + σ_12²/2G] ds,   U_shear = ∫ (σ_1n² + σ_2n²)/2G ds
    along x + s n over the material, s from +d_out (outer surface) to −r_in (the void surface a if open, else the medial axis r_MA), midpoint rule with `n_samples` points, in the
    wall frame (t1, t2, n). `sampler(points (k,3)) -> (σ (k,6) Voigt xx,yy,zz,yz,xz,xy, ρ (k,))` (array module xp).
    Returns host (U_flex, U_shear) per point (MPa·mm = N/mm per unit area)."""
    G = E / (2.0 * (1.0 + nu))
    t1, t2, n = wall_frames(np.asarray(N, float))
    X = np.asarray(X, float)
    L = np.asarray(d_out, float) + np.asarray(r_in, float)
    nv = len(X)
    Uf, Us = np.zeros(nv), np.zeros(nv)
    frac = (np.arange(n_samples) + 0.5) / n_samples
    per = max(1, chunk // n_samples)
    for a in range(0, nv, per):
        b = min(nv, a + per)
        s = np.asarray(d_out, float)[a:b, None] - frac[None, :] * L[a:b, None]          # (m, ns)
        P = X[a:b, None, :] + s[..., None] * n[a:b, None, :]
        sig, rho = sampler(xp.asarray(P.reshape(-1, 3)))
        sig = sig.reshape(b - a, n_samples, 6)
        rho = rho.reshape(b - a, n_samples)
        T = xp.stack([xp.stack([sig[..., 0], sig[..., 5], sig[..., 4]], -1),
                      xp.stack([sig[..., 5], sig[..., 1], sig[..., 3]], -1),
                      xp.stack([sig[..., 4], sig[..., 3], sig[..., 2]], -1)], -2)          # (m, ns, 3, 3)
        e1, e2, en = (xp.asarray(v[a:b])[:, None, :] for v in (t1, t2, n))

        def c(u, v):                                     # uᵀ T v per sample (u, v: (m, 1, 3) wall-frame axes)
            return (u * (T @ v[..., None])[..., 0]).sum(-1)
        s11, s22, s12, s1n, s2n = c(e1, e1), c(e2, e2), c(e1, e2), c(e1, en), c(e2, en)
        w = xp.ones_like(rho) * xp.asarray(L[a:b] / n_samples)[:, None]     # ρ only normalises σ (sampler)
        Uf[a:b] = _host(((s11 * s11 + s22 * s22) / (2 * E) + s12 * s12 / (2 * G)) * w).sum(1)
        Us[a:b] = _host(((s1n * s1n + s2n * s2n) / (2 * G)) * w).sum(1)
    return Uf, Us


class ElementStressSampler:
    """Material stress at points from FE element-centre stresses: trilinear interpolation between element centres,
    weighted by ρ_e (void elements carry no weight), σ(x) = Σ w ρ σ / Σ w ρ, ρ(x) = Σ w ρ (w trilinear weights).
    `sig_active` (Ne, 6) and `rho_active` (Ne,) on the active elements `elem_dense` of the element box `elem_shape`
    with corner node 0 at `origin` and spacing h."""

    def __init__(self, sig_active, rho_active, elem_dense, elem_shape, origin, h: float, xp=np):
        self.xp, self.h = xp, float(h)
        self.shape = tuple(int(n) for n in elem_shape)
        self.origin = xp.asarray(np.asarray(origin, float))
        self.sig = xp.asarray(sig_active, dtype=xp.float64)
        self.rho = xp.asarray(rho_active, dtype=xp.float64)
        idx = np.full(int(np.prod(self.shape)), -1, np.int64)
        idx[np.asarray(elem_dense)] = np.arange(len(elem_dense))
        self.idx = xp.asarray(idx)

    def __call__(self, P):
        xp = self.xp
        q = (P - self.origin) / self.h - 0.5                       # element-centre lattice coordinates
        i0 = xp.floor(q).astype(xp.int64)
        f = q - i0
        num = xp.zeros((len(P), 6))
        den = xp.zeros(len(P))
        rho = xp.zeros(len(P))
        nx, ny, nz = self.shape
        for c in range(8):
            o = ((c >> 2) & 1, (c >> 1) & 1, c & 1)
            ijk = [i0[:, k] + o[k] for k in range(3)]
            w = xp.ones(len(P))
            for k in range(3):
                w = w * (f[:, k] if o[k] else 1.0 - f[:, k])
            ok = (ijk[0] >= 0) & (ijk[0] < nx) & (ijk[1] >= 0) & (ijk[1] < ny) & (ijk[2] >= 0) & (ijk[2] < nz)
            flat = xp.where(ok, (ijk[0] * ny + ijk[1]) * nz + ijk[2], 0)
            a = xp.where(ok, self.idx[flat], -1)
            r = xp.where(a >= 0, self.rho[xp.maximum(a, 0)], 0.0)
            wr = w * r
            num += wr[:, None] * self.sig[xp.maximum(a, 0)]
            den += wr
            rho += wr
        return num / xp.maximum(den, 1e-300)[:, None], rho


def shear_flexure_ratio(Uf, Us, eps_rel: float = 1e-12):
    """R_v = U_shear/U_flex (0 where both vanish)."""
    Uf = np.asarray(Uf, float)
    return np.asarray(Us, float) / np.maximum(Uf, eps_rel * max(float(Uf.max(initial=0.0)), 1e-300))
