"""Design map ξ → ξ̃ → (d, a, a_ab) → ψ → χ → ρ → ρ̂, volume, sliver measure, and their reverse-mode pull-backs.

THEORY §3.2-3.10, §5.2, §6.1, with Owen's decision D1 (STATUS.md §4): the inset's lower bound depends on the outward
offset, a ≥ a_lo(d) = w·softplus((t_min − d)/w), so a wall can lie wholly outside Γ₀ once d ≥ t_min while w = a + d ≥
t_min still holds (up to the smoothing width w). The inset is a_i = a_lo(d_i) + (a^hi_i − a_lo(d_i)) ξ̃ᵃ_i.

Arrays may be NumPy (CPU, float64, tests) or CuPy (GPU); `xp` is taken from the cell arrays. Sparse P is
scipy.sparse or cupyx.scipy.sparse CSR (n_cells × N_v, 3 nnz per row).
"""
from __future__ import annotations

from dataclasses import dataclass
from typing import Any

import numpy as np


# -- scalar profiles (THEORY §3.7, §3.9, §5.2) ------------------------------------------------------------------------
def heaviside(xp, s, eps):
    """Quintic C² smoothed Heaviside H_ε(s) and its derivative (∫H′ = 1)."""
    z = xp.clip(s / eps, -1.0, 1.0)
    H = 0.5 + z * (15.0 / 16.0 - z * z * (5.0 / 8.0 - z * z * 3.0 / 16.0))
    dH = (15.0 / 16.0) * (1.0 - z * z) ** 2 / eps
    dH = xp.where(xp.abs(s) < eps, dH, 0.0)
    return H, dH


def smoothstep(xp, z):
    """S(z) = 3z² − 2z³ on clip(z, 0, 1) and S′(z)."""
    zc = xp.clip(z, 0.0, 1.0)
    return zc * zc * (3.0 - 2.0 * zc), xp.where((z > 0) & (z < 1), 6.0 * zc * (1.0 - zc), 0.0)


def softplus(xp, x, w):
    """w·log(1 + e^{x/w}) and its derivative σ(x/w), overflow-safe."""
    y = x / w
    sp = w * (xp.maximum(y, 0.0) + xp.log1p(xp.exp(-xp.abs(y))))
    return sp, 1.0 / (1.0 + xp.exp(-y))


def sliver_bump(xp, z):
    """B(z) = 16 z²(1 − z)² on 0 < z < 1, else 0 (THEORY §5.2), and B′."""
    inside = (z > 0) & (z < 1)
    B = xp.where(inside, 16.0 * z * z * (1.0 - z) ** 2, 0.0)
    dB = xp.where(inside, 32.0 * z * (1.0 - z) * (1.0 - 2.0 * z), 0.0)
    return B, dB


# -- data ------------------------------------------------------------------------------------------------------------
@dataclass
class DesignSpace:
    """Per-vertex data on Γ_h (THEORY §3.4-3.5). `filt` has apply(x) = Fx and adjoint(g) = Fᵀg."""
    r_ma: Any
    a_hi: Any
    d_hi: Any
    area: Any
    passive: Any                 # bool: fixed at ξ̃ᵃ = 1, ξ̃ᵈ = 0 after filtering
    filt: Any
    t_min: float = 1.5
    v_min: float = 3.0
    b_v: float = 0.2             # as-built bias, void side
    b_o: float = 0.0             # as-built bias, outer side
    soft_w: float = 0.1          # D1 smoothing width of a_lo(d) (mm)


@dataclass
class Discretisation:
    """Geometry-grid band and its FE blocking (THEORY §3.1, §3.6, §3.7)."""
    phi: Any                     # (n_c,) signed distance at cell centres (mm, negative inside)
    P: Any                       # (n_c, N_v) CSR interpolation at the closest point
    elem: Any                    # (n_c,) int64 flat FE element id of each cell (in the FE element box)
    n_elem: int                  # number of FE elements in the box
    k: int                       # FE element = k³ geometry cells
    h_g: float
    rho_min: float = 1e-6
    V_extra: float = 0.0         # interior volume not covered by the band (THEORY §3.10)

    @property
    def eps(self) -> float:
        return self.h_g


@dataclass
class State:
    """Forward values kept for the pull-back."""
    xit_d: Any
    xit_a: Any
    d: Any
    a: Any
    a_used: Any                  # a_ab (as-built) or a (nominal)
    a_lo: Any
    da_lo: Any
    dab_da: Any                  # ∂a_ab/∂a (1 for nominal)
    psi: Any
    outer: Any                   # outer branch of the max active
    chi: Any
    dchi: Any                    # ∂χ/∂(−ψ) = H′
    as_built: bool


def _xp(a):
    try:
        import cupy as cp
        return cp.get_array_module(a)
    except ImportError:
        return np


class DesignMap:
    """ξ = (ξᵈ, ξᵃ) on all vertices (passive entries are ignored) → ρ̂ on the FE element box, V and G_s."""

    def __init__(self, ds: DesignSpace, disc: Discretisation):
        self.ds, self.disc = ds, disc
        self.xp = _xp(disc.phi)
        self.PT = disc.P.T.tocsr()

    # -- vertex fields --------------------------------------------------------------------------------------------
    def vertex_fields(self, xi_d, xi_a):
        xp, ds = self.xp, self.ds
        xt_d = xp.clip(ds.filt.apply(xi_d), 0.0, 1.0)
        xt_a = xp.clip(ds.filt.apply(xi_a), 0.0, 1.0)
        xt_d = xp.where(ds.passive, 0.0, xt_d)
        xt_a = xp.where(ds.passive, 1.0, xt_a)
        d = ds.d_hi * xt_d
        a_lo, da_lo = softplus(xp, ds.t_min - d, ds.soft_w)          # D1; da_lo/dd = −σ
        a = a_lo + (ds.a_hi - a_lo) * xt_a
        return xt_d, xt_a, d, a, a_lo, -da_lo

    def forward(self, xi_d, xi_a, as_built: bool = True) -> State:
        xp, ds, disc = self.xp, self.ds, self.disc
        xt_d, xt_a, d, a, a_lo, da_lo = self.vertex_fields(xi_d, xi_a)
        if as_built:
            S, dS = smoothstep(xp, (ds.r_ma - a) / ds.b_v)
            a_used, dab = a - ds.b_v * S, 1.0 + dS                   # THEORY §3.9
            b_o = ds.b_o
        else:
            a_used, dab, b_o = a, xp.ones_like(a), 0.0
        d_c = disc.P @ d
        a_c = disc.P @ a_used
        o = disc.phi - d_c + b_o
        i = -disc.phi - a_c
        outer = o >= i                                                # ties to the outer branch (THEORY §6.1)
        psi = xp.where(outer, o, i)
        chi, dchi = heaviside(xp, -psi, disc.eps)
        return State(xt_d, xt_a, d, a, a_used, a_lo, da_lo, dab, psi, outer, chi, dchi, as_built)

    # -- outputs --------------------------------------------------------------------------------------------------
    def rho(self, st: State):
        """ρ_e = mean of χ over the k³ cells of e (cells outside the band count as empty)."""
        xp, disc = self.xp, self.disc
        return xp.bincount(disc.elem, weights=st.chi, minlength=disc.n_elem) / disc.k ** 3

    def rho_hat(self, st: State):
        return self.disc.rho_min + (1.0 - self.disc.rho_min) * self.rho(st)

    def volume(self, st: State) -> float:
        """V = h_g³ Σ χ (+ interior outside the band), float64 (THEORY §3.10). Use a nominal state."""
        return float(self.xp.sum(st.chi, dtype=self.xp.float64)) * self.disc.h_g ** 3 + self.disc.V_extra

    def sliver(self, st: State):
        """G_s = (1/A_Γ) Σ A_i B(z_i), z_i = (r_MA,i − a_i)/(v_min/2), nominal a; returns (G_s, ∂G_s/∂a)."""
        xp, ds = self.xp, self.ds
        half = 0.5 * ds.v_min
        B, dB = sliver_bump(xp, (ds.r_ma - st.a) / half)
        act = ~ds.passive
        A = ds.area * act
        AG = float(A.sum())
        return float((A * B).sum() / AG), -(A / AG) * dB / half

    # -- reverse mode ---------------------------------------------------------------------------------------------
    def pull_chi(self, st: State, g_chi):
        """∂R/∂ξ given ∂R/∂χ_c (THEORY §6.1 table, bottom-up)."""
        g_psi = -st.dchi * g_chi
        g_dc = self.xp.where(st.outer, -g_psi, 0.0)
        g_ac = self.xp.where(st.outer, 0.0, -g_psi)
        return self.pull_vertex(st, self.PT @ g_dc, self.PT @ g_ac)

    def pull_vertex(self, st: State, g_d, g_aused):
        """∂R/∂ξ given ∂R/∂d and ∂R/∂a_used per vertex (a_used = a_ab or a)."""
        xp, ds = self.xp, self.ds
        g_a = g_aused * st.dab_da
        g_d = g_d + g_a * st.da_lo * (1.0 - st.xit_a)                 # a = a_lo(d)(1 − ξ̃ᵃ) + a_hi ξ̃ᵃ
        g_xt_d = g_d * ds.d_hi
        g_xt_a = g_a * (ds.a_hi - st.a_lo)
        g_xt_d = xp.where(ds.passive, 0.0, g_xt_d)
        g_xt_a = xp.where(ds.passive, 0.0, g_xt_a)
        return ds.filt.adjoint(g_xt_d), ds.filt.adjoint(g_xt_a)

    def pull_rho_hat(self, st: State, g_rho_hat):
        """∂R/∂ξ given ∂R/∂ρ̂_e on the FE element box."""
        disc = self.disc
        g_chi = (1.0 - disc.rho_min) * g_rho_hat[disc.elem] / disc.k ** 3
        return self.pull_chi(st, g_chi)

    def volume_grad(self, st: State):
        return self.pull_chi(st, self.xp.full(st.chi.shape, self.disc.h_g ** 3))

    def sliver_grad(self, st: State):
        _, g_a = self.sliver(st)
        st_nom = st if not st.as_built else State(**{**st.__dict__, "dab_da": self.xp.ones_like(st.a)})
        return self.pull_vertex(st_nom, self.xp.zeros_like(st.d), g_a)


def passive_thin(a_hi, t_min: float, v_min: float, margin: float = 0.05):
    """Vertices too thin to host a void ≥ v_min under D1 (a ≥ 0 once d ≥ t_min): a^hi below max(t_min, v_min/2)."""
    return a_hi <= max(t_min, 0.5 * v_min) + margin
