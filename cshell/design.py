"""Design map ξ → ξ̃ → (d, a, a_ab) → ψ → χ → ρ → ρ̂, volume, sliver measure, and their reverse-mode pull-backs.

THEORY §3.2-3.10, §5.2, §6.1, with Owen's decision D1 (STATUS.md §4): the inset's lower bound depends on the outward
offset, a ≥ a_lo(d) = w·softplus((t_min − d)/w), so a wall can lie wholly outside Γ₀ once d ≥ t_min while w = a + d ≥
t_min still holds (up to the smoothing width w). The inset is a_i = a_lo(d_i) + (a^hi_i − a_lo(d_i)) ξ̃ᵃ_i.

Arrays may be NumPy (CPU, float64, tests) or CuPy (GPU); `xp` is taken from the cell arrays. Sparse P is
scipy.sparse or cupyx.scipy.sparse CSR (n_cells × N_v, 3 nnz per row).

`DeviceDesignMap` is the production GPU chain: per-vertex arithmetic stays on the host (NumPy, the filter is a
host factorisation), the band lives on the device in compact form (`DeviceDisc`: φ float32, closest Γ_h triangle
int32, raw barycentrics 3 × float32, FE element int32 = 24 B per cell) and fused CUDA kernels evaluate ψ, χ, ρ, V
and the pull-backs to vertices in float64 without storing any per-cell state (ψ is recomputed in the reverse pass).
"""
from __future__ import annotations

import time
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
    b_o: float = 0.0             # outer bias used in ψ (b_o as-built, 0 nominal)
    dev: Any = None              # DeviceDesignMap: (d, a_used) on the device; cell fields are recomputed, not stored


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
        self.cxp = self.xp                                            # array module of cell/element-box arrays
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
        xp, ds = self.xp, self.ds
        xt_d, xt_a, d, a, a_lo, da_lo = self.vertex_fields(xi_d, xi_a)
        if as_built:
            S, dS = smoothstep(xp, (ds.r_ma - a) / ds.b_v)
            a_used, dab = a - ds.b_v * S, 1.0 + dS                   # THEORY §3.9
            b_o = ds.b_o
        else:
            a_used, dab, b_o = a, xp.ones_like(a), 0.0
        st = State(xt_d, xt_a, d, a, a_used, a_lo, da_lo, dab, None, None, None, None, as_built, b_o)
        return self._cells(st)

    def _cells(self, st: State) -> State:
        """ψ = max(φ − P d + b_o, −φ − P a_used) and χ = H_ε(−ψ) per band cell (THEORY §3.7-3.9)."""
        xp, disc = self.xp, self.disc
        d_c = disc.P @ st.d
        a_c = disc.P @ st.a_used
        o = disc.phi - d_c + st.b_o
        i = -disc.phi - a_c
        st.outer = o >= i                                             # ties to the outer branch (THEORY §6.1)
        st.psi = xp.where(st.outer, o, i)
        st.chi, st.dchi = heaviside(xp, -st.psi, disc.eps)
        return st

    def solid_volume(self) -> float:
        """V₀ = h_g³ Σ H_ε(−φ) + V_extra of the solid part (ψ = φ, THEORY §3.8, §3.10)."""
        chi0, _ = heaviside(self.xp, -self.disc.phi, self.disc.eps)
        return float(chi0.sum()) * self.disc.h_g ** 3 + self.disc.V_extra

    def solid_rho(self, b_o: float):
        """ρ of the as-built solid (ψ_ab = φ + b_o, THEORY §3.9) on the element box."""
        disc = self.disc
        chi0, _ = heaviside(self.xp, -(disc.phi + b_o), disc.eps)
        return self.xp.bincount(disc.elem, weights=chi0, minlength=disc.n_elem) / disc.k ** 3

    def elem_mean(self, v):
        """Per element: mean over its band cells of (P v)_c (0 where no cell); used for the mount stress mask."""
        disc = self.disc
        num = self.xp.bincount(disc.elem, weights=disc.P @ v, minlength=disc.n_elem)
        return num / self.xp.maximum(self.xp.bincount(disc.elem, minlength=disc.n_elem), 1)

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



# -- GPU chain (compact band, fused kernels) -------------------------------------------------------------------------
_KERNELS = r"""
extern "C" {
__device__ __forceinline__ double cell_psi(long long c, const float* phi, const int* face, const float* bary,
        const int* Th, const double* d, const double* a, double bo, int solid,
        bool* outer, int* v, double* w) {
    double ph = (double)phi[c];
    if (solid) { *outer = true; return ph + bo; }
    int t = face[c];
    v[0] = Th[3 * t]; v[1] = Th[3 * t + 1]; v[2] = Th[3 * t + 2];
    double b0 = fmax((double)bary[3 * c], 0.0), b1 = fmax((double)bary[3 * c + 1], 0.0),
           b2 = fmax((double)bary[3 * c + 2], 0.0);
    double sb = (b0 + b1) + b2;
    w[0] = b0 / sb; w[1] = b1 / sb; w[2] = b2 / sb;
    double dc = ((w[0] * d[v[0]]) + w[1] * d[v[1]]) + w[2] * d[v[2]];
    double ac = ((w[0] * a[v[0]]) + w[1] * a[v[1]]) + w[2] * a[v[2]];
    double o = ph - dc + bo;
    double i = -ph - ac;
    *outer = o >= i;
    return *outer ? o : i;
}

__device__ __forceinline__ double block_sum(double x) {
    __shared__ double sh[32];
    for (int k = 16; k > 0; k >>= 1) x += __shfl_down_sync(0xffffffff, x, k);
    int lane = threadIdx.x & 31, wid = threadIdx.x >> 5;
    if (lane == 0) sh[wid] = x;
    __syncthreads();
    x = 0.0;
    if (wid == 0) {
        x = (lane < (blockDim.x >> 5)) ? sh[lane] : 0.0;
        for (int k = 16; k > 0; k >>= 1) x += __shfl_down_sync(0xffffffff, x, k);
    }
    return x;
}

// χ = H_ε(−ψ): Σχ into vol[0]; if want_rho, χ scattered into rho[elem] (THEORY §3.7, §3.10)
__global__ void geo_fwd(long long n, const float* phi, const int* face, const float* bary, const int* Th,
        const int* elem, const double* d, const double* a, double bo, double eps, int solid, int want_rho,
        double* rho, double* vol) {
    double acc = 0.0;
    int v[3]; double w[3]; bool outer;
    for (long long c = blockIdx.x * (long long)blockDim.x + threadIdx.x; c < n;
         c += (long long)gridDim.x * blockDim.x) {
        double s = -cell_psi(c, phi, face, bary, Th, d, a, bo, solid, &outer, v, w);
        double z = fmin(fmax(s / eps, -1.0), 1.0);
        double H = 0.5 + z * (15.0 / 16.0 - z * z * (5.0 / 8.0 - z * z * 3.0 / 16.0));
        if (H != 0.0) {
            acc += H;
            if (want_rho) atomicAdd(rho + elem[c], H);
        }
    }
    acc = block_sum(acc);
    if (threadIdx.x == 0) atomicAdd(vol, acc);
}

// reverse: g_χ = (gmul g_el[elem]) / gdiv (want_el) or gconst; g_ψ = −H′ g_χ; Pᵀ(−g_ψ) on the active branch
__global__ void geo_rev(long long n, const float* phi, const int* face, const float* bary, const int* Th,
        const int* elem, const double* d, const double* a, double bo, double eps, int want_el, const double* gel,
        double gmul, double gdiv, double gconst, double* gd, double* ga) {
    int v[3]; double w[3]; bool outer;
    for (long long c = blockIdx.x * (long long)blockDim.x + threadIdx.x; c < n;
         c += (long long)gridDim.x * blockDim.x) {
        double s = -cell_psi(c, phi, face, bary, Th, d, a, bo, 0, &outer, v, w);
        if (fabs(s) < eps) {
            double z = fmin(fmax(s / eps, -1.0), 1.0);
            double t = 1.0 - z * z;
            double dH = (15.0 / 16.0) * (t * t) / eps;
            double gchi = want_el ? (gmul * gel[elem[c]]) / gdiv : gconst;
            double g = dH * gchi;                                 // = −g_ψ
            if (g != 0.0) {
                double* out = outer ? gd : ga;
                atomicAdd(out + v[0], w[0] * g);
                atomicAdd(out + v[1], w[1] * g);
                atomicAdd(out + v[2], w[2] * g);
            }
        }
    }
}

// Re-scoring (collect, THEORY §8.6.1): nominal ψ = φ_E (exported mesh, has_phiE) or max(φ − Pd, −φ − Pa) of the
// vertex fields, as-built ψ_ab = ψ + Δ with Δ = b_o on the outer branch of max(φ − Pd, −φ − Pa), P sv on the void
// branch. Cells whose positive-weight vertices are all closed (P cl ≥ 1 − 1e-9) use a = a_closed (export closure
// rule, §8.2). solid: ψ = φ_E (or φ), Δ = b_o. vol[0] += Σ H(−ψ); rho[elem] += H(−ψ_ab) when want_rho.
__global__ void rescore_fwd(long long n, const float* phi, const int* face, const float* bary, const int* Th,
        const int* elem, const float* phiE, int has_phiE, const double* d, const double* a, const double* cl,
        const double* sv, double bo, double eps, double a_closed, int solid, int want_rho, double* rho, double* vol) {
    double acc = 0.0;
    for (long long c = blockIdx.x * (long long)blockDim.x + threadIdx.x; c < n;
         c += (long long)gridDim.x * blockDim.x) {
        double ph = (double)phi[c];
        double pn, dl;
        if (solid) {
            pn = has_phiE ? (double)phiE[c] : ph;
            dl = bo;
        } else {
            int t = face[c];
            int v0 = Th[3 * t], v1 = Th[3 * t + 1], v2 = Th[3 * t + 2];
            double b0 = fmax((double)bary[3 * c], 0.0), b1 = fmax((double)bary[3 * c + 1], 0.0),
                   b2 = fmax((double)bary[3 * c + 2], 0.0);
            double sb = (b0 + b1) + b2;
            double w0 = b0 / sb, w1 = b1 / sb, w2 = b2 / sb;
            double dc = ((w0 * d[v0]) + w1 * d[v1]) + w2 * d[v2];
            double ac = ((w0 * a[v0]) + w1 * a[v1]) + w2 * a[v2];
            double cc = ((w0 * cl[v0]) + w1 * cl[v1]) + w2 * cl[v2];
            double sc = ((w0 * sv[v0]) + w1 * sv[v1]) + w2 * sv[v2];
            if (cc >= 1.0 - 1e-9) ac = a_closed;
            double o = ph - dc, i = -ph - ac;
            bool outer = o >= i;
            pn = has_phiE ? (double)phiE[c] : (outer ? o : i);
            dl = outer ? bo : sc;
        }
        double z = fmin(fmax(-pn / eps, -1.0), 1.0);
        acc += 0.5 + z * (15.0 / 16.0 - z * z * (5.0 / 8.0 - z * z * 3.0 / 16.0));
        if (want_rho) {
            z = fmin(fmax(-(pn + dl) / eps, -1.0), 1.0);
            double H = 0.5 + z * (15.0 / 16.0 - z * z * (5.0 / 8.0 - z * z * 3.0 / 16.0));
            if (H != 0.0) atomicAdd(rho + elem[c], H);
        }
    }
    acc = block_sum(acc);
    if (threadIdx.x == 0) atomicAdd(vol, acc);
}

// Σ_c∈e (P v)_c and the cell count per element
__global__ void geo_elem_sum(long long n, const int* face, const float* bary, const int* Th, const int* elem,
        const double* f, double* num, double* cnt) {
    for (long long c = blockIdx.x * (long long)blockDim.x + threadIdx.x; c < n;
         c += (long long)gridDim.x * blockDim.x) {
        int t = face[c];
        double b0 = fmax((double)bary[3 * c], 0.0), b1 = fmax((double)bary[3 * c + 1], 0.0),
               b2 = fmax((double)bary[3 * c + 2], 0.0);
        double sb = (b0 + b1) + b2;
        double fc = ((b0 / sb) * f[Th[3 * t]] + (b1 / sb) * f[Th[3 * t + 1]]) + (b2 / sb) * f[Th[3 * t + 2]];
        atomicAdd(num + elem[c], fc);
        atomicAdd(cnt + elem[c], 1.0);
    }
}
}
"""
_MOD = None


def _kernels():
    global _MOD
    if _MOD is None:
        import cupy as cp
        _MOD = cp.RawModule(code=_KERNELS, options=("--fmad=false",))      # NumPy-identical rounding per operation
    return _MOD


class DeviceDisc:
    """Compact band for the GPU chain (THEORY §3.1, §3.6, §3.7): φ float32 (mm), closest Γ_h triangle int32, raw
    float32 barycentrics (clipped at 0 and normalised in float64 inside the kernels, as `band.P_from_compact` does on
    the CPU), FE element int32: 24 B per cell. The arrays live in pinned host memory and, when it fits, in device
    memory. `release()` drops the device copy while the FE factor needs the memory; `upload()` restores it. If the
    upload does not fit beside the FE, the kernels read the pinned host copy directly (zero-copy over NVLink-C2C on
    GH200, `mode == "mapped"`) for the rest of the phase, with a log line."""

    _CELL = ("phi", "face", "bary", "elem")

    def __init__(self, phi, face, bary, Th, elem, n_v: int, n_elem: int, k: int, h_g: float, rho_min: float = 1e-6,
                 V_extra: float = 0.0, log=print):
        import cupy as cp
        import cupyx
        if n_elem >= 2 ** 31:
            raise ValueError("element box too large for int32 element ids")
        self.n, self.n_v, self.n_elem, self.k, self.h_g = len(phi), int(n_v), int(n_elem), int(k), float(h_g)
        self.rho_min, self.V_extra, self.log = float(rho_min), float(V_extra), log

        def pinned(a, dt):
            a = np.ascontiguousarray(a, dt)
            out = cupyx.empty_pinned(a.shape, dt)         # cudaHostAlloc: device-addressable under UVA
            out[...] = a
            return out

        self._host = dict(phi=pinned(phi, np.float32), face=pinned(face, np.int32), bary=pinned(bary, np.float32),
                          elem=pinned(elem, np.int32))
        self.Th = cp.asarray(np.ascontiguousarray(Th, np.int32))
        self._dev: dict | None = None
        self.mapped = False
        self.upload()

    @property
    def eps(self) -> float:
        return self.h_g

    @property
    def nbytes(self) -> int:
        return sum(a.nbytes for a in self._host.values())

    @property
    def resident(self) -> bool:
        return self._dev is not None

    @property
    def mode(self) -> str:
        return "device" if self._dev is not None else "mapped" if self.mapped else "released"

    def upload(self) -> float:
        """Copy the band to the device unless resident or mapped; returns seconds."""
        import cupy as cp
        if self._dev is not None or self.mapped:
            return 0.0
        t0 = time.time()
        try:
            self._dev = {k: cp.asarray(v) for k, v in self._host.items()}
            cp.cuda.Device().synchronize()
        except cp.cuda.memory.OutOfMemoryError:
            self._dev = None
            cp.get_default_memory_pool().free_all_blocks()
            self.use_mapped(f"band ({self.nbytes / 2**30:.2f} GiB) does not fit beside the FE factor")
        return time.time() - t0

    def use_mapped(self, why: str = "requested") -> None:
        """Kernels read the pinned host arrays in place (no device copy) from now on."""
        import cupy as cp
        self._dev = None
        cp.get_default_memory_pool().free_all_blocks()
        if not self.mapped:
            self.log(f"[cshell]   geometry chain: {why}; kernels read it from pinned host memory (zero-copy)")
        self.mapped = True

    def release(self) -> None:
        """Free the device copy (host copy kept) and return the blocks to the driver."""
        import cupy as cp
        self._dev = None
        cp.get_default_memory_pool().free_all_blocks()

    def __getattr__(self, name):
        if name in DeviceDisc._CELL:
            if self._dev is not None:
                return self._dev[name]
            if self.mapped:                              # raw device-visible pointer to the pinned host array
                return np.uint64(self._host[name].ctypes.data)
            raise RuntimeError("DeviceDisc: band released from the device; call upload() first")
        raise AttributeError(name)


class DeviceDesignMap(DesignMap):
    """`DesignMap` with the cell part on the GPU (`DeviceDisc`). Same results to rounding: per-vertex fields, the
    sliver term and the filter pull-back are the inherited host code; ψ, χ, ρ, V and Pᵀ pull-backs are fused kernels
    with float64 arithmetic and float64 atomics (summation order differs from NumPy's)."""

    THREADS = 256

    def __init__(self, ds: DesignSpace, disc: DeviceDisc):
        import cupy as cp
        self.ds, self.disc = ds, disc
        self.xp, self.cxp = np, cp
        self._k = _kernels()

    def _grid(self) -> tuple:
        import cupy as cp
        sm = cp.cuda.Device().attributes["MultiProcessorCount"]
        return (max(1, min(-(-self.disc.n // self.THREADS), 32 * sm)),), (self.THREADS,)

    def _fwd(self, d, a, b_o: float, solid: bool, want_rho: bool):
        cp, D = self.cxp, self.disc
        rho = cp.zeros(D.n_elem if want_rho else 1, cp.float64)
        vol = cp.zeros(1, cp.float64)
        if solid:
            d = a = cp.zeros(1, cp.float64)
        self._k.get_function("geo_fwd")(*self._grid(), (
            np.int64(D.n), D.phi, D.face, D.bary, D.Th, D.elem, d, a, np.float64(b_o), np.float64(D.eps),
            np.int32(solid), np.int32(want_rho), rho, vol))
        return rho, float(vol.get()[0])

    def _cells(self, st: State) -> State:
        st.dev = (self.cxp.asarray(st.d, dtype=np.float64), self.cxp.asarray(st.a_used, dtype=np.float64))
        return st

    def rho(self, st: State) -> np.ndarray:
        rho, _ = self._fwd(*st.dev, st.b_o, False, True)
        return (rho / self.disc.k ** 3).get()

    def volume(self, st: State) -> float:
        _, s = self._fwd(*st.dev, st.b_o, False, False)
        return s * self.disc.h_g ** 3 + self.disc.V_extra

    def solid_volume(self) -> float:
        _, s = self._fwd(None, None, 0.0, True, False)
        return s * self.disc.h_g ** 3 + self.disc.V_extra

    def solid_rho(self, b_o: float) -> np.ndarray:
        rho, _ = self._fwd(None, None, b_o, True, True)
        return (rho / self.disc.k ** 3).get()

    def elem_mean(self, v) -> np.ndarray:
        cp, D = self.cxp, self.disc
        num, cnt = cp.zeros(D.n_elem), cp.zeros(D.n_elem)
        self._k.get_function("geo_elem_sum")(*self._grid(), (
            np.int64(D.n), D.face, D.bary, D.Th, D.elem, cp.asarray(v, dtype=np.float64), num, cnt))
        return (num / cp.maximum(cnt, 1)).get()

    def _rev(self, st: State, g_el, gmul: float, gdiv: float, gconst: float):
        cp, D = self.cxp, self.disc
        gd, ga = cp.zeros(D.n_v), cp.zeros(D.n_v)
        want = g_el is not None
        g_el = cp.asarray(g_el, dtype=np.float64) if want else cp.zeros(1)
        self._k.get_function("geo_rev")(*self._grid(), (
            np.int64(D.n), D.phi, D.face, D.bary, D.Th, D.elem, *st.dev, np.float64(st.b_o), np.float64(D.eps),
            np.int32(want), g_el, np.float64(gmul), np.float64(gdiv), np.float64(gconst), gd, ga))
        return self.pull_vertex(st, gd.get(), ga.get())

    def rescore(self, phiE, d, a, cl, sv, b_o: float, a_closed: float, solid: bool = False, want_rho: bool = True):
        """Collect's re-scoring cell map on the device (kernel `rescore_fwd`): ρ of ψ_ab = ψ + Δ (host, element box)
        and the nominal volume h_g³ Σ H(−ψ) + V_extra. phiE: per-cell φ_E (float32, device) or None (ψ from the
        vertex fields d, a, cl). Vertex arrays d, a, cl (closed 0/1), sv (= b_v S) are float64."""
        cp, D = self.cxp, self.disc
        rho = cp.zeros(D.n_elem if want_rho else 1, cp.float64)
        vol = cp.zeros(1, cp.float64)
        nv = D.n_v

        def dv(x):
            return cp.zeros(1, cp.float64) if x is None else cp.asarray(np.broadcast_to(x, (nv,)), dtype=np.float64)
        pe = cp.zeros(1, cp.float32) if phiE is None else cp.asarray(phiE, dtype=cp.float32)
        if phiE is not None and pe.size != D.n:
            raise ValueError(f"phiE has {pe.size} cells, the band {D.n}")
        self._k.get_function("rescore_fwd")(*self._grid(), (
            np.int64(D.n), D.phi, D.face, D.bary, D.Th, D.elem, pe, np.int32(phiE is not None), dv(d), dv(a), dv(cl),
            dv(sv), np.float64(b_o), np.float64(D.eps), np.float64(a_closed), np.int32(solid), np.int32(want_rho),
            rho, vol))
        V = float(vol.get()[0]) * D.h_g ** 3 + D.V_extra
        return ((rho / D.k ** 3).get() if want_rho else None), V

    def pull_chi(self, st: State, g_chi):
        raise NotImplementedError("DeviceDesignMap stores no per-cell state; use pull_rho_hat or volume_grad")

    def pull_rho_hat(self, st: State, g_rho_hat):
        return self._rev(st, g_rho_hat, 1.0 - self.disc.rho_min, float(self.disc.k ** 3), 0.0)

    def volume_grad(self, st: State):
        return self._rev(st, None, 0.0, 1.0, self.disc.h_g ** 3)


def passive_thin(a_hi, t_min: float, v_min: float, margin: float = 0.05):
    """Vertices too thin to host a void ≥ v_min under D1 (a ≥ 0 once d ≥ t_min): a^hi below max(t_min, v_min/2)."""
    return a_hi <= max(t_min, 0.5 * v_min) + margin
