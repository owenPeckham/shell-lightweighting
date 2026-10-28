"""
Signed-distance field of the original part and the implicit variable-thickness design.

phi(x)  signed distance to the original surface (mm, negative inside), exact (libigl, winding-number
        sign) within `band_mm` of the surface, nearest-coarse-EDT estimate further away.

A design assigns each region a solidity theta_i in {0 (hollow), 1 (solid)}. With the Gaussian blend
G_l of transition length l:

    theta(x) = G_l * theta_i(x),        Rin(x) = G_l * Rin_i(x)   (Rin_i: region depth + margin)
    a(x) = theta(x) Rin(x)              inward inset of the void wall
    d(x) = s (1 - theta(x))             outward offset of the outer skin (s: one scalar per design)
    material  <=>  psi(x) = max(phi - d, -phi - a) <= 0

The part's volume V(s) = sum clip(1/2 - psi/h, 0, 1) h^3 rises monotonically with s, so the
constant-mass constraint V(s) = V0 has one root, found by Brent's method.
"""
from __future__ import annotations

import time
import numpy as np
from scipy import ndimage
from scipy.optimize import brentq


def _winding_inside(V, T, pts):
    import igl
    w = igl.fast_winding_number(np.ascontiguousarray(V, np.float64), np.ascontiguousarray(T, np.int64),
                                np.ascontiguousarray(pts, np.float64))
    return w > 0.5


class PartGrid:
    """Fine signed-distance grid plus the coarse segmentation grid on the same origin."""

    def __init__(self, V, T, h=0.4, k_fe=2, h_seg=1.0, band=4.0, pad=4.0, verbose=True):
        import igl
        t0 = time.time()
        self.V, self.T = V, T
        self.h, self.k, self.h_seg = float(h), int(k_fe), float(h_seg)
        lo = V.min(0) - pad
        hi = V.max(0) + pad
        n = np.ceil((hi - lo) / h).astype(int) + 1
        n = (np.ceil(n / self.k) * self.k).astype(int)          # FE blocks tile the grid exactly
        self.lo, self.shape = lo, tuple(int(v) for v in n)
        # coarse grid (segmentation + far-field distance)
        nc = np.ceil((hi - lo) / h_seg).astype(int) + 2
        self.shape_c = tuple(int(v) for v in nc)
        axes = [lo[i] + h_seg * np.arange(nc[i]) for i in range(3)]
        G = np.stack(np.meshgrid(*axes, indexing="ij"), -1).reshape(-1, 3)
        self.inside_c = _winding_inside(V, T, G).reshape(self.shape_c)
        din = ndimage.distance_transform_edt(self.inside_c) * h_seg
        dout = ndimage.distance_transform_edt(~self.inside_c) * h_seg
        self.depth_c = din.astype(np.float32)                    # depth below the surface (coarse)
        phi_c = (dout - din).astype(np.float32)
        # fine grid: nearest coarse estimate, then exact values in the band
        self.ix = [np.clip(np.round(np.arange(self.shape[i]) * h / h_seg).astype(int), 0, self.shape_c[i] - 1)
                   for i in range(3)]
        phi = phi_c[np.ix_(*self.ix)]
        nb = np.argwhere(np.abs(phi) < band + 1.5 * h_seg)
        P = lo + h * nb
        S = igl.signed_distance(P, np.ascontiguousarray(V, np.float64), np.ascontiguousarray(T, np.int64))[0]
        phi[tuple(nb.T)] = S
        self.phi = phi.astype(np.float32)
        self.n_band = len(nb)
        self.t_setup = time.time() - t0
        if verbose:
            print(f"  SDF grid {self.shape} = {np.prod(self.shape) / 1e6:.1f}M cells at {h} mm "
                  f"({self.n_band / 1e6:.2f}M exact), {self.t_setup:.0f} s")

    # --- active subset: the only cells whose material state can change ----------------------
    def prepare_active(self, s_max):
        """Cells with phi < s_max + h: every other cell is outside the largest possible offset (void)."""
        self.s_max = float(s_max)
        self.active = np.flatnonzero(self.phi.ravel() < s_max + self.h)
        self.phi_a = self.phi.ravel()[self.active]
        return len(self.active)

    def design_fields_active(self, label_ext_c, solid_by_label, rin_by_label, transition_mm, mount_zone_c=None,
                             mount_mask_c=None):
        theta, rin = self.design_fields(label_ext_c, solid_by_label, rin_by_label, transition_mm, mount_zone_c, mount_mask_c)
        return theta.ravel()[self.active].copy(), rin.ravel()[self.active].copy()

    def psi_active(self, theta_a, rin_a, s, forced_a=None):
        out = np.maximum(self.phi_a - s * (1.0 - theta_a), -self.phi_a - theta_a * rin_a)
        if forced_a is not None and forced_a.any():
            out[forced_a] = np.minimum(out[forced_a], -self.h)
        return out

    def volume_active(self, psi_a) -> float:
        return float(np.clip(0.5 - psi_a / self.h, 0, 1).sum(dtype=np.float64) * self.h ** 3)

    def solve_offset_active(self, theta_a, rin_a, V0, forced_a=None, tol=1e-4):
        """Outward offset s in [0, s_max] with V(s) = V0 (Brent). (nan, n) when no root exists."""
        cnt = [0]

        def f(s):
            cnt[0] += 1
            return self.volume_active(self.psi_active(theta_a, rin_a, s, forced_a)) - V0
        f0 = f(0.0)
        if abs(f0) <= tol * V0:
            return 0.0, cnt[0]
        if f0 > 0 or f(self.s_max) < 0:
            return float("nan"), cnt[0]
        return float(brentq(f, 0.0, self.s_max, xtol=1e-3, rtol=1e-6)), cnt[0]

    def full_psi(self, psi_a):
        out = self.phi.copy()                      # outside the active set psi = phi - d >= phi - s_max > 0
        out.ravel()[self.active] = psi_a
        return out

    # --- region fields -------------------------------------------------------------------------
    def coarse_to_fine(self, arr_c):
        return arr_c[np.ix_(*self.ix)]

    def solid_volume(self) -> float:
        return float(np.clip(0.5 - self.phi / self.h, 0, 1).sum() * self.h ** 3)

    def design_fields(self, label_ext_c, solid_by_label, rin_by_label, transition_mm, mount_zone_c=None, mount_mask_c=None):
        """theta and Rin on the fine grid for a per-label solidity table (arrays indexed by label).

        mount_zone_c: coarse mask of the mounts grown by the transition length; it is treated as solid before the
        blur, so the solid-to-hollow taper lies wholly in the candidate region. mount_mask_c: the mounts themselves,
        where theta = 1 and Rin = its region value are imposed after the blur (mounts are never changed)."""
        lab = self.coarse_to_fine(label_ext_c)
        sig = transition_mm / self.h / 2.0
        S = solid_by_label.astype(np.float32)[lab]
        if mount_zone_c is not None:
            np.maximum(S, self.coarse_to_fine(mount_zone_c).astype(np.float32), out=S)
        theta = ndimage.gaussian_filter(S, sig)
        del S
        R0 = rin_by_label.astype(np.float32)[lab]
        rin = ndimage.gaussian_filter(R0, sig)
        if mount_mask_c is not None:
            m = self.coarse_to_fine(mount_mask_c)
            theta[m] = 1.0
            rin[m] = R0[m]
        del lab, R0
        return theta, rin

    def psi(self, theta, rin, s, forced_solid=None):
        out = np.maximum(self.phi - s * (1.0 - theta), -self.phi - theta * rin)
        if forced_solid is not None and forced_solid.any():
            out[forced_solid] = np.minimum(out[forced_solid], -self.h)
        return out

    def fraction(self, psi):
        return np.clip(0.5 - psi / self.h, 0, 1)

    def volume_of(self, psi) -> float:
        return float(self.fraction(psi).sum(dtype=np.float64) * self.h ** 3)

    def solve_offset(self, theta, rin, V0, s_max=8.0, forced_solid=None, tol=1e-4):
        """Scalar outward offset s with V(s) = V0. Returns (s, n_evals) or (nan, n) if no root in [0, s_max]."""
        cnt = [0]

        def f(s):
            cnt[0] += 1
            return self.volume_of(self.psi(theta, rin, s, forced_solid)) - V0
        f0 = f(0.0)
        if abs(f0) <= tol * V0:
            return 0.0, cnt[0]
        if f0 > 0:
            return float("nan"), cnt[0]
        if f(s_max) < 0:
            return float("nan"), cnt[0]
        s = brentq(f, 0.0, s_max, xtol=1e-3, rtol=1e-6)
        return float(s), cnt[0]

    def small_voids(self, psi, min_volume):
        """Enclosed void components (not connected to the outside) and those below min_volume."""
        void = psi > 0
        lab, n = ndimage.label(void)
        if n == 0:
            return np.zeros_like(void), [], lab
        border = np.unique(np.concatenate([lab[0].ravel(), lab[-1].ravel(), lab[:, 0].ravel(), lab[:, -1].ravel(),
                                           lab[:, :, 0].ravel(), lab[:, :, -1].ravel()]))
        sizes = ndimage.sum(void, lab, range(1, n + 1)) * self.h ** 3
        enclosed = [i + 1 for i in range(n) if (i + 1) not in set(border.tolist())]
        small = [i for i in enclosed if sizes[i - 1] < min_volume]
        fill = np.isin(lab, small) if small else np.zeros_like(void)
        voids = [dict(label=int(i), volume_mm3=float(sizes[i - 1])) for i in enclosed if i not in small]
        return fill, voids, lab

    def fe_fractions(self, psi):
        """Solid fraction of each FE cell (k^3 block of fine cells) and the FE grid origin and size."""
        k = self.k
        fr = self.fraction(psi)
        nx, ny, nz = (np.array(fr.shape) // k)
        vf = fr[:nx * k, :ny * k, :nz * k].reshape(nx, k, ny, k, nz, k).mean((1, 3, 5), dtype=np.float64)
        origin = self.lo - 0.5 * self.h        # fine values are cell-centred samples
        return vf.astype(np.float32), origin, k * self.h

    def fe_cell_labels(self, label_ext_c, cell_ijk):
        """Region label at each FE cell centre."""
        k = self.k
        centre_fine = cell_ijk * k + k // 2
        cc = [self.ix[i][np.clip(centre_fine[:, i], 0, self.shape[i] - 1)] for i in range(3)]
        return label_ext_c[cc[0], cc[1], cc[2]]

    def export_mesh(self, psi, path=None, step=1):
        """Marching-cubes surface of psi = 0 (optionally subsampled by `step`)."""
        from skimage.measure import marching_cubes
        import trimesh
        a = psi[::step, ::step, ::step]
        v, f, _, _ = marching_cubes(-a, 0.0, spacing=(self.h * step,) * 3)
        m = trimesh.Trimesh(v + self.lo, f[:, ::-1], process=True)
        if path:
            m.export(path)
        return m
