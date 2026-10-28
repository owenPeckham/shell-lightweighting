"""Implicit (signed-distance) variable-thickness shelling under a constant-mass constraint.

Material = { x : -a(x) <= phi(x) <= d(x) }, phi = signed distance to the ORIGINAL solid (negative inside).
  d(x) : outward offset of the outer skin   (current Fusion tool: d = t)
  a(x) : inward inset of the void wall      (current Fusion tool: a = 0; a >= local inradius -> solid)
Wall thickness w(x) = a(x) + d(x). Both fields are partition-of-unity blends of N per-region values,
smoothed over a transition length, so steps between regions become tapers.
"""
import json, time, numpy as np, trimesh, igl
from scipy import ndimage as ndi
from skimage.measure import marching_cubes
from occ_util import DATA_DIR, OUT_DIR, load_step, tessellate, volume

A_SOLID = 50.0  # mm, 'infinite' inset = solid


class ImplicitPart:
    def __init__(self, name, h=0.4, band=5.0):
        t0 = time.time()
        self.name, self.h = name, h
        shp = load_step(f'{DATA_DIR}/{name}.step'); self.V0 = volume(shp)
        V, T = tessellate(shp, lin=0.02)
        self.mesh = (V.astype(np.float64), T.astype(np.int64))
        seg = np.load(f'{OUT_DIR}/seg_{name}.npz'); self.regions = json.load(open(f'{OUT_DIR}/seg_{name}.json'))
        self.seg_lo, self.seg_h = seg['lo'], float(seg['h'])
        lab = seg['label'].astype(np.int32)
        _, ind = ndi.distance_transform_edt(lab == 0, return_indices=True)
        self.lab_ext = lab[tuple(ind)]                     # labels extended outside the part
        coarse_in = seg['inside']
        # fine grid
        pad = 4.0
        lo = V.min(0) - pad; hi = V.max(0) + pad
        self.lo = lo; self.shape = tuple(np.ceil((hi - lo) / h).astype(int) + 1)
        # approximate SDF from the coarse occupancy, upsampled
        din = ndi.distance_transform_edt(coarse_in) * self.seg_h; dout = ndi.distance_transform_edt(~coarse_in) * self.seg_h
        phi_c = (dout - din).astype(np.float32)
        g = [lo[i] + h * np.arange(self.shape[i]) for i in range(3)]
        C = np.stack(np.meshgrid(*[(g[i] - self.seg_lo[i]) / self.seg_h for i in range(3)], indexing='ij'), 0)
        phi = ndi.map_coordinates(phi_c, C.reshape(3, -1), order=1, mode='nearest').reshape(self.shape)
        # exact SDF in a band around the surface
        nb = np.argwhere(np.abs(phi) < band)
        P = lo + h * nb
        S = igl.signed_distance(P, *self.mesh)[0]
        phi[tuple(nb.T)] = S
        self.phi = phi.astype(np.float32)
        self.C = C.astype(np.float32)  # fine-grid coords in coarse index space (for field interpolation)
        self.setup_s = time.time() - t0

    def fields(self, a_r, d_r, ell=3.0):
        """Per-region values -> smooth fields on the fine grid. a_r, d_r: dict region id -> value."""
        nlab = int(self.lab_ext.max()) + 1
        A = np.zeros(nlab, np.float32); D = np.zeros(nlab, np.float32)
        for r in self.regions:
            A[r['id']] = a_r[r['id']]; D[r['id']] = d_r[r['id']]
        a_c = A[self.lab_ext]; d_c = D[self.lab_ext]
        sig = ell / self.seg_h / 2.0
        a_c = ndi.gaussian_filter(a_c, sig); d_c = ndi.gaussian_filter(d_c, sig)
        a = ndi.map_coordinates(a_c, self.C.reshape(3, -1), order=1, mode='nearest').reshape(self.shape)
        d = ndi.map_coordinates(d_c, self.C.reshape(3, -1), order=1, mode='nearest').reshape(self.shape)
        return a, d

    def psi(self, a, d):
        return np.maximum(self.phi - d, -self.phi - a)

    def volume(self, psi):
        return float(np.clip(0.5 - psi / self.h, 0, 1).sum() * self.h ** 3)

    def solve_d(self, a_r, d_shape, ell=3.0, tol=1e-3):
        """Scale factor s so that d_r = s * d_shape keeps V = V0 (bisection; V is monotonic in s)."""
        def V(s):
            a, d = self.fields(a_r, {k: s * v for k, v in d_shape.items()}, ell)
            return self.volume(self.psi(a, d))
        from scipy.optimize import brentq
        cnt = [0]
        def f(s):
            cnt[0] += 1; return V(s) - self.V0
        s = brentq(f, 0.0, 6.0, xtol=tol)
        return s, cnt[0]

    def export(self, psi, path):
        v, f, _, _ = marching_cubes(-psi, 0.0, spacing=(self.h,) * 3)
        m = trimesh.Trimesh(v + self.lo, f[:, ::-1]); m.export(path)
        return m
