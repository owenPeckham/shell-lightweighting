"""Where is Fusion's shelled surface relative to the original solid? (validates 'outside shell = true offset' and finds the split-off mounting regions)"""
import sys, numpy as np, trimesh, igl
from occ_util import DATA_DIR, OUT_DIR, load_step, tessellate
for n, t in [('Cantilever', 1.756), ('L-bracket', 1.700), ('MBB', 2.100)]:
    V0, T0 = tessellate(load_step(f'{DATA_DIR}/Solid_{n}.step'), lin=0.05)
    V1, T1 = tessellate(load_step(f'{DATA_DIR}/Shelled_{n}.step'), lin=0.05)
    m0 = trimesh.Trimesh(V0, T0); m0.merge_vertices(); m1 = trimesh.Trimesh(V1, T1); m1.merge_vertices()
    P, fi = trimesh.sample.sample_surface_even(m1, 40000, seed=1)
    S = igl.signed_distance(P, m0.vertices, m0.faces)[0]
    A = m1.area
    outer = np.abs(S - t) < 0.05; inner = np.abs(S) < 0.05
    print(f'{n}: shelled area {A:.0f} mm2; share of shelled surface at +t (outer skin) {outer.mean():.3f}, at 0 (original surface: void wall or unshelled mount) {inner.mean():.3f}, other {1-outer.mean()-inner.mean():.3f}; median of "other" {np.median(S[~(outer|inner)]):.2f}')
    np.save(f'{OUT_DIR}/fusion_{n}_samples.npy', np.c_[P, S])
