import sys, numpy as np, trimesh, matplotlib; matplotlib.use('Agg')
import matplotlib.pyplot as plt
from mpl_toolkits.mplot3d.art3d import Poly3DCollection
from occ_util import DATA_DIR, OUT_DIR, load_step, tessellate
fig = plt.figure(figsize=(15, 9))
names = ['Solid_Cantilever', 'Solid_L-bracket', 'Solid_MBB']
for k, n in enumerate(names):
    V, T = tessellate(load_step(f'{DATA_DIR}/{n}.step'), lin=0.2)
    m = trimesh.Trimesh(V, T, process=True); m.merge_vertices()
    print(n, 'watertight', m.is_watertight, 'vol', round(m.volume, 1), 'ntri', len(m.faces), 'bounds', m.bounds.round(1).tolist())
    m.export(f'{OUT_DIR}/{n}.stl')
    for j, (el, az) in enumerate([(25, -60), (90, -90)]):
        ax = fig.add_subplot(2, 3, 1 + k + 3 * j, projection='3d')
        tri = m.vertices[m.faces]
        nrm = m.face_normals; shade = 0.35 + 0.65 * np.clip(nrm @ np.array([0.3, -0.5, 0.8]) / 1.0, 0, 1)
        pc = Poly3DCollection(tri, facecolors=plt.cm.Greys(0.3 + 0.5 * shade), edgecolor='none')
        ax.add_collection3d(pc); b = m.bounds; c = b.mean(0); r = (b[1] - b[0]).max() / 2
        ax.set_xlim(c[0] - r, c[0] + r); ax.set_ylim(c[1] - r, c[1] + r); ax.set_zlim(c[2] - r, c[2] + r)
        ax.view_init(el, az); ax.set_title(f'{n} el={el}'); ax.set_xlabel('x'); ax.set_ylabel('y')
plt.tight_layout(); plt.savefig(f'{OUT_DIR}/solids.png', dpi=80)
