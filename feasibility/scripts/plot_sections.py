import os, sys, json, numpy as np
from occ_util import OUT_DIR
os.chdir(OUT_DIR)
import trimesh, matplotlib; matplotlib.use('Agg'); import matplotlib.pyplot as plt
name = sys.argv[1]; designs = sys.argv[2].split(','); axis = int(sys.argv[3]) if len(sys.argv) > 3 else 2
seg = np.load(f'seg_{name}.npz'); R = json.load(open(f'seg_{name}.json')); lab = seg['label']
joints = [r for r in R if r['kind'] == 'joint']; big = max(joints, key=lambda r: r['vol_mm3'])
cen = seg['lo'] + seg['h'] * np.argwhere(lab == big['id']).mean(0)
normal = np.eye(3)[axis]
solid = trimesh.load(f'{name}.stl')
fig, axs = plt.subplots(1, len(designs) + 1, figsize=(5 * (len(designs) + 1), 5))
D = json.load(open(f'designs_{name}.json'))
titles = {'uniform': 'uniform (current tool)', 'joints_original': 'joints kept solid, original size', 'joints_solid': 'joints solid AND offset outward'}
for ax, dsg in zip(axs, ['solid'] + designs):
    m = solid if dsg == 'solid' else trimesh.load(f'{name}_{dsg}.stl')
    sec = m.section(plane_origin=cen, plane_normal=normal)
    keep = [i for i in range(3) if i != axis]
    from matplotlib.path import Path; from matplotlib.patches import PathPatch
    verts, codes = [], []
    for e3 in sec.discrete:
        e = e3[:, keep]; verts += list(e) + [e[0]]; codes += [Path.MOVETO] + [Path.LINETO] * (len(e) - 1) + [Path.CLOSEPOLY]
    ax.add_patch(PathPatch(Path(verts, codes), facecolor='0.72', edgecolor='k', lw=0.6))
    ax.autoscale_view()
    ax.set_aspect('equal'); ax.set_xlabel('xyz'[keep[0]] + ' (mm)', fontsize=8); ax.set_ylabel('xyz'[keep[1]] + ' (mm)', fontsize=8); ax.tick_params(labelsize=7)
    t = 'original solid' if dsg == 'solid' else f"{titles.get(dsg, dsg)}\nstrut wall {D[dsg]['strut_wall_mm']:.2f} mm"
    ax.set_title(t, fontsize=10)
plt.suptitle(f'{name}: section through largest joint (plane {"xyz"[axis]} = {cen[axis]:.1f} mm), all at constant mass', fontsize=11)
plt.tight_layout(); plt.savefig(f'sections_{name}.png', dpi=110)
print('ok', cen)
