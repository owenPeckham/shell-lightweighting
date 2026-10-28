import os, sys, json, numpy as np
from occ_util import OUT_DIR
os.chdir(OUT_DIR)  # inputs are read from the output folder
import matplotlib; matplotlib.use('Agg'); import matplotlib.pyplot as plt
names = sys.argv[1:]
fig, axs = plt.subplots(1, len(names), figsize=(6 * len(names), 6), subplot_kw=dict(projection='3d'))
axs = np.atleast_1d(axs)
for ax, n in zip(axs, names):
    d = np.load(f'seg_{n}.npz'); R = {r['id']: r for r in json.load(open(f'seg_{n}.json'))}
    lab = d['label']; idx = np.argwhere(lab > 0); P = d['lo'] + d['h'] * idx; L = lab[tuple(idx.T)]
    sub = np.random.default_rng(0).random(len(P)) < 0.25; P, L = P[sub], L[sub]
    kind = np.array([R[l]['kind'] for l in L])
    cm = plt.cm.tab20
    lam = np.array([R[l].get('lam', 0) for l in L]); cmv = plt.cm.viridis
    col = np.array([(0.6, 0.6, 0.6, 1) if k == 'mount' else ((0.85, 0.1, 0.1, 1) if k == 'joint' else cmv(min(lm, 60) / 60)) for k, l, lm in zip(kind, L, lam)])
    ax.scatter(*P.T, c=col, s=1.5, depthshade=False)
    c = P.mean(0); r = np.ptp(P, 0).max() / 2
    ax.set_xlim(c[0]-r, c[0]+r); ax.set_ylim(c[1]-r, c[1]+r); ax.set_zlim(c[2]-r, c[2]+r); ax.view_init(35, -60)
    ns = sum(1 for r in R.values() if r['kind'] == 'strut'); nj = sum(1 for r in R.values() if r['kind'] == 'joint')
    ax.set_title(f'{n}: {nj} joints (red), {ns} struts (colour = slenderness, 0-60+), mounts (grey)', fontsize=10); ax.set_axis_off()
plt.tight_layout(); plt.savefig('segmentation.png', dpi=90)
