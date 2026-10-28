from occ_util import DATA_DIR, OUT_DIR
"""Uniform (reproduces the Fusion tool) vs region-wise variable thickness, at constant mass."""
import sys, json, time, numpy as np
from implicit_shell import ImplicitPart, A_SOLID
name, t_fusion = sys.argv[1], float(sys.argv[2])
p = ImplicitPart(name, h=float(sys.argv[3]) if len(sys.argv) > 3 else 0.4)
print(f'{name}: grid {p.shape} ({np.prod(p.shape)/1e6:.1f}M cells), setup {p.setup_s:.1f}s, V0 {p.V0:.0f}')
R = {r['id']: r for r in p.regions}
solid_check = p.volume(p.psi(np.full(p.shape, A_SOLID, np.float32), np.zeros(p.shape, np.float32)))
print(f'  solid volume from field {solid_check:.0f} vs B-rep {p.V0:.0f} ({100*(solid_check/p.V0-1):+.2f}%)')
mount = {k: r['kind'] == 'mount' for k, r in R.items()}
out = {}
def design(label, a_r):
    t0 = time.time()
    d_shape = {k: (0.0 if mount[k] else 1.0) for k in R}
    s, n = p.solve_d(a_r, d_shape)
    a, d = p.fields(a_r, {k: s * v for k, v in d_shape.items()}); psi = p.psi(a, d)
    m = p.export(psi, f'{OUT_DIR}/{name}_{label}.stl')
    frac_solid = sum(R[k]['vol_mm3'] for k in R if a_r[k] >= A_SOLID and not mount[k]) / sum(R[k]['vol_mm3'] for k in R if not mount[k])
    out[label] = dict(d_mm=round(s, 3), strut_wall_mm=round(s, 3), field_evals=n, secs=round(time.time() - t0, 1), stl_volume=round(m.volume), V0=round(p.V0), watertight=bool(m.is_watertight), solid_share_of_shellable=round(frac_solid, 3))
    print(' ', label, out[label], flush=True)
# 1. uniform: every non-mount region hollow, a = 0
design('uniform', {k: (A_SOLID if mount[k] else 0.0) for k in R})
# 2. joints solid, struts hollow
design('joints_solid', {k: (A_SOLID if (mount[k] or R[k]['kind'] == 'joint') else 0.0) for k in R})
# 3. joints solid + low-slenderness struts (lambda < 15) solid
design('joints_and_stubby_solid', {k: (A_SOLID if (mount[k] or R[k]['kind'] == 'joint' or R[k].get('lam', 99) < 15) else 0.0) for k in R})
# 4. joints keep their ORIGINAL solid geometry (no outward growth), struts hollow
def design_keep(label, keep):
    t0 = time.time()
    a_r = {k: (A_SOLID if (mount[k] or keep(R[k])) else 0.0) for k in R}
    d_shape = {k: (0.0 if (mount[k] or keep(R[k])) else 1.0) for k in R}
    s, n = p.solve_d(a_r, d_shape)
    a, d = p.fields(a_r, {k: s * v for k, v in d_shape.items()}); psi = p.psi(a, d)
    m = p.export(psi, f'{OUT_DIR}/{name}_{label}.stl')
    out[label] = dict(d_mm=round(s, 3), strut_wall_mm=round(s, 3), field_evals=n, secs=round(time.time() - t0, 1), stl_volume=round(m.volume), V0=round(p.V0), watertight=bool(m.is_watertight))
    print(' ', label, out[label], flush=True)
design_keep('joints_original', lambda r: r['kind'] == 'joint')
design_keep('joints_and_stubby_original', lambda r: r['kind'] == 'joint' or r.get('lam', 99) < 15)
out['fusion_t'] = t_fusion
json.dump(out, open(f'{OUT_DIR}/designs_{name}.json', 'w'), indent=1)
