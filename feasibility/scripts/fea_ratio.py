"""One-shot solid FEA (gmsh C3D10 + CalculiX) -> per-region shear/bending strain-energy ratio R_i.

BCs follow the paper's 'minimum stiffness' case: bolt-hole walls fully fixed; 200 N on the load pad.
"""
import sys, json, time, subprocess, numpy as np, gmsh
from scipy.spatial import cKDTree
from OCP.BRepAdaptor import BRepAdaptor_Surface
from OCP.GProp import GProp_GProps
from OCP.BRepGProp import BRepGProp
from occ_util import DATA_DIR, OUT_DIR, load_step, faces, surf_type

E, NU, FTOT = 1760.0, 0.42, 200.0   # MPa, -, N
G = E / (2 * (1 + NU))
CFG = {  # bolt-hole radius, load-pad selector (planar, +z normal) as a lambda on centroid
    'Solid_L-bracket': dict(r_hole=4.26, pad=lambda c: c[0] > -15),
    'Solid_Cantilever': dict(r_hole=5.0, pad=lambda c: c[0] > 130),
    'Solid_MBB': dict(r_hole=5.0, pad=lambda c: abs(c[0]) < 30),
}
name = sys.argv[1]; hmax = float(sys.argv[2]) if len(sys.argv) > 2 else 2.5
cfg = CFG[name]; path = f'{DATA_DIR}/{name}.step'
shp = load_step(path)
info = []
for f in faces(shp):
    g = GProp_GProps(); BRepGProp.SurfaceProperties_s(f, g); c = g.CentreOfMass()
    t = surf_type(f); a = BRepAdaptor_Surface(f); d = dict(type=t, c=np.array([c.X(), c.Y(), c.Z()]), A=g.Mass())
    if t == 'Cylinder': d['r'] = a.Cylinder().Radius()
    if t == 'Plane':
        n = a.Plane().Axis().Direction(); s = -1 if f.Orientation() == 1 else 1
        d['n'] = s * np.array([n.X(), n.Y(), n.Z()])
    info.append(d)
holes = [i for i, d in enumerate(info) if d['type'] == 'Cylinder' and abs(d['r'] - cfg['r_hole']) < 0.05]
zmax_pad = max(d['c'][2] for d in info if d['type'] == 'Plane' and cfg['pad'](d['c']) and d['n'][2] > 0.99)
pads = [i for i, d in enumerate(info) if d['type'] == 'Plane' and cfg['pad'](d['c']) and d['n'][2] > 0.99 and d['c'][2] > zmax_pad - 0.5]
print(name, 'hole faces', holes, 'pad faces', pads, 'pad area', round(sum(info[i]['A'] for i in pads)))

t0 = time.time()
gmsh.initialize(); gmsh.option.setNumber('General.Terminal', 0)
gmsh.model.occ.importShapes(path); gmsh.model.occ.synchronize()
gfaces = [t for _, t in gmsh.model.getEntities(2)]
gc = np.array([gmsh.model.occ.getCenterOfMass(2, t) for t in gfaces])
tree = cKDTree(gc)
def gtag(i): return gfaces[tree.query(info[i]['c'])[1]]
gmsh.option.setNumber('Mesh.MeshSizeMax', hmax); gmsh.option.setNumber('Mesh.MeshSizeMin', 0.6)
gmsh.option.setNumber('Mesh.MeshSizeFromCurvature', 16); gmsh.option.setNumber('Mesh.ElementOrder', 2)
gmsh.option.setNumber('Mesh.HighOrderOptimize', 0)
gmsh.model.mesh.generate(3)
ntag, X, _ = gmsh.model.mesh.getNodes(); X = X.reshape(-1, 3)
et, etags, enodes = gmsh.model.mesh.getElements(3)
k = list(et).index(11); E10 = enodes[k].reshape(-1, 10).astype(int); EID = etags[k].astype(int)
def face_nodes(i):
    return np.unique(gmsh.model.mesh.getNodes(2, gtag(i), includeBoundary=True)[0]).astype(int)
fix = np.unique(np.concatenate([face_nodes(i) for i in holes])); load = np.unique(np.concatenate([face_nodes(i) for i in pads]))
gmsh.finalize()
mesh_s = time.time() - t0
print(f'  mesh: {len(EID)} C3D10, {len(ntag)} nodes, {3*len(ntag)} DOF, {mesh_s:.1f}s')

# gmsh tet10 node order -> CalculiX C3D10 order (gmsh: 0-3 corners, 4:(0,1) 5:(1,2) 6:(0,2) 7:(0,3) 8:(2,3) 9:(1,3);
# ccx: 5:(1,2) 6:(2,3) 7:(1,3) 8:(1,4) 9:(2,4) 10:(3,4) with 1-based corners) -> swap gmsh 8 and 9
E10c = E10[:, [0, 1, 2, 3, 4, 5, 6, 7, 9, 8]]
job = f'{OUT_DIR}/ccx_{name}'
with open(job + '.inp', 'w') as f:
    f.write('*NODE\n'); [f.write(f'{int(t)},{x:.6f},{y:.6f},{z:.6f}\n') for t, (x, y, z) in zip(ntag, X)]
    f.write('*ELEMENT,TYPE=C3D10,ELSET=EALL\n'); [f.write(f'{e},' + ','.join(map(str, n)) + '\n') for e, n in zip(EID, E10c)]
    f.write('*NSET,NSET=FIX\n'); [f.write(f'{n},\n') for n in fix]
    f.write('*MATERIAL,NAME=PA11\n*ELASTIC\n' + f'{E},{NU}\n*SOLID SECTION,ELSET=EALL,MATERIAL=PA11\n')
    f.write('*BOUNDARY\nFIX,1,3\n*STEP\n*STATIC\n*CLOAD\n'); [f.write(f'{n},3,{-FTOT/len(load):.8f}\n') for n in load]
    f.write('*NODE PRINT,NSET=FIX,TOTALS=ONLY\nRF\n*EL PRINT,ELSET=EALL\nS\n*NODE FILE\nU\n*END STEP\n')
t1 = time.time()
subprocess.run(['ccx', '-i', job], cwd=OUT_DIR, check=True, capture_output=True, env={**__import__('os').environ, 'OMP_NUM_THREADS': '2'})
solve_s = time.time() - t1
# parse integration-point stresses
S = {}
for line in open(job + '.dat'):
    p = line.split()
    if len(p) == 8 and p[0].isdigit():
        S.setdefault(int(p[0]), []).append([float(v) for v in p[2:]])
eid_index = {e: i for i, e in enumerate(EID)}
sig = np.zeros((len(EID), 3, 3))
for e, rows in S.items():
    m = np.mean(rows, 0); sxx, syy, szz, sxy, sxz, syz = m
    sig[eid_index[e]] = [[sxx, sxy, sxz], [sxy, syy, syz], [sxz, syz, szz]]
# element centroid & volume (corner nodes)
pos = {int(t): i for i, t in enumerate(ntag)}
Pc = X[np.vectorize(pos.get)(E10[:, :4])]
cen = Pc.mean(1); vol = np.abs(np.einsum('ij,ij->i', np.cross(Pc[:, 1] - Pc[:, 0], Pc[:, 2] - Pc[:, 0]), Pc[:, 3] - Pc[:, 0])) / 6
# compliance from displacements at load nodes: U = 1/2 F.u = total strain energy
# total strain energy density (isotropic)
s = sig; tr = np.trace(s, axis1=1, axis2=2); ss = np.einsum('nij,nij->n', s, s)
w = (ss * (1 + NU) - NU * tr ** 2) / (2 * E)
# local skeleton tangent at each element
seg = np.load(f'{OUT_DIR}/seg_{name}.npz'); R = {r['id']: r for r in json.load(open(f'{OUT_DIR}/seg_{name}.json'))}
lo, h = seg['lo'], float(seg['h']); sk = np.argwhere(seg['skel']) * h + lo
kt = cKDTree(sk); tang = np.zeros_like(sk)
for i, p in enumerate(sk):
    nb = sk[kt.query_ball_point(p, 3.0)]
    tang[i] = np.linalg.svd(nb - nb.mean(0))[2][0] if len(nb) > 2 else [1, 0, 0]
tv = tang[kt.query(cen)[1]]
tmp = np.where(np.abs(tv[:, 0:1]) < 0.9, np.array([[1, 0, 0]]), np.array([[0, 1, 0]]))
n1 = np.cross(tv, tmp); n1 /= np.linalg.norm(n1, axis=1, keepdims=True); n2 = np.cross(tv, n1)
st = np.einsum('ni,nij->nj', tv, s)
s_tt = np.einsum('nj,nj->n', st, tv); t1_ = np.einsum('nj,nj->n', st, n1); t2_ = np.einsum('nj,nj->n', st, n2)
w_b = s_tt ** 2 / (2 * E); w_s = (t1_ ** 2 + t2_ ** 2) / (2 * G)
lab = seg['label']; ijk = np.clip(np.round((cen - lo) / h).astype(int), 0, np.array(lab.shape) - 1)
el_lab = lab[tuple(ijk.T)]
# elements whose centroid falls just outside the 1 mm voxelisation -> nearest labelled voxel
miss = el_lab == 0
if miss.any():
    lv = np.argwhere(lab > 0); kl = cKDTree(lv * h + lo); el_lab[miss] = lab[tuple(lv[kl.query(cen[miss])[1]].T)]
Utot = float(np.sum(w * vol))
rows = []
for rid, r in R.items():
    sel = el_lab == rid
    if sel.sum() == 0: continue
    Ub, Us, Ur = np.sum(w_b[sel] * vol[sel]), np.sum(w_s[sel] * vol[sel]), np.sum(w[sel] * vol[sel])
    rows.append(dict(id=rid, kind=r['kind'], lam=r.get('lam'), vol=float(vol[sel].sum()), U_share=float(Ur / Utot), R=float(Us / max(Ub, 1e-30)), sed_mean=float(Ur / vol[sel].sum())))
res = dict(name=name, n_el=int(len(EID)), dof=int(3 * len(ntag)), mesh_s=round(mesh_s, 1), solve_s=round(solve_s, 1), U_total_Nmm=Utot, compliance_mm_per_N=2 * Utot / FTOT ** 2, rows=rows)
json.dump(res, open(f'{OUT_DIR}/fea_{name}.json', 'w'), indent=1, default=float)
print(f'  solve {solve_s:.1f}s, U={Utot:.3f} N.mm, stiffness ~ {FTOT**2/(2*Utot)/1000:.4f} kN/mm')
for r in sorted(rows, key=lambda r: -r['U_share'])[:12]:
    lam = f"{r['lam']:.1f}" if r['lam'] else '-'
    print(f"   {r['kind']:6s} id={r['id']:3d} lam={lam:>6s} U_share={r['U_share']:.3f} R=Ushear/Ubend={r['R']:.3f}")
