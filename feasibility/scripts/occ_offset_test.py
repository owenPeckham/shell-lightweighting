"""Test OCC B-rep offset capabilities on the GD solids.

a) uniform outward offset (what Fusion's 'outside shell' does)  -> volume added vs t
b) per-face offset values via BRepOffset_MakeOffset.SetOffsetOnFace
c) thick-solid (hollow) via BRepOffsetAPI_MakeThickSolid with no removed faces
"""
import sys, time, json
import numpy as np
from OCP.BRepOffsetAPI import BRepOffsetAPI_MakeOffsetShape, BRepOffsetAPI_MakeThickSolid
from OCP.BRepOffset import BRepOffset_MakeOffset, BRepOffset_Skin
from OCP.GeomAbs import GeomAbs_Arc, GeomAbs_Intersection
from OCP.BRepCheck import BRepCheck_Analyzer
from OCP.collections import List_TopoDS_Shape as TopTools_ListOfShape
from OCP.BRepAlgoAPI import BRepAlgoAPI_Cut
from OCP.Message import Message_ProgressRange
from occ_util import DATA_DIR, OUT_DIR, load_step, volume, faces, face_area, surf_type

name = sys.argv[1]
s = load_step(f'{DATA_DIR}/{name}.step')
V0 = volume(s)
res = {'name': name, 'V0': V0, 'uniform': [], 'perface': None, 'thick': None}


def try_offset(t):
    t0 = time.time()
    try:
        mk = BRepOffsetAPI_MakeOffsetShape()
        mk.PerformByJoin(s, t, 1e-3, BRepOffset_Skin, False, False, GeomAbs_Arc, False, Message_ProgressRange())
        ok = mk.IsDone()
        sh = mk.Shape() if ok else None
        valid = ok and BRepCheck_Analyzer(sh).IsValid()
        V = volume(sh) if ok else float('nan')
    except Exception as e:  # OCC raises Standard_Failure on many offset errors
        ok, valid, V = False, False, float('nan')
    return dict(t=t, done=bool(ok), valid=bool(valid), dV=V - V0, secs=round(time.time() - t0, 1))


for t in [0.5, 1.0, 1.5, 2.0]:
    r = try_offset(t); res['uniform'].append(r); print('uniform', r, flush=True)

# b) per-face offsets: faces whose centroid lies in the first/last 25% along the longest bbox axis
#    get 0.5 mm, the rest 1.5 mm (a crude 'joint vs strut' split).
from OCP.GProp import GProp_GProps
from OCP.BRepGProp import BRepGProp
fl = faces(s)
cent = []
for f in fl:
    g = GProp_GProps(); BRepGProp.SurfaceProperties_s(f, g); c = g.CentreOfMass(); cent.append((c.X(), c.Y(), c.Z()))
cent = np.array(cent); ax = np.argmax(np.ptp(cent, 0)); lo, hi = np.percentile(cent[:, ax], [25, 75])
t0 = time.time()
try:
    mo = BRepOffset_MakeOffset()
    mo.Initialize(s, 1.5, 1e-3, BRepOffset_Skin, False, False, GeomAbs_Arc, False, False)
    n_thin = 0
    for f, c in zip(fl, cent):
        if c[ax] < lo or c[ax] > hi:
            mo.SetOffsetOnFace(f, 0.5); n_thin += 1
    mo.MakeOffsetShape(Message_ProgressRange())
    ok = mo.IsDone(); sh = mo.Shape() if ok else None
    res['perface'] = dict(done=bool(ok), valid=bool(ok and BRepCheck_Analyzer(sh).IsValid()),
                          n_faces=len(fl), n_thin=n_thin, err=str(mo.Error()), secs=round(time.time() - t0, 1),
                          dV=(volume(sh) - V0) if ok else None)
except Exception as e:
    res['perface'] = dict(done=False, exc=repr(e)[:200], secs=round(time.time() - t0, 1))
print('perface', res['perface'], flush=True)

# c) closed hollow via thick solid (inside offset -1.0 mm, no faces removed)
t0 = time.time()
try:
    mt = BRepOffsetAPI_MakeThickSolid()
    mt.MakeThickSolidByJoin(s, TopTools_ListOfShape(), -1.0, 1e-3, BRepOffset_Skin, False, False, GeomAbs_Arc, False, Message_ProgressRange())
    ok = mt.IsDone(); sh = mt.Shape() if ok else None
    res['thick'] = dict(done=bool(ok), valid=bool(ok and BRepCheck_Analyzer(sh).IsValid()), V=(volume(sh) if ok else None), secs=round(time.time() - t0, 1))
except Exception as e:
    res['thick'] = dict(done=False, exc=repr(e)[:200], secs=round(time.time() - t0, 1))
print('thick', res['thick'], flush=True)
json.dump(res, open(f'{OUT_DIR}/offset_{name}.json', 'w'), indent=1, default=float)
