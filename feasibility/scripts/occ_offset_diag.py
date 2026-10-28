import sys, time
from OCP.BRepOffsetAPI import BRepOffsetAPI_MakeOffsetShape
from OCP.BRepOffset import BRepOffset_Skin
from OCP.GeomAbs import GeomAbs_Arc, GeomAbs_Intersection
from OCP.BRepCheck import BRepCheck_Analyzer
from OCP.Message import Message_ProgressRange
from OCP.TopAbs import TopAbs_SHELL, TopAbs_SOLID, TopAbs_FACE
from OCP.ShapeAnalysis import ShapeAnalysis_FreeBounds
from OCP.BRepAdaptor import BRepAdaptor_Surface
from occ_util import DATA_DIR, OUT_DIR, load_step, volume, faces
from occ_util import count
import collections
name = sys.argv[1]
s = load_step(f'{DATA_DIR}/{name}.step'); V0 = volume(s)
# continuity of bspline faces
cont = collections.Counter()
for f in faces(s):
    a = BRepAdaptor_Surface(f)
    if 'BSpline' in str(a.GetType()):
        bs = a.BSpline(); cont[(str(bs.Continuity()).split('_')[-1], bs.UDegree(), bs.VDegree())] += 1
print(name, 'bspline continuity/degree', dict(cont))
for label, fn in [('join-arc', lambda m, t: m.PerformByJoin(s, t, 1e-3, BRepOffset_Skin, False, False, GeomAbs_Arc, False, Message_ProgressRange())),
                  ('join-int', lambda m, t: m.PerformByJoin(s, t, 1e-3, BRepOffset_Skin, False, False, GeomAbs_Intersection, False, Message_ProgressRange())),
                  ('simple', lambda m, t: m.PerformBySimple(s, t))]:
    for t in [1.0, 1.7]:
        t0 = time.time(); m = BRepOffsetAPI_MakeOffsetShape()
        try:
            fn(m, t); ok = m.IsDone()
            if ok:
                sh = m.Shape(); V = volume(sh)
                print(label, t, 'done', 'valid', BRepCheck_Analyzer(sh).IsValid(), 'solids', count(sh, TopAbs_SOLID), 'shells', count(sh, TopAbs_SHELL), 'faces', count(sh, TopAbs_FACE),
                      'V/V0', round(V / V0, 3), 'expected~', round((V0 + 0) / V0, 3), '+', f'{time.time()-t0:.1f}s', flush=True)
            else:
                print(label, t, 'not done', f'{time.time()-t0:.1f}s', flush=True)
        except Exception as e:
            print(label, t, 'EXC', repr(e)[:120], f'{time.time()-t0:.1f}s', flush=True)
