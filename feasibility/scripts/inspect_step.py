from occ_util import DATA_DIR, OUT_DIR
"""Inspect STEP files: topology counts, surface types, volume/mass, implied wall thickness."""
import sys, json, collections, time
from OCP.STEPControl import STEPControl_Reader
from OCP.IFSelect import IFSelect_RetDone
from OCP.TopExp import TopExp_Explorer
from OCP.TopAbs import TopAbs_SOLID, TopAbs_FACE, TopAbs_SHELL, TopAbs_EDGE
from OCP.TopoDS import TopoDS
from OCP.BRepAdaptor import BRepAdaptor_Surface
from OCP.GProp import GProp_GProps
from OCP.BRepGProp import BRepGProp
from OCP.Bnd import Bnd_Box
from OCP.BRepBndLib import BRepBndLib
from OCP.BRepCheck import BRepCheck_Analyzer
from OCP.GeomAbs import GeomAbs_SurfaceType

RHO = 1.04e-3  # g/mm^3 (PA11, from manuscript)


def load(path):
    r = STEPControl_Reader()
    assert r.ReadFile(path) == IFSelect_RetDone, path
    r.TransferRoots()
    return r.OneShape()


def count(shape, kind):
    n, e = 0, TopExp_Explorer(shape, kind)
    while e.More():
        n += 1; e.Next()
    return n


def surf_types(shape):
    c = collections.Counter()
    e = TopExp_Explorer(shape, TopAbs_FACE)
    while e.More():
        s = BRepAdaptor_Surface(TopoDS.Face(e.Current()))
        c[str(s.GetType()).split('.')[-1].replace('GeomAbs_', '')] += 1
        e.Next()
    return dict(c)


def props(shape):
    g = GProp_GProps(); BRepGProp.VolumeProperties_s(shape, g)
    a = GProp_GProps(); BRepGProp.SurfaceProperties_s(shape, a)
    b = Bnd_Box(); BRepBndLib.Add_s(shape, b)
    p0, p1 = b.CornerMin(), b.CornerMax()
    x0, y0, z0, x1, y1, z1 = p0.X(), p0.Y(), p0.Z(), p1.X(), p1.Y(), p1.Z()
    return g.Mass(), a.Mass(), (x1 - x0, y1 - y0, z1 - z0)


out = {}
for p in sys.argv[1:]:
    t = time.time(); s = load(p); tl = time.time() - t
    V, A, bb = props(s)
    out[p.split('/')[-1]] = dict(
        load_s=round(tl, 2), solids=count(s, TopAbs_SOLID), shells=count(s, TopAbs_SHELL),
        faces=count(s, TopAbs_FACE), edges=count(s, TopAbs_EDGE), surf_types=surf_types(s),
        volume_mm3=round(V, 1), mass_g=round(V * RHO, 2), area_mm2=round(A, 1),
        bbox_mm=[round(v, 1) for v in bb], valid=BRepCheck_Analyzer(s).IsValid())
    print(json.dumps({p.split('/')[-1]: out[p.split('/')[-1]]}), flush=True)
json.dump(out, open(f'{OUT_DIR}/step_summary.json', 'w'), indent=1)
