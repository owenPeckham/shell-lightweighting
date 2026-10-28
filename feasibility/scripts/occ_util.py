"""Small OCC helpers shared by the prototype scripts."""
import numpy as np
from OCP.STEPControl import STEPControl_Reader, STEPControl_Writer, STEPControl_AsIs
from OCP.IFSelect import IFSelect_RetDone
from OCP.TopExp import TopExp_Explorer
from OCP.TopAbs import TopAbs_FACE
from OCP.TopoDS import TopoDS
from OCP.BRepMesh import BRepMesh_IncrementalMesh
from OCP.BRep import BRep_Tool
from OCP.TopLoc import TopLoc_Location
from OCP.TopAbs import TopAbs_REVERSED
from OCP.GProp import GProp_GProps
from OCP.BRepGProp import BRepGProp
from OCP.BRepAdaptor import BRepAdaptor_Surface

import os
RHO = 1.04e-3  # g/mm^3 (PA11)
DATA_DIR = os.environ.get('SHELL_STEP_DIR', os.path.join(os.path.dirname(__file__), '..', 'step'))
OUT_DIR = os.environ.get('SHELL_OUT_DIR', os.path.join(os.path.dirname(__file__), '..', 'out'))
os.makedirs(OUT_DIR, exist_ok=True)


def load_step(path):
    r = STEPControl_Reader()
    assert r.ReadFile(path) == IFSelect_RetDone, path
    r.TransferRoots()
    return r.OneShape()


def save_step(shape, path):
    w = STEPControl_Writer(); w.Transfer(shape, STEPControl_AsIs); w.Write(path)


def faces(shape):
    out, e = [], TopExp_Explorer(shape, TopAbs_FACE)
    while e.More():
        out.append(TopoDS.Face(e.Current())); e.Next()
    return out


def volume(shape):
    g = GProp_GProps(); BRepGProp.VolumeProperties_s(shape, g); return g.Mass()


def face_area(f):
    g = GProp_GProps(); BRepGProp.SurfaceProperties_s(f, g); return g.Mass()


def surf_type(f):
    return str(BRepAdaptor_Surface(f).GetType()).split('.')[-1].replace('GeomAbs_', '')


def tessellate(shape, lin=0.05, ang=0.2, per_face=False):
    """Return vertices (n,3), triangles (m,3) [, face index per triangle]."""
    BRepMesh_IncrementalMesh(shape, lin, False, ang, True)
    V, T, F, off = [], [], [], 0
    for fi, f in enumerate(faces(shape)):
        loc = TopLoc_Location()
        tri = BRep_Tool.Triangulation_s(f, loc)
        if tri is None:
            continue
        trsf = loc.Transformation()
        for i in range(1, tri.NbNodes() + 1):
            p = tri.Node(i).Transformed(trsf)
            V.append((p.X(), p.Y(), p.Z()))
        rev = f.Orientation() == TopAbs_REVERSED
        for i in range(1, tri.NbTriangles() + 1):
            a, b, c = tri.Triangle(i).Get()
            T.append((off + a - 1, off + c - 1, off + b - 1) if rev else (off + a - 1, off + b - 1, off + c - 1))
            F.append(fi)
        off += tri.NbNodes()
    V, T, F = np.array(V), np.array(T), np.array(F)
    return (V, T, F) if per_face else (V, T)


def count(shape, kind):
    from OCP.TopExp import TopExp_Explorer as _E
    n, e = 0, _E(shape, kind)
    while e.More():
        n += 1; e.Next()
    return n
