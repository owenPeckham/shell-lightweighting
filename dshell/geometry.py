"""
STEP import (OpenCascade via OCP), face table, face selection and tessellation.
"""
from __future__ import annotations

import numpy as np

from OCP.STEPControl import STEPControl_Reader
from OCP.IFSelect import IFSelect_RetDone
from OCP.TopExp import TopExp_Explorer
from OCP.TopAbs import TopAbs_FACE, TopAbs_SOLID, TopAbs_REVERSED
from OCP.TopoDS import TopoDS
from OCP.BRepMesh import BRepMesh_IncrementalMesh
from OCP.BRep import BRep_Tool
from OCP.TopLoc import TopLoc_Location
from OCP.GProp import GProp_GProps
from OCP.BRepGProp import BRepGProp
from OCP.BRepAdaptor import BRepAdaptor_Surface
from OCP.BRepCheck import BRepCheck_Analyzer

ANALYTIC = ("Plane", "Cylinder", "Cone", "Sphere", "Torus")


def load_step(path: str):
    r = STEPControl_Reader()
    if r.ReadFile(path) != IFSelect_RetDone:
        raise IOError(f"could not read STEP file: {path}")
    r.TransferRoots()
    shape = r.OneShape()
    n_solids, e = 0, TopExp_Explorer(shape, TopAbs_SOLID)
    while e.More():
        n_solids += 1
        e.Next()
    if n_solids == 0:
        raise ValueError(f"{path}: no solid body found")
    return shape


def faces(shape) -> list:
    out, e = [], TopExp_Explorer(shape, TopAbs_FACE)
    while e.More():
        out.append(TopoDS.Face(e.Current()))
        e.Next()
    return out


def volume(shape) -> float:
    g = GProp_GProps()
    BRepGProp.VolumeProperties_s(shape, g)
    return g.Mass()


def area(shape) -> float:
    g = GProp_GProps()
    BRepGProp.SurfaceProperties_s(shape, g)
    return g.Mass()


def is_valid(shape) -> bool:
    return bool(BRepCheck_Analyzer(shape).IsValid())


def face_table(shape) -> list[dict]:
    """One dict per face: id, type, area, centroid; plane: outward normal; cylinder: radius, axis."""
    rows = []
    for i, f in enumerate(faces(shape)):
        g = GProp_GProps()
        BRepGProp.SurfaceProperties_s(f, g)
        c = g.CentreOfMass()
        a = BRepAdaptor_Surface(f)
        t = str(a.GetType()).split(".")[-1].replace("GeomAbs_", "")
        row = dict(id=i, type=t, area=float(g.Mass()), centroid=[c.X(), c.Y(), c.Z()])
        if t == "Plane":
            n = a.Plane().Axis().Direction()
            s = -1.0 if f.Orientation() == TopAbs_REVERSED else 1.0
            row["normal"] = [s * n.X(), s * n.Y(), s * n.Z()]
        elif t == "Cylinder":
            cy = a.Cylinder()
            d = cy.Axis().Direction()
            p = cy.Axis().Location()
            row["radius"] = float(cy.Radius())
            row["axis"] = [d.X(), d.Y(), d.Z()]
            row["axis_point"] = [p.X(), p.Y(), p.Z()]
        rows.append(row)
    return rows


def select_faces(table: list[dict], sel: dict) -> list[int]:
    """Face ids matching one selector (see config.py)."""
    t = sel.get("type", "ids")
    if t == "ids":
        ids = [int(i) for i in sel["ids"]]
        bad = [i for i in ids if i < 0 or i >= len(table)]
        if bad:
            raise ValueError(f"face ids out of range: {bad}")
        return ids
    if t == "cylinder":
        tol = float(sel.get("tol", 0.05))
        out = [r["id"] for r in table if r["type"] == "Cylinder" and abs(r["radius"] - float(sel["radius"])) <= tol]
    elif t == "plane":
        n = np.asarray(sel["normal"], float)
        n /= np.linalg.norm(n)
        cosmin = np.cos(np.radians(float(sel.get("tol_deg", 2.0))))
        out = [r["id"] for r in table if r["type"] == "Plane" and np.dot(r["normal"], n) >= cosmin]
    else:
        raise ValueError(f"unknown face selector type: {t}")
    box = sel.get("box") or {}
    for ax, key in enumerate("xyz"):
        if key in box:
            lo, hi = box[key]
            out = [i for i in out if lo <= table[i]["centroid"][ax] <= hi]
    if t == "plane" and sel.get("extreme") and out:
        n = np.asarray(sel["normal"], float) / np.linalg.norm(sel["normal"])
        pos = np.array([np.dot(table[i]["centroid"], n) for i in out])
        out = [i for i, p in zip(out, pos) if p >= pos.max() - float(sel.get("extreme_tol", 0.5))]
    return out


def tessellate(shape, lin: float = 0.02, ang: float = 0.2):
    """Watertight triangulation: vertices (n, 3), triangles (m, 3), face id per triangle (m,)."""
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
    V, T, F = np.asarray(V, float), np.asarray(T, np.int64), np.asarray(F, np.int64)
    # merge coincident vertices shared by neighbouring faces (needed for a watertight mesh)
    key = np.round(V / 1e-6).astype(np.int64)
    _, first, inv = np.unique(key, axis=0, return_index=True, return_inverse=True)
    V = V[first]
    T = inv.ravel()[T]
    keep = (T[:, 0] != T[:, 1]) & (T[:, 1] != T[:, 2]) & (T[:, 0] != T[:, 2])
    return V, T[keep], F[keep]


def submesh(V, T, F, face_ids):
    sel = np.isin(F, list(face_ids))
    return V, T[sel]
