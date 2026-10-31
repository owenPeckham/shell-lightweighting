"""CalculiX reproduction of the manuscript's Abaqus 3-point bend of the Study 2 PA11 tubes (pp. 14-15).

Quarter model (x >= 0 midspan symmetry, y >= 0 vertical-plane symmetry), x = tube axis, z = vertical.
Structured C3D8I hex mesh: tubes = radial (r, theta, x) sweep; solid rod A = butterfly (square core + O-ring).
bc = "contact": rigid rollers (support R 10 mm fixed, load R 25 mm displacement-controlled), meshed as one layer of
     C3D8 whose nodes are all prescribed (= rigid), frictionless surface-to-surface penalty contact.
bc = "patch":   cshell voxel BCs (tubes_pw.py): 4 mm wide load patch (z >= zmax - 0.5) of uniform vertical
     pressure, support = the node plane x = 144 with z <= zmin + 0.5, uz = 0.
Usage: python tube_ccx.py SIZE LEVEL BC [--K 1e5] [--nlgeom] [--d 3.0] [--np 72] [--tag x]
Writes <outdir>/<SIZE>_L<LEVEL>_<BC>[_tag]/{job.inp, job.dat, result.json}."""
import argparse
import json
import os
import re
import subprocess
import time
from pathlib import Path

import numpy as np

E, NU = 1760.0, 0.42
SIZES = {"A": (15.0, 0.0), "B": (15.8, 5.0), "C": (18.0, 10.0), "D": (21.2, 15.0), "E": (25.0, 20.0)}
SPAN, LENGTH = 288.0, 345.6
XS, XE = SPAN / 2, LENGTH / 2
R_SUP, R_LOAD = 10.0, 25.0
# mesh levels: radial layers, half-circumference elements, axial size at contacts, max axial size
# (L1 as run in job 7800452 used w = 6; L1/L2 below: fine zone +-w mm around each contact; nt_A = rod A override)
LEVELS = {1: dict(nr=4, nt=64, hc=0.5, hmax=3.0, w=6.0, nt_A=48), 2: dict(nr=6, nt=96, hc=0.25, hmax=3.0, w=3.0, nt_A=64),
          3: dict(nr=8, nt=128, hc=0.25, hmax=1.5, w=3.0, nt_A=96),
          # coarse uniform circumferential meshes (faceting study, patch BCs): 2*nt elements round the full circle
          **{100 + n: dict(nr=4, nt=n, hc=0.5, hmax=3.0, w=6.0, nt_A=n, a=0.0) for n in (8, 10, 12, 16, 24, 32)},
          15: dict(nr=6, nt=96, hc=0.333333333333, hmax=2.0)}
OUT = Path(__file__).resolve().parent / "runs"
FACES = np.array([[0, 1, 2, 3], [4, 7, 6, 5], [0, 4, 5, 1], [1, 5, 6, 2], [2, 6, 7, 3], [3, 7, 4, 0]])


def graded(a, b, h0, h1, hmax, g=0.12):
    """Nodes in [a, b], size h0 at a, h1 at b, growing linearly at rate g up to hmax; endpoints exact."""
    xs = [a]
    while True:
        x = xs[-1]
        h = min(hmax, h0 + g * (x - a), h1 + g * (b - x))
        if x + h >= b - 0.5 * h:
            break
        xs.append(x + h)
    xs = np.array(xs + [b])
    return a + (xs - a) * (b - a) / (xs[-1] - a)


def axial_nodes(hc, hmax, w=6.0):
    n = int(round(w / hc))
    segs = [np.linspace(0, w, n + 1), graded(w, XS - w, hc, hc, hmax), np.linspace(XS - w, XS + w, 2 * n + 1),
            graded(XS + w, XE, hc, hmax, hmax)]
    x = np.concatenate([segs[0]] + [s[1:] for s in segs[1:]])
    assert np.all(np.diff(x) > 0) and np.isclose(x, 2.0).any() and np.isclose(x, XS).any()
    return x


def theta_map(s, a):
    """s in [0,1] -> angle from bottom (0) to top (pi); a = refinement towards top/bottom (ratio (1+a)/(1-a))."""
    return np.pi * (s - a * np.sin(2 * np.pi * s) / (2 * np.pi))


def section(size, lv):
    """2D quad mesh of the half section (y >= 0): nodes (n,2) [y,z], quads (m,4)."""
    od, idd = SIZES[size]
    ro, ri = od / 2, idd / 2
    nr, nt = lv["nr"], lv["nt"]
    if ri > 0:
        r = np.linspace(ri, ro, nr + 1)
        th = theta_map(np.linspace(0, 1, nt + 1), lv.get("a", 0.5))
        R, T = np.meshgrid(r, th, indexing="ij")
        P = np.stack([R * np.sin(T), -R * np.cos(T)], -1).reshape(-1, 2)
        idx = np.arange(P.shape[0]).reshape(nr + 1, nt + 1)
        Q = np.stack([idx[:-1, :-1], idx[1:, :-1], idx[1:, 1:], idx[:-1, 1:]], -1).reshape(-1, 4)
        return P, Q
    # butterfly: core rectangle y in [0,c], z in [-c,c]; ring from core boundary to circle; uniform angular param
    c = 0.5 * ro
    nt = lv["nt_A"]
    q = nt // 4
    sg = np.linspace(0, 1, nt + 1)    # graded 3:1 per core side (fine at top/bottom and core corners), quarter points fixed
    t = sg - lv.get("a", 0.5) * np.sin(4 * np.pi * sg) / (4 * np.pi)
    yc, zc = 4 * c * t[:q + 1], -c + 4 * c * (t[q:3 * q + 1] - 0.25)
    Yc, Zc = np.meshgrid(yc, zc, indexing="ij")
    Pc = np.stack([Yc, Zc], -1).reshape(-1, 2)
    ic = np.arange(Pc.shape[0]).reshape(q + 1, 2 * q + 1)
    Qc = np.stack([ic[:-1, :-1], ic[1:, :-1], ic[1:, 1:], ic[:-1, 1:]], -1).reshape(-1, 4)
    # core boundary path from bottom (0,-c) -> (c,-c) -> (c,c) -> (0,c), nt segments
    L = 4 * c * t
    by = np.where(L <= c, L, np.where(L <= 3 * c, c, 4 * c - L))
    bz = np.where(L <= c, -c, np.where(L <= 3 * c, -c + (L - c), c))
    th = np.pi * t
    cy, cz = ro * np.sin(th), -ro * np.cos(th)
    s = np.linspace(0, 1, nr + 1)[:, None]
    Pr = np.stack([by + s * (cy - by), bz + s * (cz - bz)], -1).reshape(-1, 2)
    ir = np.arange(Pr.shape[0]).reshape(nr + 1, nt + 1) + Pc.shape[0]
    Qr = np.stack([ir[:-1, :-1], ir[1:, :-1], ir[1:, 1:], ir[:-1, 1:]], -1).reshape(-1, 4)
    P = np.vstack([Pc, Pr]); Q = np.vstack([Qc, Qr])
    key = np.round(P, 7)
    _, uniq, inv = np.unique(key, axis=0, return_index=True, return_inverse=True)
    inv = inv.ravel()
    return P[uniq], inv[Q]


def sweep(P, Q, xs):
    """Extrude the 2D section along x -> nodes (N,3), hexes (M,8) with positive Jacobian."""
    n2 = P.shape[0]
    X = np.concatenate([np.column_stack([np.full(n2, x), P]) for x in xs])
    H = np.concatenate([np.hstack([Q + k * n2, Q + (k + 1) * n2]) for k in range(len(xs) - 1)])
    H = orient(X, H)
    return X, H


def orient(X, H):
    c = X[H]
    d = np.einsum("ij,ij->i", np.cross(c[:, 1] - c[:, 0], c[:, 3] - c[:, 0]), c[:, 4] - c[:, 0])
    H = H.copy(); neg = d < 0
    H[neg] = H[neg][:, [0, 3, 2, 1, 4, 7, 6, 5]]
    return H


def outer_faces(X, H, ro, sel):
    """(elem idx, face no 1..6) with all 4 face nodes on r = ro and sel(face centroid) true."""
    r = np.hypot(X[:, 1], X[:, 2])
    out = []
    for f, fn in enumerate(FACES):
        nod = H[:, fn]
        on = np.all(np.abs(r[nod] - ro) < 1e-6, axis=1)
        cen = X[nod].mean(1)
        e = np.nonzero(on & sel(cen))[0]
        out += [(int(i), f + 1) for i in e]
    return out


def roller(centre_x, centre_z, R, psi0, psi1, ylen, hs, side):
    """One-layer C3D8 cylinder sector, axis along y; side=+1 support (contact face on top), -1 load (bottom).
    Returns nodes, hexes, and the face number of the contact face (r = R)."""
    npsi = max(4, int(np.ceil(R * (psi1 - psi0) / hs)))
    ny = max(4, int(np.ceil(ylen / 0.5)))
    psi = np.linspace(psi0, psi1, npsi + 1)
    y = np.linspace(-0.5, ylen, ny + 1)
    rr = np.array([R - 1.0, R])
    RR, PS, YY = np.meshgrid(rr, psi, y, indexing="ij")
    X = np.stack([centre_x + RR * np.sin(PS), YY, centre_z + side * RR * np.cos(PS)], -1).reshape(-1, 3)
    idx = np.arange(X.shape[0]).reshape(2, npsi + 1, ny + 1)
    H = np.stack([idx[0, :-1, :-1], idx[1, :-1, :-1], idx[1, 1:, :-1], idx[0, 1:, :-1],
                  idx[0, :-1, 1:], idx[1, :-1, 1:], idx[1, 1:, 1:], idx[0, 1:, 1:]], -1).reshape(-1, 8)
    H = orient(X, H)
    dist = np.hypot(X[:, 0] - centre_x, X[:, 2] - centre_z)
    faces = []
    for f, fn in enumerate(FACES):
        on = np.all(np.abs(dist[H[:, fn]] - R) < 1e-6, axis=1)
        faces += [(int(i), f + 1) for i in np.nonzero(on)[0]]
    return X, H, faces


def ids(a, per=16):
    a = list(a)
    return "".join(", ".join(str(int(v)) for v in a[i:i + per]) + ",\n" for i in range(0, len(a), per))


def build(size, level, bc, K, nlgeom, d, eltype="C3D8I", pw=4.0):
    lv = LEVELS[level]
    od, idd = SIZES[size]
    ro = od / 2
    xs = axial_nodes(lv["hc"], lv["hmax"], lv["w"])
    P, Q = section(size, lv)
    X, H = sweep(P, Q, xs)
    nT = X.shape[0]
    L = []
    w = L.append
    w("*HEADING\nPA11 tube %s 3-point bend quarter model, %s, level %d\n" % (size, bc, level))
    allX = [X]; blocks = []
    info = dict(size=size, level=level, bc=bc, nodes_tube=int(nT), elems_tube=int(H.shape[0]), eltype=eltype,
                nr=lv["nr"], nt=lv["nt"], hc=lv["hc"], hmax=lv["hmax"], nx=int(len(xs) - 1))
    tol = 1e-7
    sets = {}
    if bc == "contact":
        Xs, Hs, fs = roller(XS, -ro - R_SUP, R_SUP, -0.45, 0.45, ro + 2.0, 0.25 if level == 1 else 0.15, +1)
        Xl, Hl, fl = roller(0.0, ro + R_LOAD, R_LOAD, -0.04, 0.30, ro + 2.0, 0.25 if level == 1 else 0.15, -1)
        o1 = nT; o2 = nT + Xs.shape[0]; o3 = o2 + Xl.shape[0]
        allX += [Xs, Xl, np.array([[0.0, 0.0, -ro - 1.0]])]
        ground = o3
        sets["SUPN"] = np.arange(o1, o2); sets["LOADN"] = np.arange(o2, o3)
        blocks = [("SUPR", Hs + o1), ("LOADR", Hl + o2)]
    Xall = np.vstack(allX)
    w("*NODE, NSET=NALL\n")
    w("".join("%d, %.9f, %.9f, %.9f\n" % (i + 1, *p) for i, p in enumerate(Xall)))
    w("*ELEMENT, TYPE=%s, ELSET=TUBE\n" % eltype)
    w("".join("%d, %s\n" % (i + 1, ", ".join(str(v + 1) for v in h)) for i, h in enumerate(H)))
    eoff = H.shape[0]
    roller_faces = {}
    for name, Hb in blocks:
        w("*ELEMENT, TYPE=C3D8, ELSET=%s\n" % name)
        w("".join("%d, %s\n" % (eoff + i + 1, ", ".join(str(v + 1) for v in h)) for i, h in enumerate(Hb)))
        roller_faces[name] = eoff
        eoff += Hb.shape[0]
    w("*MATERIAL, NAME=PA11\n*ELASTIC\n%g, %g\n*SOLID SECTION, ELSET=TUBE, MATERIAL=PA11\n" % (E, NU))
    # symmetry
    w("*NSET, NSET=SYMX\n" + ids(np.nonzero(np.abs(X[:, 0]) < tol)[0] + 1))
    w("*NSET, NSET=SYMY\n" + ids(np.nonzero(np.abs(X[:, 1]) < tol)[0] + 1))
    bnd = ["SYMX, 1, 1", "SYMY, 2, 2"]
    zmin, zmax = X[:, 2].min(), X[:, 2].max()
    # monitoring nodes: midspan top/bottom on y = 0, support-plane top/bottom
    def nid(x, z):
        return int(np.argmin((X[:, 0] - x) ** 2 + X[:, 1] ** 2 + (X[:, 2] - z) ** 2)) + 1
    mon = dict(mid_top=nid(0, ro), mid_bot=nid(0, -ro), sup_top=nid(XS, ro), sup_bot=nid(XS, -ro),
               mid_side=int(np.argmin(X[:, 0] ** 2 + (X[:, 1] - ro) ** 2 + X[:, 2] ** 2)) + 1)
    w("*NSET, NSET=MON\n" + ids(sorted(set(mon.values()))))
    info["monitor"] = mon
    if bc == "contact":
        w("*MATERIAL, NAME=RIG\n*ELASTIC\n2.1e5, 0.3\n")
        for name, _ in blocks:
            w("*SOLID SECTION, ELSET=%s, MATERIAL=RIG\n" % name)
        w("*NSET, NSET=SUPN\n" + ids(sets["SUPN"] + 1))
        w("*NSET, NSET=LOADN\n" + ids(sets["LOADN"] + 1))
        w("*NSET, NSET=GND\n%d\n" % (ground + 1))
        # weak z spring (1e-3 N/mm) removes the tube's rigid z mode before contact closes
        w("*ELEMENT, TYPE=SPRINGA, ELSET=ESPR\n%d, %d, %d\n" % (eoff + 1, mon["mid_bot"], ground + 1))
        w("*SPRING, ELSET=ESPR\n\n1.e-3\n")
        slave_l = outer_faces(X, H, ro, lambda c: (c[:, 0] < 10.0) & (c[:, 2] > 0.6 * ro))
        slave_s = outer_faces(X, H, ro, lambda c: (np.abs(c[:, 0] - XS) < 10.0) & (c[:, 2] < -0.6 * ro))
        for nm, fl_ in (("SL_LOAD", slave_l), ("SL_SUP", slave_s)):
            w("*SURFACE, NAME=%s, TYPE=ELEMENT\n" % nm + "".join("%d, S%d\n" % (e + 1, f) for e, f in fl_))
        for nm, rf in (("M_SUP", fs), ("M_LOAD", fl)):
            off = roller_faces["SUPR" if nm == "M_SUP" else "LOADR"]
            w("*SURFACE, NAME=%s, TYPE=ELEMENT\n" % nm + "".join("%d, S%d\n" % (off + e + 1, f) for e, f in rf))
        w("*SURFACE INTERACTION, NAME=SI\n*SURFACE BEHAVIOR, PRESSURE-OVERCLOSURE=LINEAR\n%g, 1.e-6\n" % K)
        w("*CONTACT PAIR, INTERACTION=SI, TYPE=SURFACE TO SURFACE\nSL_LOAD, M_LOAD\n")
        w("*CONTACT PAIR, INTERACTION=SI, TYPE=SURFACE TO SURFACE\nSL_SUP, M_SUP\n")
        bnd += ["SUPN, 1, 3", "LOADN, 1, 2", "GND, 1, 3"]
        nsl = sorted({int(v) for e, f in slave_l for v in H[e, FACES[f - 1]]})
        nss = sorted({int(v) for e, f in slave_s for v in H[e, FACES[f - 1]]})
        w("*NSET, NSET=SLN\n" + ids(np.array(nsl) + 1))
        w("*NSET, NSET=SSN\n" + ids(np.array(nss) + 1))
        info.update(K=K, n_slave_load=len(slave_l), n_slave_sup=len(slave_s), d=d)
    else:
        sup = np.nonzero((np.abs(X[:, 0] - XS) < tol) & (X[:, 2] <= zmin + 0.5 + 1e-9))[0]
        w("*NSET, NSET=SUPP\n" + ids(sup + 1))
        bnd += ["SUPP, 3, 3"]
        top = outer_faces(X, H, ro, lambda c: np.ones(len(c), bool))
        sel = []
        Aproj = 0.0
        for e, f in top:
            nod = H[e, FACES[f - 1]]
            c = X[nod]
            if np.all(c[:, 0] <= pw / 2 + 1e-9) and np.all(c[:, 2] >= zmax - 0.5 - 1e-9):
                # projected (x-y) area of the bilinear quad
                v1, v2 = c[2, :2] - c[0, :2], c[3, :2] - c[1, :2]; a = 0.5 * abs(v1[0] * v2[1] - v1[1] * v2[0])
                Aproj += a; sel.append((e, f))
        p = 25.0 / Aproj
        info.update(pw=pw, patch_faces=len(sel), patch_Aproj_quarter=Aproj, support_nodes=int(len(sup)), pressure=p)
        w("*NSET, NSET=PATCHN\n" + ids(sorted({int(v) + 1 for e, f in sel for v in H[e, FACES[f - 1]]})))
    w("*BOUNDARY\n" + "\n".join(bnd) + "\n")
    w("*TIME POINTS, NAME=TP\n0.03, 0.1, 0.3, 1.0\n")
    w("*STEP%s, INC=1000\n" % (", NLGEOM" if nlgeom else ""))
    if bc == "contact":
        w("*STATIC, SOLVER=SPOOLES\n0.03, 1.0, 1.e-6, 0.25\n")
        w("*BOUNDARY\nLOADN, 3, 3, %g\n" % (-d))
        w("*NODE PRINT, NSET=LOADN, TOTALS=ONLY, TIME POINTS=TP\nRF\n")
        w("*NODE PRINT, NSET=SUPN, TOTALS=ONLY, TIME POINTS=TP\nRF\n")
        w("*NODE PRINT, NSET=MON, TIME POINTS=TP\nU\n")
        w("*NODE PRINT, NSET=SLN, TIME POINTS=TP\nU\n")
        w("*NODE PRINT, NSET=SSN, TIME POINTS=TP\nU\n")
        w("*EL PRINT, ELSET=TUBE, TOTALS=ONLY, TIME POINTS=TP\nELSE\n")
    else:
        w("*STATIC, SOLVER=SPOOLES\n")
        w("*DLOAD\n" + "".join("%d, P%d, %.10g\n" % (e + 1, f, p) for e, f in sel))
        w("*NODE PRINT, NSET=MON\nU\n*NODE PRINT, NSET=PATCHN\nU\n")
        w("*NODE PRINT, NSET=SUPP, TOTALS=ONLY\nRF\n")
        w("*EL PRINT, ELSET=TUBE, TOTALS=ONLY\nELSE\n")
    w("*END STEP\n")
    meta = dict(X=Xall, H=H, ro=ro, sel=sel if bc == "patch" else None)
    return "".join(L), info, meta


def parse_blocks(path):
    """-> list of (header, time, rows) from a ccx .dat file."""
    out = []; cur = None
    hdr = re.compile(r"^\s*(.+?)\s+(?:for set|for all)\s+(\S+).*time\s+([-+0-9.Ee]+)")
    for ln in open(path):
        m = hdr.match(ln)
        if m:
            cur = [m.group(1).strip() + " " + m.group(2), float(m.group(3)), []]; out.append(cur); continue
        s = ln.split()
        if cur is not None and s:
            try:
                cur[2].append([float(v) for v in s])
            except ValueError:
                cur = None
    return out


def analyse(info, meta, dat):
    B = parse_blocks(dat)
    get = lambda key: [(t, np.array(r)) for h, t, r in B if h.startswith(key)]
    res = dict(info)
    mon = info["monitor"]
    if info["bc"] == "contact":
        rfl = get("total force (fx,fy,fz) LOADN")
        rfs = get("total force (fx,fy,fz) SUPN")
        ene = get("total internal energy ELSET=TUBE") or get("total internal energy TUBE")
        uS = get("displacements (vx,vy,vz) SLN")
        uM = get("displacements (vx,vy,vz) MON")
        X, ro = meta["X"], meta["ro"]
        rows = []
        for i, (t, r) in enumerate(rfl):
            F = -4 * r[0][2]          # full-model load (N), quarter reaction x4
            dl = info["d"] * t
            um = {int(a[0]): a[1:4] for a in uM[i][1]}
            # contact footprint on the load roller: slave nodes with overclosure > 0 against the deformed roller
            us = uS[i][1]; nid_ = us[:, 0].astype(int) - 1
            p = X[nid_] + us[:, 1:4]
            cz = ro + R_LOAD - dl
            over = R_LOAD - np.hypot(p[:, 0], p[:, 2] - cz)
            c = over > 0
            row = dict(t=t, delta=dl, F=F, k=F / dl / 1000.0, F_sup=4 * r[0][2] if False else 4 * rfs[i][1][0][2],
                       ovalisation_mid=float(um[mon["mid_top"]][2] - um[mon["mid_bot"]][2]),
                       ovalisation_sup=float(um[mon["sup_top"]][2] - um[mon["sup_bot"]][2]),
                       uz_mid_bot=float(um[mon["mid_bot"]][2]), uz_mid_top=float(um[mon["mid_top"]][2]),
                       contact_half_length_x=float(X[nid_][c, 0].max()) if c.any() else 0.0,
                       contact_half_width_y=float(X[nid_][c, 1].max()) if c.any() else 0.0,
                       max_overclosure=float(over.max()), n_contact=int(c.sum()))
            if ene:
                row["energy"] = 4 * ene[i][1][0][0]
            rows.append(row)
        res["curve"] = rows
        r1 = rows[-1]
        res.update(k_100N_interp=float(np.interp(100.0, [r["F"] for r in rows], [r["k"] for r in rows])))
    else:
        ene = get("total internal energy ELSET=TUBE") or get("total internal energy TUBE")
        W = 4 * ene[-1][1][0][0]
        F = 100.0
        um = {int(a[0]): a[1:4] for a in get("displacements (vx,vy,vz) MON")[-1][1]}
        res.update(energy=W, k=F ** 2 / (2 * W) / 1000.0,
                   ovalisation_mid=float(um[mon["mid_top"]][2] - um[mon["mid_bot"]][2]),
                   uz_mid_bot=float(um[mon["mid_bot"]][2]), uz_mid_top=float(um[mon["mid_top"]][2]),
                   k_bot=F / -um[mon["mid_bot"]][2] / 1000.0,
                   F_sup=4 * get("total force (fx,fy,fz) SUPP")[-1][1][0][2])
    return res


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("size"); ap.add_argument("level", type=int); ap.add_argument("bc")
    ap.add_argument("--K", type=float, default=1e5); ap.add_argument("--nlgeom", action="store_true")
    ap.add_argument("--d", type=float, default=None); ap.add_argument("--np", type=int, default=72)
    ap.add_argument("--tag", default=""); ap.add_argument("--pw", type=float, default=4.0); ap.add_argument("--dry", action="store_true")
    a = ap.parse_args()
    d = a.d or {"A": 12.5, "B": 10.5, "C": 6.9, "D": 4.4, "E": 3.0}[a.size]
    deck, info, meta = build(a.size, a.level, a.bc, a.K, a.nlgeom, d, pw=a.pw)
    info["nlgeom"] = a.nlgeom
    name = f"{a.size}_L{a.level}_{a.bc}" + (f"_{a.tag}" if a.tag else "")
    wd = OUT / name; wd.mkdir(parents=True, exist_ok=True)
    (wd / "job.inp").write_text(deck)
    print(name, json.dumps({k: v for k, v in info.items() if k != "monitor"}), flush=True)
    if a.dry:
        return
    env = dict(os.environ, OMP_NUM_THREADS=str(a.np), CCX_NPROC_EQUATION_SOLVER=str(a.np),
               CCX_NPROC_STIFFNESS_ASSEMBLY=str(a.np), CCX_NPROC_RESULTS=str(a.np), OMP_STACKSIZE="256M")
    exe = str(Path(os.environ["CSHELL_ENV"]) / "bin" / "ccx")
    t0 = time.time()
    with open(wd / "job.log", "w") as lg:
        rc = subprocess.run([exe, "-i", "job"], cwd=wd, stdout=lg, stderr=subprocess.STDOUT, env=env).returncode
    info["seconds"] = time.time() - t0; info["rc"] = rc
    try:
        res = analyse(info, meta, wd / "job.dat")
    except Exception as e:  # keep the run record even if parsing fails
        res = dict(info, error=repr(e))
    (wd / "result.json").write_text(json.dumps(res, indent=1, default=float))
    short = {k: res.get(k) for k in ("k", "k_100N_interp", "seconds", "rc", "error", "ovalisation_mid")}
    if "curve" in res:
        short["curve"] = [(round(r["F"], 2), round(r["k"], 5), round(r["contact_half_length_x"], 2),
                           round(r["contact_half_width_y"], 2), round(r["ovalisation_mid"], 4)) for r in res["curve"]]
    print(name, json.dumps(short), flush=True)


if __name__ == "__main__":
    main()
