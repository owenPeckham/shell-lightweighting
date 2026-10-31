"""Decimated preview STLs of a full cshell run (final_design, original, uniform_shell).

Usage (CPU job, see make_previews.sh):
    python examples/cshell/runs/make_previews.py <full_run_dir> <out_dir> [--faces 500000] [--max-dvol 0.005]

Quadric edge collapse (pymeshlab, topology and boundary preserving) to ~``--faces`` triangles; writes binary
``<name>_preview.stl`` and ``previews.json`` (triangle counts, volumes, relative volume change, watertightness,
for the full and the preview mesh).  If the volume change exceeds ``--max-dvol`` the target is raised by 20 %
once (still <= ~30 MB binary STL) and the better result kept.
"""
import argparse
import json
import time
from pathlib import Path

import pymeshlab

NAMES = ("final_design", "original", "uniform_shell")


def measures(ms):
    g = ms.get_geometric_measures()
    t = ms.get_topological_measures()
    m = ms.current_mesh()
    return {
        "faces": int(m.face_number()), "vertices": int(m.vertex_number()),
        "volume_mm3": float(g.get("mesh_volume", float("nan"))),
        "boundary_edges": int(t["boundary_edges"]),
        "non_manifold_edges": int(t["non_two_manifold_edges"]),
        "non_manifold_vertices": int(t["non_two_manifold_vertices"]),
        "connected_components": int(t["connected_components_number"]),
        "watertight": int(t["boundary_edges"]) == 0 and int(t["non_two_manifold_edges"]) == 0,
    }


def decimate(full, faces):
    ms = pymeshlab.MeshSet()
    ms.add_mesh(pymeshlab.Mesh(full.vertex_matrix(), full.face_matrix()))
    ms.meshing_decimation_quadric_edge_collapse(
        targetfacenum=faces, qualitythr=0.3, preserveboundary=True, boundaryweight=1.0,
        preservenormal=True, preservetopology=True, optimalplacement=True, planarquadric=True, autoclean=True)
    return ms


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("run_dir", type=Path)
    ap.add_argument("out_dir", type=Path)
    ap.add_argument("--faces", type=int, default=500_000)
    ap.add_argument("--max-dvol", type=float, default=0.005)
    a = ap.parse_args()
    a.out_dir.mkdir(parents=True, exist_ok=True)
    report = {}
    for name in NAMES:
        src = a.run_dir / f"{name}.stl"
        if not src.exists():
            print(f"[skip] {src} missing", flush=True)
            continue
        t0 = time.time()
        ms = pymeshlab.MeshSet()
        ms.load_new_mesh(str(src))
        ms.meshing_remove_duplicate_vertices()
        full_m = measures(ms)
        full = ms.current_mesh()
        print(f"[{name}] loaded {full_m['faces']} faces in {time.time() - t0:.0f} s: {full_m}", flush=True)
        best = None
        for faces in (a.faces, int(a.faces * 1.2)):
            t1 = time.time()
            dm = decimate(full, faces)
            pm = measures(dm)
            dvol = pm["volume_mm3"] / full_m["volume_mm3"] - 1.0
            print(f"[{name}] target {faces}: {pm['faces']} faces, dvol {100 * dvol:+.3f} %, "
                  f"watertight {pm['watertight']}, {time.time() - t1:.0f} s", flush=True)
            if best is None or abs(dvol) < abs(best[2]):
                best = (dm, pm, dvol)
            if abs(dvol) <= a.max_dvol:
                break
        dm, pm, dvol = best
        out = a.out_dir / f"{name}_preview.stl"
        dm.save_current_mesh(str(out), binary=True)
        report[name] = {"source": str(src), "full": full_m, "preview": pm, "rel_volume_change": dvol,
                        "preview_bytes": out.stat().st_size, "seconds": time.time() - t0}
        print(f"[{name}] wrote {out} ({out.stat().st_size / 1e6:.1f} MB)", flush=True)
        del ms, full, dm
    (a.out_dir / "previews.json").write_text(json.dumps(report, indent=2))


if __name__ == "__main__":
    main()
