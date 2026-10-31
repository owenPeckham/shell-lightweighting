"""Body-fitted surface options on a production part (no solve). usage: python diag.py <run_dir> solid|design [variants]
A<h>: marching cubes of the SDF at s = h, gmsh volume inside it directly; B<h>: MC at s = 0.5/3, then the existing
isotropic remesh to h (remesh_surface), gmsh inside. Prints JSON lines 'RESULT ...'."""
import json
import os
import sys
import time

sys.path.insert(0, os.getcwd())
import numpy as np


def gmsh_reparam(V, T, h):
    """gmsh surface remesh of the discrete surface (classify + reparametrise) at uniform size h, then 3D."""
    import tempfile
    import gmsh
    import trimesh
    from cshell import validate as cval
    tmp = tempfile.mkdtemp(prefix="g_")
    trimesh.Trimesh(V, T, process=False).export(tmp + "/s.stl")
    gmsh.initialize()
    try:
        gmsh.option.setNumber("General.Terminal", 1)
        gmsh.option.setNumber("General.Verbosity", 2)
        gmsh.option.setNumber("General.NumThreads", int(os.environ.get("SLURM_CPUS_ON_NODE", 8)))
        gmsh.merge(tmp + "/s.stl")
        gmsh.model.mesh.classifySurfaces(40 * np.pi / 180, True, True, np.pi)
        gmsh.model.mesh.createGeometry()
        loop = gmsh.model.geo.addSurfaceLoop([e[1] for e in gmsh.model.getEntities(2)])
        gmsh.model.geo.addVolume([loop])
        gmsh.model.geo.synchronize()
        for o in ("MeshSizeFromPoints", "MeshSizeFromCurvature", "MeshSizeExtendFromBoundary"):
            gmsh.option.setNumber("Mesh." + o, 0)
        gmsh.option.setNumber("Mesh.MeshSizeMax", h)
        gmsh.option.setNumber("Mesh.MeshSizeMin", h / 8)
        gmsh.option.setNumber("Mesh.Optimize", 1)
        gmsh.option.setNumber("Mesh.OptimizeNetgen", 1)
        gmsh.model.mesh.generate(3)
        tags, xyz, _ = gmsh.model.mesh.getNodes()
        types, _, en = gmsh.model.mesh.getElements(3)
        en = en[list(types).index(4)].reshape(-1, 4)
    finally:
        gmsh.finalize()
    lut = np.zeros(int(tags.max()) + 1, np.int64)
    lut[tags] = np.arange(len(tags))
    nodes, tets = xyz.reshape(-1, 3), lut[en]
    used = np.unique(tets)
    remap = -np.ones(len(nodes), np.int64)
    remap[used] = np.arange(len(used))
    return nodes[used], cval._orient(nodes[used], remap[tets])


def main():
    run_dir, which = sys.argv[1], sys.argv[2]
    variants = sys.argv[3].split(",") if len(sys.argv) > 3 else ["A0.5", "A0.354", "B0.5"]
    from cshell import collect as C, validate as cval
    t0 = time.time()
    if which == "solid":
        from cshell import config as ccfg
        from cshell.optimise import Context
        cfg = ccfg.load(json.load(open(os.path.join(run_dir, "config_used.json"))))
        ctx = Context(cfg, log=print)
        V, T = np.asarray(ctx.V, float), np.asarray(ctx.T, np.int64)
    else:
        V, T, _ = C.load_mesh(os.path.join(run_dir, "best_candidate.stl"))
    print(f"loaded {which}: {len(V)} V {len(T)} T in {time.time() - t0:.0f} s", flush=True)
    cache = {}
    for var in variants:
        kind, hh = var[0], float(var[1:])
        s = hh if kind == "A" else float(os.environ.get("S_MC", 0.5 / 3))
        out = dict(part=os.path.basename(run_dir), which=which, variant=var)
        try:
            t1 = time.time()
            if s not in cache:
                S = cval.levelset_surface(V, T, s, log=print)
                dev = cval.surface_deviation(S[0], S[1], V, T)
                cache[s] = (S, dev)
            (Vs, Ts, info), dev = cache[s]
            out.update(levelset=info, dev_mc=dev, t_mc=time.time() - t1)
            if kind in "BT":
                t1 = time.time()
                Vs, Ts, rinfo = cval.remesh_surface(Vs, Ts, hh, **(dict(thin=float(os.environ.get("THIN", 2.0)), fine=s)
                                                                   if kind == "T" else {}))
                out.update(remesh=rinfo, dev_remesh=cval.surface_deviation(Vs, Ts, V, T), t_remesh=time.time() - t1)
            t1 = time.time()
            if kind == "G":
                nodes, tets = gmsh_reparam(Vs, Ts, hh)
            else:
                nodes, tets = cval._gmsh_inside_robust(Vs, Ts, hh)
            out["t_gmsh"] = time.time() - t1
            q0 = cval.mesh_gate(*cval.to_quadratic(nodes, tets), Vs, Ts, h=hh, enforce=False)
            q0.pop("layers", None)
            out["gate_before_peel"] = q0
            nodes, tets, out["peel"] = cval.peel_boundary_slivers(nodes, tets, hh)
            n10, t10 = cval.to_quadratic(nodes, tets)
            q = cval.mesh_gate(n10, t10, Vs, Ts, h=hh, enforce=False)
            q.pop("layers", None)
            out.update(gate=q, n_tets=int(len(tets)), n_nodes10=int(len(n10)))
        except Exception as e:
            import traceback
            traceback.print_exc()
            out["error"] = repr(e)[:800]
        print("RESULT " + json.dumps(out, default=float), flush=True)


if __name__ == "__main__":
    main()
