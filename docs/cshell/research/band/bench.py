"""Band build benchmark on the cantilever (THEORY §3.1, §4.2). Usage: python bench.py h_g [D_comp] [--extend D_new]"""
import json, resource, sys, time
from pathlib import Path
import numpy as np
sys.path.insert(0, str(Path(__file__).resolve().parents[4]))
from cshell import band as B
import warp as wp

h = float(sys.argv[1]); D = float(sys.argv[2]) if len(sys.argv) > 2 and not sys.argv[2].startswith("--") else 4.0
ext = float(sys.argv[sys.argv.index("--extend") + 1]) if "--extend" in sys.argv else None
step = Path(__file__).resolve().parents[4] / "examples" / "parts" / "Solid Cantilever.step"
wp.init(); dev = wp.get_device()
out = dict(h_g=h, D_comp=D, device=str(dev), gpu=dev.name if dev.is_cuda else None)
rss = lambda: resource.getrusage(resource.RUSAGE_SELF).ru_maxrss / 1024.0     # MiB

t = time.time()
from vshell.geometry import load_step, tessellate
V, T, _ = tessellate(load_step(str(step)), lin=0.02)
out.update(tess_seconds=time.time() - t, n_vertices=len(V), n_triangles=len(T), rss_after_tess_MiB=rss())
t = time.time()
b = B.Band.build(V, T, h, D, backend="warp", verbose=False)
out.update(build_seconds=time.time() - t, n_cells=b.n, band_MiB=b.nbytes / 2**20, grid_shape=list(b.grid.shape),
           dense_cells=int(np.prod(b.grid.shape, dtype=np.float64)), peak_rss_MiB=rss(), warp_mempool_high_MiB=None)
try:
    out["warp_mempool_high_MiB"] = wp.get_mempool_used_mem_high(dev) / 2**20
except Exception:
    pass
t = time.time(); V0 = b.V0(); out.update(V0=V0, V0_seconds=time.time() - t, V0_rel_to_theory_60942=V0 / 60942.0 - 1)
t = time.time(); b.save(B.cache_path(B.step_sha256(step), h, D, 0.02), dict(sha256=B.step_sha256(step), tol=0.02))
out["save_seconds"] = time.time() - t
t = time.time(); c = B.Band.load(B.cache_path(B.step_sha256(step), h, D, 0.02)); out["load_seconds"] = time.time() - t
assert np.array_equal(c.flat, b.flat)
del c
if ext:
    t = time.time(); e = b.extend(ext); out.update(extend_to=ext, extend_seconds=time.time() - t, extend_n_cells=e.n,
                                                    peak_rss_MiB_after_extend=rss())
print(json.dumps(out, indent=1))
Path(__file__).with_name(f"bench_h{h:g}_D{D:g}.json").write_text(json.dumps(out, indent=1))
