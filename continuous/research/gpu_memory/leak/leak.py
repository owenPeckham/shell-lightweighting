"""Does device memory come back after SparseFE.free() on an active-set rebuild? Hollow box, z=0 clamped, 3 rebuilds."""
import gc, sys, numpy as np, cupy as cp
sys.path.insert(0, ".")
from cshell.fe_sparse import ActiveGrid, SparseFE
L, t = int(sys.argv[1]), 4
mode = sys.argv[2]                       # "free" | "free_del_gc" | "reuse_handle"
GiB = 2**30
def free_gib(): return cp.cuda.runtime.memGetInfo()[0] / GiB
pool = cp.get_default_memory_pool()
print("start free", round(free_gib(), 2), flush=True)
for rep in range(3):
    i, j, k = np.meshgrid(*(np.arange(L),) * 3, indexing="ij")
    m = (i < t) | (i >= L - t) | (j < t) | (j >= L - t) | (k < t) | (k >= L - t)
    m[:, :, :] &= True
    g = ActiveGrid(m)
    nn = g.node_shape
    fixed = np.zeros(nn + (3,), bool); fixed[:, :, 0] = True
    load = np.zeros(nn + (3,)); load[:, :, -1, 0] = 1.0 / (nn[0] * nn[1])
    fe = SparseFE(g, fixed, load, 1.0, 1.0, 0.3, precision="fp64")
    fe.factor(cp.ones(g.Ne))
    u = fe.solve(fe.f)
    print(f"rep {rep}: ndof {g.ndof} after factor free {free_gib():.2f} pool used {pool.used_bytes()/GiB:.2f} "
          f"held {pool.total_bytes()/GiB:.2f} est {fe.stats.get('factor_estimate_GiB', 0):.2f}", flush=True)
    del u
    fe.free()
    pool.free_all_blocks()
    if mode == "free_reserved":
        from nvmath.internal.memory import free_reserved_memory
        free_reserved_memory(True)
    print(f"rep {rep}: after fe.free+pool free {free_gib():.2f} held {pool.total_bytes()/GiB:.2f}", flush=True)
    if mode != "free":
        fe._solver = None; fe._K = None; fe._Kf = None; fe._b = None
        del fe, g
        gc.collect(); pool.free_all_blocks(); cp.get_default_pinned_memory_pool().free_all_blocks()
        print(f"rep {rep}: after del+gc {free_gib():.2f} held {pool.total_bytes()/GiB:.2f}", flush=True)
