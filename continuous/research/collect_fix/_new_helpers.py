# =====================================================================================================================
# Host memory (PLAN_collect_fixes P5)
# =====================================================================================================================
def rss_gib() -> float:
    """Resident set size of this process (GiB, /proc/self/status)."""
    try:
        for line in open("/proc/self/status"):
            if line.startswith("VmRSS:"):
                return int(line.split()[1]) / 2 ** 20
    except OSError:
        pass
    return float("nan")


def trim_host_memory() -> None:
    """Return freed host memory to the OS: cyclic GC, CuPy's device and pinned pools, glibc malloc_trim(0)."""
    import ctypes
    import gc
    gc.collect()
    try:
        import cupy as cp
        cp.get_default_memory_pool().free_all_blocks()
        cp.get_default_pinned_memory_pool().free_all_blocks()
    except Exception:
        pass
    try:
        ctypes.CDLL("libc.so.6").malloc_trim(0)
    except OSError:
        pass


def free_model(model) -> None:
    """Release a PhaseModel's FE engine and device band, then its host arrays (caller drops its references)."""
    _release(model)
    for k in ("band", "disc", "dm", "ds", "grid_fe", "fe", "fixed", "load", "rho0"):
        if hasattr(model, k):
            setattr(model, k, None)
    trim_host_memory()


def _grid_info(model):
    """What DenseLookup / mode_magnitudes need of a model, without holding the model."""
    from types import SimpleNamespace
    return SimpleNamespace(origin=np.asarray(model.origin, float), h_e=model.h_e, elem_shape=model.elem_shape,
                           node_shape=model.node_shape)


# =====================================================================================================================
# Re-scoring cell maps (THEORY §8.6.1): host for the CPU chain, `DeviceDesignMap.rescore` for the GPU chain
# =====================================================================================================================
def stl_soup(path, c0=None) -> np.ndarray:
    """Vertices (3n, 3) float32 of the STL's triangles, unwelded (binary STL through np.memmap, chunked; ASCII via
    trimesh). c0: subtract it in float64 before rounding to float32 (the Warp frame of `band.make_backend`).
    No welding is needed for the winding-number SDF (triangle-based) or the volume."""
    size = os.path.getsize(path)
    with open(path, "rb") as f:
        head = f.read(84)
    n = int(np.frombuffer(head[80:84], "<u4")[0]) if len(head) == 84 else -1
    if n >= 0 and 84 + 50 * n == size:
        dt = np.dtype([("n", "<f4", 3), ("v", "<f4", (3, 3)), ("attr", "<u2")])
        mm = np.memmap(path, dt, "r", offset=84, shape=(n,))
        out = np.empty((3 * n, 3), np.float32)
        ch = 1 << 22
        for s in range(0, n, ch):
            v = np.asarray(mm["v"][s:s + ch]).reshape(-1, 3)
            out[3 * s:3 * s + len(v)] = v if c0 is None else (v.astype(np.float64) - c0)
        del mm
    else:
        import trimesh
        tm = trimesh.load(path, process=False, force="mesh")
        v = np.asarray(tm.triangles, np.float64).reshape(-1, 3)
        out = (v if c0 is None else v - c0).astype(np.float32)
    out += np.float32(0.0)                                   # −0 → +0 (welding by bit pattern)
    return out


def weld(soup: np.ndarray):
    """Vertex weld of an unwelded (3n, 3) float32 soup by exact coordinates: (V float64, T int64)."""
    u = soup.view(np.uint32)
    order = np.lexsort((u[:, 2], u[:, 1], u[:, 0]))
    us = u[order]
    new = np.ones(len(us), bool)
    new[1:] = np.any(us[1:] != us[:-1], axis=1)
    del us
    idx = np.cumsum(new) - 1
    inv = np.empty(len(order), np.int64)
    inv[order] = idx
    V = soup[order[new]].astype(np.float64)
    return V, inv.reshape(-1, 3)


def load_mesh(path):
    """(V, T, volume) of an STL, welded (validation mesher input)."""
    V, T = weld(stl_soup(path))
    a, b, c = V[T[:, 0]], V[T[:, 1]], V[T[:, 2]]
    vol = float(np.einsum("ij,ij->", a, np.cross(b, c))) / 6.0
    return V, T, vol


def mesh_phi(model, soup_c: np.ndarray, chunk: int = 1 << 21):
    """Signed distance (negative inside) of the closed triangle soup `soup_c` (centred by model.band.c0, float32) at
    the band cell centres: float32, on the device for the GPU chain (CuPy), else NumPy."""
    from .band import make_backend
    band = model.band
    c0 = band.c0
    T = np.arange(len(soup_c), dtype=np.int32).reshape(-1, 3)
    be = make_backend(soup_c, T)
    del T
    gpu = model.geometry_device == "gpu"
    if gpu:
        import cupy as cp
        out = cp.empty(band.n, cp.float32)
    else:
        out = np.empty(band.n, np.float32)
    for s in range(0, band.n, chunk):
        sl = slice(s, min(band.n, s + chunk))
        q = be.query(band.centres(sl) - c0, signed=True)[0]
        out[sl] = cp.asarray(q) if gpu else q
    del be
    return out


def design_fields(d, a, closed, r_ma, b_v: float) -> dict:
    """Vertex fields of a re-scored design: d, a, cl (closed 0/1) and sv = b_v S((r_MA − a)/b_v) (THEORY §3.9)."""
    from .design import smoothstep
    S, _ = smoothstep(np, (r_ma - np.asarray(a, np.float64)) / b_v)
    return dict(d=np.asarray(d, np.float64), a=np.asarray(a, np.float64), cl=np.asarray(closed, np.float64),
                sv=b_v * S)


def cell_rescore(model, phiE, f: dict | None, b_o: float, solid: bool = False, want_rho: bool = True):
    """ρ (element box, host float64) of ψ_ab = ψ + Δ and the nominal volume of ψ (THEORY §3.9, §8.6.1).

    ψ = φ_E (phiE, the exported mesh) or, with phiE None, the parametric max(φ − Pd, −φ − Pa) of `f`;
    Δ = b_o on the outer branch of max(φ − Pd, −φ − Pa), P sv on the void branch; cells whose positive-weight
    vertices are all closed use a = A_CLOSED (export closure rule, §8.2). solid: ψ = φ_E (or φ), Δ = b_o."""
    from .export import A_CLOSED
    if model.geometry_device == "gpu":
        model.disc.upload()
        f = f or {}
        return model.dm.rescore(phiE, f.get("d"), f.get("a"), f.get("cl"), f.get("sv"), b_o, A_CLOSED, solid=solid,
                                want_rho=want_rho)
    from .design import heaviside
    D = model.disc
    phi = np.asarray(D.phi, np.float64)
    if solid:
        pn = phi if phiE is None else phiE.astype(np.float64)
        dl = b_o
    else:
        P = D.P
        dc, ac = P @ f["d"], P @ f["a"]
        ac = np.where(P @ f["cl"] >= 1.0 - 1e-9, A_CLOSED, ac)
        o, i = phi - dc, -phi - ac
        del dc, ac
        outer = o >= i
        pn = np.where(outer, o, i) if phiE is None else phiE.astype(np.float64)
        del o, i
        dl = np.where(outer, b_o, P @ f["sv"])
    V = float(heaviside(np, -pn, D.eps)[0].sum()) * D.h_g ** 3 + D.V_extra
    rho = None
    if want_rho:
        rho = np.bincount(D.elem, weights=heaviside(np, -(pn + dl), D.eps)[0], minlength=D.n_elem) / D.k ** 3
    return rho, V


# =====================================================================================================================
# Boundary conditions consistent with the design (PLAN_collect_fixes P1, P4(i))
# =====================================================================================================================
_SHIFTS = [(a, b, c) for a in (0, 1) for b in (0, 1) for c in (0, 1)]


def _node_touch(mat):
    """Nodes with at least one adjacent element in `mat` (element box bool)."""
    es = mat.shape
    out = np.zeros(tuple(n + 1 for n in es), bool)
    for a, b, c in _SHIFTS:
        out[a:a + es[0], b:b + es[1], c:c + es[2]] |= mat
    return out


def _elem_touch(nodes):
    """Elements with at least one corner node in `nodes` (node box bool)."""
    es = tuple(n - 1 for n in nodes.shape)
    out = np.zeros(es, bool)
    for a, b, c in _SHIFTS:
        out |= nodes[a:a + es[0], b:b + es[1], c:c + es[2]]
    return out


class DesignBCs:
    """Material-aware BCs for re-scoring one model (THEORY §4.3 node rule restricted to nodes that touch material).

    A node touches material when one of its ≤ 8 adjacent elements has ρ ≥ `thr`. Material components (26-connected
    elements with ρ ≥ thr) that contain no element at a clamped material node are dropped (ρ = 0, P4(i): islands held
    only by ρ_min stall the PCG). Supports clamp only nodes that touch the remaining material; each load is shared
    equally among its load-face nodes that touch it (the dropped share is redistributed and reported)."""

    def __init__(self, model, ctx, thr: float = 1e-3):
        from . import fea
        self.model, self.thr = model, thr
        self.fixed0, self.load0 = model.fixed, model.load
        self.fixed_n = model.fixed.any(-1)
        self.loads = [(fea.nodes_near_faces(model.origin, model.h_e, model.node_shape, ctx.V, ctx.T, ctx.TF,
                                            L["faces"]), np.asarray(L["force"], float)) for L in ctx.loads]
        cap = ctx.cfg["capacity"]
        self.r_ex = int(np.ceil(cap["exclude_bc_mm"] / model.h_e))
        self.mount_ok = None
        if not cap["include_mounts"]:
            if model.geometry_device == "gpu":
                model.disc.upload()
            self.mount_ok = model.dm.elem_mean(ctx.mount.astype(float)).reshape(model.elem_shape) < 0.5
        self.mask0, self.ksig0 = model.stress_mask_dense, model.ksig_mask_dense
        # same rule as PhaseModel._stress_mask (elem_mean on the GPU sums with atomics: a tie at 0.5 could differ)
        self.mask_mismatch = int(np.count_nonzero(self._stress_mask(self.fixed0, self.load0) != self.mask0))
        self.current = "original"

    def _stress_mask(self, fixed, load):
        from scipy import ndimage
        bc = fixed.any(-1) | (load != 0).any(-1)
        m = ~ndimage.binary_dilation(bc, iterations=self.r_ex)[:-1, :-1, :-1]
        if self.mount_ok is not None:
            m &= self.mount_ok
        return m.ravel()

    def adjust(self, rho: np.ndarray):
        """(ρ without islands, fixed, load, info) for the design ρ (element box)."""
        from scipy import ndimage
        m = self.model
        es = m.elem_shape
        mat = (rho >= self.thr).reshape(es)
        fx = self.fixed_n & _node_touch(mat)
        lab, nlab = ndimage.label(mat, structure=np.ones((3, 3, 3), bool))
        keep = np.zeros(nlab + 1, bool)
        keep[np.unique(lab[_elem_touch(fx) & mat])] = True
        keep[0] = False
        island = mat & ~keep[lab]
        del lab
        rho2 = rho
        info = dict(components=int(nlab), islands=int(nlab - keep[1:].sum()),
                    island_mm3=float(rho[island.ravel()].sum()) * m.h_e ** 3)
        if island.any():
            rho2 = rho.copy()
            rho2[island.ravel()] = 0.0
            mat &= ~island
        touch = _node_touch(mat).ravel()
        fixed = (self.fixed_n.ravel() & touch)
        info["fixed_nodes"], info["fixed_dropped"] = int(fixed.sum()), int(self.fixed_n.sum() - fixed.sum())
        load = np.zeros((touch.size, 3))
        Fsum = sum(float(np.linalg.norm(F)) for _, F in self.loads)
        dropped = 0.0
        nd = []
        for ln, F in self.loads:
            k = ln[touch[ln]]
            if k.size == 0:
                raise RuntimeError("re-scoring: no load-face node touches material")
            load[k] += F / k.size
            dropped += float(np.linalg.norm(F)) * (1.0 - k.size / ln.size)
            nd.append(int(ln.size - k.size))
        info["load_nodes_dropped"] = nd
        info["load_share_redistributed"] = dropped / Fsum if Fsum > 0 else 0.0
        fixed = np.repeat(fixed[:, None], 3, 1).reshape(self.fixed0.shape)
        load = load.reshape(self.load0.shape)
        # check: every loaded node touches material (all 8 adjacent ρ < thr counts as isolated)
        info["isolated_load_nodes"] = int(((load != 0).any(-1).ravel() & ~touch).sum())
        return rho2, fixed, load, info

    def apply(self, fixed, load) -> None:
        """Set the model's BCs (and the stress / K_σ masks that depend on them); the FE engine is rebuilt lazily."""
        m = self.model
        same = np.array_equal(fixed, self.fixed0) and np.array_equal(load, self.load0)
        key = "original" if same else id(load)
        if key == self.current:
            return
        m.fixed, m.load = (self.fixed0, self.load0) if same else (fixed, load)
        m.stress_mask_dense = self.mask0 if same else self._stress_mask(fixed, load)
        m.ksig_mask_dense = self.ksig0 if same else m._ksig_mask()
        if m.fe is not None:
            m.fe.free()
            m.fe = None
        m.grid_fe = None
        self.current = key
