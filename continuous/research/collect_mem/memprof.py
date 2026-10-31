"""Host-memory profile of `cshell collect` per stage (diagnosis only; cshell/ is not edited, functions are wrapped).

Usage (GPU job, from the repo root):  python continuous/research/collect_mem/memprof.py <run_dir_copy>
A sampler thread records every 2 s: parent RSS, the summed PSS of the parent and its (forked) children, and the job
cgroup's anon / file memory (what the Slurm OOM killer counts). Wrapped stages log before/after/peak to
<run_dir_copy>/memprof.json. PCG failures in raw_responses are caught (NaN responses) so that every stage runs.
Validation is disabled (validate=False).
"""
import json, os, sys, threading, time
import numpy as np
import psutil

sys.path.insert(0, os.getcwd())
from cshell import collect as C, export as X, optimise as O     # noqa: E402
import trimesh                                                   # noqa: E402

run_dir = os.path.abspath(sys.argv[1])
P = psutil.Process()
GiB = 2 ** 30


def cg_path():
    for line in open("/proc/self/cgroup"):
        if line.startswith("0::"):
            return "/sys/fs/cgroup" + line.strip()[3:]
    return None


CG = cg_path()


def cg_mem():
    out = {}
    try:
        p = CG
        while p and p != "/sys/fs/cgroup" and not os.path.exists(os.path.join(p, "memory.max")):
            p = os.path.dirname(p)
        # walk up to the Slurm job level (the one with a finite memory.max)
        q = CG
        while q and q != "/sys/fs/cgroup":
            mx = open(os.path.join(q, "memory.max")).read().strip() if os.path.exists(os.path.join(q, "memory.max")) else "max"
            if mx != "max":
                break
            q = os.path.dirname(q)
        st = dict(l.split() for l in open(os.path.join(q, "memory.stat")))
        out = dict(anon=int(st["anon"]) / GiB, file=int(st["file"]) / GiB,
                   current=int(open(os.path.join(q, "memory.current")).read()) / GiB)
    except Exception:
        pass
    return out


state = dict(stage="start", peak={}, trace=[])
lock = threading.Lock()


def sample():
    rss = P.memory_info().rss / GiB
    pss = 0.0
    for p in [P] + P.children(recursive=True):
        try:
            pss += p.memory_full_info().pss / GiB
        except Exception:
            pass
    s = dict(t=time.time(), rss=rss, pss=pss, **cg_mem())
    return s


def sampler():
    while True:
        s = sample()
        with lock:
            stg = state["stage"]
            pk = state["peak"].setdefault(stg, dict(rss=0, pss=0, anon=0, current=0))
            for k in pk:
                pk[k] = max(pk[k], s.get(k, 0))
            state["trace"].append(dict(stage=stg, **{k: round(v, 3) if k != "t" else round(v, 1) for k, v in s.items()}))
        if len(state["trace"]) % 15 == 0:
            try:
                dump()
            except Exception:
                pass
        time.sleep(2.0)


threading.Thread(target=sampler, daemon=True).start()
REC = []
counter = {}


def dump():
    with lock:
        json.dump(dict(stages=REC, peak=state["peak"], trace=state["trace"]), open(os.path.join(run_dir, "memprof.json"), "w"))


def wrap(mod, name, label=None, catch=False):
    f = getattr(mod, name)
    lab0 = label or name

    def g(*a, **k):
        counter[lab0] = counter.get(lab0, 0) + 1
        lab = f"{lab0}#{counter[lab0]}"
        with lock:
            prev = state["stage"]
            state["stage"] = lab
        b = sample(); t0 = time.time()
        try:
            return f(*a, **k)
        except RuntimeError as e:
            if not catch:
                raise
            print(f"[memprof] {lab} caught {e!r}", flush=True)
            return dict(C=float("nan"), sigma_pn=float("nan"), sigma_q=float("nan"), mu_tilde=float("nan"),
                        lam1=float("nan"), dof=0, seconds=0.0, error=repr(e)[:200])
        finally:
            a_ = sample()
            with lock:
                pk = dict(state["peak"].get(lab, {}))
                state["stage"] = prev
            REC.append(dict(stage=lab, seconds=time.time() - t0, before=b, after=a_, peak=pk))
            print(f"[memprof] {lab:<28} {time.time() - t0:7.1f} s  rss {b['rss']:.1f}->{a_['rss']:.1f}  "
                  f"peak rss {pk.get('rss', 0):.1f} pss {pk.get('pss', 0):.1f} cg anon {pk.get('anon', 0):.1f} GiB",
                  flush=True)
            dump()
    setattr(mod, name, g)


if "--skip-export" in sys.argv:          # re-use the STLs / export jsons of the source run (export memory is in its json)
    SRC = os.path.abspath(sys.argv[sys.argv.index("--skip-export") + 1])
    import shutil

    def _eb(ctx, out_dir, **kw):
        os.makedirs(out_dir, exist_ok=True)
        out = {}
        for nm in ("original", "uniform_shell"):
            for ext in (".stl", "_export.json"):
                dst = os.path.join(out_dir, nm + ext)
                if not os.path.exists(dst):
                    os.symlink(os.path.join(SRC, "export_baselines", nm + ext), dst)
            out[nm] = json.load(open(os.path.join(out_dir, nm + "_export.json")))
        out["uniform_shell"].setdefault("t_final_mm", 2.0497 + out["uniform_shell"]["mass_correction"]["delta"])
        return out

    def _ed(ctx, xi, run_dir_, D_comp, name="best_candidate", **kw):
        for ext in (".stl", "_export.json"):
            dst = os.path.join(run_dir_, name + ext)
            if not os.path.exists(dst):
                os.symlink(os.path.join(SRC, name + ext), dst)
        return json.load(open(os.path.join(run_dir_, name + "_export.json")))
    X.export_baselines, X.export_design = _eb, _ed

wrap(O, "Context")
C.Context = O.Context   # collect imports it inside the function from .optimise -> patched module attribute is used
wrap(C, "make_model")
wrap(C, "CellData")
wrap(C, "raw_responses", catch=True)
wrap(C, "mesh_phi")
wrap(C, "chi_rho")
wrap(C, "asbuilt_increment")
wrap(X, "export_baselines")
wrap(X, "export_design")
wrap(X, "vertex_fields")
wrap(trimesh, "load", "trimesh.load")
t0 = time.time()
try:
    S = C.collect(run_dir, validate=False)
    print("[memprof] collect finished", flush=True)
except Exception as e:
    import traceback
    traceback.print_exc()
    print(f"[memprof] collect raised {e!r}", flush=True)
with lock:
    state["stage"] = "end"
REC.append(dict(stage="total", seconds=time.time() - t0, after=sample()))
dump()
