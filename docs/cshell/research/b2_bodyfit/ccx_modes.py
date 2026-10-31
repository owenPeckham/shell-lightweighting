"""Re-run the collect design deck with buckling mode output (*NODE FILE U) and no stress print; time variants.
usage: python ccx_modes.py <deck.inp> <tag> <threads> <n_modes> [frd 0/1] [stress 0/1]
Units mm, N, MPa. Writes work/<tag>/ and work/<tag>/modes.json (λ, wall time, mode-shape location).
"""
import json, os, re, subprocess, sys, time
from pathlib import Path
import numpy as np

deck, tag, nth, nm = Path(sys.argv[1]), sys.argv[2], int(sys.argv[3]), int(sys.argv[4])
frd = int(sys.argv[5]) if len(sys.argv) > 5 else 1
stress = int(sys.argv[6]) if len(sys.argv) > 6 else 0
wd = Path(__file__).resolve().parent / "work" / tag
wd.mkdir(parents=True, exist_ok=True)
txt = open(deck).read()
if not stress:
    txt = txt.replace("*EL PRINT, ELSET=EALL\nS\n*EL PRINT, ELSET=EALL, TOTALS=NO\nEVOL\n", "")
txt = re.sub(r"\*BUCKLE, SOLVER=SPOOLES\n\d+\n", f"*BUCKLE, SOLVER=SPOOLES\n{nm}\n", txt)
if frd:
    i = txt.rindex("*END STEP")
    txt = txt[:i] + "*NODE FILE\nU\n" + txt[i:]
open(wd / "job.inp", "w").write(txt)
env = dict(os.environ, OMP_NUM_THREADS=str(nth), CCX_NPROC_EQUATION_SOLVER=str(nth), CCX_NPROC_STIFFNESS_ASSEMBLY=str(nth),
           CCX_NPROC_RESULTS=str(nth), OMP_STACKSIZE="256M")
exe = str(Path(os.environ["CSHELL_ENV"]) / "bin" / "ccx")
t0 = time.time()
with open(wd / "job.log", "w") as lg:
    rc = subprocess.run([exe, "-i", "job"], cwd=wd, stdout=lg, stderr=subprocess.STDOUT, env=env).returncode
sec = time.time() - t0
out = dict(tag=tag, threads=nth, n_modes=nm, frd=frd, stress=stress, rc=rc, seconds=sec)
dat = open(wd / "job.dat").read()
m = re.search(r"B U C K L I N G.*?\n(.*?)(?:\n\s*\n\s*\n|\Z)", dat, re.S)
out["lambda"] = [float(l.split()[1]) for l in m.group(1).splitlines() if re.match(r"^\s*\d+\s+[-+0-9.Ee]+\s*$", l)]
blk = dat[dat.index("displacements"):]
U = np.array([l.split() for l in blk.splitlines()[1:] if re.match(r"^\s*\d+\s+[-+]", l)], float)
sys.path.insert(0, str(Path(__file__).resolve().parent))
import bf_gpu
_, _, _, ln, lv, _, _ = bf_gpu.read_deck(deck)
ud = {int(r[0]) - 1: r[1:4] for r in U}
out["C"] = float(sum(np.dot(v, ud[int(n)]) for n, v in zip(ln, lv)))
if frd:
    nodes, modes, cur = {}, [], None
    with open(wd / "job.frd") as f:
        X = None
        for ln in f:
            if ln.startswith("    2C"):
                n = int(ln.split()[1]); arr = [f.readline() for _ in range(n)]
                X = np.array([[float(a[13:25]), float(a[25:37]), float(a[37:49])] for a in arr])
                ids = np.array([int(a[3:13]) for a in arr])
            elif ln.startswith(" -4  DISP"):
                for _ in range(4):
                    f.readline()
                arr = []
                for l2 in f:
                    if l2.startswith(" -3"):
                        break
                    arr.append(l2)
                modes.append(np.array([[float(a[13:25]), float(a[25:37]), float(a[37:49])] for a in arr]))
    np.savez_compressed(wd / "modes.npz", X=X.astype(np.float32), U=np.array(modes, np.float32))
    loc = []
    for k, Um in enumerate(modes):
        a = np.linalg.norm(Um, axis=1)
        r = np.hypot(X[:, 1], X[:, 2])
        top = np.argsort(a)[-50:]
        w = a ** 2
        loc.append(dict(mode=k + 1, x_top50=[float(X[top, 0].min()), float(X[top, 0].max())],
                        r_top50=[float(r[top].min()), float(r[top].max())],
                        th_top50=np.degrees(np.arctan2(X[top, 2], X[top, 1])).round(1).tolist()[:10],
                        frac_nodes_gt_half=float((a > 0.5 * a.max()).mean()),
                        x_centroid_w=float((w * X[:, 0]).sum() / w.sum()),
                        frac_energy_in_top1pct=float(np.sort(w)[-len(w) // 100:].sum() / w.sum()),
                        xmax_node=X[np.argmax(a)].tolist()))
    out["modes"] = loc
json.dump(out, open(wd / "modes.json", "w"), indent=1)
print(json.dumps(out, indent=1))
