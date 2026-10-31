"""Gather runs/*/result.json -> summary.json (one row per run: k at 100 N for contact, k = F^2/2W for patch)."""
import json
from pathlib import Path

D = Path(__file__).resolve().parent
rows = []
for f in sorted((D / "runs").glob("*/result.json")):
    r = json.loads(f.read_text())
    row = {k: r.get(k) for k in ("size", "level", "bc", "nr", "nt", "hc", "K", "nlgeom", "pw", "seconds", "error")}
    row["run"] = f.parent.name
    if r.get("bc") == "contact" and r.get("curve"):
        c = r["curve"]
        row.update(k=r.get("k_100N_interp"), k_at_10N=[x["k"] for x in c if 9 < x["F"] < 12][0] if any(9 < x["F"] < 12 for x in c) else None,
                   F_end=c[-1]["F"], contact_half_x_end=c[-1]["contact_half_length_x"],
                   contact_half_y_end=c[-1]["contact_half_width_y"], ovalisation_mid_end=c[-1]["ovalisation_mid"])
    else:
        row.update(k=r.get("k"), ovalisation_mid=r.get("ovalisation_mid"))
    rows.append(row)
(D / "summary.json").write_text(json.dumps(rows, indent=1))
for r in rows:
    print(f"{r['run']:28s} k={r['k']}" + (f"  k10N={r.get('k_at_10N')}" if r.get("k_at_10N") else ""))
