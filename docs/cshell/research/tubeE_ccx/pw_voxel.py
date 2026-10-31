"""Step 3: voxel (cshell fe_sparse, quarter model, h = 0.2 mm) k vs load-patch width for tube E (and A), to find the
patch width that reproduces the ccx roller-contact k. Reuses tube() from ../tubeE/tubes_pw.py unchanged.
Usage: python pw_voxel.py out.json"""
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "tubeE"))
from tubes_pw import tube  # noqa: E402

rows = []
for size, h, pw in [("E", .2, .4), ("E", .2, .8), ("E", .2, 1.2), ("E", .2, 1.6), ("A", .2, 4.0), ("A", .2, 1.2)]:
    try:
        r = tube(size, h, pw=pw)
    except Exception as e:
        r = dict(size=size, h=h, pw=pw, error=repr(e))
    rows.append(r); print(json.dumps(r), flush=True)
    Path(sys.argv[1]).write_text(json.dumps(rows, indent=1))
