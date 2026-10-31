"""Design statistics of a saved ξ: python phase_stats.py <cfg> <xi.npy> <phase>"""
import json, sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parents[4]))
import numpy as np
from cshell import config
from cshell.optimise import Context, PhaseModel, phase_stats
cfg = config.load(sys.argv[1]); ctx = Context(cfg); ph = int(sys.argv[3])
m = PhaseModel(ctx, ph, cfg["design"]["outward_extent_initial_mm"])
print(json.dumps(phase_stats(m, np.load(sys.argv[2]), ctx.surf.n_vertices)))
