"""Voxel lambda_b (K_sigma exclusion) of the B2 solid and the exported best design vs element size, same design as the collect
(runs_cshell/VF_B2/best_candidate.stl), fe.engine twogrid. usage: python voxel_conv.py RUN_DIR h_e:h_g [h_e:h_g ...]"""
import copy
import json
import os
import sys
import time

sys.path.insert(0, "/home/b6bl/owenpeckham.b6bl/shell-lightweighting/.claude/worktrees/cshell-stage2")
import numpy as np
import trimesh
from cshell import collect as co, config as ccfg, export as cex
from cshell.optimise import Context


def main():
    rd = os.path.abspath(sys.argv[1])
    cfg = ccfg.load(json.load(open(os.path.join(rd, "config_used.json"))))
    cfg["fe"]["engine"] = "twogrid"
    m = cfg["manufacturing"]
    b_v, b_o = m["wall_bias_mm"], m["outer_bias_mm"]
    rows = co.start_table(rd, cfg)
    best = co.pick_best(rows)
    ctx = Context(cfg, log=print)
    nv = ctx.surf.n_vertices
    xi = np.load(os.path.join(rd, best["xi_file"]))
    D_comp = best["D_comp"]
    h_g_last = float(cfg["grid"]["phases"][-1]["h_g"])
    vf = cex.vertex_fields(ctx, xi, D_comp, h_g_last)
    tm = trimesh.load(os.path.join(rd, "best_candidate.stl"), process=True)
    VE, TE = np.asarray(tm.vertices, float), np.asarray(tm.faces, np.int64)
    a_v = np.where(vf["xt_a"] >= cex.ExportConfig.closed_thr, cex.A_CLOSED, vf["a"])
    for spec in sys.argv[2:]:
        h_e, h_g = [float(x) for x in spec.split(":")]
        t0 = time.time()
        r = dict(res=f"h{h_e}", h_e=h_e, h_g=h_g, shift=0.0)
        model = co.make_model(ctx, r, D_comp, print)
        cd = co.CellData(model, ctx)
        st = co.scoring_settings(model)
        raw0 = co.raw_responses(model, model.rho0, st)
        phiE = co.mesh_phi(cd, VE, TE)
        Dlt = co.asbuilt_increment(cd, vf["d"], a_v, ctx.r_ma, b_v, b_o)
        _, rho = co.chi_rho(cd, phiE + Dlt)
        raw = co.raw_responses(model, rho, st)
        sc = co.score(raw, raw0, st)
        print(f"RESULT h_e {h_e} h_g {h_g} dof {raw['dof']} lam_solid {raw0['lam1']:.4f} lam_best {raw['lam1']:.4f} "
              f"Pb_ratio {raw['lam1'] / raw0['lam1']:.4f} C_solid {raw0['C']:.3f} C_best {raw['C']:.3f} "
              f"K_K0 {sc['K_K0']:.4f} sigq_solid {raw0['sigma_q']:.3f} sigq_best {raw['sigma_q']:.3f} "
              f"seconds {time.time() - t0:.0f}", flush=True)
        co._release(model)
        del model, cd


if __name__ == "__main__":
    main()
