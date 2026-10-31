#!/bin/bash
# Full-objective J / dJ of the deterministic two-grid runs (bench.sh, MBB phase B ckpts 2, 3) against the hybrid fp64
# direct reference, and run-to-run (a vs b) differences. CPU only.
#SBATCH --job-name=tg_cmpdet
#SBATCH --cpus-per-task=72
#SBATCH --mem=16G
#SBATCH --time=00:20:00
#SBATCH --output=logs/%x_%j.out
set -euo pipefail
source continuous/cluster/env.sh; trap 'rm -rf "$TMPDIR"' EXIT
export CUDA_VISIBLE_DEVICES=""
B=runs_cshell/TG_bench; C="python continuous/research/twogrid/bench_design.py --compare"
echo "== direct vs previous two-grid (atomic kernels)"; $C $B/mbb_directM_ph1.npz $B/mbbATOM_ph1.npz
echo "== direct vs full-deterministic a";  $C $B/mbb_directM_ph1.npz $B/mbbFa_ph1.npz
echo "== direct vs kernels-deterministic a"; $C $B/mbb_directM_ph1.npz $B/mbbKa_ph1.npz
echo "== run to run: atomic vs atomic not run; full a vs b"; $C $B/mbbFa_ph1.npz $B/mbbFb_ph1.npz
echo "== run to run: kernels a vs b"; $C $B/mbbKa_ph1.npz $B/mbbKb_ph1.npz
