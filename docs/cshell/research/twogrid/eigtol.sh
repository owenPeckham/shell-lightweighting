#!/bin/bash
# dJ discrepancy direct vs two-grid: engine error or eigenvector truncation? Phase-0 designs of both parts, both
# engines, at a tight eigen tolerance (EIG, default 1e-5), then compare. Submit from the worktree root.
#SBATCH --job-name=tg_eigtol
#SBATCH --gpus=1
#SBATCH --mem=48G
#SBATCH --time=01:30:00
#SBATCH --output=logs/%x_%j.out
set -euo pipefail
source cshell/cluster/env.sh; trap 'rm -rf "$TMPDIR"' EXIT
D=docs/cshell/research/twogrid; OUT=runs_cshell/TG_bench; EIG=${EIG:-1e-5}
R=examples/cshell/runs_cshell
for P in mbb:fixhyb_mbb/MBB cant:fix_cantilever/Cantilever; do
  T=${P%%:*}; RUN=$R/${P#*:}
  for E in direct twogrid; do
    python $D/bench_design.py $RUN $OUT/${T}_${E}_eig${EIG}_ph0.npz --engine $E --phase 0 \
        --xi $RUN/xi_uniform_phase0.npy --eig-tol $EIG --nrhs 7
  done
  python $D/bench_design.py --compare $OUT/${T}_direct_eig${EIG}_ph0.npz $OUT/${T}_twogrid_eig${EIG}_ph0.npz
done
