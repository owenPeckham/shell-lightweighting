#!/bin/bash
# Engine accuracy of dJ without the eigenvector-dependent buckling term (Settings.buckling=False): direct vs two-grid on
# the phase-0 designs of both parts and on the voided MBB phase-B checkpoint 2 (direct = hybrid fp64, hence 240G).
#SBATCH --job-name=tg_nobuck
#SBATCH --gpus=1
#SBATCH --mem=240G
#SBATCH --time=01:00:00
#SBATCH --output=logs/%x_%j.out
set -euo pipefail
source continuous/cluster/env.sh; trap 'rm -rf "$TMPDIR"' EXIT
D=continuous/research/twogrid; OUT=runs_cshell/TG_bench; R=examples/cshell/runs_cshell
cmp() { python $D/bench_design.py --compare $OUT/$1_direct_nb.npz $OUT/$1_twogrid_nb.npz; }
for P in mbb:fixhyb_mbb/MBB cant:fix_cantilever/Cantilever; do
  T=${P%%:*}; RUN=$R/${P#*:}
  for E in direct twogrid; do
    python $D/bench_design.py $RUN $OUT/${T}0_${E}_nb.npz --engine $E --phase 0 --xi $RUN/xi_uniform_phase0.npy --no-buckling
  done
  cmp ${T}0
done
RUN=$R/fixhyb_mbb/MBB
for E in direct twogrid; do
  python $D/bench_design.py $RUN $OUT/mbb1_${E}_nb.npz --engine $E --phase 1 --ckpts 2 --no-buckling
done
cmp mbb1
