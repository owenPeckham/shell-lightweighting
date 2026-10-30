#!/bin/bash
# P6 acceptance 3: dJ without the buckling term, two-grid (deterministic engine) vs the stored direct references of
# nobuck.sh (*_direct_nb.npz, fe_sparse unchanged), phase-0 designs of both parts and the voided MBB phase-B checkpoint 2.
# Results go to new files (*_twogrid_nbD.npz); the references are not rewritten. Also compares the full-objective MBB
# runs of bench.sh (TAG mbbFa: full determinism) against the hybrid direct reference.
#SBATCH --job-name=tg_nobuckD
#SBATCH --gpus=1
#SBATCH --mem=110G
#SBATCH --time=01:00:00
#SBATCH --output=logs/%x_%j.out
set -euo pipefail
source continuous/cluster/env.sh; trap 'rm -rf "$TMPDIR"' EXIT
D=continuous/research/twogrid; OUT=runs_cshell/TG_bench; R=examples/cshell/runs_cshell
for P in mbb:fixhyb_mbb/MBB cant:fix_cantilever/Cantilever; do
  T=${P%%:*}; RUN=$R/${P#*:}
  python $D/bench_design.py $RUN $OUT/${T}0_twogrid_nbD.npz --engine twogrid --phase 0 --xi $RUN/xi_uniform_phase0.npy --no-buckling
  python $D/bench_design.py --compare $OUT/${T}0_direct_nb.npz $OUT/${T}0_twogrid_nbD.npz
done
RUN=$R/fixhyb_mbb/MBB
python $D/bench_design.py $RUN $OUT/mbb1_twogrid_nbD.npz --engine twogrid --phase 1 --ckpts 2 --no-buckling
python $D/bench_design.py --compare $OUT/mbb1_direct_nb.npz $OUT/mbb1_twogrid_nbD.npz
