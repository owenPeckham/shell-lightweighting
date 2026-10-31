#!/bin/bash
# Reproducer: two-grid forced on B2 (h_e 0.8 as in test_collect), GPU and CPU backends, 2 iterations per start. The
# CPU test run 7808904 hit "Factor is exactly singular" in the coarse SuperLU factor on this case.
#SBATCH --job-name=tg_b2
#SBATCH --gpus=1
#SBATCH --mem=32G
#SBATCH --time=00:40:00
#SBATCH --output=logs/%x_%j.out
set -uo pipefail
source continuous/cluster/env.sh; trap 'rm -rf "$TMPDIR"' EXIT
for BK in ${BACKENDS:-gpu cpu}; do for S in uniform neutral; do
  echo "== backend $BK start $S"
  python -m cshell run examples/cshell/bench/B2_slender.json --start $S \
    --set "grid.phases=[{\"h_e\":0.8,\"h_g\":0.2}]" "optimiser.max_iter=[2]" fe.engine=twogrid fe.backend=$BK \
    validation.body_fitted=false output_dir=$PWD/runs_cshell/TG_b2_$BK 2>&1 | grep -vi warn | tail -n 15
done; done
