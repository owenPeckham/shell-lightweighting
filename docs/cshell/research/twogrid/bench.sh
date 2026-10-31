#!/bin/bash
# One engine on one part: phase-B checkpoints (2, 3) then the phase-A design, each in its own process with its own
# nvidia-smi peak and /usr/bin/time RSS. Env: RUN (run dir), TAG (output prefix), ENGINE, ARGS (extra bench args).
# Submit from the worktree root: sbatch [--mem=240G] docs/cshell/research/twogrid/bench.sh
#SBATCH --job-name=tg_bench
#SBATCH --gpus=1
#SBATCH --mem=110G
#SBATCH --time=01:00:00
#SBATCH --output=logs/%x_%j.out
set -euo pipefail
source cshell/cluster/env.sh; trap 'rm -rf "$TMPDIR"' EXIT
OUT=runs_cshell/TG_bench; mkdir -p $OUT
D=docs/cshell/research/twogrid
for PH in ${PHASES:-1 0}; do
  nvidia-smi --query-gpu=memory.used --format=csv,noheader,nounits -lms 250 > $TMPDIR/mem$PH.txt &
  SMI=$!
  if [ $PH = 1 ]; then SEL="--ckpts 2,3"; else SEL="--xi $RUN/xi_uniform_phase0.npy"; fi
  /usr/bin/time -v python $D/bench_design.py $RUN $OUT/${TAG}_ph$PH.npz --engine $ENGINE --phase $PH $SEL ${ARGS:-} \
      2> $TMPDIR/time$PH.txt || { tail -30 $TMPDIR/time$PH.txt; exit 1; }
  kill $SMI
  echo "PHASE $PH PEAK_GPU_MiB $(sort -n $TMPDIR/mem$PH.txt | tail -1) BASE_MiB $(head -1 $TMPDIR/mem$PH.txt)"
  grep -E "Maximum resident|Elapsed" $TMPDIR/time$PH.txt
  grep -v "^\s" $TMPDIR/time$PH.txt | grep -iv "command being\|exit status" | tail -5 || true
done
