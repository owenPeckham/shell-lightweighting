#!/bin/bash
# Acceptance 3: 3-iteration phase-B run of one part with one engine. Env: PART (cantilever|mbb), ENG (auto|direct).
# Output runs_cshell/TG_run3_<eng>_<part>. Submit from the worktree root (direct needs --mem=240G: hybrid factor).
#SBATCH --job-name=tg_run3
#SBATCH --gpus=1
#SBATCH --mem=110G
#SBATCH --time=01:00:00
#SBATCH --output=logs/%x_%j.out
set -euo pipefail
source continuous/cluster/env.sh; trap 'rm -rf "$TMPDIR"' EXIT
OUT=$PWD/runs_cshell/TG_run3_${ENG}_${PART}
nvidia-smi --query-gpu=memory.used --format=csv,noheader,nounits -lms 500 > $TMPDIR/mem.txt &
SMI=$!
/usr/bin/time -v python -m cshell run examples/cshell/$PART.json --start uniform --set "optimiser.max_iter=[1,3]" \
    output_dir=$OUT fe.engine=$ENG 2> $TMPDIR/time.txt || { tail -40 $TMPDIR/time.txt; exit 1; }
kill $SMI; echo "PEAK_GPU_MiB $(sort -n $TMPDIR/mem.txt | tail -1)"
grep -E "Maximum resident|Elapsed" $TMPDIR/time.txt
cat $OUT/*/history.csv
