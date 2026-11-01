#!/bin/bash
# sbatch [--mem=..] docs/cshell/research/investigation/decompose.sh <Part>   (run from the worktree root)
#SBATCH --job-name=inv_decomp
#SBATCH --gpus=1
#SBATCH --time=01:00:00
#SBATCH --output=logs/%x_%j.out
set -euo pipefail
source cshell/cluster/env.sh; trap 'rm -rf "$TMPDIR"' EXIT
P=$1
RUNS=/projects/b6bl/owenpeckham.b6bl/home_data/shell-lightweighting/runs_cshell
RUN=${2:-$RUNS/SR_VAL2/$P}                         # optional: another run directory of part P
OUT=${3:-$RUNS/INV/decomp/$P}
[ -e "$OUT/summary.json" ] && { echo "exists: $OUT"; exit 1; }
python docs/cshell/research/investigation/decompose.py "$RUN" "examples/dshell/runs/$P" "$OUT"
