#!/bin/bash
#SBATCH --job-name=gpumem_multi
#SBATCH --gpus=1
#SBATCH --time=01:00:00
#SBATCH --output=/home/b6bl/owenpeckham.b6bl/shell-lightweighting/.claude/worktrees/cshell-stage2/docs/cshell/research/gpu_memory/logs/%x_%j.out
# usage: sbatch multi.sh <runlist file>; each line: <script> <part> <tag> <h> [args]; runs sequentially, continues on failure
source /home/b6bl/owenpeckham.b6bl/shell-lightweighting/.claude/worktrees/cshell-stage2/cshell/cluster/env.sh
trap 'rm -rf "$TMPDIR"' EXIT
W=/home/b6bl/owenpeckham.b6bl/shell-lightweighting/.claude/worktrees/cshell-stage2
cd $W; D=docs/cshell/research/gpu_memory
export GM_ARR_DIR=$SCRATCH/gpu_memory_arrays
nvidia-smi --query-gpu=name,memory.total --format=csv,noheader
while read -r S P TAG H REST; do
  [ -z "$S" ] && continue
  echo "=== $S $P $TAG $H $REST  $(date +%T)"
  if [ "$S" = int64_test.py ]; then python $D/int64_test.py; continue; fi
  python $D/$S examples/$P.json $D/${P}_${TAG}.json --h $H $REST || echo "=== FAILED $S $P $TAG"
done < "$1"
echo "=== end $(date +%T)"
