#!/bin/bash
# Usage: sbatch [--time ..] continuous/research/tubeE_ccx/run.sh "<args1>" "<args2>" ...  (each = tube_ccx.py args)
#SBATCH --job-name=tubeccx
#SBATCH --cpus-per-task=72
#SBATCH --mem=32G
#SBATCH --time=00:30:00
#SBATCH --output=/home/b6bl/owenpeckham.b6bl/shell-lightweighting/.claude/worktrees/cshell-stage2/continuous/research/tubeE_ccx/logs/%x_%j.out
set -uo pipefail
source /home/b6bl/owenpeckham.b6bl/shell-lightweighting/.claude/worktrees/cshell-stage2/continuous/cluster/env.sh; trap 'rm -rf "$TMPDIR"' EXIT
export CUDA_VISIBLE_DEVICES=""
D=/home/b6bl/owenpeckham.b6bl/shell-lightweighting/.claude/worktrees/cshell-stage2/continuous/research/tubeE_ccx
for a in "$@"; do /usr/bin/time -v python $D/tube_ccx.py $a 2> >(grep -E "Maximum resident|Elapsed" >&2); done
