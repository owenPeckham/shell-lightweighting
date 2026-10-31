#!/bin/bash
#SBATCH --job-name=gpumem
#SBATCH --gpus=1
#SBATCH --time=00:20:00
#SBATCH --output=/home/b6bl/owenpeckham.b6bl/shell-lightweighting/.claude/worktrees/cshell-stage2/docs/cshell/research/gpu_memory/logs/%x_%j.out
# usage: sbatch [--time ..] job.sh <part> <tag> <h> [extra args to gpumem.py]
set -euo pipefail
source /home/b6bl/owenpeckham.b6bl/shell-lightweighting/.claude/worktrees/cshell-stage2/cshell/cluster/env.sh
trap 'rm -rf "$TMPDIR"' EXIT
W=/home/b6bl/owenpeckham.b6bl/shell-lightweighting/.claude/worktrees/cshell-stage2
cd $W
P=$1; TAG=$2; H=$3; shift 3
nvidia-smi --query-gpu=name,memory.total --format=csv,noheader
python docs/cshell/research/gpu_memory/${SCRIPT:-gpumem.py} examples/$P.json docs/cshell/research/gpu_memory/${P}_${TAG}.json --h $H "$@"
