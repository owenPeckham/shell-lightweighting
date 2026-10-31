#!/bin/bash
#SBATCH --job-name=cshell_bench0
#SBATCH --gpus=1
#SBATCH --time=03:00:00
#SBATCH --array=0-2%3
#SBATCH --output=logs/%x_%A_%a.out
set -euo pipefail
source cshell/cluster/env.sh; trap 'rm -rf "$TMPDIR"' EXIT
PARTS=(cantilever l_bracket mbb)
P=${PARTS[$SLURM_ARRAY_TASK_ID]}
nvidia-smi --query-gpu=name,memory.total --format=csv,noheader
python docs/cshell/research/stage0/bench_solid.py examples/$P.json docs/cshell/research/stage0/bench_${P}${TAG:-}.json ${HS:-0.8 0.4}
