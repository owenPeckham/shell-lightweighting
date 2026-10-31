#!/bin/bash
# GPU python script job (from the repo root):
#   sbatch [--mem --time --job-name] docs/cshell/research/bodyfit_scale/gpu_py.sh <script.py> [args]
#SBATCH --job-name=bfs_gpu
#SBATCH --gpus=1
#SBATCH --mem=100G
#SBATCH --time=02:00:00
#SBATCH --output=docs/cshell/research/bodyfit_scale/%x_%j.out
set -euo pipefail
source cshell/cluster/env.sh; trap 'rm -rf "$TMPDIR"' EXIT
nvidia-smi --query-gpu=name,memory.total --format=csv,noheader
/usr/bin/time -v python "$@" 2> >(tee "$TMPDIR/time.txt" >&2)
grep -E "Maximum resident|Elapsed" "$TMPDIR/time.txt" || true
