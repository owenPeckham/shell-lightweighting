#!/bin/bash
# CPU python script job (from the repo root):
#   sbatch [--mem --time --cpus-per-task --job-name] continuous/research/export_watertight/cpu_py.sh <script.py> [args]
#SBATCH --job-name=ew_cpu
#SBATCH --cpus-per-task=8
#SBATCH --mem=64G
#SBATCH --time=00:30:00
#SBATCH --output=continuous/research/export_watertight/%x_%j.out
set -euo pipefail
source continuous/cluster/env.sh; trap 'rm -rf "$TMPDIR"' EXIT
export CUDA_VISIBLE_DEVICES=""
python "$@"
