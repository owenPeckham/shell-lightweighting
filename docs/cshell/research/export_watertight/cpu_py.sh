#!/bin/bash
# CPU python script job (from the repo root):
#   sbatch [--mem --time --cpus-per-task --job-name] docs/cshell/research/export_watertight/cpu_py.sh <script.py> [args]
#SBATCH --job-name=ew_cpu
#SBATCH --cpus-per-task=8
#SBATCH --mem=64G
#SBATCH --time=00:30:00
#SBATCH --output=docs/cshell/research/export_watertight/%x_%j.out
set -euo pipefail
source cshell/cluster/env.sh; trap 'rm -rf "$TMPDIR"' EXIT
export CUDA_VISIBLE_DEVICES=""
python "$@"
