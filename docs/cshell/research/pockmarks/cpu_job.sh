#!/bin/bash
# Throwaway diag job (CPU). Usage: sbatch [--mem=..] docs/cshell/research/pockmarks/cpu_job.sh script.py args...
#SBATCH --job-name=pockmarks_cpu
#SBATCH --cpus-per-task=72
#SBATCH --mem=96G
#SBATCH --time=01:00:00
#SBATCH --output=logs/%x_%j.out
set -euo pipefail
source cshell/cluster/env.sh; trap 'rm -rf "$TMPDIR"' EXIT
export CUDA_VISIBLE_DEVICES=""
python "$@"
