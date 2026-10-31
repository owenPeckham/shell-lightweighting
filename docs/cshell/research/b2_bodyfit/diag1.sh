#!/bin/bash
#SBATCH --job-name=b2_diag1
#SBATCH --cpus-per-task=72
#SBATCH --mem=6G
#SBATCH --time=00:40:00
#SBATCH --output=docs/cshell/research/b2_bodyfit/work/%x_%j.out
set -euo pipefail
source cshell/cluster/env.sh; trap 'rm -rf "$TMPDIR"' EXIT
export CUDA_VISIBLE_DEVICES=""
python docs/cshell/research/b2_bodyfit/diag1.py runs_cshell/C_bench/B2_slender
