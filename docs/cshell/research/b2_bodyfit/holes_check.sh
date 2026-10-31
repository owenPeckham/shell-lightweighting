#!/bin/bash
#SBATCH --job-name=b2_holes
#SBATCH --cpus-per-task=8
#SBATCH --mem=8G
#SBATCH --time=00:20:00
#SBATCH --output=docs/cshell/research/b2_bodyfit/work/%x_%j.out
set -euo pipefail
source cshell/cluster/env.sh; trap 'rm -rf "$TMPDIR"' EXIT
python docs/cshell/research/b2_bodyfit/holes_check.py runs_cshell/C_bench/B2_slender
