#!/bin/bash
#SBATCH --job-name=b2_holes
#SBATCH --cpus-per-task=8
#SBATCH --mem=8G
#SBATCH --time=00:20:00
#SBATCH --output=continuous/research/b2_bodyfit/work/%x_%j.out
set -euo pipefail
source continuous/cluster/env.sh; trap 'rm -rf "$TMPDIR"' EXIT
python continuous/research/b2_bodyfit/holes_check.py runs_cshell/C_bench/B2_slender
