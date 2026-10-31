#!/bin/bash
#SBATCH --job-name=b2_modeloc
#SBATCH --cpus-per-task=4
#SBATCH --mem=4G
#SBATCH --time=00:10:00
#SBATCH --output=continuous/research/b2_bodyfit/work/%x_%j.out
set -euo pipefail
source continuous/cluster/env.sh; trap 'rm -rf "$TMPDIR"' EXIT
python continuous/research/b2_bodyfit/mode_loc.py
