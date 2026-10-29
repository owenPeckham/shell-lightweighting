#!/bin/bash
#SBATCH --job-name=b2_ccxmodes
#SBATCH --cpus-per-task=72
#SBATCH --mem=64G
#SBATCH --time=00:45:00
#SBATCH --output=continuous/research/b2_bodyfit/work/%x_%j.out
set -euo pipefail
source continuous/cluster/env.sh; trap 'rm -rf "$TMPDIR"' EXIT
python continuous/research/b2_bodyfit/ccx_modes.py runs_cshell/C_bench/B2_slender/validation/design.inp design_m4_t72 72 4 1 0
python continuous/research/b2_bodyfit/ccx_modes.py runs_cshell/C_bench/B2_slender/validation/design.inp design_m1_t72 72 1 0 0
