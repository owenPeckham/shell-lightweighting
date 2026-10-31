#!/bin/bash
#SBATCH --job-name=b2_ccxchk
#SBATCH --cpus-per-task=72
#SBATCH --mem=56G
#SBATCH --time=00:50:00
#SBATCH --output=docs/cshell/research/b2_bodyfit/work/%x_%j.out
set -euo pipefail
source cshell/cluster/env.sh; trap 'rm -rf "$TMPDIR"' EXIT
python docs/cshell/research/b2_bodyfit/ccx_modes.py docs/cshell/research/b2_bodyfit/work/m2_design_h0.5.inp m2d05_t72_m4 72 4 0 0
python docs/cshell/research/b2_bodyfit/ccx_modes.py docs/cshell/research/b2_bodyfit/work/m2_design_h0.5.inp m2d05_t16_m4 16 4 0 0
python docs/cshell/research/b2_bodyfit/ccx_modes.py docs/cshell/research/b2_bodyfit/work/m2_design_h0.5.inp m2d05_t72_m1 72 1 0 0
