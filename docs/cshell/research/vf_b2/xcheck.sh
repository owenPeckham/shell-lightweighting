#!/bin/bash
#SBATCH --job-name=vf_b2_xcheck
#SBATCH --gpus=1
#SBATCH --mem=56G
#SBATCH --time=00:40:00
#SBATCH --output=logs/%x_%j.out
set -euo pipefail
source cshell/cluster/env.sh; trap 'rm -rf "$TMPDIR"' EXIT
python docs/cshell/research/vf_b2/xcheck.py runs_cshell/VF_B2/best_candidate.stl 0.5 72
