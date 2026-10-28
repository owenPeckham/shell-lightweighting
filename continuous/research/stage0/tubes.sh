#!/bin/bash
#SBATCH --job-name=cshell_tubes
#SBATCH --gpus=1
#SBATCH --time=01:30:00
#SBATCH --mem=220G
#SBATCH --output=logs/%x_%j.out
set -euo pipefail
source continuous/cluster/env.sh; trap 'rm -rf "$TMPDIR"' EXIT
python continuous/research/stage0/tubes.py continuous/research/stage0/tubes_patch.json 0.8 0.4 0.2
