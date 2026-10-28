#!/bin/bash
#SBATCH --job-name=cshell_tubes
#SBATCH --gpus=1
#SBATCH --time=00:20:00

#SBATCH --output=logs/%x_%j.out
set -euo pipefail
source continuous/cluster/env.sh; trap 'rm -rf "$TMPDIR"' EXIT
python continuous/research/stage0/tubes.py continuous/research/stage0/tubes_patch2.json 0.8 0.4
