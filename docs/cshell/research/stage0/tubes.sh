#!/bin/bash
#SBATCH --job-name=cshell_tubes
#SBATCH --gpus=1
#SBATCH --time=00:20:00

#SBATCH --output=logs/%x_%j.out
set -euo pipefail
source cshell/cluster/env.sh; trap 'rm -rf "$TMPDIR"' EXIT
python docs/cshell/research/stage0/tubes.py docs/cshell/research/stage0/tubes_patch2.json 0.8 0.4
