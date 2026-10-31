#!/bin/bash
#SBATCH --job-name=tubepw
#SBATCH --gpus=1
#SBATCH --time=00:40:00
#SBATCH --output=/home/b6bl/owenpeckham.b6bl/shell-lightweighting/.claude/worktrees/cshell-stage2/docs/cshell/research/tubeE_ccx/logs/%x_%j.out
set -euo pipefail
source /home/b6bl/owenpeckham.b6bl/shell-lightweighting/.claude/worktrees/cshell-stage2/cshell/cluster/env.sh; trap 'rm -rf "$TMPDIR"' EXIT
D=/home/b6bl/owenpeckham.b6bl/shell-lightweighting/.claude/worktrees/cshell-stage2/docs/cshell/research/tubeE_ccx
python $D/pw_voxel.py $D/pw_voxel.json
