#!/bin/bash
#SBATCH --job-name=tubeE
#SBATCH --gpus=1
#SBATCH --time=01:00:00
#SBATCH --output=/home/b6bl/owenpeckham.b6bl/shell-lightweighting/.claude/worktrees/cshell-stage2/continuous/research/tubeE/logs/%x_%j.out
set -euo pipefail
source /home/b6bl/owenpeckham.b6bl/shell-lightweighting/.claude/worktrees/cshell-stage2/continuous/cluster/env.sh; trap 'rm -rf "$TMPDIR"' EXIT
cd /home/b6bl/owenpeckham.b6bl/shell-lightweighting/.claude/worktrees/cshell-stage2; python continuous/research/tubeE/tubes_pw.py continuous/research/tubeE/tubeE.json
