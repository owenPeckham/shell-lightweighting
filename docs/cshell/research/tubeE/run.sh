#!/bin/bash
#SBATCH --job-name=tubeE
#SBATCH --gpus=1
#SBATCH --time=01:00:00
#SBATCH --output=/home/b6bl/owenpeckham.b6bl/shell-lightweighting/.claude/worktrees/cshell-stage2/docs/cshell/research/tubeE/logs/%x_%j.out
set -euo pipefail
source /home/b6bl/owenpeckham.b6bl/shell-lightweighting/.claude/worktrees/cshell-stage2/cshell/cluster/env.sh; trap 'rm -rf "$TMPDIR"' EXIT
cd /home/b6bl/owenpeckham.b6bl/shell-lightweighting/.claude/worktrees/cshell-stage2; python docs/cshell/research/tubeE/tubes_pw.py docs/cshell/research/tubeE/tubeE.json
