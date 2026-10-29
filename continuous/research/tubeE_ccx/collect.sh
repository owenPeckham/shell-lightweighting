#!/bin/bash
# Aggregate runs/*/result.json -> summary.json (tiny; run via sbatch, never on the login node).
#SBATCH --job-name=tubeccx_collect
#SBATCH --cpus-per-task=1
#SBATCH --mem=1G
#SBATCH --time=00:05:00
#SBATCH --output=/home/b6bl/owenpeckham.b6bl/shell-lightweighting/.claude/worktrees/cshell-stage2/continuous/research/tubeE_ccx/logs/%x_%j.out
source /home/b6bl/owenpeckham.b6bl/shell-lightweighting/.claude/worktrees/cshell-stage2/continuous/cluster/env.sh; trap 'rm -rf "$TMPDIR"' EXIT
python /home/b6bl/owenpeckham.b6bl/shell-lightweighting/.claude/worktrees/cshell-stage2/continuous/research/tubeE_ccx/collect.py
