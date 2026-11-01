#!/bin/bash
# Launch a full part on the login node: one run.sh job (24 h allocations, auto-resumed) that optimises from the
# config's start(s) and then finalises in the same job (PLAN_single_run §3: no per-start fan-out, no collect job).
#   Usage: cshell/cluster/launch_part.sh <cfg.json> [run args, e.g. --set k=v ...]
# Run from the repository (or worktree) root. Uses one queue slot. Nothing runs on the login node (no python here).
set -euo pipefail
CFG=$1; shift
mkdir -p logs
sbatch --time=${CSHELL_TIME:-24:00:00} --job-name=cshell_$(basename "$CFG" .json) cshell/cluster/run.sh "$CFG" "$@"
