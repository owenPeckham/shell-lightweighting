#!/bin/bash
# Launch a full part on the login node: one run.sh job per start (24 h allocations, auto-resumed); collect.sh is
# submitted by the last start to finish.  Usage: continuous/cluster/launch_part.sh <cfg.json> [start ...] [-- run args]
# Run from the repository (or worktree) root. Uses one queue slot per start.
set -euo pipefail
CFG=$1; shift
STARTS=(); while [ $# -gt 0 ] && [ "$1" != "--" ]; do STARTS+=("$1"); shift; done; [ "${1:-}" = "--" ] && shift
if [ ${#STARTS[@]} -eq 0 ]; then
  STARTS=($(python -c "import sys; from cshell import config; print(' '.join(config.load(sys.argv[1])['optimiser']['starts']))" "$CFG"))
fi
mkdir -p logs
for S in "${STARTS[@]}"; do
  sbatch --time=${CSHELL_TIME:-24:00:00} --job-name=cshell_${S} continuous/cluster/run.sh "$CFG" --start "$S" "$@"
done
