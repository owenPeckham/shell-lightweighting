#!/bin/bash
# Generate a demo beam STEP/config (make_demo.py), then run cshell on it (lightweight default). From the worktree root:
#   sbatch --gpus=1 --mem=32G --time=01:15:00 examples/cshell/lightweight_demo/run_demo.sh <b_y> <h_z> <L> <F> [cshell run args]
# Full runs hand over to cshell/cluster/run.sh (auto-resumed chain, finalise in the last allocation); pass --direct
# as the 5th argument for a single python -m cshell run (smoke tests, with --no-finalise etc.).
#SBATCH --job-name=cshell_lwdemo
#SBATCH --gpus=1
#SBATCH --output=logs/%x_%j.out
set -uo pipefail
source cshell/cluster/env.sh
BY=$1; HZ=$2; L=$3; F=$4; shift 4
python examples/cshell/lightweight_demo/make_demo.py "$BY" "$HZ" "$L" "$F" || exit 1
NAME=Beam_${BY%.*}x${HZ%.*}x${L%.*}
CFG=examples/cshell/lightweight_demo/$NAME.json
if [ "${1:-}" = "--direct" ]; then
  shift; trap 'rm -rf "$TMPDIR"' EXIT
  T=; [ -x /usr/bin/time ] && T="/usr/bin/time -v"
  $T python -m cshell run "$CFG" "$@"
else
  rm -rf "$TMPDIR"
  exec bash cshell/cluster/run.sh "$CFG" "$@"
fi
