#!/bin/bash
# Generate the bell-crank STEP/config (make_bellcrank.py), then run cshell (lightweight default). From the worktree root:
#   sbatch --gpus=1 --mem=48G --time=01:15:00 examples/cshell/lightweight_demo/run_bellcrank.sh <W> <T> <L1> <L2> <F> [cshell run args]
# Full runs hand over to cshell/cluster/run.sh (auto-resumed chain); 6th argument --direct: one python -m cshell run (smoke).
#SBATCH --job-name=cshell_lwcrank
#SBATCH --gpus=1
#SBATCH --output=logs/%x_%j.out
set -uo pipefail
source cshell/cluster/env.sh
W=$1; T=$2; L1=$3; L2=$4; F=$5; shift 5
GEO=(); SUF=""; if [ -n "${BC_RP:-}" ]; then GEO=("$BC_RP" "$BC_RF"); SUF=_P${BC_RP%.*}F${BC_RF%.*}; fi   # optional env BC_RP, BC_RF
python examples/cshell/lightweight_demo/make_bellcrank.py "$W" "$T" "$L1" "$L2" "$F" "${GEO[@]}" || exit 1
CFG=examples/cshell/lightweight_demo/BellCrank_${W%.*}x${T%.*}$SUF.json
if [ "${1:-}" = "--direct" ]; then
  shift; trap 'rm -rf "$TMPDIR"' EXIT
  T=; [ -x /usr/bin/time ] && T="/usr/bin/time -v"
  $T python -m cshell run "$CFG" "$@"
else
  rm -rf "$TMPDIR"
  exec bash cshell/cluster/run.sh "$CFG" "$@"
fi
