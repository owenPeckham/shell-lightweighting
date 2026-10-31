#!/bin/bash
# Submit diag.py: submit_diag.sh <variants> <which: solid|design|both> [parts...]  (from the repo root)
# parts: MBB Cantilever L-bracket (runs_cshell/FULL) and/or B2 (runs_cshell/VF_B2)
set -euo pipefail
REPO=$(pwd)
VARS=$1; WHICH=$2; shift 2
PARTS=${*:-MBB Cantilever L-bracket}
[ "$WHICH" = both ] && WHICH="solid design"
for p in $PARTS; do
  d="$REPO/runs_cshell/FULL/$p"; [ "$p" = B2 ] && d="$REPO/runs_cshell/VF_B2"
  for w in $WHICH; do
    sbatch --parsable --time=03:00:00 --mem=110G --job-name=bfs_${p}_$w \
      continuous/research/bodyfit_scale/gpu_py.sh continuous/research/bodyfit_scale/diag.py "$d" $w $VARS
  done
done
