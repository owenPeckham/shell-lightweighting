#!/bin/bash
# submit_stuff.sh <which: solid|design|both> <solve|nosolve> [parts...]  (from the repo root; parts MBB Cantilever L-bracket B2)
set -euo pipefail
REPO=$(pwd)
WHICH=$1; SOLVE=$2; shift 2
PARTS=${*:-MBB Cantilever L-bracket}
[ "$WHICH" = both ] && WHICH="solid design"
for p in $PARTS; do
  d="$REPO/runs_cshell/FULL/$p"; [ "$p" = B2 ] && d="$REPO/runs_cshell/VF_B2"
  for w in $WHICH; do
    sbatch --parsable --time=03:00:00 --mem=150G --job-name=bst_${p}_$w \
      continuous/research/bodyfit_scale/gpu_py.sh continuous/research/bodyfit_scale/stuff_check.py "$d" $w $SOLVE
  done
done
