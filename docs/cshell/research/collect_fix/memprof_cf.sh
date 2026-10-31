#!/bin/bash
# P5 acceptance: host memory of collect outside the export on the MBB smoke (h_g 0.1), STLs of SMOKE_mbb re-used
# (--skip-export), best forced through the export gate so its 94 M-triangle STL is re-scored too (--force-ok).
# Submit from the repo root:  sbatch continuous/research/collect_fix/memprof_cf.sh
#SBATCH --job-name=cf_memprof
#SBATCH --gpus=1
#SBATCH --mem=110G
#SBATCH --time=01:30:00
#SBATCH --output=continuous/research/collect_fix/%x_%j.out
set -euo pipefail
source continuous/cluster/env.sh; trap 'rm -rf "$TMPDIR"' EXIT
RD=runs_cshell/CF_mem
rm -rf $RD; mkdir -p $RD
rsync -a --exclude='*.stl' --exclude='export_baselines' --exclude='*_export.json' --exclude='collect_log.txt' \
  runs_cshell/CF_smoke/ $RD/
rm -f $RD/summary.json $RD/sensitivity.csv $RD/validation.csv $RD/*.vtp $RD/*.npz
/usr/bin/time -v python continuous/research/collect_fix/memprof_cf.py $RD --skip-export runs_cshell/SMOKE_mbb/MBB --force-ok
