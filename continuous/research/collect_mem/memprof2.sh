#!/bin/bash
# Per-stage host-memory profile of cshell collect on a copy of SMOKE_mbb (runs_cshell/CM_mem/MBB). Submit from the repo root.
#SBATCH --job-name=cm_memprof2
#SBATCH --gpus=1
#SBATCH --mem=180G
#SBATCH --time=01:30:00
#SBATCH --output=continuous/research/collect_mem/%x_%j.out
set -euo pipefail
source continuous/cluster/env.sh; trap 'rm -rf "$TMPDIR"' EXIT
/usr/bin/time -v python continuous/research/collect_mem/memprof.py runs_cshell/CM_mem2/MBB --skip-export runs_cshell/SMOKE_mbb/MBB
