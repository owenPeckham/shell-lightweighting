#!/bin/bash
# Per-stage host-memory profile of cshell collect on a copy of SMOKE_mbb (runs_cshell/CM_mem/MBB). Submit from the repo root.
#SBATCH --job-name=cm_memprof
#SBATCH --gpus=1
#SBATCH --mem=200G
#SBATCH --time=02:30:00
#SBATCH --output=docs/cshell/research/collect_mem/%x_%j.out
set -euo pipefail
source cshell/cluster/env.sh; trap 'rm -rf "$TMPDIR"' EXIT
/usr/bin/time -v python docs/cshell/research/collect_mem/memprof.py runs_cshell/CM_mem/MBB
