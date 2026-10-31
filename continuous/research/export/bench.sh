#!/bin/bash
# usage: sbatch [--mem=..] continuous/research/export/bench.sh CASE H_X [WORKERS]
#SBATCH --job-name=export_bench
#SBATCH --cpus-per-task=72
#SBATCH --mem=16G
#SBATCH --time=00:40:00
#SBATCH --output=logs/%x_%j.out
set -euo pipefail
source continuous/cluster/env.sh; trap 'rm -rf "$TMPDIR"' EXIT
export CUDA_VISIBLE_DEVICES=""
OUT=$SCRATCH/cshell_export_bench; mkdir -p $OUT
python continuous/research/export/run_export_bench.py "$1" "$2" $OUT "${3:-}" 
