#!/bin/bash
# All Task E benchmarks in one job: sbatch continuous/research/export/bench_all.sh
#SBATCH --job-name=export_bench_all
#SBATCH --cpus-per-task=72
#SBATCH --mem=100G
#SBATCH --time=00:45:00
#SBATCH --output=logs/%x_%j.out
set -euo pipefail
source continuous/cluster/env.sh; trap 'rm -rf "$TMPDIR"' EXIT
export CUDA_VISIBLE_DEVICES=""
OUT=$SCRATCH/cshell_export_bench_final; mkdir -p $OUT
for a in "b2 0.1" "b2 0.05" "cantilever 0.1" "cantilever 0.05"; do
  echo "=== $a"; python continuous/research/export/run_export_bench.py $a $OUT
done
