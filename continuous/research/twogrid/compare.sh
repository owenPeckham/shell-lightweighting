#!/bin/bash
# dJ / J comparison of bench_design.py outputs (two-grid vs direct), CPU only.
#SBATCH --job-name=tg_compare
#SBATCH --cpus-per-task=72
#SBATCH --mem=16G
#SBATCH --time=00:20:00
#SBATCH --output=logs/%x_%j.out
set -euo pipefail
source continuous/cluster/env.sh; trap 'rm -rf "$TMPDIR"' EXIT
export CUDA_VISIBLE_DEVICES=""
B=runs_cshell/TG_bench
for p in mbb cant; do for ph in 0 1; do
  echo "== $p phase $ph"; python continuous/research/twogrid/bench_design.py --compare $B/${p}_directM_ph$ph.npz $B/${p}_tgG_ph$ph.npz
done; done
