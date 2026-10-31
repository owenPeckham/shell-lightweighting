#!/bin/bash
#SBATCH --job-name=cshell_eig
#SBATCH --gpus=1
#SBATCH --mem=48G
#SBATCH --time=00:30:00
#SBATCH --output=logs/%x_%j.out
set -euo pipefail
source cshell/cluster/env.sh; trap 'rm -rf "$TMPDIR"' EXIT
export CSHELL_EIG_VERBOSE=1
python docs/cshell/research/stage0/bench_solid.py examples/dshell/cantilever.json docs/cshell/research/stage0/eig_cantilever.json 0.8
