#!/bin/bash
#SBATCH --job-name=b2_gpudeck
#SBATCH --gpus=1
#SBATCH --mem=8G
#SBATCH --time=00:30:00
#SBATCH --output=docs/cshell/research/b2_bodyfit/work/%x_%j.out
set -euo pipefail
source cshell/cluster/env.sh; trap 'rm -rf "$TMPDIR"' EXIT
nvidia-smi --query-gpu=name,memory.total --format=csv,noheader
python docs/cshell/research/b2_bodyfit/bf_gpu.py runs_cshell/C_bench/B2_slender/validation/solid.inp 4 > docs/cshell/research/b2_bodyfit/work/gpu_solid_deck.json
cat docs/cshell/research/b2_bodyfit/work/gpu_solid_deck.json
python docs/cshell/research/b2_bodyfit/bf_gpu.py runs_cshell/C_bench/B2_slender/validation/design.inp 4 > docs/cshell/research/b2_bodyfit/work/gpu_design_deck.json
cat docs/cshell/research/b2_bodyfit/work/gpu_design_deck.json
