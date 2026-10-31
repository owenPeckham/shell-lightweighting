#!/bin/bash
# Decimated preview STLs for one part (CPU job). From the repo root:
#   sbatch --mem=160G examples/cshell/runs/make_previews.sh <full_run_dir> <part>
# e.g. <full_run_dir> = /projects/b6bl/owenpeckham.b6bl/home_data/shell-lightweighting/runs_cshell/FULL/MBB
#SBATCH --job-name=cshell_previews
#SBATCH --cpus-per-task=72
#SBATCH --mem=160G
#SBATCH --time=03:00:00
#SBATCH --output=logs/%x_%j.out
set -euo pipefail
source cshell/cluster/env.sh; trap 'rm -rf "$TMPDIR"' EXIT
T=; [ -x /usr/bin/time ] && T="/usr/bin/time -v"
$T python examples/cshell/runs/make_previews.py "$1" "examples/cshell/runs/$2" "${@:3}"
