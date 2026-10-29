#!/bin/bash
# Re-run collect (new chain) on a copy of the B2 run. sbatch continuous/research/vf_b2/vf_collect.sh
#SBATCH --job-name=vf_b2_collect
#SBATCH --gpus=1
#SBATCH --mem=80G
#SBATCH --time=01:30:00
#SBATCH --output=logs/%x_%j.out
set -euo pipefail
source continuous/cluster/env.sh; trap 'rm -rf "$TMPDIR"' EXIT
nvidia-smi --query-gpu=name,memory.total --format=csv,noheader
[ -d runs_cshell/VF_B2 ] || cp -r runs_cshell/C_bench/B2_slender runs_cshell/VF_B2
rm -rf runs_cshell/VF_B2/validation
/usr/bin/time -v python -m cshell collect runs_cshell/VF_B2 2> logs/vf_b2_collect_time_$SLURM_JOB_ID.txt || { tail -30 logs/vf_b2_collect_time_$SLURM_JOB_ID.txt; exit 1; }
grep -E "Maximum resident|Elapsed" logs/vf_b2_collect_time_$SLURM_JOB_ID.txt
nvidia-smi --query-gpu=memory.used --format=csv,noheader
