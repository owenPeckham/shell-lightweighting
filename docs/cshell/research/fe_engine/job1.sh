#!/bin/bash
#SBATCH --job-name=cudss2
#SBATCH --gpus=1
#SBATCH --time=00:30:00
#SBATCH --output=/home/b6bl/owenpeckham.b6bl/shell-lightweighting/docs/cshell/research/fe_engine/cudss2_%j.out
export TMPDIR=$SCRATCH/tmp/$SLURM_JOB_ID; mkdir -p $TMPDIR
SA=~/.claude/skills/structural-analysis/scripts/sa
D=/home/b6bl/owenpeckham.b6bl/shell-lightweighting/docs/cshell/research/fe_engine
nvidia-smi --query-gpu=name,memory.total --format=csv
for a in "200 4" "300 4" "400 4"; do echo "=== $a"; timeout 500 $SA gpu python $D/cudss_shell_bench.py $a 2>&1 | grep -v -E "INFO|WARNING"; done
