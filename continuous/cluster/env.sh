# Source in every cshell job: conda env, scratch TMPDIR, FE scratch. THEORY §10.4.
export CSHELL_ENV=/projects/b6bl/owenpeckham.b6bl/home_data/shell-lightweighting/envs/cshell
export PATH=$CSHELL_ENV/bin:$PATH
export TMPDIR=$SCRATCH/tmp/${SLURM_JOB_ID:-login}; mkdir -p "$TMPDIR"
export CSHELL_FE_SCRATCH=$TMPDIR/fe
export NUMBA_CACHE_DIR=$SCRATCH/numba_cache/cshell
export PYTHONUNBUFFERED=1
cd $HOME/shell-lightweighting
