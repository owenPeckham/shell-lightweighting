#!/bin/bash
# cshell run of one part: optimise, then finalise (shifted-grid re-score, GPU export, summary, report) in the same job,
# auto-resumed across allocations (THEORY §7.6, PLAN_single_run §3).
#   sbatch [--time=24:00:00] cshell/cluster/run.sh <cfg.json> [cshell run args]
#   sbatch ... cshell/cluster/run.sh --resume <run_dir>            (what the chain submits itself)
# --max-hours defaults to the job's time limit minus CSHELL_MARGIN_H (0.4 h). When the run pauses (exit 3) or the job
# is signalled (TIMEOUT/preemption: SIGTERM), the job submits its continuation with --dependency=afterany:<this job>
# (one queued job per part at any time). The allocation in which the optimiser finishes also finalises (`cshell run`
# does it in-process): there is no collect job. CSHELL_CHAIN_MAX (default 8) caps the number of allocations.
#SBATCH --job-name=cshell_run
#SBATCH --gpus=1
#SBATCH --time=02:00:00
#SBATCH --output=logs/%x_%j.out
set -uo pipefail
HERE=$(pwd)
source cshell/cluster/env.sh; trap 'rm -rf "$TMPDIR"' EXIT
nvidia-smi --query-gpu=name,memory.total --format=csv,noheader
ARGS=("$@")
START=""; for ((i=0; i<${#ARGS[@]}; i++)); do [ "${ARGS[$i]}" = "--start" ] && START=${ARGS[$((i+1))]}; done
HAS_MAXH=0; for x in "${ARGS[@]}"; do [ "$x" = "--max-hours" ] && HAS_MAXH=1; done
if [ $HAS_MAXH = 0 ] && [ -n "${SLURM_JOB_ID:-}" ]; then
  TL=$(squeue -h -j "$SLURM_JOB_ID" -o %l)            # [d-]hh:mm:ss
  MAXH=$(python - "$TL" "${CSHELL_MARGIN_H:-0.4}" <<'PY'
import sys
t, m = sys.argv[1], float(sys.argv[2])
d, _, hms = t.rpartition("-")
p = [int(x) for x in hms.split(":")]
while len(p) < 3: p.insert(0, 0)
h = (int(d) if d else 0) * 24 + p[0] + p[1] / 60 + p[2] / 3600
print(max(h - m, 0.25 * h))
PY
)
  ARGS+=(--max-hours "$MAXH")
fi
export CSHELL_RUNDIR_FILE=$TMPDIR/run_dir
CHAIN=${CSHELL_CHAIN:-1}
resubmit() {
  local rd; rd=$(cat "$CSHELL_RUNDIR_FILE" 2>/dev/null) || return 0
  if [ "$CHAIN" -ge "${CSHELL_CHAIN_MAX:-8}" ]; then echo "[run.sh] chain limit reached ($CHAIN)"; return 0; fi
  local extra=(); [ -n "$START" ] && extra=(--start "$START")
  cd "$HERE"
  CSHELL_CHAIN=$((CHAIN + 1)) sbatch --export=ALL,CSHELL_CHAIN=$((CHAIN + 1)) --time="$(squeue -h -j "$SLURM_JOB_ID" -o %l)" \
    --dependency=afterany:"$SLURM_JOB_ID" cshell/cluster/run.sh --resume "$rd" "${extra[@]}"
}
on_term() { echo "[run.sh] signalled: submitting continuation"; resubmit; exit 143; }
trap on_term TERM
python -m cshell run "${ARGS[@]}" &
PID=$!; wait $PID; RC=$?
if [ $RC -eq 3 ]; then echo "[run.sh] paused (max-hours): submitting continuation"; resubmit; exit 0; fi
if [ $RC -ne 0 ]; then echo "[run.sh] cshell run failed (rc=$RC): no continuation"; exit $RC; fi
RD=$(cat "$CSHELL_RUNDIR_FILE" 2>/dev/null || true)
[ -n "$RD" ] && [ -f "$RD/summary.txt" ] && head -n 4 "$RD/summary.txt"
exit 0
