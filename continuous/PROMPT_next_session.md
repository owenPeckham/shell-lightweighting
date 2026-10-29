Continue building the continuous constant-mass shelling optimiser `cshell` (Isambard-AI).

Work in the worktree `/home/b6bl/owenpeckham.b6bl/shell-lightweighting/.claude/worktrees/cshell-stage2` on branch
`cshell-stage2` (enter it with EnterWorktree using that path; don't create a new one). Read first:
`continuous/HANDOVER.md` (state, my decisions, rules, remaining work), then `continuous/STATUS.md`, `continuous/VALIDATION.md`,
`continuous/BENCHMARKS.md`, `continuous/plans/PLAN_twogrid.md`, `continuous/plans/PLAN_validation_fixes.md` and the two
status notes they point to. The last commit is a WIP commit with two unfinished developer tasks; review it before
building on it, and check the logs of the jobs listed in the hand-over.

Rules: nothing runs on the login node (not even quick checks) — everything via sbatch, and pass this into every
subagent brief; the optimiser's FE must use minimum GPU and host memory, be fast and accurate, on one GPU (multi-GPU
only for validation); slot cap 8; plan before implementing.

Next, in order: (1) finish and verify the two-grid PCG engine to its plan's acceptance criteria and commit;
(2) smoke-test, then launch the full runs (3 parts × 3 starts) with auto-resume; (3) finish the body-fitted validation
fixes and re-run collect on B2; (4) the remaining Stage 2 gate items in the hand-over. Ask me only for decisions that
would waste GPU-days or can't be undone; report what you did, where it lives, and the next command.
