# Lightweighting demo: assumptions and how to re-run

- cshell defaults throughout (objective.type = lightweight, margins 1.02, PA11 E 1760 MPa, sigma_f 36.6 MPa, t_min 1.5 mm,
  void_min 3 mm, outward growth allowed). The only config change from the examples is `mounts.auto_analytic = false`:
  CAD-generated solids have only analytic faces, and auto_analytic would make every face a mount (bell crank v1 smoke
  failed with "stress p-norm: no element carries weight"). The supports and loads are the only mounts.
- Load magnitudes (500 N beams; 2 x 400 N bell crank) do not affect the ratios K/K0 and P/P0 (linear model).
- Supports are fully clamped faces (beam root face; bell-crank pivot bore, which therefore carries a moment).
- Why beams in bending: mass sits in a low-stress core near the neutral axis. Moving the skin outward (allowed) gives
  back the second moment of area lost to the hollow core. Solid lambda_b0/lambda_s0 is about 10, so buckling is far
  from governing. A thin-walled box of outer side B, wall t ~ B/12 (plate local-buckling limit at sigma_f) and
  Z >= Z0 gives B ~ 1.15 b0 and A ~ 0.44 b0^2, i.e. a ~55 % saving in mid-span; the closed ends reduce it.
- Re-run (from the worktree root, via sbatch only):
  `sbatch --gpus=1 --mem=48G --time=01:15:00 examples/cshell/lightweight_demo/run_demo.sh <b_y> <h_z> <L> <F>`
  `BC_RP=20 BC_RF=12 sbatch --gpus=1 --mem=64G --time=01:15:00 examples/cshell/lightweight_demo/run_bellcrank.sh 24 20 85 65 400`
- Not checked: body-fitted validation, sensitivity sweeps, mesh dependence beyond the cshell two-grid re-score.
