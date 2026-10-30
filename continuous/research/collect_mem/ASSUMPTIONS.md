# collect_mem diagnosis: inputs and assumptions

- Run: runs_cshell/SMOKE_mbb/MBB (MBB, 3 starts, max_iter [2,3]); best start "uniform", D_comp 4.0 mm,
  last phase h_e 0.4 mm, h_g 0.1 mm (band 281.9 M cells, element box 772x297x147). Units mm, N, MPa.
- Code state: branch cshell-stage2 working tree at 2026-10-30 (other agents editing collect/validate/export;
  the jobs imported whatever was on disk at start).
- diag.py (job 7862848, 1 GH200, 14 min): uniform shell built on the band analytically, psi = max(phi - t, -phi),
  t = 2.0497 - 0.00068 mm (export_baselines' definition and t), instead of from uniform_shell.stl; its scores match
  collect's STL-based ones to 4 digits (K/K0 0.0541, P/P0 1.7303), so the STL path is not the cause.
- Load-node isolation: a load node is "isolated" if all 8 adjacent elements have rho below a threshold (1e-3).
- memprof.py: sampler every 2 s (parent RSS, PSS of parent+children, Slurm cgroup anon/current). Job 7862907
  (full collect, --mem=200G) was OOM-killed in export_design; job 7864025 re-runs collect with the export replaced
  by the source run's STLs/jsons (export memory taken from the export's own *_export.json peak_rss fields).
