# DRAFT PLAN section: collect re-scoring correctness and host memory (MBB at h_g 0.1)

Status: draft for agreement (diagnosis by the collect_mem analyst, 2026-10-30). Evidence: jobs 7862848 (diag),
7862907 / 7864025 / 7864331 (memory profiles); scripts and logs in docs/cshell/research/collect_mem/.

## Findings (evidence)

F1. CPU and GPU geometry chains agree. Same xi on both chains gives max|rho_cpu - rho_gpu| = 7.8e-16 and the same
    J to 2e-10 (best: J 0.30553, K/K0 0.93961, P/P0 1.96080; V/V0-1 -3.2e-6). Hypothesis (a) is ruled out. The
    re-scoring-grid hypothesis (c) is also ruled out: on the optimisation grid, collect's parametric re-score of the
    best design equals the diag value exactly.

F2. Uniform shell K/K0 0.054 is caused by the BCs together with the as-built bias, not by the design. The baseline
    psi = max(phi - t, -phi) puts the original load face on the inner skin. The as-built void-branch bias b_v = 0.2 mm
    then moves the material 0.2 mm outward from that face. The load rule (nodes within 0.75 h_e = 0.3 mm of the face,
    either side) keeps nodes on the void side, so 1844 of 5304 load nodes, carrying 34.8 % of the force, have all
    8 adjacent elements at rho < 1e-3. The result is C = 1583 N mm against 85.7 N mm for the solid.
      - nominal (no as-built): K/K0 0.760, no isolated load nodes
      - as-built with the mount cells kept solid: K/K0 0.683 (P/P0 0.634)
      - as-built on the 1/2 h_e shifted grid: K/K0 0.673 (load nodes are no longer isolated)
    The 0.054 is therefore a grid-alignment artefact (MBB load face normal to z, aligned with the grid). The run's own
    uniform start is unaffected because mounts and load faces are passive there (kept solid).

F3. The best design's V/V0-1 = +14.5 % comes from the export, not from the re-scoring. The optimiser's nominal
    V/V0-1 is -3e-6. The exported STL holds 97 586 mm^3 (+14.5 %):
      - +3.79 % (+3 234 mm^3) from the closure rule. Passive vertices (xt_a = 1 >= 0.98) get a = A_CLOSED = 1e4,
        and linear barycentric interpolation of 1e4 closes every cell whose closest triangle touches a passive
        vertex, so the closed region grows by one triangle ring along every mount/thin boundary (44 034 passive of
        118 478 vertices). The diag reproduces +3.79 % on the band with 0 active vertices above the threshold.
      - +11.1 % (+9 487 mm^3) from the void_min = 3 mm opening. 4 292 open active vertices (3.6 % of the active area)
        have void width 2(r_MA - a) < 3 mm after 2-3 iterations of a smoke run.
      - +0.07 % from enclosed-void filling; escape holes give a small negative contribution.
    The mass correction then fails to bracket at ±0.35 mm (f(-0.35) = +0.14 %). export ok = False, and collect still
    re-scores and reports the STL at +14.5 %, so its K/K0 1.018 and P/P0 2.01 describe a heavier part.

F4. PCG stall (rel. residual 1.92e-8 against tol 1e-10, maxiter 100) is the best STL on the shifted grid
    (raw_responses #7). The parametric best on the same shifted grid converges in 72 iterations. The STL has
    39 bodies and 1 890 duplicated directed edges (watertight False), so floating islands held only by rho_min give
    near-singular modes. The uniform-start xi also stalls at 8.5e-10 after 100 iterations on both chains.
    maxiter 100 with tol 1e-10 has no headroom.

F5. Host memory at h_g 0.1 (cgroup anon peak per stage; parent RSS in brackets):
      Context 1 GiB | CPU-chain PhaseModel 49 (38-40 resident: phi f64, P and P^T CSR, elem int64 over 282 M cells)
      | CellData +4 (copies of phi and elem) | parametric FE 52 | export_baselines 104 (parent 86; +17 GiB stays
      resident afterwards through the module-global export._G) | export_design 225 GiB (parent RSS 120, of
      which 70 is collect state inherited by 72 forked workers, PSS 189; leaves +22 GiB resident) — the OOM cause at 115/200G | trimesh.load(process=True) of each 4.6 GB STL: peak 98 / 117 (+48 GiB transient,
      +9 GiB resident each; both meshes kept) | mesh_phi +2, asbuilt_increment +5 | second PhaseModel (shift) 95,
      because the first model's host memory is not returned (55 -> 87 GiB). Collect's peak is export_design (225 GiB cgroup); outside the export, the STL load (117) and the stacked PhaseModels (95) dominate.

## Fixes (files, approach, acceptance)

P1. Re-scoring BCs consistent with the design (collect.py; fea.py read-only reuse). For every re-scored design, load
    and clamp only nodes that touch material (all 8 adjacent elements rho >= 1e-3 counted as material), and report the
    share of force dropped and redistributed. Alternative, closer to physics: re-select the load and support nodes on
    the design's own surface, nearest-material within 0.75 h_e of the face along its normal. In addition, keep mount
    faces (load/support) solid in the uniform-shell baseline, i.e. psi = phi on the mount cells, as the optimiser does.
    That is a spec change to §7.4 S1 and needs agreement.
    Acceptance: uniform shell on opt and shift grids agrees within 5 % in K/K0. No isolated loaded node in any re-scored
    design. MBB smoke uniform shell K/K0 in [0.6, 0.8] on both grids.

P2. Export closure rule (export.py; owner: export agent). Interpolate the closure indicator, not A_CLOSED: close a
    cell where the barycentric mean of `closed` >= 0.5, or use a = A_CLOSED only when all 3 vertices are closed.
    Passive vertices should give the optimiser's a_hi, not A_CLOSED, unless they are closed in the optimiser sense.
    Acceptance: the band volume of the export rule equals the optimiser's nominal V within 0.2 % for the smoke best
    (now +3.79 %). B2 unchanged within 0.2 %.

P3. Make the mass mismatch blocking and visible (collect.py). When export ok is False or |V_STL/V0 - 1| > 2e-3, label
    the re-score "not mass-matched" and do not use it for the decision except as a reject reason (decide() already
    receives export_ok; make sure summary/report show it). The +11 % opening fill on unconverged designs is expected
    behaviour of §8.2/§8.4. For full runs, check G_s at the end. If the opening fill exceeds the ±0.3 mm correction
    capacity, report it as an optimiser non-convergence, not a re-scoring error.
    Acceptance: smoke collect reports export failure as the decision reason. No silent +14.5 % row.

P4. PCG robustness (fe_twogrid.py, config fe.twogrid.maxiter). For re-scoring: (i) drop material components
    not connected to a clamped node before the FE (the rho of islands goes to 0); (ii) maxiter 100 -> 300 in collect,
    and on stagnation (residual reduction < 2x over 20 iterations) accept rel. residual <= 1e-7 with a logged warning
    and record it in the row (compliance error is about the residual squared, so negligible at 1e-7). Do not raise.
    Acceptance: smoke collect completes all three resolutions. Converged rows unchanged to 1e-8.

P5. Host memory <= 60 GiB outside the export (collect.py, optimise.py read-only use):
    a. GPU geometry chain in collect (F1 shows it is exact). Do the re-scoring cell maps on the device: P v as a
       gather over the compact (face, bary) arrays already in DeviceDisc; mesh_phi already runs in Warp. Drop CellData's
       host phi/P/elem copies. Expected: PhaseModel host about 7 GiB (pinned band) instead of 38-49 GiB.
    b. Release each PhaseModel's host arrays before building the next (explicit del of band/disc/dm, gc.collect,
       malloc_trim(0) via ctypes), and assert RSS returns within 5 GiB.
    c. STL ingestion: trimesh.load(process=False) and a vectorised vertex weld (np.unique on quantised coordinates, or
       pass the exporter's in-memory V, F directly). Keep one mesh at a time: compute phiE for every grid when the mesh
       is loaded, or store phiE (float32, 1.1 GiB per grid) instead of the mesh.
    d. Clear export._G after each Exporter.run (owner: export agent). Run the export in a subprocess, or before the
       PhaseModel is built, so its 120-150 GiB peak does not stack on collect's state.
    Acceptance (MBB smoke, h_g 0.1): peak cgroup anon outside the export <= 60 GiB, measured with
    docs/cshell/research/collect_mem/memprof.py --skip-export. Full collect fits --mem=180G (export-bound). Scores
    unchanged against the CPU path to 1e-10 in J.

## Open questions
- Uniform-shell baseline definition (§7.4 S1): keep mounts solid as the optimiser does, or keep the pure shell and fix
  only the BC node selection? This changes the reported baseline.
- Should the decision reject at the export ok = False stage before re-scoring, to save about 20 minutes and 100 GiB?
