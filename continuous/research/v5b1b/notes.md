# V5-B1b (stub + slender rod with fillet)
Geometry (`examples/cshell/bench/make_bench.py B1b | B1b_root`): slender r 5, L 80; stub R 10, L 10 (L/D 0.5); fillet 3 mm at the
junction (2.5 mm failed: equals the 2.5 mm shoulder width); B2 load (20 N, -z), clamp at x = 0. Not specified in Appendix B: my choices.
A `B1b_stub_slender` (stub at load end, BENCH "B1b"), B `B1b_root_stub` (stub at the clamp; config only, run by `cshell run`).
Runs: `python -m cshell run ... --start uniform --set validation.body_fitted=false`, ~1 min each (jobs 7864232, 7864361);
collect OOM at 48 GB with export_h 0.05 (job 7864288, best-candidate export stage A) -> export_h 0.1 mm, 64 GB (7864592/3, 7.5 min).
`profile.py` (jobs 7864286 A, 7864594 B; B regions: "slender" row = stub x<6, "stub" row = slender x>14) gives regional closure.
