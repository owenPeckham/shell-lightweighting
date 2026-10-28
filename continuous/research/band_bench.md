# Band build benchmark (cshell/band.py, THEORY §3.1, §4.2)

Part: `examples/step/Solid Cantilever.step`, tessellated at 0.02 mm (66,764 vertices, 133,540 triangles). Band: whole interior
plus outer shell D_comp = 4 mm (+ h_g). Backend: Warp `mesh_query_point_sign_winding_number` on one GH200 (95 GiB), 1 task per job.
Script `continuous/research/band/bench.py`, raw JSON `continuous/research/band/bench_h*_D4.json`, logs `band_gpu_<jobid>.out`.

| h_g (mm) | job | band cells | dense box cells | band size | build (s) | peak host RSS | peak GPU pool | V0 / 60,942 - 1 |
| --- | --- | --- | --- | --- | --- | --- | --- | --- |
| 0.4 | 7794635 | 3.67 M | 59.9 M | 98 MiB | 1.2 | 0.6 GiB | 87 MiB | +1.1e-3 |
| 0.2 | 7794662 | 28.2 M | 0.47 G | 753 MiB | 3.6 | 1.9 GiB | 172 MiB | +2.6e-4 |
| 0.1 | 7794676 | 220.4 M | 3.72 G | 5.75 GiB | 25.8 | 9.6 GiB | 309 MiB | +5.0e-5 |

- Band = 28 B/cell (int64 flat, float32 φ, int32 triangle, 3 x float32 barycentrics); 5.9 % of the dense box at h_g = 0.1.
- V0 is h_g^3 Σ H_ε(-φ), ε = h_g, float64 (0.1 mm: 7.6 s to sum 2.2e8 cells in numpy); the tessellated solid agrees with the §4.2 value to 5e-5.
- Cache `.npz` (uncompressed): 101 MB / 756 MB / 5.75 GiB; load 0.05 / 0.33 / 2.5 s, save 0.07 / 0.32 / 2.6 s.
- `extend(4 -> 8 mm)` (§3.5): h_g = 0.4: 1.0 s, 7.56 M cells; 0.2: 6.6 s, 58.8 M cells; 0.1: 47 s, 463 M cells, peak host RSS 38 GiB (merge temporaries; the new band itself is 12 GiB). Use `--mem >= 64G` for h_g = 0.1 extension.
- Memory is bounded by tiling: queries run in chunks of 2^23 points, one 32-cell i-slab at a time; the dense box is never allocated. GPU memory is a few hundred MiB; host peak is about 1.7x the final band.
- Not measured: the cost of `Band.P` (a second closest-point query of Γ_h per cell, with 3 nnz/row: 220 M cells -> ~7 GB CSR in float64); expect about the same time as the build.
