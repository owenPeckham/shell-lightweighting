# continuous/reference

Read-only reference material for developing `cshell`, the continuous variable-thickness optimiser (see `../THEORY_continuous.md` and `../PROMPT_claude_code.md`).

**Do not edit these files.** Import the two FE modules from here: put this folder on `sys.path` inside `cshell/fea.py`. If you find a bug, work around it in `cshell` and record it in `../OPEN_QUESTIONS.md`.

| File | What it is | SHA-256 (first 16) |
| --- | --- | --- |
| `gpu_analysis.py` | Owen's GPU FE kernels (numba-cuda, optional JAX). Covers matrix-free PCG with block-Jacobi or geometric-multigrid preconditioning, element energy and stress kernels, element matrices (`q1_gauss`, `q1`, `q1_im`), engine selection and memory models. | 6f72b99601176798 |
| `analysis.py` | Owen's FE back ends `Conformal`, `Regular` and `ConjugateGradient`, with `run_analysis` and linear buckling (`run_buckling`, p-norm aggregated inverse eigenvalues with adjoint sensitivity). `ConjugateGradient` is the one to use. | 8a506ab7c20763a5 |
| `interfaces.py` | **Stub written for this hand-over.** The original module was not supplied. `analysis.py` imports only `BucklingResults` from it, and the stub reproduces that dataclass with every field `analysis.py` passes. | a5f82fad692d55d1 |
| `baseline_optimiser.py` | Owen's PCLS optimiser driver. Included only as the model for report and post-processing conventions; it imports modules that are not supplied, so it will not run. | f4f4693054704a9d |
| `main.pdf` | Manuscript: Peckham, Hicks, Goudswaard, *Hollow Promises?* (23 Oct 2026). The experimental evidence (Studies 1 and 2). | 754c07bc93e118f9 |
| `vshell_theory_and_operation.md` | Theory and operation of the discrete predecessor `vshell` (repository root). Exported from the claude.ai doc; images and diagrams were dropped. | 610002a53708847c |
| `feasibility_study.md` | The feasibility study that chose the implicit-field approach. Exported from the claude.ai doc; images dropped. | fe5419c8d24aae04 |

Quick import check (CPU-only works; the `cpu_direct` engine is selected automatically):

```python
import sys; sys.path.insert(0, "continuous/reference")
import analysis, gpu_analysis
print(gpu_analysis.ELEMENT_TYPES)          # ('q1_gauss', 'q1', 'q1_im')
```

Modules referenced in docstrings but **not** supplied:

- `docs/theory.md`, `docs/README.md`, `buckling_implementation_plan.md`
- `pcls_optimiser.py`, `problem_parser.py`, `optimise.py`
- `element_matrices.py`
- `utils/cpu_backend.py`, `utils/validate_buckling_fd.py`

Known issue found while writing the spec: for `element_type="q1_im"`, `run_buckling`'s sensitivity uses the Gauss (`q1_gauss`) solid stiffness in Term A and in the adjoint term. Directional finite differences show a constant bias of about 5.8 %, against 9·10⁻⁵ for `q1`. See THEORY §6.4.
