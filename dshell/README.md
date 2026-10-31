# dshell: constant-mass variable-thickness shelling (discrete regions)

`dshell` (called `vshell` before 2026-10-31) takes a solid STEP part (for example a Fusion generative-design outcome) and its supports and loads. It decides which regions to hollow, sizes their walls, and holds the part's mass at the original value. A hollow design is adopted only if it improves stiffness and/or load capacity. Otherwise the original solid is returned with the reason. Every run writes a PDF report.

It is the standalone successor to the Fusion add-in in `ushell/`. The theory, maths and code walkthrough are in the companion document "vshell: Variable-Thickness Shelling Toolchain - Theory and Operation" (Markdown export: `docs/cshell/reference/vshell_theory_and_operation.md`).

## Install

```
pip install -r dshell/requirements.txt
```

Python 3.10 or newer on Windows, macOS or Linux. OpenCascade comes in through build123d (OCP). No Fusion, CalculiX or GPU is needed.

## Use (from the repository root)

```
python -m dshell faces "examples/parts/Solid L-bracket.step" --png faces.png   # face ids, to pick supports/loads
python -m dshell init  "my part.step"                                           # template config
python -m dshell run   examples/dshell/l_bracket.json                           # optimise + report
python -m dshell run   examples/dshell/l_bracket.json --set search.max_fe_evaluations=6 grid.sdf_mm=0.5
python -m dshell report examples/dshell/runs/L-bracket                          # rebuild report.pdf
```

A config needs only the STEP file, the supports and the loads (paths relative to the config file):

```json
{
  "name": "L-bracket",
  "step_file": "../parts/Solid L-bracket.step",
  "supports": [{"type": "cylinder", "radius": 4.26, "tol": 0.05}],
  "loads": [{"faces": {"type": "plane", "normal": [0, 0, 1], "box": {"x": [-15, 20]}, "extreme": true},
             "force_N": [0, 0, -200]}]
}
```

Supports clamp every node on those faces in all directions (the minimum-stiffness bolt-hole case). All other settings have defaults in `dshell/config.py`: material (PA11), minimum wall 1.5 mm, 0.2 mm as-built wall bias, grid 0.4 mm, FE budget 10 designs, and the acceptance rule.

## Outputs (`<output_dir>/<name>/`)

`final_design.stl` (the part to print), `best_candidate.stl`, `original.stl`, `report.pdf`, `plots/*.png`, `summary.json/.txt`, `run_info.json` (with provenance), `regions.csv`, `evaluations.csv`, `mass_sweep.csv`, `escape_holes.csv`, `sensitivity.csv` (decision re-scored for other K_eff and failure stress), `arrays.npz`, `run_log.txt`.

The committed example results are in `examples/dshell/runs/<part>/` (from before the rename, so their texts say `vshell`; no STLs).

## Tests

```
python tests/dshell/test_basic.py            # element, patch test, beam vs Timoshenko, selectors (~1 min)
python tests/dshell/validate_tubes.py 0.8    # Study 2 tubes vs the paper's Abaqus results (~4 min)
python tests/dshell/validate_thin_wall.py    # thin-wall mesh convergence (~5 min)
```

## Resources

The three Study 1 parts (44-89 g) took 15-23 minutes each on a 0.4 mm grid, peaking at 2.4-4.5 GB of RAM on 2 cores; set grid.sdf_mm=0.5 to roughly halve memory on smaller machines. Most of the time goes on the 10 FE solves, each 30-70 s at 0.3-0.65 million DOF.
