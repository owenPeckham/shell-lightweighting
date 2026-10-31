# shell-lightweighting
Author - Owen Peckham

Description - Open access repo for constant-mass structural shelling: turning a solid part into a hollow shell part of
the same mass that is stiffer and/or stronger. It contains the original Python-based Fusion 360 add-in and two
standalone optimisers that succeed it, as associated with work done in the University of Bristol's DMF lab.

THE UNIVERSITY OF BRISTOL (UOB) PROVIDES THIS PROGRAM "AS IS" AND WITH ALL FAULTS. UOB SPECIFICALLY DISCLAIMS ANY IMPLIED WARRANTY OF MERCHANTABILITY OR FITNESS FOR A PARTICULAR USE. UOB DOES NOT WARRANT THAT THE OPERATION OF THE PROGRAM WILL BE UNINTERRUPTED OR ERROR FREE.

## The three optimisers

| Folder | What it is | Runs on | Start here |
| --- | --- | --- | --- |
| `ushell/` | **Uniform shell**: the Fusion 360 add-in (one wall thickness, constant mass) | Fusion 360 (Windows/macOS) | below |
| `dshell/` | **Discrete-region shell** (was `vshell`): hollows chosen regions, sizes their walls, voxel FE | any PC, CPU, Python ≥ 3.10 | `dshell/README.md` |
| `cshell/` | **Continuous shell**: per-vertex wall thickness/offset fields, MMA, GPU FE, buckling | GPU cluster (Isambard-AI, GH200) | `cshell/README.md` |

Example parts (STEP) are in `examples/parts/`; configs and committed results per optimiser are in `examples/dshell/`
and `examples/cshell/`. **`INDEX.md` maps every folder and where each result, report and document is.**

## Installing the Fusion 360 add-in (`ushell/`)

Copy the `ushell/` folder into Fusion's AddIns folder and rename the copy to `ShellOptimisation` (Fusion requires the
folder name to match `ShellOptimisation.py` / `ShellOptimisation.manifest`):

- Windows: `%APPDATA%\Autodesk\Autodesk Fusion 360\API\AddIns\ShellOptimisation\`
- macOS: `~/Library/Application Support/Autodesk/Autodesk Fusion 360/API/AddIns/ShellOptimisation/`

Then in Fusion: Utilities → Add-Ins → Scripts and Add-Ins → Add-Ins tab → ShellOptimisation → Run.
