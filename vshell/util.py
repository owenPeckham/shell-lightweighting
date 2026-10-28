"""Provenance, JSON cleaning and CSV helpers."""
from __future__ import annotations

import csv
import hashlib
import json
import os
import platform
import socket
import subprocess
import sys
from datetime import datetime, timezone

import numpy as np

PACKAGES = ("vshell", "numpy", "scipy", "scikit-image", "skan", "libigl", "pyamg", "trimesh", "matplotlib",
            "build123d", "cadquery-ocp", "OCP")
CODE_FILES = ("__init__.py", "config.py", "geometry.py", "field.py", "segment.py", "fe.py", "mechanics.py",
              "optimise.py", "report.py", "util.py", "cli.py")


def sha256_of(path):
    try:
        h = hashlib.sha256()
        with open(path, "rb") as f:
            for chunk in iter(lambda: f.read(1 << 20), b""):
                h.update(chunk)
        return h.hexdigest()
    except Exception:
        return None


def _git(code_dir, *args):
    try:
        out = subprocess.run(["git", "-C", code_dir, *args], capture_output=True, text=True, timeout=10)
        return out.stdout.strip() if out.returncode == 0 else None
    except Exception:
        return None


def package_versions():
    from importlib.metadata import version, PackageNotFoundError
    out = {}
    for name in PACKAGES:
        try:
            out[name] = version(name)
        except PackageNotFoundError:
            out[name] = None
        except Exception:
            out[name] = None
    return out


def provenance(step_file=None, config_file=None):
    code_dir = os.path.dirname(os.path.abspath(__file__))
    status = _git(code_dir, "status", "--porcelain")
    info = {
        "timestamp_utc": datetime.now(timezone.utc).isoformat(timespec="seconds"),
        "host": socket.gethostname(),
        "platform": platform.platform(),
        "python": sys.version.split()[0],
        "command_line": list(sys.argv),
        "working_directory": os.getcwd(),
        "code_directory": code_dir,
        "git": {"commit": _git(code_dir, "rev-parse", "HEAD"),
                "describe": _git(code_dir, "describe", "--always", "--dirty", "--tags"),
                "dirty": (bool(status) if status is not None else None)},
        "code_sha256": {f: sha256_of(os.path.join(code_dir, f)) for f in CODE_FILES
                        if os.path.isfile(os.path.join(code_dir, f))},
        "packages": package_versions(),
        "cpu_count": os.cpu_count(),
    }
    if step_file:
        info["step_file"] = {"path": os.path.abspath(step_file), "sha256": sha256_of(step_file)}
    if config_file and os.path.isfile(str(config_file)):
        info["config_file"] = {"path": os.path.abspath(config_file), "sha256": sha256_of(config_file)}
    return info


def clean_json(x):
    if isinstance(x, dict):
        return {str(k): clean_json(v) for k, v in x.items()}
    if isinstance(x, (list, tuple, set)):
        return [clean_json(v) for v in x]
    if isinstance(x, np.ndarray):
        return clean_json(x.tolist())
    if isinstance(x, (np.integer,)):
        return int(x)
    if isinstance(x, (float, np.floating)):
        x = float(x)
        if np.isnan(x):
            return None
        if np.isinf(x):
            return "inf" if x > 0 else "-inf"
        return x
    if isinstance(x, np.bool_):
        return bool(x)
    return x


def write_json(path, obj):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(clean_json(obj), f, indent=2)


def write_csv(path, rows, columns=None):
    rows = list(rows)
    if not rows:
        open(path, "w").close()
        return
    columns = columns or list(dict.fromkeys(k for r in rows for k in r.keys()))
    with open(path, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=columns, extrasaction="ignore")
        w.writeheader()
        for r in rows:
            w.writerow({k: _csv_val(r.get(k)) for k in columns})


def _csv_val(v):
    if v is None:
        return ""
    if isinstance(v, (list, tuple, set)):
        return " ".join(str(_csv_val(x)) for x in v)
    if isinstance(v, (float, np.floating)):
        return f"{float(v):.6g}"
    return v


def read_csv(path):
    if not os.path.isfile(path):
        return []
    with open(path, encoding="utf-8") as f:
        rows = list(csv.DictReader(f))
    for r in rows:
        for k, v in r.items():
            try:
                r[k] = float(v) if v not in ("", None) else None
            except ValueError:
                pass
    return rows
