"""pytest set-up for cshell: CPU by default; GPU tests are marked @pytest.mark.gpu and skipped without CUDA."""
import os
import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT))


def _have_cuda() -> bool:
    try:
        from numba import cuda
        return bool(cuda.is_available())
    except Exception:
        return False


HAVE_CUDA = _have_cuda()


def pytest_configure(config):
    config.addinivalue_line("markers", "gpu: needs a CUDA device")
    os.environ.setdefault("CSHELL_FE_SCRATCH", str(Path(os.environ.get("TMPDIR", "/tmp")) / "cshell_fe"))


def pytest_collection_modifyitems(config, items):
    if HAVE_CUDA:
        return
    skip = pytest.mark.skip(reason="no CUDA device")
    for it in items:
        if "gpu" in it.keywords:
            it.add_marker(skip)


ENGINES = [pytest.param("cpu_direct", id="cpu"), pytest.param("auto", id="gpu", marks=pytest.mark.gpu)]


@pytest.fixture(params=ENGINES)
def engine(request):
    return request.param
