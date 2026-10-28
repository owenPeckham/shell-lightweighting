"""Does nvmath/cuDSS accept int64 CSR indices (needed once nnz(K) > 2^31, i.e. MBB/cantilever at 0.2 mm)?"""
import json
import numpy as np
import cupy as cp
import cupyx.scipy.sparse as cpsp
from nvmath.sparse.advanced import DirectSolver, DirectSolverMatrixType, DirectSolverOptions

n = 1000
A = cpsp.diags([-1.0, 2.0, -1.0], [-1, 0, 1], shape=(n, n), format="csr")
out = {}
for it in (np.int32, np.int64):
    K = cpsp.csr_matrix((A.data, A.indices.astype(it), A.indptr.astype(it)), shape=A.shape)
    b = cp.asfortranarray(cp.ones((n, 1)))
    try:
        s = DirectSolver(K, b, options=DirectSolverOptions(sparse_system_type=DirectSolverMatrixType.SPD))
        s.plan(); s.factorize(); x = cp.asarray(s.solve()).reshape(-1)
        out["cupy_" + it.__name__ + "_dtype_" + str(K.indptr.dtype)] = float(cp.linalg.norm(A @ x - 1) / np.sqrt(n))
        s.free()
    except Exception as e:
        out[str(it.__name__)] = repr(e)[:300]
import scipy.sparse as sps
Ah = A.get()
for it in (np.int32, np.int64):
    K = sps.csr_matrix((Ah.data, Ah.indices.astype(it), Ah.indptr.astype(it)), shape=Ah.shape)
    K.indices = K.indices.astype(it); K.indptr = K.indptr.astype(it)
    b = np.asfortranarray(np.ones((n, 1)))
    try:
        s = DirectSolver(K, b, options=DirectSolverOptions(sparse_system_type=DirectSolverMatrixType.SPD))
        s.plan(); s.factorize(); x = np.asarray(s.solve()).reshape(-1)
        out["scipy_" + it.__name__ + "_dtype_" + str(K.indptr.dtype)] = float(np.linalg.norm(Ah @ x - 1) / np.sqrt(n))
        s.free()
    except Exception as e:
        out["scipy_" + it.__name__] = repr(e)[:300]
print("int64_test", json.dumps(out), flush=True)
