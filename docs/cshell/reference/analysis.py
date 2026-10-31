"""
FEA analysis components for the PCLS topology optimiser.

This module implements three interchangeable analysis backends (Conformal,
Regular, ConjugateGradient), each of which can run a primal linear-elastic
FEA (`run_analysis`) and, as of this revision, an optional linear buckling
eigenvalue analysis (`run_buckling`) that the optimiser can use to impose a
buckling safety-factor constraint.

See `buckling_implementation_plan.md` for the full derivation and rationale
behind the formulation used here. In summary:

  - The buckling eigenproblem is solved in terms of the INVERSE eigenvalue
    mu = 1/lambda, so that only the action of K^-1 (already available from
    the primal solve's factorisation) is needed -- no shift-invert
    infrastructure is required.
  - Because buckling eigenvalues become repeated or near-repeated
    constantly during topology optimisation, the m most critical inverse
    eigenvalues are combined with a p-norm aggregate mu_tilde before any
    sensitivity is computed.
  - The sensitivity of mu_tilde with respect to the density field is
    computed via one adjoint solve (shared across all aggregated modes),
    reusing the SAME stiffness factorisation as the eigensolver.
  - The Conformal, Regular and ConjugateGradient classes share this
    orchestration through the `Analysis` base class, and each concrete
    class supplies only: its assembled stiffness matrix and displacement
    solution, an actual and a "density-derivative direction" constitutive
    law, and a way to map an element-wise sensitivity onto the voxel grid.

IMPORTANT: the adjoint sensitivity implemented here is mathematically
delicate. Before trusting it on YOUR problem, validate it by finite
differences -- see `validate_buckling_sensitivity_regular` at the bottom
of this module, and Step 3 of the implementation plan.

Validation status at the time of writing: the ConjugateGradient
sensitivity path was checked by central finite differences against the
analytical adjoint on four axially-compressed cantilever columns --
2x5x16 (asymmetric), 3x3x16 and 4x4x14 (symmetric, doubly-degenerate
critical mode) and 4x5x14 (near-degenerate) -- with a worst-case
relative error of 0.017% across all elements, mode counts and step
sizes tested. Symmetric cross-sections are handled correctly; see
`_check_cluster_truncation` for why, and for the one degeneracy case
that does need care.

That figure pre-dates the geometric-stiffness ramp's own density term
(Term C, docs/theory.md eq. 6.10a). The current sensitivity, Term C
included, has since been re-checked by central differences on a 4x4x14
column with elements planted inside and below the ramp band
(utils/validate_buckling_fd.py): Regular agrees to 3.4e-6 with both
quadrature orders, and the ConjugateGradient host path (run through the
CPU stand-in utils/cpu_backend.py, which replaces only the CUDA kernels)
matches Regular to 2e-7; its own difference check is limited to ~1e-3 by
the float32 element properties the GPU path stores, not by the formula.
The Conformal path shares the same derivation and base-class code but has
not been executed against finite differences (it needs a remesh per
perturbation, and TetGen).

Note also that Euler's formula for a fixed-free column was reproduced to
within ~4.7% on a coarse 3x3x16 mesh, which is the expected order of
trilinear-hex discretisation error at that aspect ratio rather than a
sign of a formulation problem.
"""

import os
import io
# Optional dependencies. None of these is needed by the ConjugateGradient (CG)
# path; Conformal/Regular need the ones that fail here and will raise a
# NameError/TypeError naming them if used without.
try:
    import tetgen
except Exception:  # pragma: no cover
    tetgen = None
try:
    import trimesh
except Exception:  # pragma: no cover
    trimesh = None
import numpy as np
# (dataclasses no longer needed here: BucklingResults moved to interfaces.py)
try:
    from skfem import * # type: ignore
except Exception:  # pragma: no cover - only Conformal/Regular need scikit-fem
    pass
try:
    import pyvista as pv
except Exception:  # pragma: no cover
    pv = None
try:
    from CGAL import alpha_wrap
except Exception:  # pragma: no cover - Windows-only .pyd
    alpha_wrap = None
import matplotlib.pyplot as plt
from contextlib import redirect_stdout
from scipy.sparse import coo_matrix
from scipy.sparse.csgraph import connected_components
from scipy.interpolate import griddata
from scipy import ndimage
from timeit import default_timer as timer
from skimage.measure import marching_cubes
from scipy.sparse.linalg import eigsh, LinearOperator, ArpackNoConvergence
try:
    from pypardiso import spsolve as pardiso_solve, PyPardisoSolver
    DIRECT_SOLVER_NAME = "pypardiso"
except Exception:  # pragma: no cover - no MKL PARDISO (e.g. aarch64)
    PyPardisoSolver = None
    DIRECT_SOLVER_NAME = "scipy.splu (cached factorisation)"
    import scipy.sparse.linalg as _spla
    _SPLU_CACHE = {"A": None, "lu": None}

    def pardiso_solve(A, b):
        """SciPy fallback for pypardiso.spsolve: SuperLU, reusing the factorisation
        while called with the same matrix object (as pypardiso's own cache does)."""
        if _SPLU_CACHE["A"] is not A:
            _SPLU_CACHE["lu"] = _spla.splu(sp.csc_matrix(A, dtype=np.float64))
            _SPLU_CACHE["A"] = A
        b = np.asarray(b, dtype=np.float64)
        return _SPLU_CACHE["lu"].solve(b)
import scipy.sparse as sp
try:
    from skfem.models.elasticity import lame_parameters
except Exception:  # pragma: no cover
    def lame_parameters(E, nu):
        return E * nu / ((1. + nu) * (1. - 2. * nu)), E / (2. * (1. + nu))
from interfaces import BucklingResults
try:
    from skfem.helpers import ddot, sym_grad, grad, eye, trace
except Exception:  # pragma: no cover
    pass
from gpu_analysis import ELEMENT_TYPES, element_khat
from gpu_analysis import (
    matrix_free_pcg,
    PRECONDITIONERS,
    PCG_ENGINES,
    compute_element_strain_energy,
    compute_element_stress,
    hex_reference_data as _hex_reference_data,
    hex_dof_map as _hex_dof_map,
    lame_from_E_nu as _lame_from_E_nu,
    VOIGT_C_LAMBDA_BASIS as _VOIGT_C_LAM_BASIS,
    VOIGT_C_MU_BASIS as _VOIGT_C_MU_BASIS,
)


# NOTE: `BucklingResults` is defined in interfaces.py (imported above) so that
# pcls_optimiser.py can type-annotate against it without importing this module.


class Analysis:

    def __init__(self, iter_name: str, printouts: str="Normal", data_dir: str="analysis_data", indent="  "):
        """
        Initializes the Analysis super class with specified parameters.
        
        Args:
        - iter_name (str): Name of the current iteration.
        - printouts (str): Level of printouts ("Minimal", "Normal", "Verbose").
        - data_dir (str): Directory to store analysis data.
        """
        # Set the printout level
        if printouts not in ["Silent", "Minimal", "Normal", "Verbose"]:
            raise ValueError("printouts must be one of 'Silent', 'Minimal', 'Normal', or 'Verbose'")
        self.printouts = printouts

        self.indent = indent

        # Check is the iteration name is a valid directory path
        # if not iter_name.isidentifier():
        #     raise ValueError(f"'iter_name' must be a valid directory name, it is currently {iter_name}.")
        self.iter_name = iter_name.replace(os.path.sep, '_').replace('/', '_')

        # Create the analysis data directory if it doesn't exist
        if not iter_name == "test":
            self.analysis_dir = os.path.join(data_dir, iter_name)
            os.makedirs(self.analysis_dir, exist_ok=True)

        # --- Standardised, backend-shared per-iteration ".vtr" output ---
        #
        # Each concrete backend's run_analysis() populates these instead of
        # writing its own SED/density/displacement files directly, so that
        # everything grid-native ends up in ONE binary-encoded RectilinearGrid
        # per iteration (docs/theory.md §7.5/§7.6). run_buckling() (called
        # separately, afterwards, by the optimiser) can add to
        # _pending_point_data too, which is why the actual file write is
        # deferred to finalize_iteration_output() rather than happening
        # inside run_analysis() itself -- buckling mode shapes are not known
        # yet at that point.
        #
        # - _grid_shape: (nx, ny, nz) ELEMENT/voxel counts, set by
        #   run_analysis(). None means "no grid-native data this iteration"
        #   (e.g. save_meshes is False), and finalize_iteration_output()
        #   writes nothing in that case.
        # - _grid_hx/_grid_hy/_grid_hz: element edge length (m) per axis.
        self._pending_cell_data: dict = {}
        self._pending_point_data: dict = {}
        self._grid_shape = None
        self._grid_hx = self._grid_hy = self._grid_hz = None

    def finalize_iteration_output(self, field_data: dict | None = None,
                                   mode_shapes: np.ndarray | None = None,
                                   max_modes_to_save: int = 2):
        """
        Writes this iteration's standardised ".vtr" (see docs/theory.md
        §7.5/§7.6), combining whatever run_analysis() already stashed in
        `_pending_cell_data`/`_pending_point_data` with:

        - `field_data`: global, non-geometric scalars (compliance, lambda1,
          outer_iter, ...) that only the optimiser knows, attached as VTK
          FieldData.
        - `mode_shapes`: the (Ndofs, k) buckling mode-shape array from
          run_buckling(), if it was called this iteration. Only the leading
          `max_modes_to_save` columns are kept (buckling problems can ask
          for many more modes than are useful to actually visualise, and
          each one is a full extra nodal field). Mapped onto the grid via
          the backend's own `_map_mode_shapes_to_grid`, the same way
          displacement is mapped in `run_analysis`.

        Must be called exactly once per saved iteration, AFTER run_analysis()
        and (if it ran) run_buckling() -- this is what lets a buckling mode
        shape end up in the same file as everything else, rather than a
        second file. Safe to call even when nothing was saved this
        iteration (`save_meshes` False / `_grid_shape` still None): it then
        does nothing and returns None.

        Returns the path written, or None.
        """
        if self._grid_shape is None or not self._pending_cell_data and not self._pending_point_data:
            # Nothing grid-native was produced this iteration (save_meshes
            # was False, or this backend had zero active elements this
            # iteration and returned early). Still clear any stray state.
            self._pending_cell_data = {}
            self._pending_point_data = {}
            return None

        point_data = dict(self._pending_point_data)
        if mode_shapes is not None and mode_shapes.shape[1] > 0:
            n_modes = min(max_modes_to_save, mode_shapes.shape[1])
            try:
                mode_fields = self._map_mode_shapes_to_grid(mode_shapes[:, :n_modes])
                point_data.update(mode_fields)
            except Exception as e:
                self._print_out(
                    f"Warning: failed to map buckling mode shape(s) onto the "
                    f"output grid; the iteration's .vtr will be written "
                    f"without them. Error: {e}", "Minimal"
                )

        filename = os.path.join(self.analysis_dir, f"{self.iter_name}_iteration_data.vtr")
        try:
            _save_rectilinear_vtr(
                filename, self._grid_shape, self._grid_hx, self._grid_hy, self._grid_hz,
                cell_data=self._pending_cell_data,
                point_data=point_data,
                field_data=field_data,
            )
            self._print_out(f"Saved standardised iteration data (VTR) to {filename}", "Verbose")
        except Exception as e:
            self._print_out(f"Warning: failed to save iteration_data.vtr. Error: {e}", "Minimal")
            filename = None

        self._pending_cell_data = {}
        self._pending_point_data = {}
        return filename

    def _map_mode_shapes_to_grid(self, mode_shapes: np.ndarray) -> dict:
        """
        Maps up to k buckling mode-shape DOF vectors (shape (Ndofs, k)) onto
        the output grid as nodal displacement-MAGNITUDE point-data arrays
        named 'Buckling_Mode_1_Magnitude', 'Buckling_Mode_2_Magnitude', ...
        (direction is discarded, matching how displacement itself is stored
        -- see docs/theory.md §7.6). Each concrete backend implements this
        using whatever mapping `run_analysis` already uses for its own
        displacement field, so the two are spatially consistent with each
        other.
        """
        raise NotImplementedError(
            f"{type(self).__name__} does not implement _map_mode_shapes_to_grid."
        )

    def _print_out(self, message, message_type):
        """
        Print messages based on the printout level.

        Args:
            message (str): The message to print.
            message_type (str): The type of message ("Minimal", "Normal", "Verbose").
        """
        
        # 1. Define the indent level based *only* on the message_type
        if message_type == "Minimal":
            indent_level = 0
        elif message_type == "Normal":
            indent_level = 1
        elif message_type == "Verbose":
            indent_level = 2
        else:
            # Don't print "Silent" or unknown types
            return

        # 2. Define a hierarchy of levels
        levels = {
            "Minimal": 1,
            "Normal": 2,
            "Verbose": 3
        }

        # Get the numeric level for the current setting and message
        printout_level = levels.get(self.printouts, 0) # 0 for "Silent"
        message_level = levels.get(message_type, 0)

        # 3. Check if the message's level is at or below the global setting
        # (e.g., if setting is "Normal" (2), print "Minimal" (1) and "Normal" (2))
        if message_level > printout_level:
            return
        
        # 4. If we are here, we print!
        
        # Build the full indent prefix
        prefix = self.indent * indent_level

        # Apply the indentation to all lines in a multi-line message
        formatted_message = "\n".join(prefix + line for line in str(message).splitlines())

        # Print the result
        print(formatted_message)

    # =====================================================================
    # --- SHARED LINEAR BUCKLING MACHINERY ---
    #
    # Everything below is class-independent: pure array/sparse-matrix math
    # that has no dependency on mesh type. Concrete subclasses assemble
    # their own (K, K_sigma) and provide a small number of polymorphic
    # hooks (see the Conformal/Regular/ConjugateGradient implementations),
    # then call `_run_buckling_fe` (skfem-based classes) or drive
    # `_solve_generalized_buckling_eigenproblem` directly (ConjugateGradient).
    # =====================================================================

    @staticmethod
    def _aggregate_mu(mu_values: np.ndarray, p: float):
        """
        Numerically stable p-norm aggregation of a set of inverse buckling
        eigenvalues mu_i = 1/lambda_i.

            mu_tilde = ( sum_i mu_i^p )^(1/p)

        mu_tilde is a smooth upper bound, max_i(mu_i) <= mu_tilde <=
        m^(1/p) max_i(mu_i), whose gradient stays well defined where modes
        are degenerate or cross (Ferrari & Sigmund 2019). The optimiser uses
        that gradient as the direction of the buckling update; by default it
        rescales the VALUE to max_i(mu_i) (adaptive normalisation), so that
        the constraint is on the true critical load factor rather than on
        this conservative surrogate (docs/theory.md §6.4, eq. 6.23).

        Non-positive mu values (modes that are not actually critical for
        this load direction -- e.g. a locally tensile member) are clipped
        to zero contribution rather than aggregated: a large negative mu
        would otherwise (incorrectly) dominate the p-norm.

        Derivative (the `weights` return value), derived exactly:

            d(mu_tilde)/d(mu_j) = (sum_i mu_i^p)^(1/p - 1) * mu_j^(p-1)

        which, written in the scaled quantities used below (r_i = mu_i/mu_max
        and S = sum_i r_i^p) to avoid overflow when p is large, becomes

            w_j = r_j^(p-1) * S^(1/p - 1)

        NOTE THE S^(1/p) FACTOR. An earlier revision of this method used
        w_j = r_j^(p-1) / S, which is the above multiplied by S^(-1/p) and
        is therefore WRONG whenever S is meaningfully greater than 1. S is
        close to 1 when one mode dominates the aggregate, so the error is
        invisible on a well-separated spectrum -- but S -> (number of modes
        in the cluster) when eigenvalues are repeated, which is exactly
        what happens on a symmetric cross-section. A doubly-degenerate
        critical pair gives S = 2 and hence a uniform 2^(1/8) - 1 = 8.3%
        error in every sensitivity, which is precisely the discrepancy the
        finite-difference check reported on a square column. The
        correctness of the form used now is pinned by the Euler
        homogeneity identity sum_j w_j * mu_j == mu_tilde (mu_tilde is
        homogeneous of degree 1 in mu), which the old form failed.

        Returns
        -------
        mu_tilde : float
        weights  : np.ndarray, d(mu_tilde)/d(mu_i), same shape as mu_values
        """
        if mu_values.size == 0:
            return 0.0, np.zeros(0)

        mu_max = np.max(mu_values)
        if mu_max <= 0.0:
            # No positive (i.e. genuinely critical/compressive) eigenvalue
            # in the retained set: there is no buckling risk to aggregate.
            return 0.0, np.zeros_like(mu_values)

        ratio = np.clip(mu_values / mu_max, 0.0, None)
        s = np.sum(ratio ** p)
        mu_tilde = float(mu_max * s ** (1.0 / p))

        with np.errstate(divide='ignore', invalid='ignore'):
            weights = np.where(ratio > 0.0, ratio ** (p - 1.0) * s ** (1.0 / p - 1.0), 0.0)

        return mu_tilde, weights

    @staticmethod
    def _check_cluster_truncation(mu_values_with_probe: np.ndarray, num_modes: int,
                                   relative_tolerance: float = 1e-3):
        """
        Detects whether the retained mode set ends part-way through a
        cluster of repeated (or near-repeated) eigenvalues.

        Why this matters: at exact multiplicity the individual eigenvalues
        are only directionally differentiable, and the eigenvectors
        spanning the degenerate subspace are arbitrary (any orthonormal
        basis is equally valid, and ARPACK's choice is essentially
        random). The p-norm aggregate is nonetheless perfectly smooth
        PROVIDED the whole cluster is retained, because the aggregate then
        depends on the cluster only through the trace of the sensitivity
        over its subspace -- and a trace is invariant to the choice of
        basis within that subspace. This is why a symmetric structure with
        a doubly-degenerate critical mode is handled correctly here rather
        than needing special-case treatment.

        That invariance is lost if the cluster is cut in half by the
        `num_modes` limit: a partial trace over an arbitrarily-chosen
        half-basis is NOT basis-independent, so the sensitivity becomes
        genuinely ill-defined and will behave erratically between
        iterations. The fix is always to raise `num_modes`.

        Detecting this requires looking at the first mode BEYOND the cut,
        which is why the eigensolver deliberately computes one extra
        "probe" mode that is then discarded before aggregation. The
        condition tested is whether the probe is degenerate with the last
        RETAINED mode; a boundary that falls on a clear spectral gap is
        fine, and two equal eigenvalues sitting entirely inside the
        retained set are fine as well.

        Args:
        - mu_values_with_probe (np.ndarray): descending eigenvalues,
          INCLUDING the extra probe mode if the solver returned one.
        - num_modes (int): the number the caller actually intends to keep.

        Returns True when truncation is detected (caller should warn).
        """
        if mu_values_with_probe.size <= num_modes or num_modes < 1:
            # The solver returned no mode past the cut, so nothing was
            # sliced off the end of the retained set.
            return False
        last_kept = mu_values_with_probe[num_modes - 1]
        probe = mu_values_with_probe[num_modes]
        scale = max(abs(last_kept), abs(probe), 1e-30)
        return abs(last_kept - probe) / scale < relative_tolerance

    @staticmethod
    def _subspace_continuity(modes_new: np.ndarray, modes_old: np.ndarray | None):
        """
        Per-mode overlap between each new mode and the SPAN of the previous
        mode set, in [0, 1]. This is the quantity to threshold when deciding
        whether mode tracking has broken down.

        Why not a plain per-mode MAC: buckling modes arrive in degenerate (or
        near-degenerate) clusters on any structure with symmetry, and within
        such a cluster the eigenvectors ARPACK returns are an arbitrary
        orthonormal basis of the subspace -- they rotate freely from one
        iteration to the next even when the physics has not changed at all. A
        pairwise MAC therefore reads low and triggers spurious "mode switch"
        warnings on essentially every iteration of a symmetric problem.
        Measured on a slowly-evolving 3x3x14 symmetric column whose mode set
        was in fact perfectly stable, the pairwise MAC ran 0.36-0.51 across
        consecutive evaluations while this subspace measure was exactly
        1.0000 -- so a MAC threshold of 0.7 would have fired every single
        iteration.

        Projecting onto the span instead is invariant to that arbitrary
        in-subspace rotation, so it only drops when a mode genuinely leaves
        the tracked set -- which is the event actually worth warning about.

        Args:
        - modes_new (np.ndarray): shape (Ndofs, m_new).
        - modes_old (np.ndarray | None): shape (Ndofs, m_old), or None.

        Returns:
        - np.ndarray of shape (m_new,) with values in [0, 1], or None when
          there is no comparable history (see `_compute_mac` for the
          DOF-count case).
        """
        if modes_old is None or modes_old.size == 0 or modes_new.size == 0:
            return None
        if modes_new.shape[0] != modes_old.shape[0]:
            return None

        # Orthonormal (Euclidean) basis for the previous subspace. The stored
        # modes are K-orthonormal, not Euclidean-orthonormal, so an explicit
        # QR is needed rather than reusing them directly.
        Q, _ = np.linalg.qr(modes_old)
        norms = np.linalg.norm(modes_new, axis=0)
        norms = np.where(norms < 1e-30, 1e-30, norms)
        projected = np.linalg.norm(Q.T @ modes_new, axis=0)
        return np.clip((projected / norms) ** 2, 0.0, 1.0)

    @staticmethod
    def _compute_mac(modes_new: np.ndarray, modes_old: np.ndarray | None):
        """
        Modal Assurance Criterion between two sets of mode shapes defined
        on the same (full) DOF vector.

        Args:
        - modes_new (np.ndarray): shape (Ndofs, m_new).
        - modes_old (np.ndarray | None): shape (Ndofs, m_old), or None.

        Returns:
        - np.ndarray of shape (m_new, m_old), or None if modes_old is None
          or empty. MAC values close to 1 indicate the same physical mode;
          values close to 0 indicate unrelated mode shapes.
        """
        if modes_old is None or modes_old.size == 0 or modes_new.size == 0:
            return None

        # The two mode sets must live on the same DOF vector to be
        # comparable. They will NOT when the mesh has been rebuilt between
        # calls, which happens every single iteration for the Conformal
        # backend (marching cubes + TetGen produce a different node count
        # each time). Returning None there is correct: mode continuity is
        # genuinely undefined across a remesh, and the alternative is a
        # shape-mismatch exception midway through an optimisation run.
        if modes_new.shape[0] != modes_old.shape[0]:
            return None

        numerator = (modes_new.T @ modes_old) ** 2
        norm_new = np.sum(modes_new ** 2, axis=0)
        norm_old = np.sum(modes_old ** 2, axis=0)
        denom = np.outer(norm_new, norm_old)
        denom = np.where(denom < 1e-30, 1e-30, denom)
        return numerator / denom

    @staticmethod
    def _filter_valid_modes(mu_values: np.ndarray, modes: np.ndarray,
                             energy_fraction_in_solid: np.ndarray,
                             min_energy_fraction: float = 0.5,
                             min_mu: float = 0.0):
        """
        Discards spurious modes: those with non-positive mu (not a genuine
        critical/compressive mode for this load case) or with too little of
        their modal strain energy in elements that carry geometric stiffness
        (the signature of a spurious low-density/void mode). The energy
        share is weighted by the geometric-stiffness ramp r(rho_e) -- see
        `_run_buckling_fe` and docs/theory.md §6.5(a). It was previously a
        hard rho_e >= 0.5 cut, which on a PCLS design (most voxels sit
        between 0.3 and 0.5 for a long grey phase) rejected almost every
        genuine mode and left the run logging an infinite, i.e. "safe",
        critical load factor for designs that in fact buckled below the
        applied load.

        Args:
        - mu_values (np.ndarray): shape (m,), candidate inverse eigenvalues.
        - modes (np.ndarray): shape (Ndofs, m), candidate mode shapes.
        - energy_fraction_in_solid (np.ndarray): shape (m,), ramp-weighted
          fraction (0-1) of each mode's modal strain energy (see above).
        - min_energy_fraction (float): minimum fraction to keep a mode.
        - min_mu (float): minimum mu to keep a mode (default 0.0: keep only
          genuinely critical/compressive modes).

        Returns:
        - (mu_values_kept, modes_kept, kept_indices)
        """
        keep = (mu_values > min_mu) & (energy_fraction_in_solid >= min_energy_fraction)
        kept_indices = np.where(keep)[0]
        return mu_values[kept_indices], modes[:, kept_indices], kept_indices

    # Retry limits of the buckling eigensolve (see the docstring below): when
    # ARPACK converges NO eigenpair, the solve is repeated once with an Arnoldi
    # basis `_EIG_RETRY_NCV_FACTOR` times larger and `_EIG_RETRY_MAXITER_FACTOR`
    # times more restarts than scipy's defaults (ncv = max(2k+1, 20),
    # maxiter = 10 n).
    _EIG_RETRY_NCV_FACTOR = 3
    _EIG_RETRY_MAXITER_FACTOR = 5
    # Shift of the inertia count (eq. 6.3a): modes with mu <= this are "no compressive mode" (a load
    # factor above 1e10 is infinite for the purpose of a constraint).
    _INERTIA_SHIFT = 1e-10

    def _count_modes_above(self, K_II, G_II, sigma: float | None = None):
        """
        The number of eigenvalues mu > sigma of G_II phi = mu K_II phi, by Sylvester's law of
        inertia (docs/theory.md eq. 6.3a), or None when it cannot be determined.

        K_II is symmetric positive definite, so K_II = L L^T and G_II - sigma K_II is congruent to
        L^-1 G_II L^-T - sigma I: the two have the same inertia, and the number of positive
        eigenvalues of G_II - sigma K_II equals the number of mu > sigma. That number is a by-product
        of one sparse symmetric indefinite (LDL^T) factorisation, MKL PARDISO matrix type -2, which
        reports it in iparm(22). It needs no iteration, so it is exact where ARPACK may not converge:
        zero means there is certainly no compressive mode, a positive count that ARPACK failed to
        deliver means the buckling load is genuinely unknown.
        """
        sigma = self._INERTIA_SHIFT if sigma is None else float(sigma)
        solver = None
        try:
            A = (sp.csr_matrix(G_II, dtype=np.float64) - sigma * sp.csr_matrix(K_II, dtype=np.float64)).tocsr()
            upper = sp.triu(A, format="csr")      # matrix type -2 takes the upper triangle only
            upper.sort_indices()
            solver = PyPardisoSolver(mtype=-2)
            solver._check_A(upper)
            solver.set_phase(12)                  # analysis + numerical factorisation
            solver._call_pardiso(upper, np.zeros((A.shape[0], 1)))
            return int(solver.get_iparm(22))
        except Exception as e:
            self._print_out(f"Inertia count of the buckling pencil failed ({type(e).__name__}: {e}).", "Verbose")
            return None
        finally:
            if solver is not None:
                try:
                    solver.free_memory(everything=True)
                except Exception:
                    pass

    def _solve_generalized_buckling_eigenproblem(self, K_II, G_II, num_modes: int,
                                                 tol: float = 1e-6):
        """
        Solves G_II @ phi = mu * K_II @ phi for the `num_modes` algebraically
        largest mu (i.e. the smallest positive critical load factors
        lambda = 1/mu), reusing a single PARDISO factorisation of K_II for
        every operator application the eigensolver needs.

        This is scipy's ARPACK "regular inverse mode": since K_II is
        symmetric positive definite (it is the condensed elastic stiffness
        matrix of a well-posed, adequately-constrained structure), no
        shift-invert is required -- only the action of K_II^-1, supplied
        via `Minv`. See implementation plan sections 2.1-2.2.

        `pardiso_solve` (module-level `pypardiso.spsolve`) automatically
        detects that the SAME matrix object is being re-solved between
        calls and skips re-factorisation after the first solve, so the one
        factorisation performed here is reused for every eigensolver
        iteration AND is available to the caller for the adjoint solve
        afterwards, provided the caller reuses this exact `K_II` object.

        Args:
        - K_II (scipy.sparse matrix): condensed elastic stiffness (SPD).
        - G_II (scipy.sparse matrix): condensed geometric stiffness = -K_sigma.
        - num_modes (int): number of modes the caller intends to KEEP.
        - tol (float): ARPACK's relative accuracy for the eigenvalues, its
          stopping criterion. The default 1e-6 replaces ARPACK's own default of
          0, which means machine precision: the Lanczos iteration then keeps
          going long after the load factor is known to many more digits than
          the optimiser can use (the load factor is compared with a target to
          0.5 %, and the sensitivity is a search direction), and each extra
          iteration is a PARDISO back-substitution on the whole stiffness
          matrix. Pass 0.0 for machine precision, which the finite-difference
          validation of the sensitivity (utils/validate_buckling_fd.py) does.

        (A starting vector built from the previous evaluation's modes was
        tried and not kept: it changed the number of back-substitutions by
        -19 % to +6 % at random, no more than ARPACK's own start-to-start
        variation, because the first Arnoldi factorisation of ncv = 20 steps
        is paid whatever the start. docs/theory.md §6.10.)

        One extra "probe" mode beyond `num_modes` is computed and returned
        whenever the problem is large enough to allow it. The caller uses
        it solely to test (via `_check_cluster_truncation`) whether the
        retained set would end part-way through a cluster of repeated
        eigenvalues, then discards it before aggregating. The extra cost
        is one more Lanczos vector, which is negligible next to the
        factorisation that dominates this routine.

        PARTIAL CONVERGENCE IS NOT AN ERROR HERE. ARPACK converges slowly
        when the algebraically-largest eigenvalues are not well separated
        from the rest of the spectrum, and the worst case for that is
        exactly the benign one: a tension-dominated load case, where every
        eigenvalue is negative or near zero and there is no buckling mode to
        find. Treating `ArpackNoConvergence` as fatal would abort an entire
        optimisation run because one iteration happened to be in tension.
        The exception carries whatever subset did converge, so it is unpacked
        and used rather than re-raised. Only genuinely unexpected solver
        failures propagate.

        WHEN NOTHING CONVERGED (optimiser review follow-up, 2026-10). An
        empty result used to be read as "no compressive mode exists" (safe).
        That is true of a tension-only load case but not of a clustered
        compressive spectrum that merely hit ARPACK's limits, so the solve
        is now repeated ONCE with a larger Arnoldi basis and more restarts
        (`_EIG_RETRY_NCV_FACTOR`, `_EIG_RETRY_MAXITER_FACTOR`). If that also
        converges no eigenpair, `converged` is returned False and the caller
        reports the evaluation as UNKNOWN (BucklingResults status
        "NOT_CONVERGED"), which the optimiser neither treats as safe nor
        counts towards convergence. A solve that returns at least one
        eigenpair, or an empty one that ARPACK did not flag as failed, is
        unchanged.

        Returns:
        - (mu_values, mode_shapes, converged): mu_values shape (k,)
          descending, mode_shapes shape (n, k), where k is up to num_modes
          + 1; either may be empty. `converged` is False only when nothing
          converged even after the retry.
        """
        n = K_II.shape[0]
        k = min(num_modes + 1, max(n - 2, 1))
        if k < 1:
            return np.zeros(0), np.zeros((n, 0)), True

        def _kinv_matvec(x):
            return pardiso_solve(K_II, x)

        K_inv_op = LinearOperator(shape=(n, n), matvec=_kinv_matvec, dtype=np.float64)

        def _solve(**extra):
            """One eigsh call -> (mu_values, modes, nothing_converged)."""
            try:
                mu, modes = eigsh(G_II, k=k, M=K_II, sigma=None, Minv=K_inv_op,
                                  which='LA', tol=float(tol), **extra)
                return mu, modes, False
            except ArpackNoConvergence as e:
                mu = e.eigenvalues if e.eigenvalues is not None else np.zeros(0)
                modes = e.eigenvectors if e.eigenvectors is not None else np.zeros((n, 0))
                if mu.size == 0:
                    return np.zeros(0), np.zeros((n, 0)), True
                self._print_out(
                    f"Buckling eigensolver converged only partially "
                    f"({mu.size} of {k} requested modes). Continuing with "
                    f"the converged subset. If this recurs on a compression-loaded "
                    f"design, reduce num_modes or check the model for a "
                    f"near-singular stiffness matrix.", "Normal")
                return mu, modes, False
            except Exception as e:
                raise RuntimeError(f"Buckling eigensolver failed: {e}")

        mu_values, modes, nothing = _solve()
        converged = True
        n_compressive = None
        if nothing:
            # ARPACK delivered nothing. Count the compressive modes exactly (Sylvester's law, one LDL^T
            # factorisation): none means a tension-only design, which is a certain answer; some means
            # the load factor really is unknown unless a larger ARPACK basis finds them.
            n_compressive = self._count_modes_above(K_II, G_II)
            if n_compressive == 0:
                self._print_out("Buckling eigensolve: the inertia count finds no mode with mu > 0, so "
                                "there is no compressive mode (tension-dominated design).", "Verbose")
                return np.zeros(0), np.zeros((n, 0)), True
            # Retry once with larger ARPACK limits before giving up.
            ncv_default = min(n, max(2 * k + 1, 20))
            ncv = min(n, self._EIG_RETRY_NCV_FACTOR * ncv_default)
            if ncv > k and ncv > ncv_default:
                self._print_out(
                    f"Buckling eigensolver converged no eigenpair; retrying once with "
                    f"ncv = {ncv} (was {ncv_default}) and "
                    f"{self._EIG_RETRY_MAXITER_FACTOR}x the default restart limit.", "Normal")
                mu_values, modes, nothing = _solve(
                    ncv=ncv, maxiter=int(self._EIG_RETRY_MAXITER_FACTOR * 10 * n))
            if nothing:
                converged = False
                self._print_out(
                    "Buckling eigensolver converged no eigenpair even after the retry "
                    "with larger limits" + (f" although {n_compressive} compressive mode(s) exist (inertia count)"
                                            if n_compressive else "") + ": the buckling load is UNKNOWN "
                    "for this evaluation (it is not assumed safe).", "Minimal")
                return np.zeros(0), np.zeros((n, 0)), False

        if mu_values.size == 0:
            return np.zeros(0), np.zeros((n, 0)), converged

        # eigsh does not guarantee descending order
        order = np.argsort(mu_values)[::-1]
        return mu_values[order], modes[:, order], converged

    def _ddot_field_mean(self, field_a: np.ndarray, field_b: np.ndarray):
        """
        Per-element VOLUME INTEGRAL (not average) of the double-dot
        contraction of two (3, 3, Ne, Nqp) tensor fields:

            result_e = sum_qp( sum_ij field_a[i,j,e,qp] * field_b[i,j,e,qp] * dx[e,qp] )

        where `self.basis.dx` (shape (Ne, Nqp)) is skfem's per-quadrature-
        point differential volume element (|J| * quadrature weight) -- the
        SAME weighting skfem applies internally when a BilinearForm or
        LinearForm is assembled. This method is used OUTSIDE of any Form
        decorator (to get a per-element rather than a globally-assembled
        result), so that weighting has to be applied explicitly here; using
        a plain mean instead would silently omit the entire volume scaling
        and desynchronise this quantity from the actually-assembled K and
        K_sigma matrices it is meant to be the sensitivity of. (Caught by
        the finite-difference check in `validate_buckling_sensitivity_regular`
        during development -- see implementation plan Step 3.)

        NOTE: this returns an element INTEGRAL; `Regular._strain_energy_densities`
        returns the element's volume-AVERAGED strain energy density with the
        same dx weighting (sum sed dx / sum dx), so that F_e V_e is the element
        energy the compliance sensitivity -2 V_e F0_e needs. (It used a plain
        mean over the quadrature points until the 2026-10 review, which is
        biased for the default 4x4x4 Gauss rule.)
        """
        dot_qp = np.einsum('ijEQ,ijEQ->EQ', field_a, field_b)
        return np.sum(dot_qp * self.basis.dx, axis=1)

    def _geometric_tensor(self, mode_field: np.ndarray):
        """
        Builds the quadrature-point-resolved 'geometric strain' tensor

            Gamma_ij = sum_k (d phi_k/dx_i)(d phi_k/dx_j)

        from a full-length buckling mode shape vector (fixed DOFs assumed
        zero). This is the field that plays the role of a strain tensor in
        the geometric-stiffness bilinear form. Requires `self.basis` to be
        the same skfem vector Basis used for the primal solve.
        """
        grad_phi = grad(self.basis.interpolate(mode_field))  # shape (3,3,Ne,Nqp)
        return np.einsum('kiEQ,kjEQ->ijEQ', grad_phi, grad_phi)

    def _assemble_geometric_stiffness(self, stress_field: np.ndarray):
        """
        Assembles the geometric (stress) stiffness matrix K_sigma from a
        quadrature-point-resolved stress tensor field (shape (3,3,Ne,Nqp)),
        using the bilinear form

            a(v, u) = int  sum_k (d v_k/dx_i) sigma_ij (d u_k/dx_j) dV

        Sign convention: with sigma the ACTUAL (signed) Cauchy stress from
        the primal load case, K_sigma is negative-definite in a
        compression-dominated region, matching the standard linear
        buckling formulation (K + lambda * K_sigma) phi = 0, i.e.
        K phi = -lambda * K_sigma phi = lambda * G phi with G := -K_sigma.
        See implementation plan section 2.1.
        """
        @BilinearForm
        def geometric_stiffness(u, v, w):
            gv = grad(v)
            gu = grad(u)
            # Contract sigma's second index against gu's spatial index,
            # leaving an object with the same (component, spatial) leading
            # shape as gv so that ddot() contracts it correctly.
            sigma_gu = np.einsum('ijEQ,kjEQ->kiEQ', stress_field, gu)
            return ddot(gv, sigma_gu)

        return geometric_stiffness.assemble(self.basis)

    def _assemble_adjoint_load(self, gamma_field: np.ndarray):
        """
        Assembles the adjoint 'pseudo-load' vector g such that, for any
        admissible displacement perturbation delta_u,

            g . delta_u = sum_j w_j * phi_j^T (dK_sigma/du . delta_u) phi_j

        where `gamma_field` is the ALREADY-AGGREGATED, weighted geometric
        tensor sum_j w_j * Gamma(phi_j). Because K_sigma(u) is linear in u,
        this aggregation can happen before assembly, so only ONE adjoint
        load vector (and hence one adjoint solve) is needed regardless of
        how many modes are retained. See implementation plan section 2.4.

        Uses the ACTUAL constitutive law (`_stress_from_strain`), not its
        density-derivative direction: this is the elastic response to a
        virtual nodal displacement, not a design sensitivity.
        """
        @LinearForm
        def adjoint_load(v, w):
            stress_v = self._stress_from_strain(sym_grad(v))
            return np.einsum('ijEQ,ijEQ->EQ', stress_v, gamma_field)

        return adjoint_load.assemble(self.basis)

    def _modal_energy_fraction_in_solid(self, mode_field: np.ndarray, solid_weight=None):
        """
        Default (trivial) implementation: assumes every element in the FE
        model is "solid" (no ersatz/void material exists in the model), so
        every mode trivially passes the spurious-void-mode filter. This is
        correct for `Conformal`, whose mesh only ever covers the solid
        region. `Regular` and `ConjugateGradient`, which use a continuous
        ersatz-material interpolation over the full domain, override this
        with a real rho-based calculation. `solid_weight` (the per-element
        geometric-stiffness ramp) is accepted and ignored here so that the
        hook has one signature across backends.
        """
        return 1.0

    def _geometric_stiffness_ramp(self, threshold: float = 0.1):
        """
        Default (trivial) implementation: no ramp is needed because the
        model has no void elements (see `_modal_energy_fraction_in_solid`),
        so `threshold` is accepted and ignored to keep the hook signature
        uniform across backends. Returns a scalar multiplier broadcastable
        against a (Ne,) or (3,3,Ne,Nqp) array. `Regular` overrides this to
        ramp the geometric stiffness contribution to zero in low-density
        elements, guarding against spurious void-region buckling modes
        (section 4.1 of the implementation plan).
        """
        return 1.0

    def _geometric_stiffness_ramp_prime(self, threshold: float = 0.1):
        """
        d(ramp)/d(rho_e), the own-density-derivative of
        `_geometric_stiffness_ramp` that a full product rule for the
        sensitivity requires (see the comment above where `ramp` is
        computed in `_run_buckling_fe`). Default (trivial) implementation:
        the ramp is the constant 1.0 here (no void elements exist), so its
        derivative is 0.0. Returns a scalar or (Ne,)-shaped array
        broadcastable against the per-element sensitivity accumulator.
        `Regular` overrides this with the derivative of its cubic-
        smoothstep ramp.
        """
        return 0.0

    def _run_buckling_fe(self, voxel_field_shape: tuple, num_modes: int,
                          aggregation_p: float, prev_modes: np.ndarray | None,
                          min_energy_fraction: float = 0.5,
                          mode_continuity_warning_threshold: float = 0.7,
                          min_density_for_geometric_stiffness: float = 0.1,
                          eigensolver_tol: float = 1e-12):
        """
        Shared orchestration for the skfem-based (Conformal, Regular)
        buckling analysis. Assumes the concrete class has already, during
        its `_fea()` call, populated on `self`:

            self.basis        - the skfem vector Basis used for the primal solve
            self.u            - the primal displacement solution (full DOF vector)
            self.K            - the assembled elastic stiffness matrix (full, SPD)
            self.fixed_dofs   - the Dirichlet DOF index array from the primal solve

        and implements the polymorphic hooks:

            self._stress_from_strain(strain)        - actual constitutive law
            self._stress_from_strain_prime(strain)  - d(constitutive law)/d(rho)
            self._map_sensitivity_to_voxels(elem_vals) - per-element -> voxel grid
            self._num_elements()                    - element count
            self._modal_energy_fraction_in_solid(mode_field) - spurious-mode filter input
            self._geometric_stiffness_ramp()         - void-region safety ramp
            self._geometric_stiffness_ramp_prime()   - d(ramp)/d(rho_e), for the full product rule

        Returns a BucklingResults instance.
        """
        N = self.basis.N
        free_dofs = np.setdiff1d(np.arange(N), self.fixed_dofs)

        # 1) Stress field from the primal solution, and the geometric
        #    stiffness matrix G = -K_sigma.
        strain_u0 = sym_grad(self.basis.interpolate(self.u))
        stress_u0 = self._stress_from_strain(strain_u0)

        # Void-region safety ramp (section 4.1 of the implementation plan).
        # Applied to every stress-derived quantity feeding K_sigma below
        # (K_sigma itself, Term B, and the adjoint load), so that the
        # sensitivity differentiates the SAME ramped operator that was
        # actually used for the eigensolve. The ramp's own dependence on
        # rho_e is captured explicitly below via `ramp_prime` and Term C
        # (docs/theory.md §6.5(b)/§9.1): the full product rule for
        # d/drho_e[ramp(rho_e) * sigma(u0)] is
        #   ramp'(rho_e) * sigma(u0)  +  ramp(rho_e) * d(sigma(u0))/drho_e,
        # of which Term B below (using `_stress_from_strain_prime` and the
        # already-ramped `gamma_j`) supplies the second piece; Term C
        # supplies the first, using the UNRAMPED geometric tensor so the
        # two pieces do not double-count the ramp.
        ramp = self._geometric_stiffness_ramp(min_density_for_geometric_stiffness)
        ramp_prime = self._geometric_stiffness_ramp_prime(min_density_for_geometric_stiffness)
        K_sigma = self._assemble_geometric_stiffness(stress_u0 * ramp)
        G = -K_sigma

        K_II = self.K[free_dofs][:, free_dofs].tocsr()
        G_II = G[free_dofs][:, free_dofs].tocsr()

        # 2) Eigensolve: G_II phi = mu K_II phi, largest mu (most critical
        #    lambda = 1/mu) first.
        mu_values, modes_free, eig_converged = self._solve_generalized_buckling_eigenproblem(
            K_II, G_II, num_modes, tol=eigensolver_tol
        )

        if mu_values.size == 0:
            self._print_out(
                "Buckling eigensolver returned no modes (domain likely too "
                "small relative to the requested mode count, or no "
                "compressive mode exists)." if eig_converged else
                "Buckling eigensolver converged nothing: the buckling load is UNKNOWN.",
                "Minimal"
            )
            return BucklingResults(
                mu_values=np.zeros(0), mode_shapes=np.zeros((N, 0)),
                mu_tilde=0.0, sensitivity=np.zeros(voxel_field_shape),
                mac_matrix=None, solver_iters=0, n_candidates=0,
                eigensolver_converged=eig_converged,
            )

        modes_full = np.zeros((N, modes_free.shape[1]))
        modes_full[free_dofs, :] = modes_free
        n_candidates = int(np.count_nonzero(mu_values > 0.0))

        # 3) Filter spurious modes (non-positive mu, or too little modal
        #    energy in elements that carry geometric stiffness). The energy
        #    share is weighted by the SAME ramp r(rho_e) that scales K_sigma
        #    above (docs/theory.md §6.5(a)): a void-region mode has its
        #    energy where r = 0 and is rejected, while a genuine mode of a
        #    grey (rho ~ 0.3-0.5) design, whose energy sits where r = 1, is
        #    kept -- which the old hard rho >= 0.5 cut did not do.
        solid_weight = np.asarray(ramp, dtype=np.float64).reshape(-1) \
            if np.ndim(ramp) else np.ones(self._num_elements())
        energy_fraction = np.array([
            self._modal_energy_fraction_in_solid(modes_full[:, j], solid_weight=solid_weight)
            for j in range(modes_full.shape[1])
        ])
        mu_values, modes_full, _kept = self._filter_valid_modes(
            mu_values, modes_full, energy_fraction, min_energy_fraction
        )

        if mu_values.size == 0:
            if n_candidates > 0:
                self._print_out(
                    f"No buckling mode retained: all {n_candidates} candidate "
                    f"mode(s) with a positive eigenvalue were rejected by the "
                    f"spurious-mode filter (min_energy_fraction="
                    f"{min_energy_fraction}). The buckling load of this design "
                    f"is UNKNOWN for this evaluation (not assumed safe).", "Minimal"
                )
            else:
                self._print_out(
                    "No compressive buckling mode exists for this design and "
                    "load case (no positive eigenvalue): the buckling "
                    "constraint is satisfied.", "Normal"
                )
            return BucklingResults(
                mu_values=np.zeros(0), mode_shapes=np.zeros((N, 0)),
                mu_tilde=0.0, sensitivity=np.zeros(voxel_field_shape),
                mac_matrix=None, solver_iters=0, n_candidates=n_candidates,
            )

        mac_matrix = self._compute_mac(modes_full, prev_modes)
        overlap = self._subspace_continuity(modes_full, prev_modes)
        if overlap is not None and overlap.size:
            worst = float(overlap.min())
            if worst < mode_continuity_warning_threshold:
                self._print_out(
                    f"Note: the least-continuous buckling mode overlaps the "
                    f"previous mode subspace by only {worst:.3f} "
                    f"(< {mode_continuity_warning_threshold}). Either a mode has entered "
                    f"or left the tracked set, or the design simply moved a "
                    f"long way since the last evaluation. Both make this "
                    f"iteration's sensitivity refer to somewhat different "
                    f"physics than the last. If it persists, raise num_modes "
                    f"or reduce dt / buckling_evaluate_every.", "Minimal"
                )

        # 4) Check the probe mode for cluster truncation, then discard it
        #    and trim the retained set to num_modes before aggregating.
        if self._check_cluster_truncation(mu_values, num_modes):
            self._print_out(
                f"Warning: the retained buckling mode set would end part-way through a "
                f"cluster of repeated eigenvalues (mode {num_modes} and mode "
                f"{num_modes + 1} agree to within 0.1%). A partially-retained "
                f"degenerate cluster makes the aggregated sensitivity basis-dependent "
                f"and therefore ill-defined. Increase num_modes (currently "
                f"{num_modes}) until the retained set ends on a clear spectral gap.",
                "Minimal"
            )

        if mu_values.size > num_modes:
            mu_values = mu_values[:num_modes]
            modes_full = modes_full[:, :num_modes]

        mu_tilde, weights = self._aggregate_mu(mu_values, aggregation_p)
        self._print_out(
            f"Buckling: retained {mu_values.size} mode(s), "
            f"lambda_crit (min) = {1.0 / mu_values.max():.4f}, "
            f"mu_tilde = {mu_tilde:.6e}", "Normal"
        )

        # 5) Explicit sensitivity terms (no extra solve required) and the
        #    aggregated geometric tensor for the adjoint load, in one pass
        #    over the retained modes.
        n_elements = self._num_elements()
        explicit_sensitivity = np.zeros(n_elements)
        gamma_combined = None

        for j in range(mu_values.size):
            phi_j = modes_full[:, j]
            strain_phi = sym_grad(self.basis.interpolate(phi_j))
            # UNRAMPED geometric tensor -- needed for Term C below, which
            # supplies the ramp'(rho_e) * sigma(u0) piece of the product
            # rule and must not have the ramp baked in a second time.
            gamma_j_raw = self._geometric_tensor(phi_j)
            # Apply the SAME void-region ramp used for K_sigma itself, so
            # that Term B and the adjoint load (built from gamma_combined
            # below) differentiate the operator that was ACTUALLY used for
            # the eigensolve, not the un-ramped one.
            gamma_j = gamma_j_raw * ramp

            # Term A: -mu_j * phi_j^T (dK/drho_e) phi_j  (modal SED, using
            # the density-derivative constitutive law)
            term_a = -mu_values[j] * self._ddot_field_mean(
                self._stress_from_strain_prime(strain_phi), strain_phi
            )
            # Term B: -phi_j^T (ramp(rho_e) * dK_sigma/drho_e) phi_j -- the
            # ramp(rho_e) * d(sigma(u0))/drho_e piece of the product rule.
            term_b = -self._ddot_field_mean(
                self._stress_from_strain_prime(strain_u0), gamma_j
            )
            # Term C: -phi_j^T (ramp'(rho_e) * K_sigma-without-ramp/drho_e...) phi_j
            # -- the ramp'(rho_e) * sigma(u0) piece of the product rule that
            # was previously neglected (docs/theory.md §6.5(b)/§9.1). Uses
            # the ACTUAL (not density-derivative) constitutive law for
            # sigma(u0), since ramp_prime already carries the rho_e
            # dependence; multiplying by the unramped gamma_j_raw avoids
            # double-counting the ramp itself. `ramp_prime` is 0.0 for
            # Conformal (trivial ramp), so this term vanishes there exactly.
            term_c = -ramp_prime * self._ddot_field_mean(stress_u0, gamma_j_raw)
            explicit_sensitivity += weights[j] * (term_a + term_b + term_c)

            gamma_combined = gamma_j * weights[j] if gamma_combined is None \
                else gamma_combined + weights[j] * gamma_j

        # 6) One adjoint solve, using the aggregated geometric tensor and
        #    reusing the SAME K_II object (and therefore the SAME PARDISO
        #    factorisation) as the eigensolver.
        g_combined = self._assemble_adjoint_load(gamma_combined)
        nu_free = pardiso_solve(K_II, g_combined[free_dofs])
        nu_full = np.zeros(N)
        nu_full[free_dofs] = nu_free

        strain_nu = sym_grad(self.basis.interpolate(nu_full))
        implicit_sensitivity = self._ddot_field_mean(
            self._stress_from_strain_prime(strain_nu), strain_u0
        )

        d_mu_tilde_d_rho = explicit_sensitivity + implicit_sensitivity

        sensitivity_voxels = self._map_sensitivity_to_voxels(d_mu_tilde_d_rho)

        return BucklingResults(
            mu_values=mu_values,
            mode_shapes=modes_full,
            mu_tilde=mu_tilde,
            sensitivity=sensitivity_voxels,
            mac_matrix=mac_matrix,
            subspace_continuity=overlap,
            solver_iters=0,
            n_candidates=n_candidates,
        )


def _tet_sample_points() -> np.ndarray:
    """
    Barycentric coordinates of the 11 sample points per tet used to average
    per-tet fields over voxels (docs/theory.md §7.1, eq. 7.0c): the 4-point
    rule, the centroid, and one point towards each of the six edges.
    """
    a, b = 0.5854101966249685, 0.1381966011250105
    pts = [[a, b, b, b], [b, a, b, b], [b, b, a, b], [b, b, b, a], [0.25, 0.25, 0.25, 0.25]]
    for i in range(4):
        for j in range(i + 1, 4):
            w = [0.1, 0.1, 0.1, 0.1]
            w[i] = w[j] = 0.4
            pts.append(w)
    return np.array(pts, dtype=np.float64)


_TET_SAMPLE_POINTS = _tet_sample_points()


class Conformal(Analysis):

    def __init__(self, iter_name: str, printouts: str="Normal", data_dir: str="analysis_data", youngs_modulus: float=210.0E9, poisson_ratio: float=0.3, min_mesh_characteristic_length: float|None=None, max_mesh_characteristic_length: float|None=None, alpha_proportion:float=0.1, offset_proportion:float=0.01, save_meshes: bool=False, indent="  ", visualisation_scale: float=1.0, num_cells_cutoff: float=5.0, voxel_size_mm: float=1.0, material_interpolation: bool=True, mc_vertex_clamp: float=0.1, csg_domain_data: dict | None = None):
        """
        Initialises the Tetgen analysis class with specified parameters.

        Args:
        - iter_name (str): Name of the current iteration for data storage.
        - youngs_modulus (float): Young's modulus of the material.
        - poisson_ratio (float): Poisson's ratio of the material.
        - min_mesh_characteristic_length (float | None): Absolute point-merging tolerance (m)
          applied to the alpha-wrap surface before TetGen runs: surface vertices closer than
          this are merged, removing small/degenerate features that would otherwise force tiny
          Steiner-point tets there. TetGen has no direct "minimum tet size" switch, so this is
          how that floor is actually enforced. None/0 disables merging (§7.1, §10.6).
        - max_mesh_characteristic_length (float | None): Maximum tet edge length (m), passed to
          TetGen as a `maxvolume = length**3` coarseness ceiling. None falls back to 1% of the
          bounding-box volume.
        - alpha_proportion (float): The proportion of the geometry's bounding box diagonal to use as alpha in alpha wrap. The alpha value determines the level of detail in the mesh. A smaller alpha (0.01) captures finer details but may result in a more complex mesh, while a larger alpha (0.1) simplifies the mesh but may lose some features.
        - offset_proportion (float): The proportion of the geometry's bounding box diagonal to use as offset in alpha wrap. The offset value helps to avoid issues with degenerate faces and improves mesh quality. A small offset (0.002) creates a tight conformal mesh, while a larger offset (0.02) provides more leeway for mesh generation.
        - indent (str): Indentation string for print formatting.
        - printouts (str): Level of printouts ("Minimal", "Normal", "Verbose").
        - data_dir (str): Directory to store analysis data.
        - save_meshes (bool): Whether to save the generated meshes to files.
        - visualisation_scale (float): The scaling factor of deformation in the saved deformed vtk mesh file.
        - num_cells_cutoff (float): The number of grid cells large that a connected component has to be before it appears in the mesh
        - voxel_size_mm (float): The side length of a single design voxel in
          millimetres. This sets the scale factor between the optimiser's voxel
          coordinates and the FEA's metres, and therefore the physical size of
          the whole structure. It was previously hard-coded to 1.0 throughout
          this class while `Regular` and `ConjugateGradient` both honoured the
          parameter, so a problem file that set it produced a conformal model
          at the wrong physical scale with no warning.
        - material_interpolation (bool): give each tet the stiffness density of
          the design voxel it lies in (docs/theory.md §4.8, §7.1), the same
          penalised interpolation rho_m = c1 + (1 - c1) x^p the density
          backends use, so that grey material above the level set is as weak
          in this model as it is in theirs. The optimiser passes rho_m to
          run_analysis(stiffness_density=...). False meshes every tet with the
          full modulus, as before this revision: grey material above the level
          set is then as stiff as solid while costing only its density in the
          volume constraint, and the design has no reason to become binary.
        - mc_vertex_clamp (float): smallest distance, as a fraction of a voxel,
          between a marching-cubes vertex and the grid nodes at the two ends of
          its edge (docs/theory.md §7.1). A grid value within a hair of the level
          set puts the isosurface vertices on its edges almost ON the node,
          which gives slivers of triangles with edges of 1e-5 voxels; TetGen then
          fills the region with millions-to-one graded tets, some of them flat
          to round-off, and the FE solve returns NaN. Each vertex is moved along
          its own edge into [tau, 1 - tau]: the surface moves by at most tau
          voxels and its topology is unchanged. 0 disables the clamp.
        """

        super().__init__(iter_name, printouts, data_dir, indent=indent)

        # Set the rest of the parameters
        self.youngs_modulus = youngs_modulus
        self.poisson_ratio = poisson_ratio
        self.min_mesh_characteristic_length = min_mesh_characteristic_length
        self.max_mesh_characteristic_length = max_mesh_characteristic_length
        self.save_meshes = save_meshes
        self.alpha_proportion = alpha_proportion
        self.offset_proportion = offset_proportion
        self.visualisation_scale = visualisation_scale
        self.num_cells_cutoff = num_cells_cutoff
        self.material_interpolation = bool(material_interpolation)
        if not 0.0 <= mc_vertex_clamp < 0.5:
            raise ValueError("mc_vertex_clamp must lie in [0, 0.5).")
        self.mc_vertex_clamp = float(mc_vertex_clamp)
        # Mesh-conditioning diagnostics of the last analysis (docs/theory.md §7.1).
        self.mc_vertices_clamped = 0
        self.degenerate_tets_removed = 0
        self.min_tet_quality = None
        # Per-tet stiffness density (None = full modulus everywhere), set by
        # run_analysis when the optimiser passes a stiffness density.
        self._tet_rho_m = None
        self._stiffness_density_field = None
        self.material_interpolation_applied = False
        if voxel_size_mm <= 0.0:
            raise ValueError("voxel_size_mm must be positive.")
        self.voxel_size_m = voxel_size_mm / 1000.0
        self.voxel_volume_m3 = self.voxel_size_m ** 3

        if csg_domain_data:
            self.csg_data = csg_domain_data
        else:
            raise Exception("Required CSG domain data missing.")

        param_values = (
            f"{self.__class__.__name__} analysis component initialised with:\n"
            f"  Young's Modulus: {self.youngs_modulus:<.2e} Pa\n"
            f"  Poisson's Ratio: {self.poisson_ratio}\n"
            f"  Voxel Size: {1000.0 * self.voxel_size_m} mm\n"
            f"  CSG Geometry Enabled: {self.csg_data is not None}\n"
        )

        if self.min_mesh_characteristic_length is not None:
            param_values += f"  Min Mesh Characteristic Length: {self.min_mesh_characteristic_length}\n"
        if self.max_mesh_characteristic_length is not None:
            param_values += f"  Max Mesh Characteristic Length: {self.max_mesh_characteristic_length}\n"

        self.field_type = "phase"

        self._print_out(param_values, "Normal")
    
    @staticmethod
    def _separate_from_level(volume: np.ndarray, level: float):
        """
        Moves every grid value closer to the level set than delta = 4e-6 N
        (N the largest grid dimension, in field units) to level +- delta, on
        its own side (docs/theory.md §7.1). skimage returns the vertex
        coordinates in float32, whose spacing near N is about 1.2e-7 N: an
        edge parameter below that is rounded onto the node, and the vertex's
        edge can no longer be read from its coordinates. After the guard every
        edge parameter is at least delta / |dphi|, many float32 spacings for
        any |dphi| up to the width of the phi clamp. The change is 4e-6 N in
        phi at most and touches only values that are on the level set to
        within it.

        Returns the (possibly copied) field and the number of values moved.
        """
        delta = 4e-6 * max(volume.shape)
        near = np.abs(volume - level) < delta
        n_near = int(np.count_nonzero(near))
        if n_near == 0:
            return volume, 0
        out = np.array(volume, dtype=np.float64, copy=True)
        out[near] = np.where(out[near] >= level, level + delta, level - delta)
        return out, n_near

    @staticmethod
    def _clamp_marching_cubes_vertices(verts: np.ndarray, tau: float):
        """
        Moves each marching-cubes vertex along its own grid edge so that it is
        at least tau voxels from both nodes of that edge (docs/theory.md §7.1).

        Marching cubes puts one vertex on each grid edge the level set crosses,
        at x_a + t e_k with t = (L - phi_a) / (phi_b - phi_a) in (0, 1): two of
        its coordinates are exactly integral (in voxel units), and t is the
        fractional part of the third. A grid value phi_a almost equal to L
        gives t ~ 0 on every crossing edge at that node, and the triangles
        round it collapse into slivers with edges of order t. Clamping t to
        [tau, 1 - tau] bounds the smallest feature of the surface below by
        about tau voxels while moving no vertex by more than tau voxels; each
        vertex stays on its edge, so the connectivity (and the topology) of
        the surface is exactly that of the unclamped one. Vertices not on a
        grid edge are left alone.

        Returns the vertices (float64) and the number moved.
        """
        v = np.array(verts, dtype=np.float64, copy=True)
        if tau <= 0.0 or v.size == 0:
            return v, 0
        base = np.floor(v)
        frac = v - base
        integral = frac == 0.0
        rows = np.nonzero(integral.sum(axis=1) == 2)[0]
        axis = np.argmin(integral[rows], axis=1)
        t = frac[rows, axis]
        t_new = np.clip(t, tau, 1.0 - tau)
        v[rows, axis] = base[rows, axis] + t_new
        return v, int(np.count_nonzero(t_new != t))

    def _surface_mesh(self, phi, level_set_value):
        """
        Generates a surface mesh from the abstract field using the marching cubes algorithm.
        
        Args:
        - phi (np.ndarray): The abstract field representing the density distribution.
        - level_set_value (float): The level set value to extract the surface mesh.
        """

        self._print_out("Starting surface mesh generation using marching cubes and alpha wrap...", "Normal")
        t0 = timer()
        try:
            # Pad phi with zeros to ensure a gradient at the boundaries
            padded_phi = np.pad(phi, pad_width=1, mode='constant', constant_values=1.0)

            # Save contour plots of 'padded_phi' through the centre of the beam axis in the three orthogonal planes
            if self.save_meshes:
                try:
                    self._print_out("Saving contour plots of phi field...", "Verbose")
                    # Get center indices
                    shape = padded_phi.shape
                    center_x = shape[0] // 2
                    center_y = shape[1] // 2
                    center_z = shape[2] // 2

                    # Get the three slices
                    # Note: numpy indexing is [x, y, z] for skimage.marching_cubes
                    slice_yz = padded_phi[center_x, :, :] # Slice at X = center
                    slice_xz = padded_phi[:, center_y, :] # Slice at Y = center
                    slice_xy = padded_phi[:, :, center_z] # Slice at Z = center

                    # Create a figure with 3 subplots
                    fig, axes = plt.subplots(1, 3, figsize=(18, 5))
                    fig.suptitle(f'{self.iter_name} - Phi Field Slices', fontsize=16)

                    # --- Plot YZ Slice (at X=center) ---
                    ax0 = axes[0]
                    # Use .T to align plot axes with index names
                    contour0 = ax0.contourf(slice_yz.T, levels=30, cmap='viridis') 
                    ax0.contour(slice_yz.T, levels=[level_set_value], colors='red', linewidths=2)
                    fig.colorbar(contour0, ax=ax0, label='Phi Value')
                    ax0.set_title(f'YZ Plane (at X = {center_x})')
                    ax0.set_xlabel('Y Index')
                    ax0.set_ylabel('Z Index')
                    ax0.axis('image') # Equal aspect ratio

                    # --- Plot XZ Slice (at Y=center) ---
                    ax1 = axes[1]
                    contour1 = ax1.contourf(slice_xz.T, levels=30, cmap='viridis')
                    ax1.contour(slice_xz.T, levels=[level_set_value], colors='red', linewidths=2)
                    fig.colorbar(contour1, ax=ax1, label='Phi Value')
                    ax1.set_title(f'XZ Plane (at Y = {center_y})')
                    ax1.set_xlabel('X Index')
                    ax1.set_ylabel('Z Index')
                    ax1.axis('image')

                    # --- Plot XY Slice (at Z=center) ---
                    ax2 = axes[2]
                    contour2 = ax2.contourf(slice_xy.T, levels=30, cmap='viridis')
                    ax2.contour(slice_xy.T, levels=[level_set_value], colors='red', linewidths=2)
                    fig.colorbar(contour2, ax=ax2, label='Phi Value')
                    ax2.set_title(f'XY Plane (at Z = {center_z})')
                    ax2.set_xlabel('X Index')
                    ax2.set_ylabel('Y Index')
                    ax2.axis('image')

                    fig.tight_layout()
                    plot_path = os.path.join(self.analysis_dir, f"{self.iter_name}_phi_contours.png")
                    fig.savefig(plot_path, dpi=150)
                    plt.close(fig)
                    self._print_out(f"Contour plot saved to {plot_path}", "Verbose")
                
                except Exception as e:
                    self._print_out(f"Warning: Failed to save contour plots. Error: {e}", "Minimal")

            # Condition the isosurface (docs/theory.md §7.1): grid values within a
            # hair of the level set are first moved just off it, then every vertex
            # is kept at least mc_vertex_clamp voxels from the nodes of its edge.
            mc_field, self.mc_values_separated = self._separate_from_level(padded_phi, level_set_value)
            verts, faces, normals, _ = marching_cubes(volume=mc_field, level=level_set_value, step_size=1, allow_degenerate=False)
            verts, self.mc_vertices_clamped = self._clamp_marching_cubes_vertices(verts, self.mc_vertex_clamp)
            if self.mc_vertices_clamped or self.mc_values_separated:
                self._print_out(f"Marching cubes: {self.mc_vertices_clamped} vertex(es) moved to within "
                                f"[{self.mc_vertex_clamp:g}, {1 - self.mc_vertex_clamp:g}] of their edge, "
                                f"{self.mc_values_separated} grid value(s) moved off the level set.", "Verbose")

            # Convert from voxel coordinates to metres, which is what the FEA
            # works in. The scale factor is the voxel edge length rather than a
            # hard-coded 1 mm, so voxel_size_mm now actually reaches this class.
            verts -= 1  # Adjust for padding offset
            verts *= self.voxel_size_m

            # Check watertightedness with trimesh
            mesh = trimesh.Trimesh(vertices=verts, faces=faces, vertex_normals=normals, process=True)
            if not mesh.is_watertight:
                self._print_out("Warning: Marching cubes generated surface mesh is not watertight.", "Minimal")
            else:
                self._print_out("Generated surface mesh is watertight.", "Verbose")

            self._print_out(f"Initial marching cubes mesh has {verts.shape[0]} vertices and {faces.shape[0]} faces.", "Verbose")

            if self.save_meshes:
                # Save the unwrapped surface mesh to an STL file
                unwrapped_stl_path = os.path.join(self.analysis_dir, f"{self.iter_name}_unwrapped_surface_mesh.stl")
                mesh.export(unwrapped_stl_path)
                self._print_out(f"Surface mesh saved to {unwrapped_stl_path}", "Verbose")

            # Connected-component size cutoff, expressed as a ratio so that the
            # wrap discards components slightly smaller than one grid cell at
            # minimum (1/cutoff is used inside the wrap function as the
            # fraction). Computed HERE, outside the printout branch below:
            # it used to be computed inside the Verbose branch only, so the
            # quiet branch called wrap() without cutoff_ratio and silently kept
            # every speck of disconnected geometry. The two branches must
            # differ in logging only, never in what they compute.
            cutoff = (padded_phi.shape[0] - 2) * (padded_phi.shape[1] - 2) * (padded_phi.shape[2] - 2) / self.num_cells_cutoff

            # Use alpha wrap to robustly simplify the mesh
            if self.printouts == "Verbose":
                # 1. Create an in-memory text buffer
                log_buffer = io.StringIO()
                
                # 2. Use redirect_stdout to capture all output
                try:
                    with redirect_stdout(log_buffer):
                        # Call the C++ function.
                        # All its py::print() calls will be captured by log_buffer.
                        simplified_mesh = alpha_wrap.wrap(
                            vertices=verts, 
                            faces=faces, 
                            alpha_proportion=self.alpha_proportion, 
                            offset_proportion=self.offset_proportion,
                            cutoff_ratio=cutoff,
                            printout=True # Tell C++ to print
                        )
                    
                    # 3. Get the captured log string from the buffer
                    log_output = log_buffer.getvalue()
                    
                    # 4. Send the captured output to your custom print function
                    if log_output:
                        # Filter out the "Pre-discarding" lines
                        filtered_lines = [
                            line for line in log_output.splitlines() 
                            if not line.strip().startswith("Pre-discarding")
                        ]
                        
                        # Re-join the lines
                        filtered_log_output = "\n".join(filtered_lines)
                        
                        # Only print if there's content left after filtering
                        if filtered_log_output:
                            # .strip() removes any final blank lines
                            self._print_out(filtered_log_output.strip(), "Verbose")

                except Exception as e:
                    # 5. Make sure to print any captured logs before re-raising
                    log_output = log_buffer.getvalue()
                    if log_output:
                        self._print_out("--- C++ LOGS (from failed run) ---", "Minimal")
                        self._print_out(log_output.strip(), "Minimal")
                    raise e # Re-raise the original exception

            else:
                # If not verbose, run normally without any capture. Note the
                # argument list is otherwise IDENTICAL to the verbose branch
                # above, cutoff_ratio included.
                simplified_mesh = alpha_wrap.wrap(
                    vertices=verts,
                    faces=faces,
                    alpha_proportion=self.alpha_proportion,
                    offset_proportion=self.offset_proportion,
                    cutoff_ratio=cutoff,
                    printout=False
                )
            
            verts = np.array(simplified_mesh["vertices"])
            faces = np.array(simplified_mesh["faces"])

            self._print_out(f"Surface mesh generated from field with {verts.shape[0]} vertices and {faces.shape[0]} faces using level {level_set_value}", "Verbose")
        except Exception as e:
            raise RuntimeError(f"Marching cubes algorithm failed: {e}")
        
        stl_mesh = trimesh.Trimesh(vertices=verts, faces=faces, process=False)
        surface_volume = stl_mesh.volume
        self._print_out(f"Surface mesh volume: {1e9 * surface_volume:.6f} mm³", "Verbose")

        self.s_mesh = {"verts": verts, "faces": faces, "volume": surface_volume}

        self._print_out(f"Surface mesh generation time: {timer() - t0:.2f} seconds", "Normal")

        if self.save_meshes:
            # Save the surface mesh to an STL file
            stl_path = os.path.join(self.analysis_dir, f"{self.iter_name}_surface_mesh.stl")
            stl_mesh.export(stl_path)
            self._print_out(f"Surface mesh saved to {stl_path}", "Verbose")

    def _volume_mesh(self, s_mesh: dict | None = None):
        """
        Generates a robust volume mesh using PyVista/TetGen and stores it
        in a format optimized for scikit-fem (skfem).
        
        The output in self.v_mesh will be:
        {
            "nodes": np.array of shape (N, 3), 0-based.
            "tets": np.array of shape (M, 4), 0-based.
        }

        Args:
        - s_mesh (dict | None): The surface mesh dictionary. If None, uses self.s_mesh.
        """
        self._print_out("Starting volume mesh generation with TetGen...", "Normal")
        t0 = timer()
        if s_mesh is None:
            mesh_data = self.s_mesh
        else:
            mesh_data = s_mesh

        try:
            # 1. Create a PyVista PolyData (surface) object from your mesh data
            # PyVista requires a "padded" faces array: [3, v0, v1, v2, 3, v3, v4, v5, ...]
            num_faces = mesh_data["faces"].shape[0]
            faces_with_count = np.hstack((
                np.full((num_faces, 1), 3, dtype=np.int64), # The count '3'
                mesh_data["faces"]                               # The face indices
            )).flatten()
            
            surface_pv = pv.PolyData(mesh_data["verts"], faces=faces_with_count)

            # TetGen itself has no "minimum tet size" switch -- mindihedral/minratio below
            # are QUALITY controls (they reject slivers), not a size floor, and maxvolume is a
            # one-sided COARSENESS ceiling. Left alone, a small surface feature coming out of
            # the CGAL alpha-wrap step (a sharp corner, a near-duplicate vertex pair from two
            # primitives' boolean seam) forces TetGen to insert Steiner points and build very
            # small tets right there to conform to it exactly, which is where slow meshing and
            # exact-arithmetic failures both tend to originate. A user-supplied
            # min_mesh_characteristic_length instead sets an absolute point-merging tolerance:
            # any two surface vertices closer together than this are merged before
            # tetrahedralization, which removes exactly those degenerate small features and
            # gives min_mesh_characteristic_length real effect as a de facto minimum feature
            # size, complementing max_mesh_characteristic_length's ceiling below. See
            # docs/theory.md §7.1/§10.6.
            if self.min_mesh_characteristic_length:
                n_before = surface_pv.n_points
                surface_pv = surface_pv.clean(tolerance=self.min_mesh_characteristic_length, absolute=True)
                self._print_out(
                    f"  - min_mesh_characteristic_length: {1e3 * self.min_mesh_characteristic_length:.3g} mm "
                    f"(merged {n_before - surface_pv.n_points} surface point(s) closer together than this)",
                    "Verbose"
                )

            # 2. Run the Tetrahedralization using the 'tetgen' package API
            tet = tetgen.TetGen(surface_pv)

            # --- Set Meshing Parameters ---
            # Heuristic for max volume: 0.5% of total bounding box volume.
            # This is a critical parameter for controlling mesh coarseness.
            # You should tune this value (e.g., 0.005, 0.01) based on your needs.
            bounds = surface_pv.bounds
            vol = (bounds[1]-bounds[0]) * (bounds[3]-bounds[2]) * (bounds[5]-bounds[4])

            # Use a user-defined characteristic length if available, otherwise, heuristic
            if self.max_mesh_characteristic_length:
                # Simple cubic volume from characteristic length
                max_tet_volume = self.max_mesh_characteristic_length ** 3
            else:
                # Fallback to a heuristic based on bounding box
                max_tet_volume = vol * 0.01 # 1.0% of total volume

            self._print_out(f"Running TetGen tetrahedralize with:", "Verbose")
            self._print_out(f"  - min_dihedral: 10.0 (for FEA quality)", "Verbose")
            self._print_out(f"  - min_ratio: 1.5 (for FEA quality)", "Verbose")
            self._print_out(f"  - max_volume: {1e9 * max_tet_volume:.2e} mm³ (for coarseness)", "Verbose")

            tet.tetrahedralize(
                order=1,               # Linear tets
                mindihedral=10.0,      # Quality control (fixes slivers)
                minratio=1.5,          # Quality control (radius-edge ratio)
                maxvolume=max_tet_volume # Coarseness control
            )
            
            # Get the resulting mesh (a PyVista UnstructuredGrid)
            tet_mesh = tet.grid

            # 3. Extract nodes and tets *directly* in skfem format
            
            # Get node coordinates: shape (N, 3), 0-based
            nodes = tet_mesh.points

            # Get element data:
            # tet_mesh.cells is [4, n0, n1, n2, n3, 4, n4, n5, n6, n7, ...]
            # Reshape to (M, 5) and slice off the first column ('4')
            # The result is shape (M, 4), 0-based.
            tets = tet_mesh.cells.reshape(-1, 5)[:, 1:].astype(np.int64)

            # 3b. Element check (docs/theory.md §7.1). A tet flat to round-off has
            # a singular Jacobian: scikit-fem divides by its determinant and the
            # solve returns NaN. Such tets hold no volume, so they are dropped;
            # an INVERTED tet means TetGen's mesh folds over itself, and the
            # analysis is failed (the optimiser then retakes the step smaller).
            nodes, tets = self._check_tets(np.asarray(nodes, dtype=np.float64), tets)

            # 4. Store in the v_mesh dictionary in the new skfem-compatible format
            self.v_mesh = {
                "nodes": nodes,
                "tets": tets
            }

            if self.save_meshes:
                # Save the volume mesh to a .vtk file
                vtk_path = os.path.join(self.analysis_dir, f"{self.iter_name}_volume_mesh.vtk")
                tet_mesh.save(vtk_path)
                self._print_out(f"Volume mesh saved to {vtk_path}", "Verbose")

        except Exception as e:
            raise RuntimeError(f"PyVista/TetGen mesh generation failed: {e}")
        
        # Print the new, skfem-compatible shapes
        self._print_out(f"Volume mesh generated using TetGen with {nodes.shape[0]} nodes and {tets.shape[0]} elements", "Verbose")
        self._print_out(f"Volume mesh generation time: {timer() - t0:.2f} seconds", "Normal")

    def _check_tets(self, nodes: np.ndarray, tets: np.ndarray, flat_quality: float = 1e-8):
        """
        Drops the tets of a TetGen mesh that are flat to round-off and fails on
        inverted ones (docs/theory.md §7.1).

        The shape measure is q = 6 sqrt(2) V / l_rms^3, with V the signed
        volume and l_rms the root-mean-square edge length: q = 1 for a regular
        tet and 0 for a flat one, independent of its size. |q| < flat_quality
        is flat: its stiffness would be 0 x infinity. A flat tet is removed
        (nodes left in no tet are removed with it); its nodes stay shared by
        its neighbours, so the mesh stays connected through them. q <=
        -flat_quality is an inverted element, which a valid tetrahedralisation
        never holds, so the analysis is failed instead.

        Publishes degenerate_tets_removed and min_tet_quality (the smallest q
        kept), and returns (nodes, tets).
        """
        x = nodes[tets]
        e = np.stack([x[:, a] - x[:, b] for a, b in ((0, 1), (0, 2), (0, 3), (1, 2), (1, 3), (2, 3))], axis=1)
        l_rms = np.sqrt(np.mean(np.einsum('ijk,ijk->ij', e, e), axis=1))
        vol = np.einsum('ij,ij->i', np.cross(x[:, 1] - x[:, 0], x[:, 2] - x[:, 0]), x[:, 3] - x[:, 0]) / 6.0
        if np.count_nonzero(vol < 0.0) > tets.shape[0] // 2:
            # The opposite orientation convention: flip every tet.
            tets = tets[:, [0, 2, 1, 3]]
            vol = -vol
        q = 6.0 * np.sqrt(2.0) * vol / np.maximum(l_rms, 1e-300) ** 3
        inverted = q <= -flat_quality
        if np.any(inverted):
            raise RuntimeError(
                f"TetGen returned {int(np.count_nonzero(inverted))} inverted tetrahedra (of {tets.shape[0]}; "
                f"smallest shape measure {q.min():.2e}): the volume mesh folds over itself.")
        flat = q < flat_quality
        self.degenerate_tets_removed = int(np.count_nonzero(flat))
        if self.degenerate_tets_removed:
            keep = ~flat
            used = np.unique(tets[keep])
            remap = -np.ones(nodes.shape[0], dtype=np.int64)
            remap[used] = np.arange(used.size)
            nodes, tets, q = nodes[used], remap[tets[keep]], q[keep]
            self._print_out(f"Removed {self.degenerate_tets_removed} tetrahedra flat to round-off "
                            f"(shape measure below {flat_quality:g}).", "Normal")
        self.min_tet_quality = float(q.min()) if q.size else None
        return nodes, tets

    def _drop_floating_pieces(self):
        """
        Removes from the tetrahedral mesh every connected piece that contains
        no supported node (docs/theory.md §7.1, §9.1).

        Conformal meshes only the solid, phi >= level set. When the design
        splits -- a whole cross-section falling below the level set in the
        grey phase, or a detached fragment -- a piece that no Dirichlet
        condition reaches is free to move rigidly, the stiffness matrix is
        singular, and the solver returns meaningless displacements (negative
        or huge compliances). Removing such pieces leaves a well-posed model
        of the supported structure. What was removed is recorded:

        - floating_pieces_removed: number of pieces;
        - floating_volume_fraction: their share of the meshed volume;
        - floating_load_fraction: their share of the applied load (the norm
          of the load density at each tet's centroid times its volume). When
          this is positive the load path is CUT: part of the load cannot
          reach a support, and the analysis of what remains is not the
          design's response. The optimiser then does not use this
          iteration's strain energies (PCLS_Optimiser, docs/theory.md §8.4.1).

        The centroids of the removed tets are kept (`_removed_centroids`) so
        that the voxels of a removed piece receive zero strain energy in the
        voxel mapping, instead of the value of the nearest supported tet.
        """
        self.floating_pieces_removed = 0
        self.floating_volume_fraction = 0.0
        self.floating_load_fraction = 0.0
        self._removed_centroids = np.zeros((0, 3))
        nodes = np.asarray(self.v_mesh["nodes"], dtype=np.float64)
        tets = np.asarray(self.v_mesh["tets"], dtype=np.int64)
        n_nodes = nodes.shape[0]
        if tets.shape[0] == 0:
            return
        # Node graph: each tet joins its first node to the other three.
        rows = np.repeat(tets[:, 0], 3)
        cols = tets[:, 1:].ravel()
        graph = coo_matrix((np.ones(rows.size, dtype=np.int8), (rows, cols)), shape=(n_nodes, n_nodes))
        n_pieces, piece = connected_components(graph, directed=False)
        if n_pieces <= 1:
            return
        _voxel_index = self._voxel_of

        fixed = np.asarray(self.csg_data["fixed_nodes"]).any(axis=-1)[_voxel_index(nodes)]
        supported = np.zeros(n_pieces, dtype=bool)
        supported[np.unique(piece[fixed])] = True
        if not np.any(supported):
            # Nothing is supported at all: leave the mesh alone and let _fea
            # report the missing boundary conditions.
            return
        tet_piece = piece[tets[:, 0]]
        floating = ~supported[tet_piece]
        if not np.any(floating):
            return
        x = nodes[tets]                                      # (M, 4, 3)
        volume = np.abs(np.einsum('ij,ij->i', np.cross(x[:, 1] - x[:, 0], x[:, 2] - x[:, 0]),
                                  x[:, 3] - x[:, 0])) / 6.0
        centroids = x.mean(axis=1)
        # Each voxel's force is spread over the mesh volume in its cell (as in
        # _fea), so a tet's share of the load is its share of that volume.
        load = np.linalg.norm(np.asarray(self.csg_data["load_vectors"], dtype=np.float64), axis=-1)
        cells = _voxel_index(centroids)
        flat = np.ravel_multi_index(cells, load.shape)
        cell_volume = np.bincount(flat, weights=volume, minlength=load.size)
        tet_load = load.ravel()[flat] * volume / np.maximum(cell_volume[flat], 1e-300)
        self.floating_pieces_removed = int(np.count_nonzero(~supported))
        self.floating_volume_fraction = float(volume[floating].sum() / max(volume.sum(), 1e-300))
        total_load = float(tet_load.sum())
        self.floating_load_fraction = float(tet_load[floating].sum() / total_load) if total_load > 0 else 0.0
        self._removed_centroids = centroids[floating]
        keep = ~floating
        used = np.unique(tets[keep])
        remap = -np.ones(n_nodes, dtype=np.int64)
        remap[used] = np.arange(used.size)
        self.v_mesh = {"nodes": nodes[used], "tets": remap[tets[keep]]}
        self._print_out(
            f"Removed {self.floating_pieces_removed} piece(s) that no support reaches: "
            f"{100 * self.floating_volume_fraction:.2f}% of the meshed volume, carrying "
            f"{100 * self.floating_load_fraction:.2f}% of the load"
            + (" -- the load path is cut." if self.floating_load_fraction > 0 else "."),
            "Minimal" if self.floating_load_fraction > 0 else "Normal")

    def _C(self, T):
        """
        Linear elastic constitutive law.
        """
        return 2. * self.mu * T + self.lam * eye(trace(T), T.shape[0])

    def _voxel_of(self, points: np.ndarray):
        """
        Voxel indices (i, j, k) of points in metres: the voxel whose sample
        point (i h, j h, k h) is nearest, clipped to the grid. Marching cubes
        samples phi at those points, so voxel i's cell is [(i - 1/2) h,
        (i + 1/2) h) -- the cell the supports, the loads, the floating-piece
        check and the stiffness density all use (docs/theory.md §7.1).
        """
        pts = np.asarray(points, dtype=np.float64)
        nx, ny, nz = np.asarray(self.csg_data["fixed_nodes"]).shape[:3]
        idx = np.rint(pts / self.voxel_size_m).astype(np.int64)
        return (np.clip(idx[:, 0], 0, nx - 1), np.clip(idx[:, 1], 0, ny - 1), np.clip(idx[:, 2], 0, nz - 1))

    def _tet_scale(self):
        """Per-tet stiffness density broadcastable against a (3, 3, Ne, Nqp) field (1.0 if none)."""
        if self._tet_rho_m is None:
            return 1.0
        return self._tet_rho_m[None, None, :, None]

    def _stress_from_strain(self, strain_field: np.ndarray):
        """
        Actual constitutive law, exposed under the name the shared buckling
        machinery in the Analysis base class expects: sigma = rho_m,e C0 eps
        with the tet's stiffness density rho_m,e (docs/theory.md §7.1), or
        C0 eps when the model has the full modulus everywhere.
        """
        return self._tet_scale() * self._C(strain_field)

    def _stress_from_strain_prime(self, strain_field: np.ndarray):
        """
        d(constitutive law)/d(rho_m,e): the base-material law C0 eps.

        With the material interpolation the tet's law is rho_m,e C0, linear
        in its stiffness density, so this is exact -- the same arrangement as
        `Regular`, whose sensitivity the optimiser chains through
        d(rho_m)/d(rho) and rho'(phi). The mesh itself also moves with the
        design (the boundary is the level set); that shape dependence is not
        in this derivative, which is why Conformal's buckling sensitivity
        remains an approximation (docs/theory.md §9.3 item 2).
        """
        return self._C(strain_field)

    def _num_elements(self):
        return self.v_mesh["tets"].shape[0]

    def _map_sensitivity_to_voxels(self, element_values: np.ndarray):
        """
        Maps a per-tet scalar sensitivity onto the voxel grid.

        Uses linear interpolation from tet centroids rather than the
        nearest-neighbour lookup used for strain energy density: buckling
        sensitivity tends to be more spatially concentrated than SED, and
        nearest-neighbour interpolation visibly "blocks" it onto whichever
        voxel is nearest a tet centroid. Voxels outside the convex hull of
        the tet centroids (linear interpolation returns NaN there) fall
        back to nearest-neighbour. See implementation plan section 3.2.
        """
        element_centroids = np.mean(self.v_mesh["nodes"][self.v_mesh["tets"]], axis=1)

        grid_shape = self._buckling_voxel_shape
        grid_spacing = np.ones(3) * self.voxel_size_m

        i, j, k = np.indices(grid_shape)  # type: ignore
        query_points = np.stack([
            i.ravel() * grid_spacing[0],
            j.ravel() * grid_spacing[1],
            k.ravel() * grid_spacing[2],
        ], axis=-1)

        voxel_flat = griddata(
            element_centroids, element_values, query_points,
            method='linear', fill_value=np.nan
        )
        nan_mask = np.isnan(voxel_flat)
        if np.any(nan_mask):
            voxel_flat[nan_mask] = griddata(
                element_centroids, element_values, query_points[nan_mask],
                method='nearest', fill_value=0.0
            )

        return voxel_flat.reshape(grid_shape)

    def _map_nodal_field_to_grid_points(self, node_values: np.ndarray) -> np.ndarray:
        """
        Maps a per-mesh-NODE scalar (displacement magnitude, or a buckling
        mode-shape magnitude, both computed at TetGen mesh nodes) onto the
        (nx+1, ny+1, nz+1) voxel-grid NODE lattice, for the standardised
        iteration ".vtr" (docs/theory.md §7.6).

        Unlike `Regular`/`ConjugateGradient`, whose FEA nodes ARE the
        voxel-grid nodes exactly, `Conformal`'s mesh is rebuilt from scratch
        every iteration and its nodes sit wherever marching-cubes/alpha-wrap/
        TetGen put them -- nowhere near a regular lattice in general. This is
        therefore necessarily an interpolation, not an exact lookup, using
        the same linear-with-nearest-neighbour-fallback scheme as
        `_map_sensitivity_to_voxels`, just evaluated at grid NODE positions
        (i*h, j*h, k*h for i in 0..nx inclusive) rather than voxel/element
        centroid positions -- so that 'Displacement_Magnitude' and
        'Buckling_Mode_i_Magnitude' end up as genuine PointData on the
        output grid, matching the other two backends' field, even though
        the values themselves are approximate here.
        """
        nodes = self.v_mesh["nodes"]  # (N, 3) metres
        nx_e, ny_e, nz_e = self._buckling_voxel_shape
        grid_node_shape = (nx_e + 1, ny_e + 1, nz_e + 1)
        h = self.voxel_size_m

        i, j, k = np.indices(grid_node_shape)  # type: ignore
        query_points = np.stack([
            i.ravel() * h, j.ravel() * h, k.ravel() * h,
        ], axis=-1)

        flat = griddata(nodes, node_values, query_points, method='linear', fill_value=np.nan)
        nan_mask = np.isnan(flat)
        if np.any(nan_mask):
            flat[nan_mask] = griddata(
                nodes, node_values, query_points[nan_mask], method='nearest', fill_value=0.0
            )
        return flat.reshape(grid_node_shape)

    def _map_mode_shapes_to_grid(self, mode_shapes: np.ndarray) -> dict:
        """
        See `Analysis._map_mode_shapes_to_grid`. `mode_shapes` here is
        (Ndofs, k) on `self.basis`'s DOF numbering -- the same one `self.u`
        uses -- so each mode's per-node magnitude is computed exactly like
        `_fea`'s own displacement magnitude, then mapped with
        `_map_nodal_field_to_grid_points`.
        """
        fields = {}
        for m in range(mode_shapes.shape[1]):
            mode_xyz = np.array([mode_shapes[self.basis.nodal_dofs[i], m] for i in range(3)])
            node_mag = np.linalg.norm(mode_xyz, axis=0)
            fields[f"Buckling_Mode_{m + 1}_Magnitude"] = self._map_nodal_field_to_grid_points(node_mag)
        return fields

    def _compute_stress_field(self, displacement_field: np.ndarray):
        """
        Computes the full quadrature-point-resolved stress and strain
        tensor fields (each shape (3,3,Ne,Nqp)) for any full-length
        displacement-like field on `self.basis` (the primal solution, or a
        buckling mode shape), with the ACTUAL constitutive law (the tet's
        stiffness density included). Factored out of
        `_strain_energy_densities` so it can be reused by the buckling
        sensitivity machinery.
        """
        strain_field = sym_grad(self.basis.interpolate(displacement_field))
        stress_field = self._stress_from_strain(strain_field)
        return stress_field, strain_field

    def _tet_stiffness_density(self, phi: np.ndarray, level_set_value: float,
                               stiffness_density: np.ndarray) -> np.ndarray:
        """
        The stiffness density of every tet (docs/theory.md §4.8, eq. 4.16): that
        of the design voxel whose sample point is nearest the tet's centroid.
        Marching cubes places the surface between samples, so every tet of the
        solid lies nearer a sample inside the level set -- except in the thin
        shell that the alpha wrap's offset adds outside it, whose tets take
        the nearest sample that is inside. A binary design (phi at its wells)
        therefore gets the full modulus in every tet, exactly the model of
        `material_interpolation=False`; a grey one is as weak where it is grey
        as the density backends' model.
        """
        nodes = np.asarray(self.v_mesh["nodes"], dtype=np.float64)
        tets = np.asarray(self.v_mesh["tets"], dtype=np.int64)
        centroids = nodes[tets].mean(axis=1)
        shape = np.asarray(phi.shape)
        idx = np.rint(centroids / self.voxel_size_m).astype(np.int64)
        idx = np.clip(idx, 0, shape - 1)
        inside = np.asarray(phi) >= level_set_value
        if not np.any(inside):
            return np.ones(tets.shape[0])
        # For every voxel, the nearest voxel that is inside the level set.
        _, near = ndimage.distance_transform_edt(~inside, return_indices=True)
        si = near[0][idx[:, 0], idx[:, 1], idx[:, 2]]
        sj = near[1][idx[:, 0], idx[:, 1], idx[:, 2]]
        sk = near[2][idx[:, 0], idx[:, 1], idx[:, 2]]
        rho_m = np.asarray(stiffness_density, dtype=np.float64)[si, sj, sk]
        return np.clip(rho_m, 1e-9, 1.0)

    def uniform_material_compliance(self) -> float:
        """
        Compliance f.u of the current mesh with the FULL modulus in every tet:
        the analysis of the geometry the run would export (phi >= level set
        made solid), for comparison with the model's own compliance (eq.
        8.4i). Reuses the mesh, basis, load vector and supports of the last
        _fea call and leaves its solution, stiffness matrix and fields alone.
        """
        if self._tet_rho_m is None:
            return float(self.compliance)

        @BilinearForm
        def stiffness(u, v, w):
            return ddot(self._C(sym_grad(u)), sym_grad(v))

        k = stiffness.assemble(self.basis)
        u = solve(*condense(k, self._f, x=self.basis.zeros(), D=self.fixed_dofs), solver=pardiso_solve)
        return float(np.dot(self._f, u))

    def uniform_material_buckling(self, **kwargs):
        """
        Linear buckling of the geometry the run would export: the current mesh
        with the FULL modulus in every tet, the buckling counterpart of
        `uniform_material_compliance` (docs/theory.md §6.10). The penalised
        model's stiffness (and so its load factor) can overstate that of the
        exported part. Re-solves the primal problem at full modulus, runs
        `run_buckling(**kwargs)` on it and restores the model's own stiffness
        matrix and displacements. With no material interpolation it is
        `run_buckling` itself. Returns a BucklingResults.
        """
        if self._tet_rho_m is None:
            return self.run_buckling(**kwargs)

        @BilinearForm
        def stiffness(u, v, w):
            return ddot(self._C(sym_grad(u)), sym_grad(v))

        saved = (self._tet_rho_m, self.K, self.u)
        try:
            self._tet_rho_m = None
            k = stiffness.assemble(self.basis)
            self.u = solve(*condense(k, self._f, x=self.basis.zeros(), D=self.fixed_dofs), solver=pardiso_solve)
            self.K = k
            return self.run_buckling(**kwargs)
        finally:
            self._tet_rho_m, self.K, self.u = saved

    def _fea(self):
        """
        Sets up and runs the finite element analysis using scikit-fem.
        """
        self._print_out("Starting finite element analysis (FEA) using scikit-fem...", "Normal")
        t0 = timer()

        # 1) Load the volume mesh into skfem
        verts = self.v_mesh["nodes"].T # skfem expects shape (3, N)
        tets = self.v_mesh["tets"].T # skfem expects shape (4, M)

        try:
            m = MeshTet(verts, tets) # type: ignore
        except Exception as e:
            raise RuntimeError(f"Failed to create skfem MeshTet: {e}")
        
        if not 'tet' in m.elem.__name__.lower():
            raise NotImplementedError(f"Mesh element type {m.elem.__name__} not supported.")

        bbox = np.array([[np.min(m.p[0]), np.max(m.p[0])],
                [np.min(m.p[1]), np.max(m.p[1])],
                [np.min(m.p[2]), np.max(m.p[2])],
        ]) # Bounding box (x, y, z), shape (3, 2)
        mesh_info = (f"Loaded mesh with:\n"
                     f"  Nodes: {m.p.shape[1]}\n"
                     f"  Elements: {m.t.shape[1]} (Tets)\n"
                     f"  Bounding box X: [{bbox[0][0]:.4e}, {bbox[0][1]:.4e}]\n"
                     f"  Bounding box Y: [{bbox[1][0]:.4e}, {bbox[1][1]:.4e}]\n"
                     f"  Bounding box Z: [{bbox[2][0]:.4e}, {bbox[2][1]:.4e}]")
        self._print_out(mesh_info, "Verbose")

        # 2) Find Fixed DOFs using CSG Voxel Grid
        nx, ny, nz, _ = self.csg_data["fixed_nodes"].shape # Unpack 4 dims

        # Map conformal floating point nodes (metres) back to voxel indices:
        # the voxel whose sample point is nearest, i.e. the cell of width h
        # CENTRED on the sample, which is how marching cubes placed the
        # geometry (docs/theory.md §7.1). floor(x/h) -- used before this
        # revision -- shifted every cell by half a voxel: a support one voxel
        # thick clamped 1.5 voxels at the domain's low faces and 0.5 at its
        # high faces, and a load on the last voxel layer was half outside the
        # mesh.
        px, py, pz = self._voxel_of(m.p.T)

        # Extract the boolean mask of shape (N_nodes, 3) tracking (Tx, Ty, Tz)
        fixed_node_mask = self.csg_data["fixed_nodes"][px, py, pz]

        # 3) Setup the FEA basis. One quadrature point per tet: with linear
        # tets the strain, the stress and hence the stiffness and geometric-
        # stiffness integrands are constant in each element, so the centroid
        # rule integrates them exactly -- the default (4 points) computed the
        # same matrices with 4x the memory and time, which on a fine TetGen
        # mesh (several 10^5 tets) is gigabytes. The body force is sampled at
        # the centroid, i.e. each tet takes the load density of the voxel it
        # lies in (docs/theory.md §7.1).
        e = ElementVector(m.elem())
        self.basis = Basis(m, e, intorder=1)

        # 4) Define material properties, constraints, and forces
        self.lam, self.mu = lame_parameters(self.youngs_modulus, self.poisson_ratio)

        @BilinearForm
        def stiffness(u, v, w):
            return ddot(self._C(sym_grad(u)), sym_grad(v))

        # The tets' stiffness density (docs/theory.md §7.1): the material
        # interpolation of the density backends, applied to the body-fitted
        # mesh. None = the full modulus everywhere.
        @BilinearForm
        def stiffness_interpolated(u, v, w):
            return w.rho_m * ddot(self._C(sym_grad(u)), sym_grad(v))

        # Loads (docs/theory.md §7.1, eq. 7.0a). The CSG gives a force per
        # voxel [N]. Each voxel's force is spread uniformly over the part of
        # the mesh that lies in its cell, as a body force density
        # f_voxel / V_cell,mesh [N/m^3], so that the mesh receives every
        # voxel's force in full -- as the density backends, which put it on
        # the voxel's 8 nodes, do. Dividing by the voxel's full volume h^3
        # instead (before this revision) applied only the covered fraction of
        # the force: half of it on the last voxel layer of the domain, where
        # marching cubes puts the surface half a voxel beyond the sample.
        # Each tet belongs to the cell of its centroid.
        load_vectors = np.asarray(self.csg_data["load_vectors"], dtype=np.float64)
        tet_x = m.p[:, m.t]                                  # (3, 4, Ne)
        tet_centroid = tet_x.mean(axis=1).T                  # (Ne, 3)
        e1, e2, e3 = (tet_x[:, 1] - tet_x[:, 0]), (tet_x[:, 2] - tet_x[:, 0]), (tet_x[:, 3] - tet_x[:, 0])
        tet_volume = np.abs(np.einsum('ij,ij->j', np.cross(e1.T, e2.T).T, e3)) / 6.0
        ti, tj, tk = self._voxel_of(tet_centroid)
        cell = np.ravel_multi_index((ti, tj, tk), (nx, ny, nz))
        cell_volume = np.bincount(cell, weights=tet_volume, minlength=nx * ny * nz)
        voxel_force = load_vectors.reshape(-1, 3)
        loaded = np.any(voxel_force != 0.0, axis=1)
        missing = loaded & (cell_volume <= 0.0)
        with np.errstate(divide="ignore", invalid="ignore"):
            voxel_density = np.where(cell_volume[:, None] > 0.0, voxel_force / cell_volume[:, None], 0.0)
        tet_body_force = voxel_density[cell]                 # (Ne, 3) N/m^3
        if np.any(missing):
            # A loaded voxel with no tet centroid in its cell (a corner cell
            # the surface only clips, or a voxel whose design is void): its
            # force goes to the tet nearest its sample point, so that the
            # mesh still carries the whole load.
            from scipy.spatial import cKDTree
            miss_idx = np.flatnonzero(missing)
            sample = np.stack(np.unravel_index(miss_idx, (nx, ny, nz)), axis=1) * self.voxel_size_m
            dist, nearest = cKDTree(tet_centroid).query(sample)
            np.add.at(tet_body_force, nearest, voxel_force[miss_idx] / tet_volume[nearest, None])
            far = dist > 1.5 * self.voxel_size_m
            if np.any(far):
                self._print_out(
                    f"Warning: {int(np.count_nonzero(far))} loaded voxel(s) lie more than 1.5 voxels from "
                    f"the mesh (their design is void); their load was moved to the nearest tet.", "Minimal")
        self.requested_load_N = voxel_force.sum(axis=0)
        self.applied_load_N = (tet_body_force * tet_volume[:, None]).sum(axis=0)
        n_qp_f = self.basis.X.shape[1]
        bfx = np.repeat(tet_body_force[:, 0:1], n_qp_f, axis=1)
        bfy = np.repeat(tet_body_force[:, 1:2], n_qp_f, axis=1)
        bfz = np.repeat(tet_body_force[:, 2:3], n_qp_f, axis=1)

        @LinearForm
        def force(v, w):
            """Integrates the per-tet body force density (each voxel's force spread over its cell)."""
            return v[0] * w.bx + v[1] * w.by + v[2] * w.bz

        # ASSEMBLE MATRICES (Fix applied here!)
        self._print_out("Assembling stiffness matrix (k)...", "Verbose")
        if self._tet_rho_m is None:
            k = stiffness.assemble(self.basis)
        else:
            n_qp = self.basis.X.shape[1]
            k = stiffness_interpolated.assemble(
                self.basis, rho_m=np.repeat(self._tet_rho_m[:, None], n_qp, axis=1))

        self._print_out("Assembling force vector (f)...", "Verbose")
        f = force.assemble(self.basis, bx=bfx, by=bfy, bz=bfz)

        # 5) Apply boundary conditions
        u = self.basis.zeros()
        
        if self.basis.nodal_dofs is None:
            raise RuntimeError("Basis nodal DOFs are missing.")
            
        fixed_dofs_list = []
        for i in range(3): # For Tx, Ty, Tz
            nodes_fixed_for_dof = np.where(fixed_node_mask[:, i])[0]
            fixed_dofs_list.extend(self.basis.nodal_dofs[i, nodes_fixed_for_dof])
            
        fixed_dofs = np.unique(fixed_dofs_list)
        
        if fixed_dofs.size == 0:
            raise RuntimeError("No fixed DOFs found. Check CSG constraints.")
            
        self._print_out(f"Applying Dirichlet BCs on {fixed_dofs.size} DOFs using CSG map.", "Verbose")

        # 6) Solve the system
        if fixed_dofs.size == self.basis.N:
            error_msg = (
                "FATAL ERROR: All DOFs in the mesh are fixed. The structure has completely severed, separating the loaded regions from the constrained regions.\n\n"
                "Because strain energy cannot transfer across a broken void, the optimiser cannot recover the topology. The run is being aborted.\n\n"
                "Suggestions to prevent this in the next run:\n"
                "  1. Decrease the time step ('dt') to slow down the evolution.\n"
                "  2. Increase the initial 'mu1' or its decay factor to decrease the pressure on losing volume so fast.\n"
                "  3. Check if the target volume is simply too low for the complexity of the applied loads."
            )
            raise RuntimeError(error_msg)
        else:
            # 6) Solve the system
            self._print_out("Solving the linear system using PARDISO...", "Verbose")
            # Pass pardiso_solve to the solver argument
            self.u = solve(*condense(k, f, x=u, D=fixed_dofs), solver=pardiso_solve)
            # A singular or near-singular stiffness matrix (a degenerate
            # element, a piece held by a single node) gives non-finite or
            # non-physical displacements. The analysis is failed rather than
            # handing the optimiser NaN strain energies (docs/theory.md §7.1).
            c_check = float(np.dot(f, self.u))
            if not np.all(np.isfinite(self.u)) or not np.isfinite(c_check) or c_check <= 0.0:
                raise RuntimeError(
                    f"the FE solve returned {'non-finite displacements' if not np.all(np.isfinite(self.u)) else 'a non-physical compliance'} "
                    f"(f.u = {c_check:.3e} J) on {m.t.shape[1]} tetrahedra (smallest shape measure "
                    f"{self.min_tet_quality if self.min_tet_quality is not None else float('nan'):.2e}): the "
                    f"stiffness matrix is singular or near-singular.")

        xyz_u = 1e3 * np.array([self.u[self.basis.nodal_dofs[i]] for i in range(3)]) # Displacement field (converted to mm) # type: ignore

        self._print_out(f"Maximum displacement (mm) X: {np.max(xyz_u[0]):.6f}, Y: {np.max(xyz_u[1]):.6f}, Z: {np.max(xyz_u[2]):.6f}", "Verbose")

        # Store the assembled stiffness matrix and the fixed-DOF set so that
        # `run_buckling` can reuse them (and, via `pardiso_solve`'s
        # automatic factorisation caching, the condensed factorisation)
        # without a full re-assembly.
        self.K = k
        self.fixed_dofs = fixed_dofs
        self._f = f

        # The EXACT compliance, c = f . u, in joules. For a linear elastic
        # structure this is precisely twice the total stored strain energy, and
        # both vectors are already in hand, so it costs one dot product. The
        # optimiser reads this attribute in preference to reconstructing a
        # compliance-like quantity from the voxel-mapped SED field, which for
        # the ersatz backends double-counts the density and is not in joules.
        self.compliance = float(np.dot(f, self.u))
        self._print_out(f"Compliance (f.u): {self.compliance:.6e} J", "Verbose")

        if self.save_meshes:
            # Note: Changed 'u' to 'self.u' here to ensure we translate using the solved displacements
            m_deformed = m.translated(self.visualisation_scale * self.u[self.basis.nodal_dofs])
            output_path = os.path.join(self.analysis_dir, f"{self.iter_name}_deformed_mesh.vtk")
            m_deformed.save(output_path)
            self._print_out(f"Deformed mesh saved to {output_path}", "Verbose")

            # Standardised iteration .vtr (docs/theory.md §7.5/§7.6): the
            # deformed mesh's actual displaced coordinates above stay as the
            # authoritative unstructured-mesh output, but a voxel-grid-mapped
            # displacement MAGNITUDE also goes into `_pending_point_data`, so
            # ParaView's "Warp By Vector" on the grid isn't available here
            # (there is no vector, only magnitude, by request) but a
            # cross-backend-comparable scalar deflection field is.
            disp_xyz_m = np.array([self.u[self.basis.nodal_dofs[i]] for i in range(3)])
            node_disp_mag = np.linalg.norm(disp_xyz_m, axis=0)
            self._pending_point_data["Displacement_Magnitude"] = self._map_nodal_field_to_grid_points(node_disp_mag)

        self._print_out(f"FEA completed in {timer() - t0:.2f} seconds. It found a maximum deflection of { np.max(np.linalg.norm(xyz_u, axis=0)):.6f} mm.", "Normal")

    def _map_elements_to_voxels(self, element_values: np.ndarray, grid_shape) -> np.ndarray:
        """
        Voxel averages of per-tet values (docs/theory.md §7.1, step 6, eq. 7.0c).

        Each tet is sampled at the 11 barycentric points of _TET_SAMPLE_POINTS,
        each carrying |T|/11 of its volume, and a sample belongs to the voxel
        whose sample point is nearest (`_voxel_of`). A voxel the mesh covers
        gets the volume-weighted mean of the tets over the covered part of its
        cell. A voxel it does not cover gets the value of the nearest covered
        voxel -- a driving signal just outside the boundary -- or 0 when a
        voxel of a piece removed by _drop_floating_pieces is nearer.

        This replaced taking the value of the one tet whose centroid is
        nearest the voxel's sample point: TetGen rebuilds the mesh from
        scratch at every iteration, and a change of phi of 1e-4 moved that
        single-tet value by 11 % (median; 46 % at the 90th percentile) at the
        surface of a converging design, noise that kept boundary voxels
        jittering and the design from settling. The cell average moves by
        3 % (10 %) and agrees better with the voxel models.

        element_values: (Ne,) or (Ne, k). Returns grid_shape or grid_shape + (k,).
        """
        shape = tuple(int(v) for v in grid_shape)
        n_vox = int(np.prod(shape))
        vals = np.asarray(element_values, dtype=np.float64)
        squeeze = vals.ndim == 1
        if squeeze:
            vals = vals[:, None]
        nodes = np.asarray(self.v_mesh["nodes"], dtype=np.float64)
        tets = np.asarray(self.v_mesh["tets"], dtype=np.int64)
        n_q = _TET_SAMPLE_POINTS.shape[0]
        den = np.zeros(n_vox)
        num = np.zeros((n_vox, vals.shape[1]))
        chunk = 100_000
        for start in range(0, tets.shape[0], chunk):
            x = nodes[tets[start:start + chunk]]                      # (m, 4, 3)
            vol = np.abs(np.einsum('ij,ij->i', np.cross(x[:, 1] - x[:, 0], x[:, 2] - x[:, 0]),
                                   x[:, 3] - x[:, 0])) / 6.0
            pts = np.einsum('qn,mnd->mqd', _TET_SAMPLE_POINTS, x).reshape(-1, 3)
            flat = np.ravel_multi_index(self._voxel_of(pts), shape)
            w = np.repeat(vol / n_q, n_q)
            den += np.bincount(flat, weights=w, minlength=n_vox)
            v = np.repeat(vals[start:start + chunk], n_q, axis=0)
            for c in range(vals.shape[1]):
                num[:, c] += np.bincount(flat, weights=w * v[:, c], minlength=n_vox)
        covered = den > 0.0
        out = np.zeros((n_vox, vals.shape[1]))
        out[covered] = num[covered] / den[covered, None]
        # Sources for the voxels the mesh does not cover: covered voxels, and
        # (with value 0) the voxels of removed floating pieces.
        source = covered.copy()
        removed = getattr(self, "_removed_centroids", None)
        if removed is not None and len(removed):
            rflat = np.ravel_multi_index(self._voxel_of(np.asarray(removed, dtype=np.float64)), shape)
            source[rflat] = True
        out = out.reshape(shape + (vals.shape[1],))
        source = source.reshape(shape)
        if np.any(source) and not np.all(source):
            _, idx = ndimage.distance_transform_edt(~source, return_indices=True)
            out = out[idx[0], idx[1], idx[2]]
        return out[..., 0] if squeeze else out

    def _strain_energy_densities(self, phi):
        """
        Calculates the strain energy densities in the mesh before mapping them onto the voxels in the phase field.
        
        Args:
        - phi (np.array): Piecewise-constant phase field.

        Returns:
        - np.ndarray: Strain energy densities for each voxel in phi.
        """
        stress_field, strain_field = self._compute_stress_field(self.u)

        # Base-material strain energy density F0 = 0.5 * (C0 eps : eps) at
        # each integration point, averaged over the element. With the
        # material interpolation the stored energy density is rho_m,e * F0;
        # the optimiser receives F0 (the quantity of Wei & Wang's driving
        # term, docs/theory.md §4.5) and applies d(rho_m)/d(rho) itself.
        base_stress_field = self._C(strain_field)
        sed_qp = 0.5 * np.einsum('ijkl,ijkl->kl', base_stress_field, strain_field)
        element_sed = np.mean(sed_qp, axis=1) # Shape (n_elements,)
        tet_rho_m = self._tet_rho_m if self._tet_rho_m is not None else np.ones_like(element_sed)
        element_stored_sed = tet_rho_m * element_sed

        if self.save_meshes:
            try:
                self._print_out("Saving volume mesh with strain energy density (SED) data...", "Verbose")
                
                # 1. Get nodes and tets from the stored mesh
                nodes = self.v_mesh["nodes"]
                tets = self.v_mesh["tets"] # Shape (M, 4)
                num_tets = tets.shape[0]

                # 2. Check for data consistency
                if num_tets != element_sed.shape[0]:
                    self._print_out(f"Warning: Mismatch between tet count ({num_tets}) and SED count ({element_sed.shape[0]}). Skipping SED mesh save.", "Minimal")
                else:
                    # 3. Create a PyVista UnstructuredGrid object
                    # This is the modern, preferred way to create a mesh from
                    # a single element type.
                    cells_dict = {pv.CellType.TETRA: tets}
                    mesh_with_sed = pv.UnstructuredGrid(cells_dict, nodes)

                    # 4. Add the calculated strain energy density as cell data
                    # This will allow visualization tools (like ParaView or PyVista)
                    # to color the mesh by this value.
                    mesh_with_sed.cell_data["Strain Energy Density"] = element_stored_sed
                    mesh_with_sed.cell_data["Base Strain Energy Density"] = element_sed
                    mesh_with_sed.cell_data["Stiffness Density"] = tet_rho_m

                    # 5. Define path and save
                    output_path = os.path.join(self.analysis_dir, f"{self.iter_name}_sed_mesh.vtk")
                    mesh_with_sed.save(output_path)
                    
                    self._print_out(f"Volume mesh with SED saved to {output_path}", "Verbose")

            except Exception as e:
                # Log the error but don't crash the whole analysis
                self._print_out(f"Warning: Failed to save SED volume mesh. Error: {e}", "Minimal")

        grid_shape = np.array(phi.shape)

        # The base-material strain energy density on the voxel grid: the average
        # over each voxel's cell (eq. 7.0c), with the stored density and, for
        # the saved iterations, the von Mises stress mapped the same way.
        # Voxels of pieces removed by _drop_floating_pieces get zero energy,
        # not the value of the nearest supported tet.
        element_stress_avg = np.mean(stress_field, axis=-1)  # (3, 3, Ne)
        sxx, syy, szz = element_stress_avg[0, 0], element_stress_avg[1, 1], element_stress_avg[2, 2]
        syz, sxz, sxy = element_stress_avg[1, 2], element_stress_avg[0, 2], element_stress_avg[0, 1]
        element_von_mises = _von_mises_from_voigt(np.stack([sxx, syy, szz, syz, sxz, sxy], axis=-1))
        mapped = self._map_elements_to_voxels(
            np.stack([element_sed, element_stored_sed, element_von_mises], axis=1), grid_shape)
        voxel_sed = mapped[..., 0]

        self._print_out(f"Min/Max Strain Energy Density: {np.min(voxel_sed):.4e}/{np.max(voxel_sed):.4e}", "Verbose")

        # --- Standardised iteration .vtr (docs/theory.md §7.5/§7.6) ---
        #
        # Conformal has no ersatz density field at the analysis level (§7.1),
        # so 'Phi' takes the role 'Density' plays for Regular/ConjugateGradient,
        # and 'Active_Voxels' is the level-set threshold mask rather than a
        # rho >= 0.5 mask. Both are already exactly on this grid -- no mapping
        # needed, unlike SED/Von_Mises which are mapped from tet centroids.
        if self.save_meshes:
            grid_shape_t = tuple(int(v) for v in grid_shape)
            self._grid_shape = grid_shape_t
            self._grid_hx = self._grid_hy = self._grid_hz = self.voxel_size_m

            if self._tet_rho_m is not None:
                # Stored energy density, like the other backends' field of this
                # name; the base-material value the optimiser uses is beside it.
                self._pending_cell_data["Strain_Energy_Density"] = mapped[..., 1].astype(np.float32)
                self._pending_cell_data["Base_Strain_Energy_Density"] = voxel_sed.astype(np.float32)
                if self._stiffness_density_field is not None:
                    self._pending_cell_data["Stiffness_Density"] = np.asarray(
                        self._stiffness_density_field, dtype=np.float32)
            else:
                self._pending_cell_data["Strain_Energy_Density"] = voxel_sed.astype(np.float32)
            self._pending_cell_data["Phi"] = np.asarray(phi, dtype=np.float32)
            self._pending_cell_data["Active_Voxels"] = (
                np.asarray(phi) >= self._level_set_value
            ).astype(np.float32)

            # Von Mises stress: the same cell average as SED above (both are
            # piecewise-constant-per-tet quantities).
            self._pending_cell_data["Von_Mises_Stress"] = mapped[..., 2].astype(np.float32)

        return voxel_sed

    def run_analysis(self, field, level_set_value, stiffness_density=None):
        """
        Runs numerical analysis on the provided field.

        The steps involved are:
        1) Generate surface mesh from the abstract field using marching cubes.
        2) Generate volume mesh using Tetgen.
        3) Remove pieces no support reaches; give every tet its stiffness density.
        4) Set up and run the finite element analysis.
        
        Args:
        - field (np.ndarray): The phase field.
        - level_set_value (float): The level set value to extract the surface mesh.
        - stiffness_density (np.ndarray | None): the design's stiffness density
          rho_m per voxel (same shape as field). With material_interpolation
          on, each tet takes the value of its voxel (docs/theory.md §7.1);
          None (the exported-geometry analyses) gives every tet the full
          modulus.

        Returns:
        - voxel_sed (np.ndarray): the BASE-material strain energy density F0
          mapped onto every voxel of the phase field (equal to the stored
          energy density when every tet has the full modulus).
        - (float): The volume of the surface mesh.
        """
        self._print_out("Starting analysis run...", "Normal")
        t0 = timer()

        # Remember the voxel grid shape for the buckling sensitivity mapping
        # (run_buckling has no direct access to 'phi' otherwise), and the
        # level-set value for the standardised iteration .vtr's Active_Voxels
        # field in _strain_energy_densities (docs/theory.md §7.5/§7.6).
        self._buckling_voxel_shape = field.shape
        self._level_set_value = level_set_value

        # Generate a surface mesh
        self._surface_mesh(field, level_set_value)

        # Generate a volume mesh
        self._volume_mesh()

        # Remove pieces of the mesh that no support reaches (docs/theory.md
        # §7.1): without them the stiffness matrix is singular.
        self._drop_floating_pieces()

        # Material interpolation (docs/theory.md §4.8, §7.1).
        self._stiffness_density_field = None
        self._tet_rho_m = None
        if self.material_interpolation and stiffness_density is not None:
            if np.shape(stiffness_density) != np.shape(field):
                raise ValueError(f"stiffness_density has shape {np.shape(stiffness_density)}, "
                                 f"the phase field {np.shape(field)}.")
            self._stiffness_density_field = np.asarray(stiffness_density)
            self._tet_rho_m = self._tet_stiffness_density(field, level_set_value, stiffness_density)
        self.material_interpolation_applied = self._tet_rho_m is not None

        # Run the finite element analysis
        self._fea()

        voxel_sed = self._strain_energy_densities(field)

        self._print_out(f"Total analysis time: {timer() - t0:.2f} seconds", "Minimal")

        return voxel_sed, self.s_mesh["volume"]

    def run_buckling(self, num_modes: int = 6, aggregation_p: float = 8.0,
                      prev_modes: np.ndarray | None = None,
                      min_energy_fraction: float = 0.5,
                      mode_continuity_warning_threshold: float = 0.7,
                      min_density_for_geometric_stiffness: float = 0.1,
                      eigensolver_tol: float = 1e-12) -> BucklingResults:
        """
        Runs a linear buckling eigenvalue analysis using the primal solution
        and mesh already computed by the most recent `run_analysis` call.

        Must be called AFTER `run_analysis` in the same iteration: it reuses
        `self.basis`, `self.u`, `self.K` and `self.fixed_dofs` set there.

        Signature matches `BucklingAnalysisInterface` so the optimiser can
        call any backend identically. Two arguments are accepted but have
        no effect here, because the conformal mesh covers only the solid
        region and contains no ersatz/void material for spurious modes to
        form in: `min_energy_fraction` (every mode trivially passes the
        solid-energy filter) and `min_density_for_geometric_stiffness`
        (there is no density field to ramp).

        NOTE ON `prev_modes`: the conformal mesh is rebuilt from scratch
        every iteration, so consecutive mode sets live on different DOF
        vectors and the MAC is undefined across the remesh.
        `_compute_mac` detects the size change and returns None rather than
        raising, so `mac_matrix` will normally be None for this backend.
        Mode-continuity tracking is only meaningful for the fixed-mesh
        backends (Regular, ConjugateGradient).

        Args:
        - num_modes (int): number of buckling modes to retain (m in the plan).
        - aggregation_p (float): p-norm aggregation exponent.
        - prev_modes (np.ndarray | None): previous call's mode_shapes.
        - min_energy_fraction (float): accepted, no effect (see above).
        - mode_continuity_warning_threshold (float): MAC below which a mode-switch
          warning is logged.
        - min_density_for_geometric_stiffness (float): accepted, no effect.
        - eigensolver_tol (float): ARPACK's relative eigenvalue tolerance
          (default 1e-12; 0.0 = machine precision), see
          `_solve_generalized_buckling_eigenproblem`.

        Returns:
        - BucklingResults
        """
        if not hasattr(self, 'basis') or not hasattr(self, 'u'):
            raise RuntimeError("run_buckling() called before run_analysis(); no primal solution available.")

        return self._run_buckling_fe(
            voxel_field_shape=self._buckling_voxel_shape,
            num_modes=num_modes,
            aggregation_p=aggregation_p,
            prev_modes=prev_modes,
            min_energy_fraction=min_energy_fraction,
            mode_continuity_warning_threshold=mode_continuity_warning_threshold,
            min_density_for_geometric_stiffness=min_density_for_geometric_stiffness,
            eigensolver_tol=eigensolver_tol,
        )


class Regular(Analysis):

    def __init__(self, iter_name: str, printouts: str="Normal", data_dir: str="analysis_data", youngs_modulus: float=210.0E9, poisson_ratio: float=0.3, save_meshes: bool=False, indent="  ", visualisation_scale: float=1.0, voxel_size_mm: float=1.0, quadrature_intorder: int | None = None, csg_domain_data: dict | None = None):
        """
        Initialises the Regular grid analysis class with specified parameters.
        This class performs FEA directly on the active voxels of the phi grid.

        Args:
        - iter_name (str): Name of the current iteration for data storage.
        - youngs_modulus (float): Young's modulus of the material.
        - poisson_ratio (float): Poisson's ratio of the material.
        - save_meshes (bool): Whether to save the generated meshes to files.
        - indent (str): Indentation string for print formatting.
        - printouts (str): Level of printouts ("Minimal", "Normal", "Verbose").
        - data_dir (str): Directory to store analysis data.
        - visualisation_scale (float): The scaling factor of deformation in the saved deformed vtk mesh file.
        - voxel_size_mm (float): The side length of a single voxel in millimeters.
        - quadrature_intorder (int | None): Polynomial degree the element
          quadrature rule must integrate exactly. None uses scikit-fem's
          default, which for a trilinear hex is a 4x4x4 (64-point) Gauss
          rule. Setting 3 selects 2x2x2 (8-point), which is what the
          ConjugateGradient backend's CUDA kernels use.

          This matters ONLY for buckling. For the elastic stiffness a 2x2x2
          rule is already exact on a trilinear hex, so compliance results are
          identical either way. The GEOMETRIC stiffness integrand is of
          higher polynomial degree, so the two rules give genuinely different
          (both valid, both convergent) discrete K_sigma operators, and hence
          buckling eigenvalues that differ by a fraction of a percent on a
          coarse mesh. Pass 3 if you want Regular and ConjugateGradient
          buckling numbers to agree closely; leave it None for the more
          accurate integration.
        """
        super().__init__(iter_name, printouts, data_dir, indent=indent)

        self.youngs_modulus = youngs_modulus
        self.poisson_ratio = poisson_ratio
        self.save_meshes = save_meshes
        self.visualisation_scale = visualisation_scale
        self.voxel_size_m = voxel_size_mm / 1000.0
        self.quadrature_intorder = quadrature_intorder

        if csg_domain_data:
            self.csg_data = csg_domain_data
        else:
            raise Exception("Required CSG domain data missing.")

        param_values = (
            f"{self.__class__.__name__} analysis component initialised with:\n"
            f"  Young's Modulus: {self.youngs_modulus:<.2e} Pa\n"
            f"  Poisson's Ratio: {self.poisson_ratio}\n"
            f"  Voxel Size: {voxel_size_mm} mm\n"
            f"  Quadrature intorder: {self.quadrature_intorder if self.quadrature_intorder is not None else 'scikit-fem default (4x4x4)'}\n"
            f"  CSG Geometry Enabled: {self.csg_data is not None}\n"
        )

        self.field_type = "density"

        self._print_out(param_values, "Normal")

    def _build_grid_index_maps(self):
        """
        Builds the mapping between the (i, j, k) voxel grid and scikit-fem's
        internal node and element numbering, DERIVED FROM MESH COORDINATES
        rather than assumed from a flatten order.

        This exists because scikit-fem's `MeshHex.init_tensor` does not number
        its nodes or elements in either C or Fortran order. As of scikit-fem
        12.0.2 the ordering runs y fastest, then x, then z -- equivalently
        `np.transpose(A, (2, 0, 1)).ravel()` -- which coincides with a
        Fortran-order flatten only for fields that are constant in x and y.

        An earlier revision of this class assumed `order='F'` throughout, which
        silently permuted the density field, the applied loads and the strain
        energy mapped back to the voxel grid whenever the field varied in x or
        y. A uniform test case cannot detect this: it was found by comparing
        the assembled load vector against the ConjugateGradient backend
        node-by-node in physical space.

        Deriving the maps from `self.m.p` rather than hard-coding the
        permutation means this stays correct if scikit-fem ever renumbers.

        Sets:
        - self._node_idx (nx+1, ny+1, nz+1) int array: skfem node index of
          each grid node.
        - self._elem_idx (nx, ny, nz) int array: skfem element index of each
          voxel.

        Usage: `flat[idx] = field_3d` converts a grid field into skfem order,
        and `field_3d = flat[idx]` converts back.
        """
        h = self.voxel_size_m
        nx_e, ny_e, nz_e = self.rho_shape

        # --- Nodes ---
        p = self.m.p                                   # (3, Nnodes)
        ni = np.rint(p[0] / h).astype(np.int64)
        nj = np.rint(p[1] / h).astype(np.int64)
        nk = np.rint(p[2] / h).astype(np.int64)
        self._node_idx = np.empty((nx_e + 1, ny_e + 1, nz_e + 1), dtype=np.int64)
        self._node_idx[ni, nj, nk] = np.arange(p.shape[1], dtype=np.int64)

        # --- Elements (located by centroid) ---
        cent = self.m.p[:, self.m.t].mean(axis=1)      # (3, Ne)
        ei = np.rint(cent[0] / h - 0.5).astype(np.int64)
        ej = np.rint(cent[1] / h - 0.5).astype(np.int64)
        ek = np.rint(cent[2] / h - 0.5).astype(np.int64)
        self._elem_idx = np.empty((nx_e, ny_e, nz_e), dtype=np.int64)
        self._elem_idx[ei, ej, ek] = np.arange(cent.shape[1], dtype=np.int64)

        # Cheap bijectivity check: a permutation must hit every index once.
        if (np.unique(self._elem_idx).size != nx_e * ny_e * nz_e
                or np.unique(self._node_idx).size != (nx_e + 1) * (ny_e + 1) * (nz_e + 1)):
            raise RuntimeError(
                "Failed to build a one-to-one map between the voxel grid and the "
                "scikit-fem mesh numbering. This usually means the mesh spacing "
                "does not match voxel_size_mm."
            )

    def _grid_to_element_array(self, field_3d):
        """Converts an (nx, ny, nz) voxel field into skfem element order."""
        flat = np.empty(field_3d.size, dtype=field_3d.dtype)
        flat[self._elem_idx] = field_3d
        return flat

    def _element_array_to_grid(self, flat):
        """Converts a per-element array in skfem order back to the voxel grid."""
        return flat[self._elem_idx]

    def _node_array_to_grid(self, flat):
        """
        Converts a per-NODE array in skfem order back to the (nx+1, ny+1,
        nz+1) voxel-grid node lattice, using `self._node_idx` (built by
        `_build_grid_index_maps`) -- the exact node-level analogue of
        `_element_array_to_grid`. Used for the standardised iteration .vtr's
        Displacement_Magnitude / Buckling_Mode_i_Magnitude PointData
        (docs/theory.md §7.6): unlike `Conformal`, this mesh's nodes ARE the
        voxel-grid nodes, so this is an exact lookup, not an interpolation.
        """
        return flat[self._node_idx]

    def _volume_mesh(self, rho):
        """
        Generates a regular hexahedral mesh from the rho field, creating spatially-varying material properties from the density field instead of removing elements.
        
        Args:
        - rho (np.ndarray): The density field.
        """
        self._print_out("Starting regular volume mesh generation...", "Normal")
        self.rho_shape = rho.shape
        dims = np.array(self.rho_shape)

        if self.save_meshes:
            try:
                self._print_out("Saving contour plots of rho field...", "Verbose")
                # Get center indices
                shape = rho.shape
                center_x = shape[0] // 2
                center_y = shape[1] // 2
                center_z = shape[2] // 2

                # Get the three slices
                # Note: numpy indexing is [x, y, z] for skimage.marching_cubes
                slice_yz = rho[center_x, :, :] # Slice at X = center
                slice_xz = rho[:, center_y, :] # Slice at Y = center
                slice_xy = rho[:, :, center_z] # Slice at Z = center

                # Create a figure with 3 subplots
                fig, axes = plt.subplots(1, 3, figsize=(18, 6))
                fig.suptitle(f'{self.iter_name} - Rho Field Slices', fontsize=16)

                # --- Plot YZ Slice (at X=center) ---
                ax0 = axes[0]
                # Use .T to align plot axes with index names
                contour0 = ax0.contourf(slice_yz.T, levels=30, cmap='viridis') 
                ax0.contour(slice_yz.T, levels=[0.5], colors='red', linewidths=2)
                fig.colorbar(contour0, ax=ax0, label='Rho Value')
                ax0.set_title(f'YZ Plane (at X = {center_x})')
                ax0.set_xlabel('Y Index')
                ax0.set_ylabel('Z Index')
                ax0.axis('image') # Equal aspect ratio

                # --- Plot XZ Slice (at Y=center) ---
                ax1 = axes[1]
                contour1 = ax1.contourf(slice_xz.T, levels=30, cmap='viridis')
                ax1.contour(slice_xz.T, levels=[0.5], colors='red', linewidths=2)
                fig.colorbar(contour1, ax=ax1, label='Rho Value')
                ax1.set_title(f'XZ Plane (at Y = {center_y})')
                ax1.set_xlabel('X Index')
                ax1.set_ylabel('Z Index')
                ax1.axis('image')

                # --- Plot XY Slice (at Z=center) ---
                ax2 = axes[2]
                contour2 = ax2.contourf(slice_xy.T, levels=30, cmap='viridis')
                ax2.contour(slice_xy.T, levels=[0.5], colors='red', linewidths=2)
                fig.colorbar(contour2, ax=ax2, label='Rho Value')
                ax2.set_title(f'XY Plane (at Z = {center_z})')
                ax2.set_xlabel('X Index')
                ax2.set_ylabel('Y Index')
                ax2.axis('image')

                fig.tight_layout()
                plot_path = os.path.join(self.analysis_dir, f"{self.iter_name}_rho_contours.png")
                fig.savefig(plot_path, dpi=150)
                plt.close(fig)
                self._print_out(f"Contour plot saved to {plot_path}", "Verbose")
            
            except Exception as e:
                self._print_out(f"Warning: Failed to save contour plots. Error: {e}", "Minimal")
        
        # 1. Create the full tensor mesh
        self._print_out("Initialising full tensor mesh...", "Verbose")
        t0 = timer()
        try:
            # Create node vectors
            x = np.linspace(0, (dims[0]) * self.voxel_size_m, dims[0] + 1)
            y = np.linspace(0, (dims[1]) * self.voxel_size_m, dims[1] + 1)
            z = np.linspace(0, (dims[2]) * self.voxel_size_m, dims[2] + 1)
            
            # Create the full mesh
            m = MeshHex.init_tensor(x, y, z)
        except Exception as e:
            raise RuntimeError(f"Failed to create skfem MeshHex.init_tensor: {e}")
        
        self.m = m # Store the full mesh
        self.original_element_count = m.t.shape[1]
        self._print_out(f"Full mesh time: {timer() - t0:.2f}s", "Verbose")
        
        # 2. Define strong material properties (base stiffness)
        self.lam_strong, self.mu_strong = lame_parameters(
            self.youngs_modulus, self.poisson_ratio
        )
        self._print_out(f"Strong Lame (lam, mu) for rho=1.0: {self.lam_strong:.2e}, {self.mu_strong:.2e}", "Verbose")

        # 3. Create continuous material property maps
        # Element ordering is taken from the mesh geometry, NOT assumed to be
        # a Fortran-order flatten (see _build_grid_index_maps).
        self._build_grid_index_maps()
        self.rho_flat = self._grid_to_element_array(np.ascontiguousarray(rho))
        if self.rho_flat.size != self.original_element_count:
            raise RuntimeError(f"Element count ({self.original_element_count}) does not match rho size ({self.rho_flat.size}).")

        # Create Lame parameter maps for all elements
        # Properties are directly proportional to the density field 'rho'
        # (which ranges from void_density to 1.0).
        # This matches the formulation rho(phi) * E_ijkl in the paper.
        lambda_map = self.rho_flat * self.lam_strong
        mu_map = self.rho_flat * self.mu_strong
        
        self._print_out(f"Continuous lambda_map (min, mean, max): {lambda_map.min():.2e}, {lambda_map.mean():.2e}, {lambda_map.max():.2e}", "Verbose")
        self._print_out(f"Continuous mu_map (min, mean, max): {mu_map.min():.2e}, {mu_map.mean():.2e}, {mu_map.max():.2e}", "Verbose")

        # 4. Create skfem DiscreteFields for material properties
        # Use a constant-per-element (P0) field
        e_material = ElementHex0()
        basis_material = Basis(self.m, e_material)

        self.lambda_map_field = basis_material.zeros()
        self.mu_map_field = basis_material.zeros()
        
        # Assign the flat arrays directly to the field's DOFs
        # self.lambda_map_field is a (n_elements,) array
        self.lambda_map_field[:] = lambda_map
        self.mu_map_field[:] = mu_map

        # 5. Calculate volume (based on binary threshold, for reporting)
        # This metric is kept for compatibility, but the 'true'
        # material volume is np.mean(rho_flat) * total_domain_volume
        self.active_element_indices = np.where(self.rho_flat >= 0.5)[0]
        self.num_active = self.active_element_indices.size
        
        voxel_volume = self.voxel_size_m ** 3
        volume = self.num_active * voxel_volume
        self.mesh_volume = volume # Store for compatibility
        self._print_out(f"Calculated 'active' (rho >= {0.5:.2f}) volume: {1e9 * volume:.6f} mm³ ({self.num_active} active voxels)", "Verbose")
        self._print_out(f"Total elements in mesh: {self.original_element_count}", "Verbose")
        
        self._print_out(f"Regular volume mesh and continuous material maps generated in {timer() - t0:.2f}s", "Normal")

    def _fea(self):
        """
        Sets up and runs the finite element analysis on the full,
        heterogeneous regular mesh.
        """
        self._print_out("Starting finite element analysis (FEA) on regular mesh...", "Normal")
        t0 = timer()

        m = self.m # Use the full mesh
        
        bbox = np.array([
            [np.min(m.p[0]), np.max(m.p[0])],
            [np.min(m.p[1]), np.max(m.p[1])],
            [np.min(m.p[2]), np.max(m.p[2])],
        ]) # Bounding box (x, y, z), shape (3, 2)
        mesh_info = (f"Loaded mesh with:\n"
                     f"  Nodes: {m.p.shape[1]}\n"
                     f"  Elements: {m.t.shape[1]} (Hex)\n"
                     f"  Bounding box X: [{bbox[0][0]:.4e}, {bbox[0][1]:.4e}]\n"
                     f"  Bounding box Y: [{bbox[1][0]:.4e}, {bbox[1][1]:.4e}]\n"
                     f"  Bounding box Z: [{bbox[2][0]:.4e}, {bbox[2][1]:.4e}]")
        self._print_out(mesh_info, "Verbose")

        # 2) Define boundaries and loads using CSG data
        nx, ny, nz, _ = self.csg_data["fixed_nodes"].shape
        
        # Ensure mask accounts for all 3 DOFs (Tx, Ty, Tz)
        fixed_node_mask = np.zeros((nx+1, ny+1, nz+1, 3), dtype=bool)
        nodal_loads = np.zeros((nx+1, ny+1, nz+1, 3))
        
        elem_fixed = self.csg_data["fixed_nodes"] # Shape: (nx, ny, nz, 3)
        load_vecs = self.csg_data["load_vectors"]

        # Smear the element-wise properties onto the connected nodes
        for dx in (0, 1):
            for dy in (0, 1):
                for dz in (0, 1):
                    fixed_node_mask[dx:nx+dx, dy:ny+dy, dz:nz+dz, :] |= elem_fixed
                    nodal_loads[dx:nx+dx, dy:ny+dy, dz:nz+dz] += load_vecs / 8.0

        # Map the (i, j, k) node grid onto scikit-fem's node numbering using the
        # coordinate-derived index map rather than a Fortran-order flatten.
        any_fixed = fixed_node_mask.any(axis=-1)
        is_fixed_node = np.zeros(any_fixed.size, dtype=bool)
        is_fixed_node[self._node_idx] = any_fixed
        fixed_nodes_indices = np.where(is_fixed_node)[0]

        if fixed_nodes_indices.size == 0:
            raise RuntimeError("No fixed nodes found in CSG data.")

        # 3) Setup the FEA basis
        e = ElementVector(ElementHex1())
        if self.quadrature_intorder is None:
            self.basis = Basis(m, e)
        else:
            self.basis = Basis(m, e, intorder=self.quadrature_intorder)

        @BilinearForm 
        def stiffness(u, v, w):
            lambda_val = self.lambda_map_field[None, :, None]
            mu_val = self.mu_map_field[None, :, None]
            T = sym_grad(u) 
            V = sym_grad(v) 
            C_T = 2. * mu_val * T + lambda_val * eye(trace(T), T.shape[0])
            return ddot(C_T, V)

        # ASSEMBLE MATRICES (Fix applied here!)
        self._print_out("Assembling stiffness matrix (k)...", "Verbose")
        k_assemble_start = timer()
        k = stiffness.assemble(self.basis) 
        self._print_out(f"Stiffness matrix assembly time: {timer() - k_assemble_start:.2f}s", "Verbose")

        # Assemble Force Vector directly onto nodal DOFs (No integrals needed!)
        self._print_out("Assembling force vector (f) from CSG nodes...", "Verbose")

        # Guard to satisfy the type checker before doing DOF assignment
        if self.basis.nodal_dofs is None:
            raise RuntimeError("Basis nodal DOFs are missing.")

        f = np.zeros(self.basis.N)
        for i in range(3):
            component = np.zeros(nodal_loads[..., i].size)
            component[self._node_idx] = nodal_loads[..., i]
            f[self.basis.nodal_dofs[i]] = component

        # 5) Apply boundary conditions
        u = self.basis.zeros()
        
        fixed_dofs_list = []
        for i in range(3):
            is_fixed_node_for_dof = np.zeros(fixed_node_mask[..., i].size, dtype=bool)
            is_fixed_node_for_dof[self._node_idx] = fixed_node_mask[..., i]
            fixed_nodes_indices = np.where(is_fixed_node_for_dof)[0]
            fixed_dofs_list.extend(self.basis.nodal_dofs[i, fixed_nodes_indices])

        fixed_dofs = np.unique(fixed_dofs_list)

        if fixed_dofs.size == 0:
            raise RuntimeError("No fixed DOFs found in CSG data.")

        self._print_out(f"Applying Dirichlet BCs on {fixed_dofs.size} DOFs.", "Verbose")

        # 6) Solve the system
        self._print_out("Solving the linear system...", "Verbose")
        solve_start = timer()
        self.u = solve(*condense(k, f, x=u, D=fixed_dofs)) # type: ignore
        self._print_out(f"Linear system solve time: {timer() - solve_start:.2f}s", "Verbose")

        # Store the assembled stiffness matrix and the fixed-DOF set so that
        # `run_buckling` can reuse them without a full re-assembly.
        self.K = k
        self.fixed_dofs = fixed_dofs

        # The EXACT compliance, c = f . u, in joules -- see the equivalent line
        # in Conformal._fea for why this is published rather than reconstructed
        # from the SED field by the optimiser.
        self.compliance = float(np.dot(f, self.u))
        self._print_out(f"Compliance (f.u): {self.compliance:.6e} J", "Verbose")

        # A load applied to a constrained DOF does no work and is silently
        # absorbed by the reaction when `condense` removes that row. Say so,
        # rather than letting a mis-drawn CSG region quietly lose its load.
        loaded_fixed = int(np.count_nonzero(f[fixed_dofs] != 0.0))
        if loaded_fixed:
            self._print_out(
                f"Warning: {loaded_fixed} DOF(s) carry both a load and a Dirichlet "
                f"constraint. Those load components do no work and are discarded by "
                f"the constraint; check the CSG regions if that was not intended.",
                "Minimal")

        # Post-process: calculate displacement in mm
        xyz_u = 1e3 * np.array([self.u[self.basis.nodal_dofs[i]] for i in range(3)]) # type: ignore
        self._print_out(f"Maximum displacement (mm) X: {np.max(xyz_u[0]):.6f}, Y: {np.max(xyz_u[1]):.6f}, Z: {np.max(xyz_u[2]):.6f}", "Verbose")

        # Standardised iteration .vtr (docs/theory.md §7.5/§7.6): displacement
        # magnitude, mapped from skfem node order onto the exact voxel-grid
        # node lattice via `_node_array_to_grid`. Replaces the old separate
        # `_deformed_mesh_regular.vtk` (baked-in displaced coordinates) --
        # this backend's mesh never moves or is rebuilt, so a scalar
        # displacement field over the fixed grid carries the same
        # information (docs/theory.md §7.6 on why a RectilinearGrid stores
        # deformation as data rather than as moved geometry).
        if self.save_meshes:
            disp_xyz_m = np.array([self.u[self.basis.nodal_dofs[i]] for i in range(3)])  # type: ignore
            node_disp_mag = np.linalg.norm(disp_xyz_m, axis=0)
            self._pending_point_data["Displacement_Magnitude"] = self._node_array_to_grid(node_disp_mag).astype(np.float32)

        self._print_out(f"FEA completed in {timer() - t0:.2f} seconds. It found a maximum deflection of { np.max(np.linalg.norm(xyz_u, axis=0)):.6f} mm.", "Normal")

    def _compute_stress_field(self, displacement_field: np.ndarray):
        """
        Computes the full quadrature-point-resolved stress and strain
        tensor fields (each shape (3,3,Ne,Nqp)) for any full-length
        displacement-like field on `self.basis`, using the CURRENT
        (rho-scaled) constitutive law. Factored out of
        `_strain_energy_densities` so it can be reused by the buckling
        sensitivity machinery.
        """
        strain_field = sym_grad(self.basis.interpolate(displacement_field)) # type: ignore
        stress_field = self._stress_from_strain(strain_field)
        return stress_field, strain_field

    def _stress_from_strain(self, strain_field: np.ndarray):
        """
        Actual (rho-scaled) constitutive law: sigma = C(rho_e) : epsilon.
        """
        lambda_qp = self.lambda_map_field[None, :, None]
        mu_qp = self.mu_map_field[None, :, None]
        return 2. * mu_qp * strain_field + lambda_qp * eye(trace(strain_field), strain_field.shape[0])

    def _stress_from_strain_prime(self, strain_field: np.ndarray):
        """
        d(constitutive law)/d(rho_e), evaluated at strain_field.

        Because the PCLS density interpolation is LINEAR in rho
        (C_e = rho_e * C_strong; see `_rho_prime` in the optimiser, which
        is a constant), this derivative is simply the constant "strong"
        material constitutive law -- it does not itself depend on rho_e.
        """
        return 2. * self.mu_strong * strain_field + self.lam_strong * eye(trace(strain_field), strain_field.shape[0])

    def _num_elements(self):
        return self.original_element_count

    def _modal_energy_fraction_in_solid(self, mode_field: np.ndarray, solid_weight=None):
        """
        Weighted fraction (0-1) of a mode's modal strain energy carried by
        elements that contribute geometric stiffness:

            sum_e w_e E_e / sum_e E_e,   E_e = phi^T K_e phi,

        with w_e the geometric-stiffness ramp r(rho_e) passed in by
        `_run_buckling_fe` (docs/theory.md §6.5(a)). Used to discard
        spurious modes localised in the low-stiffness ersatz/void material.
        With `solid_weight=None` the legacy hard indicator w_e = [rho_e >=
        0.5] is used.
        """
        strain_phi = sym_grad(self.basis.interpolate(mode_field)) # type: ignore
        stress_phi = self._stress_from_strain(strain_phi)
        element_energy = self._ddot_field_mean(stress_phi, strain_phi)
        element_energy = np.clip(element_energy, 0.0, None)
        total = element_energy.sum()
        if total <= 0.0:
            return 0.0
        if solid_weight is None:
            solid_weight = (self.rho_flat >= 0.5).astype(np.float64)
        return float(np.dot(element_energy, solid_weight) / total)

    def _geometric_stiffness_ramp(self, threshold: float = 0.1):
        """
        Smooth (cubic smoothstep) ramp that zeroes the geometric stiffness
        contribution of elements below `threshold` density and fully
        restores it by 2*threshold, guarding against spurious buckling
        modes forming in the low-stiffness ersatz/void material. See
        implementation plan section 4.1.

        Returns a (1,1,Ne,1)-broadcastable array.
        """
        upper = 2.0 * threshold
        x = np.clip((self.rho_flat - threshold) / max(upper - threshold, 1e-12), 0.0, 1.0)
        ramp = x * x * (3.0 - 2.0 * x)  # smoothstep
        return ramp[None, None, :, None]

    def _geometric_stiffness_ramp_prime(self, threshold: float = 0.1):
        """
        d(ramp)/d(rho_e) for the cubic smoothstep above, evaluated in
        closed form: with x = clip((rho_e-threshold)/(upper-threshold), 0, 1)
        and ramp(x) = x^2(3-2x), the chain rule gives

            d(ramp)/d(rho_e) = (6x - 6x^2) / (upper - threshold)   for 0 < x < 1
                              = 0                                   for x <= 0 or x >= 1.

        Because 6x(1-x) is itself exactly 0 at x=0 and x=1, using the
        SAME clipped `x` as `_geometric_stiffness_ramp` (rather than a
        separate unclipped derivative) already gives the correct value
        everywhere, including at and beyond the threshold/2*threshold
        boundary -- no piecewise branch is needed. See docs/theory.md
        §6.5(b) for where this term feeds into the buckling sensitivity.

        Returns an (Ne,)-shaped array, matching the shape `_ddot_field_mean`
        returns (unlike `_geometric_stiffness_ramp`, which returns a
        (1,1,Ne,1)-broadcastable array for multiplying tensor fields
        directly).
        """
        upper = 2.0 * threshold
        denom = max(upper - threshold, 1e-12)
        x = np.clip((self.rho_flat - threshold) / denom, 0.0, 1.0)
        return (6.0 * x * (1.0 - x)) / denom

    def _map_sensitivity_to_voxels(self, element_values: np.ndarray):
        """
        Maps a per-element (P0) sensitivity back onto the phi/rho voxel grid
        using the coordinate-derived element index map, which is the inverse
        of the map used to place the density field onto the mesh in the first
        place (see `_build_grid_index_maps`).
        """
        return self._element_array_to_grid(element_values)

    def _strain_energy_densities(self):
        """
        Calculates the strain energy densities in all elements and
        maps them back to the original full phi grid.

        Returns:
        - np.ndarray: Strain energy densities for each voxel in the original phi grid.
        """
        self._print_out("Calculating and mapping strain energy densities...", "Normal")
        
        stress_field, strain_field = self._compute_stress_field(self.u)

        # sed = 0.5 * (sigma : epsilon)
        # Sum over i and j, leaving (n_elements, n_integration_points)
        sed_dot_product_qp = np.einsum('ijkl,ijkl->kl', stress_field, strain_field)
        sed_qp = 0.5 * sed_dot_product_qp

        # Volume average over the element: sum_qp sed * dx / sum_qp dx, dx =
        # |J| * quadrature weight (skfem's basis.dx), so that F_e V_e is
        # exactly the element strain energy 1/2 u_e^T K_e u_e that the
        # compliance sensitivities (-2 V_e F0_e) and the energy identity
        # c = f.u = 2 sum_e F_e V_e assume (docs/theory.md §4.5, §7.2). A
        # plain np.mean over the quadrature points equals this only for
        # equal weights (the 2x2x2 rule, intorder 3); with scikit-fem's
        # default 4x4x4 Gauss rule for hex1 the weights differ and the plain
        # mean overstated the energy by 3 % on a random design and by up to
        # 28 % in a one-voxel bending member (optimiser review, 2026-10).
        dx = np.asarray(self.basis.dx, dtype=np.float64)
        element_sed = np.sum(sed_qp * dx, axis=1) / np.sum(dx, axis=1)  # Shape (n_elements,)

        # 'element_sed' is the flat, skfem-ordered array over the full grid.
        # Map it back to (x, y, z) using the same coordinate-derived index map
        # that placed the density onto the mesh, so the round trip is exact.
        voxel_sed = self._element_array_to_grid(element_sed)

        self._print_out(f"Min/Max Strain Energy Density: {np.min(voxel_sed):.4e}/{np.max(voxel_sed):.4e}", "Verbose")

        # --- Standardised iteration .vtr (docs/theory.md §7.5/§7.6) ---
        #
        # Replaces the old `_sed_mesh_regular_active.vtk` (SED restricted to
        # active elements only, in its own file): SED here covers the FULL
        # domain, active and void alike, matching how ConjugateGradient's
        # Strain_Energy_Density/Active_Voxels have always worked -- "active"
        # is now a separate mask field on the same grid rather than a
        # differently-scoped mesh.
        if self.save_meshes:
            self._grid_shape = self.rho_shape
            self._grid_hx = self._grid_hy = self._grid_hz = self.voxel_size_m

            self._pending_cell_data["Strain_Energy_Density"] = voxel_sed.astype(np.float32)
            self._pending_cell_data["Density"] = self._element_array_to_grid(self.rho_flat).astype(np.float32)
            self._pending_cell_data["Active_Voxels"] = self._element_array_to_grid(
                (self.rho_flat >= 0.5).astype(np.float32)
            )

            # Von Mises stress: same (3,3,Ne,Nqp) stress tensor already
            # computed above for SED, averaged over quadrature points and
            # reduced to the Voigt (xx,yy,zz,yz,xz,xy) convention shared with
            # gpu_analysis.py, then mapped exactly like SED (no interpolation --
            # this mesh IS the voxel grid).
            element_stress_avg = np.mean(stress_field, axis=-1)  # (3, 3, Ne), skfem order
            sxx, syy, szz = element_stress_avg[0, 0], element_stress_avg[1, 1], element_stress_avg[2, 2]
            syz, sxz, sxy = element_stress_avg[1, 2], element_stress_avg[0, 2], element_stress_avg[0, 1]
            element_stress_voigt = np.stack([sxx, syy, szz, syz, sxz, sxy], axis=-1)  # (Ne, 6), skfem order
            element_von_mises = _von_mises_from_voigt(element_stress_voigt)
            self._pending_cell_data["Von_Mises_Stress"] = self._element_array_to_grid(element_von_mises).astype(np.float32)

        return voxel_sed

    def _map_mode_shapes_to_grid(self, mode_shapes: np.ndarray) -> dict:
        """
        See `Analysis._map_mode_shapes_to_grid`. `mode_shapes` is (Ndofs, k)
        on `self.basis`'s DOF numbering (the same one `self.u` uses), so each
        mode's per-node magnitude is computed exactly like `_fea`'s own
        displacement magnitude, then mapped with the exact `_node_array_to_grid`
        lookup (this mesh's nodes ARE the voxel-grid nodes).
        """
        fields = {}
        for m in range(mode_shapes.shape[1]):
            mode_xyz = np.array([mode_shapes[self.basis.nodal_dofs[i], m] for i in range(3)])
            node_mag = np.linalg.norm(mode_xyz, axis=0)
            fields[f"Buckling_Mode_{m + 1}_Magnitude"] = self._node_array_to_grid(node_mag).astype(np.float32)
        return fields

    def run_analysis(self, field, level_set_value=None):
        """
        Runs numerical analysis on the provided field.

        The steps involved are:
        1) Generate the full regular hex mesh and material property maps.
        2) Set up and run the finite element analysis.
        3) Calculate strain energy densities and map them back to the grid.
        
        Args:
        - field (np.ndarray): The density field.
        - level_set_value (float): Placeholder to make 'interface' functionality work.

        Returns:
        - voxel_sed (np.ndarray): A mapping of the strain energy density in every cell of the density field.
        - (float): The volume of the *active* elements.
        """
        self._print_out("Starting regular analysis run...", "Normal")
        t0 = timer()

        try:
            # Generate the full volume mesh and material maps
            self._volume_mesh(field)

            # No "zero active elements -> skip the FEA" short-circuit here.
            # The ersatz stiffness rho_e*C0 is positive everywhere
            # (rho_e >= void_density > 0), so the problem is well-posed for
            # ANY density field, and returning an all-zero SED (as this used
            # to whenever no voxel had reached rho >= 0.5) silently switched
            # the stiffness driver off for exactly the low-volume-fraction
            # designs that most need it. num_active is kept for reporting.

            # Run the finite element analysis
            self._fea()

            # Calculate strain energy densities
            voxel_sed = self._strain_energy_densities()
            
            self._print_out(f"Total analysis time: {timer() - t0:.2f} seconds", "Minimal")

            return voxel_sed, self.mesh_volume

        except Exception as e:
            raise RuntimeError(f"Regular analysis run failed: {e}")

    def run_buckling(self, num_modes: int = 6, aggregation_p: float = 8.0,
                      prev_modes: np.ndarray | None = None,
                      min_energy_fraction: float = 0.5,
                      mode_continuity_warning_threshold: float = 0.7,
                      min_density_for_geometric_stiffness: float = 0.1,
                      eigensolver_tol: float = 1e-12) -> BucklingResults:
        """
        Runs a linear buckling eigenvalue analysis using the primal solution
        and mesh already computed by the most recent `run_analysis` call.

        Must be called AFTER `run_analysis` in the same iteration: it reuses
        `self.basis`, `self.u`, `self.K`, `self.fixed_dofs`, `self.rho_flat`
        and `self.lambda_map_field`/`self.mu_map_field` set there.

        Args:
        - num_modes (int): number of buckling modes to retain (m in the plan).
        - aggregation_p (float): p-norm aggregation exponent.
        - prev_modes (np.ndarray | None): previous call's mode_shapes, for
          MAC-based mode-continuity tracking. Pass None on the first call.
        - min_energy_fraction (float): minimum fraction of a mode's modal
          strain energy carried by elements that contribute geometric
          stiffness (weighted by the ramp r(rho_e)) for the mode to be
          retained (spurious-void-mode filter, docs/theory.md §6.5(a)).
        - mode_continuity_warning_threshold (float): MAC below which a mode-switch
          warning is logged.
        - min_density_for_geometric_stiffness (float): rho threshold below
          which the geometric stiffness contribution is smoothly ramped to
          zero, guarding against spurious buckling modes forming in the
          low-stiffness ersatz material (section 4.1 of the plan).
        - eigensolver_tol (float): ARPACK's relative eigenvalue tolerance
          (default 1e-12; 0.0 = machine precision), see
          `_solve_generalized_buckling_eigenproblem`.

        Returns:
        - BucklingResults
        """
        if not hasattr(self, 'basis') or not hasattr(self, 'u'):
            raise RuntimeError("run_buckling() called before run_analysis(); no primal solution available.")
        if not np.any(self.rho_flat > min_density_for_geometric_stiffness):
            # The ramp zeroes K_sigma everywhere: no element can buckle.
            self._print_out("No element above the geometric-stiffness threshold: "
                            "no compressive buckling mode can exist.", "Minimal")
            return BucklingResults(
                mu_values=np.zeros(0), mode_shapes=np.zeros((self.basis.N, 0)),
                mu_tilde=0.0, sensitivity=np.zeros(self.rho_shape),
                mac_matrix=None, solver_iters=0, n_candidates=0,
            )

        return self._run_buckling_fe(
            voxel_field_shape=self.rho_shape,
            num_modes=num_modes,
            aggregation_p=aggregation_p,
            prev_modes=prev_modes,
            min_energy_fraction=min_energy_fraction,
            mode_continuity_warning_threshold=mode_continuity_warning_threshold,
            min_density_for_geometric_stiffness=min_density_for_geometric_stiffness,
            eigensolver_tol=eigensolver_tol,
        )


def _save_rectilinear_vtr(filename, grid_shape, hx, hy, hz,
                           cell_data: dict | None = None,
                           point_data: dict | None = None,
                           field_data: dict | None = None):
    """
    Save a rectilinear grid (.vtr) carrying an arbitrary number of named
    cell-, point- and field-data arrays (docs/theory.md §7.5/§7.6), written
    with VTK XML's binary "appended" encoding via `pyvista.RectilinearGrid`.
    The installed VTK's default writer settings zlib-compress that encoding
    automatically -- ParaView and PyVista both read it natively, no plugin
    required.

    Replaces an earlier hand-written-XML version of this function that only
    ever wrote exactly three fixed, ASCII-encoded arrays (SED/Density/
    Active_Voxels) for `ConjugateGradient`. This version is backend-agnostic
    and field-count-agnostic: every backend's `run_analysis` (and
    `Analysis.finalize_iteration_output`) builds its own `cell_data`/
    `point_data`/`field_data` dicts and passes them straight through.

    Args:
    - filename (str): output .vtr path.
    - grid_shape (tuple[int, int, int]): (nx, ny, nz) ELEMENT (voxel) counts.
    - hx, hy, hz (float): element edge length in metres along each axis.
    - cell_data (dict[str, np.ndarray] | None): each array shaped
      (nx, ny, nz) in standard numpy [i, j, k] indexing -- one value per
      voxel/element (e.g. Strain_Energy_Density, Density, Active_Voxels,
      Von_Mises_Stress).
    - point_data (dict[str, np.ndarray] | None): each array shaped
      (nx+1, ny+1, nz+1) -- one value per grid node (e.g.
      Displacement_Magnitude, Buckling_Mode_i_Magnitude).
    - field_data (dict[str, scalar | array] | None): global, non-geometric
      metadata (compliance, lambda1, outer_iter, ...) attached to the file
      as a whole rather than to the grid.
    """
    nx, ny, nz = grid_shape
    x = np.linspace(0, nx * hx, nx + 1)
    y = np.linspace(0, ny * hy, ny + 1)
    z = np.linspace(0, nz * hz, nz + 1)

    grid = pv.RectilinearGrid(x, y, z)

    # VTK expects the data stream to vary X fastest (Fortran order) for a
    # structured/rectilinear grid -- true of both cell and point arrays.
    if cell_data:
        for name, arr in cell_data.items():
            arr = np.asarray(arr, dtype=np.float32)
            if arr.shape != (nx, ny, nz):
                raise ValueError(
                    f"cell_data['{name}'] has shape {arr.shape}, expected "
                    f"{(nx, ny, nz)} (the element/voxel grid)."
                )
            grid.cell_data[name] = arr.flatten(order='F')

    if point_data:
        for name, arr in point_data.items():
            arr = np.asarray(arr, dtype=np.float32)
            if arr.shape != (nx + 1, ny + 1, nz + 1):
                raise ValueError(
                    f"point_data['{name}'] has shape {arr.shape}, expected "
                    f"{(nx + 1, ny + 1, nz + 1)} (the node grid)."
                )
            grid.point_data[name] = arr.flatten(order='F')

    if field_data:
        for name, val in field_data.items():
            # Strings (e.g. the final save's termination_reason) are stored
            # as VTK string arrays; everything else as float64.
            if isinstance(val, str):
                grid.field_data[name] = np.array([val])
            else:
                grid.field_data[name] = np.atleast_1d(np.asarray(val, dtype=np.float64))

    grid.save(filename, binary=True)


def _von_mises_from_voigt(stress_voigt: np.ndarray) -> np.ndarray:
    """
    Von Mises equivalent stress from an (..., 6)-shaped array of Voigt-form
    Cauchy stress components, ordered (xx, yy, zz, yz, xz, xy) -- the same
    convention used throughout this module and `gpu_analysis.py` (see the "Voigt
    ordering" note above `gpu_analysis.VOIGT_C_LAMBDA_BASIS`).

    sigma_vm = sqrt(0.5*((sxx-syy)^2+(syy-szz)^2+(szz-sxx)^2)
                    + 3*(syz^2+sxz^2+sxy^2))
    """
    s = np.asarray(stress_voigt)
    sxx, syy, szz, syz, sxz, sxy = (s[..., i] for i in range(6))
    return np.sqrt(
        0.5 * ((sxx - syy) ** 2 + (syy - szz) ** 2 + (szz - sxx) ** 2)
        + 3.0 * (syz ** 2 + sxz ** 2 + sxy ** 2)
    )


# =========================================================================
# --- ConjugateGradient buckling: design notes ---
#
# The primal solve in ConjugateGradient stays fully matrix-free and
# GPU-resident, exactly as before. For the buckling side-computation,
# however, this module deliberately assembles EXPLICIT sparse K and
# K_sigma matrices on the host with vectorised NumPy, then reuses the
# same PyPardiso + ARPACK machinery as the Regular and Conformal classes.
#
# This is a considered trade-off, not an oversight:
#   - The buckling eigenproblem and its adjoint involve far fewer
#     floating-point operations than the primal PCG solve (a handful of
#     factorisation-backed solves vs. thousands of matrix-free matvecs),
#     so host assembly is unlikely to dominate run time.
#   - An eigensolver in float32 is not viable; the host path is float64
#     throughout, which is what makes the eigenvalues trustworthy.
#   - It let the whole sensitivity be validated by finite differences
#     without a GPU (see the module docstring).
#
# The element formulation itself is NOT duplicated here. The host-side
# shape-function data, Lame conversion, Voigt constitutive bases and
# element->DOF map are all imported from gpu_analysis.py, which holds the single
# canonical definition next to the kernels they must agree with. Run
# `gpu_analysis.verify_host_kernel_consistency()` on a GPU machine to confirm
# the host and kernel paths still agree after any edit to either.
#
# If profiling later shows the host assembly IS a bottleneck, the next
# step is the dedicated `_element_geometric_action_kernel` described in
# Step 5 of the implementation plan; `compute_element_stress` in
# gpu_analysis.py already provides the GPU-side stress it would need.
# =========================================================================


def _engines_force_classic():
    """gpu_analysis.FORCE_CLASSIC (the regression guard pins every engine option to "classic")."""
    import gpu_analysis
    return gpu_analysis.FORCE_CLASSIC


class ConjugateGradient(Analysis):
    
    def __init__(self, iter_name: str, printouts: str="Normal", data_dir: str="analysis_data", youngs_modulus: float=210.0E9, poisson_ratio: float=0.3, voxel_size_mm: float=1.0, pcg_max_iters: int=10_000, pcg_tol: float=1e-3, pcg_rel_tol: float=1e-6, save_meshes: bool=False, indent: str="  ", csg_domain_data: dict | None = None, warm_start_cg: bool=False, warm_start_max_rho_change: float=0.15, warm_start_probe_iters: int=8, warm_start_min_reduction: float=0.5, warm_start_state: dict | None = None, element_type: str="q1", preconditioner: str="gmg", gmg_levels: int | None=None, gmg_smoother_steps: int=2, gmg_coarse_max_nodes: int=400, gmg_reuse_max_rel_change: float=0.0, analysis_refinement: int=1, engine: str="auto", pcg_engine: str="auto", pcg_check_every: int | None=None):

        super().__init__(iter_name, printouts, data_dir, indent=indent)

        """
        Initialises the Conjugate Gradient analysis class with specified parameters.

        Args:
        - iter_name (str): Name of the current iteration for data storage.
        - printouts (str): Level of printouts ("Silent", "Minimal", "Normal", "Verbose").
        - data_dir (str): Directory to store analysis data.
        - youngs_modulus (float): Young's modulus of the material in Pa.
        - poisson_ratio (float): Poisson's ratio of the material.
        - voxel_size_mm (float): Edge length of each voxel in millimetres.
        - pcg_max_iters (int): Maximum number of iterations for the PCG solver.
        - pcg_tol (float): ABSOLUTE convergence tolerance for the PCG solver,
          compared against the residual 2-norm in newtons.
        - pcg_rel_tol (float): RELATIVE convergence tolerance. The solver stops
          at ||r|| <= max(pcg_tol, pcg_rel_tol * ||b||), so this is the floor
          on solve accuracy that survives a change of load scale. With an
          absolute tolerance alone the effective accuracy tracked the load
          magnitude: 1e-3 N is a relative 1.7e-6 against a 600 N load but a
          relative 1e-3 against a 1 N one, so the same problem file gave
          wildly different sensitivity quality depending on how the loads
          happened to be scaled. Set to 0.0 to restore purely absolute
          behaviour.
        - save_meshes (bool): Whether to save analysis meshes and outputs to files.
        - indent (str): Indentation string used for print formatting.
        - warm_start_cg (bool): Seed each PCG solve from the previous inner
          iteration's converged displacement field instead of zero (docs/
          theory.md, "CG warm start"). Empirically cuts iterations 18-52%
          while the design is coarse, but the benefit shrinks and can
          reverse once the void/solid stiffness contrast has grown late in
          a run; see warm_start_max_rho_change/warm_start_probe_iters/
          warm_start_min_reduction below for the two safeguards that bound
          the downside. Requires warm_start_state (below) to actually do
          anything -- pcls_optimiser.py supplies it automatically at the
          real inner-loop call site only, so probe and final/exported
          solves are never warm-started even when this is True.
        - warm_start_max_rho_change (float): cheap pre-solve gate. Skip the
          warm start (solve cold) when the penalised stiffness density
          field has changed, at any single voxel, by more than this since
          the field the stored displacement solved (L-infinity, not L2,
          because a single voxel flipping void<->solid changes the LOCAL
          stiffness contrast sharply even when it barely moves an
          aggregate norm). In [0, 1] since that is rho_model's own range.
        - warm_start_probe_iters (int): the real safeguard. After this many
          PCG iterations of a warm-started solve, check whether it has
          actually paid off (see warm_start_min_reduction) and, if not,
          restart cold in place for the remaining budget. Bounds a bad
          warm start's worst case at this many wasted iterations, never a
          full duplicate solve.
        - warm_start_min_reduction (float): fraction ||r|| must have fallen,
          relative to the warm start's own initial residual ||b - A u0||,
          by warm_start_probe_iters, or the solve is judged not to be
          benefiting and is restarted cold. In [0, 1).
        - warm_start_state (dict | None): the mutable container the warm
          start reads its previous displacement/density from and writes its
          own into, shared across the many ConjugateGradient instances one
          run creates (a fresh one every inner iteration). Never set this
          from a problem file; it is Python-only plumbing supplied by
          pcls_optimiser.py.
        - element_type (str): "q1" (default since 2026-10-22; the trilinear hex as a
          precomputed 24x24 matrix), "q1_gauss" (the original 2x2x2 Gauss-point
          kernels, the pre-2026-10-22 default; select it to reproduce old runs) or "q1_im" (Wilson-Taylor incompatible-modes hex,
          QM6, condensed to a 24x24 matrix; removes shear locking of thin members in
          bending). Applies to the PCG action, the block-Jacobi diagonal, the element
          strain energy (sensitivities, compliance) and the stiffness K of the buckling
          analysis. The geometric stiffness G, the stress used for it and the von Mises
          output stay trilinear (from the compatible displacement field); for "q1_im"
          the buckling sensitivity therefore neglects the incompatible-mode correction
          of the stress (docs/theory.md §13). See gpu_analysis.py.
        - preconditioner (str): "gmg" (default since 2026-10-22), a
          geometric-multigrid V-cycle (Galerkin coarse operators, Chebyshev smoother,
          dense coarsest solve) that cuts the CG iteration count by one to two orders
          of magnitude (docs/theory.md §14). Works with every element_type and
          the CG warm start. On a grid too small to coarsen (fewer than 2 levels) it
          falls back to "block_jacobi" automatically (one printed line, and
          run_info.json "preconditioner_used"). "block_jacobi" (the pre-2026-10-22
          default) is selectable explicitly.
        - gmg_levels (int | None): total number of multigrid levels (>= 2); None =
          coarsen until at most gmg_coarse_max_nodes nodes remain.
        - gmg_smoother_steps (int): Chebyshev degree, pre and post (default 2).
        - gmg_coarse_max_nodes (int): coarsest-level node limit for gmg_levels=None
          (default 400; the coarsest level is inverted densely).
        - gmg_reuse_max_rel_change (float): 0 (default) rebuilds the multigrid hierarchy in
          every solve. > 0 reuses the previous solve's coarse operators and Chebyshev
          bounds (process-wide, see gpu_analysis) while max|E - E_prev| / max(E_prev) is at most
          this value and the grid, constraints and element are unchanged; a stale
          preconditioner is still SPD, so the CG solution is the same to tolerance, only
          the iteration count may rise. Holds ~0.3-2 GB of device memory between solves.
        - analysis_refinement (int): 1 (default, bit-identical), 2 or 3. Each design voxel
          is analysed as r^3 equal fine elements with the voxel's E (h/r), the voxel
          constraints and loads replicated (load / r^3 per child, so total force and
          moment are preserved exactly). The returned strain-energy density, von Mises
          stress and exports stay per design voxel (energy = sum of the children, von
          Mises = mean), so sensitivities are those of the fine model; displacements
          are exported on the design-voxel nodes. Removes the over-stiffness of
          under-resolved staircase members; needs preconditioner="gmg" in practice
          (r^3 more DOFs). Not available with buckling (run_buckling raises).
          docs/theory.md §15.
        - engine (str): elastic-analysis engine (gpu_analysis.py, docs/README.md ("Solver engines")). "classic": exactly the
          settings above. "auto" (default since 2026-10-23): chosen per process from the machine's capability
          -- "gpu_gmg" on a CUDA device, "gpu_gmg_lean" (GMG with hierarchy reuse 0.05) on cards with
          <= 4 GB, "cpu_direct" (host sparse Cholesky) without CUDA -- with analysis_refinement capped
          by the device-memory model. Explicit "gpu_gmg", "gpu_gmg_lean", "gpu_block_jacobi" or
          "cpu_direct" override preconditioner accordingly.
        - pcg_engine (str): "auto" (default since 2026-10-24: jax where capable and faster, else fast), "fast" (gpu_analysis, docs/theory.md §16) or "classic" (the previous
          solver): the fast engine is the same GMG-preconditioned CG with node-gather kernels (no atomics, bitwise reproducible run to run), fused
          smoother kernels, device-side CG scalars and a batched Galerkin setup. Needs preconditioner="gmg" (q1/q1_im
          element, coarsenable grid); otherwise the solve runs classic (recorded in the solve info).
          "jax": the same GMG-PCG written in JAX/XLA (gpu_analysis.py, docs/theory.md §17; needs JAX with CUDA, imported lazily and
          optional): one jit for the hierarchy setup and one for the whole PCG loop, cubic voxels, q1/q1_im; falls back to "fast" when it
          cannot run (the reason is in the solve info and gpu_analysis.RECORD). "auto": gpu_analysis.pcg_choice picks "jax" or "fast" per grid
          (jax only where measured faster and its memory fits).
        - pcg_check_every (int | None): iterations between host reads of the fast engine's stopping flag; None =
          adaptive. Does not change the result (the stop is exact), only the number of wasted iterations.
        """

        if element_type not in ELEMENT_TYPES:
            raise ValueError(f"element_type must be one of {ELEMENT_TYPES}, got {element_type!r}.")
        self.element_type = element_type
        if preconditioner not in PRECONDITIONERS:
            raise ValueError(f"preconditioner must be one of {PRECONDITIONERS}, got {preconditioner!r}.")
        if gmg_levels is not None and int(gmg_levels) < 2:
            raise ValueError("gmg_levels must be >= 2 (or None for automatic).")
        if int(gmg_smoother_steps) < 1:
            raise ValueError("gmg_smoother_steps must be >= 1.")
        if int(gmg_coarse_max_nodes) < 8:
            raise ValueError("gmg_coarse_max_nodes must be >= 8.")
        self.preconditioner = preconditioner
        if pcg_engine not in PCG_ENGINES:
            raise ValueError(f"pcg_engine must be one of {PCG_ENGINES}, got {pcg_engine!r}.")
        if pcg_check_every is not None and int(pcg_check_every) < 1:
            raise ValueError("pcg_check_every must be >= 1 (or None for adaptive).")
        self.pcg_engine = pcg_engine
        self.pcg_check_every = None if pcg_check_every is None else int(pcg_check_every)
        self.gmg_levels = None if gmg_levels is None else int(gmg_levels)
        self.gmg_smoother_steps = int(gmg_smoother_steps)
        self.gmg_coarse_max_nodes = int(gmg_coarse_max_nodes)
        self.gmg_info = None
        if float(gmg_reuse_max_rel_change) < 0.0:
            raise ValueError("gmg_reuse_max_rel_change must be >= 0.")
        self.gmg_reuse_max_rel_change = float(gmg_reuse_max_rel_change)
        if analysis_refinement not in (1, 2, 3):
            raise ValueError(f"analysis_refinement must be 1, 2 or 3, got {analysis_refinement!r}.")
        self.analysis_refinement = int(analysis_refinement)
        from gpu_analysis import ELASTIC_ENGINES
        if engine not in ELASTIC_ENGINES:
            raise ValueError(f"engine must be one of {ELASTIC_ENGINES}, got {engine!r}.")
        self.engine = engine
        self.youngs_modulus = youngs_modulus
        self.poisson_ratio = poisson_ratio
        self.voxel_size_m = voxel_size_mm / 1000.0
        self.pcg_max_iters = pcg_max_iters
        self.pcg_tol = pcg_tol
        self.pcg_rel_tol = max(0.0, float(pcg_rel_tol))
        self.save_meshes = save_meshes

        self.warm_start_cg = bool(warm_start_cg)
        self.warm_start_max_rho_change = float(warm_start_max_rho_change)
        if not (0.0 <= self.warm_start_max_rho_change <= 1.0):
            raise ValueError("warm_start_max_rho_change must be in [0, 1].")
        self.warm_start_probe_iters = int(warm_start_probe_iters)
        if self.warm_start_probe_iters < 1:
            raise ValueError("warm_start_probe_iters must be >= 1.")
        self.warm_start_min_reduction = float(warm_start_min_reduction)
        if not (0.0 <= self.warm_start_min_reduction < 1.0):
            raise ValueError("warm_start_min_reduction must be in [0, 1).")
        self.warm_start_state = warm_start_state

        if csg_domain_data:
            self.csg_data = csg_domain_data
        else:
            raise Exception("Required CSG domain data missing.")

        param_values = (
            f"{self.__class__.__name__} analysis component initialised with:\n"
            f"  Young's Modulus: {self.youngs_modulus:<.2e} Pa\n"
            f"  Poisson's Ratio: {self.poisson_ratio}\n"
            f"  Voxel Size: {voxel_size_mm} mm"
            f"  CSG Geometry Enabled: {self.csg_data is not None}"
            f"  CG warm start: {'enabled' if self.warm_start_cg else 'disabled'}"
            f"  Element type: {self.element_type}"
            f"  Preconditioner: {self.preconditioner}"
            f"  Analysis refinement: {self.analysis_refinement}"
        )

        self.field_type = "density"

        self._print_out(param_values, "Verbose")

    def _rep(self, a):
        """Replicate a per-design-voxel array (axes 0-2) onto the r^3 fine elements of each voxel."""
        r = self.analysis_refinement
        if r == 1:
            return a
        return np.repeat(np.repeat(np.repeat(a, r, axis=0), r, axis=1), r, axis=2)

    def _refined_csg(self):
        """(elem_fixed, load_vecs) per fine element: constraints replicated, load / r^3 per child."""
        r = self.analysis_refinement
        elem_fixed = self.csg_data["fixed_nodes"]
        load_vecs = self.csg_data["load_vectors"]
        if r == 1:
            return elem_fixed, load_vecs
        return self._rep(elem_fixed), self._rep(np.asarray(load_vecs)) / float(r ** 3)

    def _mean_children(self, flat_fine, dtype=np.float32):
        """Per-design-voxel mean (C-order flat, coarse) of a flat fine-element array (r = 1: unchanged)."""
        r = self.analysis_refinement
        if r == 1:
            return flat_fine
        nxr, nyr, nzr = self._fine_elem_shape
        a = np.asarray(flat_fine, dtype=np.float64).reshape(nxr // r, r, nyr // r, r, nzr // r, r)
        return a.mean(axis=(1, 3, 5)).reshape(-1).astype(dtype)

    def _volume_mesh(self, rho: np.ndarray):
        """
        Create geometry counters, per-element E and nu (from lambda/mu = rho * lambda_strong/mu_strong),
        nodal RHS vector b, and fixed_mask for Dirichlet nodes using CSG mapping.

        Returns:
            nx, ny, nz (node counts),
            hx, hy, hz (element sizes in m),
            E_elem (Ne,), nu_elem (Ne,),
            b (Ndofs,), fixed_mask (Nnodes,),
            rho_flat (Ne,)
        """
        if rho.ndim != 3:
            raise ValueError("rho must be a 3D array")

        rho = self._rep(rho)          # analysis_refinement r: r^3 fine elements per design voxel (r = 1: unchanged)
        self._fine_elem_shape = tuple(rho.shape)
        nx_elems, ny_elems, nz_elems = rho.shape
        nx_nodes = nx_elems + 1
        ny_nodes = ny_elems + 1
        nz_nodes = nz_elems + 1

        hx = hy = hz = np.float32(self.voxel_size_m / self.analysis_refinement)

        Ne = nx_elems * ny_elems * nz_elems
        Nnodes = nx_nodes * ny_nodes * nz_nodes
        Ndofs = 3 * Nnodes

        rho = np.asfortranarray(rho)
        # Flatten in C order to match the GPU node ID logic
        rho_flat = rho.flatten(order='C').astype(np.float32)
        if rho_flat.size != Ne:
            raise RuntimeError("Mismatch between flattened rho size and expected element count")

        # Lame for solid (strong)
        lam_strong, mu_strong = lame_parameters(self.youngs_modulus, self.poisson_ratio)

        # Per-element Lame scaled by rho
        lambda_map = rho_flat * lam_strong
        mu_map = rho_flat * mu_strong

        # Convert lambda, mu to E and nu per element safely (avoid divide-by-zero)
        denom = (lambda_map + mu_map)
        tiny = 1e-12
        denom_safe = np.where(np.abs(denom) < tiny, tiny, denom)

        # E = mu*(3*lambda + 2*mu)/(lambda + mu)
        E_elem = (mu_map * (3.0 * lambda_map + 2.0 * mu_map) / denom_safe).astype(np.float32)
        # nu = lambda / (2*(lambda + mu))
        nu_elem = (lambda_map / (2.0 * denom_safe)).astype(np.float32)

        # --- Build nodal arrays from CSG definitions ---
        # The CG solver uses C-order (z varies fastest)
        fixed_node_mask = np.zeros((nx_nodes, ny_nodes, nz_nodes, 3), dtype=bool)
        nodal_loads = np.zeros((nx_nodes, ny_nodes, nz_nodes, 3), dtype=np.float32)
        
        elem_fixed, load_vecs = self._refined_csg()   # Shape: (nx, ny, nz, 3); r > 1: replicated

        # Smear the element-wise properties onto the 8 connected corners
        for dx in (0, 1):
            for dy in (0, 1):
                for dz in (0, 1):
                    fixed_node_mask[dx:nx_elems+dx, dy:ny_elems+dy, dz:nz_elems+dz, :] |= elem_fixed
                    nodal_loads[dx:nx_elems+dx, dy:ny_elems+dy, dz:nz_elems+dz] += load_vecs / 8.0
        
        # Flatten the (nx, ny, nz, 3) array directly
        fixed_node_mask_flat = fixed_node_mask.flatten(order='C')
        fixed_mask = fixed_node_mask_flat.astype(np.int32)
        
        nodal_loads_flat = nodal_loads.reshape(-1, 3, order='C')
        
        # Build nodal force vector b (host)
        b = np.zeros(Ndofs, dtype=np.float32)

        # The CG solver structures DOF arrays as [ux0, uy0, uz0, ux1, uy1, uz1...]
        b[0::3] = nodal_loads_flat[:, 0]
        b[1::3] = nodal_loads_flat[:, 1]
        b[2::3] = nodal_loads_flat[:, 2]

        # Drop any load sitting on a constrained DOF.
        #
        # Physically this changes nothing: the DOF cannot move, so the load does
        # no work and is carried entirely by the reaction. It matters for the
        # SOLVER, because the matrix-free PCG imposes the Dirichlet conditions
        # by zeroing Ap at fixed DOFs. The residual r = b - A u then keeps its
        # initial value b there for ever, so ||r|| could never fall below
        # ||b_fixed|| and the solve ran to pcg_max_iters on every single call
        # -- silently, and at full cost -- whenever a CSG region put a load on a
        # voxel that was also constrained. Removing the component at source
        # makes the residual an honest measure of the free-DOF error again.
        loaded_and_fixed = int(np.count_nonzero((fixed_mask == 1) & (b != 0.0)))
        if loaded_and_fixed:
            self._print_out(
                f"Warning: {loaded_and_fixed} DOF(s) carry both a load and a Dirichlet "
                f"constraint. The load there does no work and has been removed from the "
                f"right-hand side; check the CSG regions if that was not intended.",
                "Minimal")
            b[fixed_mask == 1] = 0.0

        return (nx_nodes, ny_nodes, nz_nodes,
                hx, hy, hz,
                E_elem, nu_elem, b, fixed_mask,
                rho_flat)

    def run_analysis(self, field: np.ndarray, level_set_value=None):
        """
        Run the GPU matrix-free PCG using density field rho.
        Returns voxel_sed (same shape as rho) and mesh_volume (active volume m^3).
        """
        t0 = timer()
        if self.engine != "classic" and not _engines_force_classic() and self._apply_elastic_engine() == "cpu_direct":
            return self._run_analysis_cpu(field)
        self._print_out("Starting ConjugateGradient (matrix-free GPU) analysis...", "Verbose")

        (nx, ny, nz,
         hx, hy, hz,
         E_elem, nu_elem, b, fixed_mask,
         rho_flat) = self._volume_mesh(field)

        # active volume consistent with Regular (rho >= 0.5)
        r_ref = self.analysis_refinement
        if r_ref == 1:
            active_indices = np.where(rho_flat >= 0.5)[0]
            num_active = active_indices.size
        else:       # active design voxels (rho is constant over the children of a voxel)
            num_active = int(np.count_nonzero(self._mean_children(rho_flat) >= 0.5))
        voxel_vol = float(self.voxel_size_m ** 3)
        mesh_volume = num_active * voxel_vol

        # No "zero active elements -> return zero SED" short-circuit: the
        # ersatz stiffness is positive everywhere, so the solve is well-posed
        # for any density field (see the matching note in Regular.run_analysis).
        self._print_out(f"Active elements: {num_active}/{field.size}. Launching GPU PCG...", "Verbose")

        # --- CG warm start (docs/theory.md, "CG warm start") ---
        # `field` here IS the penalised stiffness density rho_model handed to
        # this backend (pcls_optimiser.py's _rho_model), so comparing it
        # against the field the stored displacement solved captures BOTH a
        # raw density change AND a penalisation-power-ramp change in one
        # signal -- exactly what determines the stiffness matrix.
        Ndofs = 3 * nx * ny * nz
        u0 = None
        self.warm_start_attempted = False
        self.warm_start_bailed_to_cold = False
        self.warm_start_skip_reason = None
        state = self.warm_start_state
        if self.warm_start_cg:
            if state is None:
                self.warm_start_skip_reason = "no warm_start_state supplied for this call"
            else:
                u_prev = state.get("u_prev")
                rho_prev = state.get("rho_prev")
                if u_prev is None or rho_prev is None:
                    self.warm_start_skip_reason = "no previous solve stored yet"
                elif u_prev.shape[0] != Ndofs or rho_prev.shape != field.shape:
                    self.warm_start_skip_reason = "mesh/domain size changed since the stored solve"
                else:
                    rho_change = float(np.max(np.abs(field.astype(np.float32) - rho_prev)))
                    if rho_change > self.warm_start_max_rho_change:
                        self.warm_start_skip_reason = (
                            f"stiffness density changed by {rho_change:.4f} at its most-changed "
                            f"voxel (> warm_start_max_rho_change={self.warm_start_max_rho_change})")
                    else:
                        u0 = u_prev
                        self.warm_start_attempted = True
            if self.warm_start_skip_reason:
                self._print_out(f"CG warm start: cold-starting this solve ({self.warm_start_skip_reason}).", "Verbose")

        # Call the high-level GPU PCG solver from gpu_analysis.py
        try:
            u_h, history, cg_info = matrix_free_pcg(
                nx, ny, nz, np.float32(hx), np.float32(hy), np.float32(hz),
                E_elem, nu_elem, b, fixed_mask,
                max_iters=self.pcg_max_iters,
                tol=self.pcg_tol,
                rel_tol=self.pcg_rel_tol,
                verbose=False,
                u0=u0,
                warm_start_probe_iters=self.warm_start_probe_iters,
                warm_start_min_reduction=self.warm_start_min_reduction,
                element_type=self.element_type, nu=self.poisson_ratio,
                preconditioner=self.preconditioner, gmg_levels=self.gmg_levels,
                gmg_smoother_steps=self.gmg_smoother_steps,
                gmg_coarse_max_nodes=self.gmg_coarse_max_nodes,
                gmg_reuse_max_rel_change=self.gmg_reuse_max_rel_change,
                **({} if self.pcg_engine == "classic" else dict(pcg_engine=self.pcg_engine, pcg_check_every=self.pcg_check_every)),
            )
        except Exception as e:
            raise RuntimeError(f"matrix_free_pcg_optimized failed: {e}")
        self.gmg_info = cg_info.get("gmg")
        self.pcg_engine_used = cg_info.get("pcg_engine_used")

        self.warm_start_bailed_to_cold = bool(cg_info.get("bailed_to_cold", False))
        if self.warm_start_bailed_to_cold:
            self._print_out(
                f"CG warm start: bailed to a cold restart at iteration {cg_info.get('bailout_iter')} "
                f"(not converging fast enough); {len(history)} total PCG iterations this solve.",
                "Verbose")

        # Refresh the shared state with THIS solve's own result, whether it
        # ran warm or cold, so the next inner iteration has the freshest
        # possible starting guess. Only the real inner-loop call site ever
        # supplies warm_start_state (pcls_optimiser.py), so probe and
        # final/exported solves -- which never receive it -- correctly never
        # read OR write it, and can't contaminate the real trajectory.
        if state is not None:
            state["u_prev"] = u_h.astype(np.float32, copy=True)
            state["rho_prev"] = field.astype(np.float32, copy=True)

        # Report a solve that never met its tolerance. This used to be visible
        # only in the (throttled) CG log file, so a run could spend its entire
        # budget on unconverged solves without anything saying so.
        stop_tol = max(self.pcg_tol, self.pcg_rel_tol * float(np.linalg.norm(b)))
        if not history or history[-1] >= stop_tol:
            self._print_out(
                f"Warning: PCG did not converge in {self.pcg_max_iters} iterations "
                f"(final residual {history[-1] if history else float('nan'):.3e} vs "
                f"tolerance {stop_tol:.3e}). The strain energy field, and therefore "
                f"this iteration's sensitivity, is only as accurate as that residual.",
                "Minimal")

        # Solver statistics, logged per iteration by the optimiser
        # (solver_iterations in variables.csv).
        self.solver_iterations = len(history)
        self.solver_final_residual = float(history[-1]) if history else float('nan')
        self.solver_converged = bool(history) and history[-1] < stop_tol

        # The EXACT compliance, c = b . u, in joules. See Conformal._fea.
        self.compliance = float(np.dot(b.astype(np.float64), u_h.astype(np.float64)))

        save_log = False
        if "inner_" in self.iter_name:
            try:
                # Extract the number at the very end of the string (e.g., "010" -> 10)
                inner_iter = int(self.iter_name.split("_")[-1])
                if inner_iter % 10 == 0:
                    save_log = True
            except ValueError:
                pass

        if save_log:
            log_filename = os.path.join(self.analysis_dir, f"{self.iter_name}_cg_log.txt")
            try:
                with open(log_filename, "w") as f:
                    f.write(f"Conjugate Gradient Convergence Log\n")
                    f.write(f"Iteration Name: {self.iter_name}\n")
                    f.write(f"Max Iters: {self.pcg_max_iters} | Absolute tol: {self.pcg_tol} "
                            f"| Relative tol: {self.pcg_rel_tol} | Effective tol: {stop_tol:.6e}\n")
                    f.write("-" * 40 + "\n")

                    for i, rnorm in enumerate(history):
                        f.write(f"Iter {i:4d} rnorm {rnorm:.6e}\n")

                    if history and history[-1] < stop_tol:
                        f.write("-" * 40 + "\nConverged successfully.\n")
                    else:
                        f.write("-" * 40 + "\nReached maximum iterations without converging.\n")
                        
                self._print_out(f"Saved CG convergence log to {log_filename}", "Verbose")
            except Exception as e:
                self._print_out(f"Warning: Failed to save CG log. Error: {e}", "Minimal")

        # Compute per-element energies on GPU and convert to SED
        try:
            element_energies = compute_element_strain_energy(
                nx, ny, nz, np.float32(hx), np.float32(hy), np.float32(hz),
                E_elem, nu_elem, u_h,
                element_type=self.element_type, nu=self.poisson_ratio
            )
        except Exception as e:
            raise RuntimeError(f"compute_element_strain_energy failed: {e}")

        elem_vol = float(hx * hy * hz)
        element_sed = (element_energies.astype(np.float32) / np.float32(elem_vol)).astype(np.float32)
        # r > 1: per-design-voxel SED = mean of the children's = sum of their energies / voxel volume
        element_sed = self._mean_children(element_sed)

        # correct element grid dimensions (design voxels)
        nx_e = (nx - 1) // r_ref
        ny_e = (ny - 1) // r_ref
        nz_e = (nz - 1) // r_ref

        # reshape into element grid (NOT field.shape)
        voxel_sed = element_sed.reshape((nx_e, ny_e, nz_e), order='C')

        self._print_out(f"ConjugateGradient finished in {timer() - t0:.2f}s."
                        f"\n  Max displacement (mm): {1000 * np.max(u_h):.4e}"
                        f"\n  SED min/max: {voxel_sed.min():.4e}/{voxel_sed.max():.4e}", "Verbose")

        if self.save_meshes:
            # Standardised iteration .vtr (docs/theory.md §7.5/§7.6). This
            # backend already wrote exactly this trio of CellData arrays in
            # one file before this revision -- it's the pattern the other
            # two backends are now brought in line with -- so the only
            # change here is deferring the actual write to
            # finalize_iteration_output() (via _pending_cell_data /
            # _pending_point_data) instead of calling _save_rectilinear_vtr
            # directly, so Displacement_Magnitude/Von_Mises/buckling-mode
            # fields computed below (and any buckling mode shape added
            # later, by the optimiser) land in the SAME file.
            self._grid_shape = (nx_e, ny_e, nz_e)
            self._grid_hx, self._grid_hy, self._grid_hz = float(hx) * r_ref, float(hy) * r_ref, float(hz) * r_ref

            rho_vox = self._mean_children(rho_flat)
            voxel_rho = rho_vox.reshape((nx_e, ny_e, nz_e), order='C')
            voxel_active = (rho_vox >= 0.5).astype(np.float32).reshape((nx_e, ny_e, nz_e), order='C')

            self._pending_cell_data["Strain_Energy_Density"] = voxel_sed.astype(np.float32)
            self._pending_cell_data["Density"] = voxel_rho.astype(np.float32)
            self._pending_cell_data["Active_Voxels"] = voxel_active

            # Displacement magnitude (PointData, exact -- no interpolation:
            # gpu_analysis's node id nid(i,j,k) = i*(ny*nz) + j*nz + k matches a
            # plain C-order reshape directly, and DOFs are interleaved
            # [ux0, uy0, uz0, ux1, ...] per node -- see hex_dof_map).
            u_nodal = u_h.reshape((nx, ny, nz, 3), order='C')[::r_ref, ::r_ref, ::r_ref]   # design-voxel nodes
            self._pending_point_data["Displacement_Magnitude"] = np.linalg.norm(
                u_nodal, axis=-1
            ).astype(np.float32)

            # Von Mises stress from gpu_analysis's element-averaged Voigt stress
            # kernel (compute_element_stress / _element_stress_kernel).
            # Previously only ever invoked from run_buckling's
            # verify_against_gpu cross-check path -- this is the same,
            # already-validated kernel, just also called here so its output
            # reaches every saved iteration's .vtr, not only a debug check.
            try:
                element_stress = compute_element_stress(
                    nx, ny, nz, np.float32(hx), np.float32(hy), np.float32(hz),
                    E_elem, nu_elem, u_h
                )
                element_von_mises = _von_mises_from_voigt(element_stress.astype(np.float64))
                element_von_mises = self._mean_children(element_von_mises)    # r > 1: volume mean of the children
                self._pending_cell_data["Von_Mises_Stress"] = element_von_mises.reshape(
                    (nx_e, ny_e, nz_e), order='C'
                ).astype(np.float32)
            except Exception as e:
                self._print_out(f"Warning: compute_element_stress failed; Von_Mises_Stress will be "
                                 f"omitted from this iteration's .vtr. Error: {e}", "Minimal")

            self._print_out(f"Prepared standardised CG iteration data (written by "
                             f"finalize_iteration_output).", "Verbose")

        # Store enough state for run_buckling() to reuse without a full
        # re-solve. u_h and rho_flat in particular are the expensive-to-
        # obtain results this is meant to avoid recomputing.
        self._last_nx, self._last_ny, self._last_nz = nx, ny, nz
        self._last_hx, self._last_hy, self._last_hz = float(hx), float(hy), float(hz)
        self._last_E_elem, self._last_nu_elem = E_elem, nu_elem
        self._last_u_h = u_h
        self._last_fixed_mask = fixed_mask
        self._last_rho_flat = rho_flat
        self._last_num_active = num_active
        self._buckling_voxel_shape = (nx_e, ny_e, nz_e)

        return voxel_sed, mesh_volume

    # ---------------------------------------------------------------
    # --- Engine selection (gpu_analysis.py; docs/PLAN_engine_buckling.md) ---
    # ---------------------------------------------------------------

    def _apply_elastic_engine(self):
        """Resolve `engine` (not "classic") once per process and apply it to this instance's preconditioner /
        GMG reuse / refinement. Returns the engine name."""
        import gpu_analysis
        cap = gpu_analysis.probe(gpu=self.engine != "cpu_direct")
        # design-voxel grid from the CSG constraint mask (elements, shape (nx_e, ny_e, nz_e, 3))
        nx, ny, nz = (int(d) + 1 for d in np.asarray(self.csg_data["fixed_nodes"]).shape[:3])
        eng, pre, reuse, r, reason = gpu_analysis.elastic_choice(cap, nx, ny, nz, self.engine, self.analysis_refinement)
        if eng.startswith("gpu") and not cap.get("cuda"):
            raise RuntimeError(f"elastic engine {eng!r} needs a CUDA device ({cap.get('cuda_error', 'none found')}).")
        if gpu_analysis.record("elastic", requested=self.engine, selected=eng, reason=reason, preconditioner=pre,
                          gmg_reuse_max_rel_change=reuse, analysis_refinement=r):
            print(f"engines: elastic = {eng} ({reason}). {gpu_analysis.summary_lines(cap)}", flush=True)
        gpu_analysis.record_use("elastic", eng)
        if pre is not None:
            self.preconditioner = pre
        if reuse is not None:
            self.gmg_reuse_max_rel_change = reuse
        self.analysis_refinement = r
        return eng

    def _run_analysis_cpu(self, field: np.ndarray):
        """Elastic engine "cpu_direct": the run_analysis contract on a machine without CUDA. Structured-grid
        stencil assembly of the same K as the GPU kernels (element_type honoured) and one host sparse Cholesky
        (MKL PARDISO / CHOLMOD / SuperLU); u stored in float32 as the GPU path does."""
        t0 = timer()
        (nx, ny, nz, hx, hy, hz, E_elem, nu_elem, b, fixed_mask, rho_flat) = self._volume_mesh(field)
        r_ref = self.analysis_refinement
        if r_ref == 1:
            num_active = int(np.count_nonzero(rho_flat >= 0.5))
        else:
            num_active = int(np.count_nonzero(self._mean_children(rho_flat) >= 0.5))
        mesh_volume = num_active * float(self.voxel_size_m ** 3)
        Ndofs = 3 * nx * ny * nz
        lam_elem, mu_elem = _lame_from_E_nu(E_elem.astype(np.float64), nu_elem.astype(np.float64))
        grid = grid_for(nx, ny, nz, (float(hx), float(hy), float(hz)), fixed_mask)
        if self.element_type == "q1_gauss":
            ka, kb, KA, KB = lam_elem, mu_elem, grid.Klam, grid.Kmu
        else:
            KA = element_khat(self.element_type, (float(hx), float(hy), float(hz)), self.poisson_ratio)
            ka, kb, KB = 2.0 * mu_elem * (1.0 + self.poisson_ratio), np.zeros_like(mu_elem), np.zeros((24, 24))
        K_II = assemble_K_cpu(grid, ka, kb, KA, KB)
        solver = HostCholesky().factor(K_II)
        bb = b.astype(np.float64)
        u = np.zeros(Ndofs)
        u[grid.free] = solver.solve(bb[grid.free])
        solver.free()
        res = float(np.linalg.norm(K_II @ u[grid.free] - bb[grid.free]) / max(np.linalg.norm(bb[grid.free]), 1e-300))
        del K_II
        self.compliance = float(np.dot(bb, u))
        u_e = u[grid.dof_map]
        if self.element_type == "q1_gauss":
            element_energy = 0.5 * (lam_elem * np.einsum("ec,cd,ed->e", u_e, grid.Klam, u_e)
                                    + mu_elem * np.einsum("ec,cd,ed->e", u_e, grid.Kmu, u_e))
        else:
            element_energy = 0.5 * ka * np.einsum("ec,cd,ed->e", u_e, KA, u_e)
        elem_vol = float(hx * hy * hz)
        nx_e, ny_e, nz_e = (nx - 1) // r_ref, (ny - 1) // r_ref, (nz - 1) // r_ref
        voxel_sed = self._mean_children(element_energy / elem_vol, np.float64).reshape(
            (nx_e, ny_e, nz_e), order='C').astype(np.float32)
        self.solver_iterations = 1
        self.solver_final_residual = res
        self.solver_converged = bool(res < 1e-8)
        if self.save_meshes:
            self._grid_shape = (nx_e, ny_e, nz_e)
            self._grid_hx, self._grid_hy, self._grid_hz = float(hx) * r_ref, float(hy) * r_ref, float(hz) * r_ref
            rho_vox = self._mean_children(rho_flat)
            self._pending_cell_data["Strain_Energy_Density"] = voxel_sed
            self._pending_cell_data["Density"] = rho_vox.reshape((nx_e, ny_e, nz_e), order='C').astype(np.float32)
            self._pending_cell_data["Active_Voxels"] = (rho_vox >= 0.5).astype(np.float32).reshape(
                (nx_e, ny_e, nz_e), order='C')
            self._pending_point_data["Displacement_Magnitude"] = np.linalg.norm(
                u.reshape((nx, ny, nz, 3), order='C')[::r_ref, ::r_ref, ::r_ref], axis=-1).astype(np.float32)
            B_all, _Bg, detJ = _hex_reference_data(float(hx), float(hy), float(hz))
            strain = self._cg_strain_voigt(u, grid.dof_map, B_all)
            stress = self._cg_stress_voigt(strain, lam_elem, mu_elem)
            self._pending_cell_data["Von_Mises_Stress"] = self._mean_children(
                _von_mises_from_voigt(stress.mean(axis=1)), np.float64).reshape(
                (nx_e, ny_e, nz_e), order='C').astype(np.float32)
        self._last_nx, self._last_ny, self._last_nz = nx, ny, nz
        self._last_hx, self._last_hy, self._last_hz = float(hx), float(hy), float(hz)
        self._last_E_elem, self._last_nu_elem = E_elem, nu_elem
        self._last_u_h = u.astype(np.float32)
        self._last_fixed_mask = fixed_mask
        self._last_rho_flat = rho_flat
        self._last_num_active = num_active
        self._buckling_voxel_shape = (nx_e, ny_e, nz_e)
        self._print_out(f"ConjugateGradient (cpu_direct, {solver.kind}) finished in {timer() - t0:.2f}s, "
                        f"c={self.compliance:.4e}, rel. residual {res:.1e}.", "Verbose")
        return voxel_sed, mesh_volume

    def _map_mode_shapes_to_grid(self, mode_shapes: np.ndarray) -> dict:
        """
        See `Analysis._map_mode_shapes_to_grid`. `mode_shapes` is (Ndofs, k)
        on this backend's own node/DOF numbering (`self._last_nx/_ny/_nz`,
        interleaved [ux,uy,uz] per node) -- the same convention `run_analysis`
        uses for `Displacement_Magnitude`, so this is the same direct C-order
        reshape, one column at a time.
        """
        nx, ny, nz = self._last_nx, self._last_ny, self._last_nz
        fields = {}
        for m in range(mode_shapes.shape[1]):
            mode_nodal = mode_shapes[:, m].reshape((nx, ny, nz, 3), order='C')
            fields[f"Buckling_Mode_{m + 1}_Magnitude"] = np.linalg.norm(mode_nodal, axis=-1).astype(np.float32)
        return fields

    # ---------------------------------------------------------------
    # --- Linear buckling analysis (host-side sparse assembly) ---
    #
    # See the module-level comment above `_hex_reference_data` for why
    # this is implemented via explicit host-assembled sparse matrices
    # rather than new CUDA kernels, and the implementation plan (section
    # 2 and 3.2, ConjugateGradient class notes) for the underlying
    # mathematics -- identical to the Regular/Conformal formulation, just
    # expressed in flat Voigt-array form instead of skfem tensor fields.
    # ---------------------------------------------------------------

    def _cg_strain_voigt(self, u_vector: np.ndarray, dof_map: np.ndarray, B_all: np.ndarray):
        """Per-element, per-Gauss-point Voigt strain: shape (Ne, 8, 6)."""
        u_e = u_vector[dof_map]  # (Ne, 24)
        return np.einsum('gvc,ec->egv', B_all, u_e)

    def _cg_stress_voigt(self, strain_voigt: np.ndarray, lam_elem: np.ndarray, mu_elem: np.ndarray):
        """Actual (rho-scaled) constitutive law in Voigt form: shape (Ne, 8, 6)."""
        stress_lam = np.einsum('vw,egw->egv', _VOIGT_C_LAM_BASIS, strain_voigt)
        stress_mu = np.einsum('vw,egw->egv', _VOIGT_C_MU_BASIS, strain_voigt)
        return lam_elem[:, None, None] * stress_lam + mu_elem[:, None, None] * stress_mu

    def _cg_stress_voigt_prime(self, strain_voigt: np.ndarray, lam_strong: float, mu_strong: float):
        """
        d(constitutive law)/d(rho_e), evaluated at strain_voigt. As in
        Regular, this is rho-independent because the PCLS interpolation is
        linear in rho (C_e = rho_e * C_strong).
        """
        stress_lam = np.einsum('vw,egw->egv', _VOIGT_C_LAM_BASIS, strain_voigt)
        stress_mu = np.einsum('vw,egw->egv', _VOIGT_C_MU_BASIS, strain_voigt)
        return lam_strong * stress_lam + mu_strong * stress_mu

    def _cg_gamma_voigt(self, mode_vector: np.ndarray, dof_map: np.ndarray, Bg_all: np.ndarray):
        """
        Builds the 'geometric strain' Gamma_ij = sum_k (d phi_k/dx_i)(d phi_k/dx_j)
        in the same doubled-off-diagonal Voigt convention as strain
        (so that a plain Voigt dot product against a Voigt stress gives
        the correct full tensor double contraction sigma:Gamma). Shape
        (Ne, 8, 6): (G00, G11, G22, 2*G12, 2*G02, 2*G01).
        """
        phi_e = mode_vector[dof_map].reshape(-1, 8, 3)  # (Ne, node, component)
        # grad_phi[e,gp,k,i] = d phi_k / d x_i
        grad_phi = np.einsum('gia,eak->egki', Bg_all, phi_e)

        g00 = np.einsum('egk,egk->eg', grad_phi[:, :, :, 0], grad_phi[:, :, :, 0])
        g11 = np.einsum('egk,egk->eg', grad_phi[:, :, :, 1], grad_phi[:, :, :, 1])
        g22 = np.einsum('egk,egk->eg', grad_phi[:, :, :, 2], grad_phi[:, :, :, 2])
        g12 = np.einsum('egk,egk->eg', grad_phi[:, :, :, 1], grad_phi[:, :, :, 2])
        g02 = np.einsum('egk,egk->eg', grad_phi[:, :, :, 0], grad_phi[:, :, :, 2])
        g01 = np.einsum('egk,egk->eg', grad_phi[:, :, :, 0], grad_phi[:, :, :, 1])

        return np.stack([g00, g11, g22, 2.0 * g12, 2.0 * g02, 2.0 * g01], axis=-1)

    def _cg_ddot_mean(self, field_a: np.ndarray, field_b: np.ndarray, detJ: float):
        """
        Per-element VOLUME INTEGRAL (not average) Voigt dot product:

            result_e = detJ * sum_gp( sum_v field_a[e,gp,v] * field_b[e,gp,v] )

        `detJ` is the (constant, for this uniform grid) Jacobian
        determinant per Gauss point -- see `_ddot_field_mean` on the
        Analysis base class for why the volume weighting must be applied
        explicitly here rather than averaged away.
        """
        return detJ * np.sum(np.einsum('egv,egv->eg', field_a, field_b), axis=1)

    def _cg_assemble_stiffness(self, dof_map, lam_elem, mu_elem, B_all, detJ, Ndofs):
        """Vectorized host assembly of the elastic stiffness matrix K."""
        if self.element_type == "q1_gauss":
            Ke_lam = np.einsum('gvc,vw,gwd->cd', B_all, _VOIGT_C_LAM_BASIS, B_all) * detJ
            Ke_mu = np.einsum('gvc,vw,gwd->cd', B_all, _VOIGT_C_MU_BASIS, B_all) * detJ
            Ke_all = lam_elem[:, None, None] * Ke_lam[None, :, :] + mu_elem[:, None, None] * Ke_mu[None, :, :]
        else:
            # K_e = E_e * Khat (precomputed "q1" / condensed "q1_im" matrix); E_e = 2 mu_e (1 + nu)
            Khat = element_khat(self.element_type,
                                (float(self._last_hx), float(self._last_hy), float(self._last_hz)),
                                self.poisson_ratio)
            E_e = 2.0 * mu_elem * (1.0 + self.poisson_ratio)
            Ke_all = E_e[:, None, None] * Khat[None, :, :]

        rows = np.repeat(dof_map[:, :, None], 24, axis=2)
        cols = np.repeat(dof_map[:, None, :], 24, axis=1)
        K = coo_matrix((Ke_all.ravel(), (rows.ravel(), cols.ravel())), shape=(Ndofs, Ndofs))
        return K.tocsr()

    def _cg_assemble_geometric_stiffness(self, dof_map, stress_voigt, Bg_all, detJ, Ndofs):
        """
        Vectorized host assembly of K_sigma from a per-element, per-Gauss-
        point Voigt stress field (Ne, 8, 6). See implementation plan
        section 2.1 for the sign convention (G = -K_sigma is formed by the
        caller).
        """
        Ne = stress_voigt.shape[0]
        sigma_tensor = np.zeros((Ne, 8, 3, 3))
        sigma_tensor[:, :, 0, 0] = stress_voigt[:, :, 0]
        sigma_tensor[:, :, 1, 1] = stress_voigt[:, :, 1]
        sigma_tensor[:, :, 2, 2] = stress_voigt[:, :, 2]
        sigma_tensor[:, :, 1, 2] = sigma_tensor[:, :, 2, 1] = stress_voigt[:, :, 3]
        sigma_tensor[:, :, 0, 2] = sigma_tensor[:, :, 2, 0] = stress_voigt[:, :, 4]
        sigma_tensor[:, :, 0, 1] = sigma_tensor[:, :, 1, 0] = stress_voigt[:, :, 5]

        # Kg_scalar[e,a,b] = sum_gp detJ * sum_ij Bg[gp,i,a] * sigma[e,gp,i,j] * Bg[gp,j,b]
        Kg_scalar = np.einsum('gia,egij,gjb->eab', Bg_all, sigma_tensor, Bg_all) * detJ  # (Ne,8,8)

        Kg_full = np.zeros((Ne, 24, 24))
        for c in range(3):
            Kg_full[:, c::3, c::3] = Kg_scalar

        rows = np.repeat(dof_map[:, :, None], 24, axis=2)
        cols = np.repeat(dof_map[:, None, :], 24, axis=1)
        Ksig = coo_matrix((Kg_full.ravel(), (rows.ravel(), cols.ravel())), shape=(Ndofs, Ndofs))
        return Ksig.tocsr()

    def _cg_assemble_adjoint_load(self, dof_map, gamma_combined_voigt, lam_elem, mu_elem, B_all, detJ, Ndofs):
        """
        Vectorized host assembly of the adjoint pseudo-load vector g from
        the already-aggregated, weighted geometric tensor field
        `gamma_combined_voigt` (Ne, 8, 6). See implementation plan section
        2.4 -- this is the CG-specific (Voigt-array) equivalent of the
        skfem-based `Analysis._assemble_adjoint_load`.
        """
        # CB_lam[gp,v,a] = (C_lam_basis @ B_all[gp])[v,a]; universal, no rho.
        CB_lam = np.einsum('vw,gwa->gva', _VOIGT_C_LAM_BASIS, B_all)
        CB_mu = np.einsum('vw,gwa->gva', _VOIGT_C_MU_BASIS, B_all)

        proj_lam = np.einsum('gva,egv->ega', CB_lam, gamma_combined_voigt)  # (Ne,8,24)
        proj_mu = np.einsum('gva,egv->ega', CB_mu, gamma_combined_voigt)

        g_local = detJ * (
            lam_elem[:, None] * proj_lam.sum(axis=1) + mu_elem[:, None] * proj_mu.sum(axis=1)
        )  # (Ne, 24)

        g_global = np.zeros(Ndofs)
        np.add.at(g_global, dof_map.ravel(), g_local.ravel())
        return g_global

    def _modal_energy_fraction_in_solid(self, strain_voigt_mode: np.ndarray, lam_elem, mu_elem, rho_flat, detJ,
                                        solid_weight=None):
        """CG-specific spurious-void-mode filter input (ramp-weighted energy share); see Regular's docstring."""
        stress_mode = self._cg_stress_voigt(strain_voigt_mode, lam_elem, mu_elem)
        element_energy = np.clip(self._cg_ddot_mean(stress_mode, strain_voigt_mode, detJ), 0.0, None)
        total = element_energy.sum()
        if total <= 0.0:
            return 0.0
        if solid_weight is None:
            solid_weight = (rho_flat >= 0.5).astype(np.float64)
        return float(np.dot(element_energy, solid_weight) / total)

    def run_buckling(self, num_modes: int = 6, aggregation_p: float = 8.0,
                      prev_modes: np.ndarray | None = None,
                      min_energy_fraction: float = 0.5,
                      mode_continuity_warning_threshold: float = 0.7,
                      min_density_for_geometric_stiffness: float = 0.1,
                      verify_against_gpu: bool = False,
                      eigensolver_tol: float = 1e-12,
                      engine: str | None = None,
                      lobpcg_tol: float | None = None) -> BucklingResults:
        """
        Runs a linear buckling eigenvalue analysis using the primal
        solution computed by the most recent `run_analysis` call.

        Must be called AFTER `run_analysis` in the same iteration: it
        reuses the mesh/material/solution state stored there. See the
        "ConjugateGradient buckling: design notes" comment block above
        this class for why this assembles explicit sparse matrices on the
        host rather than using new CUDA kernels, and implementation plan
        section 2 for the mathematics.

        Args: as in Regular.run_buckling, plus
        - min_density_for_geometric_stiffness (float): rho threshold below
          which the geometric stiffness contribution is ramped to zero
          (section 4.1 of the implementation plan).
        - verify_against_gpu (bool): if True, cross-check the host-computed
          element stress against `gpu_analysis.compute_element_stress` on this
          exact problem and warn on disagreement. The host path and the
          CUDA kernels are two implementations of the same element
          formulation; this catches the case where one has been edited and
          the other has not. Costs one extra GPU round trip, so it is off
          by default -- but running it once at the start of a session is
          cheap insurance. `gpu_analysis.verify_host_kernel_consistency()` is the
          more thorough offline equivalent.
        - eigensolver_tol (float): ARPACK's relative eigenvalue tolerance
          (default 1e-12; 0.0 = machine precision). See
          `_solve_generalized_buckling_eigenproblem`.
        - engine (str | None): buckling engine (gpu_analysis.py, docs/README.md ("Solver engines")); None = the process
          default set by the optimiser from optimiser.buckling.engine ("auto" unless configured).
          "classic" is the code below, unchanged. Any other engine runs `_run_buckling_fast` (same
          formulation, filter, aggregation and outputs) and falls back along its chain on failure.
        - lobpcg_tol (float | None): LOBPCG relative residual tolerance of the fast engines (None =
          optimiser.buckling.lobpcg_tol, default 1e-3; the eigenvalue error is ~ tol^2).

        Returns:
        - BucklingResults
        """
        if self.analysis_refinement != 1:
            raise NotImplementedError("run_buckling is not available with analysis_refinement > 1 "
                                      "(the buckling path assumes one element per design voxel).")
        if not hasattr(self, '_last_u_h'):
            raise RuntimeError("run_buckling() called before run_analysis(); no primal solution available.")
        if not np.any(self._last_rho_flat > min_density_for_geometric_stiffness):
            # The ramp zeroes K_sigma everywhere: no element can buckle.
            self._print_out("No element above the geometric-stiffness threshold: "
                            "no compressive buckling mode can exist.", "Minimal")
            Ndofs = 3 * self._last_nx * self._last_ny * self._last_nz
            return BucklingResults(
                mu_values=np.zeros(0), mode_shapes=np.zeros((Ndofs, 0)),
                mu_tilde=0.0, sensitivity=np.zeros(self._buckling_voxel_shape),
                mac_matrix=None, solver_iters=0, n_candidates=0,
            )

        if engine is None or lobpcg_tol is None:
            from gpu_analysis import buckling_defaults
            _bd = buckling_defaults()
            engine = _bd["buckling_engine"] if engine is None else engine
            lobpcg_tol = _bd["lobpcg_tol"] if lobpcg_tol is None else lobpcg_tol
        if engine != "classic" and not _engines_force_classic():
            fast = self._run_buckling_fast(engine, float(lobpcg_tol), num_modes, aggregation_p, prev_modes,
                                           min_energy_fraction, mode_continuity_warning_threshold,
                                           min_density_for_geometric_stiffness, eigensolver_tol)
            if fast is not None:
                return fast
            # the engine chain ended at "classic": continue below

        nx, ny, nz = self._last_nx, self._last_ny, self._last_nz
        hx, hy, hz = self._last_hx, self._last_hy, self._last_hz
        rho_flat = self._last_rho_flat
        fixed_mask = self._last_fixed_mask
        u_h = self._last_u_h.astype(np.float64)
        Ndofs = 3 * nx * ny * nz

        lam_elem, mu_elem = _lame_from_E_nu(
            self._last_E_elem.astype(np.float64), self._last_nu_elem.astype(np.float64)
        )
        lam_strong, mu_strong = lame_parameters(self.youngs_modulus, self.poisson_ratio)

        dof_map = _hex_dof_map(nx, ny, nz)
        B_all, Bg_all, detJ = _hex_reference_data(hx, hy, hz)

        # 1) Elastic stiffness, primal stress field, and G = -K_sigma
        K = self._cg_assemble_stiffness(dof_map, lam_elem, mu_elem, B_all, detJ, Ndofs)
        strain_u0 = self._cg_strain_voigt(u_h, dof_map, B_all)
        stress_u0_actual = self._cg_stress_voigt(strain_u0, lam_elem, mu_elem)

        if verify_against_gpu:
            try:
                gpu_stress = compute_element_stress(
                    nx, ny, nz, np.float32(hx), np.float32(hy), np.float32(hz),
                    self._last_E_elem, self._last_nu_elem, self._last_u_h
                ).astype(np.float64)
                host_stress = np.mean(stress_u0_actual, axis=1)
                scale = max(np.abs(gpu_stress).max(), 1e-30)
                err = np.abs(host_stress - gpu_stress).max() / scale
                if err < 2e-4:
                    self._print_out(
                        f"GPU stress cross-check passed (max relative difference "
                        f"{err:.2e}, within float32 round-off).", "Verbose"
                    )
                else:
                    self._print_out(
                        f"Warning: host-computed element stress disagrees with "
                        f"gpu_analysis.compute_element_stress by {err:.2e} relative, which "
                        f"exceeds float32 round-off. The host reference data in "
                        f"gpu_analysis.hex_reference_data may have drifted from the CUDA "
                        f"kernels. Run gpu_analysis.verify_host_kernel_consistency() -- "
                        f"buckling results are not trustworthy until this is "
                        f"resolved.", "Minimal"
                    )
            except Exception as e:
                self._print_out(f"GPU stress cross-check could not be run: {e}", "Minimal")

        # Void-region safety ramp (section 4.1). Applied consistently to
        # the stress used for K_sigma AND to the explicit/adjoint terms
        # below that derive from it. The ramp's own dependence on rho_e is
        # captured explicitly via `ramp_prime` and Term C below (same
        # product-rule split as the skfem-based `_run_buckling_fe`; see
        # docs/theory.md §6.5(b)/§9.1).
        threshold = min_density_for_geometric_stiffness
        upper = 2.0 * threshold
        denom = max(upper - threshold, 1e-12)
        ramp_x = np.clip((rho_flat.astype(np.float64) - threshold) / denom, 0.0, 1.0)
        ramp = (ramp_x * ramp_x * (3.0 - 2.0 * ramp_x))[:, None, None]  # (Ne,1,1), broadcasts over (Ne,8,6)
        # d(ramp)/d(rho_e); see Regular._geometric_stiffness_ramp_prime for
        # why re-using the clipped ramp_x already gives the right value
        # (including 0) at and beyond the threshold/2*threshold boundary.
        ramp_prime = (6.0 * ramp_x * (1.0 - ramp_x)) / denom  # (Ne,)

        stress_u0_ramped = stress_u0_actual * ramp

        K_sigma = self._cg_assemble_geometric_stiffness(dof_map, stress_u0_ramped, Bg_all, detJ, Ndofs)
        G = -K_sigma

        free_dofs = np.where(fixed_mask == 0)[0]
        K_II = K[free_dofs][:, free_dofs].tocsr()
        G_II = G[free_dofs][:, free_dofs].tocsr()

        # 2) Eigensolve
        mu_values, modes_free, eig_converged = self._solve_generalized_buckling_eigenproblem(
            K_II, G_II, num_modes, tol=eigensolver_tol)

        if mu_values.size == 0:
            self._print_out("Buckling eigensolver returned no modes (no compressive mode found)."
                            if eig_converged else
                            "Buckling eigensolver converged nothing: the buckling load is UNKNOWN.",
                            "Minimal")
            return BucklingResults(
                mu_values=np.zeros(0), mode_shapes=np.zeros((Ndofs, 0)),
                mu_tilde=0.0, sensitivity=np.zeros(self._buckling_voxel_shape),
                mac_matrix=None, solver_iters=0, n_candidates=0,
                eigensolver_converged=eig_converged,
            )

        modes_full = np.zeros((Ndofs, modes_free.shape[1]))
        modes_full[free_dofs, :] = modes_free
        n_candidates = int(np.count_nonzero(mu_values > 0.0))

        # 3) Filter spurious modes, weighting each element's share of the
        #    modal strain energy by the same geometric-stiffness ramp that
        #    scales K_sigma (docs/theory.md §6.5(a)).
        solid_weight = ramp.reshape(-1)
        energy_fraction = np.zeros(modes_full.shape[1])
        for j in range(modes_full.shape[1]):
            strain_phi_j = self._cg_strain_voigt(modes_full[:, j], dof_map, B_all)
            energy_fraction[j] = self._modal_energy_fraction_in_solid(strain_phi_j, lam_elem, mu_elem, rho_flat, detJ,
                                                                      solid_weight=solid_weight)

        mu_values, modes_full, _kept = self._filter_valid_modes(
            mu_values, modes_full, energy_fraction, min_energy_fraction
        )

        if mu_values.size == 0:
            if n_candidates > 0:
                self._print_out(
                    f"No buckling mode retained: all {n_candidates} candidate mode(s) with a "
                    f"positive eigenvalue were rejected by the spurious-mode filter "
                    f"(min_energy_fraction={min_energy_fraction}). The buckling load of this "
                    f"design is UNKNOWN for this evaluation (not assumed safe).", "Minimal"
                )
            else:
                self._print_out("No compressive buckling mode exists for this design and load "
                                "case: the buckling constraint is satisfied.", "Normal")
            return BucklingResults(
                mu_values=np.zeros(0), mode_shapes=np.zeros((Ndofs, 0)),
                mu_tilde=0.0, sensitivity=np.zeros(self._buckling_voxel_shape),
                mac_matrix=None, solver_iters=0, n_candidates=n_candidates,
            )

        mac_matrix = self._compute_mac(modes_full, prev_modes)
        overlap = self._subspace_continuity(modes_full, prev_modes)
        if overlap is not None and overlap.size:
            worst = float(overlap.min())
            if worst < mode_continuity_warning_threshold:
                self._print_out(
                    f"Note: the least-continuous buckling mode overlaps the previous "
                    f"mode subspace by only {worst:.3f} (< {mode_continuity_warning_threshold}). Either "
                    f"a mode entered or left the tracked set, or the design moved a long way "
                    f"since the last evaluation. If it persists, raise num_modes or reduce "
                    f"dt / buckling_evaluate_every.", "Minimal"
                )

        # 4) Check the probe mode for cluster truncation, then discard it
        #    and trim the retained set to num_modes before aggregating.
        if self._check_cluster_truncation(mu_values, num_modes):
            self._print_out(
                f"Warning: the retained buckling mode set would end part-way through a "
                f"cluster of repeated eigenvalues (mode {num_modes} and mode "
                f"{num_modes + 1} agree to within 0.1%). A partially-retained "
                f"degenerate cluster makes the aggregated sensitivity basis-dependent "
                f"and therefore ill-defined. Increase num_modes (currently "
                f"{num_modes}) until the retained set ends on a clear spectral gap.",
                "Minimal"
            )

        if mu_values.size > num_modes:
            mu_values = mu_values[:num_modes]
            modes_full = modes_full[:, :num_modes]

        mu_tilde, weights = self._aggregate_mu(mu_values, aggregation_p)
        self._print_out(
            f"Buckling: retained {mu_values.size} mode(s), "
            f"lambda_crit (min) = {1.0 / mu_values.max():.4f}, "
            f"mu_tilde = {mu_tilde:.6e}", "Normal"
        )

        # 5) Explicit sensitivity terms + aggregated geometric tensor
        Ne = rho_flat.size
        explicit_sensitivity = np.zeros(Ne)
        gamma_combined = None

        for j in range(mu_values.size):
            phi_j = modes_full[:, j]
            strain_phi = self._cg_strain_voigt(phi_j, dof_map, B_all)
            # UNRAMPED geometric tensor -- needed for Term C below (the
            # ramp'(rho_e) * sigma(u0) piece of the product rule); must not
            # have the ramp baked in a second time.
            gamma_j_raw = self._cg_gamma_voigt(phi_j, dof_map, Bg_all)
            # Ramped for the same reason as in _run_buckling_fe: Term B and
            # the adjoint load (via gamma_combined) must differentiate the
            # SAME ramped K_sigma operator that was actually assembled above.
            gamma_j = gamma_j_raw * ramp

            term_a = -mu_values[j] * self._cg_ddot_mean(
                self._cg_stress_voigt_prime(strain_phi, lam_strong, mu_strong), strain_phi, detJ
            )
            # ramp(rho_e) * d(sigma(u0))/drho_e piece of the product rule.
            term_b = -self._cg_ddot_mean(
                self._cg_stress_voigt_prime(strain_u0, lam_strong, mu_strong), gamma_j, detJ
            )
            # ramp'(rho_e) * sigma(u0) piece of the product rule, previously
            # neglected (docs/theory.md §6.5(b)/§9.1). Uses the ACTUAL
            # (unramped) stress and the unramped gamma_j_raw so the ramp
            # itself is not double-counted.
            term_c = -ramp_prime * self._cg_ddot_mean(stress_u0_actual, gamma_j_raw, detJ)
            explicit_sensitivity += weights[j] * (term_a + term_b + term_c)

            gamma_combined = gamma_j * weights[j] if gamma_combined is None \
                else gamma_combined + weights[j] * gamma_j

        # 6) One adjoint solve, reusing the SAME K_II object (and PARDISO
        #    factorisation) as the eigensolver.
        g_combined = self._cg_assemble_adjoint_load(dof_map, gamma_combined, lam_elem, mu_elem, B_all, detJ, Ndofs)
        nu_free = pardiso_solve(K_II, g_combined[free_dofs])
        nu_full = np.zeros(Ndofs)
        nu_full[free_dofs] = nu_free

        strain_nu = self._cg_strain_voigt(nu_full, dof_map, B_all)
        implicit_sensitivity = self._cg_ddot_mean(
            self._cg_stress_voigt_prime(strain_nu, lam_strong, mu_strong), strain_u0, detJ
        )

        d_mu_tilde_d_rho = explicit_sensitivity + implicit_sensitivity

        nx_e, ny_e, nz_e = self._buckling_voxel_shape
        sensitivity_voxels = d_mu_tilde_d_rho.reshape((nx_e, ny_e, nz_e), order='C')

        return BucklingResults(
            mu_values=mu_values,
            mode_shapes=modes_full,
            mu_tilde=mu_tilde,
            sensitivity=sensitivity_voxels,
            mac_matrix=mac_matrix,
            subspace_continuity=overlap,
            solver_iters=0,
            n_candidates=n_candidates,
        )


    def _run_buckling_fast(self, requested, lobpcg_tol, num_modes, aggregation_p, prev_modes, min_energy_fraction,
                           mode_continuity_warning_threshold, threshold, eigensolver_tol):
        """
        run_buckling through a fast engine (the former buckling_fast.py section below) chosen by gpu_analysis.buckling_chain. Same outputs as the
        classic path below it. A failing engine (e.g. cuDSS ALLOC_FAILED, a non-finite float32 factor, an import
        error) is excluded for the rest of the process and the next engine of the chain is tried, with one printed
        line and an entry in gpu_analysis.RECORD["fallbacks"]. Returns None when the chain reaches "classic".
        """
        import gpu_analysis
        nx, ny, nz = self._last_nx, self._last_ny, self._last_nz
        n_free = int(np.count_nonzero(np.asarray(self._last_fixed_mask) == 0))
        cap = gpu_analysis.probe(gpu=requested not in ("cpu_cholesky_lobpcg", "cpu_cholesky_eigsh"))
        chain, skipped = gpu_analysis.buckling_chain(cap, nx, ny, nz, n_free, num_modes, requested)
        chain = [c for c in chain if c[0] not in gpu_analysis.FAILED]
        if not chain:
            raise RuntimeError(f"no buckling engine can run here: {skipped}")
        if gpu_analysis.record("buckling", requested=requested, selected=chain[0][0], reason=chain[0][1],
                          chain=[c[0] for c in chain], skipped=skipped, lobpcg_tol=lobpcg_tol):
            print(f"engines: buckling = {chain[0][0]} ({chain[0][1]}); fallback chain "
                  f"{' -> '.join(c[0] for c in chain)}. {gpu_analysis.summary_lines(cap)}", flush=True)
        for i, (name, _reason) in enumerate(chain):
            if name == "classic":
                gpu_analysis.record_use("buckling", "classic")
                return None
            try:
                res = self._buckling_fast_eval(name, cap, lobpcg_tol, num_modes, aggregation_p, prev_modes,
                                               min_energy_fraction, mode_continuity_warning_threshold, threshold,
                                               eigensolver_tol)
                gpu_analysis.record_use("buckling", name)
                return res
            except Exception as e:  # noqa: BLE001 -- any engine failure: next engine
                gpu_analysis.FAILED.add(name)
                nxt = chain[i + 1][0] if i + 1 < len(chain) else None
                print(f"engines: buckling engine {name} failed ({type(e).__name__}: {str(e)[:200]}); "
                      f"falling back to {nxt}.", flush=True)
                gpu_analysis.record_fallback("buckling", name, nxt, f"{type(e).__name__}: {e}")
                reset_state()
                if nxt is None:
                    raise
        return None

    def _buckling_fast_eval(self, name, cap, lobpcg_tol, num_modes, aggregation_p, prev_modes, min_energy_fraction,
                            mode_continuity_warning_threshold, threshold, eigensolver_tol):
        """One evaluation with fast engine `name`: eigenpairs, then the classic filter / continuity / truncation /
        aggregation steps (same code and messages), then the engine's sensitivity."""
        import gpu_analysis
        t_start = timer()
        nx, ny, nz = self._last_nx, self._last_ny, self._last_nz
        Ndofs = 3 * nx * ny * nz
        lam_strong, mu_strong = lame_parameters(self.youngs_modulus, self.poisson_ratio)
        h = (self._last_hx, self._last_hy, self._last_hz)
        Khat = None if self.element_type == "q1_gauss" else element_khat(self.element_type, h, self.poisson_ratio)
        eng = BucklingEngine(
            name, nx, ny, nz, h, self._last_fixed_mask, self._last_E_elem, self._last_nu_elem,
            self._last_u_h.astype(np.float64), self._last_rho_flat, threshold, lam_strong, mu_strong,
            self.element_type, self.poisson_ratio, Khat, lobpcg_tol=lobpcg_tol, eigensolver_tol=eigensolver_tol,
            hybrid_limit_bytes=gpu_analysis.hybrid_limit_bytes(cap) if "hybrid" in name else None,
            gmg_kwargs=dict(levels=self.gmg_levels, smoother_steps=self.gmg_smoother_steps,
                            coarse_max_nodes=self.gmg_coarse_max_nodes),
            log=lambda m: self._print_out(m, "Normal"))
        try:
            mu_values, modes_free, eig_converged, _cert = eng.eigen(num_modes)
            empty = dict(mu_values=np.zeros(0), mode_shapes=np.zeros((Ndofs, 0)), mu_tilde=0.0,
                         sensitivity=np.zeros(self._buckling_voxel_shape), mac_matrix=None, solver_iters=0)
            if mu_values.size == 0:
                self._print_out("Buckling eigensolver returned no modes (no compressive mode found)."
                                if eig_converged else
                                "Buckling eigensolver converged nothing: the buckling load is UNKNOWN.", "Minimal")
                return self._fast_done(eng, t_start, BucklingResults(**empty, n_candidates=0,
                                                                      eigensolver_converged=eig_converged))
            n_candidates = int(np.count_nonzero(mu_values > 0.0))
            energy_fraction = eng.energy_fraction(modes_free)
            keep = (mu_values > 0.0) & (energy_fraction >= min_energy_fraction)    # _filter_valid_modes
            kept = np.where(keep)[0]
            mu_values = mu_values[kept]
            modes_free = modes_free[:, eng.xp.asarray(kept)] if kept.size else modes_free[:, :0]
            if mu_values.size == 0:
                if n_candidates > 0:
                    self._print_out(
                        f"No buckling mode retained: all {n_candidates} candidate mode(s) with a "
                        f"positive eigenvalue were rejected by the spurious-mode filter "
                        f"(min_energy_fraction={min_energy_fraction}). The buckling load of this "
                        f"design is UNKNOWN for this evaluation (not assumed safe).", "Minimal")
                else:
                    self._print_out("No compressive buckling mode exists for this design and load "
                                    "case: the buckling constraint is satisfied.", "Normal")
                return self._fast_done(eng, t_start, BucklingResults(**empty, n_candidates=n_candidates))
            modes_full = eng.modes_full_host(modes_free)
            mac_matrix = self._compute_mac(modes_full, prev_modes)
            overlap = self._subspace_continuity(modes_full, prev_modes)
            if overlap is not None and overlap.size:
                worst = float(overlap.min())
                if worst < mode_continuity_warning_threshold:
                    self._print_out(
                        f"Note: the least-continuous buckling mode overlaps the previous "
                        f"mode subspace by only {worst:.3f} (< {mode_continuity_warning_threshold}). Either "
                        f"a mode entered or left the tracked set, or the design moved a long way "
                        f"since the last evaluation. If it persists, raise num_modes or reduce "
                        f"dt / buckling_evaluate_every.", "Minimal")
            if self._check_cluster_truncation(mu_values, num_modes):
                self._print_out(
                    f"Warning: the retained buckling mode set would end part-way through a "
                    f"cluster of repeated eigenvalues (mode {num_modes} and mode "
                    f"{num_modes + 1} agree to within 0.1%). A partially-retained "
                    f"degenerate cluster makes the aggregated sensitivity basis-dependent "
                    f"and therefore ill-defined. Increase num_modes (currently "
                    f"{num_modes}) until the retained set ends on a clear spectral gap.", "Minimal")
            if mu_values.size > num_modes:
                mu_values = mu_values[:num_modes]
                modes_full = modes_full[:, :num_modes]
                modes_free = modes_free[:, :num_modes]
            mu_tilde, weights = self._aggregate_mu(mu_values, aggregation_p)
            self._print_out(
                f"Buckling: retained {mu_values.size} mode(s), "
                f"lambda_crit (min) = {1.0 / mu_values.max():.4f}, "
                f"mu_tilde = {mu_tilde:.6e}", "Normal")
            sens = eng.sensitivity(mu_values, modes_free, weights)
            nx_e, ny_e, nz_e = self._buckling_voxel_shape
            return self._fast_done(eng, t_start, BucklingResults(
                mu_values=mu_values, mode_shapes=modes_full, mu_tilde=mu_tilde,
                sensitivity=sens.reshape((nx_e, ny_e, nz_e), order='C'), mac_matrix=mac_matrix,
                subspace_continuity=overlap, solver_iters=int(eng.diag.get("lobpcg_iters", 0)),
                n_candidates=n_candidates))
        finally:
            eng.release()

    def _fast_done(self, eng, t_start, res):
        """Timing / diagnostics of a fast evaluation (self.buckling_engine_info; gpu_analysis.RECORD summary)."""
        import gpu_analysis
        wall = timer() - t_start
        info = dict(eng.diag, wall_s=wall, **{f"t_{k}": v for k, v in eng.t.items()})
        self.buckling_engine_info = info
        tm = gpu_analysis.RECORD.setdefault("buckling", {}).setdefault("timing", {})
        tm.setdefault(eng.name, {"n": 0, "total_s": 0.0, "max_s": 0.0})
        e = tm[eng.name]
        e["n"] += 1; e["total_s"] += wall; e["max_s"] = max(e["max_s"], wall)
        gpu_analysis.RECORD["buckling"]["last"] = {k: (float(v) if isinstance(v, (np.floating, float)) else v)
                                              for k, v in info.items()}
        self._print_out(f"Buckling engine {eng.name}: {wall:.2f} s ("
                        + ", ".join(f"{k} {v:.2f}" for k, v in eng.t.items()) + ")", "Verbose")
        return res

def validate_buckling_sensitivity_regular(analysis_factory, rho_field: np.ndarray,
                                           test_element_indices=None, delta: float = 5e-3,
                                           num_modes: int = 6, aggregation_p: float = 8.0):
    """
    Finite-difference validation of the `Regular` class's buckling
    sensitivity -- Step 3 of the implementation plan. Perturbs the
    DENSITY field at a handful of elements, re-solves the full primal
    problem and buckling eigenproblem at rho+delta and rho-delta, and
    compares the central-difference estimate of d(mu_tilde)/d(rho_e) to
    the analytical adjoint sensitivity from a single baseline
    (unperturbed) run.

    Run this once per problem type / mesh configuration to build
    confidence in the sensitivity before trusting it in production, and
    again after any change to the buckling code in this module. It is
    deliberately NOT wired into the optimisation loop.

    Symmetric cross-sections are supported and expected to pass. A square
    column produces exactly-degenerate eigenvalue pairs (it is equally
    happy buckling about either principal axis), but the p-norm aggregate
    remains smooth through exact multiplicity because it depends on a
    degenerate cluster only through the TRACE of the sensitivity over its
    subspace, and a trace does not care which orthonormal basis ARPACK
    happened to return. Measured worst-case error on a 3x3x16 square
    column is ~0.017%, against ~0.003% on an asymmetric 2x5x16 section --
    the same order.

    The one degeneracy-related situation that IS genuinely ill-defined is
    a `num_modes` cut that lands part-way through a degenerate cluster,
    since a partial trace over an arbitrary half-basis is not
    basis-independent. `run_buckling` detects this with a probe mode and
    warns; the fix is always to increase `num_modes`. If this check
    reports errors of a few percent that do not shrink as delta shrinks,
    look for that warning in the log first.

    `Conformal`'s and `ConjugateGradient`'s sensitivities are not checked
    by this function (Conformal needs a full remesh per perturbation;
    ConjugateGradient's host-numpy assembly is exercised identically to
    Regular's mathematics, just with a different array layout) -- use
    this function's structure as the template if a dedicated check for
    either is needed.

    Args:
    - analysis_factory (callable): takes a single `iter_name` keyword
      argument and returns a FRESH `Regular` instance with everything
      else (youngs_modulus, poisson_ratio, csg_domain_data, save_meshes=
      False, ...) already bound, e.g.
          functools.partial(Regular, printouts="Silent",
                             youngs_modulus=210e9, poisson_ratio=0.3,
                             csg_domain_data=my_csg_data, save_meshes=False)
      A new instance is required per perturbation because `Regular`
      accumulates mesh/material state across calls.
    - rho_field (np.ndarray): the (small!) density field to perturb.
      Keep this test problem small -- each of the ~2*len(test_element_
      indices)+1 evaluations is a full primal + buckling solve.
    - test_element_indices (list[int] | None): flat (Fortran-order)
      element indices to test. Defaults to 5 random active (rho >= 0.5)
      elements.
    - delta (float): rho perturbation magnitude for the central difference.
    - num_modes, aggregation_p: as in `run_buckling`.

    Returns:
    - dict[int, tuple[float, float, float]] mapping element index to
      (analytical, finite_difference, relative_error).
    """
    baseline = analysis_factory(iter_name="test")
    baseline.run_analysis(rho_field, None)
    baseline_result = baseline.run_buckling(num_modes=num_modes, aggregation_p=aggregation_p)

    if baseline_result.mu_values.size == 0:
        raise RuntimeError(
            "Baseline run found no valid buckling modes; choose a more "
            "heavily (compressively) loaded test problem."
        )

    if test_element_indices is None:
        rng = np.random.default_rng(0)
        candidates = np.where(rho_field.ravel(order='F') >= 0.5)[0]
        if candidates.size == 0:
            raise RuntimeError("No active (rho >= 0.5) elements in rho_field to test.")
        test_element_indices = rng.choice(
            candidates, size=min(5, candidates.size), replace=False
        )

    analytical_sensitivity_flat = baseline_result.sensitivity.flatten(order='F')

    results = {}
    for e in test_element_indices:
        e = int(e)
        flat_plus = rho_field.flatten(order='F').copy()
        flat_minus = rho_field.flatten(order='F').copy()
        flat_plus[e] = np.clip(flat_plus[e] + delta, 0.0, 1.0)
        flat_minus[e] = np.clip(flat_minus[e] - delta, 0.0, 1.0)
        rho_plus = flat_plus.reshape(rho_field.shape, order='F')
        rho_minus = flat_minus.reshape(rho_field.shape, order='F')

        comp_plus = analysis_factory(iter_name="test")
        comp_plus.run_analysis(rho_plus, None)
        result_plus = comp_plus.run_buckling(num_modes=num_modes, aggregation_p=aggregation_p)

        comp_minus = analysis_factory(iter_name="test")
        comp_minus.run_analysis(rho_minus, None)
        result_minus = comp_minus.run_buckling(num_modes=num_modes, aggregation_p=aggregation_p)

        actual_delta = (rho_plus.flatten(order='F')[e] - rho_minus.flatten(order='F')[e]) / 2.0
        fd_estimate = (result_plus.mu_tilde - result_minus.mu_tilde) / (2.0 * actual_delta)
        analytical = float(analytical_sensitivity_flat[e])

        denom = max(abs(analytical), abs(fd_estimate), 1e-12)
        rel_error = abs(analytical - fd_estimate) / denom

        results[e] = (analytical, fd_estimate, rel_error)

        print(
            f"Element {e}: analytical={analytical:.6e}  "
            f"finite-difference={fd_estimate:.6e}  relative error={rel_error:.4f}"
        )

    return results


# ======================================================================================================================
# === former buckling_fast.py: fast linear-buckling engines (cuDSS / GMG-LOBPCG / host Cholesky) ===
# ======================================================================================================================
"""
Fast linear-buckling engines for the ConjugateGradient backend (docs/PLAN_engine_buckling.md, parts B, C, E3).

Same formulation as `analysis.ConjugateGradient.run_buckling` (the "classic" engine): the largest k = m + 1
values of mu = 1/lambda of the symmetric-definite pencil

    G phi = mu K phi,   G = -K_sigma(r(rho) sigma(u)),   K SPD on the free DOFs,

followed by the spurious-mode filter, p-norm aggregation and the adjoint sensitivity. Only the linear algebra
is replaced (ported from bench/buckling, verified there against run_buckling to 1e-12):

  assembly   structured-grid node-stencil assembly of K and G in CSR order (numba on the CPU, a CuPy RawKernel
             on the GPU), or matrix-free K x / G x products on the GPU;
  K^-1       cuDSS sparse Cholesky through nvmath-python (float64 "cudss64", float32 "cudss32", optionally in
             hybrid host/device memory), the PCLS geometric-multigrid V-cycle ("gmg_lobpcg32"), or a host
             Cholesky (MKL PARDISO mtype 2, CHOLMOD, SuperLU as the last resort);
  eigen      block LOBPCG for the largest mu (preconditioner = the K^-1 above, residuals in float64), warm-started
             from the previous evaluation's whole block, with an eigsh fallback when it stalls; or ARPACK eigsh
             with an exact K^-1 ("cpu_cholesky_eigsh", the accuracy reference);
  adjoint    one solve K nu = g with the same factor (float64), or PCG with the float64 operator preconditioned
             by the float32 factor / the V-cycle.

Tension-only designs (part C): LOBPCG returns the largest mu directly, so a design without a compressive mode
shows mu_max <= 0. When LOBPCG did not converge, or mu_max is within noise of zero, the number of mu > sigma is
certified by Sylvester's law of inertia of G - sigma K (cuDSS symmetric-indefinite LDL^T inertia when it fits
on the device, else MKL PARDISO mtype -2), exactly as `Analysis._count_modes_above`; None = unknown.

Only numpy and scipy are imported at module level. CuPy, nvmath-python, numba, scikit-sparse and pypardiso are
imported inside the engines that need them, so the module imports on any machine (Windows, aarch64, CPU-only).
"""
import gc
import time

import numpy as np
import scipy.sparse as sp

# ============================================================================ element reference data
# Corner order n0..n7 and DOF order (ux, uy, uz per node) of the PCLS GPU kernels (gpu_analysis.hex_dof_map).
_DI = np.array([0, 1, 1, 0, 0, 1, 1, 0])
_DJ = np.array([0, 0, 1, 1, 0, 0, 1, 1])
_DK = np.array([0, 0, 0, 0, 1, 1, 1, 1])
_C_LAM = np.zeros((6, 6)); _C_LAM[:3, :3] = 1.0
_C_MU = np.diag([2.0, 2.0, 2.0, 1.0, 1.0, 1.0])
_VOIGT_PAIRS = [[(0, 0)], [(1, 1)], [(2, 2)], [(1, 2), (2, 1)], [(0, 2), (2, 0)], [(0, 1), (1, 0)]]
_OFFS = np.array([[(_DI[b] - _DI[a] + 1) * 9 + (_DJ[b] - _DJ[a] + 1) * 3 + (_DK[b] - _DK[a] + 1)
                   for b in range(8)] for a in range(8)], dtype=np.int32)

# Engine names (see gpu_analysis.py for the selection and the memory model).
GPU_DIRECT = ("cudss64", "cudss32", "cudss64_hybrid", "cudss32_hybrid", "cudss64_eigsh")
FAST_ENGINES = GPU_DIRECT + ("gmg_lobpcg32", "cpu_cholesky_lobpcg", "cpu_cholesky_eigsh")

# LOBPCG settings (plan B): block = k + guards; iteration caps after which the eigsh fallback runs.
LOBPCG_GUARDS = 3
LOBPCG_MAXITER_DIRECT = 150
LOBPCG_MAXITER_GMG = 400
EIGSH_FALLBACK_TOL = 1e-6
ADJOINT_PCG_TOL = 1e-10          # relative residual of the PCG adjoint / exact K^-1 (float32 factor, GMG)
INERTIA_SHIFT = 1e-10            # = Analysis._INERTIA_SHIFT
ELEM_CHUNK = 1 << 15             # elements per chunk in the element-level steps (bounds their temporaries)
NEAR_ZERO_REL = 1e-8             # |mu_max| <= this * max|theta| counts as "within noise of 0"


def hex_dof_map(nx, ny, nz):
    """(Ne, 24) global DOFs, C-order nodes (z fastest), elements in C order -- gpu_analysis.hex_dof_map."""
    ei, ej, ek = np.meshgrid(np.arange(nx - 1), np.arange(ny - 1), np.arange(nz - 1), indexing="ij")
    ei, ej, ek = ei.ravel(), ej.ravel(), ek.ravel()
    nodes = np.stack([(ei + _DI[a]) * (ny * nz) + (ej + _DJ[a]) * nz + (ek + _DK[a]) for a in range(8)], axis=1)
    return (3 * nodes[:, :, None] + np.arange(3)[None, None, :]).reshape(nodes.shape[0], 24).astype(np.int64)


# ============================================================================ grid and CSR patterns
class Grid:
    """Structured hex grid + Dirichlet mask: reference element operators and (lazily) the CSR patterns.

    B_all, Bg_all, detJ are passed in from gpu_analysis.hex_reference_data so the element algebra is the classic
    path's own."""

    def __init__(self, nx, ny, nz, h, fixed_mask, B_all, Bg_all, detJ):
        self.nx, self.ny, self.nz = int(nx), int(ny), int(nz)
        self.Nn = self.nx * self.ny * self.nz
        self.Nd = 3 * self.Nn
        self.Ne = (self.nx - 1) * (self.ny - 1) * (self.nz - 1)
        self.h = tuple(float(v) for v in h)
        self.fixed = np.asarray(fixed_mask).astype(bool).ravel()
        if self.fixed.size != self.Nd:
            raise ValueError("fixed_mask size does not match the grid")
        self.free = np.flatnonzero(~self.fixed)
        self.nfree = int(self.free.size)
        self.free_index = np.full(self.Nd, -1, dtype=np.int64)
        self.free_index[self.free] = np.arange(self.nfree)
        self.dof_map = hex_dof_map(self.nx, self.ny, self.nz)
        self.B = np.asarray(B_all, np.float64)
        self.Bg = np.asarray(Bg_all, np.float64)
        self.detJ = float(detJ)
        d = self.detJ
        self.Klam = d * np.einsum("gvc,vw,gwd->cd", self.B, _C_LAM, self.B)
        self.Kmu = d * np.einsum("gvc,vw,gwd->cd", self.B, _C_MU, self.B)
        # stress operators: sigma_voigt(e, gp, v) = lam_e (DL u_e) + mu_e (DM u_e)
        self.DL = np.einsum("vw,gwc->gvc", _C_LAM, self.B).reshape(48, 24)
        self.DM = np.einsum("vw,gwc->gvc", _C_MU, self.B).reshape(48, 24)
        # geometric stiffness: S_e[a, b] = sum_{gp, v} sigma_voigt(e, gp, v) T6[gp, v, a, b]
        T6 = np.zeros((8, 6, 8, 8))
        for v, pairs in enumerate(_VOIGT_PAIRS):
            for (i, j) in pairs:
                T6[:, v] += d * np.einsum("ga,gb->gab", self.Bg[:, i, :], self.Bg[:, j, :])
        self.T6 = T6.reshape(48, 64)
        self._pat = {}

    def key(self):
        return (self.nx, self.ny, self.nz, self.h, hash(self.fixed.tobytes()))

    def pattern(self, kind):
        """CSR pattern (indptr, indices, gidx) of the free x free matrix; data = storage.ravel()[gidx].
        kind 'K': storage A[n, o, c, d] (Nn, 27, 3, 3); kind 'G': diagonal blocks only, S[n, o] (Nn, 27).
        Entries are generated in (n, c, o, d) order, which is already CSR order (no sort)."""
        if kind not in self._pat:
            try:
                self._pat[kind] = _pattern_numba(self, kind)
            except ImportError:
                self._pat[kind] = _pattern_numpy(self, kind)
        return self._pat[kind]


_NUMBA = {}


def _pattern_numba(g, kind):
    import numba  # noqa: F401
    from numba import njit, prange
    if "pattern" not in _NUMBA:
        @njit(parallel=True, cache=True)
        def count(nx, ny, nz, fi, diag, counts):
            Nn = nx * ny * nz
            for n in prange(Nn):
                i = n // (ny * nz); j = (n // nz) % ny; k = n % nz
                for c in range(3):
                    if fi[3 * n + c] < 0:
                        continue
                    cnt = 0
                    for o in range(27):
                        di = o // 9 - 1; dj = (o // 3) % 3 - 1; dk = o % 3 - 1
                        if i + di < 0 or i + di >= nx or j + dj < 0 or j + dj >= ny or k + dk < 0 or k + dk >= nz:
                            continue
                        m = n + di * ny * nz + dj * nz + dk
                        if diag:
                            if fi[3 * m + c] >= 0:
                                cnt += 1
                        else:
                            for d in range(3):
                                if fi[3 * m + d] >= 0:
                                    cnt += 1
                    counts[fi[3 * n + c]] = cnt

        @njit(parallel=True, cache=True)
        def fill(nx, ny, nz, fi, diag, indptr, cols, gidx):
            Nn = nx * ny * nz
            for n in prange(Nn):
                i = n // (ny * nz); j = (n // nz) % ny; k = n % nz
                for c in range(3):
                    r = fi[3 * n + c]
                    if r < 0:
                        continue
                    t = indptr[r]
                    for o in range(27):
                        di = o // 9 - 1; dj = (o // 3) % 3 - 1; dk = o % 3 - 1
                        if i + di < 0 or i + di >= nx or j + dj < 0 or j + dj >= ny or k + dk < 0 or k + dk >= nz:
                            continue
                        m = n + di * ny * nz + dj * nz + dk
                        if diag:
                            q = fi[3 * m + c]
                            if q >= 0:
                                cols[t] = q; gidx[t] = n * 27 + o; t += 1
                        else:
                            for d in range(3):
                                q = fi[3 * m + d]
                                if q >= 0:
                                    cols[t] = q; gidx[t] = n * 243 + o * 9 + c * 3 + d; t += 1
        _NUMBA["pattern"] = (count, fill)
    count, fill = _NUMBA["pattern"]
    diag = kind == "G"
    counts = np.zeros(g.nfree, dtype=np.int64)
    count(g.nx, g.ny, g.nz, g.free_index, diag, counts)
    indptr = np.zeros(g.nfree + 1, dtype=np.int64)
    np.cumsum(counts, out=indptr[1:])
    cols = np.empty(indptr[-1], dtype=np.int32)
    gidx = np.empty(indptr[-1], dtype=np.int64)
    fill(g.nx, g.ny, g.nz, g.free_index, diag, indptr, cols, gidx)
    return indptr, cols, gidx


def _pattern_numpy(g, kind):
    nx, ny, nz, Nn = g.nx, g.ny, g.nz, g.Nn
    n = np.arange(Nn, dtype=np.int64)
    i, j, k = n // (ny * nz), (n // nz) % ny, n % nz
    o = np.arange(27)
    di, dj, dk = o // 9 - 1, (o // 3) % 3 - 1, o % 3 - 1
    ok = ((i[:, None] + di >= 0) & (i[:, None] + di < nx) & (j[:, None] + dj >= 0) & (j[:, None] + dj < ny)
          & (k[:, None] + dk >= 0) & (k[:, None] + dk < nz))
    m = n[:, None] + di * ny * nz + dj * nz + dk
    fi = g.free_index
    if kind == "K":
        c = np.arange(3)[None, :, None, None]; dd = np.arange(3)[None, None, None, :]
        row = 3 * n[:, None, None, None] + c
        col = 3 * np.where(ok, m, 0)[:, None, :, None] + dd
        valid = ok[:, None, :, None] & (fi[row] >= 0) & (fi[col] >= 0)
        gidx = np.broadcast_to(n[:, None, None, None] * 243 + o[None, None, :, None] * 9 + c * 3 + dd, valid.shape)[valid]
        cols = np.broadcast_to(fi[col], valid.shape)[valid]
    else:
        c = np.arange(3)[None, :, None]
        row = 3 * n[:, None, None]
        col = 3 * np.where(ok, m, 0)[:, None, :] + c
        valid = ok[:, None, :] & (fi[row + c] >= 0) & (fi[col] >= 0)
        gidx = np.broadcast_to(n[:, None, None] * 27 + o[None, None, :], valid.shape)[valid]
        cols = fi[col][valid]
    counts = valid.reshape(Nn * 3, -1).sum(axis=1)[fi >= 0]
    indptr = np.zeros(g.nfree + 1, dtype=np.int64)
    np.cumsum(counts, out=indptr[1:])
    return indptr, cols.astype(np.int32), gidx.astype(np.int64)


# ============================================================================ element data
def _to(xp, a):
    return a if xp is np else xp.asarray(a)


class ElementData:
    """Per-design element quantities on numpy or cupy, built from the classic path's own inputs.

    K_e = ka_e KA + kb_e KB: (lam_e, mu_e, Klam, Kmu) for q1_gauss, (E_e, 0, Khat, 0) for q1 / q1_im with
    E_e = 2 mu_e (1 + nu) -- exactly ConjugateGradient._cg_assemble_stiffness. Stress, filter energy, the
    sensitivity terms and the adjoint load use the Gauss (compatible) operators, as the classic path does."""

    def __init__(self, xp, grid, E_elem, nu_elem, u, rho_flat, threshold, lam_s, mu_s, element_type, nu_run, Khat):
        from gpu_analysis import lame_from_E_nu
        self.xp = xp
        g = grid
        lam_e, mu_e = lame_from_E_nu(np.asarray(E_elem).astype(np.float64), np.asarray(nu_elem).astype(np.float64))
        self.lam_s, self.mu_s = float(lam_s), float(mu_s)
        if element_type == "q1_gauss":
            ka, kb, KA, KB = lam_e, mu_e, g.Klam, g.Kmu
        else:
            ka, kb, KA, KB = 2.0 * mu_e * (1.0 + float(nu_run)), np.zeros_like(mu_e), np.asarray(Khat, np.float64), np.zeros((24, 24))
        self.ka, self.kb = _to(xp, ka), _to(xp, kb)
        self.KA, self.KB = _to(xp, np.ascontiguousarray(KA)), _to(xp, np.ascontiguousarray(KB))
        thr = float(threshold)
        den = max(2.0 * thr - thr, 1e-12)
        x = np.clip((np.asarray(rho_flat).astype(np.float64) - thr) / den, 0.0, 1.0)
        self.lam, self.mu = _to(xp, lam_e), _to(xp, mu_e)
        self.ramp = _to(xp, x * x * (3.0 - 2.0 * x))
        self.rampp = _to(xp, (6.0 * x * (1.0 - x)) / den)
        self.dof_map = _to(xp, g.dof_map)
        self.DL, self.DM, self.T6 = _to(xp, g.DL), _to(xp, g.DM), _to(xp, g.T6)
        self.Klam, self.Kmu, self.Bg = _to(xp, g.Klam), _to(xp, g.Kmu), _to(xp, g.Bg)
        self.K0 = self.lam_s * self.Klam + self.mu_s * self.Kmu          # solid element stiffness (Gauss)
        uu = _to(xp, np.asarray(u, np.float64))
        self.Ue = uu[self.dof_map]                                          # (Ne, 24)
        del uu
        Ne = g.Ne
        self.sig_act = xp.empty((Ne, 48))                                    # actual stress
        self.sig_prime = xp.empty((Ne, 48))                                  # d sigma / d rho (C linear in rho)
        self.S_ramped = xp.empty((Ne, 64))                                   # K_sigma element blocks
        for c0 in range(0, Ne, ELEM_CHUNK):
            c = slice(c0, min(Ne, c0 + ELEM_CHUNK))
            eL = self.Ue[c] @ self.DL.T
            eM = self.Ue[c] @ self.DM.T
            self.sig_act[c] = self.lam[c, None] * eL + self.mu[c, None] * eM
            self.sig_prime[c] = self.lam_s * eL + self.mu_s * eM
            self.S_ramped[c] = (self.sig_act[c] * self.ramp[c, None]) @ self.T6
            del eL, eM


# ============================================================================ assembly
def _numba_kernels():
    if "stencil_K" in _NUMBA:
        return _NUMBA
    from numba import njit, prange

    @njit(parallel=True, cache=True)
    def stencil_K(nx, ny, nz, ka, kb, KA, KB, offs, di, dj, dk, out):
        Nn = nx * ny * nz
        for n in prange(Nn):
            i = n // (ny * nz); j = (n // nz) % ny; k = n % nz
            for a in range(8):
                ei = i - di[a]; ej = j - dj[a]; ek = k - dk[a]
                if ei < 0 or ej < 0 or ek < 0 or ei >= nx - 1 or ej >= ny - 1 or ek >= nz - 1:
                    continue
                e = ei * (ny - 1) * (nz - 1) + ej * (nz - 1) + ek
                l = ka[e]; m = kb[e]
                for b in range(8):
                    o = offs[a, b]
                    for c in range(3):
                        for d in range(3):
                            out[n, o, c, d] += l * KA[3 * a + c, 3 * b + d] + m * KB[3 * a + c, 3 * b + d]

    @njit(parallel=True, cache=True)
    def stencil_G(nx, ny, nz, S, offs, di, dj, dk, out):
        Nn = nx * ny * nz
        for n in prange(Nn):
            i = n // (ny * nz); j = (n // nz) % ny; k = n % nz
            for a in range(8):
                ei = i - di[a]; ej = j - dj[a]; ek = k - dk[a]
                if ei < 0 or ej < 0 or ek < 0 or ei >= nx - 1 or ej >= ny - 1 or ek >= nz - 1:
                    continue
                e = ei * (ny - 1) * (nz - 1) + ej * (nz - 1) + ek
                for b in range(8):
                    out[n, offs[a, b]] += S[e, 8 * a + b]

    @njit(parallel=True, cache=True)
    def gather(src, idx, out):
        for t in prange(idx.size):
            out[t] = src[idx[t]]

    _NUMBA.update(stencil_K=stencil_K, stencil_G=stencil_G, gather=gather)
    return _NUMBA


def assemble_K_cpu(grid, ka, kb, KA, KB):
    """Free-DOF K = sum_e (ka_e KA + kb_e KB) by node-stencil assembly (numba CPU, no atomics) + CSR gather."""
    nb = _numba_kernels()
    g = grid
    AK = np.zeros((g.Nn, 27, 3, 3))
    nb["stencil_K"](g.nx, g.ny, g.nz, np.ascontiguousarray(ka, np.float64), np.ascontiguousarray(kb, np.float64),
                    np.ascontiguousarray(KA, np.float64), np.ascontiguousarray(KB, np.float64), _OFFS, _DI, _DJ, _DK, AK)
    ipK, inK, gK = g.pattern("K")
    dK = np.empty(gK.size); nb["gather"](AK.ravel(), gK, dK)
    del AK
    return sp.csr_matrix((dK, inK, ipK), shape=(g.nfree, g.nfree))


def grid_for(nx, ny, nz, h, fixed_mask):
    """The process-wide Grid (CSR patterns cached) for these dimensions and Dirichlet mask."""
    return _state_for((nx, ny, nz, h, fixed_mask))["grid"]


def assemble_cpu(grid, ed, want_G=True):
    """Node-stencil assembly of K and G on the CPU (numba, one thread per node, no atomics) + CSR gather."""
    nb = _numba_kernels()
    g = grid
    K_II = assemble_K_cpu(g, ed.ka, ed.kb, ed.KA, ed.KB)
    G_II = None
    if want_G:
        AG = np.zeros((g.Nn, 27))
        nb["stencil_G"](g.nx, g.ny, g.nz, -ed.S_ramped, _OFFS, _DI, _DJ, _DK, AG)    # G = -K_sigma
        ipG, inG, gG = g.pattern("G")
        dG = np.empty(gG.size); nb["gather"](AG.ravel(), gG, dG)
        G_II = sp.csr_matrix((dG, inG, ipG), shape=(g.nfree, g.nfree))
    return K_II, G_II


_CUDA_SRC = r"""
__constant__ int c_DI[8] = {0,1,1,0,0,1,1,0};
__constant__ int c_DJ[8] = {0,0,1,1,0,0,1,1};
__constant__ int c_DK[8] = {0,0,0,0,1,1,1,1};

extern "C" __global__ void stencil_K(const int nx, const int ny, const int nz,
        const double* __restrict__ ka, const double* __restrict__ kb,
        const double* __restrict__ KA, const double* __restrict__ KB, double* __restrict__ out) {
    long n = (long)blockIdx.x * blockDim.x + threadIdx.x;
    if (n >= (long)nx * ny * nz) return;
    int i = n / (ny * nz), j = (n / nz) % ny, k = n % nz;
    for (int a = 0; a < 8; ++a) {
        int ei = i - c_DI[a], ej = j - c_DJ[a], ek = k - c_DK[a];
        if (ei < 0 || ej < 0 || ek < 0 || ei >= nx - 1 || ej >= ny - 1 || ek >= nz - 1) continue;
        long e = ((long)ei * (ny - 1) + ej) * (nz - 1) + ek;
        double l = ka[e], m = kb[e];
        for (int b = 0; b < 8; ++b) {
            int o = (c_DI[b] - c_DI[a] + 1) * 9 + (c_DJ[b] - c_DJ[a] + 1) * 3 + (c_DK[b] - c_DK[a] + 1);
            for (int c = 0; c < 3; ++c)
                for (int d = 0; d < 3; ++d)
                    out[n * 243 + o * 9 + c * 3 + d] += l * KA[(3*a+c)*24 + 3*b+d] + m * KB[(3*a+c)*24 + 3*b+d];
        }
    }
}

extern "C" __global__ void stencil_G(const int nx, const int ny, const int nz,
        const double* __restrict__ S, const double sgn, double* __restrict__ out) {
    long n = (long)blockIdx.x * blockDim.x + threadIdx.x;
    if (n >= (long)nx * ny * nz) return;
    int i = n / (ny * nz), j = (n / nz) % ny, k = n % nz;
    for (int a = 0; a < 8; ++a) {
        int ei = i - c_DI[a], ej = j - c_DJ[a], ek = k - c_DK[a];
        if (ei < 0 || ej < 0 || ek < 0 || ei >= nx - 1 || ej >= ny - 1 || ek >= nz - 1) continue;
        long e = ((long)ei * (ny - 1) + ej) * (nz - 1) + ek;
        for (int b = 0; b < 8; ++b) {
            int o = (c_DI[b] - c_DI[a] + 1) * 9 + (c_DJ[b] - c_DJ[a] + 1) * 3 + (c_DK[b] - c_DK[a] + 1);
            out[n * 27 + o] += sgn * S[e * 64 + 8 * a + b];
        }
    }
}

// Matrix-free y = K x (mode 0) or y = s * K_sigma x (mode 1) for nv vectors stored row-major (Nd, nv).
// One thread per (node, vector); gathers from the <= 8 elements of its node: no atomics.
// x must vanish at fixed DOFs; y is zeroed there.
extern "C" __global__ void apply_op(const int nx, const int ny, const int nz, const int nv, const int mode,
        const double* __restrict__ ka, const double* __restrict__ kb,
        const double* __restrict__ KA, const double* __restrict__ KB,
        const double* __restrict__ S, const double s,
        const signed char* __restrict__ fixed, const double* __restrict__ x, double* __restrict__ y) {
    __shared__ double sK[2 * 576];
    if (mode == 0) {
        for (int t = threadIdx.x; t < 576; t += blockDim.x) { sK[t] = KA[t]; sK[576 + t] = KB[t]; }
    }
    __syncthreads();
    long t = (long)blockIdx.x * blockDim.x + threadIdx.x;
    long Nn = (long)nx * ny * nz;
    if (t >= Nn * nv) return;
    long n = t / nv; int v = t % nv;
    int i = n / (ny * nz), j = (n / nz) % ny, k = n % nz;
    double y0 = 0.0, y1 = 0.0, y2 = 0.0;
    for (int a = 0; a < 8; ++a) {
        int ei = i - c_DI[a], ej = j - c_DJ[a], ek = k - c_DK[a];
        if (ei < 0 || ej < 0 || ek < 0 || ei >= nx - 1 || ej >= ny - 1 || ek >= nz - 1) continue;
        long e = ((long)ei * (ny - 1) + ej) * (nz - 1) + ek;
        long base = ((long)ei * ny + ej) * nz + ek;     // node of corner 0
        if (mode == 0) {
            double l = ka[e], m = kb[e];
            for (int b = 0; b < 8; ++b) {
                long mb = base + ((long)c_DI[b] * ny + c_DJ[b]) * nz + c_DK[b];
                double x0 = x[(3*mb) * nv + v], x1 = x[(3*mb+1) * nv + v], x2 = x[(3*mb+2) * nv + v];
                const double* KL = sK + (3*a) * 24 + 3*b;
                const double* KM = sK + 576 + (3*a) * 24 + 3*b;
                y0 += l * (KL[0]*x0 + KL[1]*x1 + KL[2]*x2) + m * (KM[0]*x0 + KM[1]*x1 + KM[2]*x2);
                y1 += l * (KL[24]*x0 + KL[25]*x1 + KL[26]*x2) + m * (KM[24]*x0 + KM[25]*x1 + KM[26]*x2);
                y2 += l * (KL[48]*x0 + KL[49]*x1 + KL[50]*x2) + m * (KM[48]*x0 + KM[49]*x1 + KM[50]*x2);
            }
        } else {
            for (int b = 0; b < 8; ++b) {
                long mb = base + ((long)c_DI[b] * ny + c_DJ[b]) * nz + c_DK[b];
                double sab = s * S[e * 64 + 8 * a + b];
                y0 += sab * x[(3*mb) * nv + v]; y1 += sab * x[(3*mb+1) * nv + v]; y2 += sab * x[(3*mb+2) * nv + v];
            }
        }
    }
    y[(3*n) * nv + v]     = fixed[3*n]     ? 0.0 : y0;
    y[(3*n+1) * nv + v]   = fixed[3*n+1]   ? 0.0 : y1;
    y[(3*n+2) * nv + v]   = fixed[3*n+2]   ? 0.0 : y2;
}
"""
_CUPY_MOD = {}


def ensure_cuda_headers():
    """CuPy compiles the kernels above with NVRTC, which needs the CUDA headers. pip wheels bring them with
    `cupy-cuda12x[ctk]`; a conda env has them under <prefix>/targets/<arch>/include (and <prefix>/Library/include on
    Windows). Point CUDA_PATH there when it is unset and the default lookup would fail."""
    import glob
    import os
    import sys
    if os.environ.get("CUDA_PATH"):
        return
    for pre in {sys.prefix, os.environ.get("CONDA_PREFIX") or sys.prefix}:
        for cand in glob.glob(os.path.join(pre, "targets", "*")) + [os.path.join(pre, "Library"), pre]:
            if os.path.exists(os.path.join(cand, "include", "cuda_runtime.h")):
                os.environ["CUDA_PATH"] = cand
                return


def _cuda_kernels():
    if not _CUPY_MOD:
        ensure_cuda_headers()
        import cupy as cp
        mod = cp.RawModule(code=_CUDA_SRC, options=("-std=c++14",))
        mod.compile()
        _CUPY_MOD.update(stencil_K=mod.get_function("stencil_K"), stencil_G=mod.get_function("stencil_G"),
                         apply_op=mod.get_function("apply_op"))
    return _CUPY_MOD


class GPUGrid:
    """Device copies of the grid data used by the GPU paths (kept for the whole run)."""

    def __init__(self, grid):
        import cupy as cp
        self.g = grid
        self.fixed = cp.asarray(grid.fixed.astype(np.int8))
        self.free = cp.asarray(grid.free)
        self._pat = {}

    def pattern(self, kind):
        import cupy as cp
        if kind not in self._pat:
            ip, ind, gi = self.g.pattern(kind)
            self._pat[kind] = (cp.asarray(ip.astype(np.int32)), cp.asarray(ind), cp.asarray(gi))
        return self._pat[kind]


def assemble_gpu(gg, ed, kind="K"):
    """Stencil assembly on the GPU + CSR gather: cupyx CSR K_II (kind 'K') or G_II (kind 'G'), float64."""
    import cupy as cp
    import cupyx.scipy.sparse as csp
    kk = _cuda_kernels()
    g = gg.g
    th = 128
    bl = (g.Nn + th - 1) // th
    if kind == "K":
        A = cp.zeros((g.Nn, 27, 3, 3))
        kk["stencil_K"]((bl,), (th,), (np.int32(g.nx), np.int32(g.ny), np.int32(g.nz), ed.ka, ed.kb, ed.KA, ed.KB, A))
    else:
        A = cp.zeros((g.Nn, 27))
        kk["stencil_G"]((bl,), (th,), (np.int32(g.nx), np.int32(g.ny), np.int32(g.nz), ed.S_ramped,
                                       np.float64(-1.0), A))
    ip, ind, gi = gg.pattern(kind)
    M = csp.csr_matrix((A.ravel()[gi], ind, ip), shape=(g.nfree, g.nfree))
    del A
    return M


class MatrixFreeOps:
    """y = K x, y = G x on free-DOF blocks (n_free, nv) via the matrix-free CUDA kernel (float64)."""

    def __init__(self, gg, ed):
        self.gg, self.ed = gg, ed
        self.k = _cuda_kernels()["apply_op"]
        self.nK = 0
        self.nG = 0

    def _apply(self, X, mode):
        import cupy as cp
        g = self.gg.g
        one = X.ndim == 1
        X2 = X[:, None] if one else X
        nv = X2.shape[1]
        full = cp.zeros((g.Nd, nv))
        full[self.gg.free] = X2
        y = cp.empty_like(full)
        th = 128
        tot = g.Nn * nv
        ed = self.ed
        self.k(((tot + th - 1) // th,), (th,),
               (np.int32(g.nx), np.int32(g.ny), np.int32(g.nz), np.int32(nv), np.int32(mode),
                ed.ka, ed.kb, ed.KA, ed.KB, ed.S_ramped, np.float64(-1.0), self.gg.fixed, full, y))
        del full
        out = y[self.gg.free]
        return out[:, 0] if one else out

    def K(self, X):
        self.nK += 1 if X.ndim == 1 else X.shape[1]
        return self._apply(X, 0)

    def G(self, X):
        self.nG += 1 if X.ndim == 1 else X.shape[1]
        return self._apply(X, 1)


# ============================================================================ K^-1 backends
class HostCholesky:
    """Host sparse Cholesky: MKL PARDISO mtype 2 (x86 / Windows), else CHOLMOD (scikit-sparse), else SuperLU.
    A fresh factorisation each evaluation (scikit-sparse's numeric refactorisation ran single-threaded and
    took > 15 min on the 0.67 M DOF column; a fresh cho_factor took ~20 s)."""

    def __init__(self, prefer=None):
        self.kind = prefer or host_cholesky_kind()
        if self.kind is None:
            raise ImportError("no host sparse direct solver")
        self._f = None

    def factor(self, K):
        self.free()
        t0 = time.perf_counter()
        if self.kind == "pardiso":
            from pypardiso import PyPardisoSolver
            A = sp.triu(K, format="csr"); A.sort_indices()
            self.A = A
            self._f = PyPardisoSolver(mtype=2)
            self._f.factorize(A)
        elif self.kind == "cholmod":
            from sksparse import cholmod
            Kc = sp.csc_matrix(K)
            if hasattr(cholmod, "cho_factor"):
                self._f = cholmod.cho_factor(Kc)
                self._solve = self._f.solve
            else:
                self._f = cholmod.cholesky(Kc)
                self._solve = self._f
        else:
            import scipy.sparse.linalg as spla
            self._f = spla.splu(sp.csc_matrix(K))
        self.factor_s = time.perf_counter() - t0
        return self

    def solve(self, b):
        b = np.asarray(b, dtype=np.float64)
        if self.kind == "pardiso":
            return self._f.solve(self.A, b)
        if self.kind == "cholmod":
            return self._solve(b)
        return self._f.solve(b)

    def free(self):
        if self._f is not None and self.kind == "pardiso":
            try:
                self._f.free_memory(everything=True)
            except Exception:
                pass
        self._f = None
        gc.collect()


def host_cholesky_kind():
    """'pardiso' | 'cholmod' | 'splu' | None, in that order of preference."""
    try:
        import pypardiso  # noqa: F401
        return "pardiso"
    except Exception:
        pass
    try:
        from sksparse import cholmod  # noqa: F401
        return "cholmod"
    except Exception:
        pass
    return "splu"


class CuDSSSolver:
    """cuDSS sparse Cholesky (SPD) through nvmath-python, float64 or float32, optionally hybrid memory.
    The plan (reordering + symbolic analysis) is made once per run while the sparsity pattern is unchanged; each
    evaluation then copies the new values into the held operand (keeps the plan) and refactorises. nvmath binds
    the solver to one right-hand-side shape, so a fixed (n, nrhs) column-major buffer is used."""

    def __init__(self, dtype=np.float64, nrhs=1, hybrid_limit_bytes=None):
        self.dtype = np.dtype(dtype)
        self.nrhs = int(nrhs)
        self.hybrid_limit_bytes = hybrid_limit_bytes
        self.ds = None
        self._key = None
        self.plan_s = 0.0

    def factor(self, K64):
        """K64: cupyx CSR float64 (consumed: the float32 variant keeps only its own cast)."""
        import cupy as cp
        import cupyx.scipy.sparse as csp
        import nvmath
        from nvmath.sparse.advanced import DirectSolver, DirectSolverOptions
        n = K64.shape[0]
        key = (n, int(K64.nnz), self.nrhs, self.dtype.str)
        cp.cuda.Device().synchronize()
        t0 = time.perf_counter()
        if self.ds is None or self._key != key:
            self.free()
            Kd = K64 if self.dtype == np.float64 else csp.csr_matrix(
                (K64.data.astype(self.dtype), K64.indices, K64.indptr), shape=K64.shape)
            self._b = cp.zeros((n, self.nrhs), dtype=self.dtype, order="F")
            opts = DirectSolverOptions(sparse_system_type=nvmath.sparse.advanced.DirectSolverMatrixType.SPD)
            execution = None
            if self.hybrid_limit_bytes:
                from nvmath.sparse.advanced import ExecutionCUDA, HybridMemoryModeOptions
                execution = ExecutionCUDA(hybrid_memory_mode_options=HybridMemoryModeOptions(
                    hybrid_memory_mode=True, hybrid_device_memory_limit=int(self.hybrid_limit_bytes)))
            self.K = Kd
            self.ds = DirectSolver(self.K, self._b, options=opts, execution=execution)
            self.ds.plan()
            self._key = key
            cp.cuda.Device().synchronize()
            self.plan_s = time.perf_counter() - t0
        else:
            self.K.data[...] = K64.data                  # same pattern: update values in place (keeps the plan)
            self.plan_s = 0.0
        t1 = time.perf_counter()
        self.ds.factorize()
        cp.cuda.Device().synchronize()
        self.factor_s = time.perf_counter() - t1
        return self

    def solve_gpu(self, B):
        import cupy as cp
        one = B.ndim == 1
        B2 = B[:, None] if one else B
        k = B2.shape[1]
        out = cp.empty((B2.shape[0], k), dtype=np.float64)
        for j0 in range(0, k, self.nrhs):
            j1 = min(k, j0 + self.nrhs)
            # unit-max columns (keeps a float32 right-hand side away from under/overflow; linear, so exact)
            sc = cp.abs(B2[:, j0:j1]).max(axis=0)
            sc = cp.where(sc > 0, sc, 1.0)
            self._b[:, : j1 - j0] = B2[:, j0:j1] / sc[None, :]
            if j1 - j0 < self.nrhs:
                self._b[:, j1 - j0:] = 0.0
            X = self.ds.solve()
            out[:, j0:j1] = X[:, : j1 - j0].astype(np.float64) * sc[None, :]
        if not bool(cp.all(cp.isfinite(out))):
            raise FloatingPointError(f"cuDSS {self.dtype.name} solve produced non-finite values")
        return out[:, 0] if one else out

    def free(self):
        if self.ds is not None:
            try:
                self.ds.free()
            except Exception:
                pass
        self.ds = None
        self._key = None
        self.K = None


def cudss_inertia(A_gpu):
    """(n_positive, n_negative) of a symmetric (indefinite) cupyx CSR float64 matrix by a cuDSS LDL^T, or None.
    nvmath's DirectSolverFactorizationInfo.inertia = [positive, negative] pivots (checked by
    tests/test_engines.py on a matrix with known inertia)."""
    import cupy as cp
    import nvmath
    from nvmath.sparse.advanced import DirectSolver, DirectSolverOptions
    b = cp.zeros((A_gpu.shape[0], 1), order="F")
    opts = DirectSolverOptions(sparse_system_type=nvmath.sparse.advanced.DirectSolverMatrixType.SYMMETRIC)
    ds = DirectSolver(A_gpu, b, options=opts)
    try:
        ds.plan()
        info = ds.factorize()
        inert = np.asarray(info.inertia).astype(np.int64)
        return int(inert[0]), int(inert[1])
    finally:
        ds.free()


def pardiso_inertia(K_II, G_II, sigma):
    """Number of positive eigenvalues of G - sigma K by MKL PARDISO mtype -2 (Analysis._count_modes_above)."""
    from pypardiso import PyPardisoSolver
    A = (sp.csr_matrix(G_II, dtype=np.float64) - sigma * sp.csr_matrix(K_II, dtype=np.float64)).tocsr()
    upper = sp.triu(A, format="csr"); upper.sort_indices()
    s = PyPardisoSolver(mtype=-2)
    try:
        s._check_A(upper)
        s.set_phase(12)
        s._call_pardiso(upper, np.zeros((A.shape[0], 1)))
        return int(s.get_iparm(22))
    finally:
        try:
            s.free_memory(everything=True)
        except Exception:
            pass


class GMGWrapper:
    """The PCLS geometric-multigrid V-cycle (gpu_analysis.GMGPreconditioner, numba-CUDA, float32) applied column by
    column to free-DOF float64 CuPy blocks; the element matrix is the run's element_type."""

    def __init__(self, grid, E_elem, nu_elem, element_type, nu_run, levels=None, smoother_steps=2,
                 coarse_max_nodes=400):
        import cupy as cp
        from numba import cuda
        import gpu_analysis as gc_
        self.cp, self.cuda = cp, cuda
        g = grid
        nx, ny, nz = g.nx, g.ny, g.nz
        th = gc_.THREADS_PER_BLOCK
        E = np.asarray(E_elem, np.float32)
        nu_e = np.asarray(nu_elem, np.float32)
        etype = element_type if element_type != "q1_gauss" else "q1"   # V-cycle only: same element, matrix form
        self.d_E = cuda.to_device(E)
        self.d_Khat = gc_._khat_to_device(etype, *g.h, E, nu_e, nu_run)
        self.d_fixed = cuda.to_device(g.fixed.astype(np.int32))
        Nn, Ne = g.Nn, g.Ne
        be = (Ne + th - 1) // th; bn = (Nn + th - 1) // th
        d_diag = cuda.device_array(Nn * 9, dtype=np.float32)
        gc_._zero_array_kernel[(d_diag.size + th - 1) // th, th](d_diag)
        gc_._element_diag_kernel_mat[be, th](nx, ny, nz, self.d_E, self.d_Khat, d_diag)
        self.d_Minv = cuda.device_array(Nn * 9, dtype=np.float32)
        gc_._invert_diag_blocks_kernel[bn, th](Nn, d_diag, self.d_Minv, self.d_fixed)
        del d_diag
        d_E, d_Khat = self.d_E, self.d_Khat

        def applyK0(d_in, d_out):
            gc_._element_action_kernel_mat[be, th](nx, ny, nz, d_E, d_Khat, d_in, d_out)
        self.pre = gc_.GMGPreconditioner(nx, ny, nz, self.d_fixed, applyK0, self.d_Minv, d_E, d_Khat,
                                         levels=levels, smoother_steps=smoother_steps,
                                         coarse_max_nodes=coarse_max_nodes)
        self.r32 = cp.zeros(g.Nd, dtype=np.float32)
        self.z32 = cp.zeros(g.Nd, dtype=np.float32)
        self.free = cp.asarray(g.free)

    def apply_free(self, R):
        cp = self.cp
        one = R.ndim == 1
        R2 = R[:, None] if one else R
        out = cp.empty_like(R2)
        for j in range(R2.shape[1]):
            col = R2[:, j]
            s = float(cp.abs(col).max())
            if s == 0.0:
                out[:, j] = 0.0
                continue
            self.r32.fill(0)
            self.r32[self.free] = (col / s).astype(np.float32)
            self.pre.apply(self.cuda.as_cuda_array(self.r32), self.cuda.as_cuda_array(self.z32))
            out[:, j] = self.z32[self.free].astype(np.float64) * s
        return out[:, 0] if one else out


# ============================================================================ solvers
def pcg(xp, Kop, Mop, b, tol, maxiter=500):
    """Preconditioned CG (float64) for K x = b, relative residual tol; returns (x, iterations, rel_residual)."""
    x = xp.zeros_like(b)
    r = b.copy()
    nb = float(xp.linalg.norm(b))
    if nb == 0.0:
        return x, 0, 0.0
    z = Mop(r)
    p = z.copy()
    rz = float(r @ z)
    it = 0
    rel = 1.0
    for it in range(1, maxiter + 1):
        Ap = Kop(p)
        alpha = rz / float(p @ Ap)
        x += alpha * p
        r -= alpha * Ap
        rel = float(xp.linalg.norm(r)) / nb
        if rel <= tol:
            break
        z = Mop(r)
        rz_new = float(r @ z)
        p = z + (rz_new / rz) * p
        rz = rz_new
    return x, it, rel


def _svqb(xp, X, BX, drop=1e-12):
    """B-orthonormalise the columns of X (Stathopoulos-Wu SVQB); drops near-dependent directions."""
    Gm = X.T @ BX
    Gm = 0.5 * (Gm + Gm.T)
    d = xp.sqrt(xp.clip(xp.diag(Gm), 1e-300, None))
    w, V = xp.linalg.eigh(Gm / d[:, None] / d[None, :])
    keep = w > drop * float(w.max())
    V = V[:, keep]; w = w[keep]
    Z = (V / d[:, None]) / xp.sqrt(w)[None, :]
    return X @ Z, BX @ Z, Z


def _bmat(xp, B):
    """Dense block matrix from a nested list (cupy has no `block`)."""
    return xp.concatenate([xp.concatenate(row, axis=1) for row in B], axis=0)


def lobpcg_largest(xp, applyA, applyB, applyT, X0, k_conv, tol, maxiter=200, recompute_every=15):
    """Block LOBPCG (Knyazev 2001) for the largest eigenpairs of A x = mu B x, B SPD, A symmetric (the bench's
    verified iteration). Rayleigh-Ritz on span[X, W, P] with SVQB B-orthonormalisation and soft locking; the
    Gram matrices are formed block by block (no concatenated n x 3nb copies: memory-lean).
    Convergence: ||A x - mu B x||_2 / (|mu| ||B x||_2) < tol for the first k_conv columns.
    Returns (mu desc, X (B-orthonormal), iterations, final max rel. residual, converged)."""
    X = X0
    BX = applyB(X)
    X, BX, _ = _svqb(xp, X, BX)
    AX = applyA(X)
    nb = X.shape[1]
    P = AP = BP = None
    th = None
    last = np.inf
    it = 0
    for it in range(1, maxiter + 1):
        if P is None:
            GA = X.T @ AX; GA = 0.5 * (GA + GA.T)
            th, C = xp.linalg.eigh(GA)
            idx = xp.argsort(-th)
            th, C = th[idx], C[:, idx]
            X, AX, BX = X @ C, AX @ C, BX @ C
        R = AX - BX * th[None, :]
        rn = xp.linalg.norm(R, axis=0) / (xp.abs(th) * xp.linalg.norm(BX, axis=0) + 1e-300)
        rn_h = np.asarray(rn.get() if hasattr(rn, "get") else rn)
        last = float(rn_h[:k_conv].max())
        if np.all(rn_h[:k_conv] < tol):
            return th, X, it, last, True
        act = np.flatnonzero(rn_h >= tol)
        W = applyT(R[:, act]) if act.size < nb else applyT(R)
        del R
        BW = applyB(W)
        W = W - X @ (BX.T @ W)
        BW = BW - BX @ (X.T @ BW)
        W, BW, _ = _svqb(xp, W, BW)
        AW = applyA(W)
        blocks = [(X, AX, BX), (W, AW, BW)] + ([(P, AP, BP)] if P is not None else [])
        nbl = len(blocks)
        GA = [[None] * nbl for _ in range(nbl)]
        GB = [[None] * nbl for _ in range(nbl)]
        for i in range(nbl):
            for j in range(i, nbl):
                GA[i][j] = blocks[i][0].T @ blocks[j][1]
                GB[i][j] = blocks[i][0].T @ blocks[j][2]
                if j > i:
                    GA[j][i] = GA[i][j].T
                    GB[j][i] = GB[i][j].T
        GAm = _bmat(xp, GA); GAm = 0.5 * (GAm + GAm.T)
        GBm = _bmat(xp, GB); GBm = 0.5 * (GBm + GBm.T)
        d = xp.sqrt(xp.clip(xp.diag(GBm), 1e-300, None))
        w, V = xp.linalg.eigh(GBm / d[:, None] / d[None, :])
        keep = w > 1e-13 * float(w.max())
        Z = (V[:, keep] / d[:, None]) / xp.sqrt(w[keep])[None, :]
        th_all, Y = xp.linalg.eigh(Z.T @ GAm @ Z)
        idx = xp.argsort(-th_all)[:nb]
        th = th_all[idx]
        Cf = Z @ Y[:, idx]
        sizes = [b[0].shape[1] for b in blocks]
        offs = np.cumsum([0] + sizes)
        Cs = [Cf[offs[i]:offs[i + 1]] for i in range(nbl)]
        Pn = blocks[1][0] @ Cs[1]; APn = blocks[1][1] @ Cs[1]; BPn = blocks[1][2] @ Cs[1]
        if nbl == 3:
            Pn += P @ Cs[2]; APn += AP @ Cs[2]; BPn += BP @ Cs[2]
        del W, AW, BW
        X = X @ Cs[0] + Pn; AX = AX @ Cs[0] + APn; BX = BX @ Cs[0] + BPn
        P, AP, BP = Pn, APn, BPn
        del Pn, APn, BPn, blocks
        if recompute_every and it % recompute_every == 0:
            BX = applyB(X); X, BX, _ = _svqb(xp, X, BX); AX = applyA(X)
            P = AP = BP = None
        else:
            P, BP, Zp = _svqb(xp, P, BP)
            AP = AP @ Zp
    return th, X, it, last, False


def eigsh_largest(applyA, applyB, kinv, n, k, tol, v0=None):
    """ARPACK regular-inverse mode, eigsh(G, M=K, Minv=K^-1, which='LA') on host vectors (the classic call),
    with partial convergence unpacked. Returns (mu desc, modes (n, k) host, nothing_converged, n_Kinv)."""
    from scipy.sparse.linalg import eigsh, LinearOperator, ArpackNoConvergence, ArpackError
    cnt = [0]

    def mv(x):
        cnt[0] += 1
        return kinv(x)
    A = LinearOperator((n, n), matvec=applyA, dtype=np.float64)
    M = LinearOperator((n, n), matvec=applyB, dtype=np.float64)
    Minv = LinearOperator((n, n), matvec=mv, dtype=np.float64)
    kw = {} if v0 is None else {"v0": v0}
    try:
        mu, modes = eigsh(A, k=k, M=M, sigma=None, Minv=Minv, which="LA", tol=float(tol), **kw)
        nothing = False
    except ArpackNoConvergence as e:
        mu = e.eigenvalues if e.eigenvalues is not None else np.zeros(0)
        modes = e.eigenvectors if e.eigenvectors is not None else np.zeros((n, 0))
        nothing = mu.size == 0
    except ArpackError:                       # e.g. "starting vector is zero" when G = 0: nothing is known
        mu, modes, nothing = np.zeros(0), np.zeros((n, 0)), True
    order = np.argsort(mu)[::-1]
    return mu[order], modes[:, order], nothing, cnt[0]


# ============================================================================ the engine
# Process-wide state kept across evaluations (the analysis object is rebuilt every iteration): the grid with
# its CSR patterns, the device copies, the cuDSS plans, and the previous evaluation's LOBPCG block (warm start).
_STATE = {"key": None}


def reset_state():
    """Free the factorisations and forget the warm-start block (also called when the grid / DOF set changes)."""
    for s in list(_STATE.get("solvers", {}).values()):
        try:
            s.free()
        except Exception:
            pass
    _STATE.clear()
    _STATE["key"] = None
    try:
        import cupy as cp
        cp.get_default_memory_pool().free_all_blocks()
    except Exception:
        pass


def _state_for(grid_args):
    from gpu_analysis import hex_reference_data
    nx, ny, nz, h, fixed_mask = grid_args
    key = (nx, ny, nz, tuple(float(v) for v in h), hash(np.asarray(fixed_mask).astype(bool).tobytes()))
    if _STATE.get("key") != key:
        reset_state()
        B_all, Bg_all, detJ = hex_reference_data(*h)
        _STATE.update(key=key, grid=Grid(nx, ny, nz, h, fixed_mask, B_all, Bg_all, detJ), gg=None, solvers={},
                      prev_block=None, gmg=None)
    return _STATE


class EngineError(RuntimeError):
    pass


class BucklingEngine:
    """One fast buckling evaluation: `eigen()` (mu, K-normalised free-DOF modes), then the classic path's filter /
    aggregation in analysis.py, then `sensitivity()` for the retained modes. Device data live on this object for
    the evaluation; the factorisation / plan and the warm-start block persist in _STATE."""

    def __init__(self, name, nx, ny, nz, h, fixed_mask, E_elem, nu_elem, u_h, rho_flat, threshold, lam_s, mu_s,
                 element_type, nu_run, Khat, lobpcg_tol=1e-3, eigensolver_tol=1e-12, hybrid_limit_bytes=None,
                 gmg_kwargs=None, log=print):
        if name not in FAST_ENGINES:
            raise ValueError(f"unknown fast buckling engine {name!r}")
        self.name = name
        self.st = _state_for((nx, ny, nz, h, fixed_mask))
        self.g = self.st["grid"]
        self.gpu = name in GPU_DIRECT or name == "gmg_lobpcg32"
        self.lobpcg_tol = float(lobpcg_tol)
        self.eigensolver_tol = float(eigensolver_tol)
        self.hybrid_limit_bytes = hybrid_limit_bytes
        self.gmg_kwargs = dict(gmg_kwargs or {})
        self.log = log
        self.diag = {"engine": name}
        self.t = {}
        t0 = time.perf_counter()
        if self.gpu:
            import cupy as cp
            self.xp = cp
            if self.st.get("gg") is None:
                self.st["gg"] = GPUGrid(self.g)
            self.gg = self.st["gg"]
        else:
            self.xp = np
            self.gg = None
        self.ed = ElementData(self.xp, self.g, E_elem, nu_elem, u_h, rho_flat, threshold, lam_s, mu_s,
                              element_type, nu_run, Khat)
        self._E_elem, self._nu_elem, self._etype, self._nu_run = E_elem, nu_elem, element_type, nu_run
        self._u_h, self._rho_h, self._thr, self._Khat = u_h, rho_flat, threshold, Khat   # host inertia fallback
        self.K_cpu = self.G_cpu = None
        self.solver = None
        self.gmg = None
        self.ops = None
        self._sync()
        self.t["element_data"] = time.perf_counter() - t0

    # ------------------------------------------------------------------ helpers
    def _sync(self):
        if self.xp is not np:
            self.xp.cuda.Device().synchronize()

    def _host(self, a):
        return a.get() if hasattr(a, "get") else np.asarray(a)

    def _setup(self):
        """Assembly + factorisation / preconditioner set-up."""
        t0 = time.perf_counter()
        nm = self.name
        g = self.g
        if nm in GPU_DIRECT:
            K = assemble_gpu(self.gg, self.ed, "K")
            self._sync(); self.t["assembly"] = time.perf_counter() - t0; t0 = time.perf_counter()
            dtype = np.float32 if nm.startswith("cudss32") else np.float64
            skey = (nm, self.hybrid_limit_bytes if "hybrid" in nm else None)
            s = self.st["solvers"].get(skey)
            if s is None:
                for other in list(self.st["solvers"]):        # one factor on the device at a time
                    self.st["solvers"].pop(other).free()
                s = CuDSSSolver(dtype=dtype, nrhs=self.nb_block(), hybrid_limit_bytes=self.hybrid_limit_bytes
                                if "hybrid" in nm else None)
                self.st["solvers"][skey] = s
            s.factor(K)
            del K
            self.solver = s
            self.diag.update(plan_s=s.plan_s, factor_s=s.factor_s)
            self.ops = MatrixFreeOps(self.gg, self.ed)
        elif nm == "gmg_lobpcg32":
            self.t["assembly"] = 0.0
            self.gmg = GMGWrapper(g, self._E_elem, self._nu_elem, self._etype, self._nu_run, **self.gmg_kwargs)
            self.ops = MatrixFreeOps(self.gg, self.ed)
        else:
            self.K_cpu, self.G_cpu = assemble_cpu(g, self.ed, want_G=True)
            self.t["assembly"] = time.perf_counter() - t0; t0 = time.perf_counter()
            s = HostCholesky()
            s.factor(self.K_cpu)
            self.solver = s
            self.diag.update(host_solver=s.kind, factor_s=s.factor_s)
        self._sync()
        self.t["factor_setup"] = time.perf_counter() - t0

    def nb_block(self):
        return self._k + LOBPCG_GUARDS

    # operators on free-DOF blocks
    def applyK(self, X):
        return self.ops.K(X) if self.ops is not None else self.K_cpu @ X

    def applyG(self, X):
        return self.ops.G(X) if self.ops is not None else self.G_cpu @ X

    def applyT(self, R):
        """Preconditioner T ~ K^-1 for LOBPCG."""
        if self.gmg is not None:
            return self.gmg.apply_free(R)
        if self.xp is not np:
            return self.solver.solve_gpu(R)
        return self.solver.solve(R)

    def kinv_exact(self, b):
        """K^-1 b to working accuracy: the float64 factor, or PCG preconditioned by the float32 factor / V-cycle."""
        if self.name in ("cudss64", "cudss64_hybrid", "cudss64_eigsh"):
            return self.solver.solve_gpu(b)
        if self.xp is np:
            return self.solver.solve(b)
        x, it, rel = pcg(self.xp, self.applyK, self.applyT, b, ADJOINT_PCG_TOL)
        self.diag["pcg_iters"] = self.diag.get("pcg_iters", 0) + it
        if rel > ADJOINT_PCG_TOL * 10:
            raise EngineError(f"{self.name}: PCG for K^-1 stalled at relative residual {rel:.2e}")
        return x

    # ------------------------------------------------------------------ eigen
    def eigen(self, num_modes):
        """Returns (mu desc (host), modes_free (xp, K-normalised, n x k), converged, n_compressive_cert).
        converged False = nothing is known (NOT_CONVERGED). n_compressive_cert: inertia count when it ran."""
        xp = self.xp
        n = self.g.nfree
        self._k = k = min(num_modes + 1, max(n - 2, 1))
        self._setup()
        t0 = time.perf_counter()
        cert = None
        if self.name in ("cpu_cholesky_eigsh", "cudss64_eigsh"):
            # reference engines: the classic ARPACK call (eigensolver_tol) with an exact K^-1
            if self.xp is np:
                fA, fB, fK = self.applyG, self.applyK, self.solver.solve
            else:
                fA = lambda x: xp.asnumpy(self.applyG(xp.asarray(x)))          # noqa: E731
                fB = lambda x: xp.asnumpy(self.applyK(xp.asarray(x)))          # noqa: E731
                fK = lambda x: xp.asnumpy(self.solver.solve_gpu(xp.asarray(x)))  # noqa: E731
            mu, modes, nothing, nKinv = eigsh_largest(fA, fB, fK, n, k, self.eigensolver_tol)
            if xp is not np:
                modes = xp.asarray(modes)
            self.diag.update(eig="eigsh", n_Kinv=nKinv)
            converged = not nothing
            if nothing:
                cert = self.inertia_count()
                self.diag["inertia"] = cert
                converged = cert == 0
                mu, modes = np.zeros(0), np.zeros((n, 0))
            self.t["eigensolve"] = time.perf_counter() - t0
            return mu, modes, converged, cert
        nb = self.nb_block()
        rng = np.random.default_rng(0)
        X0 = xp.asarray(rng.standard_normal((n, nb)))
        prev = self.st.get("prev_block")
        warm = prev is not None and prev.shape[0] == n
        if warm:
            pv = xp.asarray(prev)
            m = min(nb, pv.shape[1])
            X0[:, :m] = pv[:, :m]
        maxiter = LOBPCG_MAXITER_GMG if self.gmg is not None else LOBPCG_MAXITER_DIRECT
        th, X, iters, relres, conv = lobpcg_largest(xp, self.applyG, self.applyK, self.applyT, X0, k,
                                                    self.lobpcg_tol, maxiter=maxiter)
        del X0
        self.diag.update(eig="lobpcg", lobpcg_iters=iters, lobpcg_relres=relres, lobpcg_converged=conv, warm=warm)
        mu = self._host(th[:k]).astype(np.float64)
        modes = X[:, :k]
        if conv:
            self.st["prev_block"] = X                     # the whole block (modes + probe + guards) for the next start
        scale = float(np.max(np.abs(self._host(th)))) if th.size else 0.0
        near_zero = mu.size == 0 or abs(float(mu[0])) <= NEAR_ZERO_REL * max(scale, 1e-300)
        if not conv:
            # stalled: eigsh with an exact K^-1 at a moderate tolerance (plan B)
            self.log(f"Buckling ({self.name}): LOBPCG did not reach tol {self.lobpcg_tol:g} in {iters} iterations "
                     f"(residual {relres:.2e}); falling back to eigsh at tol {EIGSH_FALLBACK_TOL:g}.")
            v0 = self._host(X[:, :k]).sum(axis=1)
            if xp is np:
                fA, fB, fK = self.applyG, self.applyK, self.kinv_exact
            else:
                fA = lambda x: xp.asnumpy(self.applyG(xp.asarray(x)))          # noqa: E731
                fB = lambda x: xp.asnumpy(self.applyK(xp.asarray(x)))          # noqa: E731
                fK = lambda x: xp.asnumpy(self.kinv_exact(xp.asarray(x)))      # noqa: E731
            mu, modes_h, nothing, nKinv = eigsh_largest(fA, fB, fK, n, k, EIGSH_FALLBACK_TOL, v0=v0)
            self.diag.update(eig="lobpcg->eigsh", n_Kinv=nKinv)
            modes = xp.asarray(modes_h) if xp is not np else modes_h
            if nothing:
                cert = self.inertia_count()
                self.diag["inertia"] = cert
                self.t["eigensolve"] = time.perf_counter() - t0
                return np.zeros(0), modes[:, :0], cert == 0, cert
            near_zero = abs(float(mu[0])) <= NEAR_ZERO_REL * max(float(np.max(np.abs(mu))), 1e-300)
        if near_zero:
            # mu_max within noise of 0: certify the sign by the inertia count (None = keep the eigensolver's answer)
            cert = self.inertia_count()
            self.diag["inertia"] = cert
            if cert == 0:
                mu = np.minimum(mu, 0.0)                    # certified: no mu > sigma
        self.t["eigensolve"] = time.perf_counter() - t0
        return mu, modes, True, cert

    def inertia_count(self):
        """Number of mu > INERTIA_SHIFT by Sylvester's law (part C), or None when no LDL^T is available."""
        t0 = time.perf_counter()
        try:
            if self.xp is not np:
                try:
                    import cupy as cp
                    K = assemble_gpu(self.gg, self.ed, "K")
                    G = assemble_gpu(self.gg, self.ed, "G")
                    A = (G - INERTIA_SHIFT * K).tocsr()
                    del K, G
                    A.sum_duplicates()
                    pos, _neg = cudss_inertia(A)
                    del A
                    cp.get_default_memory_pool().free_all_blocks()
                    self.diag["inertia_by"] = "cudss"
                    return pos
                except Exception as e:  # noqa: BLE001 -- e.g. out of device memory: try the host
                    self.diag["inertia_gpu_error"] = f"{type(e).__name__}: {e}"[:200]
                ed_h = ElementData(np, self.g, self._E_elem, self._nu_elem, self._u_h, self._rho_h, self._thr,
                                   self.ed.lam_s, self.ed.mu_s, self._etype, self._nu_run, self._Khat)
                K, G = assemble_cpu(self.g, ed_h)
            else:
                K, G = self.K_cpu, self.G_cpu
            n = pardiso_inertia(K, G, INERTIA_SHIFT)
            self.diag["inertia_by"] = "pardiso"
            return n
        except Exception as e:  # noqa: BLE001
            self.diag["inertia_error"] = f"{type(e).__name__}: {e}"[:200]
            return None
        finally:
            self.t["inertia"] = time.perf_counter() - t0

    # ------------------------------------------------------------------ filter input
    def energy_fraction(self, modes_free):
        """Ramp-weighted share of each mode's element strain energy (classic _modal_energy_fraction_in_solid)."""
        xp, ed, g = self.xp, self.ed, self.g
        free = _to(xp, g.free)
        k = modes_free.shape[1]
        frac = np.zeros(k)
        full = xp.zeros(g.Nd)
        for j in range(k):
            full[free] = modes_free[:, j]
            tot = 0.0
            wsum = 0.0
            for c0 in range(0, g.Ne, ELEM_CHUNK):
                c = slice(c0, min(g.Ne, c0 + ELEM_CHUNK))
                Pe = full[ed.dof_map[c]]
                en = ed.lam[c] * xp.einsum("ec,ec->e", Pe @ ed.Klam, Pe) + ed.mu[c] * xp.einsum("ec,ec->e", Pe @ ed.Kmu, Pe)
                en = xp.clip(en, 0.0, None)
                tot += float(en.sum())
                wsum += float((en * ed.ramp[c]).sum())
            frac[j] = wsum / tot if tot > 0 else 0.0
        return frac

    def modes_full_host(self, modes_free):
        g = self.g
        out = np.zeros((g.Nd, modes_free.shape[1]))
        out[g.free, :] = self._host(modes_free)
        return out

    # ------------------------------------------------------------------ sensitivity
    def sensitivity(self, mu_k, modes_free, weights):
        """d mu_tilde / d rho_e (Ne,) host: explicit terms A, B, C + one adjoint solve (classic steps 5-6)."""
        t0 = time.perf_counter()
        xp, ed, g = self.xp, self.ed, self.g
        Ne = g.Ne
        d = g.detJ
        free = _to(xp, g.free)
        expl = xp.zeros(Ne)
        gam_comb = xp.zeros((Ne, 8, 6))
        sigp = ed.sig_prime.reshape(Ne, 8, 6)
        siga = ed.sig_act.reshape(Ne, 8, 6)
        full = xp.zeros(g.Nd)
        for j in range(len(mu_k)):
            full[free] = modes_free[:, j]
            wj = float(weights[j])
            for c0 in range(0, Ne, ELEM_CHUNK):
                c = slice(c0, min(Ne, c0 + ELEM_CHUNK))
                Pe = full[ed.dof_map[c]]
                nc = Pe.shape[0]
                term_a = -float(mu_k[j]) * xp.einsum("ec,ec->e", Pe @ ed.K0, Pe)
                gp = xp.einsum("gia,eak->egik", ed.Bg, Pe.reshape(nc, 8, 3))    # d phi_k / d x_i
                G00 = (gp[:, :, 0, :] ** 2).sum(-1); G11 = (gp[:, :, 1, :] ** 2).sum(-1); G22 = (gp[:, :, 2, :] ** 2).sum(-1)
                G12 = (gp[:, :, 1, :] * gp[:, :, 2, :]).sum(-1); G02 = (gp[:, :, 0, :] * gp[:, :, 2, :]).sum(-1)
                G01 = (gp[:, :, 0, :] * gp[:, :, 1, :]).sum(-1)
                del gp
                gam = xp.stack([G00, G11, G22, 2 * G12, 2 * G02, 2 * G01], axis=-1)   # (nc, 8, 6) unramped
                del G00, G11, G22, G12, G02, G01
                term_b = -d * ed.ramp[c] * xp.einsum("egv,egv->e", sigp[c], gam)
                term_c = -d * ed.rampp[c] * xp.einsum("egv,egv->e", siga[c], gam)
                expl[c] += wj * (term_a + term_b + term_c)
                gam_comb[c] += wj * ed.ramp[c, None, None] * gam
                del gam, Pe
        DLr = ed.DL.reshape(8, 6, 24); DMr = ed.DM.reshape(8, 6, 24)
        gfull = xp.zeros(g.Nd)
        for c0 in range(0, Ne, ELEM_CHUNK):
            c = slice(c0, min(Ne, c0 + ELEM_CHUNK))
            gl = d * (ed.lam[c, None] * xp.einsum("gva,egv->ea", DLr, gam_comb[c])
                      + ed.mu[c, None] * xp.einsum("gva,egv->ea", DMr, gam_comb[c]))
            if xp is np:
                np.add.at(gfull, g.dof_map[c].ravel(), gl.ravel())
            else:
                import cupyx
                cupyx.scatter_add(gfull, ed.dof_map[c].ravel(), gl.ravel())
            del gl
        del gam_comb
        self.t["explicit"] = time.perf_counter() - t0
        t0 = time.perf_counter()
        nu_free = self.kinv_exact(gfull[free])
        self._sync()
        self.t["adjoint"] = time.perf_counter() - t0
        nu_full = xp.zeros(g.Nd); nu_full[free] = nu_free
        for c0 in range(0, Ne, ELEM_CHUNK):
            c = slice(c0, min(Ne, c0 + ELEM_CHUNK))
            expl[c] += xp.einsum("ec,ec->e", nu_full[ed.dof_map[c]] @ ed.K0, ed.Ue[c])
        return self._host(expl).astype(np.float64)

    def release(self):
        """Drop this evaluation's device data (the factor and warm block persist in _STATE)."""
        had_gmg = self.gmg is not None
        self.ed = self.ops = self.gmg = self.K_cpu = self.G_cpu = None
        if self.xp is not np:
            try:
                self.xp.get_default_memory_pool().free_all_blocks()
            except Exception:
                pass
        if had_gmg:                        # numba defers frees; release the V-cycle hierarchy now
            try:
                gc.collect()
                from numba import cuda
                cuda.current_context().deallocations.clear()
            except Exception:
                pass
