"""Guard and run MACalc Batch on exactly one Stage 3 workbook."""

from __future__ import annotations

import argparse
import hashlib
import importlib
import os
import subprocess
import sys
import traceback
from pathlib import Path

PROVENANCE_FILES = (
    "MACalcrunmanager.py",
    "MACalc.py",
    "MACE.py",
    "calcSCR.py",
    "calcmp.py",
    "MACalcutils.py",
)


def _normalised(path: str | Path) -> str:
    return os.path.normcase(str(Path(path).resolve()))


def _model_provenance(model_dir: Path) -> None:
    """Print the working-copy identity and hashes before changing a workbook."""
    result = subprocess.run(
        ["git", "rev-parse", "HEAD"],
        cwd=str(model_dir),
        capture_output=True,
        text=True,
        errors="replace",
        shell=False,
    )
    if result.returncode:
        head = f"unavailable ({result.stderr.strip()})"
    else:
        head = result.stdout.strip()
    print(f"MACalc Batch git HEAD: {head}", flush=True)

    status = subprocess.run(
        ["git", "status", "--porcelain"],
        cwd=str(model_dir),
        capture_output=True,
        text=True,
        errors="replace",
        shell=False,
    )
    if status.returncode:
        modified = f"unknown ({status.stderr.strip()})"
    else:
        modified = "yes" if status.stdout.strip() else "no"
    print(f"MACalc Batch working copy modified: {modified}", flush=True)
    for filename in PROVENANCE_FILES:
        path = model_dir / filename
        if not path.is_file():
            print(
                f"MACalc Batch SHA-256 {filename}: unavailable (file missing)",
                flush=True,
            )
            continue
        digest = hashlib.sha256(path.read_bytes()).hexdigest()
        print(f"MACalc Batch SHA-256 {filename}: {digest}", flush=True)


def _configure_console_streams() -> None:
    """Line-buffer console streams when the runtime supports reconfiguration."""
    for stream in (sys.stdout, sys.stderr):
        reconfigure = getattr(stream, "reconfigure", None)
        if callable(reconfigure):
            reconfigure(line_buffering=True)


def run(
    model_dir: Path,
    root: Path,
    expected: Path,
    downgrade_run: bool,
) -> int:
    """Require the run manager's complete workbook list to equal the audit."""
    model_dir = model_dir.resolve(strict=True)
    root = root.resolve(strict=True)
    expected = expected.resolve()
    frozen = getattr(sys, "frozen", False)
    if not frozen and not (model_dir / "MACalcrunmanager.py").is_file():
        raise ValueError(f"MACalc Batch entry point is missing: {model_dir}")
    if expected.parent != root:
        raise ValueError(
            f"Expected workbook is not directly inside the audited folder: {expected}"
        )

    sys.path.insert(0, str(model_dir))
    MACalcrunmanager = sys.modules.get("MACalcrunmanager")
    if not callable(getattr(MACalcrunmanager, "find_workbooks", None)):
        sys.modules.pop("MACalcrunmanager", None)
        MACalcrunmanager = importlib.import_module("MACalcrunmanager")
    workbooks = MACalcrunmanager.find_workbooks(root)
    print("MACalc Batch workbook scan:", flush=True)
    for workbook in workbooks:
        print(f"  {workbook}", flush=True)
    if len(workbooks) != 1 or _normalised(workbooks[0]) != _normalised(expected):
        print(
            "Stage 4 stopped: MACalc Batch must find exactly the audited Stage 3 workbook.",
            flush=True,
        )
        return 2
    if frozen:
        print("MACalc Batch packaged provenance: release.json", flush=True)
    else:
        _model_provenance(model_dir)

    expected_normalised = _normalised(expected)

    def confirm_workbooks(found: list[Path]) -> bool:
        print("MACalc Batch confirmation list:", flush=True)
        for workbook in found:
            print(f"  {workbook}", flush=True)
        return len(found) == 1 and _normalised(found[0]) == expected_normalised

    MACalcrunmanager.confirm_workbooks = confirm_workbooks
    try:
        results = MACalcrunmanager.main(
            root_repository=root,
            use_binary_optimisation=False,
            run_optimisation=False,
            is_downgrade_run=downgrade_run,
            floor_cod_at_zero=True,
        )
    except SystemExit as error:
        if error.code == 1:
            traceback.print_exc()
            raise
        print(f"MACalc Batch stopped with exit code {error.code}", flush=True)
        return int(error.code or 2)

    if not isinstance(results, list) or len(results) != 1:
        print("MACalc Batch did not process exactly one workbook.", flush=True)
        return 2
    result = results[0]
    if not isinstance(result, (tuple, list)) or len(result) < 3:
        print(f"MACalc Batch returned an invalid result: {result!r}", flush=True)
        return 2
    workbook = Path(result[0])
    csv_path = Path(result[2])
    if _normalised(workbook) != expected_normalised:
        print(f"MACalc Batch processed an unexpected workbook: {workbook}", flush=True)
        return 2
    if not csv_path.is_file() or csv_path.stat().st_size == 0:
        print(f"MACalc Batch results CSV is missing or empty: {csv_path}", flush=True)
        return 2
    print(f"Processed 1 workbook(s). Results: {csv_path}", flush=True)
    return 0


def main(arguments: list[str] | None = None) -> int:
    """Run a packaged self-test or the guarded Stage 4 model batch."""
    _configure_console_streams()
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--model-dir", type=Path)
    parser.add_argument("--root", type=Path)
    parser.add_argument("--expected", type=Path)
    parser.add_argument(
        "--downgrade-run",
        action=argparse.BooleanOptionalAction,
        default=False,
    )
    parser.add_argument("--self-test", action="store_true")
    options = parser.parse_args(arguments)

    model_dir = options.model_dir
    if getattr(sys, "frozen", False):
        model_dir = Path(
            getattr(sys, "_MEIPASS", Path(sys.executable).resolve().parent)
        )
    if model_dir is None:
        parser.error("--model-dir is required")
    original_directory = Path.cwd()
    try:
        root = options.root.resolve() if options.root is not None else None
        expected = options.expected.resolve() if options.expected is not None else None
        model_dir = model_dir.resolve(strict=True)
        sys.path.insert(0, str(model_dir))
        os.chdir(model_dir)
        if options.self_test:
            import MACalcrunmanager  # noqa: F401
            import pulp

            if not pulp.PULP_CBC_CMD().available():
                raise RuntimeError("Bundled MACalc CBC solver is unavailable")

            print("MACALC_SELF_TEST_OK", flush=True)
            return 0
        if root is None or expected is None:
            parser.error("--root and --expected are required")
        return run(model_dir, root, expected, options.downgrade_run)
    except SystemExit:
        raise
    except Exception:
        traceback.print_exc()
        return 1
    finally:
        os.chdir(original_directory)


if __name__ == "__main__":
    raise SystemExit(main())
