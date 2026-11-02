"""Frozen Python 3.11 entry point for the unchanged Bonds Lite batch model."""

from __future__ import annotations

import argparse
import csv
import json
import multiprocessing
import os
import sys
import traceback
from pathlib import Path


def preflight() -> None:
    """Check runtime dependencies without first-run numerical compilation."""
    print("Loading numerical libraries...", flush=True)
    import dateutil
    import numpy as np
    import pandas as pd
    import scipy
    import xlwings as xw

    print("Loading Bonds Lite model modules...", flush=True)
    import pyBondsHMutils as utils
    from BondsModelHandler import BondsModelHandler

    print(json.dumps({
        "status": "MODEL_PREFLIGHT_OK", "python": sys.version.split()[0],
        "numpy": np.__version__, "pandas": pd.__version__,
        "scipy": scipy.__version__, "xlwings": xw.__version__,
        "numba": utils.numba is not None,
    }), flush=True)


def self_test(excel: bool = False) -> None:
    """Check bundled numerical kernels and optionally a disposable Excel session."""
    import dateutil
    import numpy as np
    import pandas as pd
    import scipy
    import xlwings as xw

    import pyBondsHMutils as utils
    from BondsModelHandler import BondsModelHandler

    output = np.empty((1, 1, 1, 1), dtype=np.float64)
    result = utils.discount_factors_kernel(
        np.zeros((1, 1, 1)), np.zeros(1), np.ones(1), output
    )
    np.testing.assert_allclose(result, np.ones_like(output))
    spreads = utils.solve_zspreads_batch(
        np.array([0.0, 1.0]),
        np.array([[-100.0, 105.0]]),
        np.array([0.0, 0.05]),
    )
    if not np.isfinite(spreads).all():
        raise RuntimeError("Bundled spread solver self-test failed")
    if excel:
        application = xw.App(visible=False, add_book=True)
        try:
            book = application.books[0]
            book.sheets[0].range("A1").value = 42
            if book.sheets[0].range("A1").value != 42:
                raise RuntimeError("Bundled Excel COM read/write self-test failed")
            book.close()
        finally:
            application.api.DisplayAlerts = False
            application.api.Quit()
    print(json.dumps({
        "status": "MODEL_SELF_TEST_OK", "python": sys.version.split()[0],
        "numpy": np.__version__, "pandas": pd.__version__,
        "scipy": scipy.__version__, "xlwings": xw.__version__,
        "numba": utils.numba is not None, "excel": excel,
    }), flush=True)


def main(arguments: list[str] | None = None) -> int:
    """Run a bounded self-test or the folder batch without a Python CLI dependency."""
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("folder", type=Path, nargs="?")
    parser.add_argument("--self-test", action="store_true")
    parser.add_argument("--excel-self-test", action="store_true")
    parser.add_argument("--preflight", action="store_true")
    options = parser.parse_args(arguments)
    cache = Path(os.environ.get("LOCALAPPDATA", str(Path.home()))) / "BondsLite" / "numba-cache"
    cache.mkdir(parents=True, exist_ok=True)
    os.environ["NUMBA_CACHE_DIR"] = str(cache)
    try:
        if options.preflight:
            preflight()
            return 0
        if options.self_test or options.excel_self_test:
            self_test(excel=options.excel_self_test)
            return 0
        if options.folder is None:
            parser.error("Specify the staged input folder or --self-test")
        folder = options.folder.resolve(strict=True)
        expected = {"Bonds Lite input_Bonds.xlsx", "Bonds Lite input_MAP.xlsx"}
        actual = {
            path.name for path in folder.glob("*.xlsx")
            if not path.name.startswith("~$")
        }
        if actual != expected:
            raise ValueError("Packaged model requires exactly the Bonds and MAP input pair")
        from run_bonds_batch import run_batch

        summary = run_batch(folder)
        with Path(summary).open(encoding="utf-8-sig", newline="") as stream:
            rows = list(csv.DictReader(stream))
        if len(rows) != 2 or {row["input_file"] for row in rows} != expected:
            raise ValueError("Batch did not complete both datasets")
        if any(row["status"] != "success" or row["error"].strip() for row in rows):
            raise ValueError("Batch summary reports a failed model input")
        return 0
    except Exception:
        traceback.print_exc()
        return 1


if __name__ == "__main__":
    multiprocessing.freeze_support()
    raise SystemExit(main())
