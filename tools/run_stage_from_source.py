"""Run an SST stage from this project's source against the open Draft.

Maintainer helper for testing code changes without building or promoting a
packaged release. It finds the Draft open in desktop Excel, then runs the
same command-line entry point that the packaged launcher runs, using the
working-tree ``src`` folder.

Example::

    python tools\\run_stage_from_source.py P1
"""

from __future__ import annotations

import argparse
import os
import subprocess
import sys
from pathlib import Path

PROJECT = Path(__file__).resolve().parents[1]
DEFAULT_DRAFT = PROJECT / "Draft Input spreadsheet for SST.xlsm"


def find_open_draft(draft: Path) -> tuple[str, int]:
    """Return the open Draft's Excel ``FullName`` and its Excel window handle."""
    import xlwings as xw

    matches = [
        (book.api.FullName, int(application.api.Hwnd))
        for application in xw.apps
        for book in application.books
        if book.name.casefold() == draft.name.casefold()
    ]
    if len(matches) != 1:
        raise SystemExit(
            f"Open exactly one copy of {draft.name} in desktop Excel "
            f"(found {len(matches)})"
        )
    return matches[0]


def main() -> int:
    """Locate the open Draft and run the requested stage from source."""
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("stage", choices=("1", "2", "3", "4", "P1", "P2", "P3", "P4"))
    parser.add_argument("--draft", type=Path, default=DEFAULT_DRAFT)
    arguments = parser.parse_args()
    draft = arguments.draft.expanduser().resolve(strict=True)
    full_name, hwnd = find_open_draft(draft)
    environment = os.environ.copy()
    environment["PYTHONPATH"] = os.pathsep.join(
        filter(None, [str(PROJECT / "src"), environment.get("PYTHONPATH")])
    )
    command = [
        sys.executable, "-u", "-m", "bonds_lite_input",
        "--caller", full_name,
        "--caller-local", str(draft),
        "--excel-hwnd", str(hwnd),
        "--stage", arguments.stage,
    ]
    return subprocess.run(command, env=environment, check=False).returncode


if __name__ == "__main__":
    raise SystemExit(main())
