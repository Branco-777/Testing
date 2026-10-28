"""Run only against a generated manifest in an isolated pytest run folder."""

from __future__ import annotations

import csv
import json
import os
import sys
from pathlib import Path


def main() -> int:
    """Emit flushed progress and controlled results without Excel or model imports."""
    folder = Path(sys.argv[1]).resolve(strict=True)
    manifest = json.loads((folder / "runner_case.json").read_text(encoding="utf-8"))
    print("STAGE_TWO_STDOUT", flush=True)
    print("STAGE_TWO_STDERR", file=sys.stderr, flush=True)
    Path(manifest["observed_path"]).write_text(
        json.dumps({"argv": sys.argv[1:], "cwd": str(Path.cwd()), "pid": os.getpid()}),
        encoding="utf-8",
    )
    allowed_folders = {
        Path(output_folder).resolve()
        for output_folder in manifest["output_folders"].values()
    }
    for dataset, names in manifest["artifacts"].items():
        output_folder = Path(manifest["output_folders"][dataset]).resolve()
        if output_folder not in allowed_folders:
            raise ValueError("Fixture outputs must use a saved expected output folder")
        for name in names:
            if Path(name).name != name:
                raise ValueError("Fixture artefact names must stay within the output folder")
            (output_folder / name).write_bytes(b"Controlled fresh result\n")
    rows = manifest["rows"]
    with (folder / "bonds_batch_summary_20261028_090000.csv").open(
        "w", encoding="utf-8", newline="",
    ) as stream:
        writer = csv.DictWriter(stream, fieldnames=(
            "input_file", "status", "seconds", "output_folder_path", "output_suffix", "error",
        ))
        writer.writeheader()
        writer.writerows(rows)
    print("STAGE_TWO_PARTIAL", end="", flush=True)
    return int(manifest["exit_code"])


if __name__ == "__main__":
    raise SystemExit(main())