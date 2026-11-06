"""Console entry point for the packaged Bonds Lite controller."""

from __future__ import annotations

import argparse
import multiprocessing
import sys


def _startup_message(arguments: list[str]) -> str:
    """Identify the requested stage before importing the workflow."""
    parser = argparse.ArgumentParser(add_help=False)
    parser.add_argument("--caller")
    parser.add_argument("--stage", choices=("1", "2", "3", "4"), default="1")
    parser.add_argument("--self-test", action="store_true")
    parser.add_argument("--install-button")
    parser.add_argument("--install-stage-three")
    parser.add_argument("--install-stage-four")
    parser.add_argument("-h", "--help", action="store_true")
    options, _ = parser.parse_known_args(arguments)
    if not options.caller or any(
        (
            options.self_test,
            options.install_button,
            options.install_stage_three,
            options.install_stage_four,
            options.help,
        )
    ):
        return "Starting Bonds Lite controller..."
    return f"Starting Stage {options.stage}..."


print(_startup_message(sys.argv[1:]), flush=True)

from bonds_lite_input import main


if __name__ == "__main__":
    multiprocessing.freeze_support()
    raise SystemExit(main())
