"""Console entry point for the packaged Bonds Lite controller."""

import multiprocessing
print("Starting Bonds Lite controller...", flush=True)

from bonds_lite_input import main


if __name__ == "__main__":
    multiprocessing.freeze_support()
    raise SystemExit(main())