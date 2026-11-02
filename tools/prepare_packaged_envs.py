"""Prepare pinned, isolated environments for the controller and model builds."""

from __future__ import annotations

import argparse
import os
import subprocess
import sys
import winreg
from pathlib import Path


def current_model_python() -> Path:
    """Resolve the approved Python 3.11 base interpreter from settings."""
    setting = os.environ.get("BONDS_LITE_MODEL_PYTHON", "").strip()
    if not setting:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, "Environment") as key:
            setting = winreg.QueryValueEx(key, "BONDS_LITE_MODEL_PYTHON")[0]
    path = Path(os.path.expandvars(setting)).expanduser()
    if path.is_dir():
        path = path / "python.exe"
    if not path.is_file():
        raise FileNotFoundError(f"Model Python not found: {path}")
    return path.resolve()


def prepare(base: Path, environment: Path, requirements: Path) -> Path:
    """Create or reuse one venv and install its pinned build requirements."""
    base = base.resolve(strict=True)
    environment = environment.resolve()
    python = environment / "Scripts" / "python.exe"
    if not python.is_file():
        environment.parent.mkdir(parents=True, exist_ok=True)
        subprocess.run(
            [str(base), "-m", "venv", str(environment)],
            check=True,
            shell=False,
        )
    subprocess.run(
        [str(python), "-m", "pip", "install", "--disable-pip-version-check", "-r", str(requirements)],
        check=True,
        shell=False,
    )
    print(f"Prepared {environment} from {base}", flush=True)
    return python


def main() -> int:
    """Prepare isolated controller Python 3.13 and model Python 3.11 venvs."""
    project = Path(__file__).resolve().parents[1]
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--controller-base", type=Path, default=Path(sys.executable))
    parser.add_argument("--model-base", type=Path, default=current_model_python())
    parser.add_argument(
        "--env-root",
        type=Path,
        default=Path(os.environ.get("LOCALAPPDATA", Path.home()))
        / "BondsLiteBuild" / "venvs",
    )
    options = parser.parse_args()
    if os.name != "nt":
        parser.error("The Windows/Excel build environments must be prepared on Windows")
    controller = prepare(
        options.controller_base,
        options.env_root / "controller-py313",
        project / "tools" / "requirements-controller-build.txt",
    )
    model = prepare(
        options.model_base,
        options.env_root / "model-py311",
        project / "tools" / "requirements-model-build.txt",
    )
    print(f"CONTROLLER_PYTHON={controller}", flush=True)
    print(f"MODEL_PYTHON={model}", flush=True)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
