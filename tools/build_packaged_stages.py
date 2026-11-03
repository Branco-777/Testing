"""Build and smoke-test both Windows one-folder applications before publishing."""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
from datetime import datetime, timezone
from pathlib import Path

CENTRAL_RELEASE_ROOT = Path(
    r"\\server090\internal model$\29 Financial Resilience\Sandbox\SST spreadsheet executable file"
)


def run(
    command: list[str], cwd: Path, timeout: int = 1800, capture: bool = True,
) -> str:
    """Run build tooling with captured diagnostics and a finite deadline."""
    print("Running: " + subprocess.list2cmdline(command), flush=True)
    result = subprocess.run(
        command, cwd=cwd, capture_output=capture, text=True, errors="replace",
        timeout=timeout, shell=False,
    )
    if result.returncode:
        if capture:
            print(result.stdout[-12000:], flush=True)
            print(result.stderr[-12000:], flush=True)
        raise RuntimeError(f"Command failed with exit code {result.returncode}")
    return result.stdout.strip() if capture else ""


def pyinstaller_command(
    python: Path, entry: Path, name: str, dist: Path, work: Path, spec: Path,
    project: Path, model: Path | None = None, clean: bool = False,
) -> list[str]:
    """Create an incremental PyInstaller command; clean rebuilds are opt-in."""
    options = [
        str(python), "-m", "PyInstaller", "--noconfirm", "--onedir",
        "--console", "--noupx", "--name", name,
        "--distpath", str(dist), "--workpath", str(work),
        "--specpath", str(spec), "--paths", str(project / "src"),
        "--hidden-import", "pythoncom", "--hidden-import", "pywintypes",
        "--hidden-import", "sst_stage_three",
        "--exclude-module", "matplotlib", "--exclude-module", "IPython",
        "--exclude-module", "notebook", "--exclude-module", "tkinter",
    ]
    if clean:
        options.append("--clean")
    if model is not None:
        options += [
            "--paths", str(model),
            "--add-data", f"{model / 'pyBondsHMutils.py'}{os.pathsep}.",
        ]
    else:
        for module in ("scipy", "numba", "llvmlite", "pandas", "numpy"):
            options += ["--exclude-module", module]
    return options + [str(entry)]


def build(
    python: Path, entry: Path, name: str, build_root: Path, dist: Path,
    project: Path, model: Path | None = None, clean: bool = False,
) -> Path:
    """Build one app while reusing its analysis cache unless explicitly cleaned."""
    work = build_root / "work" / name
    spec = build_root / "spec"
    command = pyinstaller_command(
        python, entry, name, dist, work, spec, project, model=model, clean=clean,
    )
    run(command, project, capture=False)
    return dist / name


def collect_licences(python: Path, destination: Path, project: Path) -> None:
    """Include available runtime dependency licences alongside the application."""
    probe = (
        "import importlib.metadata as m,json; "
        "names=['pyinstaller','numpy','pandas','scipy','numba','llvmlite','xlwings',"
        "'openpyxl','python-dateutil','pywin32']; out=[]; "
        "[(out.append({'name':d.metadata['Name'],'version':d.version,'files':"
        "[str(d.locate_file(f)) for f in (d.files or []) "
        "if any(s in str(f).lower() for s in ('license','licence','copying','notice'))]})) "
        "for d in m.distributions() if d.metadata['Name'].lower() in names]; print(json.dumps(out))"
    )
    packages = json.loads(run([str(python), "-c", probe], project, timeout=60))
    destination.mkdir(parents=True)
    for package in packages:
        folder = destination / f"{package['name']}-{package['version']}"
        folder.mkdir()
        for index, filename in enumerate(package["files"]):
            source = Path(filename)
            if source.is_file():
                shutil.copy2(source, folder / f"{index}-{source.name}")
    (destination / "versions.json").write_text(json.dumps(packages, indent=2), encoding="utf-8")


def main() -> int:
    """Build both runtimes without changing the model source or financial inputs."""
    project = Path(__file__).resolve().parents[1]
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--controller-python", type=Path, default=Path(sys.executable))
    parser.add_argument("--model-python", type=Path, required=True)
    parser.add_argument("--model-dir", type=Path, required=True)
    parser.add_argument("--build-root", type=Path)
    parser.add_argument("--release-root", type=Path, default=CENTRAL_RELEASE_ROOT)
    parser.add_argument("--release-id", default=datetime.now().strftime("%Y%m%d-stages-v1"))
    parser.add_argument("--clean", action="store_true", help="Discard PyInstaller analysis caches")
    parser.add_argument(
        "--promote", action="store_true",
        help="Update current.txt only after the release passes all build checks",
    )
    options = parser.parse_args()
    if os.name != "nt":
        parser.error("The Windows/Excel bundle must be built on Windows")
    if not re.fullmatch(r"[A-Za-z0-9][A-Za-z0-9._-]{0,79}", options.release_id):
        parser.error("Release ID may contain only letters, numbers, dot, underscore and hyphen")
    release_root = options.release_root.expanduser().resolve()
    release_parent = release_root / "releases" / options.release_id
    output = release_parent / "Bonds Lite App"
    if release_parent.exists():
        parser.error("Release ID already exists; choose another to preserve previous releases")
    model = options.model_dir.resolve(strict=True)
    source_names = ("run_bonds_batch.py", "BondsModelHandler.py", "pyBondsHMutils.py")
    hashes = {name: hashlib.sha256((model / name).read_bytes()).hexdigest() for name in source_names}
    local_app = os.environ.get("LOCALAPPDATA")
    build_root = options.build_root or (
        Path(local_app) / "BondsLiteBuild" if local_app else Path.home() / ".bonds-lite-build"
    )
    build_root = build_root.resolve()
    build_root.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix="release-stage-", dir=build_root) as stage:
        stage_root = Path(stage)
        dist = stage_root / "dist"
        controller = build(
            options.controller_python.resolve(strict=True),
            project / "tools" / "controller_entry.py", "BondsLiteController",
            build_root, dist, project, clean=options.clean,
        )
        model_app = build(
            options.model_python.resolve(strict=True),
            project / "src" / "bonds_lite_model_entry.py", "BondsLiteModel",
            build_root, dist, project, model=model, clean=options.clean,
        )
        bundle = stage_root / "Bonds Lite App"
        bundle.mkdir()
        shutil.move(str(controller), bundle / "Controller")
        shutil.move(str(model_app), bundle / "Model")
        controller_exe = bundle / "Controller" / "BondsLiteController.exe"
        model_exe = bundle / "Model" / "BondsLiteModel.exe"
        print(run([str(model_exe), "--self-test"], bundle, timeout=300), flush=True)
        print(run([str(controller_exe), "--self-test"], bundle, timeout=180), flush=True)
        collect_licences(
            options.controller_python.resolve(strict=True), bundle / "Licences" / "Controller", project
        )
        collect_licences(options.model_python.resolve(strict=True), bundle / "Licences" / "Model", project)
        if hashes != {name: hashlib.sha256((model / name).read_bytes()).hexdigest() for name in source_names}:
            raise RuntimeError("Model sources changed during the build; release not published")
        manifest = {
            "release_id": options.release_id,
            "built_utc": datetime.now(timezone.utc).isoformat(),
            "model_source_sha256": hashes,
            "controller_sha256": hashlib.sha256(controller_exe.read_bytes()).hexdigest(),
            "model_executable_sha256": hashlib.sha256(model_exe.read_bytes()).hexdigest(),
            "layout": "Controller/BondsLiteController.exe; Model/BondsLiteModel.exe",
            "prerequisites": "Windows desktop Excel, JRC add-in, macro/executable permissions, scenario access",
        }
        (bundle / "release.json").write_text(json.dumps(manifest, indent=2), encoding="utf-8")
        release_parent.mkdir(parents=True)
        shutil.copytree(bundle, output)
    if options.promote:
        release_root.mkdir(parents=True, exist_ok=True)
        pointer_temporary = release_root / f"current-{os.getpid()}.tmp"
        pointer_temporary.write_text(options.release_id, encoding="ascii")
        pointer_temporary.replace(release_root / "current.txt")
    print(f"Published: {output}", flush=True)
    if options.promote:
        print(f"Promoted: {release_root / 'current.txt'} -> {options.release_id}", flush=True)
    print("The central release is a package source; runtime executables launch from LocalAppData", flush=True)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())