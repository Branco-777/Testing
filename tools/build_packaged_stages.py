"""Build and smoke-test the three Windows applications before publishing."""

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
import time
from datetime import datetime, timezone
from pathlib import Path

CENTRAL_RELEASE_ROOT = Path(
    r"\\server090\internal model$\29 Financial Resilience\Sandbox\SST spreadsheet executable file"
)
MACALC_SOURCE_NAMES = (
    "MACalcrunmanager.py",
    "MACalc.py",
    "MACE.py",
    "calcSCR.py",
    "calcmp.py",
    "MACalcutils.py",
)
COMPONENT_CACHE_DIR = "components"
EXCLUDED_SOURCE_PARTS = {"__pycache__", ".git", ".venv", "venv", "build", "dist"}
MACALC_PACKAGE_POLICY = "windows-x64-cbc-v1"


def run(
    command: list[str],
    cwd: Path,
    timeout: int = 1800,
    capture: bool = True,
) -> str:
    """Run build tooling with captured diagnostics and a finite deadline."""
    print("Running: " + subprocess.list2cmdline(command), flush=True)
    result = subprocess.run(
        command,
        cwd=cwd,
        capture_output=capture,
        text=True,
        errors="replace",
        timeout=timeout,
        shell=False,
    )
    if result.returncode:
        if capture:
            print(result.stdout[-12000:], flush=True)
            print(result.stderr[-12000:], flush=True)
        raise RuntimeError(f"Command failed with exit code {result.returncode}")
    return result.stdout.strip() if capture else ""


def pyinstaller_command(
    python: Path,
    entry: Path,
    name: str,
    dist: Path,
    work: Path,
    spec: Path,
    project: Path,
    model: Path | None = None,
    clean: bool = False,
) -> list[str]:
    """Create an incremental PyInstaller command; clean rebuilds are opt-in."""
    options = [
        str(python),
        "-m",
        "PyInstaller",
        "--noconfirm",
        "--onedir",
        "--console",
        "--noupx",
        "--name",
        name,
        "--distpath",
        str(dist),
        "--workpath",
        str(work),
        "--specpath",
        str(spec),
        "--paths",
        str(project / "src"),
        "--hidden-import",
        "pythoncom",
        "--hidden-import",
        "pywintypes",
        "--exclude-module",
        "matplotlib",
        "--exclude-module",
        "IPython",
        "--exclude-module",
        "notebook",
        "--exclude-module",
        "tkinter",
    ]
    if clean:
        options.append("--clean")
    if model is not None:
        options += ["--paths", str(model)]
        if name == "MACalcModel":
            options += [
                "--paths",
                str(model / "LPIcurvestool"),
                "--collect-all",
                "pulp",
                "--collect-all",
                "smithwilson",
            ]
            for module in (
                *(Path(filename).stem for filename in MACALC_SOURCE_NAMES),
                "LPIcurvestool.LPIcurves",
                "utils",
                "JustStats.other_distns",
            ):
                options += ["--hidden-import", module]
        else:
            options += [
                "--add-data",
                f"{model / 'pyBondsHMutils.py'}{os.pathsep}.",
            ]
    else:
        options += [
            "--hidden-import",
            "sst_stage_three",
            "--hidden-import",
            "sst_stage_four",
            "--hidden-import",
            "sst_property_stage_one",
            "--hidden-import",
            "sst_property_stage_two",
            "--hidden-import",
            "sst_property_stage_two_watchdog",
        ]
        for module in ("scipy", "numba", "llvmlite", "pandas", "numpy"):
            options += ["--exclude-module", module]
    return options + [str(entry)]


def build(
    python: Path,
    entry: Path,
    name: str,
    build_root: Path,
    dist: Path,
    project: Path,
    model: Path | None = None,
    clean: bool = False,
) -> Path:
    """Build one app while reusing its analysis cache unless explicitly cleaned."""
    work = build_root / "work" / name
    spec = build_root / "spec"
    command = pyinstaller_command(
        python,
        entry,
        name,
        dist,
        work,
        spec,
        project,
        model=model,
        clean=clean,
    )
    run(command, project, capture=False)
    return dist / name


def prune_macalc_solver_binaries(application: Path) -> None:
    """Keep only the Windows 64-bit CBC solver in the MACalc bundle."""
    solver_root = (
        application / "_internal" / "pulp" / "solverdir" / "cbc"
    )
    windows_x64_solver = solver_root / "win" / "i64" / "cbc.exe"
    if not windows_x64_solver.is_file():
        raise RuntimeError(
            f"Windows 64-bit CBC solver is missing from {application}"
        )
    platforms = list(solver_root.iterdir())
    for platform in platforms:
        if not platform.is_dir():
            continue
        if platform.name.casefold() == "licenses":
            continue
        if platform.name.casefold() != "win":
            _preserve_solver_licences(platform, solver_root)
            shutil.rmtree(platform)
            continue
        for architecture in platform.iterdir():
            if architecture.is_dir() and architecture.name.casefold() != "i64":
                _preserve_solver_licences(architecture, solver_root)
                shutil.rmtree(architecture)


def _preserve_solver_licences(directory: Path, solver_root: Path) -> None:
    """Keep licence and notice files from pruned solver platforms."""
    licence_tokens = ("license", "licence", "copying", "notice")
    for path in directory.rglob("*"):
        if path.is_file() and any(
            token in path.name.casefold() for token in licence_tokens
        ):
            destination = solver_root / "licenses" / path.relative_to(solver_root)
            destination.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(path, destination)


def sha256_file(path: Path) -> str:
    """Return the SHA-256 hex digest of one file."""
    return hashlib.sha256(path.read_bytes()).hexdigest()


def python_environment(python: Path, project: Path) -> list[str]:
    """Describe the interpreter version and every installed distribution."""
    probe = (
        "import sys,json,importlib.metadata as m; print(json.dumps([sys.version]+"
        "sorted(str(d.metadata['Name'])+'=='+str(d.version) for d in m.distributions())))"
    )
    return json.loads(run([str(python), "-c", probe], project, timeout=60))


def component_fingerprint(
    python: Path,
    entry: Path,
    name: str,
    project: Path,
    source_dir: Path,
    environment: list[str],
) -> str:
    """Fingerprint every input that determines a model application build.

    The fingerprint covers the PyInstaller options, the entry script, every
    Python file in the model source folder and the model interpreter's
    installed packages, so a cached build is reused only when none changed.
    """
    command = pyinstaller_command(
        python,
        entry,
        name,
        Path("<dist>"),
        Path("<work>"),
        Path("<spec>"),
        project,
        model=source_dir,
    )
    digest = hashlib.sha256()
    digest.update(json.dumps({"command": command, "environment": environment}).encode())
    if name == "MACalcModel":
        digest.update(MACALC_PACKAGE_POLICY.encode())
    digest.update(sha256_file(entry).encode())
    for path in sorted(source_dir.rglob("*.py")):
        relative = path.relative_to(source_dir)
        if EXCLUDED_SOURCE_PARTS.intersection(relative.parts):
            continue
        digest.update(f"{relative.as_posix()}:{sha256_file(path)}".encode())
    return digest.hexdigest()


def reuse_component(
    build_root: Path, name: str, fingerprint: str, dist: Path
) -> Path | None:
    """Copy a previously verified build into ``dist`` if its fingerprint matches."""
    cache = build_root / COMPONENT_CACHE_DIR / name
    try:
        record = json.loads((cache / "component.json").read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return None
    executable = cache / "app" / f"{name}.exe"
    if (
        record.get("fingerprint") != fingerprint
        or not executable.is_file()
        or sha256_file(executable) != record.get("executable_sha256")
    ):
        return None
    application = dist / name
    shutil.copytree(cache / "app", application)
    return application


def store_component(
    build_root: Path, name: str, fingerprint: str, application: Path
) -> None:
    """Cache a self-tested build; the record is written last so partial copies are ignored."""
    cache = build_root / COMPONENT_CACHE_DIR / name
    if cache.exists():
        shutil.rmtree(cache)
    shutil.copytree(application, cache / "app")
    record = {
        "fingerprint": fingerprint,
        "executable_sha256": sha256_file(application / f"{name}.exe"),
    }
    (cache / "component.json").write_text(json.dumps(record, indent=2), encoding="utf-8")


def publish_bundle(bundle: Path, output: Path) -> None:
    """Copy the verified bundle to the release folder with multi-threaded robocopy.

    The central share is the slow step of a release, so files are copied in
    parallel. Every staged file must then exist in the release with the same
    size; otherwise the partial release folder is removed.
    """
    command = [
        "robocopy",
        str(bundle),
        str(output),
        "/E",
        "/MT:16",
        "/R:2",
        "/W:5",
        "/NP",
        "/NFL",
        "/NDL",
        "/NJH",
    ]
    print("Running: " + subprocess.list2cmdline(command), flush=True)
    try:
        # Robocopy exit codes below 8 mean success; 8 or above mean failures.
        result = subprocess.run(command, capture_output=True, text=True, errors="replace")
        if result.returncode >= 8:
            print(result.stdout[-12000:], result.stderr[-12000:], flush=True)
            raise RuntimeError(f"robocopy failed with exit code {result.returncode}")
        missing = [
            str(source.relative_to(bundle))
            for source in bundle.rglob("*")
            if source.is_file()
            and (
                not (target := output / source.relative_to(bundle)).is_file()
                or target.stat().st_size != source.stat().st_size
            )
        ]
        if missing:
            raise RuntimeError(
                f"{len(missing)} file(s) missing or incomplete after publishing, "
                f"e.g. {missing[0]}"
            )
    except BaseException:
        shutil.rmtree(output.parent, ignore_errors=True)
        raise


def publish_component(
    application: Path,
    release_root: Path,
    name: str,
    fingerprint: str,
) -> bool:
    """Publish one immutable component; return whether a verified copy existed."""
    if name not in {"BondsLiteModel", "MACalcModel"}:
        raise ValueError(f"Unknown packaged component: {name}")
    if not re.fullmatch(r"[0-9a-f]{64}", fingerprint):
        raise ValueError(f"Invalid component fingerprint: {fingerprint}")
    executable = application / f"{name}.exe"
    if not executable.is_file():
        raise RuntimeError(f"Packaged component executable is missing: {executable}")
    executable_hash = sha256_file(executable)
    component_parent = release_root / "components" / name
    component_parent.mkdir(parents=True, exist_ok=True)
    component = component_parent / fingerprint
    if component.exists():
        record_path = component / "component.json"
        marker_path = component / "component.id"
        try:
            record = json.loads(record_path.read_text(encoding="utf-8"))
            marker = marker_path.read_text(encoding="ascii")
        except (OSError, ValueError) as error:
            raise RuntimeError(
                f"Shared component record is missing or invalid: {component}"
            ) from error
        published_executable = component / "app" / executable.name
        if (
            not isinstance(record, dict)
            or record.get("name") != name
            or record.get("fingerprint") != fingerprint
            or record.get("executable_sha256") != executable_hash
            or marker != f"{name}={fingerprint}"
            or not published_executable.is_file()
            or published_executable.stat().st_size != executable.stat().st_size
            or sha256_file(published_executable) != executable_hash
        ):
            raise RuntimeError(
                f"Shared component does not match the current build: {component}"
            )
        return True

    with tempfile.TemporaryDirectory(
        prefix=f".{fingerprint}.partial-", dir=component_parent
    ) as temporary:
        staging = Path(temporary)
        publish_bundle(application, staging / "app")
        if sha256_file(staging / "app" / executable.name) != executable_hash:
            raise RuntimeError(f"Shared component copy hash mismatch: {name}")
        record = {
            "name": name,
            "fingerprint": fingerprint,
            "executable_sha256": executable_hash,
        }
        (staging / "component.id").write_text(
            f"{name}={fingerprint}", encoding="ascii"
        )
        (staging / "component.json").write_text(
            json.dumps(record, indent=2), encoding="utf-8"
        )
        for attempt in range(5):
            try:
                staging.rename(component)
                break
            except PermissionError:
                if attempt == 4:
                    raise
                time.sleep(0.1 * (attempt + 1))
    return False


def collect_licences(python: Path, destination: Path, project: Path) -> None:
    """Include available runtime dependency licences alongside the application."""
    probe = (
        "import importlib.metadata as m,json; "
        "names=['pyinstaller','numpy','pandas','scipy','numba','llvmlite','xlwings',"
        "'openpyxl','python-dateutil','pywin32','pulp','smithwilson']; out=[]; "
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
    (destination / "versions.json").write_text(
        json.dumps(packages, indent=2), encoding="utf-8"
    )


def main() -> int:
    """Build all runtimes without changing model sources or financial inputs."""
    project = Path(__file__).resolve().parents[1]
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--controller-python", type=Path, default=Path(sys.executable))
    parser.add_argument("--model-python", type=Path, required=True)
    parser.add_argument("--model-dir", type=Path, required=True)
    parser.add_argument("--macalc-dir", type=Path, required=True)
    parser.add_argument("--build-root", type=Path)
    parser.add_argument("--release-root", type=Path, default=CENTRAL_RELEASE_ROOT)
    parser.add_argument(
        "--release-id", default=datetime.now().strftime("%Y%m%d-stages-v1")
    )
    parser.add_argument(
        "--clean",
        action="store_true",
        help="Discard PyInstaller analysis caches and rebuild every application",
    )
    parser.add_argument(
        "--rebuild-all",
        action="store_true",
        help="Rebuild the model and MACalc applications even if their inputs are unchanged",
    )
    parser.add_argument(
        "--promote",
        action="store_true",
        help="Update current.txt only after the release passes all build checks",
    )
    options = parser.parse_args()
    if os.name != "nt":
        parser.error("The Windows/Excel bundle must be built on Windows")
    if not re.fullmatch(r"[A-Za-z0-9][A-Za-z0-9._-]{0,79}", options.release_id):
        parser.error(
            "Release ID may contain only letters, numbers, dot, underscore and hyphen"
        )
    release_root = options.release_root.expanduser().resolve()
    release_parent = release_root / "releases" / options.release_id
    output = release_parent / "Bonds Lite App"
    if release_parent.exists():
        parser.error(
            "Release ID already exists; choose another to preserve previous releases"
        )
    model = options.model_dir.resolve(strict=True)
    source_names = ("run_bonds_batch.py", "BondsModelHandler.py", "pyBondsHMutils.py")
    hashes = {
        name: hashlib.sha256((model / name).read_bytes()).hexdigest()
        for name in source_names
    }
    macalc = options.macalc_dir.resolve(strict=True)
    macalc_hashes = {
        name: hashlib.sha256((macalc / name).read_bytes()).hexdigest()
        for name in MACALC_SOURCE_NAMES
    }
    local_app = os.environ.get("LOCALAPPDATA")
    build_root = options.build_root or (
        Path(local_app) / "BondsLiteBuild"
        if local_app
        else Path.home() / ".bonds-lite-build"
    )
    build_root = build_root.resolve()
    build_root.mkdir(parents=True, exist_ok=True)
    model_python = options.model_python.resolve(strict=True)
    environment = python_environment(model_python, project)
    components = {
        "BondsLiteModel": (project / "src" / "bonds_lite_model_entry.py", model),
        "MACalcModel": (project / "src" / "macalc_batch_entry.py", macalc),
    }

    def fingerprints() -> dict[str, str]:
        return {
            name: component_fingerprint(
                model_python, entry, name, project, source, environment
            )
            for name, (entry, source) in components.items()
        }

    expected_fingerprints = fingerprints()
    allow_reuse = not (options.clean or options.rebuild_all)
    reused: list[str] = []
    with tempfile.TemporaryDirectory(prefix="release-stage-", dir=build_root) as stage:
        stage_root = Path(stage)
        dist = stage_root / "dist"

        def model_application(name: str) -> Path:
            entry, source = components[name]
            if allow_reuse:
                application = reuse_component(
                    build_root, name, expected_fingerprints[name], dist
                )
                if application is not None:
                    print(
                        f"Reusing unchanged {name} build "
                        f"({expected_fingerprints[name][:12]})",
                        flush=True,
                    )
                    reused.append(name)
                    return application
            return build(
                model_python,
                entry,
                name,
                build_root,
                dist,
                project,
                model=source,
                clean=options.clean,
            )

        controller = build(
            options.controller_python.resolve(strict=True),
            project / "tools" / "controller_entry.py",
            "BondsLiteController",
            build_root,
            dist,
            project,
            clean=options.clean,
        )
        model_app = model_application("BondsLiteModel")
        macalc_app = model_application("MACalcModel")
        prune_macalc_solver_binaries(macalc_app)
        bundle = stage_root / "Bonds Lite App"
        bundle.mkdir()
        shutil.move(str(controller), bundle / "Controller")
        shutil.move(str(model_app), bundle / "Model")
        shutil.move(str(macalc_app), bundle / "MACalc")
        controller_exe = bundle / "Controller" / "BondsLiteController.exe"
        model_exe = bundle / "Model" / "BondsLiteModel.exe"
        macalc_exe = bundle / "MACalc" / "MACalcModel.exe"
        print(run([str(model_exe), "--self-test"], bundle, timeout=300), flush=True)
        print(
            run([str(controller_exe), "--self-test"], bundle, timeout=180), flush=True
        )
        print(run([str(macalc_exe), "--self-test"], bundle, timeout=300), flush=True)
        collect_licences(
            options.controller_python.resolve(strict=True),
            bundle / "Licences" / "Controller",
            project,
        )
        collect_licences(
            options.model_python.resolve(strict=True),
            bundle / "Licences" / "Model",
            project,
        )
        if hashes != {
            name: hashlib.sha256((model / name).read_bytes()).hexdigest()
            for name in source_names
        }:
            raise RuntimeError(
                "Model sources changed during the build; release not published"
            )
        if macalc_hashes != {
            name: hashlib.sha256((macalc / name).read_bytes()).hexdigest()
            for name in MACALC_SOURCE_NAMES
        }:
            raise RuntimeError(
                "MACalc sources changed during the build; release not published"
            )
        if fingerprints() != expected_fingerprints:
            raise RuntimeError(
                "Model or MACalc build inputs changed during the build; "
                "release not published"
            )
        for name, folder in (("BondsLiteModel", "Model"), ("MACalcModel", "MACalc")):
            if name not in reused:
                store_component(
                    build_root, name, expected_fingerprints[name], bundle / folder
                )
        shared_reuse = {
            name: publish_component(
                bundle / folder,
                release_root,
                name,
                expected_fingerprints[name],
            )
            for name, folder in (
                ("BondsLiteModel", "Model"),
                ("MACalcModel", "MACalc"),
            )
        }
        manifest = {
            "release_id": options.release_id,
            "built_utc": datetime.now(timezone.utc).isoformat(),
            "layout_version": 2,
            "model_source_sha256": hashes,
            "macalc_source_sha256": macalc_hashes,
            "component_fingerprints": expected_fingerprints,
            "reused_components": reused,
            "shared_components_reused": [
                name for name, was_reused in shared_reuse.items() if was_reused
            ],
            "controller_sha256": hashlib.sha256(
                controller_exe.read_bytes()
            ).hexdigest(),
            "model_executable_sha256": hashlib.sha256(
                model_exe.read_bytes()
            ).hexdigest(),
            "macalc_executable_sha256": hashlib.sha256(
                macalc_exe.read_bytes()
            ).hexdigest(),
            "platform": "windows-x86_64",
            "layout": "Controller; content-addressed Model and MACalc components",
            "prerequisites": (
                "64-bit Windows, desktop Excel, JRC add-in, "
                "macro/executable permissions, scenario access"
            ),
        }
        component_map = (
            f"model={expected_fingerprints['BondsLiteModel']}\n"
            f"macalc={expected_fingerprints['MACalcModel']}\n"
        )
        release_bundle = stage_root / "Bonds Lite Release"
        release_bundle.mkdir()
        shutil.copytree(bundle / "Controller", release_bundle / "Controller")
        shutil.copytree(bundle / "Licences", release_bundle / "Licences")
        (release_bundle / "components.txt").write_text(
            component_map, encoding="ascii"
        )
        (release_bundle / "release.json").write_text(
            json.dumps(manifest, indent=2), encoding="utf-8"
        )
        release_parent.mkdir(parents=True)
        publish_bundle(release_bundle, output)
    if options.promote:
        release_root.mkdir(parents=True, exist_ok=True)
        pointer_temporary = release_root / f"current-{os.getpid()}.tmp"
        pointer_temporary.write_text(options.release_id, encoding="ascii")
        pointer_temporary.replace(release_root / "current.txt")
    print(f"Published: {output}", flush=True)
    if options.promote:
        print(
            f"Promoted: {release_root / 'current.txt'} -> {options.release_id}",
            flush=True,
        )
    print(
        "The central release is a package source; runtime executables launch from LocalAppData",
        flush=True,
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
