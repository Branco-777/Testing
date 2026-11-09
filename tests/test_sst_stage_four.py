"""Contracts for the Stage 4 handoff to the unchanged MACalc Batch model."""

from __future__ import annotations

from io import StringIO
import logging
import sys
from pathlib import Path
from types import SimpleNamespace
from typing import Any

import pytest


def _stage_four() -> Any:
    """Import lazily so contract tests collect before Stage 4 is implemented."""
    import sst_stage_four

    return sst_stage_four


def _handoff_case(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    *,
    workbook_name: str = "MACalc input_Test_0123456789abcdef.xlsb",
    status: object = "Succeeded",
    input_run: object = "SST Output_Test/Credit",
    downgrade: object = True,
) -> tuple[Any, Path, Path, dict[str, object]]:
    """Create the live folder shape and isolate the Draft's named audit reads."""
    import bonds_lite_input as shared
    import sst_stage_three as stage_three

    stage_four = _stage_four()
    caller_path = tmp_path / "Draft.xlsm"
    caller_path.write_bytes(b"fake Draft")
    run_folder = tmp_path / "SST Output_Test" / "Credit"
    run_folder.mkdir(parents=True)
    workbook = run_folder / workbook_name
    if workbook_name:
        workbook.write_bytes(b"fake MACalc input")
    fields: dict[str, object] = {
        "status": status,
        "input_run": input_run,
        "last_output": f"SST Output_Test/Credit/{workbook_name}",
    }
    caller = SimpleNamespace(fullname=str(caller_path))

    monkeypatch.setattr(shared, "resolve_caller_path", lambda book: caller_path)
    monkeypatch.setattr(shared, "_caller_path", lambda book: caller_path)
    monkeypatch.setattr(shared, "resolve_stage_two_results", lambda book: run_folder)
    monkeypatch.setattr(
        shared, "portable_output_path", lambda path, output: "SST Output_Test/Credit"
    )
    monkeypatch.setattr(
        shared,
        "read_stress_flags",
        lambda book: {
            "SST_Credit_DG_standalone": downgrade,
            "SST_Credit_spread_standalone": False,
        },
    )
    monkeypatch.setattr(
        stage_three, "_read_stage_three_fields", lambda book: fields.copy()
    )
    return caller, caller_path, run_folder, fields


def test_resolve_stage_four_input_accepts_the_current_audited_stage_three_workbook(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    stage_four = _stage_four()
    caller, _, run_folder, _ = _handoff_case(tmp_path, monkeypatch)

    handoff = stage_four.resolve_stage_four_input(caller)

    assert handoff.run_folder == run_folder
    assert handoff.workbook == run_folder / "MACalc input_Test_0123456789abcdef.xlsb"
    assert handoff.is_downgrade_run is True


@pytest.mark.parametrize(
    "status", ["Failed", "Running", "Interrupted", "Not Run", "", None]
)
def test_resolve_stage_four_input_rejects_non_succeeded_stage_three_audit(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    status: object,
) -> None:
    stage_four = _stage_four()
    caller, _, _, _ = _handoff_case(tmp_path, monkeypatch, status=status)

    with pytest.raises(stage_four.StageFourError, match=r"(?i)rerun Stage 3"):
        stage_four.resolve_stage_four_input(caller)


def test_resolve_stage_four_input_rejects_stage_three_from_another_stage_one_run(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    stage_four = _stage_four()
    caller, _, _, fields = _handoff_case(tmp_path, monkeypatch)
    fields["input_run"] = "SST Output_Other/Credit"

    with pytest.raises(stage_four.StageFourError, match=r"(?i)rerun Stage 3"):
        stage_four.resolve_stage_four_input(caller)


def test_resolve_stage_four_input_propagates_stage_two_not_ready(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    import bonds_lite_input as shared

    stage_four = _stage_four()
    caller, _, _, _ = _handoff_case(tmp_path, monkeypatch)

    def reject_stage_two(book: object) -> Path:
        raise shared.BondsLiteError("Stage 2 has not succeeded for the current run")

    monkeypatch.setattr(shared, "resolve_stage_two_results", reject_stage_two)
    with pytest.raises(shared.BondsLiteError, match="Stage 2"):
        stage_four.resolve_stage_four_input(caller)


@pytest.mark.parametrize(
    "bad_name",
    [
        "",
        "MACalc input.xlsb",
        "MACalc input_Test_0123456789abcde.xlsb",
        "MACalc input_Test_0123456789abcdef0.xlsb",
        "MACalc input_Test_0123456789ABCDEF.xlsb",
        "MACalc input_Other_0123456789abcdef.xlsb",
    ],
)
def test_resolve_stage_four_input_rejects_invalid_handoff_names(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    bad_name: str,
) -> None:
    stage_four = _stage_four()
    caller, _, run_folder, fields = _handoff_case(
        tmp_path, monkeypatch, workbook_name=bad_name
    )
    fields["last_output"] = f"SST Output_Test/Credit/{bad_name}" if bad_name else None

    with pytest.raises(stage_four.StageFourError):
        stage_four.resolve_stage_four_input(caller)


@pytest.mark.parametrize(
    "target", ["missing", "folder", "nested", "staging", "other stress"]
)
def test_resolve_stage_four_input_rejects_non_published_or_wrong_folder_paths(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    target: str,
) -> None:
    stage_four = _stage_four()
    caller, caller_path, run_folder, fields = _handoff_case(tmp_path, monkeypatch)
    good_name = "MACalc input_Test_0123456789abcdef.xlsb"
    (run_folder / good_name).unlink()
    if target == "missing":
        path = run_folder / good_name
    elif target == "folder":
        path = run_folder / good_name
        path.mkdir()
    elif target == "nested":
        path = run_folder / "subfolder" / good_name
        path.parent.mkdir()
        path.write_bytes(b"not published at the stress-folder root")
    elif target == "staging":
        path = run_folder.parent / "Stage 3 staging" / "token" / good_name
        path.parent.mkdir(parents=True)
        path.write_bytes(b"incomplete staged workbook")
    else:
        path = run_folder.parent / "Property" / good_name
        path.parent.mkdir()
        path.write_bytes(b"wrong stress folder")
    fields["last_output"] = path.relative_to(caller_path.parent).as_posix()

    with pytest.raises(stage_four.StageFourError):
        stage_four.resolve_stage_four_input(caller)


def test_resolve_stage_four_input_rejects_symlink_handoff(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    stage_four = _stage_four()
    caller, caller_path, run_folder, fields = _handoff_case(tmp_path, monkeypatch)
    linked = run_folder / "MACalc input_Test_0123456789abcdef.xlsb"
    linked.unlink()
    real = run_folder / "actual workbook.xlsb"
    real.write_bytes(b"target")
    try:
        linked.symlink_to(real)
    except (OSError, NotImplementedError) as error:
        pytest.skip(f"Symlinks are unavailable: {error}")
    fields["last_output"] = linked.relative_to(caller_path.parent).as_posix()

    with pytest.raises(stage_four.StageFourError):
        stage_four.resolve_stage_four_input(caller)


def test_resolve_stage_four_input_stops_when_macalc_workbook_is_locked(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    stage_four = _stage_four()
    caller, _, run_folder, _ = _handoff_case(tmp_path, monkeypatch)
    (run_folder / "~$MACalc input_Test_0123456789abcdef.xlsb").write_bytes(b"lock")

    with pytest.raises(stage_four.StageFourError, match=r"(?i)lock|open"):
        stage_four.resolve_stage_four_input(caller)


@pytest.mark.parametrize("downgrade", [True, False])
def test_resolve_stage_four_input_preserves_boolean_credit_downgrade_setting(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    downgrade: bool,
) -> None:
    stage_four = _stage_four()
    caller, _, _, _ = _handoff_case(tmp_path, monkeypatch, downgrade=downgrade)

    handoff = stage_four.resolve_stage_four_input(caller)

    assert handoff.is_downgrade_run is downgrade


@pytest.mark.parametrize("invalid", ["TRUE", 1, None, ""])
def test_resolve_stage_four_input_rejects_non_boolean_downgrade_flag(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    invalid: object,
) -> None:
    import bonds_lite_input as shared

    stage_four = _stage_four()
    caller, _, _, _ = _handoff_case(tmp_path, monkeypatch, downgrade=invalid)

    def reject_flag(book: object) -> dict[str, bool]:
        raise shared.BondsLiteError("Stress flag must be TRUE or FALSE")

    monkeypatch.setattr(shared, "read_stress_flags", reject_flag)
    with pytest.raises(shared.BondsLiteError, match="TRUE or FALSE"):
        stage_four.resolve_stage_four_input(caller)


def test_load_macalc_config_uses_the_configured_user_python_and_model_folder(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    stage_four = _stage_four()
    caller = tmp_path / "Draft.xlsm"
    caller.write_bytes(b"Draft")
    python = tmp_path / "a different user's py311" / "python.exe"
    python.parent.mkdir()
    python.write_bytes(b"fake interpreter")
    model = tmp_path / "MACalc Batch working copy"
    model.mkdir()
    (model / "MACalcrunmanager.py").write_text("# fake manager\n", encoding="utf-8")
    monkeypatch.setenv("BONDS_LITE_MODEL_PYTHON", str(python))
    monkeypatch.setenv("SST_MACALC_DIR", str(model))

    config = stage_four.load_macalc_config(caller)

    assert config.python_executable == python.resolve()
    assert config.model_dir == model.resolve()


def test_load_macalc_config_defaults_model_folder_relative_to_the_draft(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    stage_four = _stage_four()
    caller = tmp_path / "work" / "Draft.xlsm"
    caller.parent.mkdir()
    caller.write_bytes(b"Draft")
    python = tmp_path / "py311" / "python.exe"
    python.parent.mkdir()
    python.write_bytes(b"fake interpreter")
    expected = (caller.parent / "../../00 Ad Hoc/MACalc_Batch").resolve()
    expected.mkdir(parents=True)
    (expected / "MACalcrunmanager.py").write_text("# fake manager\n", encoding="utf-8")
    monkeypatch.setenv("BONDS_LITE_MODEL_PYTHON", str(python))
    monkeypatch.delenv("SST_MACALC_DIR", raising=False)

    config = stage_four.load_macalc_config(caller)

    assert config.model_dir == expected


@pytest.mark.parametrize("setting", ["", "not a python executable"])
def test_load_macalc_config_rejects_missing_or_invalid_model_python_setting(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    setting: str,
) -> None:
    import bonds_lite_input as shared

    stage_four = _stage_four()
    caller = tmp_path / "Draft.xlsm"
    monkeypatch.setenv("BONDS_LITE_MODEL_PYTHON", setting)
    monkeypatch.setattr(shared, "_user_environment_value", lambda name: "")
    with pytest.raises(stage_four.StageFourError):
        stage_four.load_macalc_config(caller)


def test_load_macalc_config_requires_the_macalc_run_manager(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    stage_four = _stage_four()
    caller = tmp_path / "Draft.xlsm"
    python = tmp_path / "python.exe"
    python.write_bytes(b"fake interpreter")
    model = tmp_path / "model without entry"
    model.mkdir()
    monkeypatch.setenv("BONDS_LITE_MODEL_PYTHON", str(python))
    monkeypatch.setenv("SST_MACALC_DIR", str(model))

    with pytest.raises(stage_four.StageFourError, match="MACalcrunmanager.py"):
        stage_four.load_macalc_config(caller)


def test_preflight_macalc_environment_probes_python_311_from_model_directory(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    stage_four = _stage_four()
    python = tmp_path / "py311" / "python.exe"
    python.parent.mkdir()
    python.write_bytes(b"fake interpreter")
    model = tmp_path / "MACalc Batch"
    model.mkdir()
    calls: list[dict[str, Any]] = []

    def successful_probe(command: list[str], **kwargs: Any) -> Any:
        calls.append({"command": command, **kwargs})
        return SimpleNamespace(
            returncode=0, stdout="MACalc environment ready", stderr=""
        )

    monkeypatch.setattr(stage_four.subprocess, "run", successful_probe)
    config = SimpleNamespace(python_executable=python, model_dir=model, packaged=False)

    stage_four.preflight_macalc_environment(config)

    call = calls[0]
    assert call["command"][0] == str(python)
    probe = " ".join(call["command"])
    assert "MACalcrunmanager" in probe
    assert "3,11" in probe or "3.11" in probe
    assert call["cwd"] == str(model)


def _fake_macalc_model(model_dir: Path) -> None:
    """Write a tiny manager that records calls and creates the expected CSV."""
    model_dir.mkdir(parents=True, exist_ok=True)
    (model_dir / "MACalcrunmanager.py").write_text(
        """
import json
import os
from pathlib import Path

def find_workbooks(root):
    allowed = {".xls", ".xlsb", ".xlsm", ".xlsx"}
    return sorted(
        path for path in Path(root).rglob("*")
        if path.is_file()
        and not path.name.startswith("~$")
        and "_opt" not in path.stem
        and "macalc" in path.name.casefold()
        and path.suffix.casefold() in allowed
    )

def confirm_workbooks(workbooks):
    return input("fake manager prompt")

def main(**kwargs):
    root = Path(kwargs["root_repository"])
    workbooks = find_workbooks(root)
    with open(os.environ["MACALC_RECORD"], "w", encoding="utf-8") as stream:
        json.dump({"kwargs": {key: str(value) if isinstance(value, Path) else value
                              for key, value in kwargs.items()},
                   "workbooks": [str(path) for path in workbooks]}, stream)
    if os.environ.get("MACALC_MODE") == "failure":
        raise SystemExit(1)
    if not confirm_workbooks(workbooks):
        return []
    if os.environ.get("MACALC_MODE") == "empty result":
        return []
    workbook = workbooks[0]
    csv_path = workbook.with_name(workbook.stem + "_unoptimised_assets.csv")
    if os.environ.get("MACALC_MODE") != "missing csv":
        csv_path.write_text(
            "" if os.environ.get("MACALC_MODE") == "empty csv" else "ISIN,PV\\nX,1\\n",
            encoding="utf-8",
        )
    return [(workbook, None, csv_path, None)]
""",
        encoding="utf-8",
    )


@pytest.fixture
def fake_macalc_entry(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> Any:
    """Provide an isolated run manager, stress folder and audited workbook."""
    import macalc_batch_entry

    model_dir = tmp_path / "fake MACalc model"
    _fake_macalc_model(model_dir)
    root = tmp_path / "SST Output_Test" / "Credit"
    root.mkdir(parents=True)
    expected = root / "MACalc input_Test_0123456789abcdef.xlsb"
    expected.write_bytes(b"fake input")
    record = tmp_path / "manager calls.json"
    monkeypatch.setenv("MACALC_RECORD", str(record))
    monkeypatch.delenv("MACALC_MODE", raising=False)
    monkeypatch.syspath_prepend(str(model_dir))
    monkeypatch.delitem(sys.modules, "MACalcrunmanager", raising=False)
    yield macalc_batch_entry, model_dir, root, expected
    sys.modules.pop("MACalcrunmanager", None)


def _entry_arguments(
    model_dir: Path,
    root: Path,
    expected: Path,
    downgrade: bool = True,
) -> list[str]:
    return [
        "--model-dir",
        str(model_dir),
        "--root",
        str(root),
        "--expected",
        str(expected),
        "--downgrade-run" if downgrade else "--no-downgrade-run",
    ]


@pytest.mark.parametrize("downgrade", [True, False])
def test_macalc_entry_confirms_only_the_audited_workbook_and_passes_options(
    fake_macalc_entry: tuple[Any, Path, Path, Path],
    monkeypatch: pytest.MonkeyPatch,
    capsys: pytest.CaptureFixture[str],
    tmp_path: Path,
    downgrade: bool,
) -> None:
    import json

    entry, model_dir, root, expected = fake_macalc_entry
    record = tmp_path / "manager calls.json"
    monkeypatch.setenv("MACALC_RECORD", str(record))
    monkeypatch.setattr(
        "builtins.input", lambda *args: pytest.fail("MACalc confirmation was prompted")
    )

    result = entry.main(_entry_arguments(model_dir, root, expected, downgrade))

    assert result == 0
    observed = json.loads(record.read_text(encoding="utf-8"))
    kwargs = observed["kwargs"]
    assert kwargs["root_repository"] == str(root)
    assert kwargs["use_binary_optimisation"] is False
    assert kwargs["run_optimisation"] is False
    assert kwargs["is_downgrade_run"] is downgrade
    assert kwargs["floor_cod_at_zero"] is True
    assert observed["workbooks"] == [str(expected)]
    assert (root / f"{expected.stem}_unoptimised_assets.csv").stat().st_size > 0
    assert str(expected) in capsys.readouterr().out


def test_macalc_entry_stops_before_model_run_when_another_workbook_is_listed(
    fake_macalc_entry: tuple[Any, Path, Path, Path],
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    entry, model_dir, root, expected = fake_macalc_entry
    record = tmp_path / "manager calls.json"
    monkeypatch.setenv("MACALC_RECORD", str(record))
    (root / "MACalc input_other.xlsb").write_bytes(b"extra")

    result = entry.main(_entry_arguments(model_dir, root, expected))

    assert result == 2
    assert not record.exists()
    assert not (root / f"{expected.stem}_unoptimised_assets.csv").exists()


@pytest.mark.skipif(
    sys.platform != "win32", reason="Case-insensitive paths are Windows-specific"
)
def test_macalc_entry_treats_case_only_expected_path_difference_as_equal(
    fake_macalc_entry: tuple[Any, Path, Path, Path],
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    entry, model_dir, root, expected = fake_macalc_entry
    record = tmp_path / "manager calls.json"
    monkeypatch.setenv("MACALC_RECORD", str(record))

    assert (
        entry.main(_entry_arguments(model_dir, root, Path(str(expected).upper()))) == 0
    )


def test_macalc_entry_reports_no_candidates_without_running_the_model(
    fake_macalc_entry: tuple[Any, Path, Path, Path],
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    entry, model_dir, root, expected = fake_macalc_entry
    expected.unlink()
    record = tmp_path / "manager calls.json"
    monkeypatch.setenv("MACALC_RECORD", str(record))

    result = entry.main(_entry_arguments(model_dir, root, expected))

    assert result == 2
    assert not record.exists()


def test_macalc_entry_rejects_a_different_single_workbook_and_lists_it(
    fake_macalc_entry: tuple[Any, Path, Path, Path],
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
    capsys: pytest.CaptureFixture[str],
) -> None:
    entry, model_dir, root, expected = fake_macalc_entry
    expected.unlink()
    other = root / "MACalc input_other.xlsb"
    other.write_bytes(b"not the audited workbook")
    record = tmp_path / "manager calls.json"
    monkeypatch.setenv("MACALC_RECORD", str(record))

    assert entry.main(_entry_arguments(model_dir, root, expected)) == 2

    assert not record.exists()
    assert str(other) in capsys.readouterr().out


@pytest.mark.parametrize("mode", ["empty result", "missing csv", "empty csv"])
def test_macalc_entry_requires_one_nonempty_csv_result(
    fake_macalc_entry: tuple[Any, Path, Path, Path],
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
    mode: str,
) -> None:
    entry, model_dir, root, expected = fake_macalc_entry
    monkeypatch.setenv("MACALC_MODE", mode)
    monkeypatch.setenv("MACALC_RECORD", str(tmp_path / "manager calls.json"))

    assert entry.main(_entry_arguments(model_dir, root, expected)) == 2


def test_macalc_entry_propagates_model_failure_with_traceback(
    fake_macalc_entry: tuple[Any, Path, Path, Path],
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
    capsys: pytest.CaptureFixture[str],
) -> None:
    entry, model_dir, root, expected = fake_macalc_entry
    monkeypatch.setenv("MACALC_MODE", "failure")
    monkeypatch.setenv("MACALC_RECORD", str(tmp_path / "manager calls.json"))

    with pytest.raises(SystemExit) as caught:
        entry.main(_entry_arguments(model_dir, root, expected))

    assert caught.value.code == 1
    assert "Traceback" in capsys.readouterr().err


def test_macalc_entry_reports_git_provenance_and_hashes_of_six_model_sources(
    fake_macalc_entry: tuple[Any, Path, Path, Path],
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
    capsys: pytest.CaptureFixture[str],
) -> None:
    import hashlib
    import subprocess

    entry, model_dir, root, expected = fake_macalc_entry
    monkeypatch.setenv("MACALC_RECORD", str(tmp_path / "manager calls.json"))
    source_names = (
        "MACalcrunmanager.py",
        "MACalc.py",
        "MACE.py",
        "calcSCR.py",
        "calcmp.py",
        "MACalcutils.py",
    )
    hashes = {}
    for name in source_names:
        source = model_dir / name
        if name != "MACalcrunmanager.py":
            source.write_text(f"# provenance sentinel: {name}\n", encoding="utf-8")
        hashes[name] = hashlib.sha256(source.read_bytes()).hexdigest()

    def fake_git(command: list[str], **kwargs: object) -> Any:
        joined = " ".join(command)
        if "rev-parse" in joined:
            return subprocess.CompletedProcess(command, 0, "deadbeef\n", "")
        if "status" in joined:
            return subprocess.CompletedProcess(command, 0, " M MACalc.py\n", "")
        return subprocess.CompletedProcess(command, 0, "", "")

    monkeypatch.setattr(entry.subprocess, "run", fake_git)

    assert entry.main(_entry_arguments(model_dir, root, expected)) == 0

    output = capsys.readouterr().out
    assert "deadbeef" in output
    assert "modified" in output.casefold()
    assert all(value in output for value in hashes.values())


@pytest.mark.parametrize("solver_available", [True, False])
def test_macalc_entry_frozen_self_test_imports_bundled_runtime_without_excel(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    capsys: pytest.CaptureFixture[str],
    solver_available: bool,
) -> None:
    import builtins

    import macalc_batch_entry as entry

    executable = tmp_path / "MACalc" / "MACalcModel.exe"
    executable.parent.mkdir()
    executable.write_bytes(b"fake packaged entry")
    bundled = executable.parent / "_internal"
    bundled.mkdir()
    original_cwd = Path.cwd()
    original_import = builtins.__import__
    imports: list[str] = []
    solver_checks: list[Path] = []

    def forbidden_model_call(*args: Any, **kwargs: Any) -> None:
        pytest.fail("Self-test must not scan workbooks, run the model or open Excel")

    manager = SimpleNamespace(
        find_workbooks=forbidden_model_call, main=forbidden_model_call
    )

    def solver() -> Any:
        solver_checks.append(Path.cwd())
        return SimpleNamespace(available=lambda: solver_available)

    pulp = SimpleNamespace(PULP_CBC_CMD=solver)

    def import_bundled(name: str, *args: Any, **kwargs: Any) -> Any:
        if name in {"MACalcrunmanager", "pulp"}:
            imports.append(name)
            assert Path.cwd() == bundled
            assert str(bundled) in sys.path
            return manager if name == "MACalcrunmanager" else pulp
        if name.split(".")[0] in {"xlwings", "win32com"}:
            pytest.fail("Self-test must not open Excel")
        return original_import(name, *args, **kwargs)

    monkeypatch.setattr(entry.sys, "frozen", True, raising=False)
    monkeypatch.setattr(entry.sys, "_MEIPASS", str(bundled), raising=False)
    monkeypatch.setattr(entry.sys, "executable", str(executable))
    monkeypatch.setattr(sys, "path", sys.path.copy())
    monkeypatch.setattr(builtins, "__import__", import_bundled)
    monkeypatch.setattr(entry, "run", forbidden_model_call)
    monkeypatch.chdir(original_cwd)

    # A frozen executable must ignore a stale external source-folder override.
    result = entry.main(
        ["--self-test", "--model-dir", str(tmp_path / "unavailable external sources")]
    )

    assert result == (0 if solver_available else 1)
    assert {"MACalcrunmanager", "pulp"} <= set(imports)
    assert solver_checks == [bundled]
    captured = capsys.readouterr()
    if solver_available:
        assert "MACALC_SELF_TEST_OK" in captured.out
    else:
        assert "MACALC_SELF_TEST_OK" not in captured.out
        assert any(
            word in (captured.out + captured.err).casefold()
            for word in ("cbc", "solver")
        )
    assert Path.cwd() == original_cwd
    assert not list(bundled.glob("*.py"))


def test_macalc_entry_line_buffers_piped_stdout_before_exit() -> None:
    """Expose a plain MACalc print while the packaged model is still running."""
    import os
    import queue
    import subprocess
    import threading

    import macalc_batch_entry as entry

    environment = os.environ.copy()
    environment.pop("PYTHONUNBUFFERED", None)
    source_folder = str(Path(entry.__file__).resolve().parent)
    script = (
        "import sys; "
        f"sys.path.insert(0, {source_folder!r}); "
        "from macalc_batch_entry import _configure_console_streams; "
        "_configure_console_streams(); "
        "print('MACALC_LINE_BUFFERED'); "
        "sys.stdin.readline()"
    )
    child = subprocess.Popen(
        [sys.executable, "-c", script],
        stdin=subprocess.PIPE,
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        text=True,
        env=environment,
    )
    lines: queue.Queue[str] = queue.Queue()
    reader = threading.Thread(
        target=lambda: lines.put(child.stdout.readline()), daemon=True
    )
    reader.start()
    try:
        assert lines.get(timeout=5.0).strip() == "MACALC_LINE_BUFFERED"
        assert child.poll() is None
    finally:
        if child.poll() is None:
            child.stdin.write("continue\n")
            child.stdin.flush()
        child.wait(timeout=5.0)
        reader.join(timeout=1.0)
        child.stdin.close()
        child.stdout.close()


@pytest.mark.parametrize("model_fails", [False, True])
def test_macalc_entry_frozen_main_runs_in_bundle_and_restores_cwd(
    fake_macalc_entry: tuple[Any, Path, Path, Path],
    monkeypatch: pytest.MonkeyPatch,
    model_fails: bool,
) -> None:
    import importlib

    entry, model_dir, root, expected = fake_macalc_entry
    manager = importlib.import_module("MACalcrunmanager")
    (model_dir / "MACalcrunmanager.py").unlink()
    bundled = model_dir / "_internal"
    bundled.mkdir()
    original_cwd = Path.cwd()
    original_main = manager.main
    calls: list[Path] = []

    def model_main(**kwargs: Any) -> Any:
        calls.append(Path.cwd())
        assert Path.cwd() == bundled
        return original_main(**kwargs)

    def forbidden_provenance(*args: Any, **kwargs: Any) -> None:
        pytest.fail(
            "Frozen runtime provenance belongs in release.json, not git/source checks"
        )

    monkeypatch.setattr(manager, "main", model_main)
    monkeypatch.setattr(entry.sys, "frozen", True, raising=False)
    monkeypatch.setattr(entry.sys, "_MEIPASS", str(bundled), raising=False)
    monkeypatch.setattr(entry.sys, "executable", str(model_dir / "MACalcModel.exe"))
    monkeypatch.setattr(sys, "path", sys.path.copy())
    monkeypatch.setattr(entry, "_model_provenance", forbidden_provenance)
    monkeypatch.setattr(entry.subprocess, "run", forbidden_provenance)
    monkeypatch.setattr(entry.hashlib, "sha256", forbidden_provenance)
    monkeypatch.chdir(original_cwd)
    if model_fails:
        monkeypatch.setenv("MACALC_MODE", "failure")
    arguments = _entry_arguments(model_dir / "missing external source", root, expected)

    if model_fails:
        with pytest.raises(SystemExit) as caught:
            entry.main(arguments)
        assert caught.value.code == 1
    else:
        assert entry.main(arguments) == 0
    assert calls == [bundled]
    assert Path.cwd() == original_cwd
    assert not list(bundled.glob("*.py"))


def test_macalc_entry_frozen_run_needs_no_python_sources_or_runtime_provenance(
    fake_macalc_entry: tuple[Any, Path, Path, Path],
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    import importlib

    entry, model_dir, root, expected = fake_macalc_entry
    importlib.import_module("MACalcrunmanager")
    (model_dir / "MACalcrunmanager.py").unlink()

    def forbidden_provenance(*args: Any, **kwargs: Any) -> None:
        pytest.fail("Frozen runs must not query git or hash absent Python sources")

    monkeypatch.setattr(entry.sys, "frozen", True, raising=False)
    monkeypatch.setattr(sys, "path", sys.path.copy())
    monkeypatch.setattr(entry, "_model_provenance", forbidden_provenance)
    monkeypatch.setattr(entry.subprocess, "run", forbidden_provenance)
    monkeypatch.setattr(entry.hashlib, "sha256", forbidden_provenance)

    assert entry.run(model_dir, root, expected, True) == 0
    assert (root / f"{expected.stem}_unoptimised_assets.csv").is_file()


def test_packaged_macalc_config_resolves_sibling_application_and_rejects_missing(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    stage_four = _stage_four()
    controller = tmp_path / "release" / "Controller" / "BondsLiteController.exe"
    controller.parent.mkdir(parents=True)
    controller.write_bytes(b"controller")
    macalc_exe = tmp_path / "release" / "MACalc" / "MACalcModel.exe"
    macalc_exe.parent.mkdir()
    macalc_exe.write_bytes(b"MACalc model")
    monkeypatch.setattr(stage_four.sys, "frozen", True, raising=False)
    monkeypatch.setattr(stage_four.sys, "executable", str(controller))
    monkeypatch.setenv("SST_MACALC_DIR", str(tmp_path / "missing source tree"))
    monkeypatch.setattr(
        stage_four.shared,
        "resolve_model_python",
        lambda: pytest.fail("Frozen config must not resolve a user Python environment"),
    )

    config = stage_four.load_macalc_config(tmp_path / "Draft.xlsm")

    assert config.packaged is True
    assert config.python_executable == macalc_exe.resolve()
    assert config.model_dir == macalc_exe.parent.resolve()
    macalc_exe.unlink()
    with pytest.raises(
        stage_four.StageFourError, match=r"(?i)MACalc.*missing|missing.*MACalc"
    ):
        stage_four.load_macalc_config(tmp_path / "Draft.xlsm")


def test_packaged_macalc_config_resolves_shared_component(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    import hashlib
    import json

    stage_four = _stage_four()
    release_root = tmp_path / "release source"
    bundle = (
        release_root
        / "releases"
        / "shared-release"
        / "Bonds Lite App"
    )
    controller = bundle / "Controller" / "BondsLiteController.exe"
    controller.parent.mkdir(parents=True)
    controller.write_bytes(b"controller")
    macalc_fingerprint = "b" * 64
    component = (
        release_root
        / "components"
        / "MACalcModel"
        / macalc_fingerprint
    )
    app = component / "app"
    app.mkdir(parents=True)
    macalc = app / "MACalcModel.exe"
    macalc.write_bytes(b"MACalc model")
    macalc_hash = hashlib.sha256(macalc.read_bytes()).hexdigest()
    (component / "component.id").write_text(
        f"MACalcModel={macalc_fingerprint}", encoding="ascii"
    )
    (component / "component.json").write_text(
        json.dumps(
            {
                "name": "MACalcModel",
                "fingerprint": macalc_fingerprint,
                "executable_sha256": macalc_hash,
            }
        ),
        encoding="utf-8",
    )
    (bundle / "components.txt").write_text(
        f"model={'a' * 64}\nmacalc={macalc_fingerprint}\n",
        encoding="ascii",
    )
    model_fingerprint = "a" * 64
    model_component = (
        release_root
        / "components"
        / "BondsLiteModel"
        / model_fingerprint
    )
    model_app = model_component / "app"
    model_app.mkdir(parents=True)
    model_exe = model_app / "BondsLiteModel.exe"
    model_exe.write_bytes(b"model")
    model_hash = hashlib.sha256(model_exe.read_bytes()).hexdigest()
    (model_component / "component.id").write_text(
        f"BondsLiteModel={model_fingerprint}", encoding="ascii"
    )
    (model_component / "component.json").write_text(
        json.dumps(
            {
                "name": "BondsLiteModel",
                "fingerprint": model_fingerprint,
                "executable_sha256": model_hash,
            }
        ),
        encoding="utf-8",
    )
    (bundle / "release.json").write_text(
        json.dumps(
            {
                "release_id": "shared-release",
                "layout_version": 2,
                "component_fingerprints": {
                    "BondsLiteModel": model_fingerprint,
                    "MACalcModel": macalc_fingerprint,
                },
                "model_executable_sha256": model_hash,
                "macalc_executable_sha256": macalc_hash,
            }
        ),
        encoding="utf-8",
    )
    monkeypatch.setattr(stage_four.sys, "frozen", True, raising=False)
    monkeypatch.setattr(stage_four.sys, "executable", str(controller))

    config = stage_four.load_macalc_config(tmp_path / "Draft.xlsm")

    assert config.python_executable == macalc.resolve()
    assert config.model_dir == app.resolve()
    assert config.packaged is True


def test_stage_four_audit_names_are_confined_to_stage_four_fields() -> None:
    stage_four = _stage_four()

    assert set(stage_four.STAGE_FOUR_FIELD_NAMES.values()) == {
        f"SST_Stage4_{suffix}"
        for suffix in (
            "LastRun",
            "Start",
            "Finish",
            "Elapsed",
            "User",
            "Status",
            "InputRun",
            "LastSuccessfulOutput",
        )
    }
    assert stage_four.STAGE_FOUR_PANEL_NAME == "SST_Stage4_Panel"


def test_write_stage_four_audit_rejects_unknown_fields_without_writing(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    stage_four = _stage_four()
    writes: list[object] = []
    cell = SimpleNamespace(value="Not Run")
    caller = SimpleNamespace(
        app=SimpleNamespace(api=SimpleNamespace(EnableEvents=True)),
        api=SimpleNamespace(Saved=True),
        save=lambda: writes.append("save"),
    )
    monkeypatch.setattr(stage_four, "require_stage_four_panel", lambda book: None)
    monkeypatch.setattr(
        stage_four,
        "_stage_four_cells",
        lambda book: {"status": cell},
        raising=False,
    )

    with pytest.raises(stage_four.StageFourError):
        stage_four.write_stage_four_audit(caller, {"stage_three_status": "Succeeded"})

    assert cell.value == "Not Run"
    assert writes == []
    assert caller.app.api.EnableEvents is True


def test_write_stage_four_audit_requires_the_draft_to_save_each_change(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    stage_four = _stage_four()
    cell = SimpleNamespace(value="Not Run")
    caller = SimpleNamespace(
        app=SimpleNamespace(api=SimpleNamespace(EnableEvents=True)),
        api=SimpleNamespace(Saved=False),
        save=lambda: None,
    )
    monkeypatch.setattr(stage_four, "require_stage_four_panel", lambda book: None)
    monkeypatch.setattr(
        stage_four,
        "_stage_four_cells",
        lambda book: {"status": cell},
        raising=False,
    )

    with pytest.raises(stage_four.StageFourError):
        stage_four.write_stage_four_audit(caller, {"status": "Running"})

    assert cell.value == "Running"
    assert caller.app.api.EnableEvents is True


def test_missing_stage_four_panel_stops_with_installation_guidance() -> None:
    stage_four = _stage_four()

    with pytest.raises(
        stage_four.StageFourError, match=r"(?i)install the Stage 4 controls"
    ):
        stage_four.require_stage_four_panel(SimpleNamespace())


def _stage_four_run_case(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    *,
    child_returncode: int = 0,
    downgrade: bool = True,
    interrupt: bool = False,
) -> tuple[Any, Any, Path, Path, Path, list[dict[str, object]], list[dict[str, Any]]]:
    """Stub only Excel and process boundaries while retaining real backup I/O."""
    import contextlib
    import bonds_lite_input as shared

    stage_four = _stage_four()
    caller_path = tmp_path / "Draft.xlsm"
    caller_path.write_bytes(b"fake Draft")
    workbook = (
        tmp_path
        / "SST Output_Test"
        / "Credit"
        / ("MACalc input_Test_0123456789abcdef.xlsb")
    )
    workbook.parent.mkdir(parents=True)
    workbook.write_bytes(b"input before MACalc")
    model_dir = tmp_path / "MACalc Batch"
    model_dir.mkdir()
    python = tmp_path / "py311" / "python.exe"
    python.parent.mkdir()
    python.write_bytes(b"fake interpreter")
    csv_path = workbook.with_name(workbook.stem + "_unoptimised_assets.csv")
    stage_input = SimpleNamespace(
        run_folder=workbook.parent,
        workbook=workbook,
        is_downgrade_run=downgrade,
    )
    config = SimpleNamespace(
        python_executable=python,
        model_dir=model_dir,
        packaged=False,
    )
    caller = SimpleNamespace(
        fullname=str(caller_path),
        api=SimpleNamespace(Saved=True, ReadOnly=False),
        save=lambda: None,
    )
    audit_writes: list[dict[str, object]] = []
    process_calls: list[dict[str, Any]] = []
    guard_calls: list[tuple[Path, str | None]] = []

    monkeypatch.setattr(stage_four, "require_stage_four_panel", lambda book: None)
    monkeypatch.setattr(
        stage_four, "resolve_stage_four_input", lambda book: stage_input
    )
    monkeypatch.setattr(stage_four, "load_macalc_config", lambda path: config)
    monkeypatch.setattr(stage_four, "preflight_macalc_environment", lambda loaded: None)
    monkeypatch.setattr(
        stage_four,
        "write_stage_four_audit",
        lambda book, fields: audit_writes.append(dict(fields)),
    )
    monkeypatch.setattr(
        stage_four,
        "_read_stage_four_fields",
        lambda book: {"last_output": "SST Output_Test/Credit/previous.csv"},
        raising=False,
    )
    monkeypatch.setattr(shared, "resolve_caller_path", lambda book: caller_path)
    monkeypatch.setattr(shared, "_caller_path", lambda book: caller_path)

    def run_guard(path: Path, *, conflict_message: str | None = None) -> Any:
        guard_calls.append((Path(path), conflict_message))
        return contextlib.nullcontext()

    monkeypatch.setattr(shared, "workbook_run_guard", run_guard)
    monkeypatch.setattr(
        shared,
        "portable_output_path",
        lambda path, output: output.relative_to(caller_path.parent).as_posix(),
    )
    csv_path.write_text("ISIN,PV\nX,1\n", encoding="utf-8")

    class Process:
        def __init__(self, *args: Any, **kwargs: Any) -> None:
            process_calls.append({"command": args[0], **kwargs})
            self.returncode = child_returncode
            self.stdout = StringIO("fake MACalc output\n")

        def wait(self) -> int:
            if interrupt:
                raise KeyboardInterrupt
            if self.returncode == 0:
                csv_path.write_text("ISIN,PV\nX,2\n", encoding="utf-8")
            return self.returncode

        def poll(self) -> int:
            return self.returncode

        def __enter__(self) -> "Process":
            return self

        def __exit__(self, *args: object) -> None:
            return None

    monkeypatch.setattr(stage_four.subprocess, "Popen", Process)
    caller.guard_calls = guard_calls
    return (
        stage_four,
        caller,
        caller_path,
        workbook,
        csv_path,
        audit_writes,
        process_calls,
    )


@pytest.mark.parametrize("downgrade", [True, False])
def test_run_stage_four_backups_input_launches_in_model_folder_and_audits_success(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    downgrade: bool,
) -> None:
    (
        stage_four,
        caller,
        caller_path,
        workbook,
        csv_path,
        writes,
        processes,
    ) = _stage_four_run_case(tmp_path, monkeypatch, downgrade=downgrade)

    assert stage_four.run_stage_four(caller) == csv_path

    backup = caller_path.parent / "SST Output_Test" / "Stage 4 backup" / workbook.name
    assert backup.read_bytes() == b"input before MACalc"
    assert writes[0]["status"] == "Running"
    assert writes[0]["last_output"] == "SST Output_Test/Credit/previous.csv"
    assert writes[0]["input_run"] == (
        "SST Output_Test/Credit/MACalc input_Test_0123456789abcdef.xlsb"
    )
    assert writes[-1]["status"] == "Succeeded"
    assert writes[-1]["last_output"] == (
        "SST Output_Test/Credit/MACalc input_Test_0123456789abcdef"
        "_unoptimised_assets.csv"
    )
    call = processes[0]
    command = call["command"]
    expected_flag = "--downgrade-run" if downgrade else "--no-downgrade-run"
    assert command == [
        str(tmp_path / "py311" / "python.exe"),
        "-B",
        str(Path(stage_four.__file__).with_name("macalc_batch_entry.py")),
        "--model-dir",
        str(tmp_path / "MACalc Batch"),
        "--root",
        str(workbook.parent),
        "--expected",
        str(workbook),
        expected_flag,
    ]
    assert call["cwd"] == str(tmp_path / "MACalc Batch")
    assert call["shell"] is False
    assert call["env"]["PYTHONDONTWRITEBYTECODE"] == "1"
    assert "timeout" not in call
    assert caller.guard_calls == [
        (caller_path, None),
        (
            workbook.parent.parent / "Stage 4 backup",
            "A Stage 4 run is already in progress for this output folder",
        ),
    ]


@pytest.mark.parametrize("log_level", [logging.INFO, logging.DEBUG])
def test_run_stage_four_child_command_is_debug_only(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    caplog: pytest.LogCaptureFixture,
    log_level: int,
) -> None:
    stage_four, caller, _, _, csv_path, _, processes = _stage_four_run_case(
        tmp_path, monkeypatch
    )
    monkeypatch.setattr(stage_four.logger, "handlers", [caplog.handler])
    monkeypatch.setattr(stage_four.logger, "propagate", False)

    with caplog.at_level(log_level, logger=stage_four.logger.name):
        assert stage_four.run_stage_four(caller) == csv_path

    assert "MACalc options:" in caplog.text
    commands = [
        record
        for record in caplog.records
        if record.getMessage().startswith("MACalc child command:")
    ]
    if log_level == logging.INFO:
        assert commands == []
    else:
        assert len(commands) == 1
        assert commands[0].levelno == logging.DEBUG
        assert commands[0].getMessage() == (
            "MACalc child command: "
            + stage_four.subprocess.list2cmdline(processes[0]["command"])
        )


def test_run_stage_four_replaces_fixed_backup_on_rerun(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    stage_four, caller, caller_path, workbook, _, _, _ = _stage_four_run_case(
        tmp_path, monkeypatch
    )
    backup = caller_path.parent / "SST Output_Test" / "Stage 4 backup" / workbook.name

    stage_four.run_stage_four(caller)
    workbook.write_bytes(b"second input state")
    stage_four.run_stage_four(caller)

    assert backup.read_bytes() == b"second input state"
    assert list(backup.parent.iterdir()) == [backup]


def test_backup_stage_four_input_retains_only_the_latest_backup(
    tmp_path: Path,
) -> None:
    stage_four = _stage_four()
    run_folder = tmp_path / "SST Output_Test" / "Credit"
    run_folder.mkdir(parents=True)
    old_workbook = run_folder / "MACalc input_Test_0123456789abcdef.xlsb"
    old_workbook.write_bytes(b"old input")
    new_workbook = run_folder / "MACalc input_Test_fedcba9876543210.xlsb"
    new_workbook.write_bytes(b"new input")
    backup_folder = run_folder.parent / "Stage 4 backup"
    unrelated = backup_folder / "notes.txt"
    backup_folder.mkdir()
    unrelated.write_text("not a Stage 4 backup", encoding="utf-8")
    old_backup = backup_folder / old_workbook.name
    old_backup.write_bytes(b"old backup")

    latest = stage_four.backup_stage_four_input(new_workbook)

    assert latest.read_bytes() == b"new input"
    assert sorted(path.name for path in backup_folder.iterdir()) == [
        latest.name,
        unrelated.name,
    ]


def test_backup_stage_four_input_reports_failure_to_remove_old_backup(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    stage_four = _stage_four()
    run_folder = tmp_path / "SST Output_Test" / "Credit"
    run_folder.mkdir(parents=True)
    old_workbook = run_folder / "MACalc input_Test_0123456789abcdef.xlsb"
    old_workbook.write_bytes(b"old input")
    backup_folder = run_folder.parent / "Stage 4 backup"
    backup_folder.mkdir()
    old_backup = backup_folder / old_workbook.name
    old_backup.write_bytes(b"old backup")
    new_workbook = run_folder / "MACalc input_Test_fedcba9876543210.xlsb"
    new_workbook.write_bytes(b"new input")
    real_unlink = stage_four.os.unlink

    def deny_old_backup(path: str | Path, *args: Any, **kwargs: Any) -> None:
        if Path(path) == old_backup:
            raise OSError("backup is locked")
        real_unlink(path, *args, **kwargs)

    monkeypatch.setattr(stage_four.os, "unlink", deny_old_backup)

    with pytest.raises(stage_four.StageFourError, match="Could not remove previous"):
        stage_four.backup_stage_four_input(new_workbook)

    assert old_backup.read_bytes() == b"old backup"
    assert (backup_folder / new_workbook.name).read_bytes() == b"new input"


def test_run_stage_four_does_not_launch_when_backup_creation_fails(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    stage_four, caller, _, _, _, _, processes = _stage_four_run_case(
        tmp_path, monkeypatch
    )

    def fail_copy(*args: object, **kwargs: object) -> None:
        raise OSError("backup denied")

    monkeypatch.setattr(stage_four.shutil, "copy2", fail_copy)
    with pytest.raises(stage_four.StageFourError):
        stage_four.run_stage_four(caller)

    assert processes == []


@pytest.mark.parametrize("exit_code", [1, 2])
def test_run_stage_four_records_failure_and_retains_previous_successful_output(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    exit_code: int,
) -> None:
    stage_four, caller, _, _, _, writes, _ = _stage_four_run_case(
        tmp_path, monkeypatch, child_returncode=exit_code
    )

    with pytest.raises(stage_four.StageFourError, match=str(exit_code)):
        stage_four.run_stage_four(caller)

    assert writes[0]["status"] == "Running"
    assert writes[-1]["status"] == "Failed"
    assert writes[-1]["last_output"] == "SST Output_Test/Credit/previous.csv"


def test_run_stage_four_records_interrupted_attempt_and_keeps_previous_output(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    stage_four, caller, _, _, _, writes, _ = _stage_four_run_case(
        tmp_path, monkeypatch, interrupt=True
    )

    with pytest.raises(KeyboardInterrupt):
        stage_four.run_stage_four(caller)

    assert writes[0]["status"] == "Running"
    assert writes[-1]["status"] == "Interrupted"
    assert writes[-1]["last_output"] == "SST Output_Test/Credit/previous.csv"


def test_run_stage_four_stale_handoff_stops_before_audit_write(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    stage_four, caller, _, _, _, writes, processes = _stage_four_run_case(
        tmp_path, monkeypatch
    )

    def stale(book: object) -> object:
        raise stage_four.StageFourError("Stage 3 is stale; rerun Stage 3")

    monkeypatch.setattr(stage_four, "resolve_stage_four_input", stale)
    with pytest.raises(stage_four.StageFourError, match="rerun Stage 3"):
        stage_four.run_stage_four(caller)

    assert writes == []
    assert processes == []


def test_install_stage_four_controls_refuses_a_draft_open_in_excel(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    import xlwings as xw

    stage_four = _stage_four()
    draft = tmp_path / "Draft.xlsm"
    draft.write_bytes(b"fake Draft")
    opened = SimpleNamespace(name="Draft.xlsm")
    app = SimpleNamespace(books=[opened])
    monkeypatch.setattr(xw, "apps", [app])

    with pytest.raises(stage_four.StageFourError, match=r"(?i)close .* in Excel"):
        stage_four.install_stage_four_controls(draft)

    assert not (tmp_path / "references" / "backups").exists()


def test_install_stage_four_controls_backs_up_before_requiring_stage_three_panel(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    import xlwings as xw
    import bonds_lite_input as shared
    import sst_stage_three as stage_three

    stage_four = _stage_four()
    draft = tmp_path / "Draft.xlsm"
    draft.write_bytes(b"fake Draft")

    class OpenBook:
        name = "Draft.xlsm"

        def __init__(self) -> None:
            self.api = SimpleNamespace(ReadOnly=False)
            self.closed = False

        def close(self, *args: object, **kwargs: object) -> None:
            self.closed = True

    book = OpenBook()
    application = SimpleNamespace(
        api=SimpleNamespace(
            AutomationSecurity=None,
            EnableEvents=True,
            Calculation=None,
            Quit=lambda: None,
        ),
        display_alerts=True,
        books=SimpleNamespace(open=lambda *args, **kwargs: book),
    )
    monkeypatch.setattr(xw, "apps", [])
    monkeypatch.setattr(xw, "App", lambda **kwargs: application)
    monkeypatch.setattr(shared, "require_stage_two_panel", lambda caller: None)
    monkeypatch.setattr(
        stage_three,
        "require_stage_three_panel",
        lambda caller: (_ for _ in ()).throw(
            stage_four.StageFourError("Stage 3 panel is required")
        ),
    )
    monkeypatch.setattr(
        stage_four,
        "require_stage_three_panel",
        stage_three.require_stage_three_panel,
        raising=False,
    )

    with pytest.raises(stage_four.StageFourError, match="Stage 3 panel"):
        stage_four.install_stage_four_controls(draft)

    backups = list((tmp_path / "references" / "backups").glob("Draft_*.xlsm"))
    assert len(backups) == 1
    assert backups[0].read_bytes() == b"fake Draft"
    assert book.closed


def _stage_four_panel_install_book(
    stage_four: Any,
    counta: int = 0,
    merge_cells: bool = False,
    shapes: list[Any] | None = None,
) -> Any:
    import sst_stage_three as stage_three

    panel = SimpleNamespace(
        row=40,
        column=8,
        rows=SimpleNamespace(count=11),
        left=100,
    )
    names = [
        SimpleNamespace(
            name=stage_three.STAGE_THREE_PANEL_NAME,
            refers_to_range=panel,
        )
    ]

    class Names(list):
        def __getitem__(self, key: object) -> Any:
            if isinstance(key, str):
                return next(name for name in self if name.name == key)
            return super().__getitem__(key)

    area = SimpleNamespace(
        api=SimpleNamespace(MergeCells=merge_cells),
        left=100,
        top=100,
        width=100,
        height=100,
    )
    sheet = SimpleNamespace(
        api=SimpleNamespace(Shapes=shapes or []),
        range=lambda *args: area,
    )
    book = SimpleNamespace(
        names=Names(names),
        sheets={"I. Control": sheet},
        app=SimpleNamespace(
            api=SimpleNamespace(
                EnableEvents=True,
                WorksheetFunction=SimpleNamespace(CountA=lambda target: counta),
            )
        ),
    )
    return book


@pytest.mark.parametrize(
    ("counta", "merge_cells"),
    [(1, False), (0, True)],
    ids=["occupied", "merged"],
)
def test_install_stage_four_panel_rejects_occupied_or_merged_area(
    monkeypatch: pytest.MonkeyPatch,
    counta: int,
    merge_cells: bool,
) -> None:
    import sst_stage_three as stage_three

    stage_four = _stage_four()
    monkeypatch.setattr(stage_three, "require_stage_three_panel", lambda book: None)
    book = _stage_four_panel_install_book(
        stage_four, counta=counta, merge_cells=merge_cells
    )

    with pytest.raises(stage_four.StageFourError, match="area is occupied"):
        stage_four._install_stage_four_panel(book)


def test_install_stage_four_panel_rejects_overlapping_shape(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    import sst_stage_three as stage_three

    stage_four = _stage_four()
    monkeypatch.setattr(stage_three, "require_stage_three_panel", lambda book: None)
    shape = SimpleNamespace(Name="Existing", Left=110, Width=10, Top=110, Height=10)
    book = _stage_four_panel_install_book(stage_four, shapes=[shape])

    with pytest.raises(stage_four.StageFourError, match="overlaps a shape"):
        stage_four._install_stage_four_panel(book)


def test_install_stage_four_panel_rejects_incomplete_names(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    import sst_stage_three as stage_three

    stage_four = _stage_four()
    monkeypatch.setattr(stage_three, "require_stage_three_panel", lambda book: None)
    book = _stage_four_panel_install_book(stage_four)
    book.names.append(SimpleNamespace(name=stage_four.STAGE_FOUR_PANEL_NAME))

    with pytest.raises(stage_four.StageFourError, match="names are incomplete"):
        stage_four._install_stage_four_panel(book)


def test_install_stage_four_panel_validates_a_complete_existing_install(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    import sst_stage_three as stage_three

    stage_four = _stage_four()
    monkeypatch.setattr(stage_three, "require_stage_three_panel", lambda book: None)
    book = _stage_four_panel_install_book(stage_four)
    book.names.extend(
        SimpleNamespace(name=name)
        for name in (
            stage_four.STAGE_FOUR_PANEL_NAME,
            *stage_four.STAGE_FOUR_FIELD_NAMES.values(),
        )
    )
    validated: list[Any] = []
    monkeypatch.setattr(
        stage_four,
        "require_stage_four_panel",
        lambda current: validated.append(current),
    )

    stage_four._install_stage_four_panel(book)

    assert validated == [book]


def test_install_stage_four_button_rejects_overlapping_shape() -> None:
    stage_four = _stage_four()
    anchor = SimpleNamespace(
        sheet=None,
        width=100,
        left=100,
        top=100,
        height=40,
        row=20,
        column=4,
        rows=SimpleNamespace(count=3),
    )
    shape = SimpleNamespace(Name="Existing", Left=110, Width=10, Top=110, Height=10)
    covered = SimpleNamespace(api=SimpleNamespace(MergeCells=False))
    sheet = SimpleNamespace(
        api=SimpleNamespace(Shapes=[shape]),
        range=lambda *args: (SimpleNamespace(left=400) if len(args) == 1 else covered),
    )
    anchor.sheet = sheet
    book = SimpleNamespace(
        app=SimpleNamespace(
            api=SimpleNamespace(
                WorksheetFunction=SimpleNamespace(CountA=lambda target: 0)
            )
        )
    )

    with pytest.raises(stage_four.StageFourError, match="overlaps an existing shape"):
        stage_four._install_button(book, anchor)


def test_stage_four_installer_inserts_run_macalc_macro_only_once(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    import xlwings as xw
    import sst_stage_three as stage_three

    stage_four = _stage_four()
    draft = tmp_path / "Draft.xlsm"
    draft.write_bytes(b"fake Draft")

    class CodeModule:
        def __init__(self) -> None:
            self.text = (
                "Public Sub GenerateBondsLiteInput()\n    LaunchBondsLite 1\nEnd Sub\n"
                "Public Sub RunBondsLite()\n    LaunchBondsLite 2\nEnd Sub\n"
                "Public Sub PrepareMACalc()\n    LaunchBondsLite 3\nEnd Sub\n"
            )

        @property
        def CountOfLines(self) -> int:
            return len(self.text.splitlines())

        def Lines(self, start: int, count: int) -> str:
            return "\n".join(self.text.splitlines()[start - 1 : start + count - 1])

        def InsertLines(self, line: int, text: str) -> None:
            self.text += text

    code = CodeModule()
    component = SimpleNamespace(CodeModule=code)
    panel = SimpleNamespace(
        row=90,
        column=8,
        rows=SimpleNamespace(count=11),
        left=500,
    )
    names = [
        SimpleNamespace(name=name, refers_to_range=panel)
        for name in (
            stage_four.STAGE_FOUR_PANEL_NAME,
            *stage_four.STAGE_FOUR_FIELD_NAMES.values(),
        )
    ]

    class Names(list):
        def __getitem__(self, key: object) -> Any:
            if isinstance(key, str):
                return next(name for name in self if name.name == key)
            return super().__getitem__(key)

    sheet = SimpleNamespace(
        range=lambda *args: SimpleNamespace(left=400),
    )
    book = SimpleNamespace(
        name="Draft.xlsm",
        names=Names(names),
        sheets={"I. Control": sheet},
        api=SimpleNamespace(
            ReadOnly=False,
            VBProject=SimpleNamespace(
                VBComponents=SimpleNamespace(Item=lambda name: component)
            ),
        ),
        close=lambda *args, **kwargs: None,
        save=lambda: None,
    )
    application = SimpleNamespace(
        api=SimpleNamespace(
            AutomationSecurity=None,
            EnableEvents=True,
            Calculation=None,
            Quit=lambda: None,
        ),
        display_alerts=True,
        books=SimpleNamespace(open=lambda *args, **kwargs: book),
    )
    monkeypatch.setattr(xw, "apps", [])
    monkeypatch.setattr(xw, "App", lambda **kwargs: application)
    monkeypatch.setattr(stage_three, "require_stage_three_panel", lambda caller: None)
    monkeypatch.setattr(
        stage_four,
        "require_stage_three_panel",
        stage_three.require_stage_three_panel,
        raising=False,
    )
    monkeypatch.setattr(stage_four, "_install_stage_four_panel", lambda caller: None)
    monkeypatch.setattr(stage_four, "require_stage_four_panel", lambda caller: None)
    monkeypatch.setattr(stage_four, "_stage_four_button_row", lambda panel, left: 90)
    monkeypatch.setattr(stage_four, "_install_button", lambda *args: None)

    stage_four.install_stage_four_controls(draft)
    stage_four.install_stage_four_controls(draft)

    assert code.text.count("Public Sub RunMACalc()") == 1
    assert code.text.count("LaunchBondsLite 4") == 1
