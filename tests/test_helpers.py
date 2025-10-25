"""Tests for helper utilities defined in ClassroomTools prior to the Qt imports."""

from __future__ import annotations

import ast
import contextlib
import os
import sys
import tempfile
import types
from pathlib import Path
from typing import Iterable, List, Optional, Set, Tuple


def _load_helper_module() -> types.ModuleType:
    """Load selected helper utilities from ``ClassroomTools`` using the AST."""

    module = types.ModuleType("ctools_helpers")
    path = Path(__file__).resolve().parents[1] / "ClassroomTools.py"
    source = path.read_text(encoding="utf-8").lstrip("\ufeff")
    tree = ast.parse(source, filename=str(path))

    namespace = module.__dict__
    namespace.update(
        {
            "__file__": str(path),
            "__name__": module.__name__,
            "os": os,
            "sys": sys,
            "contextlib": contextlib,
            "tempfile": tempfile,
            "Optional": Optional,
            "Tuple": Tuple,
            "List": List,
            "Iterable": Iterable,
            "Set": Set,
        }
    )
    sys.modules[module.__name__] = module

    targets = {
        "_ensure_directory",
        "_ensure_writable_directory",
        "_preferred_app_directory",
        "_choose_writable_target",
        "str_to_bool",
    }
    for node in tree.body:
        if isinstance(node, ast.FunctionDef) and node.name in targets:
            targets.remove(node.name)
            submodule = ast.Module(body=[node], type_ignores=[])
            ast.fix_missing_locations(submodule)
            exec(compile(submodule, str(path), "exec"), namespace)
        if not targets:
            break
    missing = sorted(targets)
    if missing:  # pragma: no cover - protects against refactors
        raise RuntimeError(f"Failed to load helper functions: {missing}")
    return module


helpers = _load_helper_module()


def test_str_to_bool_recognises_common_values() -> None:
    assert helpers.str_to_bool("True") is True
    assert helpers.str_to_bool("false") is False
    assert helpers.str_to_bool("YES") is True
    assert helpers.str_to_bool("off") is False
    assert helpers.str_to_bool("1") is True
    assert helpers.str_to_bool("0") is False


def test_str_to_bool_respects_default_for_unknown_values() -> None:
    assert helpers.str_to_bool("maybe", default=True) is True
    assert helpers.str_to_bool("", default=False) is False


def test_str_to_bool_handles_non_string_inputs() -> None:
    assert helpers.str_to_bool(True) is True
    assert helpers.str_to_bool(False) is False
    assert helpers.str_to_bool(3) is True
    assert helpers.str_to_bool(0) is False


def test_choose_writable_target_falls_back_when_parent_is_file(tmp_path: Path) -> None:
    blocker = tmp_path / "blocker"
    blocker.write_text("cannot use as directory")
    candidate = blocker / "students.xlsx"
    result = helpers._choose_writable_target(  # type: ignore[attr-defined]
        (str(candidate),),
        is_dir=False,
        fallback_name="students.xlsx",
    )
    assert Path(result).name == "students.xlsx"
    assert Path(result).parent != blocker
