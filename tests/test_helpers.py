"""Tests for helper utilities defined in ClassroomTools prior to the Qt imports."""

from __future__ import annotations

import ast
import enum
import math
import contextlib
import os
import sys
import tempfile
import types
from functools import singledispatch
from pathlib import Path
from typing import Iterable, List, Mapping, Optional, Set, Tuple


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
            "Mapping": Mapping,
            "Enum": enum.Enum,
            "math": math,
            "singledispatch": singledispatch,
        }
    )
    sys.modules[module.__name__] = module

    targets = {
        "_TRUE_STRINGS",
        "_FALSE_STRINGS",
        "_BooleanParseResult",
        "_coerce_bool",
        "parse_bool",
        "_ensure_directory",
        "_ensure_writable_directory",
        "_preferred_app_directory",
        "_choose_writable_target",
        "str_to_bool",
        "_resolve_word_control_conflict",
    }
    def _should_include_function(node: ast.FunctionDef) -> bool:
        if node.name in targets:
            return True
        for decorator in node.decorator_list:
            candidate = decorator
            if isinstance(candidate, ast.Call):
                candidate = candidate.func
            if (
                isinstance(candidate, ast.Attribute)
                and isinstance(candidate.value, ast.Name)
                and candidate.value.id == "_coerce_bool"
                and candidate.attr == "register"
            ):
                return True
        return False

    for node in tree.body:
        if isinstance(node, ast.FunctionDef) and _should_include_function(node):
            targets.discard(node.name)
            submodule = ast.Module(body=[node], type_ignores=[])
            ast.fix_missing_locations(submodule)
            exec(compile(submodule, str(path), "exec"), namespace)
        elif isinstance(node, ast.ClassDef) and node.name in targets:
            targets.remove(node.name)
            submodule = ast.Module(body=[node], type_ignores=[])
            ast.fix_missing_locations(submodule)
            exec(compile(submodule, str(path), "exec"), namespace)
        elif isinstance(node, ast.Assign):
            constant_names = {
                target.id
                for target in node.targets
                if isinstance(target, ast.Name) and target.id in targets
            }
            if constant_names:
                for name in constant_names:
                    targets.remove(name)
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


def test_str_to_bool_handles_bytes_and_enums() -> None:
    class _SampleEnum(enum.Enum):
        ENABLED = "yes"
        DISABLED = "no"

    assert helpers.str_to_bool(b"YES") is True
    assert helpers.str_to_bool(b"off") is False
    assert helpers.str_to_bool(_SampleEnum.ENABLED) is True
    assert helpers.str_to_bool(_SampleEnum.DISABLED) is False


def test_str_to_bool_interprets_numeric_strings() -> None:
    assert helpers.str_to_bool(" 2 ") is True
    assert helpers.str_to_bool("-1") is True
    assert helpers.str_to_bool("0.0") is False
    assert helpers.str_to_bool("+0") is False


def test_str_to_bool_ignores_nan_strings() -> None:
    assert helpers.str_to_bool("nan", default=True) is True
    assert helpers.str_to_bool("nan", default=False) is False


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


def test_resolve_word_control_conflict_prefers_last_toggle() -> None:
    resolve = helpers._resolve_word_control_conflict  # type: ignore[attr-defined]
    assert resolve(True, True, source_flags={"_last_word_toggle": "ms_word"}) == (True, False)
    assert resolve(True, True, source_flags={"_last_word_toggle": "wps_word"}) == (False, True)


def test_resolve_word_control_conflict_uses_previous_when_no_last_toggle() -> None:
    resolve = helpers._resolve_word_control_conflict  # type: ignore[attr-defined]
    assert resolve(
        True,
        True,
        previous_flags={"ms_word": True, "wps_word": False},
    ) == (True, False)
    assert resolve(
        True,
        True,
        previous_flags={"ms_word": False, "wps_word": True},
    ) == (False, True)


def test_resolve_word_control_conflict_defaults_to_ms_word() -> None:
    resolve = helpers._resolve_word_control_conflict  # type: ignore[attr-defined]
    assert resolve(True, True) == (True, False)
    assert resolve(True, False) == (True, False)
    assert resolve(False, True) == (False, True)
