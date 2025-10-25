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
            "win32gui": None,
            "_user32_top_level_hwnd": lambda hwnd: 0,
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
        "_compute_presentation_category",
        "_PresentationWindowMixin",
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


_WPS_WRITER_CLASSES = {
    "kwpsframeclass",
    "kwpsmainframe",
    "wpsframeclass",
    "wpsmainframe",
    "kwpsdocview",
    "wpsdocview",
}


def _stub_has_wps_presentation_signature(name: str) -> bool:
    if not name:
        return False
    lowered = name.lower()
    if lowered.startswith("kwpp") or "kwpp" in lowered:
        return True
    if lowered.startswith("wpsshow") or "wpsshow" in lowered:
        return True
    return False


def _stub_is_wps_slideshow_class(name: str) -> bool:
    lowered = name.lower()
    return lowered in {
        "kwppshowframeclass",
        "kwppshowframe",
        "kwppshowwndclass",
        "kwpsshowframe",
        "wpsshowframe",
    }


def _stub_has_wps_writer_signature(name: str) -> bool:
    return name.lower() in _WPS_WRITER_CLASSES


def _stub_is_word_like_class(name: str) -> bool:
    lowered = name.lower()
    if lowered in _WPS_WRITER_CLASSES:
        return True
    return lowered in {
        "opusapp",
        "nuidocumentwindow",
        "netuihwnd",
        "documentwindow",
        "mdiclient",
        "paneclassdc",
        "worddocument",
        "_wwg",
        "_wwb",
    }


def _stub_has_ms_presentation_signature(name: str) -> bool:
    lowered = name.lower()
    if lowered in {"screenclass", "pptviewwndclass", "pptframeclass"}:
        return True
    return "ppt" in lowered


def _stub_is_wps_presentation_process(process_name: str, *classes: str) -> bool:
    lowered = process_name.strip().lower()
    if not lowered:
        return False
    if lowered.startswith(("wpp", "wppt")):
        return True
    if "wpspresentation" in lowered:
        return True
    if lowered.startswith("wps"):
        for cls in classes:
            if not cls:
                continue
            lowered_cls = cls.lower()
            if lowered_cls in {"screenclass", "kwppshowframeclass", "kwppshowframe", "kwppshowwndclass"}:
                return True
            if _stub_has_wps_presentation_signature(cls):
                return True
    return False


def _stub_is_wps_writer_process(process_name: str, *classes: str) -> bool:
    lowered = process_name.strip().lower()
    if not lowered:
        return False
    if _stub_is_wps_presentation_process(process_name, *classes):
        return False
    if "wpswriter" in lowered:
        return True
    if any(_stub_has_wps_writer_signature(cls) for cls in classes if cls):
        return True
    if not lowered.startswith("wps"):
        return False
    for cls in classes:
        if not cls:
            continue
        lowered_cls = cls.lower()
        if lowered_cls in {"screenclass", "kwppshowframeclass", "kwppshowframe", "kwppshowwndclass"}:
            return False
    return False


def _compute_category(class_name: str, top_class: str, process_name: str) -> str:
    return helpers._compute_presentation_category(  # type: ignore[attr-defined]
        class_name,
        top_class,
        process_name,
        has_wps_presentation_signature=_stub_has_wps_presentation_signature,
        is_wps_slideshow_class=_stub_is_wps_slideshow_class,
        has_wps_writer_signature=_stub_has_wps_writer_signature,
        is_word_like_class=_stub_is_word_like_class,
        has_ms_presentation_signature=_stub_has_ms_presentation_signature,
        is_wps_presentation_process=_stub_is_wps_presentation_process,
        is_wps_writer_process=_stub_is_wps_writer_process,
    )


def test_presentation_category_prefers_wps_slideshow() -> None:
    assert (
        _compute_category("screenclass", "kwpsframeclass", "wppt.exe")
        == "wps_ppt"
    )


def test_presentation_category_identifies_wps_writer() -> None:
    assert _compute_category("kwpsframeclass", "", "wps.exe") == "wps_word"


def test_presentation_category_detects_ms_powerpoint() -> None:
    assert _compute_category("screenclass", "", "powerpnt.exe") == "ms_ppt"


def test_presentation_category_handles_wps_hosted_screenclass() -> None:
    assert _compute_category("screenclass", "", "wps.exe") == "wps_ppt"


class _MixinHarness(helpers._PresentationWindowMixin):  # type: ignore[attr-defined]
    def _overlay_widget(self):  # pragma: no cover - interface requirement
        return None


def test_is_wps_writer_process_requires_writer_hints() -> None:
    harness = _MixinHarness()
    assert harness._is_wps_writer_process("wpswriter.exe", "") is True
    assert harness._is_wps_writer_process("wps.exe", "kwpsframeclass") is True
    assert (
        harness._is_wps_writer_process("wps.exe", "screenclass", "kwppshowframeclass")
        is False
    )
