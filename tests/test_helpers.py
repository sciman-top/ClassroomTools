"""Tests for helper utilities defined in ClassroomTools prior to the Qt imports."""

from __future__ import annotations

import ast
import dataclasses
import enum
import functools
import logging
import math
import contextlib
import os
import shutil
import sys
import tempfile
import types

import pytest
from functools import singledispatch
from pathlib import Path
from typing import Any, Callable, Dict, Iterable, List, Mapping, Optional, Set, Tuple, cast


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
            "Dict": Dict,
            "functools": functools,
            "cast": cast,
            "Callable": Callable,
            "Any": Any,
            "dataclass": dataclasses.dataclass,
            "Enum": enum.Enum,
            "math": math,
            "singledispatch": singledispatch,
            "shutil": shutil,
            "win32gui": None,
            "_user32_top_level_hwnd": lambda hwnd: 0,
            "logger": logging.getLogger("ctools_helpers"),
        }
    )
    sys.modules[module.__name__] = module

    targets = {
        "_TRUE_STRINGS",
        "_FALSE_STRINGS",
        "_BooleanParseResult",
        "_coerce_bool",
        "_casefold_cached",
        "_coerce_to_text",
        "_normalize_text_token",
        "_normalize_class_token",
        "parse_bool",
        "_ensure_directory",
        "_ensure_writable_directory",
        "_preferred_app_directory",
        "_collect_resource_roots",
        "_ResourceLocator",
        "_ResolvedPathGroup",
        "_get_resource_locator",
        "_resolve_writable_resource",
        "_choose_writable_target",
        "_iter_unique_paths",
        "_normalize_path_marker",
        "_candidate_path_pool",
        "_remove_file_candidates",
        "_replicate_file_to_candidates",
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


def test_normalize_class_token_handles_varied_inputs() -> None:
    assert helpers._normalize_class_token("  KwppFrameClass  ") == "kwppframeclass"
    assert helpers._normalize_class_token(bytearray(b"WPSMainFrame")) == "wpsmainframe"
    assert helpers._normalize_class_token(memoryview(b" kwpsDocView")) == "kwpsdocview"
    assert (
        helpers._normalize_class_token("\x00KwppFrameClass\x00") == "kwppframeclass"
    )

    class _Explosive:
        def __str__(self) -> str:
            raise RuntimeError("boom")

    assert helpers._normalize_class_token(_Explosive()) == ""


def test_coerce_to_text_handles_common_inputs() -> None:
    assert helpers._coerce_to_text("value") == "value"
    assert helpers._coerce_to_text(b"value") == "value"
    assert helpers._coerce_to_text(bytearray(b"valu\xe9")) == "valu"
    assert helpers._coerce_to_text(memoryview(b"value")) == "value"


def test_coerce_to_text_rejects_missing_or_invalid_values() -> None:
    class _Explosive:
        def __str__(self) -> str:
            raise RuntimeError("boom")

    class _Stringy:
        def __str__(self) -> str:
            return "converted"

    assert helpers._coerce_to_text(None) is None
    assert helpers._coerce_to_text(_Explosive()) is None
    assert helpers._coerce_to_text(_Stringy()) == "converted"


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


def test_remove_file_candidates_deduplicates_and_respects_skip(tmp_path: Path) -> None:
    primary = tmp_path / "students.xlsx"
    sibling = tmp_path / "copy" / "students.xlsx"
    sibling.parent.mkdir()
    primary.write_text("plain")
    sibling.write_text("plain")
    keep = tmp_path / "keep.xlsx"
    keep.write_text("stay")
    helpers._remove_file_candidates(  # type: ignore[attr-defined]
        [str(primary), str(primary), str(sibling), str(keep)],
        skip=str(keep),
    )
    assert not primary.exists()
    assert not sibling.exists()
    assert keep.exists()


def test_replicate_file_to_candidates_creates_missing_directories(tmp_path: Path) -> None:
    source = tmp_path / "students.xlsx.enc"
    source.write_text("secret")
    target_a = tmp_path / "mirror" / "students.xlsx.enc"
    target_b = tmp_path / "students.xlsx.enc"
    helpers._replicate_file_to_candidates(  # type: ignore[attr-defined]
        str(source), [str(target_a), str(target_b), str(source)]
    )
    assert target_a.read_text() == "secret"
    assert target_b.read_text() == "secret"


def test_collect_resource_roots_includes_onefile_parent(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    parent = tmp_path / "onefile_parent"
    parent.mkdir()
    script_dir = tmp_path / "runtime"
    script_dir.mkdir()
    exe_dir = tmp_path / "launcher"
    exe_dir.mkdir()
    app_dir = tmp_path / "appdata"
    initial = tmp_path / "initial"
    initial.mkdir()

    monkeypatch.setenv("NUITKA_ONEFILE_PARENT", str(parent))
    monkeypatch.setattr(helpers, "_INITIAL_CWD", str(initial), raising=False)  # type: ignore[attr-defined]
    monkeypatch.setattr(helpers, "_preferred_app_directory", lambda: str(app_dir), raising=False)  # type: ignore[attr-defined]
    monkeypatch.setattr(helpers.sys, "argv", [str(script_dir / "ClassroomTools.exe")])
    monkeypatch.setattr(helpers.sys, "frozen", True, raising=False)
    monkeypatch.setattr(helpers.sys, "executable", str(exe_dir / "ClassroomTools.exe"), raising=False)
    helpers._get_resource_locator.cache_clear()  # type: ignore[attr-defined]

    roots = helpers._collect_resource_roots()  # type: ignore[attr-defined]

    normalized = {Path(root) for root in roots}
    assert Path(parent) in normalized
    assert Path(script_dir) in normalized
    assert Path(initial) in normalized
    assert Path(app_dir) in normalized
    assert Path(exe_dir) in normalized


def test_resolve_writable_resource_prefers_existing_candidate(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    script_dir = tmp_path / "script"
    script_dir.mkdir()
    existing = tmp_path / "data" / "students.xlsx"
    existing.parent.mkdir()
    existing.write_text("payload")
    app_dir = tmp_path / "appdata"
    initial = tmp_path / "initial"
    initial.mkdir()

    monkeypatch.setattr(helpers, "_INITIAL_CWD", str(initial), raising=False)  # type: ignore[attr-defined]
    monkeypatch.setattr(helpers, "_preferred_app_directory", lambda: str(app_dir), raising=False)  # type: ignore[attr-defined]
    monkeypatch.setattr(helpers.sys, "argv", [str(script_dir / "ClassroomTools.py")])
    monkeypatch.setattr(helpers.sys, "frozen", False, raising=False)
    monkeypatch.setenv("NUITKA_ONEFILE_PARENT", "")
    helpers._get_resource_locator.cache_clear()  # type: ignore[attr-defined]

    group = helpers._resolve_writable_resource(  # type: ignore[attr-defined]
        "students.xlsx",
        extra_candidates=(str(existing),),
        is_dir=False,
        copy_from_candidates=False,
    )

    assert group.primary == str(existing)


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


def test_presentation_category_accepts_non_string_hints(tmp_path: Path) -> None:
    path = tmp_path / "WPP.EXE"
    path.write_text("dummy")
    assert (
        _compute_category(b"KwppShowFrameClass", memoryview(b"kwpsframeclass"), path)
        == "wps_ppt"
    )
    assert _compute_category(bytearray(), None, None) == "other"


def test_prefix_keyword_classifier_normalizes_tokens() -> None:
    classifier = helpers._PresentationWindowMixin._PrefixKeywordClassifier(  # type: ignore[attr-defined]
        prefixes=[" WPP", "Wpp", "kwp"],
        keywords=["Show", "SHOW"],
        excludes=["Beta", "BETA"],
        canonical=["KwppsFrame"],
    )
    assert classifier.prefixes == ("wpp", "kwp")
    assert classifier.keywords == ("show",)
    assert classifier.excludes == ("beta",)
    assert "kwppsframe" in classifier.canonical
    assert classifier.has_signature("KWPPSFRAME") is True
    assert classifier.has_signature(" wppBetaWindow ") is False
    assert classifier.has_signature("wpp main show window") is True


def test_prefix_keyword_classifier_accepts_normalized_signature() -> None:
    classifier = helpers._PresentationWindowMixin._PrefixKeywordClassifier(  # type: ignore[attr-defined]
        prefixes=["kwps"],
        keywords=["frame"],
        excludes=["show"],
        canonical=["kwpsframeclass"],
    )
    assert classifier.has_signature("KWPSFrameClass") is True
    assert classifier.has_normalized_signature("kwpsframeclass") is True
    assert classifier.has_normalized_signature("kwpsframeclass ") is True


def test_class_tokens_freeze_normalizes_inputs() -> None:
    tokens = helpers._PresentationWindowMixin._ClassTokens.freeze(  # type: ignore[attr-defined]
        [" KWPPShowFrameClass "],
        {"KwpsDocView", "kwpsdocview"},
        b"WPSMainFrame",
        None,
    )
    assert tokens == {
        "kwppshowframeclass",
        "kwpsdocview",
        "wpsmainframe",
    }


def test_normalize_process_name_handles_bytes_and_failures() -> None:
    harness = _MixinHarness()

    assert harness._normalize_process_name(b"  WPS.EXE  ") == "wps.exe"
    assert harness._normalize_process_name(memoryview(b"WpPt.exe")) == "wppt.exe"
    assert harness._normalize_process_name("WPS.EXE\x00") == "wps.exe"
    assert harness._normalize_process_name(None) == ""

    class _Explosive:
        def __str__(self) -> str:
            raise ValueError("boom")

    assert harness._normalize_process_name(_Explosive()) == ""


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


def test_is_wps_presentation_process_detects_ms_screenclass() -> None:
    harness = _MixinHarness()
    assert harness._is_wps_presentation_process("wps.exe", "screenclass") is True


def test_summarize_wps_process_hints_collects_flags() -> None:
    harness = _MixinHarness()
    normalized = harness._normalized_class_hints(
        "KwppShowFrameClass", "kwpsframeclass", "pptviewwndclass"
    )
    hints = harness._summarize_wps_process_hints(normalized)
    assert hints.classes == normalized
    assert hints.has_slideshow is True
    assert hints.has_wps_presentation_signature is True
    assert hints.has_ms_presentation_signature is True
    assert hints.has_writer_signature is True


def test_summarize_wps_process_hints_ignores_predicate_errors() -> None:
    class _FlakyHarness(_MixinHarness):
        def _class_has_wps_presentation_signature(self, class_name: str) -> bool:
            if "boom" in class_name:
                raise RuntimeError("predicate failure")
            return super()._class_has_wps_presentation_signature(class_name)

    harness = _FlakyHarness()
    normalized = harness._normalized_class_hints("BoomWindow", "kwpsframeclass")
    hints = harness._summarize_wps_process_hints(normalized)
    assert hints.classes == normalized
    assert hints.has_wps_presentation_signature is False
    assert hints.has_writer_signature is True


def test_summarize_wps_process_hints_logs_with_fallback_logger() -> None:
    class _FlakyHarness(_MixinHarness):
        def _class_has_wps_presentation_signature(self, class_name: str) -> bool:
            if "boom" in class_name:
                raise RuntimeError("predicate failure")
            return super()._class_has_wps_presentation_signature(class_name)

    harness = _FlakyHarness()
    normalized = harness._normalized_class_hints("BoomWindow", "kwpsframeclass")
    original_logger = getattr(helpers, "logger", None)
    try:
        helpers.logger = object()  # type: ignore[assignment]
        hints = harness._summarize_wps_process_hints(normalized)
    finally:
        if original_logger is None:
            try:
                del helpers.logger
            except AttributeError:
                pass
        else:
            helpers.logger = original_logger
    assert hints.classes == normalized
    assert hints.has_wps_presentation_signature is False
    assert hints.has_writer_signature is True


def test_summarize_wps_process_hints_updates_cached_logger() -> None:
    class _FlakyHarness(_MixinHarness):
        def __init__(self) -> None:
            super().__init__()
            self.calls = 0

        def _class_has_wps_presentation_signature(self, class_name: str) -> bool:
            self.calls += 1
            raise RuntimeError("predicate failure")

    class _Recorder:
        def __init__(self) -> None:
            self.calls = 0

        def debug(self, *args: Any, **kwargs: Any) -> None:
            self.calls += 1

    harness = _FlakyHarness()
    normalized = harness._normalized_class_hints("BoomWindow", "kwpsframeclass")
    original_logger = getattr(helpers, "logger", None)
    recorder_one = _Recorder()
    recorder_two = _Recorder()

    try:
        helpers.logger = recorder_one  # type: ignore[assignment]
        hints = harness._summarize_wps_process_hints(normalized)
        assert hints.has_writer_signature is True
        assert recorder_one.calls == 2

        hints = harness._summarize_wps_process_hints(normalized)
        assert recorder_one.calls == 2

        helpers.logger = recorder_two  # type: ignore[assignment]
        hints = harness._summarize_wps_process_hints(normalized)
        assert recorder_two.calls == 2
    finally:
        if original_logger is None:
            try:
                del helpers.logger
            except AttributeError:
                pass
        else:
            helpers.logger = original_logger

    assert harness.calls == 4
    assert recorder_one.calls == 2


def test_summarize_wps_process_hints_caches_duplicate_classes() -> None:
    class _CountingHarness(_MixinHarness):
        def __init__(self) -> None:
            super().__init__()
            self.calls = 0

        def _class_has_wps_presentation_signature(self, class_name: str) -> bool:
            self.calls += 1
            return super()._class_has_wps_presentation_signature(class_name)

    harness = _CountingHarness()
    normalized = harness._normalized_class_hints("KwppShowFrameClass", "KwppShowFrameClass")
    hints = harness._summarize_wps_process_hints(normalized)
    assert hints.has_wps_presentation_signature is True
    assert harness.calls == 1


def test_summarize_wps_process_hints_ignores_duplicate_specs() -> None:
    class _DuplicateHarness(_MixinHarness):
        def __init__(self) -> None:
            super().__init__()
            self.calls = 0

        def _is_wps_slideshow_class(self, class_name: str) -> bool:
            self.calls += 1
            return super()._is_wps_slideshow_class(class_name)

        def _wps_hint_predicate_specs(self):  # type: ignore[override]
            specs = list(super()._wps_hint_predicate_specs())
            spec_type = type(specs[0])
            specs.append(
                spec_type(
                    specs[0].flag_name,
                    specs[0].predicate,
                    specs[0].normalized_predicate,
                    specs[0].base_impl,
                )
            )
            return tuple(specs)

    harness = _DuplicateHarness()
    normalized = harness._normalized_class_hints("KwppShowFrameClass")
    hints = harness._summarize_wps_process_hints(normalized)
    assert hints.has_slideshow is True
    assert harness.calls == 1


def test_summarize_wps_process_hints_skips_specs_without_flag_names() -> None:
    class _NamelessHarness(_MixinHarness):
        def _wps_hint_predicate_specs(self):  # type: ignore[override]
            specs = list(super()._wps_hint_predicate_specs())
            spec_type = type(specs[0])
            return (
                spec_type(
                    "",
                    specs[0].predicate,
                    specs[0].normalized_predicate,
                    specs[0].base_impl,
                ),
            )

    harness = _NamelessHarness()
    normalized = harness._normalized_class_hints("KwppShowFrameClass")
    hints = harness._summarize_wps_process_hints(normalized)
    assert hints.has_slideshow is False
    assert hints.has_wps_presentation_signature is False
    assert hints.has_ms_presentation_signature is False
    assert hints.has_writer_signature is False


def test_summarize_wps_process_hints_accepts_unbound_predicates() -> None:
    harness = _MixinHarness()
    specs = list(harness._wps_hint_predicate_specs())
    spec_type = type(specs[0])

    def _custom_specs(self):
        base = specs[0]
        return (
            spec_type(
                base.flag_name,
                base.base_impl,
                base.normalized_predicate,
                base.base_impl,
            ),
        )

    harness._wps_hint_predicate_specs = types.MethodType(_custom_specs, harness)
    normalized = harness._normalized_class_hints("KwppShowFrameClass")
    hints = harness._summarize_wps_process_hints(normalized)
    assert hints.has_slideshow is True


def test_summarize_wps_process_hints_respects_wrapped_overrides() -> None:
    class _WrappedHarness(_MixinHarness):
        def __init__(self) -> None:
            super().__init__()
            self.calls: List[str] = []

        @functools.wraps(
            helpers._PresentationWindowMixin._class_has_wps_presentation_signature  # type: ignore[attr-defined]
        )
        def _class_has_wps_presentation_signature(self, class_name: str) -> bool:
            self.calls.append(class_name)
            return False

    harness = _WrappedHarness()
    normalized = harness._normalized_class_hints("KwppShowFrameClass")
    hints = harness._summarize_wps_process_hints(normalized)
    assert hints.has_wps_presentation_signature is False
    assert harness.calls == ["kwppshowframeclass"]
    hints = harness._summarize_wps_process_hints(normalized)
    assert hints.has_wps_presentation_signature is False
    assert harness.calls == ["kwppshowframeclass"]
