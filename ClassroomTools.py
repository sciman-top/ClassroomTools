# -*- coding: utf-8 -*-
from __future__ import annotations

import base64
import configparser
import contextlib
import ctypes
import importlib
import io
import json
import math
import os
import random
import re
import shutil
import subprocess
import sys
import tempfile
import threading
import time
from queue import Empty, Queue
from dataclasses import dataclass
from typing import Callable, Dict, List, Optional, Mapping

from PyQt6.QtCore import (
    QByteArray,
    QPoint,
    QPointF,
    QRect,
    QSize,
    Qt,
    QTimer,
    QEvent,
    pyqtSignal,
    QObject,
)
from PyQt6.QtGui import (
    QBrush,
    QColor,
    QCursor,
    QFont,
    QFontDatabase,
    QFontMetrics,
    QIcon,
    QPainter,
    QPainterPath,
    QPainterPathStroker,
    QPen,
    QPixmap,
    QKeyEvent,
    QResizeEvent,
    QScreen,
)
from PyQt6.QtWidgets import (
    QApplication,
    QButtonGroup,
    QCheckBox,
    QDialog,
    QDialogButtonBox,
    QFrame,
    QGridLayout,
    QHBoxLayout,
    QLabel,
    QMenu,
    QMessageBox,
    QPushButton,
    QSpacerItem,
    QSizePolicy,
    QSlider,
    QSpinBox,
    QStackedWidget,
    QToolButton,
    QVBoxLayout,
    QWidget,
    QToolTip,
)

# ---------- 运行环境准备 ----------

def _prepare_windows_tts_environment() -> None:
    """确保 Windows 打包环境下的语音依赖可以写入缓存。"""

    if sys.platform != "win32":
        return
    cache_dir = os.environ.get("COMTYPES_CACHE_DIR", "").strip()
    if cache_dir and os.path.isdir(cache_dir):
        return
    try:
        base = os.environ.get("LOCALAPPDATA")
        if not base:
            base = os.path.join(os.path.expanduser("~"), "AppData", "Local")
        cache_dir = os.path.join(base, "ClassroomTools", "comtypes_cache")
        os.makedirs(cache_dir, exist_ok=True)
        os.environ["COMTYPES_CACHE_DIR"] = cache_dir
    except Exception:
        # 打包环境下若目录创建失败，也不要阻塞主程序。
        pass


_prepare_windows_tts_environment()

# ---------- 图标 ----------
class IconManager:
    """集中管理浮动工具条的 SVG 图标，方便后续统一换肤。"""
    _icons: Dict[str, str] = {
        "cursor": "PHN2ZyB4bWxucz0naHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmcnIHZpZXdCb3g9JzAgMCAyNCAyNCc+CiAgICA8cGF0aCBmaWxsPScjZjFmM2Y0JyBkPSdNNCAzLjMgMTEuNCAyMWwxLjgtNS44IDYuMy0yLjF6Jy8+CiAgICA8cGF0aCBmaWxsPScjOGFiNGY4JyBkPSdtMTIuNiAxNC40IDQuOCA0LjgtMi4xIDIuMS00LjItNC4yeicvPgo8L3N2Zz4=",
        "shape": "PHN2ZyB4bWxucz0naHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmcnIHZpZXdCb3g9JzAgMCAyNCAyNCc+CiAgICA8cmVjdCB4PSczLjUnIHk9JzMuNScgd2lkdGg9JzknIGhlaWdodD0nOScgcng9JzInIGZpbGw9JyNmMWYzZjQnLz4KICAgIDxjaXJjbGUgY3g9JzE2LjUnIGN5PScxNi41JyByPSc1LjUnIGZpbGw9J25vbmUnIHN0cm9rZT0nI2YxZjNmNCcgc3Ryb2tlLXdpZHRoPScxLjgnLz4KICAgIDxjaXJjbGUgY3g9JzE2LjUnIGN5PScxNi41JyByPSczLjUnIGZpbGw9JyM4YWI0ZjgnIGZpbGwtb3BhY2l0eT0nMC4zNScvPgo8L3N2Zz4=",
        "eraser": "PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHZpZXdCb3g9IjAgMCAyNCAyNCI+CiAgPHBhdGggZD0iTTQuNiAxNC4yIDExLjMgNy40YTIgMiAwIDAgMSAyLjggMGwzLjUgMy41YTIgMiAwIDAgMSAwIDIuOGwtNC44IDQuOEg5LjRhMiAyIDAgMCAxLTEuNC0uNmwtMy0zYTIgMiAwIDAgMSAwLTIuOHoiIGZpbGw9IiNmNGE5YjciLz4KICA8cGF0aCBkPSJNOS4yIDE5LjZoNi4xYy42IDAgMS4xLS4yIDEuNS0uNmwxLjctMS43IiBmaWxsPSJub25lIiBzdHJva2U9IiM1ZjYzNjgiIHN0cm9rZS13aWR0aD0iMS42IiBzdHJva2UtbGluZWNhcD0icm91bmQiLz4KICA8cGF0aCBkPSJtNy4yIDEyLjMgNC41IDQuNSIgZmlsbD0ibm9uZSIgc3Ryb2tlPSIjZmZmZmZmIiBzdHJva2Utd2lkdGg9IjEuNiIgc3Rya2UtbGluZWNhcD0icm91bmQiLz4KICA8cGF0aCBkPSJNMy42IDE4LjZoNiIgc3Ryb2tlPSIjNWY2MzY4IiBzdHJva2Utd2lkdGg9IjEuNiIgc3Rya2UtbGluZWNhcD0icm91bmQiLz4KPC9zdmc+",
        "clear_all": "PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHZpZXdCb3g9IjAgMCAyNCAyNCI+CiAgPGRlZnM+CiAgICA8bGluZWFyR3JhZGllbnQgaWQ9ImciIHgxPSIwIiB4Mj0iMCIgeTE9IjAiIHkyPSIxIj4KICAgICAgPHN0b3Agb2Zmc2V0PSIwIiBzdG9wLWNvbG9yPSIjOGFiNGY4Ii8+CiAgICAgIDxzdG9wIG9mZnNldD0iMSIgc3RvcC1jb2xvcj0iIzFhNzNlOCIvPgogICAgPC9saW5lYXJHcmFkaWVudD4KICA8L2RlZnM+CiAgPHBhdGggZD0iTTUuNSA4aDEzbC0uOSAxMS4yQTIgMiAwIDAgMSAxNS42IDIxSDguNGEyIDIgMCAwIDEtMS45LTEuOEw1LjUgOHoiIGZpbGw9InVybCgjZykiIHN0cm9rZT0iIzFhNzNlOCIgc3Rya2Utd2lkdGg9IjEuMiIvPgogIDxwYXRoIGQ9Ik05LjUgNS41IDEwLjMgNGgzLjRsLjggMS41aDQuNSIgZmlsbD0ibm9uZSIgc3Rya2U9IiM1ZjYzNjgiIHN0cm9rZS13aWR0aD0iMS42IiBzdHJva2UtbGluZWNhcD0icm91bmQiIHN0cm9rZS1saW5lam9pbj0icm91bmQiLz4KICA8cGF0aCBkPSJNNSA1LjVoNCIgc3Ryb2tlPSIjNWY2MzY4IiBzdHJva2Utd2lkdGg9IjEuNiIgc3Rya2UtbGluZWNhcD0icm91bmQiLz4KICA8cGF0aCBkPSJNMTAgMTEuMnY2LjFNMTQgMTEuMnY2LjEiIHN0cm9rZT0iI2ZmZmZmZiIgc3Rya2Utd2lkdGg9IjEuNCIgc3Rya2UtbGluZWNhcD0icm91bmQiLz4KICA8cGF0aCBkPSJNOC4yIDExLjJ2Ni4xIiBzdHJva2U9IiMzYjc4ZTciIHN0cm9rZS13aWR0aD0iMS40IiBzdHJva2UtbGluZWNhcD0icm91bmQiIG9wYWNpdHk9Ii43Ii8+CiAgPHBhdGggZD0iTTE1LjggMTEuMnY2LjEiIHN0cm9rZT0iIzNiNzhlNyIgc3Rya2Utd2lkdGg9IjEuNCIgc3Rya2UtbGluZWNhcD0icm91bmQiIG9wYWNpdHk9Ii43Ii8+CiAgPHBhdGggZD0iTTYuMiAzLjYgNy40IDIuNCIgc3Ryb2tlPSIjZmJiYzA0IiBzdHJva2Utd2lkdGg9IjEuNCIgc3Rya2UtbGluZWNhcD0icm91bmQiLz4KICA8cGF0aCBkPSJtMTguNCAzLjQgMS40LTEuNCIgc3Rya2U9IiMzNGE4NTMiIHN0cm9rZS13aWR0aD0iMS40IiBzdHJva2UtbGluZWNhcD0icm91bmQiLz4KPC9zdmc+",
        "settings": "PHN2ZyB4bWxucz0naHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmcnIHZpZXdCb3g9JzAgMCAyNCAyNCc+CiAgICA8Y2lyY2xlIGN4PScxMicgY3k9JzEyJyByPSczLjUnIGZpbGw9JyM4YWI0ZjgnLz4KICAgIDxwYXRoIGZpbGw9J25vbmUnIHN0cm9rZT0nI2YxZjNmNCcgc3Ryb2tlLXdpZHRoPScxLjYnIHN0cm9rZS1saW5lY2FwPSdyb3VuZCcgc3Ryb2tlLWxpbmVqb2luPSdyb3VuZCcKICAgICAgICBkPSdNMTIgNC41VjIuOG0wIDE4LjR2LTEuN203LjEtNy41SDIwbS0xOCAwaDEuNk0xNy42IDZsMS4yLTEuMk01LjIgMTguNCA2LjQgMTcuMk02LjQgNiA1LjIgNC44bTEzLjYgMTMuNi0xLjItMS4yJy8+Cjwvc3ZnPg==",
        "whiteboard": "PHN2ZyB4bWxucz0naHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmcnIHZpZXdCb3g9JzAgMCAyNCAyNCc+CiAgICA8cmVjdCB4PSczJyB5PSc0JyB3aWR0aD0nMTgnIGhlaWdodD0nMTInIHJ4PScyJyByeT0nMicgZmlsbD0nI2YxZjNmNCcgZmlsbC1vcGFjaXR5PScwLjEyJyBzdHJva2U9JyNmMWYzZjQnIHN0cm9rZS13aWR0aD0nMS42Jy8+CiAgICA8cGF0aCBkPSdtNyAxOCA1LTUgNSA1JyBmaWxsPSdub25lJyBzdHJva2U9JyM4YWI0ZjgnIHN0cm9rZS13aWR0aD0nMS44JyBzdHJva2UtbGluZWNhcD0ncm91bmQnIHN0cm9rZS1saW5lam9pbj0ncm91bmQnLz4KICAgIDxwYXRoIGQ9J004IDloOG0tOCAzaDUnIHN0cm9rZT0nI2YxZjNmNCcgc3Ryb2tlLXdpZHRoPScxLjYnIHN0cm9rZS1saW5lY2FwPSdyb3VuZCcvPgo8L3N2Zz4=",
        "undo": "PHN2ZyB4bWxucz0naHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmcnIHZpZXdCb3g9JzAgMCAyNCAyNCc+CiAgPHBhdGggZmlsbD0nI2YxZjNmNCcgZD0nTTguNCA1LjJMMyAxMC42bDUuNCA1LjQgMS40LTEuNC0yLjMtMi4zaDUuNWMzLjIgMCA1LjggMi42IDUuOCA1LjggMCAuNS0uMSAxLS4yIDEuNWwyLjEuNmMuMi0uNy4zLTEuNC4zLTIuMSAwLTQuNC0zLjYtOC04LThINy41bDIuMy0yLjMtMS40LTEuNHonLz4KPC9zdmc+",
    }
    _cache: Dict[str, QIcon] = {}

    @classmethod
    def get_brush_icon(cls, color_hex: str) -> QIcon:
        key = f"brush_{color_hex.lower()}"
        if key in cls._cache:
            return cls._cache[key]
        pixmap = QPixmap(28, 28)
        pixmap.fill(Qt.GlobalColor.transparent)
        painter = QPainter(pixmap)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        brush_color = QColor(color_hex)
        if not brush_color.isValid():
            brush_color = QColor("#999999")
        painter.setBrush(QBrush(brush_color))
        painter.setPen(QPen(QColor(0, 0, 0, 140), 1.4))
        painter.drawEllipse(5, 6, 18, 18)
        painter.setPen(QPen(QColor(255, 255, 255, 230), 3, Qt.PenStyle.SolidLine, Qt.PenCapStyle.RoundCap))
        painter.drawLine(9, 10, 18, 19)
        painter.setPen(QPen(QColor(0, 0, 0, 90), 2, Qt.PenStyle.SolidLine, Qt.PenCapStyle.RoundCap))
        painter.drawLine(10, 9, 19, 18)
        painter.end()
        icon = QIcon(pixmap)
        cls._cache[key] = icon
        return icon

    @classmethod
    def get_icon(cls, name: str) -> QIcon:
        """返回缓存的图标，如果未缓存则即时加载。"""
        if name == "clear":
            name = "clear_all"  # 兼容旧配置
        if name in cls._cache:
            return cls._cache[name]
        data = cls._icons.get(name)
        if not data:
            return QIcon()
        try:
            pixmap = QPixmap()
            pixmap.loadFromData(QByteArray.fromBase64(data.encode("ascii")), "SVG")
            icon = QIcon(pixmap)
            cls._cache[name] = icon
            return icon
        except Exception:
            return QIcon()


# ---------- 可选依赖 ----------
try:
    import pandas as pd
    PANDAS_AVAILABLE = True
except ImportError:
    pd = None
    PANDAS_AVAILABLE = False

try:
    import openpyxl  # noqa: F401
    OPENPYXL_AVAILABLE = True
except ImportError:
    OPENPYXL_AVAILABLE = False

try:
    import pyttsx3
    PYTTSX3_AVAILABLE = True
except ImportError:
    pyttsx3 = None
    PYTTSX3_AVAILABLE = False

try:
    import sounddevice as sd
    import numpy as np
    SOUNDDEVICE_AVAILABLE = True
except ImportError:
    sd = None
    np = None
    SOUNDDEVICE_AVAILABLE = False

if sys.platform == "win32":
    try:
        import winreg
        WINREG_AVAILABLE = True
    except ImportError:
        WINREG_AVAILABLE = False
else:
    WINREG_AVAILABLE = False


# ---------- DPI ----------
def ensure_high_dpi_awareness() -> None:
    if sys.platform != "win32":
        return
    os.environ.setdefault("QT_ENABLE_HIGHDPI_SCALING", "1")
    os.environ.setdefault("QT_SCALE_FACTOR_ROUNDING_POLICY", "PassThrough")
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(2)
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass


# ---------- 工具 ----------
def geometry_to_text(widget: QWidget) -> str:
    """以不含边框的尺寸记录窗口几何信息，避免重复放大。"""

    frame = widget.frameGeometry()
    inner = widget.geometry()
    width = inner.width() or frame.width()
    height = inner.height() or frame.height()
    return f"{width}x{height}+{frame.x()}+{frame.y()}"


def apply_geometry_from_text(widget: QWidget, geometry: str) -> None:
    if not geometry:
        return
    parts = geometry.split("+")
    if len(parts) != 3:
        return
    size_part, x_str, y_str = parts
    if "x" not in size_part:
        return
    width_str, height_str = size_part.split("x", 1)
    try:
        width = int(width_str)
        height = int(height_str)
        x = int(x_str)
        y = int(y_str)
    except ValueError:
        return

    base_min_width = getattr(widget, "_base_minimum_width", widget.minimumWidth())
    base_min_height = getattr(widget, "_base_minimum_height", widget.minimumHeight())

    custom_min_width = getattr(widget, "_ensure_min_width", 160)
    custom_min_height = getattr(widget, "_ensure_min_height", 120)

    min_width = max(base_min_width, custom_min_width)
    min_height = max(base_min_height, custom_min_height)

    screen = QApplication.screenAt(QPoint(x, y))
    if screen is None:
        try:
            screen = widget.screen() or QApplication.primaryScreen()
        except Exception:
            screen = QApplication.primaryScreen()
    if screen is not None:
        available = screen.availableGeometry()
        max_width = max(min_width, 320, int(available.width() * 0.9))
        max_height = max(min_height, 240, int(available.height() * 0.9))
        width = max(min_width, min(width, max_width))
        height = max(min_height, min(height, max_height))
        x = max(available.left(), min(x, available.right() - width))
        y = max(available.top(), min(y, available.bottom() - height))
    target_width = max(min_width, width)
    target_height = max(min_height, height)
    widget.resize(target_width, target_height)
    widget.move(x, y)


def ensure_widget_within_screen(widget: QWidget) -> None:
    screen = None
    try:
        screen = widget.screen()
    except Exception:
        screen = None
    if screen is None:
        screen = QApplication.primaryScreen()
    if screen is None:
        return
    base_min_width = getattr(widget, "_base_minimum_width", widget.minimumWidth())
    base_min_height = getattr(widget, "_base_minimum_height", widget.minimumHeight())

    custom_min_width = getattr(widget, "_ensure_min_width", 160)
    custom_min_height = getattr(widget, "_ensure_min_height", 120)

    min_width = max(base_min_width, custom_min_width)
    min_height = max(base_min_height, custom_min_height)

    available = screen.availableGeometry()
    geom = widget.frameGeometry()
    width = widget.width() or geom.width() or widget.sizeHint().width()
    height = widget.height() or geom.height() or widget.sizeHint().height()
    max_width = min(available.width(), max(min_width, int(available.width() * 0.9)))
    max_height = min(available.height(), max(min_height, int(available.height() * 0.9)))
    width = max(min_width, min(width, max_width))
    height = max(min_height, min(height, max_height))
    left_limit = available.x()
    top_limit = available.y()
    right_limit = max(left_limit, available.x() + available.width() - width)
    bottom_limit = max(top_limit, available.y() + available.height() - height)
    x = geom.x() if geom.width() else widget.x()
    y = geom.y() if geom.height() else widget.y()
    x = max(left_limit, min(x, right_limit))
    y = max(top_limit, min(y, bottom_limit))
    widget.resize(width, height)
    widget.move(x, y)


def str_to_bool(value: str, default: bool = False) -> bool:
    if isinstance(value, str):
        return value.strip().lower() in {"1", "true", "yes", "on"}
    return default


def bool_to_str(value: bool) -> str:
    return "True" if value else "False"


def dedupe_strings(values: List[str]) -> List[str]:
    seen: set[str] = set()
    unique: List[str] = []
    for value in values:
        normalized = value.strip() if isinstance(value, str) else ""
        if not normalized or normalized in seen:
            continue
        seen.add(normalized)
        unique.append(normalized)
    return unique


def _try_import_module(module: str) -> bool:
    try:
        importlib.import_module(module)
        return True
    except Exception:
        return False


def _count_windows_voice_tokens() -> tuple[int, Optional[str]]:
    if not WINREG_AVAILABLE:
        return -1, "无法访问 Windows 注册表"
    token_names: set[str] = set()
    path = r"SOFTWARE\\Microsoft\\Speech\\Voices\\Tokens"
    flags = {0}
    for attr in ("KEY_WOW64_32KEY", "KEY_WOW64_64KEY"):
        flag = getattr(winreg, attr, 0) if WINREG_AVAILABLE else 0  # type: ignore[name-defined]
        if flag:
            flags.add(flag)
    try:
        for hive in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):  # type: ignore[name-defined]
            for flag in flags:
                access = getattr(winreg, "KEY_READ", 0) | flag  # type: ignore[name-defined]
                try:
                    handle = winreg.OpenKey(hive, path, 0, access)  # type: ignore[name-defined]
                except FileNotFoundError:
                    continue
                except OSError as exc:
                    return 0, str(exc)
                with contextlib.closing(handle):
                    index = 0
                    while True:
                        try:
                            name = winreg.EnumKey(handle, index)  # type: ignore[name-defined]
                        except OSError:
                            break
                        token_names.add(str(name))
                        index += 1
    except Exception as exc:
        return 0, str(exc)
    return len(token_names), None


def detect_speech_environment_issues() -> tuple[str, List[str]]:
    issues: List[str] = []
    suggestions: List[str] = []
    if sys.platform == "win32":
        missing: List[str] = []
        module_hints = (
            ("pyttsx3", "请安装 pyttsx3（pip install pyttsx3）后重新启动程序。"),
            ("comtypes.client", "请安装 comtypes（pip install comtypes）后重新启动程序。"),
            ("win32com.client", "请安装 pywin32（pip install pywin32）后重新启动程序。"),
        )
        for module_name, hint in module_hints:
            if not _try_import_module(module_name):
                base_name = module_name.split(".")[0]
                if base_name not in missing:
                    missing.append(base_name)
                suggestions.append(hint)
        if missing:
            issues.append(f"缺少语音依赖：{'、'.join(sorted(missing))}")
        token_count, token_error = _count_windows_voice_tokens()
        if token_error:
            issues.append(f"无法读取语音包：{token_error}")
        elif token_count == 0:
            issues.append("系统未检测到任何语音包")
            suggestions.append("请在 Windows 的“设置 -> 时间和语言 -> 语音”中下载并启用语音包。")
        powershell = shutil.which("pwsh") or shutil.which("powershell")
        if not powershell:
            issues.append("未检测到 PowerShell，可用的语音播报方式受限")
            suggestions.append("请确认系统已安装 PowerShell 5+ 或 PowerShell 7，并在环境变量中可用。")
        if getattr(sys, "frozen", False):
            suggestions.append("若为打包版本，请在打包配置中包含 pyttsx3、comtypes、pywin32 等语音依赖或在目标电脑上单独安装它们。")
        suggestions.append("如已安装语音组件，请尝试以管理员权限首次运行程序以初始化语音服务。")
    else:
        suggestions.append("请确保系统已配置可用的语音引擎后重新启动程序。")
    return "；".join(issues), dedupe_strings(suggestions)


class QuietInfoPopup(QWidget):
    """提供一个静音的小型提示窗口，避免系统提示音干扰课堂。"""

    _active_popups: List["QuietInfoPopup"] = []

    def __init__(self, parent: Optional[QWidget], text: str, title: str) -> None:
        flags = (
            Qt.WindowType.Tool
            | Qt.WindowType.WindowTitleHint
            | Qt.WindowType.WindowCloseButtonHint
            | Qt.WindowType.CustomizeWindowHint
        )
        super().__init__(parent, flags)
        self.setWindowTitle(title)
        self.setAttribute(Qt.WidgetAttribute.WA_DeleteOnClose, True)
        self.setWindowModality(Qt.WindowModality.ApplicationModal)
        self.setMinimumWidth(240)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(18, 18, 18, 12)
        layout.setSpacing(12)

        self.message_label = QLabel(text, self)
        self.message_label.setWordWrap(True)
        self.message_label.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignTop)
        layout.addWidget(self.message_label)

        button_row = QHBoxLayout()
        button_row.addStretch(1)
        self.ok_button = QPushButton("确定", self)
        self.ok_button.setCursor(Qt.CursorShape.PointingHandCursor)
        self.ok_button.setDefault(True)
        self.ok_button.clicked.connect(self.close)
        button_row.addWidget(self.ok_button)
        layout.addLayout(button_row)

        QuietInfoPopup._active_popups.append(self)
        self.destroyed.connect(self._cleanup)

    def showEvent(self, event) -> None:  # type: ignore[override]
        super().showEvent(event)
        self.adjustSize()
        self._relocate()
        self.activateWindow()
        self.ok_button.setFocus(Qt.FocusReason.ActiveWindowFocusReason)

    def _relocate(self) -> None:
        target_rect: Optional[QRect] = None
        parent = self.parentWidget()
        if parent and parent.isVisible():
            target_rect = parent.frameGeometry()
        else:
            screen = QApplication.primaryScreen()
            if screen:
                target_rect = screen.availableGeometry()
        if not target_rect:
            return
        geo = self.frameGeometry()
        geo.moveCenter(target_rect.center())
        self.move(geo.topLeft())

    def _cleanup(self, *_args) -> None:
        try:
            QuietInfoPopup._active_popups.remove(self)
        except ValueError:
            pass


def show_quiet_information(parent: Optional[QWidget], text: str, title: str = "提示") -> None:
    popup = QuietInfoPopup(parent, text, title)
    popup.show()


# ---------- 配置 ----------
class SettingsManager:
    """负责读取/写入配置文件的轻量封装。"""

    def __init__(self, filename: str = "settings.ini") -> None:
        # 统一维护配置文件的存放路径，优先使用用户配置目录，保证跨次启动仍能读取到历史点名状态。
        self._mirror_targets: set[str] = set()
        self.filename = self._prepare_storage_path(filename)
        self._mirror_targets.add(self.filename)
        self.config = configparser.ConfigParser()
        self.config.optionxform = str
        self._settings_cache: Optional[Dict[str, Dict[str, str]]] = None
        self.defaults: Dict[str, Dict[str, str]] = {
            "Launcher": {
                "x": "120",
                "y": "120",
                "minimized": "False",
                "bubble_x": "120",
                "bubble_y": "120",
            },
            "Startup": {"autostart_enabled": "False"},
            "RollCallTimer": {
                "geometry": "480x280+180+180",
                "show_id": "True",
                "show_name": "True",
                "speech_enabled": "False",
                "speech_voice_id": "",
                "current_group": "全部",
                "timer_countdown_minutes": "5",
                "timer_countdown_seconds": "0",
                "timer_sound_enabled": "True",
                "mode": "roll_call",
                "timer_mode": "countdown",
                "timer_seconds_left": "300",
                "timer_stopwatch_seconds": "0",
                "timer_running": "False",
                "id_font_size": "48",
                "name_font_size": "60",
                "timer_font_size": "56",
                "scoreboard_order": "rank",
            },
            "Paint": {"x": "260", "y": "260", "brush_size": "12", "brush_color": "#ff0000"},
        }

    def _get_config_dir(self) -> str:
        """返回当前系统下建议的配置目录。"""

        home = os.path.expanduser("~")
        if sys.platform.startswith("win"):
            base = os.environ.get("APPDATA") or os.path.join(home, "AppData", "Roaming")
            return os.path.join(base, "ClassroomTools")
        if sys.platform == "darwin":
            return os.path.join(home, "Library", "Application Support", "ClassroomTools")
        base = os.environ.get("XDG_CONFIG_HOME") or os.path.join(home, ".config")
        return os.path.join(base, "ClassroomTools")

    def _prepare_storage_path(self, filename: str) -> str:
        """根据平台选择合适的设置文件路径，并在需要时迁移旧文件。"""

        base_name = os.path.basename(filename) or "settings.ini"
        legacy_path = os.path.abspath(filename)
        config_dir = self._get_config_dir()
        try:
            os.makedirs(config_dir, exist_ok=True)
        except OSError:
            config_dir = os.path.dirname(legacy_path) or os.getcwd()
        target_path = os.path.join(config_dir, base_name)

        if os.path.exists(legacy_path):
            same_file = False
            try:
                same_file = os.path.samefile(legacy_path, target_path)
            except (OSError, FileNotFoundError):
                same_file = False
            if not same_file and not os.path.exists(target_path):
                try:
                    shutil.copy2(legacy_path, target_path)
                except Exception:
                    target_path = legacy_path
            if not same_file:
                legacy_dir = os.path.dirname(legacy_path) or os.getcwd()
                if os.access(legacy_dir, os.W_OK):
                    self._mirror_targets.add(legacy_path)

        try:
            with open(target_path, "a", encoding="utf-8"):
                pass
        except OSError:
            target_path = legacy_path

        return target_path

    def get_defaults(self) -> Dict[str, Dict[str, str]]:
        return {section: values.copy() for section, values in self.defaults.items()}

    def load_settings(self) -> Dict[str, Dict[str, str]]:
        if self._settings_cache is not None:
            return {section: values.copy() for section, values in self._settings_cache.items()}

        settings = self.get_defaults()
        if os.path.exists(self.filename):
            try:
                self.config.read(self.filename, encoding="utf-8")
                for section in self.config.sections():
                    if section not in settings:
                        settings[section] = {}
                    for key, value in self.config.items(section):
                        settings[section][key] = value
            except configparser.Error:
                settings = self.get_defaults()

        self._settings_cache = {section: values.copy() for section, values in settings.items()}
        return {section: values.copy() for section, values in self._settings_cache.items()}

    def _write_atomic(self, path: str, data: str) -> None:
        directory = os.path.dirname(path)
        if directory and not os.path.exists(directory):
            os.makedirs(directory, exist_ok=True)
        target_dir = directory or os.getcwd()
        fd, tmp_path = tempfile.mkstemp(prefix="ctools_", suffix=".tmp", dir=target_dir)
        try:
            with os.fdopen(fd, "w", encoding="utf-8") as handle:
                handle.write(data)
            os.replace(tmp_path, path)
        except Exception:
            try:
                os.remove(tmp_path)
            except OSError:
                pass
            raise

    def save_settings(self, settings: Dict[str, Dict[str, str]]) -> None:
        config = configparser.ConfigParser()
        config.optionxform = str
        snapshot: Dict[str, Dict[str, str]] = {}
        for section, options in settings.items():
            snapshot[section] = {key: str(value) for key, value in options.items()}
            config[section] = snapshot[section]

        buffer = io.StringIO()
        config.write(buffer)
        data = buffer.getvalue()
        buffer.close()

        failed: List[str] = []
        for path in sorted(self._mirror_targets):
            try:
                self._write_atomic(path, data)
            except Exception:
                failed.append(path)
        if failed and self.filename not in failed:
            print(f"无法写入备用配置文件: {failed}")

        self._settings_cache = {section: values.copy() for section, values in snapshot.items()}

    def clear_roll_call_history(self) -> None:
        """清除点名历史信息，仅在用户主动重置时调用。"""

        settings = self.load_settings()
        section = settings.get("RollCallTimer", {})
        removed = False
        for key in ("group_remaining", "group_last", "global_drawn"):
            if key in section:
                section.pop(key, None)
                removed = True
        if removed:
            settings["RollCallTimer"] = section
            self.save_settings(settings)


# ---------- 自绘置顶 ToolTip ----------
class TipWindow(QWidget):
    """一个轻量的自绘 ToolTip，确保位于所有置顶窗之上。"""
    def __init__(self) -> None:
        super().__init__(None, Qt.WindowType.Tool | Qt.WindowType.FramelessWindowHint | Qt.WindowType.WindowStaysOnTopHint)
        self.setAttribute(Qt.WidgetAttribute.WA_TransparentForMouseEvents, True)
        self.setAttribute(Qt.WidgetAttribute.WA_ShowWithoutActivating, True)
        self.setStyleSheet(
            """
            QLabel {
                color: #f1f3f4;
                background: rgba(32, 33, 36, 230);
                border: 1px solid rgba(255, 255, 255, 45);
                border-radius: 6px;
                padding: 4px 8px;
                font-size: 12px;
            }
            """
        )
        self._label = QLabel("", self)
        self._label.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
        self._hide_timer = QTimer(self); self._hide_timer.setSingleShot(True); self._hide_timer.timeout.connect(self.hide)

    def show_tip(self, text: str, pos: QPoint, duration_ms: int = 2500) -> None:
        self._label.setText(text or "")
        self._label.adjustSize()
        self.resize(self._label.size())
        self.move(pos + QPoint(12, 16))
        self.show()
        self.raise_()
        self._hide_timer.start(duration_ms)

    def hide_tip(self) -> None:
        self._hide_timer.stop()
        self.hide()


# ---------- 对话框 ----------
class PenSettingsDialog(QDialog):
    """画笔粗细与颜色选择对话框。"""
    COLORS = {
        "#FFFF00": "黄",
        "#FFA500": "橙",
        "#24B47E": "绿",
        "#800080": "紫",
        "#FFFFFF": "白",
    }

    def __init__(self, parent: Optional[QWidget] = None, initial_size: int = 12, initial_color: str = "#FF0000") -> None:
        super().__init__(parent)
        self.setWindowTitle("画笔设置")
        self.setWindowFlag(Qt.WindowType.WindowStaysOnTopHint, True)
        self.pen_color = QColor(initial_color)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(8, 8, 8, 8)
        layout.setSpacing(6)

        size_layout = QHBoxLayout()
        size_label = QLabel("粗细:")
        self.size_slider = QSlider(Qt.Orientation.Horizontal)
        self.size_slider.setRange(2, 40)
        self.size_slider.setValue(int(initial_size))
        self.size_slider.setMinimumWidth(120)
        self.size_value = QLabel(str(initial_size))
        self.size_slider.valueChanged.connect(lambda value: self.size_value.setText(str(value)))

        size_layout.addWidget(size_label)
        size_layout.addWidget(self.size_slider, 1)
        size_layout.addWidget(self.size_value)
        layout.addLayout(size_layout)

        layout.addWidget(QLabel("颜色:"))
        color_layout = QGridLayout()
        color_layout.setContentsMargins(0, 0, 0, 0)
        color_layout.setSpacing(6)
        for index, (color_hex, name) in enumerate(self.COLORS.items()):
            button = QPushButton()
            button.setFixedSize(24, 24)
            button.setCursor(Qt.CursorShape.PointingHandCursor)
            button.setStyleSheet(
                f"background-color: {color_hex}; border: 1px solid rgba(0, 0, 0, 60); border-radius: 12px;"
            )
            button.setToolTip(name)
            button.clicked.connect(lambda _checked, c=color_hex: self._select_color(c))
            color_layout.addWidget(button, index // 3, index % 3)
        layout.addLayout(color_layout)

        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

        self.setFixedSize(self.sizeHint())

    def _select_color(self, color_hex: str) -> None:
        self.pen_color = QColor(color_hex)

    def get_settings(self) -> tuple[int, QColor]:
        return self.size_slider.value(), self.pen_color

    def showEvent(self, event) -> None:  # type: ignore[override]
        super().showEvent(event)
        ensure_widget_within_screen(self)


class ShapeSettingsDialog(QDialog):
    """图形工具的快捷选择窗口。"""
    SHAPES = {"line": "直线", "dashed_line": "虚线", "rect": "矩形", "circle": "圆形"}

    def __init__(self, parent: Optional[QWidget] = None) -> None:
        super().__init__(parent)
        self.setWindowTitle("选择图形")
        self.setWindowFlag(Qt.WindowType.WindowStaysOnTopHint, True)
        self.selected_shape: Optional[str] = None

        layout = QGridLayout(self)
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setHorizontalSpacing(6)
        layout.setVerticalSpacing(6)

        for index, (shape_key, name) in enumerate(self.SHAPES.items()):
            button = QPushButton(name)
            button.setCursor(Qt.CursorShape.PointingHandCursor)
            button.setFixedWidth(68)
            button.clicked.connect(lambda _checked, key=shape_key: self._select_shape(key))
            layout.addWidget(button, index // 2, index % 2)

        self.setFixedSize(self.sizeHint())

    def _select_shape(self, shape: str) -> None:
        self.selected_shape = shape
        self.accept()

    def get_shape(self) -> Optional[str]:
        return self.selected_shape

    def showEvent(self, event) -> None:  # type: ignore[override]
        super().showEvent(event)
        ensure_widget_within_screen(self)


class BoardColorDialog(QDialog):
    """白板背景颜色选择对话框。"""
    COLORS = {"#FFFFFF": "白板", "#000000": "黑板", "#0E4020": "绿板"}

    def __init__(self, parent: Optional[QWidget] = None) -> None:
        super().__init__(parent)
        self.setWindowTitle("选择颜色")
        self.setWindowFlag(Qt.WindowType.WindowStaysOnTopHint, True)
        self.selected_color: Optional[QColor] = None

        layout = QHBoxLayout(self)
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(6)

        for color_hex, name in self.COLORS.items():
            button = QPushButton(name)
            button.setCursor(Qt.CursorShape.PointingHandCursor)
            button.setFixedSize(72, 28)
            button.clicked.connect(lambda _checked, c=color_hex: self._select_color(c))
            layout.addWidget(button)

        self.setFixedSize(self.sizeHint())

    def _select_color(self, color_hex: str) -> None:
        self.selected_color = QColor(color_hex)
        self.accept()

    def get_color(self) -> Optional[QColor]:
        return self.selected_color

    def showEvent(self, event) -> None:  # type: ignore[override]
        super().showEvent(event)
        ensure_widget_within_screen(self)


# ---------- 标题栏 ----------
class TitleBar(QWidget):
    """浮动工具条的标题栏，负责拖拽移动。"""

    def __init__(self, toolbar: "FloatingToolbar") -> None:
        super().__init__(toolbar)
        self.toolbar = toolbar
        self._dragging = False
        self._drag_offset = QPoint()
        self.setCursor(Qt.CursorShape.OpenHandCursor)

        self.setAutoFillBackground(True)
        palette = self.palette()
        palette.setColor(self.backgroundRole(), QColor(36, 37, 41, 235))
        self.setPalette(palette)
        self.setFixedHeight(22)

        layout = QHBoxLayout(self)
        layout.setContentsMargins(6, 0, 6, 0)
        title = QLabel("屏幕画笔")
        font = title.font()
        font.setBold(True)
        title.setFont(font)
        title.setStyleSheet("color: #f1f3f4; font-size: 10.5px;")
        layout.addWidget(title)
        layout.addStretch()

    def mousePressEvent(self, event) -> None:
        if event.button() == Qt.MouseButton.LeftButton:
            self._dragging = True
            self._drag_offset = event.globalPosition().toPoint() - self.toolbar.pos()
            self.setCursor(Qt.CursorShape.ClosedHandCursor)
            event.accept()
        super().mousePressEvent(event)

    def mouseMoveEvent(self, event) -> None:
        if self._dragging:
            self.toolbar.move(event.globalPosition().toPoint() - self._drag_offset)
            event.accept()
        super().mouseMoveEvent(event)

    def mouseReleaseEvent(self, event) -> None:
        if self._dragging:
            self.toolbar.overlay.save_window_position()
        self._dragging = False
        self.setCursor(Qt.CursorShape.OpenHandCursor)
        try:
            self.toolbar.overlay.raise_toolbar()
        except Exception:
            pass
        event.accept()
        super().mouseReleaseEvent(event)


# ---------- 浮动工具条（画笔/白板） ----------
class FloatingToolbar(QWidget):
    """悬浮工具条：提供画笔、图形、白板等常用按钮。"""

    def __init__(self, overlay: "OverlayWindow", settings_manager: SettingsManager) -> None:
        super().__init__(
            None,
            Qt.WindowType.Tool | Qt.WindowType.FramelessWindowHint | Qt.WindowType.WindowStaysOnTopHint,
        )
        self.overlay = overlay
        self.settings_manager = settings_manager
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        self.setAttribute(Qt.WidgetAttribute.WA_AlwaysShowToolTips, True)
        self.setWindowFlag(Qt.WindowType.WindowDoesNotAcceptFocus, True)
        self._tip = TipWindow()
        self._build_ui()

        settings = self.settings_manager.load_settings().get("Paint", {})
        self.move(int(settings.get("x", "260")), int(settings.get("y", "260")))
        self.adjustSize()
        self.setFixedSize(self.sizeHint())
        self._base_minimum_width = self.width()
        self._base_minimum_height = self.height()
        self._ensure_min_width = self.width()
        self._ensure_min_height = self.height()

    def _build_ui(self) -> None:
        self.setStyleSheet(
            """
            #container {
                background-color: rgba(28, 29, 32, 235);
                border-radius: 10px;
                border: 1px solid rgba(255, 255, 255, 45);
            }
            QPushButton {
                color: #f1f3f4;
                background: rgba(60, 64, 67, 240);
                border: 1px solid rgba(255, 255, 255, 45);
                border-radius: 6px;
                padding: 3px;
                min-width: 28px;
                min-height: 28px;
            }
            QPushButton:hover {
                background: rgba(138, 180, 248, 245);
                border-color: rgba(138, 180, 248, 255);
                color: #202124;
            }
            QPushButton:checked {
                background: rgba(138, 180, 248, 255);
                color: #202124;
            }
            QPushButton#eraserButton {
                background: rgba(241, 243, 244, 235);
                color: #3c4043;
                border-color: rgba(138, 180, 248, 170);
            }
            QPushButton#eraserButton:hover,
            QPushButton#eraserButton:checked {
                background: rgba(202, 225, 255, 255);
                border-color: #1a73e8;
                color: #174ea6;
            }
            QPushButton#clearButton {
                background: rgba(255, 236, 232, 240);
                color: #a03a1e;
                border-color: rgba(255, 173, 153, 230);
            }
            QPushButton#clearButton:hover,
            QPushButton#clearButton:checked {
                background: rgba(255, 210, 204, 255);
                border-color: rgba(255, 138, 101, 255);
                color: #5f2121;
            }
            #whiteboardButtonActive {
                background: rgba(255, 214, 102, 240);
                border-color: rgba(251, 188, 5, 255);
                color: #202124;
            }
            """
        )

        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        container = QWidget(self)
        container.setObjectName("container")
        root.addWidget(container)

        layout = QVBoxLayout(container)
        layout.setContentsMargins(6, 5, 6, 6)
        layout.setSpacing(5)
        self.title_bar = TitleBar(self)
        layout.addWidget(self.title_bar)

        button_row = QHBoxLayout()
        button_row.setContentsMargins(0, 0, 0, 0)
        button_row.setSpacing(3)

        self.btn_cursor = QPushButton(IconManager.get_icon("cursor"), "")
        self.brush_color_buttons: Dict[str, QPushButton] = {}
        brush_configs = [
            ("#000000", "黑色画笔"),
            ("#FF0000", "红色画笔"),
            ("#1E90FF", "蓝色画笔"),
        ]  # 预设的高频画笔颜色
        brush_buttons = []
        for color_hex, name in brush_configs:
            button = QPushButton(IconManager.get_brush_icon(color_hex), "")
            button.setToolTip(name)
            self.brush_color_buttons[color_hex.lower()] = button
            brush_buttons.append(button)
        self.btn_shape = QPushButton(IconManager.get_icon("shape"), "")
        self.btn_undo = QPushButton(IconManager.get_icon("undo"), "")
        self.btn_eraser = QPushButton(IconManager.get_icon("eraser"), "")
        self.btn_eraser.setObjectName("eraserButton")
        self.btn_clear_all = QPushButton(IconManager.get_icon("clear_all"), "")
        self.btn_clear_all.setObjectName("clearButton")
        self.btn_whiteboard = QPushButton(IconManager.get_icon("whiteboard"), "")
        self.btn_settings = QPushButton(IconManager.get_icon("settings"), "")

        buttons = (
            [self.btn_cursor]
            + brush_buttons
            + [
                self.btn_shape,
                self.btn_undo,
                self.btn_eraser,
                self.btn_clear_all,
                self.btn_whiteboard,
                self.btn_settings,
            ]
        )
        for btn in buttons:
            btn.setIconSize(QSize(18, 18))
            btn.setCursor(Qt.CursorShape.PointingHandCursor)
            btn.setFocusPolicy(Qt.FocusPolicy.NoFocus)
            button_row.addWidget(btn)
        layout.addLayout(button_row)

        tooltip_text = {
            self.btn_cursor: "光标",
            self.btn_shape: "图形",
            self.btn_undo: "撤销",
            self.btn_eraser: "橡皮（再次点击恢复画笔）",
            self.btn_clear_all: "清除（并恢复画笔）",
            self.btn_whiteboard: "白板（单击开关 / 双击换色）",
            self.btn_settings: "画笔设置",
        }
        for button in brush_buttons:
            tooltip_text[button] = button.toolTip() or "画笔"
        for btn, tip_text in tooltip_text.items():
            btn.setToolTip(tip_text)
            btn.installEventFilter(self)

        self.tool_buttons = QButtonGroup(self)
        for btn in (self.btn_cursor, *brush_buttons, self.btn_shape, self.btn_eraser):
            btn.setCheckable(True)
            self.tool_buttons.addButton(btn)
        self.tool_buttons.setExclusive(True)

        self.btn_cursor.clicked.connect(self.overlay.toggle_cursor_mode)
        for color_hex, button in zip([c for c, _ in brush_configs], brush_buttons):
            button.clicked.connect(lambda _checked, c=color_hex: self.overlay.use_brush_color(c))
        self.btn_shape.clicked.connect(self._select_shape)
        self.btn_undo.clicked.connect(self.overlay.undo_last_action)
        self.btn_eraser.clicked.connect(self.overlay.toggle_eraser_mode)
        self.btn_clear_all.clicked.connect(self.overlay.clear_all)
        self.btn_settings.clicked.connect(self.overlay.open_pen_settings)
        self.btn_whiteboard.clicked.connect(self._handle_whiteboard_click)

        self._wb_click_timer = QTimer(self)
        self._wb_click_timer.setInterval(QApplication.instance().doubleClickInterval())
        self._wb_click_timer.setSingleShot(True)
        # 使用单次定时器来区分白板按钮的单击与双击行为
        self._wb_click_timer.timeout.connect(self.overlay.toggle_whiteboard)

        self.btn_undo.setEnabled(False)

        for widget in (self, container, self.title_bar):
            widget.installEventFilter(self)

    def update_tool_states(self, mode: str, pen_color: QColor) -> None:
        color_key = pen_color.name().lower()
        for hex_key, button in self.brush_color_buttons.items():
            prev = button.blockSignals(True)
            button.setChecked(mode == "brush" and hex_key == color_key)
            button.blockSignals(prev)
        for tool, button in (
            ("cursor", self.btn_cursor),
            ("shape", self.btn_shape),
            ("eraser", self.btn_eraser),
        ):
            prev = button.blockSignals(True)
            button.setChecked(mode == tool)
            button.blockSignals(prev)
        if mode == "brush" and color_key not in self.brush_color_buttons:
            for button in (self.btn_cursor, self.btn_shape, self.btn_eraser):
                prev = button.blockSignals(True)
                button.setChecked(False)
                button.blockSignals(prev)

    def update_undo_state(self, enabled: bool) -> None:
        self.btn_undo.setEnabled(enabled)

    def eventFilter(self, obj, event):
        event_type = event.type()
        if isinstance(obj, QPushButton) and event_type == QEvent.Type.ToolTip:
            try:
                self.overlay.raise_toolbar()
            except Exception:
                pass
            self._tip.show_tip(obj.toolTip(), QCursor.pos())
            return True
        if event_type in (
            QEvent.Type.Leave,
            QEvent.Type.MouseButtonPress,
            QEvent.Type.MouseButtonDblClick,
        ):
            self._tip.hide_tip()
        return super().eventFilter(obj, event)

    def _select_shape(self) -> None:
        dialog = ShapeSettingsDialog(self)
        if dialog.exec():
            shape = dialog.get_shape()
            if shape:
                self.overlay.set_mode("shape", shape_type=shape)
        else:
            self.overlay.update_toolbar_state()

    def _handle_whiteboard_click(self) -> None:
        if self._wb_click_timer.isActive():
            self._wb_click_timer.stop()
            self.overlay.open_board_color_dialog()
        else:
            self._wb_click_timer.start()

    def update_whiteboard_button_state(self, active: bool) -> None:
        self.btn_whiteboard.setObjectName("whiteboardButtonActive" if active else "")
        self.style().polish(self.btn_whiteboard)

    def enterEvent(self, event) -> None:
        self.overlay.raise_toolbar()
        super().enterEvent(event)

    def showEvent(self, event) -> None:  # type: ignore[override]
        super().showEvent(event)
        ensure_widget_within_screen(self)


# ---------- 叠加层（画笔/白板） ----------
class OverlayWindow(QWidget):
    def __init__(self, settings_manager: SettingsManager) -> None:
        super().__init__(None, Qt.WindowType.FramelessWindowHint | Qt.WindowType.WindowStaysOnTopHint)
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        self.settings_manager = settings_manager
        s = self.settings_manager.load_settings().get("Paint", {})
        self.pen_size = int(s.get("brush_size", "12"))
        self.pen_color = QColor(s.get("brush_color", "#ff0000"))
        self.mode = "brush"
        self.current_shape: Optional[str] = None
        self.shape_start_point: Optional[QPoint] = None
        self.drawing = False
        self.last_point = QPointF(); self.prev_point = QPointF()
        self.last_width = float(self.pen_size); self.last_time = time.time()
        self._last_brush_color = QColor(self.pen_color)
        self._last_brush_size = max(1, self.pen_size)
        self._last_draw_mode = "brush"
        self._last_shape_type: Optional[str] = None
        self._restoring_tool = False
        self._eraser_last_point: Optional[QPoint] = None
        self._stroke_points: List[QPointF] = []
        self._stroke_timestamps: List[float] = []
        self._stroke_speed: float = 0.0
        self._stroke_last_midpoint: Optional[QPointF] = None
        self.whiteboard_active = False
        self.whiteboard_color = QColor(0, 0, 0, 0); self.last_board_color = QColor("#ffffff")
        self.cursor_pixmap = QPixmap()
        self.setFocusPolicy(Qt.FocusPolicy.StrongFocus)
        self.setMouseTracking(True)

        self._build_scene()
        self.history: List[QPixmap] = []
        self._history_limit = 30
        self.toolbar = FloatingToolbar(self, self.settings_manager)
        self.set_mode("brush", initial=True)
        self.toolbar.update_undo_state(False)

    def raise_toolbar(self) -> None:
        if getattr(self, "toolbar", None) is not None:
            self.toolbar.show()
            self.toolbar.raise_()

    def _build_scene(self) -> None:
        virtual = QRect()
        for screen in QApplication.screens():
            virtual = virtual.united(screen.geometry())
        self.setGeometry(virtual)
        self.canvas = QPixmap(self.size()); self.canvas.fill(Qt.GlobalColor.transparent)
        self.temp_canvas = QPixmap(self.size()); self.temp_canvas.fill(Qt.GlobalColor.transparent)

    def show_overlay(self) -> None:
        self.show(); self.toolbar.show(); self.raise_toolbar()
        self.set_mode(self.mode, self.current_shape)

    def hide_overlay(self) -> None:
        self.hide(); self.toolbar.hide()
        self.save_settings(); self.save_window_position()

    def open_pen_settings(self) -> None:
        pm, ps = self.mode, self.current_shape
        d = PenSettingsDialog(self.toolbar, self.pen_size, self.pen_color.name())
        if d.exec():
            self.pen_size, self.pen_color = d.get_settings()
            self.update_cursor()
        self.set_mode(pm, ps)
        self.raise_toolbar()

    def open_board_color_dialog(self) -> None:
        d = BoardColorDialog(self.toolbar)
        if d.exec():
            c = d.get_color()
            if c:
                self.last_board_color = c
                self.whiteboard_color = c
                self.whiteboard_active = True
                self.toolbar.update_whiteboard_button_state(True)
                self._update_visibility_for_mode(initial=False)
                self.raise_toolbar()
                self.update()

    def toggle_whiteboard(self) -> None:
        self.whiteboard_active = not self.whiteboard_active
        self.whiteboard_color = self.last_board_color if self.whiteboard_active else QColor(0, 0, 0, 0)
        self._update_visibility_for_mode(initial=False)
        self.raise_toolbar()
        self.toolbar.update_whiteboard_button_state(self.whiteboard_active)
        self.update()

    def set_mode(self, mode: str, shape_type: Optional[str] = None, *, initial: bool = False) -> None:
        prev_mode = getattr(self, "mode", None)
        self.mode = mode
        if shape_type is not None or mode != "shape":
            self.current_shape = shape_type
        if mode != "shape":
            self.shape_start_point = None
        if self.mode != "eraser":
            self._eraser_last_point = None
        if self.mode in {"brush", "shape"} and not self._restoring_tool:
            self._update_last_tool_snapshot()
        self._update_visibility_for_mode(initial=initial)
        if not initial:
            self.raise_toolbar()
        self.update_toolbar_state()
        self.update_cursor()

    def update_toolbar_state(self) -> None:
        if not getattr(self, 'toolbar', None):
            return
        self.toolbar.update_tool_states(self.mode, self.pen_color)

    def _update_last_tool_snapshot(self) -> None:
        if self.pen_color.isValid():
            self._last_brush_color = QColor(self.pen_color)
        if self.pen_size > 0:
            self._last_brush_size = max(1, int(self.pen_size))
        if self.mode in {"brush", "shape"}:
            self._last_draw_mode = self.mode
            if self.mode == "shape":
                self._last_shape_type = self.current_shape

    def _restore_last_tool(self) -> None:
        if isinstance(self._last_brush_color, QColor) and self._last_brush_color.isValid():
            self.pen_color = QColor(self._last_brush_color)
        if isinstance(self._last_brush_size, int) and self._last_brush_size > 0:
            self.pen_size = max(1, int(self._last_brush_size))
        target_mode = self._last_draw_mode if self._last_draw_mode in {"brush", "shape"} else "brush"
        target_shape = self._last_shape_type if target_mode == "shape" else None
        self._restoring_tool = True
        try:
            self.set_mode(target_mode, shape_type=target_shape)
        finally:
            self._restoring_tool = False
        self._update_last_tool_snapshot()

    def toggle_eraser_mode(self) -> None:
        """切换橡皮模式；再次点击会恢复上一次的画笔配置。"""
        if self.mode == "eraser":
            self._restore_last_tool()
        else:
            self._update_last_tool_snapshot()
            self.set_mode("eraser")

    def toggle_cursor_mode(self) -> None:
        """切换光标模式；再次点击恢复最近的画笔或图形设置。"""
        if self.mode == "cursor":
            self._restore_last_tool()
            return
        self._update_last_tool_snapshot()
        self.set_mode("cursor")

    def update_cursor(self) -> None:
        if self.mode == "cursor":
            self.setCursor(Qt.CursorShape.ArrowCursor); return
        if self.mode == "shape":
            self.setCursor(Qt.CursorShape.CrossCursor); return
        d = max(10, int(self.pen_size * (3.2 if self.mode == "eraser" else 2.2)))
        self.cursor_pixmap = QPixmap(d, d); self.cursor_pixmap.fill(Qt.GlobalColor.transparent)
        p = QPainter(self.cursor_pixmap); p.setRenderHint(QPainter.RenderHint.Antialiasing)
        if self.mode == "eraser":
            p.setBrush(QBrush(Qt.GlobalColor.white)); p.setPen(QPen(QColor("#555"), 2))
        else:
            p.setBrush(QBrush(self.pen_color)); p.setPen(QPen(Qt.GlobalColor.black, 2))
        p.drawEllipse(1, 1, d - 2, d - 2); p.end()
        self.setCursor(QCursor(self.cursor_pixmap, d // 2, d // 2))

    # ---- 系统级穿透 ----
    def _apply_input_passthrough(self, enabled: bool) -> None:
        # 开启后可让鼠标穿透画布，方便回到课件操作
        self.setAttribute(Qt.WidgetAttribute.WA_TransparentForMouseEvents, enabled)
        self.setWindowFlag(Qt.WindowType.WindowTransparentForInput, enabled)
        if self.isVisible():
            super().show()  # 立即生效

    def _update_visibility_for_mode(self, *, initial: bool = False) -> None:
        passthrough = (self.mode == "cursor") and (not self.whiteboard_active)
        self._apply_input_passthrough(passthrough)
        if initial:
            return
        if not passthrough:
            self.show(); self.raise_()
            self.setFocus(Qt.FocusReason.ActiveWindowFocusReason)
        else:
            if not self.isVisible():
                self.show()

    def _push_history(self) -> None:
        if not isinstance(self.canvas, QPixmap):
            return
        self.history.append(self.canvas.copy())
        if len(self.history) > self._history_limit:
            self.history.pop(0)
        self._update_undo_button()

    def _update_undo_button(self) -> None:
        if getattr(self, "toolbar", None):
            self.toolbar.update_undo_state(bool(self.history))

    def clear_all(self) -> None:
        """清除整块画布，同时根据需要恢复画笔模式。"""
        restore_needed = self.mode not in {"brush", "shape"}
        self._push_history()
        self.canvas.fill(Qt.GlobalColor.transparent)
        self.temp_canvas.fill(Qt.GlobalColor.transparent)
        self.update()
        self._eraser_last_point = None
        if restore_needed:
            self._restore_last_tool()
        else:
            if self.mode in {"brush", "shape"}:
                self._update_last_tool_snapshot()
            self.raise_toolbar()
        self._update_undo_button()

    def use_brush_color(self, color_hex: str) -> None:
        """根据传入的十六进制颜色值启用画笔模式。"""
        color = QColor(color_hex)
        if not color.isValid():
            return
        self.pen_color = color
        self.set_mode("brush")

    def undo_last_action(self) -> None:
        if not self.history:
            return
        last = self.history.pop()
        if isinstance(last, QPixmap):
            self.canvas = last
        else:
            self._update_undo_button()
            return
        self.temp_canvas.fill(Qt.GlobalColor.transparent)
        self.drawing = False
        self.update()
        self.raise_toolbar()
        self._update_undo_button()

    def save_settings(self) -> None:
        settings = self.settings_manager.load_settings()
        paint = settings.get("Paint", {})
        paint["brush_size"] = str(self.pen_size)
        paint["brush_color"] = self.pen_color.name()
        settings["Paint"] = paint
        self.settings_manager.save_settings(settings)

    def save_window_position(self) -> None:
        settings = self.settings_manager.load_settings()
        paint = settings.get("Paint", {})
        pos = self.toolbar.pos()
        paint["x"] = str(pos.x()); paint["y"] = str(pos.y())
        settings["Paint"] = paint
        self.settings_manager.save_settings(settings)

    # ---- 画图事件 ----
    def mousePressEvent(self, e) -> None:
        if e.button() == Qt.MouseButton.LeftButton and self.mode != "cursor":
            self._push_history()
            self.drawing = True
            pointf = e.position(); self.last_point = pointf; self.prev_point = pointf
            now = time.time()
            self.last_time = now
            self._stroke_points = [QPointF(pointf)]
            self._stroke_timestamps = [now]
            self._stroke_speed = 0.0
            self._stroke_last_midpoint = QPointF(pointf)
            self.last_width = max(1.0, float(self.pen_size) * 0.4)
            self.shape_start_point = e.pos() if self.mode == "shape" else None
            self._eraser_last_point = e.pos() if self.mode == "eraser" else None
            self.raise_toolbar()
            e.accept()
        super().mousePressEvent(e)

    def mouseMoveEvent(self, e) -> None:
        if self.drawing and self.mode != "cursor":
            p = e.pos(); pf = e.position()
            if self.mode == "brush": self._draw_brush_line(pf)
            elif self.mode == "eraser": self._erase_at(p)
            elif self.mode == "shape" and self.current_shape: self._draw_shape_preview(p)
            self.update()
            self.raise_toolbar()
        super().mouseMoveEvent(e)

    def mouseReleaseEvent(self, e) -> None:
        if e.button() == Qt.MouseButton.LeftButton and self.drawing:
            if self.mode == "shape" and self.current_shape: self._draw_shape_final(e.pos())
            self.drawing = False; self.shape_start_point = None; self.update()
            if self.mode == "eraser":
                self._eraser_last_point = None
            if self.mode == "brush":
                self._stroke_points.clear(); self._stroke_timestamps.clear()
                self._stroke_last_midpoint = None; self._stroke_speed = 0.0
            self.raise_toolbar()
        super().mouseReleaseEvent(e)

    def keyPressEvent(self, e: QKeyEvent) -> None:
        if e.key() == Qt.Key.Key_Escape:
            self.set_mode("cursor"); return
        super().keyPressEvent(e)

    def _draw_brush_line(self, cur: QPointF) -> None:
        now = time.time()
        cur_point = QPointF(cur)
        if not self._stroke_points:
            self._stroke_points = [QPointF(cur_point)]
            self._stroke_timestamps = [now]
            self.prev_point = QPointF(cur_point)
            self.last_point = QPointF(cur_point)
            self._stroke_last_midpoint = QPointF(cur_point)
            self.last_time = now
            return

        last_point = QPointF(self._stroke_points[-1])
        self._stroke_points.append(cur_point)
        self._stroke_timestamps.append(now)
        if len(self._stroke_points) > 5:
            self._stroke_points.pop(0)
            self._stroke_timestamps.pop(0)

        elapsed = max(1e-4, now - self._stroke_timestamps[-2])
        distance = math.hypot(cur_point.x() - last_point.x(), cur_point.y() - last_point.y())
        speed = distance / elapsed
        self._stroke_speed = self._stroke_speed * 0.65 + speed * 0.35

        curvature = 0.0
        if len(self._stroke_points) >= 3:
            p0 = self._stroke_points[-3]
            p1 = self._stroke_points[-2]
            p2 = self._stroke_points[-1]
            v1x, v1y = p1.x() - p0.x(), p1.y() - p0.y()
            v2x, v2y = p2.x() - p1.x(), p2.y() - p1.y()
            denom = math.hypot(v1x, v1y) * math.hypot(v2x, v2y)
            if denom > 1e-5:
                curvature = abs(v1x * v2y - v1y * v2x) / denom

        pressure = min(1.0, (now - self.last_time) * 4.5)
        self.last_time = now

        base_size = float(max(1, self.pen_size))
        speed_scale = 1.0 / (1.0 + self._stroke_speed / (base_size * 18.0 + 36.0))
        curve_scale = min(1.0, curvature * base_size * 0.65)
        target_w = base_size * (0.42 + 0.45 * speed_scale + 0.25 * curve_scale)
        target_w *= 1.0 + 0.3 * pressure
        cur_w = self.last_width * 0.55 + target_w * 0.45

        last_mid = QPointF(self._stroke_last_midpoint) if self._stroke_last_midpoint else QPointF(last_point)
        current_mid = (last_point + cur_point) / 2.0

        path = QPainterPath(last_mid)
        path.quadTo(last_point, current_mid)
        painter = QPainter(self.canvas)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        fade = QColor(self.pen_color)
        fade_alpha = int(max(30, min(200, 200 * speed_scale)))
        fade.setAlpha(fade_alpha)
        painter.setPen(QPen(fade, cur_w * 1.35, Qt.PenStyle.SolidLine, Qt.PenCapStyle.RoundCap, Qt.PenJoinStyle.RoundJoin))
        painter.drawPath(path)
        painter.setPen(QPen(self.pen_color, cur_w, Qt.PenStyle.SolidLine, Qt.PenCapStyle.RoundCap, Qt.PenJoinStyle.RoundJoin))
        painter.drawPath(path)
        painter.end()

        self.prev_point = QPointF(last_point)
        self.last_point = QPointF(cur_point)
        self._stroke_last_midpoint = QPointF(current_mid)
        self.last_width = cur_w

    def _erase_at(self, pos) -> None:
        current = QPointF(pos) if isinstance(pos, QPointF) else QPointF(QPoint(pos))
        if not isinstance(self._eraser_last_point, QPoint):
            self._eraser_last_point = current.toPoint()
        start_point = QPointF(self._eraser_last_point)
        path = QPainterPath(start_point)
        path.lineTo(current)

        radius = max(8.0, float(self.pen_size) * 1.6)
        stroker = QPainterPathStroker()
        stroker.setWidth(max(12.0, radius * 2.0))
        stroker.setCapStyle(Qt.PenCapStyle.RoundCap)
        stroker.setJoinStyle(Qt.PenJoinStyle.RoundJoin)
        erase_path = stroker.createStroke(path)
        erase_path.addEllipse(current, radius, radius)

        painter = QPainter(self.canvas)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        painter.setCompositionMode(QPainter.CompositionMode.CompositionMode_Clear)
        painter.fillPath(erase_path, QColor(0, 0, 0, 0))
        painter.end()

        self._eraser_last_point = current.toPoint()

    def _draw_shape_preview(self, end_point) -> None:
        if not self.shape_start_point: return
        self.temp_canvas.fill(Qt.GlobalColor.transparent)
        p = QPainter(self.temp_canvas); p.setRenderHint(QPainter.RenderHint.Antialiasing)
        pen = QPen(self.pen_color, self.pen_size)
        if self.current_shape and "dashed" in self.current_shape: pen.setStyle(Qt.PenStyle.DashLine)
        p.setPen(pen); self._draw_shape(p, self.shape_start_point, end_point); p.end()
        self.raise_toolbar()

    def _draw_shape_final(self, end_point) -> None:
        if not self.shape_start_point: return
        p = QPainter(self.canvas); p.setRenderHint(QPainter.RenderHint.Antialiasing)
        pen = QPen(self.pen_color, self.pen_size)
        if self.current_shape and "dashed" in self.current_shape: pen.setStyle(Qt.PenStyle.DashLine)
        p.setPen(pen); self._draw_shape(p, self.shape_start_point, end_point); p.end()
        self.temp_canvas.fill(Qt.GlobalColor.transparent)
        self.raise_toolbar()

    def _draw_shape(self, painter: QPainter, start_point, end_point) -> None:
        rect = QRect(start_point, end_point)
        shape = (self.current_shape or "line").replace("dashed_", "")
        if shape == "rect": painter.drawRect(rect.normalized())
        elif shape == "circle": painter.drawEllipse(rect.normalized())
        else: painter.drawLine(start_point, end_point)

    def paintEvent(self, e) -> None:
        p = QPainter(self)
        if self.whiteboard_active:
            p.fillRect(self.rect(), self.whiteboard_color)
        else:
            p.fillRect(self.rect(), QColor(0, 0, 0, 1))
        p.drawPixmap(0, 0, self.canvas)
        if self.drawing and self.mode == "shape": p.drawPixmap(0, 0, self.temp_canvas)
        p.end()

    def showEvent(self, e) -> None:
        super().showEvent(e)
        self.raise_toolbar()

    def closeEvent(self, e) -> None:
        self.save_settings(); self.save_window_position()
        super().closeEvent(e)


# ---------- 语音 ----------
class TTSManager(QObject):
    """简单封装语音播报，优先使用 pyttsx3，必要时回退到 PowerShell。"""

    def __init__(self, preferred_voice_id: str = "", parent: Optional[QObject] = None) -> None:
        super().__init__(parent)
        self.engine = None
        self.voice_ids: List[str] = []
        self.default_voice_id = ""
        self.current_voice_id = ""
        self.failure_reason = ""
        self.failure_suggestions: List[str] = []
        self.supports_voice_selection = False
        self._mode: str = "none"
        self._powershell_path = ""
        self._powershell_busy = False
        self._queue: Queue[str] = Queue()
        self._timer = QTimer(self)
        self._timer.timeout.connect(self._pump)
        if pyttsx3 is not None:
            try:
                init_kwargs = {"driverName": "sapi5"} if sys.platform == "win32" else {}
                self.engine = pyttsx3.init(**init_kwargs)
                voices = self.engine.getProperty("voices") or []
                self.voice_ids = [v.id for v in voices if getattr(v, "id", None)]
                if not self.voice_ids:
                    self._record_failure("未检测到任何可用的发音人")
                    self.engine = None
                else:
                    self.default_voice_id = self.voice_ids[0]
                    self.current_voice_id = (
                        preferred_voice_id if preferred_voice_id in self.voice_ids else self.default_voice_id
                    )
                    if self.current_voice_id:
                        try:
                            self.engine.setProperty("voice", self.current_voice_id)
                        except Exception as exc:
                            self._record_failure("无法设置默认发音人", exc)
                            self.engine = None
                    if self.engine is not None:
                        self.supports_voice_selection = True
                        self._mode = "pyttsx3"
                        self.engine.startLoop(False)
                        self._timer.start(100)
                        return
            except Exception as exc:
                self._record_failure("初始化语音引擎失败", exc)
                self.engine = None
        self._init_powershell_fallback()

    @property
    def available(self) -> bool:
        return self._mode in {"pyttsx3", "powershell"}

    def diagnostics(self) -> tuple[str, List[str]]:
        return self.failure_reason, list(self.failure_suggestions)

    def _init_powershell_fallback(self) -> None:
        if sys.platform != "win32":
            return
        path = shutil.which("pwsh") or shutil.which("powershell")
        if not path:
            system_root = os.environ.get("SystemRoot") or os.environ.get("WINDIR")
            candidate_paths: List[str] = []
            if system_root:
                candidate_paths.extend(
                    [
                        os.path.join(system_root, "System32", "WindowsPowerShell", "v1.0", "pwsh.exe"),
                        os.path.join(system_root, "System32", "WindowsPowerShell", "v1.0", "powershell.exe"),
                        os.path.join(system_root, "SysWOW64", "WindowsPowerShell", "v1.0", "powershell.exe"),
                    ]
                )
            candidate_paths.extend(
                [
                    os.path.join("C:\\Program Files\\PowerShell\\7", "pwsh.exe"),
                    os.path.join("C:\\Program Files\\PowerShell\\6", "pwsh.exe"),
                ]
            )
            for candidate in candidate_paths:
                if candidate and os.path.exists(candidate):
                    path = candidate
                    break
        if not path:
            if not self.failure_reason:
                self._record_failure("未检测到 PowerShell，可用的语音播报方式受限")
            return
        self._powershell_path = os.path.abspath(path)
        self.engine = object()
        self.voice_ids = []
        self.default_voice_id = ""
        self.current_voice_id = ""
        self.supports_voice_selection = False
        self.failure_reason = ""
        self.failure_suggestions = []
        self._mode = "powershell"
        self._timer.start(120)

    def _record_failure(self, fallback: str, exc: Optional[Exception] = None) -> None:
        message = ""
        if exc is not None:
            message = str(exc).strip()
        if message and message not in fallback:
            reason = f"{fallback}：{message}"
        else:
            reason = fallback
        self.failure_reason = reason
        suggestions: List[str] = []
        lower = message.lower()
        if "comtypes" in lower:
            suggestions.append("请安装 comtypes（pip install comtypes）后重新启动程序。")
        if "pywin32" in lower or "win32" in lower:
            suggestions.append("请安装 pywin32（pip install pywin32）后重新启动程序。")
        platform_hint = []
        if sys.platform == "win32":
            platform_hint.append("请确认 Windows 已启用 SAPI5 中文语音包。")
        elif sys.platform == "darwin":
            platform_hint.append("请在系统“辅助功能 -> 语音”中启用所需的语音包。")
        else:
            platform_hint.append("请确保系统已安装可用的语音引擎（如 espeak）并重新启动程序。")
        platform_hint.append("可尝试重新安装 pyttsx3 或检查语音服务状态后重启软件。")
        for hint in platform_hint:
            if hint not in suggestions:
                suggestions.append(hint)
        self.failure_suggestions = suggestions

    def set_voice(self, voice_id: str) -> None:
        if not self.supports_voice_selection:
            return
        if voice_id in self.voice_ids:
            self.current_voice_id = voice_id
            if self.engine:
                try:
                    self.engine.setProperty("voice", voice_id)
                except Exception:
                    pass

    def speak(self, text: str) -> None:
        if not self.available:
            return
        while not self._queue.empty():
            try:
                self._queue.get_nowait()
            except Empty:
                break
        self._queue.put(text)

    def _pump(self) -> None:
        if self._mode == "pyttsx3":
            if not self.engine:
                return
            try:
                text = self._queue.get_nowait()
                self.engine.stop()
                if self.current_voice_id:
                    self.engine.setProperty("voice", self.current_voice_id)
                self.engine.say(text)
            except Empty:
                pass
            try:
                self.engine.iterate()
            except Exception as exc:
                self._record_failure("语音引擎运行异常", exc)
                self.shutdown()
        elif self._mode == "powershell":
            if self._powershell_busy:
                return
            try:
                text = self._queue.get_nowait()
            except Empty:
                return
            self._powershell_busy = True
            worker = threading.Thread(target=self._run_powershell_speech, args=(text,), daemon=True)
            worker.start()

    def _run_powershell_speech(self, text: str) -> None:
        try:
            if not text or not self._powershell_path:
                return
            payload = base64.b64encode(text.encode("utf-8")).decode("ascii")
            script = (
                "$msg = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String('" + payload + "'));"
                "Add-Type -AssemblyName System.Speech;"
                "$sp = New-Object System.Speech.Synthesis.SpeechSynthesizer;"
                "$sp.Speak($msg);"
            )
            startupinfo = None
            if os.name == "nt":
                startupinfo = subprocess.STARTUPINFO()
                startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            subprocess.run(
                [self._powershell_path, "-NoLogo", "-NonInteractive", "-NoProfile", "-Command", script],
                check=True,
                timeout=30,
                startupinfo=startupinfo,
            )
        except Exception as exc:
            self._record_failure("PowerShell 语音播报失败", exc)
            QTimer.singleShot(0, self.shutdown)
        finally:
            self._powershell_busy = False

    def shutdown(self) -> None:
        if self._mode == "pyttsx3" and self.engine:
            try:
                self.engine.endLoop()
            except Exception:
                pass
            try:
                self.engine.stop()
            except Exception:
                pass
        self.engine = None
        self._mode = "none"
        self._powershell_busy = False
        self._timer.stop()


# ---------- 点名/计时 ----------
class CountdownSettingsDialog(QDialog):
    """设置倒计时分钟和秒数的小窗口。"""

    def __init__(self, parent: Optional[QWidget], minutes: int, seconds: int) -> None:
        super().__init__(parent)
        self.setWindowTitle("设置倒计时")
        self.setWindowFlag(Qt.WindowType.WindowStaysOnTopHint, True)
        self.result: Optional[tuple[int, int]] = None

        layout = QVBoxLayout(self); layout.setContentsMargins(10, 10, 10, 10); layout.setSpacing(6)

        ml = QHBoxLayout(); ml.addWidget(QLabel("分钟 (0-25):"))
        self.minutes_spin = QSpinBox(); self.minutes_spin.setRange(0, 300); self.minutes_spin.setValue(max(0, min(300, minutes)))
        minute_slider = QSlider(Qt.Orientation.Horizontal); minute_slider.setRange(0, 25)
        minute_slider.setValue(min(self.minutes_spin.value(), minute_slider.maximum()))
        minute_slider.valueChanged.connect(self.minutes_spin.setValue)

        def sync(v: int, slider=minute_slider):
            if v <= slider.maximum():
                prev = slider.blockSignals(True); slider.setValue(v); slider.blockSignals(prev)
        self.minutes_spin.valueChanged.connect(sync)
        ml.addWidget(self.minutes_spin); layout.addLayout(ml); layout.addWidget(minute_slider)

        sl = QHBoxLayout(); sl.addWidget(QLabel("秒 (0-59):"))
        self.seconds_spin = QSpinBox(); self.seconds_spin.setRange(0, 59); self.seconds_spin.setValue(max(0, min(59, seconds)))
        second_slider = QSlider(Qt.Orientation.Horizontal); second_slider.setRange(0, 59)
        second_slider.setValue(self.seconds_spin.value())
        second_slider.valueChanged.connect(self.seconds_spin.setValue); self.seconds_spin.valueChanged.connect(second_slider.setValue)
        sl.addWidget(self.seconds_spin); layout.addLayout(sl); layout.addWidget(second_slider)

        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        buttons.accepted.connect(self._accept); buttons.rejected.connect(self.reject); layout.addWidget(buttons)
        self.setFixedSize(self.sizeHint())

    def _accept(self) -> None:
        self.result = (self.minutes_spin.value(), self.seconds_spin.value()); self.accept()

    def showEvent(self, event) -> None:  # type: ignore[override]
        super().showEvent(event)
        ensure_widget_within_screen(self)


class ClickableFrame(QFrame):
    clicked = pyqtSignal()
    def mousePressEvent(self, e) -> None:
        if e.button() == Qt.MouseButton.LeftButton: self.clicked.emit()
        super().mousePressEvent(e)


def preferred_calligraphy_font(default: str = "Microsoft YaHei UI") -> str:
    """返回系统中更具书法风格的字体，若不可用则回退到默认字体。"""

    try:
        families = set(QFontDatabase().families())
    except Exception:
        return default
    for candidate in ("楷体", "KaiTi", "Kaiti SC", "STKaiti", "DFKai-SB", "FZKai-Z03S"):
        if candidate in families:
            return candidate
    return default


class StudentListDialog(QDialog):
    def __init__(self, parent: Optional[QWidget], students: List[tuple[str, str, int]]) -> None:
        super().__init__(parent)
        self.setWindowTitle("学生名单")
        self.setModal(True)
        self._selected_index: Optional[int] = None

        layout = QVBoxLayout(self)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(12)

        grid = QGridLayout()
        grid.setContentsMargins(0, 0, 0, 0)
        grid.setHorizontalSpacing(6)
        grid.setVerticalSpacing(6)
        grid.setAlignment(Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignHCenter)

        button_font = QFont("Microsoft YaHei UI", 10, QFont.Weight.Medium)
        metrics = QFontMetrics(button_font)
        max_text = max((metrics.horizontalAdvance(f"{sid} {name}") for sid, name, _ in students), default=120)
        min_button_width = max(120, max_text + 24)
        button_height = max(36, metrics.height() + 14)

        screen = QApplication.primaryScreen()
        available_width = screen.availableGeometry().width() if screen else 1280
        max_width_per_button = max(96, int((available_width * 0.9 - 40) / 10))
        button_width = min(min_button_width, max_width_per_button)
        button_size = QSize(button_width, button_height)

        total_rows = max(1, math.ceil(len(students) / 10))

        for column in range(10):
            grid.setColumnStretch(column, 0)
            grid.setColumnMinimumWidth(column, button_width)

        for row in range(total_rows):
            grid.setRowStretch(row, 0)
            grid.setRowMinimumHeight(row, button_height)

        for position, (sid, name, data_index) in enumerate(students):
            row = position // 10
            column = position % 10
            button = QPushButton(f"{sid} {name}")
            button.setFont(button_font)
            button.setCursor(Qt.CursorShape.PointingHandCursor)
            button.setFixedSize(button_size)
            button.setSizePolicy(QSizePolicy.Policy.Fixed, QSizePolicy.Policy.Fixed)
            button.clicked.connect(lambda _checked=False, value=data_index: self._select_student(value))
            grid.addWidget(button, row, column, Qt.AlignmentFlag.AlignCenter)

        layout.addLayout(grid)

        box = QDialogButtonBox(QDialogButtonBox.StandardButton.Close, parent=self)
        box.rejected.connect(self.reject)
        layout.addWidget(box)

        if screen is not None:
            available = screen.availableGeometry()
            rows = total_rows
            h_spacing = grid.horizontalSpacing() if grid.horizontalSpacing() is not None else 6
            v_spacing = grid.verticalSpacing() if grid.verticalSpacing() is not None else 6
            preferred_width = min(int(available.width() * 0.9), button_width * 10 + h_spacing * 9 + 40)
            preferred_height = min(
                int(available.height() * 0.85),
                rows * button_height + max(0, rows - 1) * v_spacing + box.sizeHint().height() + 48,
            )
            self.resize(preferred_width, preferred_height)

    def _select_student(self, index: int) -> None:
        self._selected_index = index
        self.accept()

    @property
    def selected_index(self) -> Optional[int]:
        return self._selected_index


class ScoreboardDialog(QDialog):
    ORDER_RANK = "rank"
    ORDER_ID = "id"

    @dataclass
    class _CardMetrics:
        count: int
        columns: int
        rows: int
        card_width: int
        card_height: int
        padding_h: int
        padding_v: int
        inner_spacing: int
        font_size: int
        horizontal_spacing: int
        vertical_spacing: int

    def __init__(
        self,
        parent: Optional[QWidget],
        students: List[tuple[str, str, int]],
        order: str = "rank",
        order_changed: Optional[Callable[[str], None]] = None,
    ) -> None:
        super().__init__(parent)
        self.setWindowTitle("成绩展示")
        self.setModal(True)
        self.setObjectName("ScoreboardDialog")
        self._pending_maximize = True

        self.students = list(students)
        self._order_changed_callback = order_changed
        self._order = order if order in {self.ORDER_RANK, self.ORDER_ID} else self.ORDER_RANK
        self._grid_row_count = 0
        self._grid_column_count = 0
        self._card_metrics: Optional[ScoreboardDialog._CardMetrics] = None
        self._card_metrics_key: Optional[tuple[int, int, int]] = None

        calligraphy_font = preferred_calligraphy_font()
        self._calligraphy_font = calligraphy_font

        layout = QVBoxLayout(self)
        layout.setContentsMargins(32, 32, 32, 32)
        layout.setSpacing(24)

        title = QLabel("成绩展示")
        title.setObjectName("ScoreboardHeader")
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title.setFont(QFont(calligraphy_font, 44, QFont.Weight.Bold))
        layout.addWidget(title)

        order_layout = QHBoxLayout()
        order_layout.setContentsMargins(0, 0, 0, 0)
        order_layout.setSpacing(12)
        order_label = QLabel("排序方式：")
        order_label.setFont(QFont(calligraphy_font, 28, QFont.Weight.Bold))
        order_layout.addWidget(order_label, 0, Qt.AlignmentFlag.AlignLeft)

        button_font = QFont(calligraphy_font, 22, QFont.Weight.Bold)
        self.order_button_group = QButtonGroup(self)
        self.order_button_group.setExclusive(True)
        self.order_buttons: Dict[str, QPushButton] = {}
        for key, text in ((self.ORDER_RANK, "按排名"), (self.ORDER_ID, "按学号")):
            button = QPushButton(text)
            button.setCheckable(True)
            button.setCursor(Qt.CursorShape.PointingHandCursor)
            button.setProperty("class", "orderButton")
            button.setFont(button_font)
            button.setFixedSize(140, 44)
            self.order_button_group.addButton(button)
            button.clicked.connect(lambda _checked=False, value=key: self._on_order_button_clicked(value))
            order_layout.addWidget(button, 0, Qt.AlignmentFlag.AlignLeft)
            self.order_buttons[key] = button
        order_layout.addStretch(1)
        layout.addLayout(order_layout)

        grid_container = QWidget()
        grid_container.setObjectName("ScoreboardGridContainer")
        grid_container.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        self.grid_layout = QGridLayout(grid_container)
        self.grid_layout.setContentsMargins(18, 18, 18, 18)
        self.grid_layout.setHorizontalSpacing(20)
        self.grid_layout.setVerticalSpacing(20)
        layout.addWidget(grid_container, 1)

        box = QDialogButtonBox(QDialogButtonBox.StandardButton.Close, parent=self)
        box.setFont(QFont(calligraphy_font, 22, QFont.Weight.Bold))
        close_button = box.button(QDialogButtonBox.StandardButton.Close)
        if close_button is not None:
            close_button.setText("关闭")
            close_button.setMinimumHeight(42)
            close_button.setCursor(Qt.CursorShape.PointingHandCursor)
            close_button.setFont(QFont(calligraphy_font, 22, QFont.Weight.Bold))
        box.rejected.connect(self.reject)
        layout.addWidget(box)

        self.setStyleSheet(
            "#ScoreboardDialog {"
            "    background-color: qlineargradient(x1:0, y1:0, x2:0, y2:1,"
            "        stop:0 #f7f9fc, stop:1 #e3edff);"
            "}"
            "#ScoreboardGridContainer {"
            "    background-color: rgba(255, 255, 255, 0.85);"
            "    border-radius: 24px;"
            "}"
            "QLabel#ScoreboardHeader {"
            "    color: #0b3d91;"
            "}"
            "QLabel[class=\"scoreboardName\"] {"
            "    color: #103d73;"
            "}"
            "QLabel[class=\"scoreboardScore\"] {"
            "    color: #103d73;"
            "}"
            "QWidget[class=\"scoreboardWrapper\"] {"
            "    background-color: rgba(255, 255, 255, 0.95);"
            "    border-radius: 18px;"
            "    border: 1px solid rgba(16, 61, 115, 0.12);"
            "}"
            "QPushButton[class=\"orderButton\"] {"
            "    background-color: rgba(255, 255, 255, 0.88);"
            "    border-radius: 22px;"
            "    border: 1px solid rgba(16, 61, 115, 0.24);"
            "    padding: 4px 18px;"
            "    color: #0b3d91;"
            "}"
            "QPushButton[class=\"orderButton\"]:hover {"
            "    border-color: #1a73e8;"
            "    background-color: rgba(26, 115, 232, 0.12);"
            "}"
            "QPushButton[class=\"orderButton\"]:checked {"
            "    background-color: #1a73e8;"
            "    border-color: #1a73e8;"
            "    color: #ffffff;"
            "}"
        )

        screen = QApplication.primaryScreen()
        self._available_geometry = screen.availableGeometry() if screen is not None else QRect(0, 0, 1920, 1080)

        self._update_order_buttons()
        self._populate_grid()

        if screen is not None:
            self.setGeometry(self._available_geometry)

    def _update_order_buttons(self) -> None:
        for key, button in self.order_buttons.items():
            block = button.blockSignals(True)
            button.setChecked(key == self._order)
            button.blockSignals(block)

    def _on_order_button_clicked(self, order: str) -> None:
        if order not in {self.ORDER_RANK, self.ORDER_ID}:
            self._update_order_buttons()
            return
        if order == self._order:
            self._update_order_buttons()
            return
        self._order = order
        if callable(self._order_changed_callback):
            try:
                self._order_changed_callback(order)
            except Exception:
                pass
        self._update_order_buttons()
        self._populate_grid()

    def _clear_grid(self) -> None:
        while self.grid_layout.count():
            item = self.grid_layout.takeAt(0)
            widget = item.widget()
            if widget is not None:
                widget.deleteLater()
        for row in range(self._grid_row_count):
            self.grid_layout.setRowStretch(row, 0)
            self.grid_layout.setRowMinimumHeight(row, 0)
        for column in range(self._grid_column_count):
            self.grid_layout.setColumnStretch(column, 0)
            self.grid_layout.setColumnMinimumWidth(column, 0)
        self._grid_row_count = 0
        self._grid_column_count = 0

    def _collect_display_candidates(self) -> tuple[List[tuple[int, str, str]], List[str], List[str]]:
        sorted_students = self._sort_students()
        alternate_order = (
            self.ORDER_ID if self._order == self.ORDER_RANK else self.ORDER_RANK
        )
        alternate_students = self._sort_students(alternate_order)

        display_entries: List[tuple[int, str, str]] = []
        display_candidates: List[str] = []
        score_candidates: List[str] = []

        for idx, (sid, name, score) in enumerate(sorted_students):
            display_text = self._format_display_text(idx, sid, name)
            score_text = self._format_score_text(score)
            display_entries.append((idx, display_text, score_text))
            display_candidates.append(display_text)
            score_candidates.append(score_text)

        for idx, (sid, name, score) in enumerate(alternate_students):
            display_candidates.append(
                self._format_display_text_for_order(
                    alternate_order, idx, sid, name
                )
            )
            score_candidates.append(self._format_score_text(score))

        return display_entries, display_candidates, score_candidates

    def _compute_card_metrics(self) -> Optional[_CardMetrics]:
        count = len(self.students)
        if count == 0:
            return None

        available = self._available_geometry
        key = (count, available.width(), available.height())
        if self._card_metrics is not None and self._card_metrics_key == key:
            return self._card_metrics

        columns = 10
        rows = max(1, math.ceil(count / columns))

        usable_width = max(available.width() - 160, 640)
        usable_height = max(available.height() - 240, 520)

        margins = self.grid_layout.contentsMargins()
        horizontal_spacing = max(14, int(usable_width * 0.01))
        vertical_spacing = max(18, int(usable_height * 0.035 / rows))

        spacing_total_x = horizontal_spacing * max(0, columns - 1)
        spacing_total_y = vertical_spacing * max(0, rows - 1)

        available_width_for_cards = (
            usable_width - margins.left() - margins.right() - spacing_total_x
        )
        available_height_for_cards = (
            usable_height - margins.top() - margins.bottom() - spacing_total_y
        )

        per_card_width = max(1.0, available_width_for_cards / columns)
        per_card_height = max(1.0, available_height_for_cards / rows)

        card_width = int(math.floor(per_card_width))
        card_height = int(math.floor(per_card_height))

        if per_card_width >= 120:
            card_width = max(card_width, 120)
        if per_card_height >= 180:
            card_height = max(card_height, 180)

        padding_h = max(12, int(card_width * 0.08))
        padding_v = max(14, int(card_height * 0.1))
        inner_spacing = max(6, int(card_height * 0.045))

        display_entries, display_candidates, score_candidates = self._collect_display_candidates()

        if not display_candidates:
            return None

        calligraphy_font = self._calligraphy_font
        probe_font = QFont(calligraphy_font, 64, QFont.Weight.Bold)
        metrics = QFontMetrics(probe_font)

        widest_display = max(
            display_candidates,
            key=lambda text: metrics.tightBoundingRect(text).width(),
        )
        widest_score = max(
            score_candidates,
            key=lambda text: metrics.tightBoundingRect(text).width(),
        )
        name_label.setFont(QFont(calligraphy_font, font_size, QFont.Weight.Bold))
        name_label.setStyleSheet("margin: 0px; padding: 0px;")
        layout.addWidget(name_label)

        score_label = QLabel(score_text)
        score_label.setProperty("class", "scoreboardScore")
        score_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        score_label.setFont(QFont(calligraphy_font, font_size, QFont.Weight.Bold))
        score_label.setStyleSheet(f"margin-top: {max(6, inner_spacing // 2)}px;")
        layout.addWidget(score_label)

        layout.addStretch(1)
        return wrapper

    def _format_display_text(self, index: int, sid: str, name: str) -> str:
        return self._format_display_text_for_order(self._order, index, sid, name)

    def _format_display_text_for_order(
        self, order: str, index: int, sid: str, name: str
    ) -> str:
        clean_name = (name or "").strip() or "未命名"
        if order == self.ORDER_ID:
            sid_display = str(sid).strip() or "—"
            return f"{sid_display}.{clean_name}"
        return f"{index + 1}.{clean_name}"

    @staticmethod
    def _format_score_text(score: int | float | str) -> str:
        text = "—"
        try:
            value = float(score)
        except (TypeError, ValueError):
            score_str = str(score).strip()
            if score_str and score_str.lower() != "none":
                text = score_str
        else:
            if math.isfinite(value):
                if abs(value - int(value)) < 1e-6:
                    text = str(int(round(value)))
                else:
                    text = f"{value:.2f}".rstrip("0").rstrip(".")
        return f"{text} 分"

    @staticmethod
    def _fit_font_size(
        text: str,
        family: str,
        weight: QFont.Weight,
        max_width: int,
        max_height: int,
        minimum: int,
        maximum: int,
    ) -> int:
        if not text:
            return max(6, min(minimum, maximum))
        if max_width <= 0 or max_height <= 0:
            return max(6, min(minimum, maximum))
        lower = max(6, min(minimum, maximum))
        upper = max(6, max(minimum, maximum))
        for size in range(upper, lower - 1, -1):
            font = QFont(family, size, weight)
            metrics = QFontMetrics(font)
            rect = metrics.tightBoundingRect(text)
            if rect.width() <= max_width and rect.height() <= max_height:
                return size
        return lower

    def _populate_grid(self) -> None:
        self._clear_grid()
        count = len(self.students)
        layout = self.grid_layout
        calligraphy_font = self._calligraphy_font

        usable_name_width = max(60, card_width - 2 * padding_h)
        content_height = max(80, card_height - 2 * padding_v - inner_spacing)
        name_height = int(content_height * 0.58)
        score_height = max(32, content_height - name_height)

        font_upper_bound = int(min(card_width * 0.28, card_height * 0.36))
        font_upper_bound = max(20, font_upper_bound)
        fit_minimum = 14

        name_fit = self._fit_font_size(
            widest_display,
            calligraphy_font,
            QFont.Weight.Bold,
            usable_name_width,
            name_height,
            fit_minimum,
            font_upper_bound,
        )
        score_fit = self._fit_font_size(
            widest_score,
            calligraphy_font,
            QFont.Weight.Bold,
            usable_name_width,
            score_height,
            fit_minimum,
            font_upper_bound,
        )

        final_font_size = max(fit_minimum, min(name_fit, score_fit, font_upper_bound))

        self._card_metrics = ScoreboardDialog._CardMetrics(
            count=count,
            columns=columns,
            rows=rows,
            card_width=card_width,
            card_height=card_height,
            padding_h=padding_h,
            padding_v=padding_v,
            inner_spacing=inner_spacing,
            font_size=final_font_size,
            horizontal_spacing=horizontal_spacing,
            vertical_spacing=vertical_spacing,
        )
        self._card_metrics_key = key
        return self._card_metrics

    def _ensure_metrics(self) -> Optional[_CardMetrics]:
        metrics = self._compute_card_metrics()
        if metrics is not None:
            return metrics
        self._card_metrics = None
        self._card_metrics_key = None
        return None

    def _sort_students(self, order: Optional[str] = None) -> List[tuple[str, str, int]]:
        data = list(self.students)
        current_order = self._order if order is None else order
        if current_order == self.ORDER_ID:
            def _id_key(item: tuple[str, str, int]) -> tuple[int, str, str]:
                sid_text = str(item[0]).strip()
                try:
                    sid_value = int(sid_text)
                except (TypeError, ValueError):
                    sid_value = sys.maxsize
                return (sid_value, sid_text, item[1])

            data.sort(key=_id_key)
        else:
            def _rank_key(item: tuple[str, str, int]) -> tuple[int, int, str]:
                sid_text = str(item[0]).strip()
                try:
                    sid_value = int(sid_text)
                except (TypeError, ValueError):
                    sid_value = sys.maxsize
                return (-item[2], sid_value, item[1])

            data.sort(key=_rank_key)
        return data

    def _create_card(
        self,
        display_text: str,
        score_text: str,
        card_width: int,
        card_height: int,
        font_size: int,
        padding_h: int,
        padding_v: int,
        inner_spacing: int,
    ) -> QWidget:
        calligraphy_font = self._calligraphy_font
        wrapper = QWidget()
        wrapper.setProperty("class", "scoreboardWrapper")
        wrapper.setFixedSize(card_width, card_height)
        wrapper.setSizePolicy(QSizePolicy.Policy.Fixed, QSizePolicy.Policy.Fixed)

        layout = QVBoxLayout(wrapper)
        layout.setContentsMargins(padding_h, padding_v, padding_h, padding_v)
        layout.setSpacing(inner_spacing)

        name_label = QLabel(display_text or "未命名")
        name_label.setProperty("class", "scoreboardName")
        name_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        name_label.setWordWrap(False)
        name_label.setSizePolicy(
            QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed
        )
        name_label.setFont(QFont(calligraphy_font, font_size, QFont.Weight.Bold))
        name_label.setStyleSheet("margin: 0px; padding: 0px;")
        layout.addWidget(name_label)

        score_label = QLabel(score_text)
        score_label.setProperty("class", "scoreboardScore")
        score_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        score_label.setWordWrap(False)
        score_label.setSizePolicy(
            QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed
        )
        score_label.setFont(QFont(calligraphy_font, font_size, QFont.Weight.Bold))
        score_label.setStyleSheet(f"margin-top: {max(6, inner_spacing // 2)}px;")
        layout.addWidget(score_label)

        layout.addStretch(1)
        return wrapper

    def _format_display_text(self, index: int, sid: str, name: str) -> str:
        return self._format_display_text_for_order(self._order, index, sid, name)

    def _format_display_text_for_order(
        self, order: str, index: int, sid: str, name: str
    ) -> str:
        clean_name = (name or "").strip() or "未命名"
        if order == self.ORDER_ID:
            sid_display = str(sid).strip() or "—"
            return f"{sid_display}.{clean_name}"
        return f"{index + 1}.{clean_name}"

    @staticmethod
    def _format_score_text(score: int | float | str) -> str:
        text = "—"
        try:
            value = float(score)
        except (TypeError, ValueError):
            score_str = str(score).strip()
            if score_str and score_str.lower() != "none":
                text = score_str
        else:
            if math.isfinite(value):
                if abs(value - int(value)) < 1e-6:
                    text = str(int(round(value)))
                else:
                    text = f"{value:.2f}".rstrip("0").rstrip(".")
        return f"{text} 分"

    @staticmethod
    def _fit_font_size(
        text: str,
        family: str,
        weight: QFont.Weight,
        max_width: int,
        max_height: int,
        minimum: int,
        maximum: int,
    ) -> int:
        if not text:
            return max(6, min(minimum, maximum))
        if max_width <= 0 or max_height <= 0:
            return max(6, min(minimum, maximum))
        lower = max(6, min(minimum, maximum))
        upper = max(6, max(minimum, maximum))
        for size in range(upper, lower - 1, -1):
            font = QFont(family, size, weight)
            metrics = QFontMetrics(font)
            rect = metrics.tightBoundingRect(text)
            if rect.width() <= max_width and rect.height() <= max_height:
                return size
        return lower

    def _populate_grid(self) -> None:
        self._clear_grid()
        count = len(self.students)
        layout = self.grid_layout
        calligraphy_font = self._calligraphy_font

        if count == 0:
            empty = QLabel("暂无成绩数据")
            empty.setAlignment(Qt.AlignmentFlag.AlignCenter)
            empty.setFont(QFont(calligraphy_font, 28, QFont.Weight.Bold))
            empty.setStyleSheet("color: #103d73;")
            layout.addWidget(empty, 0, 0, Qt.AlignmentFlag.AlignCenter)
            layout.setRowStretch(0, 1)
            layout.setColumnStretch(0, 1)
            self._grid_row_count = 1
            self._grid_column_count = 1
            return

        metrics = self._ensure_metrics()
        if metrics is None:
            return

        layout.setHorizontalSpacing(metrics.horizontal_spacing)
        layout.setVerticalSpacing(metrics.vertical_spacing)

        for column in range(metrics.columns):
            layout.setColumnStretch(column, 1)
            layout.setColumnMinimumWidth(column, metrics.card_width)
        for row in range(metrics.rows):
            layout.setRowStretch(row, 1)
            layout.setRowMinimumHeight(row, metrics.card_height)

        self._grid_row_count = metrics.rows
        self._grid_column_count = metrics.columns

        container = layout.parentWidget()
        if container is not None:
            margins = layout.contentsMargins()
            total_width = (
                metrics.columns * metrics.card_width
                + metrics.horizontal_spacing * max(0, metrics.columns - 1)
                + margins.left()
                + margins.right()
            )
            total_height = (
                metrics.rows * metrics.card_height
                + metrics.vertical_spacing * max(0, metrics.rows - 1)
                + margins.top()
                + margins.bottom()
            )
            container.setMinimumSize(total_width, total_height)

        display_entries, _, _ = self._collect_display_candidates()

        for idx, display_text, score_text in display_entries:
            row = idx // metrics.columns
            column = idx % metrics.columns
            card = self._create_card(
                display_text,
                score_text,
                metrics.card_width,
                metrics.card_height,
                metrics.font_size,
                metrics.padding_h,
                metrics.padding_v,
                metrics.inner_spacing,
            )
            layout.addWidget(card, row, column, Qt.AlignmentFlag.AlignCenter)

    def showEvent(self, event) -> None:  # type: ignore[override]
        super().showEvent(event)
        if self._pending_maximize:
            self._pending_maximize = False
            self.setWindowState(self.windowState() | Qt.WindowState.WindowMaximized)


class RollCallTimerWindow(QWidget):
    """集成点名与计时的主功能窗口。"""
    window_closed = pyqtSignal()
    visibility_changed = pyqtSignal(bool)

    STUDENT_FILE = "students.xlsx"
    MIN_FONT_SIZE = 5
    MAX_FONT_SIZE = 220

    def __init__(self, settings_manager: SettingsManager, student_data, parent: Optional[QWidget] = None) -> None:
        super().__init__(parent)
        self.setWindowTitle("点名 / 计时")
        flags = (
            Qt.WindowType.Window
            | Qt.WindowType.WindowTitleHint
            | Qt.WindowType.WindowCloseButtonHint
            | Qt.WindowType.WindowStaysOnTopHint
            | Qt.WindowType.CustomizeWindowHint
        )
        self.setWindowFlags(flags)
        self.setAttribute(Qt.WidgetAttribute.WA_DeleteOnClose)
        self.settings_manager = settings_manager
        self.student_data = student_data
        try:
            self._rng = random.SystemRandom()
        except NotImplementedError:
            self._rng = random.Random()

        s = self.settings_manager.load_settings().get("RollCallTimer", {})
        def _get_int(key: str, default: int) -> int:
            try:
                return int(s.get(key, str(default)))
            except (TypeError, ValueError):
                return default
        apply_geometry_from_text(self, s.get("geometry", "420x240+180+180"))
        self.setMinimumSize(260, 160)
        # 记录初始最小宽高，供后续还原窗口尺寸时使用
        self._base_minimum_width = self.minimumWidth()
        self._base_minimum_height = self.minimumHeight()
        self._ensure_min_width = self._base_minimum_width
        self._ensure_min_height = self._base_minimum_height

        self.mode = s.get("mode", "roll_call") if s.get("mode", "roll_call") in {"roll_call", "timer"} else "roll_call"
        self.timer_modes = ["countdown", "stopwatch", "clock"]
        self.timer_mode_index = self.timer_modes.index(s.get("timer_mode", "countdown")) if s.get("timer_mode", "countdown") in self.timer_modes else 0
        self._active_timer_mode: Optional[str] = None

        self.timer_countdown_minutes = _get_int("timer_countdown_minutes", 5)
        self.timer_countdown_seconds = _get_int("timer_countdown_seconds", 0)
        self.timer_sound_enabled = str_to_bool(s.get("timer_sound_enabled", "True"), True)

        self.show_id = str_to_bool(s.get("show_id", "True"), True)
        self.show_name = str_to_bool(s.get("show_name", "True"), True)
        if not self.show_id and not self.show_name: self.show_id = True

        self.current_group_name = s.get("current_group", "全部")
        self.groups = ["全部"]
        if not self.student_data.empty:
            gs = sorted({str(g).strip().upper() for g in self.student_data["分组"].dropna() if str(g).strip()})
            self.groups.extend(gs)
        if self.current_group_name not in self.groups: self.current_group_name = "全部"

        self.current_student_index: Optional[int] = None
        self._placeholder_on_show = True
        self._group_all_indices: Dict[str, List[int]] = {}
        self._group_remaining_indices: Dict[str, List[int]] = {}
        self._group_last_student: Dict[str, Optional[int]] = {}
        # 记录各分组初始的随机顺序，便于在界面切换时保持未点名名单不被重新洗牌
        self._group_initial_sequences: Dict[str, List[int]] = {}
        # 记录每个分组已点过名的学生索引，便于核对剩余名单
        self._group_drawn_history: Dict[str, set[int]] = {}
        # 统一维护一个全局已点名集合，确保“全部”分组与子分组状态一致
        self._global_drawn_students: set[int] = set()
        self._student_groups: Dict[int, set[str]] = {}
        self._rebuild_group_indices()
        self._restore_group_state(s)
        self.timer_seconds_left = max(0, _get_int("timer_seconds_left", self.timer_countdown_minutes * 60 + self.timer_countdown_seconds))
        self.timer_stopwatch_seconds = max(0, _get_int("timer_stopwatch_seconds", 0))
        self.timer_running = str_to_bool(s.get("timer_running", "False"), False)

        order_value = str(s.get("scoreboard_order", "rank")).strip().lower()
        self.scoreboard_order = order_value if order_value in {"rank", "id"} else "rank"

        self.last_id_font_size = max(self.MIN_FONT_SIZE, _get_int("id_font_size", 48))
        self.last_name_font_size = max(self.MIN_FONT_SIZE, _get_int("name_font_size", 60))
        self.last_timer_font_size = max(self.MIN_FONT_SIZE, _get_int("timer_font_size", 56))

        self.count_timer = QTimer(self); self.count_timer.setInterval(1000); self.count_timer.timeout.connect(self._on_count_timer)
        self.clock_timer = QTimer(self); self.clock_timer.setInterval(1000); self.clock_timer.timeout.connect(self._update_clock)

        self.tts_manager: Optional[TTSManager] = None
        self.speech_enabled = str_to_bool(s.get("speech_enabled", "False"), False)
        self.selected_voice_id = s.get("speech_voice_id", "")
        manager = TTSManager(self.selected_voice_id, parent=self)
        self.tts_manager = manager
        if not manager.available:
            self.speech_enabled = False
        self._speech_issue_reported = False
        self._speech_check_scheduled = False
        self._pending_passive_student: Optional[int] = None
        self._score_persist_failed = False
        self._score_write_lock = threading.Lock()

        # QFontDatabase 在 Qt 6 中以静态方法为主，这里直接调用类方法避免实例化失败
        families_list = []
        get_families = getattr(QFontDatabase, "families", None)
        if callable(get_families):
            try:
                families_list = list(get_families())
            except TypeError:
                # 个别绑定版本可能要求显式写入枚举参数
                try:
                    families_list = list(get_families(QFontDatabase.WritingSystem.Any))  # type: ignore[arg-type]
                except Exception:
                    families_list = []
        families = set(families_list)
        self.name_font_family = "楷体" if "楷体" in families else ("KaiTi" if "KaiTi" in families else "Microsoft YaHei UI")

        # 使用轻量级的延迟写入机制，避免频繁操作磁盘。
        self._save_timer = QTimer(self)
        self._save_timer.setSingleShot(True)
        self._save_timer.setInterval(250)
        self._save_timer.timeout.connect(self.save_settings)

        self._build_ui()
        self._apply_saved_fonts()
        self._update_menu_state()
        self.update_mode_ui(force_timer_reset=self.mode == "timer")
        self.on_group_change(initial=True)
        self.display_current_student()

    def _build_ui(self) -> None:
        self.setStyleSheet("background-color: #f4f5f7;")
        layout = QVBoxLayout(self)
        layout.setContentsMargins(8, 8, 8, 8)
        layout.setSpacing(6)

        toolbar_layout = QVBoxLayout()
        toolbar_layout.setContentsMargins(0, 0, 0, 0)
        toolbar_layout.setSpacing(2)

        top = QHBoxLayout()
        top.setContentsMargins(0, 0, 0, 0)
        top.setSpacing(4)
        self.title_label = QLabel("点名"); f = QFont("Microsoft YaHei UI", 10, QFont.Weight.Bold)
        self.title_label.setFont(f); self.title_label.setStyleSheet("color: #202124;")
        self.title_label.setSizePolicy(QSizePolicy.Policy.Fixed, QSizePolicy.Policy.Fixed)
        top.addWidget(self.title_label, 0, Qt.AlignmentFlag.AlignLeft)

        self.mode_button = QPushButton("切换到计时"); self.mode_button.setCursor(Qt.CursorShape.PointingHandCursor)
        self.mode_button.clicked.connect(self.toggle_mode); self.mode_button.setFixedHeight(28)
        mode_font = QFont("Microsoft YaHei UI", 9, QFont.Weight.Medium)
        self.mode_button.setFont(mode_font)
        fm = self.mode_button.fontMetrics()
        max_text = max(("切换到计时", "切换到点名"), key=lambda t: fm.horizontalAdvance(t))
        target_width = fm.horizontalAdvance(max_text) + 24
        self.mode_button.setFixedWidth(target_width)  # 固定宽度，避免文本切换导致按钮位置跳动
        self.mode_button.setSizePolicy(QSizePolicy.Policy.Fixed, QSizePolicy.Policy.Fixed)
        top.addWidget(self.mode_button, 0, Qt.AlignmentFlag.AlignLeft)

        compact_font = QFont("Microsoft YaHei UI", 9, QFont.Weight.Medium)
        button_style = (
            "QPushButton {"
            "    padding: 2px 10px;"
            "    border-radius: 11px;"
            "    border: 1px solid #c3c7cf;"
            "    background-color: #ffffff;"
            "    color: #1a1c1f;"
            "}"
            "QPushButton:hover {"
            "    border-color: #1a73e8;"
            "    background-color: #eaf2ff;"
            "}"
            "QPushButton:pressed {"
            "    background-color: #d7e7ff;"
            "}"
        )

        def _setup_secondary_button(button: QPushButton) -> None:
            button.setCursor(Qt.CursorShape.PointingHandCursor)
            button.setFixedHeight(28)
            button.setSizePolicy(QSizePolicy.Policy.Fixed, QSizePolicy.Policy.Fixed)
            button.setFont(compact_font)
            button.setStyleSheet(button_style)

        control_bar = QWidget()
        control_bar.setSizePolicy(QSizePolicy.Policy.Fixed, QSizePolicy.Policy.Fixed)
        control_layout = QHBoxLayout(control_bar)
        control_layout.setContentsMargins(0, 0, 0, 0)
        control_layout.setSpacing(0)

        self.list_button = QPushButton("名单"); _setup_secondary_button(self.list_button)
        self.list_button.clicked.connect(self.show_student_selector)
        control_layout.addWidget(self.list_button)

        self.add_score_button = QPushButton("加分"); _setup_secondary_button(self.add_score_button)
        self.add_score_button.setEnabled(False)
        self.add_score_button.clicked.connect(self.increment_current_score)
        control_layout.addWidget(self.add_score_button)

        self.showcase_button = QPushButton("展示"); _setup_secondary_button(self.showcase_button)
        self.showcase_button.clicked.connect(self.show_scoreboard)
        control_layout.addWidget(self.showcase_button)

        self.reset_button = QPushButton("重置"); _setup_secondary_button(self.reset_button)
        self.reset_button.clicked.connect(self.reset_roll_call_pools)
        control_layout.addWidget(self.reset_button)

        top.addWidget(control_bar, 0, Qt.AlignmentFlag.AlignLeft)
        top.addStretch(1)

        self.menu_button = QToolButton(); self.menu_button.setText("..."); self.menu_button.setPopupMode(QToolButton.ToolButtonPopupMode.InstantPopup)
        self.menu_button.setFixedSize(28, 28); self.menu_button.setStyleSheet("font-size: 18px; padding-bottom: 6px;")
        self.main_menu = self._build_menu(); self.menu_button.setMenu(self.main_menu)
        top.addWidget(self.menu_button, 0, Qt.AlignmentFlag.AlignRight)
        toolbar_layout.addLayout(top)

        group_row = QHBoxLayout()
        group_row.setContentsMargins(0, 0, 0, 0)
        group_row.setSpacing(4)

        self.group_label = QLabel("分组")
        self.group_label.setFont(QFont("Microsoft YaHei UI", 9, QFont.Weight.Medium))
        self.group_label.setStyleSheet("color: #3c4043;")
        self.group_label.setFixedHeight(28)
        self.group_label.setAlignment(Qt.AlignmentFlag.AlignVCenter | Qt.AlignmentFlag.AlignLeft)
        self.group_label.setSizePolicy(QSizePolicy.Policy.Fixed, QSizePolicy.Policy.Fixed)
        group_row.addWidget(self.group_label, 0, Qt.AlignmentFlag.AlignLeft)

        self.group_bar = QWidget()
        self.group_bar.setFixedHeight(28)
        self.group_bar.setSizePolicy(QSizePolicy.Policy.Preferred, QSizePolicy.Policy.Fixed)
        self.group_bar_layout = QHBoxLayout(self.group_bar)
        self.group_bar_layout.setContentsMargins(0, 0, 0, 0)
        self.group_bar_layout.setSpacing(2)
        self.group_button_group = QButtonGroup(self)
        self.group_button_group.setExclusive(True)
        self.group_buttons: Dict[str, QPushButton] = {}
        self._rebuild_group_buttons_ui()
        group_row.addWidget(self.group_bar, 1, Qt.AlignmentFlag.AlignLeft)
        group_row.addStretch(1)
        toolbar_layout.addLayout(group_row)
        layout.addLayout(toolbar_layout)

        self.stack = QStackedWidget(); layout.addWidget(self.stack, 1)

        self.roll_call_frame = ClickableFrame(); self.roll_call_frame.setFrameShape(QFrame.Shape.NoFrame)
        rl = QGridLayout(self.roll_call_frame); rl.setContentsMargins(6, 6, 6, 6); rl.setSpacing(6)
        self.id_label = QLabel(""); self.name_label = QLabel("")
        for lab in (self.id_label, self.name_label):
            lab.setAlignment(Qt.AlignmentFlag.AlignCenter)
            lab.setStyleSheet("color: #ffffff; background-color: #1a73e8; border-radius: 8px; padding: 8px;")
            lab.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        self.score_label = QLabel("成绩：--")
        self.score_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.score_label.setFont(QFont("Microsoft YaHei UI", 24, QFont.Weight.DemiBold))
        self.score_label.setStyleSheet(
            "color: #0b57d0;"
            " background-color: #e8f0fe;"
            " border-radius: 12px;"
            " padding: 2px 16px;"
            " margin: 0px;"
        )

        rl.addWidget(self.id_label, 0, 0); rl.addWidget(self.name_label, 0, 1)
        rl.addWidget(self.score_label, 1, 0, 1, 2)
        self.stack.addWidget(self.roll_call_frame)

        self.timer_frame = QWidget(); tl = QVBoxLayout(self.timer_frame); tl.setContentsMargins(6, 6, 6, 6); tl.setSpacing(8)
        self.time_display_label = QLabel("00:00"); self.time_display_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.time_display_label.setStyleSheet("color: #ffffff; background-color: #202124; border-radius: 8px; padding: 8px;")
        self.time_display_label.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        tl.addWidget(self.time_display_label, 1)

        ctrl = QHBoxLayout(); ctrl.setSpacing(6)
        self.timer_mode_button = QPushButton("倒计时"); self.timer_mode_button.clicked.connect(self.toggle_timer_mode)
        self.timer_start_pause_button = QPushButton("开始"); self.timer_start_pause_button.clicked.connect(self.start_pause_timer)
        self.timer_reset_button = QPushButton("重置"); self.timer_reset_button.clicked.connect(self.reset_timer)
        self.timer_set_button = QPushButton("设定"); self.timer_set_button.clicked.connect(self.set_countdown_time)
        for b in (self.timer_mode_button, self.timer_start_pause_button, self.timer_reset_button, self.timer_set_button):
            b.setCursor(Qt.CursorShape.PointingHandCursor); b.setFixedHeight(30); b.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed); ctrl.addWidget(b)
        tl.addLayout(ctrl); self.stack.addWidget(self.timer_frame)

        self.roll_call_frame.clicked.connect(self.roll_student)
        self.id_label.installEventFilter(self); self.name_label.installEventFilter(self)

    def _apply_saved_fonts(self) -> None:
        id_font = QFont("Microsoft YaHei UI", self.last_id_font_size, QFont.Weight.Bold)
        name_weight = QFont.Weight.Normal if self.name_font_family in {"楷体", "KaiTi"} else QFont.Weight.Bold
        name_font = QFont(self.name_font_family, self.last_name_font_size, name_weight)
        timer_font = QFont("Consolas", self.last_timer_font_size, QFont.Weight.Bold)
        self.id_label.setFont(id_font)
        self.name_label.setFont(name_font)
        self.time_display_label.setFont(timer_font)

    def _build_menu(self) -> QMenu:
        menu = QMenu(self)
        disp = menu.addMenu("显示选项")
        self.show_id_action = disp.addAction("显示学号"); self.show_id_action.setCheckable(True); self.show_id_action.setChecked(self.show_id)
        self.show_id_action.toggled.connect(self._on_display_option_changed)
        self.show_name_action = disp.addAction("显示姓名"); self.show_name_action.setCheckable(True); self.show_name_action.setChecked(self.show_name)
        self.show_name_action.toggled.connect(self._on_display_option_changed)

        menu.addSeparator()
        speech = menu.addMenu("语音播报")
        manager = self.tts_manager
        checked = bool(self.speech_enabled and manager and manager.available)
        self.speech_enabled = checked
        self.speech_enabled_action = speech.addAction("启用语音播报"); self.speech_enabled_action.setCheckable(True)
        self.speech_enabled_action.setChecked(checked); self.speech_enabled_action.toggled.connect(self._toggle_speech)
        self.voice_menu = speech.addMenu("选择发音人"); self.voice_actions = []
        if manager and manager.available:
            self.speech_enabled_action.setEnabled(True)
            self.speech_enabled_action.setToolTip("点名时自动朗读当前学生姓名。")
            if manager.supports_voice_selection and manager.voice_ids:
                for vid in manager.voice_ids:
                    act = self.voice_menu.addAction(vid); act.setCheckable(True); act.setChecked(vid == manager.current_voice_id)
                    act.triggered.connect(lambda _c, v=vid: self._set_voice(v)); self.voice_actions.append(act)
            else:
                self.voice_menu.setEnabled(False)
                self.voice_menu.setToolTip("当前语音引擎不支持切换发音人。")
        else:
            self.voice_menu.setEnabled(False)
            self.speech_enabled_action.setEnabled(False)
            reason, suggestions = self._collect_speech_issue_details()
            tooltip_lines = [reason] if reason else []
            tooltip_lines.extend(suggestions)
            if tooltip_lines:
                self.speech_enabled_action.setToolTip("\n".join(tooltip_lines))

        menu.addSeparator()
        self.timer_sound_action = menu.addAction("倒计时结束提示音"); self.timer_sound_action.setCheckable(True)
        self.timer_sound_action.setChecked(self.timer_sound_enabled); self.timer_sound_action.toggled.connect(self._toggle_timer_sound)
        return menu

    def _update_menu_state(self) -> None:
        if self.show_id_action.isChecked() != self.show_id: self.show_id_action.setChecked(self.show_id)
        if self.show_name_action.isChecked() != self.show_name: self.show_name_action.setChecked(self.show_name)
        self.timer_sound_action.setChecked(self.timer_sound_enabled)
        manager = self.tts_manager
        if manager and manager.available:
            self.speech_enabled_action.setEnabled(True)
            self.speech_enabled_action.setChecked(self.speech_enabled)
            self.speech_enabled_action.setToolTip("点名时自动朗读当前学生姓名。")
            if manager.supports_voice_selection and manager.voice_ids:
                self.voice_menu.setEnabled(True)
                for act in self.voice_actions:
                    act.setChecked(act.text() == manager.current_voice_id)
                self.voice_menu.setToolTip("")
            else:
                self.voice_menu.setEnabled(False)
                self.voice_menu.setToolTip("当前语音引擎不支持切换发音人。")
        else:
            self.voice_menu.setEnabled(False)
            self.speech_enabled_action.setEnabled(False)
            self.speech_enabled_action.setChecked(False)
            reason, suggestions = self._collect_speech_issue_details()
            tooltip_lines = [reason] if reason else []
            tooltip_lines.extend(suggestions)
            if tooltip_lines:
                self.speech_enabled_action.setToolTip("\n".join(tooltip_lines))
            if not self._speech_issue_reported:
                self._diagnose_speech_engine()

    def _default_speech_suggestions(self) -> List[str]:
        hints: List[str] = []
        if sys.platform == "win32":
            hints.append("请确认 Windows 已启用 SAPI5 中文语音包。")
        elif sys.platform == "darwin":
            hints.append("请在系统“辅助功能 -> 语音”中启用所需的语音包。")
        else:
            hints.append("请确保系统已安装可用的语音引擎（如 espeak）并重新启动程序。")
        hints.append("可尝试重新安装 pyttsx3 或检查语音服务状态后重启软件。")
        return hints

    def _dedupe_suggestions(self, values: List[str]) -> List[str]:
        return dedupe_strings(values)

    def _collect_speech_issue_details(self) -> tuple[str, List[str]]:
        manager = self.tts_manager
        reason = ""
        suggestions: List[str] = []
        if manager is None:
            reason = "无法初始化系统语音引擎"
            suggestions = self._default_speech_suggestions()
        elif not manager.available:
            reason, suggestions = manager.diagnostics()
            reason = reason or "无法初始化系统语音引擎"
            if not suggestions:
                suggestions = self._default_speech_suggestions()
        elif manager.supports_voice_selection and not getattr(manager, "voice_ids", []):
            reason = "未检测到任何可用的发音人"
            suggestions = self._default_speech_suggestions()
            suggestions.append("请在操作系统语音设置中添加语音包后重新启动程序。")

        env_reason, env_suggestions = detect_speech_environment_issues()
        if env_reason:
            if not reason:
                reason = env_reason
            elif env_reason not in reason:
                reason = f"{reason}；{env_reason}"
        suggestions.extend(env_suggestions)
        if not suggestions and reason:
            suggestions = self._default_speech_suggestions()
        return reason, self._dedupe_suggestions(suggestions)

    def _ensure_speech_manager(self) -> Optional[TTSManager]:
        manager = self.tts_manager
        if manager and manager.available:
            return manager
        if manager is not None:
            try:
                manager.shutdown()
            except Exception:
                pass
        manager = TTSManager(self.selected_voice_id, parent=self)
        self.tts_manager = manager
        if manager.available:
            self._speech_issue_reported = False
            QTimer.singleShot(0, self._update_menu_state)
        return manager

    def _diagnose_speech_engine(self) -> None:
        if self._speech_issue_reported:
            return
        action = getattr(self, "speech_enabled_action", None)
        if action is None or action.isEnabled():
            return
        if not self.isVisible():
            if not self._speech_check_scheduled:
                self._speech_check_scheduled = True
                QTimer.singleShot(200, self._diagnose_speech_engine)
            return
        reason, suggestions = self._collect_speech_issue_details()
        if not reason:
            return
        advice = "\n".join(f"· {line}" for line in suggestions)
        message = f"语音播报功能当前不可用：{reason}"
        if advice:
            message = f"{message}\n{advice}"
        show_quiet_information(self, message, "语音播报提示")
        self._speech_issue_reported = True
        self._speech_check_scheduled = False

    def eventFilter(self, obj, e):
        if obj in (self.id_label, self.name_label) and e.type() == QEvent.Type.MouseButtonPress:
            if e.button() == Qt.MouseButton.LeftButton:
                self.roll_student(); return True
        return super().eventFilter(obj, e)

    def _on_display_option_changed(self) -> None:
        if not self.show_id_action.isChecked() and not self.show_name_action.isChecked():
            self.show_id_action.setChecked(True)
        self.show_id = self.show_id_action.isChecked()
        self.show_name = self.show_name_action.isChecked()
        self.update_display_layout()
        self.display_current_student()
        self._schedule_save()

    def _toggle_speech(self, enabled: bool) -> None:
        if not enabled:
            self.speech_enabled = False
            self._schedule_save()
            return
        manager = self._ensure_speech_manager()
        if not manager or not manager.available:
            reason, suggestions = self._collect_speech_issue_details()
            message = reason or "未检测到语音引擎，无法开启语音播报。"
            advice = "\n".join(f"· {line}" for line in suggestions)
            if advice:
                message = f"{message}\n{advice}"
            show_quiet_information(self, message, "语音播报提示")
            self.speech_enabled_action.setChecked(False)
            self._speech_issue_reported = True
            return
        if manager.supports_voice_selection and not getattr(manager, "voice_ids", []):
            reason, suggestions = self._collect_speech_issue_details()
            message = reason or "未检测到可用的发音人。"
            advice = "\n".join(f"· {line}" for line in suggestions)
            if advice:
                message = f"{message}\n{advice}"
            show_quiet_information(self, message, "语音播报提示")
            self.speech_enabled_action.setChecked(False)
            self._speech_issue_reported = True
            return
        self.speech_enabled = enabled
        self._schedule_save()

    def _set_voice(self, voice_id: str) -> None:
        manager = self._ensure_speech_manager()
        if not manager or not manager.supports_voice_selection:
            return
        manager.set_voice(voice_id); self.selected_voice_id = voice_id
        for a in self.voice_actions: a.setChecked(a.text() == voice_id)
        self._schedule_save()

    def _toggle_timer_sound(self, enabled: bool) -> None:
        self.timer_sound_enabled = enabled
        self._schedule_save()

    def _set_scoreboard_order(self, order: str) -> None:
        normalized = str(order).strip().lower()
        if normalized not in {"rank", "id"}:
            return
        if self.scoreboard_order == normalized:
            return
        self.scoreboard_order = normalized
        self._schedule_save()

    def _speak_text(self, text: str) -> None:
        if not text:
            return
        manager = self.tts_manager
        if not (self.speech_enabled and manager and manager.available):
            return
        manager.speak(text)

    def _announce_current_student(self) -> None:
        if (
            not self.speech_enabled
            or self.tts_manager is None
            or not self.tts_manager.available
            or self.current_student_index is None
            or self.student_data is None
            or self.student_data.empty
        ):
            return
        try:
            stu = self.student_data.loc[self.current_student_index]
        except Exception:
            return
        name_value = stu.get("姓名", "")
        if isinstance(name_value, str):
            name = name_value.strip()
        else:
            name = str(name_value).strip() if pd.notna(name_value) else ""
        if name:
            self._speak_text(name)

    def show_student_selector(self) -> None:
        if self.mode != "roll_call":
            return
        if self.student_data is None or self.student_data.empty:
            show_quiet_information(self, "暂无学生数据，无法显示名单。")
            return
        records: List[tuple[int, str, str, int]] = []
        for idx, row in self.student_data.iterrows():
            sid_value = row.get("学号", "")
            sid_display = re.sub(r"\s+", "", _normalize_text(sid_value))
            name = re.sub(r"\s+", "", _normalize_text(row.get("姓名", "")))
            try:
                sort_key = int(sid_display) if sid_display else sys.maxsize
            except (TypeError, ValueError):
                sort_key = sys.maxsize
            records.append((sort_key, sid_display, name, idx))
        if not records:
            show_quiet_information(self, "当前没有可显示的学生名单。")
            return
        records.sort(key=lambda item: (item[0], item[1]))
        dialog_data = []
        for _, sid, name, data_idx in records:
            display_sid = sid if sid else "无学号"
            display_name = name or "未命名"
            dialog_data.append((display_sid, display_name, data_idx))
        dialog = StudentListDialog(self, dialog_data)
        if dialog.exec() == QDialog.DialogCode.Accepted and dialog.selected_index is not None:
            selected = dialog.selected_index
            if selected in self.student_data.index:
                self.current_student_index = selected
                self._pending_passive_student = None
                self.display_current_student()
                self._announce_current_student()

    def increment_current_score(self) -> None:
        if self.mode != "roll_call":
            return
        if self.current_student_index is None:
            show_quiet_information(self, "请先选择需要加分的学生。")
            return
        if self.student_data is None:
            return
        if "成绩" not in self.student_data.columns:
            self.student_data["成绩"] = 0
        value = self.student_data.at[self.current_student_index, "成绩"]
        try:
            base = int(value)
        except (TypeError, ValueError):
            base = 0
        new_score = base + 1
        self.student_data.at[self.current_student_index, "成绩"] = new_score
        self._pending_passive_student = None
        self._update_score_display()
        self._persist_student_scores()
        self._speak_text("加分")

    def show_scoreboard(self) -> None:
        if self.mode != "roll_call":
            return
        if self.student_data is None or self.student_data.empty:
            show_quiet_information(self, "暂无学生数据，无法展示成绩。")
            return
        if "成绩" not in self.student_data.columns:
            self.student_data["成绩"] = 0
        records: List[tuple[str, str, int]] = []
        for _, row in self.student_data.iterrows():
            sid_value = row.get("学号", "")
            sid_display = re.sub(r"\s+", "", _normalize_text(sid_value))
            name = re.sub(r"\s+", "", _normalize_text(row.get("姓名", ""))) or "未命名"
            value = row.get("成绩", 0)
            try:
                score = int(value)
            except (TypeError, ValueError):
                score = 0
            records.append((sid_display, name, score))
        dialog = ScoreboardDialog(
            self,
            records,
            order=self.scoreboard_order,
            order_changed=self._set_scoreboard_order,
        )
        dialog.exec()

    def _persist_student_scores(self) -> None:
        if not (PANDAS_AVAILABLE and OPENPYXL_AVAILABLE):
            return
        if self.student_data is None:
            return
        try:
            with self._score_write_lock:
                df = self.student_data.copy()
                df = _normalize_student_dataframe(df, drop_incomplete=False)
                _write_student_workbook(self.STUDENT_FILE, df)
            self._score_persist_failed = False
        except Exception as exc:
            if not self._score_persist_failed:
                show_quiet_information(self, f"保存成绩失败：{exc}")
                self._score_persist_failed = True

    def toggle_mode(self) -> None:
        self.mode = "timer" if self.mode == "roll_call" else "roll_call"
        if self.mode == "roll_call":
            self._placeholder_on_show = True
        self.update_mode_ui(force_timer_reset=self.mode == "timer")
        self._schedule_save()

    def update_mode_ui(self, force_timer_reset: bool = False) -> None:
        is_roll = self.mode == "roll_call"
        self.title_label.setText("点名" if is_roll else "计时")
        self.mode_button.setText("切换到计时" if is_roll else "切换到点名")
        self.group_label.setVisible(is_roll)
        if hasattr(self, "group_bar"):
            self.group_bar.setVisible(is_roll)
        self._update_roll_call_controls()
        if is_roll:
            if self._placeholder_on_show:
                self.current_student_index = None
            self.stack.setCurrentWidget(self.roll_call_frame)
            self.count_timer.stop(); self.clock_timer.stop(); self.timer_running = False; self.timer_start_pause_button.setText("开始")
            self.update_display_layout(); self.display_current_student()
            self.schedule_font_update()
            self._placeholder_on_show = False
        else:
            self.stack.setCurrentWidget(self.timer_frame)
            changed = False
            if force_timer_reset:
                changed = self.reset_timer(persist=False)
            self.update_timer_mode_ui()
            if changed:
                self._schedule_save()
            self.schedule_font_update()
        self.reset_button.setVisible(is_roll)
        self.updateGeometry()

    def _handle_timer_mode_transition(self, previous_mode: Optional[str], new_mode: str) -> None:
        if previous_mode == new_mode:
            return
        if new_mode in {"countdown", "stopwatch"}:
            self.timer_running = False
            self.count_timer.stop()
            self.timer_start_pause_button.setText("开始")
            if new_mode == "countdown":
                total = max(0, self.timer_countdown_minutes * 60 + self.timer_countdown_seconds)
                self.timer_seconds_left = total
            else:
                self.timer_stopwatch_seconds = 0

    def update_timer_mode_ui(self) -> None:
        mode = self.timer_modes[self.timer_mode_index]
        previous_mode = self._active_timer_mode
        if previous_mode is not None:
            self._handle_timer_mode_transition(previous_mode, mode)
        self._active_timer_mode = mode
        self.clock_timer.stop()
        if mode == "countdown":
            self.timer_mode_button.setText("倒计时")
            self.timer_start_pause_button.setEnabled(True); self.timer_reset_button.setEnabled(True); self.timer_set_button.setEnabled(True)
            self.timer_start_pause_button.setText("暂停" if self.timer_running else "开始")
            if self.timer_running and not self.count_timer.isActive(): self.count_timer.start()
            self.update_timer_display()
        elif mode == "stopwatch":
            self.timer_mode_button.setText("秒表")
            self.timer_start_pause_button.setEnabled(True); self.timer_reset_button.setEnabled(True); self.timer_set_button.setEnabled(False)
            self.timer_start_pause_button.setText("暂停" if self.timer_running else "开始")
            if self.timer_running and not self.count_timer.isActive(): self.count_timer.start()
            self.update_timer_display()
        else:
            self.timer_mode_button.setText("时钟")
            self.timer_start_pause_button.setEnabled(False); self.timer_reset_button.setEnabled(False); self.timer_set_button.setEnabled(False)
            self.timer_running = False; self.count_timer.stop(); self._update_clock(); self.clock_timer.start()
            self.timer_start_pause_button.setText("开始")
        self.schedule_font_update()

    def toggle_timer_mode(self) -> None:
        if self.timer_running: return
        self.timer_mode_index = (self.timer_mode_index + 1) % len(self.timer_modes); self.update_timer_mode_ui()
        self._schedule_save()

    def start_pause_timer(self) -> None:
        if self.timer_modes[self.timer_mode_index] == "clock": return
        self.timer_running = not self.timer_running
        if self.timer_running:
            self.timer_start_pause_button.setText("暂停")
            if not self.count_timer.isActive(): self.count_timer.start()
        else:
            self.timer_start_pause_button.setText("开始")
            self.count_timer.stop()
        self._schedule_save()

    def reset_timer(self, persist: bool = True) -> bool:
        changed = self.timer_running
        self.timer_running = False; self.count_timer.stop(); self.timer_start_pause_button.setText("开始")
        m = self.timer_modes[self.timer_mode_index]
        if m == "countdown":
            baseline = max(0, self.timer_countdown_minutes * 60 + self.timer_countdown_seconds)
            if self.timer_seconds_left != baseline:
                self.timer_seconds_left = baseline
                changed = True
        elif m == "stopwatch":
            if self.timer_stopwatch_seconds != 0:
                self.timer_stopwatch_seconds = 0
                changed = True
        self.update_timer_display()
        if persist and changed:
            self._schedule_save()
        return changed

    def set_countdown_time(self) -> None:
        d = CountdownSettingsDialog(self, self.timer_countdown_minutes, self.timer_countdown_seconds)
        if d.exec() and d.result:
            mi, se = d.result; self.timer_countdown_minutes = mi; self.timer_countdown_seconds = se
            changed = self.reset_timer()
            if not changed:
                self._schedule_save()

    def _on_count_timer(self) -> None:
        m = self.timer_modes[self.timer_mode_index]
        if m == "countdown":
            if self.timer_seconds_left > 0: self.timer_seconds_left -= 1
            else:
                self.count_timer.stop(); self.timer_running = False; self.timer_start_pause_button.setText("开始"); self.update_timer_display()
                if self.timer_sound_enabled: self.play_timer_sound()
                return
        elif m == "stopwatch":
            self.timer_stopwatch_seconds += 1
        self.update_timer_display()

    def update_timer_display(self) -> None:
        m = self.timer_modes[self.timer_mode_index]
        if m == "countdown": seconds = max(0, self.timer_seconds_left)
        elif m == "stopwatch": seconds = max(0, self.timer_stopwatch_seconds)
        else: seconds = 0
        if m in {"countdown", "stopwatch"}:
            mi, se = divmod(seconds, 60); self.time_display_label.setText(f"{int(mi):02d}:{int(se):02d}")
        else:
            self.time_display_label.setText(time.strftime("%H:%M:%S"))
        self.schedule_font_update()

    def _update_clock(self) -> None:
        self.time_display_label.setText(time.strftime("%H:%M:%S"))
        self.schedule_font_update()

    def play_timer_sound(self) -> None:
        if SOUNDDEVICE_AVAILABLE:
            def _play() -> None:
                try:
                    fs = 44100; duration = 0.5; frequency = 880
                    t = np.linspace(0, duration, int(fs * duration), endpoint=False)
                    data = 0.4 * np.sin(2 * np.pi * frequency * t)
                    sd.play(data.astype(np.float32), fs); sd.wait()
                except Exception:
                    pass
            threading.Thread(target=_play, daemon=True).start()

    def on_group_change(self, group_name: Optional[str] = None, initial: bool = False) -> None:
        if not self.groups:
            return
        if group_name is None:
            group_name = self.current_group_name
        if group_name not in self.groups:
            group_name = "全部" if "全部" in self.groups else self.groups[0]
        previous_group = self.current_group_name
        self.current_group_name = group_name
        self._update_group_button_state(group_name)
        if self.student_data.empty:
            self.current_student_index = None
            self.display_current_student()
            if not initial and previous_group != group_name:
                self._schedule_save()
            return
        self._pending_passive_student = None
        self._ensure_group_pool(group_name)
        self.current_student_index = None
        self.display_current_student()
        if not initial and previous_group != group_name:
            self._schedule_save()

    def roll_student(self, speak: bool = True) -> None:
        if self.mode != "roll_call": return
        group_name = self.current_group_name
        pool = self._group_remaining_indices.get(group_name)
        if pool is None:
            self._ensure_group_pool(group_name)
            pool = self._group_remaining_indices.get(group_name, [])
        if not pool:
            base_total = self._group_all_indices.get(group_name, [])
            if not base_total:
                show_quiet_information(self, f"'{group_name}' 分组当前没有可点名的学生。")
                self.current_student_index = None
                self.display_current_student()
                return
            if self._all_groups_completed():
                show_quiet_information(self, "所有学生都已完成点名，请点击“重置”按钮重新开始。")
            else:
                show_quiet_information(self, f"'{group_name}' 的同学已经全部点到，请切换其他分组或点击“重置”按钮。")
            return
        self.current_student_index = pool.pop()
        self._pending_passive_student = self.current_student_index
        self._group_last_student[group_name] = self.current_student_index
        self._mark_student_drawn(self.current_student_index)
        self.display_current_student()
        if speak:
            self._announce_current_student()
        # 即时同步保存配置，防止异常退出导致未点名名单丢失。
        self.save_settings()

    def _all_groups_completed(self) -> bool:
        """判断是否所有分组的学生都已点名完毕。"""

        total_students = len(self._group_all_indices.get("全部", []))
        if total_students == 0:
            return True
        if len(self._global_drawn_students) < total_students:
            return False
        for group, base in self._group_all_indices.items():
            if not base:
                continue
            remaining = self._group_remaining_indices.get(group, [])
            if remaining:
                return False
        return True

    def _reset_roll_call_state(self) -> None:
        """清空全部点名历史并重新洗牌。"""

        self.settings_manager.clear_roll_call_history()
        self._pending_passive_student = None
        self._rebuild_group_indices()
        self._ensure_group_pool(self.current_group_name)

    def _shuffle(self, values: List[int]) -> None:
        try:
            self._rng.shuffle(values)
        except Exception:
            random.shuffle(values)

    def reset_roll_call_pools(self) -> None:
        """根据当前分组执行重置：子分组独立重置，“全部”重置所有。"""

        group_name = self.current_group_name
        if group_name == "全部":
            self._reset_roll_call_state()
        else:
            self._reset_single_group(group_name)
        self.current_student_index = None
        self._pending_passive_student = None
        self.display_current_student()
        self.save_settings()

    def _reset_single_group(self, group_name: str) -> None:
        """仅重置指定分组，同时保持其它分组及全局状态不变。"""

        if group_name == "全部":
            return
        base_indices_raw = self._group_all_indices.get(group_name)
        if base_indices_raw is None:
            return
        base_indices: List[int] = []
        for value in base_indices_raw:
            try:
                base_indices.append(int(value))
            except (TypeError, ValueError):
                continue
        shuffled = list(base_indices)
        self._shuffle(shuffled)
        self._group_remaining_indices[group_name] = shuffled
        self._group_initial_sequences[group_name] = list(shuffled)
        self._group_last_student[group_name] = None
        if self._pending_passive_student in shuffled:
            self._pending_passive_student = None

        history = self._group_drawn_history.setdefault(group_name, set())
        if history:
            for idx in list(history):
                self._remove_from_global_history(idx, ignore_group=group_name)
            history.clear()

        last_all = self._group_last_student.get("全部")
        if last_all is not None:
            try:
                last_all_key = int(last_all)
            except (TypeError, ValueError):
                last_all_key = None
            if last_all_key is not None and last_all_key not in self._global_drawn_students:
                self._group_last_student["全部"] = None

        self._refresh_all_group_pool()

    def _rebuild_group_indices(self) -> None:
        """重新构建各分组的学生索引池。"""

        all_indices: Dict[str, List[int]] = {}
        remaining: Dict[str, List[int]] = {}
        last_student: Dict[str, Optional[int]] = {}
        student_groups: Dict[int, set[str]] = {}
        initial_sequences: Dict[str, List[int]] = {}

        if self.student_data.empty:
            all_indices["全部"] = []
        else:
            all_indices["全部"] = list(self.student_data.index)
            for idx in all_indices["全部"]:
                student_groups.setdefault(int(idx), set()).add("全部")
            group_series = self.student_data["分组"].astype(str).str.strip().str.upper()
            for group_name in self.groups:
                if group_name == "全部":
                    continue
                mask = group_series == group_name
                all_indices[group_name] = list(self.student_data[mask].index)
                for idx in all_indices[group_name]:
                    student_groups.setdefault(int(idx), set()).add(group_name)

        for group_name, indices in all_indices.items():
            pool = list(indices)
            self._shuffle(pool)
            remaining[group_name] = pool
            initial_sequences[group_name] = list(pool)
            last_student[group_name] = None

        self._group_all_indices = all_indices
        self._group_remaining_indices = remaining
        self._group_last_student = last_student
        self._group_initial_sequences = initial_sequences
        self._student_groups = student_groups
        self._group_drawn_history = {group: set() for group in all_indices}
        # “全部”分组直接引用全局集合，避免重复维护两份数据造成不一致
        if "全部" in self._group_drawn_history:
            self._global_drawn_students.clear()
            self._group_drawn_history["全部"] = self._global_drawn_students
        else:
            self._group_drawn_history["全部"] = self._global_drawn_students

        self._refresh_all_group_pool()

    def _remove_from_global_history(self, student_index: int, ignore_group: Optional[str] = None) -> None:
        """若学生未在其它分组被点名，则从全局记录中移除。"""

        try:
            student_key = int(student_index)
        except (TypeError, ValueError):
            return
        for group, history in self._group_drawn_history.items():
            if group == "全部" or group == ignore_group:
                continue
            if student_key in history:
                return
        self._global_drawn_students.discard(student_key)

    def _restore_group_state(self, section: Mapping[str, str]) -> None:
        """从配置中恢复各分组剩余学生池，保持未抽学生不重复。"""

        def _load_dict(key: str) -> Dict[str, object]:
            raw = section.get(key, "")
            if not raw:
                return {}
            try:
                data = json.loads(raw)
            except Exception:
                return {}
            return data if isinstance(data, dict) else {}

        remaining_data = _load_dict("group_remaining")
        last_data = _load_dict("group_last")

        # 读取保存的全局已点名名单，保证窗口被关闭后重新打开时仍能继承上一轮的状态
        global_drawn_raw = section.get("global_drawn", "")
        restored_global: set[int] = set()
        if global_drawn_raw:
            try:
                payload = json.loads(global_drawn_raw)
            except Exception:
                payload = []
            if isinstance(payload, list):
                for value in payload:
                    try:
                        restored_global.add(int(value))
                    except (TypeError, ValueError):
                        continue

        self._global_drawn_students.clear()
        self._global_drawn_students.update(restored_global)
        self._group_drawn_history["全部"] = self._global_drawn_students

        # 先记录一份备份，稍后重新计算所有集合时需要与持久化信息交叉验证
        existing_global = set(self._global_drawn_students)

        # 从头构建全局集合，避免旧对象上的引用导致状态被意外清空
        self._global_drawn_students = set()
        self._group_drawn_history["全部"] = self._global_drawn_students

        for group, indices in remaining_data.items():
            if group not in self._group_all_indices:
                continue
            base_list = self._group_all_indices[group]
            base_set = set(base_list)
            restored: List[int] = []
            if isinstance(indices, list):
                seen: set[int] = set()
                for value in indices:
                    try:
                        idx = int(value)
                    except (TypeError, ValueError):
                        continue
                    if idx not in base_set or idx in seen:
                        continue
                    restored.append(idx)
                    seen.add(idx)
            self._group_remaining_indices[group] = restored

        for group, value in last_data.items():
            if group not in self._group_all_indices:
                continue
            if value is None:
                self._group_last_student[group] = None
                continue
            try:
                idx = int(value)
            except (TypeError, ValueError):
                continue
            if idx in self._group_all_indices[group]:
                self._group_last_student[group] = idx

        # 根据恢复后的剩余名单推导出每个分组已点名的学生集合
        # 重新整理所有分组的已点名学生，并同步更新全局集合
        for group, base_indices in self._group_all_indices.items():
            normalized_base: List[int] = []
            for value in base_indices:
                try:
                    normalized_base.append(int(value))
                except (TypeError, ValueError):
                    continue
            remaining_set: set[int] = set()
            for value in self._group_remaining_indices.get(group, []):
                try:
                    remaining_set.add(int(value))
                except (TypeError, ValueError):
                    continue
            drawn = {idx for idx in normalized_base if idx not in remaining_set}
            if group != "全部" and existing_global:
                # 在恢复时合并先前记录的全局名单，防止由于意外写入丢失导致遗漏
                drawn.update(idx for idx in existing_global if idx in normalized_base)
            seq = list(self._group_remaining_indices.get(group, []))
            seq.extend(idx for idx in normalized_base if idx not in seq)
            self._group_initial_sequences[group] = seq
            if group == "全部":
                # “全部”分组的已点名集合以全局集合为准
                self._global_drawn_students.update(drawn)
            else:
                self._group_drawn_history[group] = drawn
                self._global_drawn_students.update(drawn)

        # 合并持久化阶段记录的全局集合，防止遗漏尚未恢复的记录
        if existing_global:
            self._global_drawn_students.update(existing_global)

        # 最后重新指定“全部”分组引用当前全局集合，保持一致性
        self._group_drawn_history["全部"] = self._global_drawn_students
        self._refresh_all_group_pool()

    def _ensure_group_pool(self, group_name: str, force_reset: bool = False) -> None:
        """确保指定分组仍有待抽取的学生，必要时重新洗牌。"""

        if group_name not in self._group_all_indices:
            if self.student_data.empty:
                base_list: List[int] = []
            elif group_name == "全部":
                base_list = list(self.student_data.index)
            else:
                group_series = self.student_data["分组"].astype(str).str.strip().str.upper()
                base_list = list(self.student_data[group_series == group_name].index)
            self._group_all_indices[group_name] = base_list
            self._group_remaining_indices[group_name] = []
            self._group_last_student.setdefault(group_name, None)
            if group_name == "全部":
                self._group_drawn_history[group_name] = self._global_drawn_students
            else:
                self._group_drawn_history.setdefault(group_name, set())
            for idx in base_list:
                entry = self._student_groups.setdefault(int(idx), set())
                entry.add(group_name)
                entry.add("全部")
            # 新增分组时同步生成初始顺序
            shuffled = list(base_list)
            self._shuffle(shuffled)
            self._group_initial_sequences[group_name] = shuffled

        base_indices: List[int] = []
        for value in self._group_all_indices.get(group_name, []):
            try:
                base_indices.append(int(value))
            except (TypeError, ValueError):
                continue
        if group_name == "全部":
            drawn_history = self._group_drawn_history.setdefault("全部", self._global_drawn_students)
            reference_drawn = self._global_drawn_students
        else:
            drawn_history = self._group_drawn_history.setdefault(group_name, set())
            reference_drawn = drawn_history

        if group_name == "全部":
            # “全部”分组直接依据全局集合生成剩余名单，避免与各子分组脱节
            if group_name not in self._group_initial_sequences:
                shuffled = list(base_indices)
                self._shuffle(shuffled)
                self._group_initial_sequences[group_name] = shuffled
            self._refresh_all_group_pool()
            self._group_last_student.setdefault(group_name, None)
            return

        if force_reset or group_name not in self._group_remaining_indices:
            drawn_history.clear()
            pool = list(base_indices)
            self._shuffle(pool)
            self._group_remaining_indices[group_name] = pool
            self._group_last_student.setdefault(group_name, None)
            self._group_initial_sequences[group_name] = list(pool)
            self._refresh_all_group_pool()
            return

        raw_pool = self._group_remaining_indices.get(group_name, [])
        normalized_pool: List[int] = []
        seen: set[int] = set()
        for value in raw_pool:
            try:
                idx = int(value)
            except (TypeError, ValueError):
                continue
            if idx in base_indices and idx not in seen and idx not in reference_drawn:
                normalized_pool.append(idx)
                seen.add(idx)

        source_order = self._group_initial_sequences.get(group_name)
        if source_order is None:
            # 如果未记录初始顺序，则退回数据原有顺序
            source_order = list(base_indices)
            self._group_initial_sequences[group_name] = list(source_order)

        for idx in source_order:
            if idx in reference_drawn or idx in seen or idx not in base_indices:
                continue
            normalized_pool.append(idx)
            seen.add(idx)

        self._group_remaining_indices[group_name] = normalized_pool
        self._group_last_student.setdefault(group_name, None)
        self._refresh_all_group_pool()

    def _mark_student_drawn(self, student_index: int) -> None:
        """抽中学生后，从所有关联分组的候选列表中移除该学生。"""

        student_key = None
        try:
            student_key = int(student_index)
        except (TypeError, ValueError):
            return

        groups = self._student_groups.get(student_key)
        if not groups:
            return
        self._global_drawn_students.add(student_key)
        for group in groups:
            if group == "全部":
                history = self._group_drawn_history.setdefault("全部", self._global_drawn_students)
            else:
                history = self._group_drawn_history.setdefault(group, set())
            history.add(student_key)
            pool = self._group_remaining_indices.get(group)
            if not pool:
                continue
            cleaned: List[int] = []
            for value in pool:
                try:
                    idx = int(value)
                except (TypeError, ValueError):
                    continue
                if idx != student_key:
                    cleaned.append(idx)
            self._group_remaining_indices[group] = cleaned

        self._refresh_all_group_pool()

    def _refresh_all_group_pool(self) -> None:
        """同步“全部”分组的剩余名单，使其与各子分组保持一致。"""

        base_all = self._group_all_indices.get("全部", [])
        if "全部" not in self._group_initial_sequences:
            shuffled = list(base_all)
            self._shuffle(shuffled)
            self._group_initial_sequences["全部"] = shuffled
        order = list(self._group_initial_sequences.get("全部", []))
        normalized: List[int] = []
        seen: set[int] = set()
        for value in order:
            try:
                idx = int(value)
            except (TypeError, ValueError):
                continue
            if idx in seen:
                continue
            seen.add(idx)
            if idx in self._global_drawn_students:
                continue
            normalized.append(idx)
        for value in base_all:
            try:
                idx = int(value)
            except (TypeError, ValueError):
                continue
            if idx in seen:
                continue
            seen.add(idx)
            if idx in self._global_drawn_students:
                continue
            normalized.append(idx)
        self._group_remaining_indices["全部"] = normalized

    def display_current_student(self) -> None:
        if self.current_student_index is None:
            self.id_label.setText("学号" if self.show_id else "")
            self.name_label.setText("学生" if self.show_name else "")
        else:
            stu = self.student_data.loc[self.current_student_index]
            raw_sid = stu.get("学号", "")
            raw_name = stu.get("姓名", "")
            sid = re.sub(r"\s+", "", _normalize_text(raw_sid))
            name = re.sub(r"\s+", "", _normalize_text(raw_name))
            self.id_label.setText(sid if self.show_id else ""); self.name_label.setText(name if self.show_name else "")
            if not self.show_id: self.id_label.setText("")
            if not self.show_name: self.name_label.setText("")
        self.update_display_layout()
        self._update_score_display()
        self._update_roll_call_controls()
        self.schedule_font_update()

    def update_display_layout(self) -> None:
        self.id_label.setVisible(self.show_id); self.name_label.setVisible(self.show_name)
        layout: QGridLayout = self.roll_call_frame.layout()
        layout.setColumnStretch(0, 1); layout.setColumnStretch(1, 1)
        if not self.show_id: layout.setColumnStretch(0, 0)
        if not self.show_name: layout.setColumnStretch(1, 0)
        self.schedule_font_update()

    def _rebuild_group_buttons_ui(self) -> None:
        if not hasattr(self, "group_bar_layout"):
            return
        for button in list(self.group_button_group.buttons()):
            self.group_button_group.removeButton(button)
        while self.group_bar_layout.count():
            item = self.group_bar_layout.takeAt(0)
            widget = item.widget()
            if widget is not None:
                widget.deleteLater()
        self.group_buttons = {}
        if not self.groups:
            return
        button_font = QFont("Microsoft YaHei UI", 9, QFont.Weight.Medium)
        for group in self.groups:
            button = QPushButton(group)
            button.setCheckable(True)
            button.setCursor(Qt.CursorShape.PointingHandCursor)
            button.setFixedHeight(28)
            button.setFont(button_font)
            button.setStyleSheet(
                "QPushButton {"
                "    padding: 4px 14px;"
                "    border-radius: 14px;"
                "    border: 1px solid #d0d7de;"
                "    background-color: #ffffff;"
                "    color: #1b1f24;"
                "}"
                "QPushButton:hover {"
                "    border-color: #1a73e8;"
                "}"
                "QPushButton:checked {"
                "    background-color: #1a73e8;"
                "    color: #ffffff;"
                "    border-color: #1a73e8;"
                "}"
            )
            button.clicked.connect(lambda _checked=False, name=group: self.on_group_change(name))
            self.group_bar_layout.addWidget(button)
            self.group_button_group.addButton(button)
            self.group_buttons[group] = button
        self.group_bar_layout.addStretch(1)
        self._update_group_button_state(self.current_group_name)

    def _update_group_button_state(self, active_group: str) -> None:
        if not hasattr(self, "group_buttons"):
            return
        for name, button in self.group_buttons.items():
            block = button.blockSignals(True)
            button.setChecked(name == active_group)
            button.blockSignals(block)

    def _update_score_display(self) -> None:
        if not hasattr(self, "score_label"):
            return
        if (
            self.current_student_index is None
            or self.student_data is None
            or self.student_data.empty
            or "成绩" not in self.student_data.columns
        ):
            self.score_label.setText("成绩：--")
            return
        value = self.student_data.at[self.current_student_index, "成绩"]
        try:
            score = int(value)
        except (TypeError, ValueError):
            score = 0
        self.score_label.setText(f"成绩：{score}")

    def _update_roll_call_controls(self) -> None:
        if not all(hasattr(self, attr) for attr in ("list_button", "add_score_button", "showcase_button")):
            return
        is_roll = self.mode == "roll_call"
        self.list_button.setVisible(is_roll)
        self.add_score_button.setVisible(is_roll)
        self.showcase_button.setVisible(is_roll)
        self.score_label.setVisible(is_roll)
        has_student = self.current_student_index is not None
        has_data = self.student_data is not None and not getattr(self.student_data, "empty", True)
        self.add_score_button.setEnabled(is_roll and has_student)
        self.list_button.setEnabled(is_roll and has_data)
        self.showcase_button.setEnabled(is_roll and has_data)

    def schedule_font_update(self) -> None:
        QTimer.singleShot(0, self.update_dynamic_fonts)

    def update_dynamic_fonts(self) -> None:
        name_font_size = self.last_name_font_size
        for lab in (self.id_label, self.name_label):
            if not lab.isVisible(): continue
            w = max(40, lab.width()); h = max(40, lab.height()); text = lab.text()
            size = self._calc_font_size(w, h, text)
            if lab is self.name_label:
                weight = QFont.Weight.Normal if self.name_font_family in {"楷体", "KaiTi"} else QFont.Weight.Bold
                lab.setFont(QFont(self.name_font_family, size, weight))
                self.last_name_font_size = size
                name_font_size = size
            else:
                lab.setFont(QFont("Microsoft YaHei UI", size, QFont.Weight.Bold))
                self.last_id_font_size = size
        if hasattr(self, "score_label") and self.score_label.isVisible():
            base = name_font_size if name_font_size > 0 else self.last_name_font_size
            if base <= 0:
                base = self.MIN_FONT_SIZE * 4
            score_size = max(1, int(round(base / 4)))
            self.score_label.setFont(QFont("Microsoft YaHei UI", score_size, QFont.Weight.Bold))
        if self.timer_frame.isVisible():
            text = self.time_display_label.text()
            w = max(60, self.time_display_label.width())
            h = max(60, self.time_display_label.height())
            size = self._calc_font_size(w, h, text, monospace=True)
            self.time_display_label.setFont(QFont("Consolas", size, QFont.Weight.Bold))
            self.last_timer_font_size = size

    def _calc_font_size(self, w: int, h: int, text: str, monospace: bool = False) -> int:
        if not text or w < 20 or h < 20:
            return self.MIN_FONT_SIZE
        w_eff = max(1, w - 16)
        h_eff = max(1, h - 16)
        is_cjk = any("\u4e00" <= c <= "\u9fff" for c in text)
        length = max(1, len(text))
        width_char_factor = 1.0 if is_cjk else (0.58 if monospace else 0.6)
        size_by_width = w_eff / (length * width_char_factor)
        size_by_height = h_eff * 0.70
        final_size = int(min(size_by_width, size_by_height))
        return max(self.MIN_FONT_SIZE, min(self.MAX_FONT_SIZE, final_size))

    def showEvent(self, e) -> None:
        super().showEvent(e)
        if self.mode == "roll_call" and self._placeholder_on_show:
            self.current_student_index = None
            self.display_current_student()
            self._placeholder_on_show = False
        elif self.mode == "timer":
            active_mode = self.timer_modes[self.timer_mode_index]
            if active_mode in {"countdown", "stopwatch"}:
                if self.reset_timer(persist=False):
                    self._schedule_save()
                self.update_timer_mode_ui()
        self.visibility_changed.emit(True)
        self.schedule_font_update()
        ensure_widget_within_screen(self)
        if not self._speech_issue_reported:
            self._diagnose_speech_engine()

    def resizeEvent(self, e: QResizeEvent) -> None:
        super().resizeEvent(e)
        self.schedule_font_update()

    def hideEvent(self, e) -> None:
        super().hideEvent(e)
        self._placeholder_on_show = True
        self.save_settings()
        self.visibility_changed.emit(False)

    def closeEvent(self, e) -> None:
        self.save_settings()
        self.count_timer.stop()
        self.clock_timer.stop()
        if self.tts_manager: self.tts_manager.shutdown()
        self.window_closed.emit()
        super().closeEvent(e)

    def _schedule_save(self) -> None:
        """延迟写入设置，避免频繁保存导致的磁盘抖动。"""

        if self._save_timer.isActive():
            self._save_timer.stop()
        self._save_timer.start()

    def save_settings(self) -> None:
        if self._save_timer.isActive():
            self._save_timer.stop()
        settings = self.settings_manager.load_settings()
        sec = settings.get("RollCallTimer", {})
        sec["geometry"] = geometry_to_text(self)
        sec["show_id"] = bool_to_str(self.show_id)
        sec["show_name"] = bool_to_str(self.show_name)
        sec["speech_enabled"] = bool_to_str(self.speech_enabled)
        sec["speech_voice_id"] = self.selected_voice_id
        sec["current_group"] = self.current_group_name
        sec["timer_countdown_minutes"] = str(self.timer_countdown_minutes)
        sec["timer_countdown_seconds"] = str(self.timer_countdown_seconds)
        sec["timer_sound_enabled"] = bool_to_str(self.timer_sound_enabled)
        sec["mode"] = self.mode
        sec["timer_mode"] = self.timer_modes[self.timer_mode_index]
        sec["timer_seconds_left"] = str(self.timer_seconds_left)
        sec["timer_stopwatch_seconds"] = str(self.timer_stopwatch_seconds)
        sec["timer_running"] = bool_to_str(self.timer_running)
        sec["id_font_size"] = str(self.last_id_font_size)
        sec["name_font_size"] = str(self.last_name_font_size)
        sec["timer_font_size"] = str(self.last_timer_font_size)
        sec["scoreboard_order"] = self.scoreboard_order
        remaining_payload: Dict[str, List[int]] = {}
        for group, indices in self._group_remaining_indices.items():
            cleaned: List[int] = []
            for idx in indices:
                try:
                    cleaned.append(int(idx))
                except (TypeError, ValueError):
                    continue
            remaining_payload[group] = cleaned
        last_payload: Dict[str, Optional[int]] = {}
        all_groups = set(self._group_all_indices.keys()) | set(self._group_last_student.keys())
        for group in all_groups:
            value = self._group_last_student.get(group)
            if value is None:
                last_payload[group] = None
            else:
                try:
                    last_payload[group] = int(value)
                except (TypeError, ValueError):
                    last_payload[group] = None
        try:
            sec["group_remaining"] = json.dumps(remaining_payload, ensure_ascii=False)
        except TypeError:
            sec["group_remaining"] = json.dumps({}, ensure_ascii=False)
        try:
            sec["group_last"] = json.dumps(last_payload, ensure_ascii=False)
        except TypeError:
            sec["group_last"] = json.dumps({}, ensure_ascii=False)
        try:
            global_drawn_payload = sorted(int(idx) for idx in self._global_drawn_students)
            sec["global_drawn"] = json.dumps(global_drawn_payload, ensure_ascii=False)
        except TypeError:
            sec["global_drawn"] = json.dumps([], ensure_ascii=False)
        settings["RollCallTimer"] = sec
        self.settings_manager.save_settings(settings)


# ---------- 关于 ----------
class AboutDialog(QDialog):
    def __init__(self, parent: Optional[QWidget] = None) -> None:
        super().__init__(parent)
        self.setWindowTitle("关于")
        self.setWindowFlag(Qt.WindowType.WindowStaysOnTopHint, True)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(18, 18, 18, 18)
        layout.setSpacing(12)
        info = QLabel(
            "<b>ClassroomTools</b><br>"
            "作者：广州番禺王耀强<br>"
            "知乎主页：<a href='https://www.zhihu.com/people/sciman/columns'>sciman</a><br>"
            "公众号：sciman逸居<br>"
            "“初中物理教研”Q群：323728546"
        )
        info.setOpenExternalLinks(True)
        info.setTextFormat(Qt.TextFormat.RichText)
        layout.addWidget(info)
        btn = QPushButton("确定")
        btn.clicked.connect(self.accept)
        btn.setCursor(Qt.CursorShape.PointingHandCursor)
        layout.addWidget(btn, alignment=Qt.AlignmentFlag.AlignCenter)
        self.setFixedSize(self.sizeHint())

    def showEvent(self, event) -> None:  # type: ignore[override]
        super().showEvent(event)
        ensure_widget_within_screen(self)


# ---------- 数据 ----------
def _normalize_text(value: object) -> str:
    if PANDAS_AVAILABLE and pd is not None:
        if pd.isna(value):
            return ""
    else:
        if value is None:
            return ""
        if isinstance(value, float) and math.isnan(value):
            return ""
        if isinstance(value, str) and not value:
            return ""
    text = str(value).strip()
    lowered = text.lower()
    if lowered in {"nan", "none", "nat"}:
        return ""
    return text


def _normalize_student_dataframe(df: pd.DataFrame, drop_incomplete: bool = True) -> pd.DataFrame:
    if not (PANDAS_AVAILABLE and pd is not None):
        return df.copy()

    normalized = df.copy()
    for column in ("学号", "姓名", "分组", "成绩"):
        if column not in normalized.columns:
            normalized[column] = pd.NA

    normalized["姓名"] = normalized["姓名"].apply(lambda v: re.sub(r"\s+", "", _normalize_text(v)))
    normalized["分组"] = normalized["分组"].apply(lambda v: re.sub(r"\s+", "", _normalize_text(v))).str.upper()

    id_series = normalized["学号"].apply(lambda v: re.sub(r"\s+", "", _normalize_text(v)))
    id_numeric = pd.to_numeric(id_series.replace("", pd.NA), errors="coerce")
    if not id_numeric.empty:
        fractional_mask = id_numeric.notna() & (id_numeric != id_numeric.round())
        id_numeric = id_numeric.where(~fractional_mask)
    normalized["学号"] = id_numeric.round().astype("Int64")

    score_series = normalized["成绩"].apply(lambda v: re.sub(r"\s+", "", _normalize_text(v)))
    score_numeric = pd.to_numeric(score_series.replace("", pd.NA), errors="coerce").fillna(0)
    score_numeric = score_numeric.round()
    normalized["成绩"] = score_numeric.astype("Int64")

    for column in normalized.select_dtypes(include=["object"]).columns:
        if column in {"姓名", "分组"}:
            continue
        normalized[column] = normalized[column].apply(_normalize_text)

    ordered_columns = [col for col in ["学号", "姓名", "分组", "成绩"] if col in normalized.columns]
    extra_columns = [col for col in normalized.columns if col not in ordered_columns]
    normalized = normalized[ordered_columns + extra_columns]

    if drop_incomplete:
        normalized = normalized[(normalized["学号"].notna()) & (normalized["姓名"] != "")].copy()
        normalized.reset_index(drop=True, inplace=True)

    return normalized


def _write_student_workbook(file_path: str, df: pd.DataFrame) -> None:
    try:
        export_df = _normalize_student_dataframe(df, drop_incomplete=False)
    except Exception:
        export_df = df.copy()

    if not OPENPYXL_AVAILABLE:
        export_df.to_excel(file_path, index=False)
        return

    try:
        from openpyxl import Workbook
        from openpyxl.styles import Font
    except Exception:
        export_df.to_excel(file_path, index=False)
        return

    try:
        workbook = Workbook()
        worksheet = workbook.active
        worksheet.title = "students"

        headers = list(export_df.columns)
        worksheet.append(headers)
        header_font = Font(name="等线", size=12, bold=True)
        body_font = Font(name="等线", size=12)
        for cell in worksheet[1]:
            cell.font = header_font

        for row_values in export_df.itertuples(index=False, name=None):
            normalized_row = []
            for value in row_values:
                if pd.isna(value):
                    normalized_row.append(None)
                else:
                    normalized_row.append(value)
            worksheet.append(tuple(normalized_row))

        for row in worksheet.iter_rows(min_row=2):
            for cell in row:
                cell.font = body_font
                if isinstance(cell.value, str):
                    cell.value = cell.value.strip()

        workbook.save(file_path)
    except Exception:
        export_df.to_excel(file_path, index=False)


def load_student_data(parent: Optional[QWidget]) -> Optional[pd.DataFrame]:
    """从 students.xlsx 读取点名所需的数据，不存在时自动生成模板。"""
    if not (PANDAS_AVAILABLE and OPENPYXL_AVAILABLE):
        QMessageBox.warning(parent, "提示", "未安装 pandas/openpyxl，点名功能不可用。")
        return None
    file_path = RollCallTimerWindow.STUDENT_FILE
    if not os.path.exists(file_path):
        try:
            df = pd.DataFrame(
                {"学号": [101, 102, 103], "姓名": ["张三", "李四", "王五"], "分组": ["A", "B", "A"], "成绩": [0, 0, 0]}
            )
            df = _normalize_student_dataframe(df)
            _write_student_workbook(file_path, df)
            show_quiet_information(parent, f"未找到学生名单，已为您创建模板文件：{file_path}")
        except Exception as exc:
            QMessageBox.critical(parent, "错误", f"创建模板文件失败：{exc}")
            return None
    try:
        df = pd.read_excel(file_path)
        df = _normalize_student_dataframe(df)
        _write_student_workbook(file_path, df)
        return df
    except Exception as exc:
        QMessageBox.critical(parent, "错误", f"无法加载学生名单，请检查文件格式。\n错误：{exc}")
        return None


# ---------- 启动器 ----------
class LauncherBubble(QWidget):
    """启动器缩小时显示的悬浮圆球，负责发出恢复指令。"""

    restore_requested = pyqtSignal()
    position_changed = pyqtSignal(QPoint)

    def __init__(self, diameter: int = 42) -> None:
        super().__init__(
            None,
            Qt.WindowType.Tool
            | Qt.WindowType.FramelessWindowHint
            | Qt.WindowType.WindowStaysOnTopHint,
        )
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground, True)
        self.setAttribute(Qt.WidgetAttribute.WA_NoSystemBackground, True)
        self.setAttribute(Qt.WidgetAttribute.WA_ShowWithoutActivating, True)
        self.setCursor(Qt.CursorShape.PointingHandCursor)
        self.setWindowTitle("ClassroomTools Bubble")
        self._diameter = max(32, diameter)
        self.setFixedSize(self._diameter, self._diameter)
        self.setWindowOpacity(0.74)
        self._ensure_min_width = self._diameter
        self._ensure_min_height = self._diameter
        self._dragging = False
        self._drag_offset = QPoint()
        self._moved = False

    def place_near(self, target: QPoint, screen: Optional[QScreen]) -> None:
        """将气泡吸附到距离 target 最近的屏幕边缘。"""

        if screen is None:
            screen = QApplication.screenAt(target) or QApplication.primaryScreen()
        if screen is None:
            return
        available = screen.availableGeometry()
        margin = 6
        bubble_size = self.size()
        center = QPoint(int(target.x()), int(target.y()))
        center.setX(max(available.left(), min(center.x(), available.right())))
        center.setY(max(available.top(), min(center.y(), available.bottom())))

        distances = {
            "left": abs(center.x() - available.left()),
            "right": abs(available.right() - center.x()),
            "top": abs(center.y() - available.top()),
            "bottom": abs(available.bottom() - center.y()),
        }
        nearest_edge = min(distances, key=distances.get)
        if nearest_edge == "left":
            x = available.left() + margin
            y = center.y() - bubble_size.height() // 2
        elif nearest_edge == "right":
            x = available.right() - bubble_size.width() - margin
            y = center.y() - bubble_size.height() // 2
        elif nearest_edge == "top":
            y = available.top() + margin
            x = center.x() - bubble_size.width() // 2
        else:
            y = available.bottom() - bubble_size.height() - margin
            x = center.x() - bubble_size.width() // 2

        x = max(available.left() + margin, min(x, available.right() - bubble_size.width() - margin))
        y = max(available.top() + margin, min(y, available.bottom() - bubble_size.height() - margin))
        self.move(int(x), int(y))
        self.position_changed.emit(self.pos())

    def snap_to_edge(self) -> None:
        screen = self.screen() or QApplication.screenAt(self.frameGeometry().center()) or QApplication.primaryScreen()
        if screen is None:
            return
        self.place_near(self.frameGeometry().center(), screen)

    def paintEvent(self, event) -> None:
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing, True)
        rect = self.rect()
        color = QColor(32, 33, 36, 150)
        highlight = QColor(138, 180, 248, 160)
        painter.setBrush(QBrush(color))
        painter.setPen(Qt.PenStyle.NoPen)
        painter.drawEllipse(rect)
        painter.setBrush(QBrush(highlight))
        painter.drawEllipse(rect.adjusted(rect.width() // 3, rect.height() // 3, -rect.width() // 6, -rect.height() // 6))

    def mousePressEvent(self, event) -> None:
        if event.button() == Qt.MouseButton.LeftButton:
            self._dragging = True
            self._drag_offset = event.position().toPoint()
            self._moved = False
        super().mousePressEvent(event)

    def mouseMoveEvent(self, event) -> None:
        if self._dragging and event.buttons() & Qt.MouseButton.LeftButton:
            new_pos = event.globalPosition().toPoint() - self._drag_offset
            self.move(new_pos)
            self._moved = True
        super().mouseMoveEvent(event)

    def mouseReleaseEvent(self, event) -> None:
        if event.button() == Qt.MouseButton.LeftButton:
            if not self._moved:
                self.restore_requested.emit()
            else:
                self.snap_to_edge()
                self.position_changed.emit(self.pos())
            self._dragging = False
        super().mouseReleaseEvent(event)

    def showEvent(self, event) -> None:  # type: ignore[override]
        super().showEvent(event)
        ensure_widget_within_screen(self)


class LauncherWindow(QWidget):
    def __init__(self, settings_manager: SettingsManager, student_data: Optional[pd.DataFrame]) -> None:
        super().__init__(None, Qt.WindowType.Tool | Qt.WindowType.FramelessWindowHint | Qt.WindowType.WindowStaysOnTopHint)
        self.settings_manager = settings_manager
        self.student_data = student_data
        self.overlay: Optional[OverlayWindow] = None
        self.roll_call_window: Optional[RollCallTimerWindow] = None
        self._dragging = False; self._drag_offset = QPoint()
        self.bubble: Optional[LauncherBubble] = None
        self._last_position = QPoint()
        self._bubble_position = QPoint()
        self._minimized = False
        self._minimized_on_start = False

        self.setAttribute(Qt.WidgetAttribute.WA_StyledBackground, True)
        self.setStyleSheet(
            """
            QWidget#launcherContainer {
                background-color: rgba(28, 29, 32, 232);
                border-radius: 9px;
                border: 1px solid rgba(255, 255, 255, 40);
            }
            QPushButton {
                color: #f1f3f4;
                background-color: rgba(60, 64, 67, 230);
                border: 1px solid rgba(255, 255, 255, 35);
                border-radius: 6px;
                padding: 3px 9px;
                min-height: 26px;
            }
            QPushButton:hover {
                background-color: rgba(138, 180, 248, 240);
                border-color: rgba(138, 180, 248, 255);
                color: #202124;
            }
            QCheckBox {
                color: #f1f3f4;
                spacing: 4px;
            }
            QCheckBox::indicator {
                width: 14px;
                height: 14px;
                border-radius: 4px;
                border: 1px solid rgba(255, 255, 255, 60);
                background: rgba(32, 33, 36, 240);
            }
            QCheckBox::indicator:checked {
                background-color: #8ab4f8;
                border-color: transparent;
            }
            """
        )
        container = QWidget(self); container.setObjectName("launcherContainer")
        layout = QVBoxLayout(self); layout.setContentsMargins(0, 0, 0, 0); layout.addWidget(container)
        v = QVBoxLayout(container); v.setContentsMargins(8, 8, 8, 8); v.setSpacing(5)

        # 通过三段可伸缩空白保证“画笔”“点名/计时”两侧及中间留白一致
        row = QGridLayout(); row.setContentsMargins(0, 0, 0, 0); row.setHorizontalSpacing(0)
        for col in (0, 2, 4):
            row.setColumnMinimumWidth(col, 12)
            row.setColumnStretch(col, 1)
        self.paint_button = QPushButton("画笔")
        self.paint_button.clicked.connect(self.toggle_paint)
        row.addItem(QSpacerItem(0, 0, QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Minimum), 0, 0)
        row.addWidget(self.paint_button, 0, 1)
        row.addItem(QSpacerItem(0, 0, QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Minimum), 0, 2)
        self.roll_call_button = QPushButton("点名/计时")
        self.roll_call_button.clicked.connect(self.toggle_roll_call)
        row.addWidget(self.roll_call_button, 0, 3)
        row.addItem(QSpacerItem(0, 0, QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Minimum), 0, 4)
        unified_width = self._action_button_width()
        self.paint_button.setFixedWidth(unified_width)
        self.roll_call_button.setFixedWidth(unified_width)
        v.addLayout(row)

        bottom = QHBoxLayout(); bottom.setSpacing(3)
        self.minimize_button = QPushButton("缩小"); self.minimize_button.clicked.connect(self.minimize_launcher)
        bottom.addWidget(self.minimize_button)

        self.autostart_check = QCheckBox("开机启动"); self.autostart_check.stateChanged.connect(self.toggle_autostart); bottom.addWidget(self.autostart_check)
        bottom.addStretch(1)

        self.about_button = QPushButton("关于"); self.about_button.clicked.connect(self.show_about); bottom.addWidget(self.about_button)
        self.exit_button = QPushButton("退出")
        self.exit_button.clicked.connect(self.request_exit)
        bottom.addWidget(self.exit_button)
        v.addLayout(bottom)

        aux_width = max(self.minimize_button.sizeHint().width(), 52)
        self.minimize_button.setFixedWidth(aux_width)
        about_width = max(self.about_button.sizeHint().width(), 52)
        exit_width = max(self.exit_button.sizeHint().width(), 52)
        self.about_button.setFixedWidth(about_width)
        self.exit_button.setFixedWidth(exit_width)

        button_height = max(
            self.paint_button.sizeHint().height(),
            self.roll_call_button.sizeHint().height(),
            self.minimize_button.sizeHint().height(),
            self.about_button.sizeHint().height(),
            self.exit_button.sizeHint().height(),
        )
        for btn in (self.paint_button, self.roll_call_button, self.minimize_button, self.about_button, self.exit_button):
            btn.setFixedHeight(button_height)

        s = self.settings_manager.load_settings().get("Launcher", {})
        x = int(s.get("x", "120")); y = int(s.get("y", "120"))
        self.move(x, y)
        self._last_position = QPoint(x, y)
        bubble_x = int(s.get("bubble_x", str(x)))
        bubble_y = int(s.get("bubble_y", str(y)))
        self._bubble_position = QPoint(bubble_x, bubble_y)
        minimized = str_to_bool(s.get("minimized", "False"), False)
        self._minimized = minimized
        self._minimized_on_start = minimized

        startup = self.settings_manager.load_settings().get("Startup", {})
        autostart_enabled = str_to_bool(startup.get("autostart_enabled", "False"), False)
        self.autostart_check.setChecked(autostart_enabled and WINREG_AVAILABLE)
        self.autostart_check.setEnabled(WINREG_AVAILABLE)

        if not PANDAS_AVAILABLE or not OPENPYXL_AVAILABLE or self.student_data is None:
            self.roll_call_button.setEnabled(False)

        for w in (self, container, self.paint_button, self.roll_call_button, self.minimize_button, self.autostart_check):
            w.installEventFilter(self)

        # 锁定启动器的推荐尺寸，避免误拖拽造成遮挡
        self.adjustSize()
        self.setFixedSize(self.sizeHint())
        self._base_minimum_width = self.minimumWidth()
        self._base_minimum_height = self.minimumHeight()
        self._ensure_min_width = self.width()
        self._ensure_min_height = self.height()

    def _action_button_width(self) -> int:
        """计算“画笔”与“点名/计时”按钮的统一宽度，保证观感一致。"""

        paint_metrics = QFontMetrics(self.paint_button.font())
        roll_metrics = QFontMetrics(self.roll_call_button.font())
        paint_texts = ["画笔", "隐藏画笔"]
        roll_texts = ["点名/计时", "显示点名", "隐藏点名"]
        max_width = max(
            max(paint_metrics.horizontalAdvance(text) for text in paint_texts),
            max(roll_metrics.horizontalAdvance(text) for text in roll_texts),
        )
        return max_width + 28

    def showEvent(self, e) -> None:
        super().showEvent(e)
        ensure_widget_within_screen(self)
        self._last_position = self.pos()
        if self._minimized_on_start:
            QTimer.singleShot(0, self._restore_minimized_state)

    def eventFilter(self, obj, e) -> bool:
        if e.type() == QEvent.Type.MouseButtonPress and e.button() == Qt.MouseButton.LeftButton:
            self._dragging = True; self._drag_offset = e.globalPosition().toPoint() - self.pos()
        elif e.type() == QEvent.Type.MouseMove and self._dragging and e.buttons() & Qt.MouseButton.LeftButton:
            self.move(e.globalPosition().toPoint() - self._drag_offset)
        elif e.type() == QEvent.Type.MouseButtonRelease and e.button() == Qt.MouseButton.LeftButton:
            self._dragging = False
            self._last_position = self.pos()
            self.save_position()
        return super().eventFilter(obj, e)

    def save_position(self) -> None:
        settings = self.settings_manager.load_settings()
        pos = self._last_position if (self._minimized and not self._last_position.isNull()) else self.pos()
        launcher = settings.get("Launcher", {})
        launcher["x"] = str(pos.x())
        launcher["y"] = str(pos.y())
        bubble_pos = self._bubble_position if not self._bubble_position.isNull() else pos
        launcher["bubble_x"] = str(bubble_pos.x())
        launcher["bubble_y"] = str(bubble_pos.y())
        launcher["minimized"] = bool_to_str(self._minimized)
        settings["Launcher"] = launcher
        startup = settings.get("Startup", {}); startup["autostart_enabled"] = bool_to_str(self.autostart_check.isChecked()); settings["Startup"] = startup
        self.settings_manager.save_settings(settings)

    def toggle_paint(self) -> None:
        """打开或隐藏屏幕画笔覆盖层。"""
        if self.overlay is None: self.overlay = OverlayWindow(self.settings_manager)
        if self.overlay.isVisible():
            self.overlay.hide_overlay(); self.paint_button.setText("画笔")
        else:
            self.overlay.show_overlay(); self.paint_button.setText("隐藏画笔")

    def toggle_roll_call(self) -> None:
        """切换点名/计时窗口的显示状态，必要时先创建窗口。"""
        if self.student_data is None:
            QMessageBox.warning(self, "提示", "学生数据加载失败，无法打开点名器。"); return
        if self.roll_call_window is None:
            self.roll_call_window = RollCallTimerWindow(self.settings_manager, self.student_data)
            self.roll_call_window.window_closed.connect(self.on_roll_call_window_closed)
            self.roll_call_window.visibility_changed.connect(self.on_roll_call_visibility_changed)
            self.roll_call_window.show()
            self.roll_call_button.setText("隐藏点名")
        else:
            if self.roll_call_window.isVisible():
                self.roll_call_window.hide()
                self.roll_call_button.setText("显示点名")
            else:
                self.roll_call_window.show(); self.roll_call_window.raise_()
                self.roll_call_button.setText("隐藏点名")

    def on_roll_call_window_closed(self) -> None:
        self.roll_call_window = None
        self.roll_call_button.setText("点名/计时")

    def on_roll_call_visibility_changed(self, visible: bool) -> None:
        self.roll_call_button.setText("隐藏点名" if visible else "显示点名")

    def toggle_autostart(self) -> None:
        if not WINREG_AVAILABLE: return
        enabled = self.autostart_check.isChecked()
        self.set_autostart(enabled); self.save_position()

    def request_exit(self) -> None:
        """优雅地关闭应用程序，确保所有窗口在退出前持久化状态。"""

        app = QApplication.instance()
        if not self.close():
            return
        if app is not None:
            QTimer.singleShot(0, app.quit)

    def handle_about_to_quit(self) -> None:
        """在应用退出前的最后一道保险，保证关键状态已经写入配置。"""

        self.save_position()
        window = self.roll_call_window
        if window is not None:
            try:
                window.save_settings()
            except RuntimeError:
                pass

    def minimize_launcher(self, from_settings: bool = False) -> None:
        """将启动器收纳为悬浮圆球。"""

        if self._minimized and not from_settings:
            return
        if self.bubble is None:
            self.bubble = LauncherBubble()
            self.bubble.restore_requested.connect(self.restore_from_bubble)
            self.bubble.position_changed.connect(self._on_bubble_position_changed)
        target_center = self.frameGeometry().center()
        screen = self.screen() or QApplication.screenAt(target_center) or QApplication.primaryScreen()
        if not from_settings:
            self._last_position = self.pos()
        self.hide()
        if from_settings and not self._bubble_position.isNull():
            self.bubble.place_near(self._bubble_position, screen)
        else:
            self.bubble.place_near(target_center, screen)
        self.bubble.setWindowOpacity(0.74)
        self.bubble.show()
        self.bubble.raise_()
        self._minimized = True
        self.save_position()

    def _restore_minimized_state(self) -> None:
        if not self._minimized_on_start:
            return
        self._minimized_on_start = False
        self.minimize_launcher(from_settings=True)

    def restore_from_bubble(self) -> None:
        """从悬浮球恢复主启动器窗口。"""

        self._minimized = False
        target_pos: Optional[QPoint] = None
        screen = None
        if self.bubble:
            self._bubble_position = self.bubble.pos()
            bubble_geom = self.bubble.frameGeometry()
            bubble_center = bubble_geom.center()
            screen = self.bubble.screen() or QApplication.screenAt(bubble_center) or QApplication.primaryScreen()
            margin = 12
            width = self.width() or self.sizeHint().width()
            height = self.height() or self.sizeHint().height()
            if screen is not None:
                available = screen.availableGeometry()
                distances = {
                    "left": abs(bubble_center.x() - available.left()),
                    "right": abs(available.right() - bubble_center.x()),
                    "top": abs(bubble_center.y() - available.top()),
                    "bottom": abs(available.bottom() - bubble_center.y()),
                }
                nearest_edge = min(distances, key=distances.get)
                if nearest_edge == "left":
                    x = bubble_geom.right() + margin
                    y = bubble_center.y() - height // 2
                elif nearest_edge == "right":
                    x = bubble_geom.left() - width - margin
                    y = bubble_center.y() - height // 2
                elif nearest_edge == "top":
                    y = bubble_geom.bottom() + margin
                    x = bubble_center.x() - width // 2
                else:
                    y = bubble_geom.top() - height - margin
                    x = bubble_center.x() - width // 2
                x = max(available.left(), min(int(x), available.right() - width))
                y = max(available.top(), min(int(y), available.bottom() - height))
                target_pos = QPoint(x, y)
            self.bubble.hide()
        if target_pos is None and not self._last_position.isNull():
            target_pos = QPoint(self._last_position)
        if target_pos is not None:
            self.move(target_pos)
        self.show()
        self.raise_()
        self.activateWindow()
        ensure_widget_within_screen(self)
        self._last_position = self.pos()
        self.save_position()

    def _on_bubble_position_changed(self, pos: QPoint) -> None:
        self._bubble_position = QPoint(pos.x(), pos.y())
        if self._minimized:
            self.save_position()

    def set_autostart(self, enabled: bool) -> None:
        if not WINREG_AVAILABLE: return
        key_path = r"Software\\Microsoft\\Windows\\CurrentVersion\\Run"
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path, 0, winreg.KEY_SET_VALUE) as key:
                if enabled:
                    script_path = self.get_script_path()
                    if script_path.lower().endswith((".py", ".pyw")):
                        pythonw = os.path.join(sys.prefix, "pythonw.exe")
                        if not os.path.exists(pythonw): pythonw = sys.executable
                        value = f'"{pythonw}" "{script_path}"'
                    else:
                        value = f'"{script_path}"'
                    winreg.SetValueEx(key, "ClassroomTools", 0, winreg.REG_SZ, value)
                else:
                    try: winreg.DeleteValue(key, "ClassroomTools")
                    except FileNotFoundError: pass
        except PermissionError:
            QMessageBox.warning(self, "提示", "写入开机启动失败，请以管理员身份运行。")
        except Exception as exc:
            QMessageBox.warning(self, "提示", f"设置开机启动失败：{exc}")

    def get_script_path(self) -> str:
        if getattr(sys, "frozen", False): return sys.executable
        return os.path.abspath(sys.argv[0])

    def show_about(self) -> None:
        AboutDialog(self).exec()

    def closeEvent(self, e) -> None:
        self.save_position()
        if self.bubble is not None:
            self.bubble.close()
        if self.roll_call_window is not None: self.roll_call_window.close()
        if self.overlay is not None: self.overlay.close()
        super().closeEvent(e)


# ---------- 入口 ----------
def main() -> None:
    """应用程序入口：初始化 DPI、加载设置并启动启动器窗口。"""
    ensure_high_dpi_awareness()
    app = QApplication(sys.argv)
    app.setQuitOnLastWindowClosed(False)
    QToolTip.setFont(QFont("Microsoft YaHei UI", 9))

    settings_manager = SettingsManager()
    student_data = load_student_data(None) if PANDAS_AVAILABLE else None

    window = LauncherWindow(settings_manager, student_data)
    app.aboutToQuit.connect(window.handle_about_to_quit)
    window.show()
    sys.exit(app.exec())


# Nuitka 打包指令（根据当前依赖整理的推荐参数，保持在单行便于复制）：
# 单文件：python -m nuitka --onefile --enable-plugin=pyqt6 --include-qt-plugins=sensible --windows-disable-console --windows-icon-from-ico=icon.ico --include-data-file=students.xlsx=students.xlsx --include-data-file=settings.ini=settings.ini ClassroomTools.py
# 独立目录：python -m nuitka --standalone --enable-plugin=pyqt6 --include-qt-plugins=sensible --windows-disable-console --windows-icon-from-ico=icon.ico --output-dir=dist --include-data-file=students.xlsx=students.xlsx --include-data-file=settings.ini=settings.ini ClassroomTools.py
if __name__ == "__main__":
    main()
