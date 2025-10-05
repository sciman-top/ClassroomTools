# -*- coding: utf-8 -*-
from __future__ import annotations

import base64
import configparser
import ctypes
import json
import os
import random
import sys
import threading
import time
from queue import Empty, Queue
from typing import Dict, List, Optional, Mapping

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
    QComboBox,
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

    screen = QApplication.screenAt(QPoint(x, y))
    if screen is None:
        try:
            screen = widget.screen() or QApplication.primaryScreen()
        except Exception:
            screen = QApplication.primaryScreen()
    if screen is not None:
        available = screen.availableGeometry()
        max_width = max(base_min_width, 320, int(available.width() * 0.9))
        max_height = max(base_min_height, 240, int(available.height() * 0.9))
        width = max(base_min_width, min(width, max_width))
        height = max(base_min_height, min(height, max_height))
        x = max(available.left(), min(x, available.right() - width))
        y = max(available.top(), min(y, available.bottom() - height))
    target_width = max(base_min_width, max(160, width))
    target_height = max(base_min_height, max(120, height))
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

    available = screen.availableGeometry()
    geom = widget.frameGeometry()
    width = widget.width() or geom.width() or widget.sizeHint().width()
    height = widget.height() or geom.height() or widget.sizeHint().height()
    max_width = min(available.width(), max(base_min_width, int(available.width() * 0.9)))
    max_height = min(available.height(), max(base_min_height, int(available.height() * 0.9)))
    width = max(base_min_width, min(width, max_width))
    height = max(base_min_height, min(height, max_height))
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
        self.filename = filename
        self.config = configparser.ConfigParser()
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
            },
            "Paint": {"x": "260", "y": "260", "brush_size": "12", "brush_color": "#ff0000"},
        }

    def get_defaults(self) -> Dict[str, Dict[str, str]]:
        return {section: values.copy() for section, values in self.defaults.items()}

    def load_settings(self) -> Dict[str, Dict[str, str]]:
        settings = self.get_defaults()
        if not os.path.exists(self.filename):
            return settings
        try:
            self.config.read(self.filename, encoding="utf-8")
            for section in self.config.sections():
                if section not in settings:
                    settings[section] = {}
                for key, value in self.config.items(section):
                    settings[section][key] = value
        except configparser.Error:
            return self.get_defaults()
        return settings

    def save_settings(self, settings: Dict[str, Dict[str, str]]) -> None:
        config = configparser.ConfigParser()
        for section, options in settings.items():
            config[section] = {key: str(value) for key, value in options.items()}
        try:
            with open(self.filename, "w", encoding="utf-8") as handle:
                config.write(handle)
        except IOError:
            print(f"无法写入配置文件: {self.filename}")

    def clear_roll_call_history(self) -> None:
        """清空点名相关的历史记录，保证新启动时重新开始。"""

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

        self.btn_cursor.clicked.connect(lambda: self.overlay.set_mode("cursor"))
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
        self._eraser_last_point: Optional[QPoint] = None
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
        if self.mode == "brush" and (not initial or prev_mode != "brush"):
            self._remember_brush_state()
        self._update_visibility_for_mode(initial=initial)
        if not initial:
            self.raise_toolbar()
        self.update_toolbar_state()
        self.update_cursor()

    def update_toolbar_state(self) -> None:
        if not getattr(self, 'toolbar', None):
            return
        self.toolbar.update_tool_states(self.mode, self.pen_color)

    def _remember_brush_state(self) -> None:
        if self.pen_color.isValid():
            self._last_brush_color = QColor(self.pen_color)
        if self.pen_size > 0:
            self._last_brush_size = max(1, int(self.pen_size))

    def _restore_brush_mode(self) -> None:
        if isinstance(self._last_brush_color, QColor) and self._last_brush_color.isValid():
            self.pen_color = QColor(self._last_brush_color)
        if isinstance(self._last_brush_size, int) and self._last_brush_size > 0:
            self.pen_size = max(1, int(self._last_brush_size))
        self.set_mode("brush")

    def toggle_eraser_mode(self) -> None:
        """切换橡皮模式；再次点击会恢复上一次的画笔配置。"""
        if self.mode == "eraser":
            self._restore_brush_mode()
        else:
            self._remember_brush_state()
            self.set_mode("eraser")

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
        restore_needed = self.mode != "brush"
        self._push_history()
        self.canvas.fill(Qt.GlobalColor.transparent)
        self.temp_canvas.fill(Qt.GlobalColor.transparent)
        self.update()
        self._eraser_last_point = None
        if restore_needed:
            self._restore_brush_mode()
        else:
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
            self.last_width = self.pen_size * 0.4
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
            self.raise_toolbar()
        super().mouseReleaseEvent(e)

    def keyPressEvent(self, e: QKeyEvent) -> None:
        if e.key() == Qt.Key.Key_Escape:
            self.set_mode("cursor"); return
        super().keyPressEvent(e)

    def _draw_brush_line(self, cur: QPointF) -> None:
        cur = QPointF(cur); self.last_point = QPointF(self.last_point); self.prev_point = QPointF(self.prev_point)
        distance = (cur - self.last_point).manhattanLength()
        speed = distance / max(1.0, self.pen_size)
        speed_factor = 1.0 - min(1.0, speed)
        delta_prev = self.last_point - self.prev_point; delta_current = cur - self.last_point
        angle = abs(delta_prev.x() * delta_current.y() - delta_prev.y() * delta_current.x())
        angle_factor = min(1.0, angle / (self.pen_size * 20.0))
        elapsed = time.time() - self.last_time; self.last_time = time.time()
        pressure = min(1.0, elapsed * 5.0)
        target_w = self.pen_size * (0.35 + 0.55 * speed_factor + 0.35 * angle_factor) * (1.0 + 0.4 * pressure)
        cur_w = self.last_width * 0.65 + target_w * 0.35
        mid = (self.last_point + cur) / 2.0

        path = QPainterPath(); path.moveTo(self.prev_point); path.quadTo(self.last_point, mid)
        p = QPainter(self.canvas); p.setRenderHint(QPainter.RenderHint.Antialiasing)
        fade = QColor(self.pen_color); fade.setAlpha(70)
        p.setPen(QPen(fade, cur_w * 1.4, Qt.PenStyle.SolidLine, Qt.PenCapStyle.RoundCap, Qt.PenJoinStyle.RoundJoin)); p.drawPath(path)
        p.setPen(QPen(self.pen_color, cur_w, Qt.PenStyle.SolidLine, Qt.PenCapStyle.RoundCap, Qt.PenJoinStyle.RoundJoin)); p.drawPath(path)
        p.end()

        self.prev_point = self.last_point; self.last_point = cur; self.last_width = cur_w

    def _erase_at(self, pos) -> None:
        if isinstance(pos, QPointF):
            pos = pos.toPoint()
        start = self._eraser_last_point or pos
        p = QPainter(self.canvas)
        p.setRenderHint(QPainter.RenderHint.Antialiasing)
        p.setCompositionMode(QPainter.CompositionMode.CompositionMode_Clear)
        diameter = max(10, int(self.pen_size * 1.8) * 2)
        pen = QPen(QColor(255, 255, 255, 0), diameter, Qt.PenStyle.SolidLine, Qt.PenCapStyle.RoundCap, Qt.PenJoinStyle.RoundJoin)
        p.setPen(pen)
        p.drawLine(start, pos)
        p.end()
        self._eraser_last_point = QPoint(pos)

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
    """简单封装 pyttsx3，实现点名时的语音播报。"""

    def __init__(self, preferred_voice_id: str = "", parent: Optional[QObject] = None) -> None:
        super().__init__(parent)
        self.engine = None
        self.voice_ids: List[str] = []
        self.default_voice_id = ""
        self.current_voice_id = ""
        self._queue: Queue[str] = Queue()
        self._timer = QTimer(self)
        self._timer.timeout.connect(self._pump)
        try:
            self.engine = pyttsx3.init()
            voices = self.engine.getProperty("voices") or []
            self.voice_ids = [v.id for v in voices]
            if self.voice_ids: self.default_voice_id = self.voice_ids[0]
            self.current_voice_id = preferred_voice_id if preferred_voice_id in self.voice_ids else self.default_voice_id
            if self.current_voice_id: self.engine.setProperty("voice", self.current_voice_id)
            self.engine.startLoop(False); self._timer.start(100)
        except Exception:
            self.engine = None

    @property
    def available(self) -> bool:
        return self.engine is not None

    def set_voice(self, voice_id: str) -> None:
        if voice_id in self.voice_ids:
            self.current_voice_id = voice_id
            if self.engine:
                try:
                    self.engine.setProperty("voice", voice_id)
                except Exception:
                    pass

    def speak(self, text: str) -> None:
        if not self.engine: return
        while not self._queue.empty():
            try: self._queue.get_nowait()
            except Empty: break
        self._queue.put(text)

    def _pump(self) -> None:
        if not self.engine: return
        try:
            text = self._queue.get_nowait()
            self.engine.stop()
            if self.current_voice_id: self.engine.setProperty("voice", self.current_voice_id)
            self.engine.say(text)
        except Empty:
            pass
        try:
            self.engine.iterate()
        except Exception:
            self.shutdown()

    def shutdown(self) -> None:
        if self.engine:
            try: self.engine.endLoop()
            except Exception: pass
            try: self.engine.stop()
            except Exception: pass
        self.engine = None; self._timer.stop()


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


class ClickableFrame(QFrame):
    clicked = pyqtSignal()
    def mousePressEvent(self, e) -> None:
        if e.button() == Qt.MouseButton.LeftButton: self.clicked.emit()
        super().mousePressEvent(e)


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
        self.setWindowFlag(Qt.WindowType.WindowStaysOnTopHint, True)
        self.setAttribute(Qt.WidgetAttribute.WA_DeleteOnClose)
        self.settings_manager = settings_manager
        self.student_data = student_data

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

        self.mode = s.get("mode", "roll_call") if s.get("mode", "roll_call") in {"roll_call", "timer"} else "roll_call"
        self.timer_modes = ["countdown", "stopwatch", "clock"]
        self.timer_mode_index = self.timer_modes.index(s.get("timer_mode", "countdown")) if s.get("timer_mode", "countdown") in self.timer_modes else 0

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

        self.last_id_font_size = max(self.MIN_FONT_SIZE, _get_int("id_font_size", 48))
        self.last_name_font_size = max(self.MIN_FONT_SIZE, _get_int("name_font_size", 60))
        self.last_timer_font_size = max(self.MIN_FONT_SIZE, _get_int("timer_font_size", 56))

        self.count_timer = QTimer(self); self.count_timer.setInterval(1000); self.count_timer.timeout.connect(self._on_count_timer)
        self.clock_timer = QTimer(self); self.clock_timer.setInterval(1000); self.clock_timer.timeout.connect(self._update_clock)

        self.tts_manager: Optional[TTSManager] = None
        self.speech_enabled = str_to_bool(s.get("speech_enabled", "False"), False) and PYTTSX3_AVAILABLE
        self.selected_voice_id = s.get("speech_voice_id", "")
        if PYTTSX3_AVAILABLE:
            self.tts_manager = TTSManager(self.selected_voice_id, parent=self)
            if not self.tts_manager.available: self.tts_manager = None; self.speech_enabled = False
        else:
            self.speech_enabled = False

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

        self._build_ui()
        self._apply_saved_fonts()
        self._update_menu_state()
        self.update_mode_ui()
        self.on_group_change(initial=True)
        self.display_current_student()

    def _build_ui(self) -> None:
        self.setStyleSheet("background-color: #f4f5f7;")
        layout = QVBoxLayout(self); layout.setContentsMargins(8, 8, 8, 8); layout.setSpacing(6)

        top = QHBoxLayout(); top.setSpacing(4)
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

        self.group_label = QLabel("分组")
        self.group_label.setFont(QFont("Microsoft YaHei UI", 9, QFont.Weight.Medium))
        self.group_label.setStyleSheet("color: #3c4043;")
        self.group_label.setFixedHeight(28)
        self.group_label.setAlignment(Qt.AlignmentFlag.AlignVCenter | Qt.AlignmentFlag.AlignLeft)
        self.group_label.setSizePolicy(QSizePolicy.Policy.Fixed, QSizePolicy.Policy.Fixed)
        top.addWidget(self.group_label, 0, Qt.AlignmentFlag.AlignLeft)

        self.group_combo = QComboBox(); self.group_combo.addItems(self.groups); self.group_combo.setCurrentText(self.current_group_name)
        self.group_combo.setFixedHeight(28)
        self.group_combo.setMinimumContentsLength(4)
        self.group_combo.setSizeAdjustPolicy(QComboBox.SizeAdjustPolicy.AdjustToMinimumContentsLengthWithIcon)
        self.group_combo.currentTextChanged.connect(self.on_group_change)

        self.group_placeholder = QWidget()
        self.group_placeholder.setFixedHeight(28)

        self.group_stack = QStackedWidget()
        self.group_stack.setFixedHeight(28)
        self.group_stack.addWidget(self.group_combo)
        self.group_stack.addWidget(self.group_placeholder)
        combo_hint = max(180, min(260, self.group_combo.sizeHint().width()))
        self.group_combo.setFixedWidth(combo_hint)
        self.group_placeholder.setFixedWidth(combo_hint)
        self.group_stack.setFixedWidth(combo_hint)
        self.group_combo.setSizePolicy(QSizePolicy.Policy.Fixed, QSizePolicy.Policy.Fixed)
        self.group_placeholder.setSizePolicy(QSizePolicy.Policy.Fixed, QSizePolicy.Policy.Fixed)
        self.group_stack.setSizePolicy(QSizePolicy.Policy.Fixed, QSizePolicy.Policy.Fixed)
        top.addWidget(self.group_stack, 0, Qt.AlignmentFlag.AlignLeft)
        self.group_combo.view().setMinimumWidth(combo_hint)

        self.reset_button = QPushButton("重置")
        self.reset_button.setCursor(Qt.CursorShape.PointingHandCursor)
        self.reset_button.setFixedHeight(28)
        self.reset_button.setSizePolicy(QSizePolicy.Policy.Fixed, QSizePolicy.Policy.Fixed)
        self.reset_button.clicked.connect(self.reset_roll_call_pools)
        top.addWidget(self.reset_button, 0, Qt.AlignmentFlag.AlignLeft)

        self.menu_button = QToolButton(); self.menu_button.setText("..."); self.menu_button.setPopupMode(QToolButton.ToolButtonPopupMode.InstantPopup)
        self.menu_button.setFixedSize(28, 28); self.menu_button.setStyleSheet("font-size: 18px; padding-bottom: 6px;")
        self.main_menu = self._build_menu(); self.menu_button.setMenu(self.main_menu)
        top.addStretch(1)
        top.addWidget(self.menu_button, 0, Qt.AlignmentFlag.AlignRight)
        layout.addLayout(top)

        self.stack = QStackedWidget(); layout.addWidget(self.stack, 1)

        self.roll_call_frame = ClickableFrame(); self.roll_call_frame.setFrameShape(QFrame.Shape.NoFrame)
        rl = QGridLayout(self.roll_call_frame); rl.setContentsMargins(6, 6, 6, 6); rl.setSpacing(6)
        self.id_label = QLabel(""); self.name_label = QLabel("")
        for lab in (self.id_label, self.name_label):
            lab.setAlignment(Qt.AlignmentFlag.AlignCenter)
            lab.setStyleSheet("color: #ffffff; background-color: #1a73e8; border-radius: 8px; padding: 8px;")
            lab.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        rl.addWidget(self.id_label, 0, 0); rl.addWidget(self.name_label, 0, 1); self.stack.addWidget(self.roll_call_frame)

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
        self.speech_enabled_action = speech.addAction("启用语音播报"); self.speech_enabled_action.setCheckable(True)
        self.speech_enabled_action.setChecked(self.speech_enabled); self.speech_enabled_action.toggled.connect(self._toggle_speech)
        self.voice_menu = speech.addMenu("选择发音人"); self.voice_actions = []
        if hasattr(self, "tts_manager") and self.tts_manager and self.tts_manager.voice_ids:
            for vid in self.tts_manager.voice_ids:
                act = self.voice_menu.addAction(vid); act.setCheckable(True); act.setChecked(vid == self.tts_manager.current_voice_id)
                act.triggered.connect(lambda _c, v=vid: self._set_voice(v)); self.voice_actions.append(act)
        else:
            self.voice_menu.setEnabled(False); self.speech_enabled_action.setEnabled(False)

        menu.addSeparator()
        self.timer_sound_action = menu.addAction("倒计时结束提示音"); self.timer_sound_action.setCheckable(True)
        self.timer_sound_action.setChecked(self.timer_sound_enabled); self.timer_sound_action.toggled.connect(self._toggle_timer_sound)
        return menu

    def _update_menu_state(self) -> None:
        if self.show_id_action.isChecked() != self.show_id: self.show_id_action.setChecked(self.show_id)
        if self.show_name_action.isChecked() != self.show_name: self.show_name_action.setChecked(self.show_name)
        self.timer_sound_action.setChecked(self.timer_sound_enabled)
        if self.tts_manager:
            self.speech_enabled_action.setEnabled(True); self.speech_enabled_action.setChecked(self.speech_enabled)
        else:
            self.speech_enabled_action.setEnabled(False); self.speech_enabled_action.setChecked(False)

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

    def _toggle_speech(self, enabled: bool) -> None:
        if not self.tts_manager or not self.tts_manager.available:
            show_quiet_information(self, "未检测到语音引擎，无法开启语音播报。")
            self.speech_enabled_action.setChecked(False)
            return
        self.speech_enabled = enabled

    def _set_voice(self, voice_id: str) -> None:
        if not self.tts_manager: return
        self.tts_manager.set_voice(voice_id); self.selected_voice_id = voice_id
        for a in self.voice_actions: a.setChecked(a.text() == voice_id)

    def _toggle_timer_sound(self, enabled: bool) -> None:
        self.timer_sound_enabled = enabled

    def toggle_mode(self) -> None:
        self.mode = "timer" if self.mode == "roll_call" else "roll_call"
        self.update_mode_ui()

    def update_mode_ui(self) -> None:
        is_roll = self.mode == "roll_call"
        self.title_label.setText("点名" if is_roll else "计时")
        self.mode_button.setText("切换到计时" if is_roll else "切换到点名")
        self.group_label.setVisible(is_roll)
        if is_roll:
            self.stack.setCurrentWidget(self.roll_call_frame); self.group_stack.setCurrentWidget(self.group_combo)
            self.count_timer.stop(); self.clock_timer.stop(); self.timer_running = False; self.timer_start_pause_button.setText("开始")
            self.update_display_layout(); self.display_current_student()
            self.schedule_font_update()
        else:
            self.stack.setCurrentWidget(self.timer_frame); self.group_stack.setCurrentWidget(self.group_placeholder); self.update_timer_mode_ui()
            self.schedule_font_update()
        self.reset_button.setVisible(is_roll)
        self.updateGeometry()

    def update_timer_mode_ui(self) -> None:
        mode = self.timer_modes[self.timer_mode_index]
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

    def start_pause_timer(self) -> None:
        if self.timer_modes[self.timer_mode_index] == "clock": return
        self.timer_running = not self.timer_running
        if self.timer_running:
            self.timer_start_pause_button.setText("暂停")
            if not self.count_timer.isActive(): self.count_timer.start()
        else:
            self.timer_start_pause_button.setText("开始")
            self.count_timer.stop()

    def reset_timer(self) -> None:
        self.timer_running = False; self.count_timer.stop(); self.timer_start_pause_button.setText("开始")
        m = self.timer_modes[self.timer_mode_index]
        if m == "countdown":
            self.timer_seconds_left = max(0, self.timer_countdown_minutes * 60 + self.timer_countdown_seconds)
        elif m == "stopwatch":
            self.timer_stopwatch_seconds = 0
        self.update_timer_display()

    def set_countdown_time(self) -> None:
        d = CountdownSettingsDialog(self, self.timer_countdown_minutes, self.timer_countdown_seconds)
        if d.exec() and d.result:
            mi, se = d.result; self.timer_countdown_minutes = mi; self.timer_countdown_seconds = se; self.reset_timer()

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
        if group_name is None:
            group_name = self.group_combo.currentText()
        if group_name not in self.groups:
            group_name = "全部"
            if self.group_combo.currentText() != group_name:
                self.group_combo.setCurrentText(group_name)
        self.current_group_name = group_name
        if self.student_data.empty:
            self.current_student_index = None
            self.display_current_student()
            return
        self._ensure_group_pool(group_name)
        self.current_student_index = None
        self.display_current_student()

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
        self._group_last_student[group_name] = self.current_student_index
        self._mark_student_drawn(self.current_student_index)
        self.display_current_student()
        if speak and self.speech_enabled and self.tts_manager and self.tts_manager.available:
            stu = self.student_data.loc[self.current_student_index]
            name = str(stu["姓名"]) if "姓名" in stu and pd.notna(stu["姓名"]) else ""
            if name: self.tts_manager.speak(name)

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

        self._rebuild_group_indices()
        self._ensure_group_pool(self.current_group_name)

    def reset_roll_call_pools(self) -> None:
        """根据当前分组执行重置：子分组独立重置，“全部”重置所有。"""

        group_name = self.current_group_name
        if group_name == "全部":
            self._reset_roll_call_state()
        else:
            self._reset_single_group(group_name)
        self.current_student_index = None
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
        random.shuffle(shuffled)
        self._group_remaining_indices[group_name] = shuffled
        self._group_initial_sequences[group_name] = list(shuffled)
        self._group_last_student[group_name] = None

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
            random.shuffle(pool)
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
            random.shuffle(shuffled)
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
                random.shuffle(shuffled)
                self._group_initial_sequences[group_name] = shuffled
            self._refresh_all_group_pool()
            self._group_last_student.setdefault(group_name, None)
            return

        if force_reset or group_name not in self._group_remaining_indices:
            drawn_history.clear()
            pool = list(base_indices)
            random.shuffle(pool)
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
            random.shuffle(shuffled)
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
            sid = str(stu["学号"]) if pd.notna(stu["学号"]) else ""; name = str(stu["姓名"]) if pd.notna(stu["姓名"]) else ""
            self.id_label.setText(sid if self.show_id else ""); self.name_label.setText(name if self.show_name else "")
            if not self.show_id: self.id_label.setText("")
            if not self.show_name: self.name_label.setText("")
        self.update_display_layout()
        self.schedule_font_update()

    def update_display_layout(self) -> None:
        self.id_label.setVisible(self.show_id); self.name_label.setVisible(self.show_name)
        layout: QGridLayout = self.roll_call_frame.layout()
        layout.setColumnStretch(0, 1); layout.setColumnStretch(1, 1)
        if not self.show_id: layout.setColumnStretch(0, 0)
        if not self.show_name: layout.setColumnStretch(1, 0)
        self.schedule_font_update()

    def schedule_font_update(self) -> None:
        QTimer.singleShot(0, self.update_dynamic_fonts)

    def update_dynamic_fonts(self) -> None:
        for lab in (self.id_label, self.name_label):
            if not lab.isVisible(): continue
            w = max(40, lab.width()); h = max(40, lab.height()); text = lab.text()
            size = self._calc_font_size(w, h, text)
            if lab is self.name_label:
                weight = QFont.Weight.Normal if self.name_font_family in {"楷体", "KaiTi"} else QFont.Weight.Bold
                lab.setFont(QFont(self.name_font_family, size, weight))
                self.last_name_font_size = size
            else:
                lab.setFont(QFont("Microsoft YaHei UI", size, QFont.Weight.Bold))
                self.last_id_font_size = size
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
        self.visibility_changed.emit(True)
        self.schedule_font_update()
        ensure_widget_within_screen(self)

    def resizeEvent(self, e: QResizeEvent) -> None:
        super().resizeEvent(e)
        self.schedule_font_update()

    def hideEvent(self, e) -> None:
        super().hideEvent(e)
        self.save_settings()
        self.visibility_changed.emit(False)

    def closeEvent(self, e) -> None:
        self.save_settings()
        self.count_timer.stop()
        self.clock_timer.stop()
        if self.tts_manager: self.tts_manager.shutdown()
        self.window_closed.emit()
        super().closeEvent(e)

    def save_settings(self) -> None:
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


# ---------- 数据 ----------
def load_student_data(parent: Optional[QWidget]) -> Optional[pd.DataFrame]:
    """从 students.xlsx 读取点名所需的数据，不存在时自动生成模板。"""
    if not (PANDAS_AVAILABLE and OPENPYXL_AVAILABLE):
        QMessageBox.warning(parent, "提示", "未安装 pandas/openpyxl，点名功能不可用。")
        return None
    file_path = RollCallTimerWindow.STUDENT_FILE
    if not os.path.exists(file_path):
        try:
            df = pd.DataFrame({"学号": [101, 102, 103], "姓名": ["张三", "李四", "王五"], "分组": ["A", "B", "A"]})
            df.to_excel(file_path, index=False)
            show_quiet_information(parent, f"未找到学生名单，已为您创建模板文件：{file_path}")
        except Exception as exc:
            QMessageBox.critical(parent, "错误", f"创建模板文件失败：{exc}")
            return None
    try:
        df = pd.read_excel(file_path, usecols=["学号", "姓名", "分组"])
        df["学号"] = pd.to_numeric(df["学号"], errors="coerce").astype("Int64")
        df["姓名"] = df["姓名"].astype(str).str.strip()
        df["分组"] = df["分组"].astype(str).str.strip().str.upper()
        df.dropna(subset=["学号", "姓名"], inplace=True)
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
        self.exit_button.clicked.connect(QApplication.instance().quit)
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
        # 启动器关闭视为重新开课，清理点名历史以便下次全新开始
        self.settings_manager.clear_roll_call_history()
        super().closeEvent(e)


# ---------- 入口 ----------
def main() -> None:
    """应用程序入口：初始化 DPI、加载设置并启动启动器窗口。"""
    ensure_high_dpi_awareness()
    app = QApplication(sys.argv)
    app.setQuitOnLastWindowClosed(False)
    QToolTip.setFont(QFont("Microsoft YaHei UI", 9))

    settings_manager = SettingsManager()
    # 每次启动器运行时先清空上一轮的点名记录，确保不会延续上一课堂的名单
    settings_manager.clear_roll_call_history()
    student_data = load_student_data(None) if PANDAS_AVAILABLE else None

    window = LauncherWindow(settings_manager, student_data)
    window.show()
    sys.exit(app.exec())


# Nuitka 打包指令（根据当前依赖整理的推荐参数，保持在单行便于复制）：
# 单文件：python -m nuitka --onefile --enable-plugin=pyqt6 --include-qt-plugins=sensible --windows-disable-console --windows-icon-from-ico=icon.ico --include-data-file=students.xlsx=students.xlsx --include-data-file=settings.ini=settings.ini ClassroomTools.py
# 独立目录：python -m nuitka --standalone --enable-plugin=pyqt6 --include-qt-plugins=sensible --windows-disable-console --windows-icon-from-ico=icon.ico --output-dir=dist --include-data-file=students.xlsx=students.xlsx --include-data-file=settings.ini=settings.ini ClassroomTools.py
if __name__ == "__main__":
    main()
