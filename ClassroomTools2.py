# -*- coding: utf-8 -*-
from __future__ import annotations

import base64
import configparser
import ctypes
import os
import random
import sys
import threading
import time
from queue import Empty, Queue
from typing import Dict, List, Optional

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
    QIcon,
    QPainter,
    QPainterPath,
    QPen,
    QPixmap,
    QKeyEvent,
    QResizeEvent,
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
    QSizePolicy,
    QSlider,
    QSpinBox,
    QStackedWidget,
    QToolButton,
    QVBoxLayout,
    QWidget,
)

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
    # [FIX] 必须在 QApplication 之前设置，否则会出现 setHighDpiScaleFactorRoundingPolicy 报错
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
    rect = widget.frameGeometry()
    return f"{rect.width()}x{rect.height()}+{rect.x()}+{rect.y()}"


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
    widget.resize(max(160, width), max(120, height))
    widget.move(x, y)


def str_to_bool(value: str, default: bool = False) -> bool:
    if isinstance(value, str):
        return value.strip().lower() in {"1", "true", "yes", "on"}
    return default


def bool_to_str(value: bool) -> str:
    return "True" if value else "False"


# ---------- 配置 ----------
class SettingsManager:
    def __init__(self, filename: str = "settings.ini") -> None:
        self.filename = filename
        self.config = configparser.ConfigParser()
        self.defaults: Dict[str, Dict[str, str]] = {
            "Launcher": {"x": "120", "y": "120"},
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


# ---------- 图标 ----------
class IconManager:
    _icons: Dict[str, str] = {
        "cursor": "PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHdpZHRoPSIyNCIgaGVpZHRoPSIyNCIgdmlld0JveD0iMCAwIDI0IDI0IiBmaWxsPSJub25lIiBzdHJva2U9ImN1cnJlbnRDb2xvciIgc3Ryb2tlLXdpZHRoPSIyIiBzdHJva2UtbGluZWNhcD0icm91bmQiIHN0cm9rZS1saW5lam9pbj0icm91bmQiPjxwYXRoIGQ9Ik0zIDNsNy4wNyAxNy4wNyAyLjUxLTcuMzkgNy4zOS0yLjUxTDMgM3oiLz48cGF0aCBkPSJNMTMgMTNsNiA2Ii8+PC9zdmc+",
        "brush": "PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHdpZHRoPSIyNCIgaGVpZHRoPSIyNCIgdmlld0JveD0iMCAwIDI0IDI0IiBmaWxsPSJub25lIiBzdHJva2U9ImN1cnJlbnRDb2xvciIgc3Ryb2tlLXdpZHRoPSIyIiBzdHJva2UtbGluZWNhcD0icm91bmQiIHN0cm9rZS1saW5lam9pbj0icm91bmQiPjxwYXRoIGQ9Ik0xNyAzYTIuODI4IDIuODI4IDAgMSAxIDQgNEw3LjUgMjAuNSAyIDIybDEuNS01LjVMMTcgM3oiPjwvcGF0aD48L3N2Zz4=",
        "shape": "PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHdpZHRoPSIyNCIgaGVpZHRoPSIyNCIgdmlld0JveD0iMCAwIDI0IDI0IiBmaWxsPSJub25lIiBzdHJva2U9ImN1cnJlbnRDb2xvciIgc3Ryb2tlLXdpZHRoPSIyIiBzdHJva2UtbGluZWNhcD0icm91bmQiIHN0cm9rZS1saW5lam9pbj0icm91bmQiPjxjaXJjbGUgY3g9IjE1IiBjeT0iMTUiIHI9IjciLz48cmVjdCB4PSIyIiB5PSIyIiB3aWR0aD0iMTIiIGhlaWdodD0iMTIiIHJ4PSIyIi8+PC9zdmc+",
        "eraser": "PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHdpZHRoPSIyNCIgaGVpZHRoPSIyNCIgdmlld0JveD0iMCAwIDI0IDI0IiBmaWxsPSJub25lIiBzdHJva2U9ImN1cnJlbnRDb2xvciIgc3Ryb2tlLXdpZHRoPSIyIiBzdHJva2UtbGluZWNhcD0icm91bmQiIHN0cm9rZS1saW5lam9pbj0icm91bmQiPjxwYXRoIGQ9Im03IDIxLTQuMy00LjNjLTEtMS0xLTIuNSAwLTMuNGw5LjYtOS42YzEtMSAyLjUtMSAzLjQgMGw1LjYgNS42YzEgMSAxIDIuNSAwIDMuNEwxMyAyMUg3WiIvPjxwYXRoIGQ9Ik0yMiAyMUg3Ii8+PHBhdGggZD0ibTUgMTItNSA1Ii8+PC9zdmc+",
        "clear": "PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHdpZHRoPSIyNCIgaGVpZHRoPSIyNCIgdmlld0JveD0iMCAwIDI0IDI0IiBmaWxsPSJub25lIiBzdHJva2U9ImN1cnJlbnRDb2xvciIgc3Ryb2tlLXdpZHRoPSIyIiBzdHJva2UtbGluZWNhcD0icm91bmQiIHN0cm9rZS1saW5lam9pbj0icm91bmQiPjxwb2x5bGluZSBwb2ludHM9IjMgNiA1IDYgMjEgNiI+PC9wb2x5bGluZT48cGF0aCBkPSJNMTkgNnYxNGEyIDIgMCAwIDEtMiAySDdhMiAyIDAgMCAxLTItMlY2bTMgMFY0YTIgMiAwIDAgMSAyLTJoNGEyIDIgMCAwIDEgMiAydjIiPjwvcGF0aD48bGluZSB4MT0iMTAiIHkxPSIxMSIgeDI9IjEwIiB5Mj0iMTciPjwvbGluZT48bGluZSB4MT0iMTQiIHkxPSIxMSIgeDI9IjE0IiB5Mj0iMTciPjwvbGluZT48L3N2Zz4=",
        "settings": "PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHdpZHRoPSIyNCIgaGVpZHRoPSIyNCIgdmlld0JveD0iMCAwIDI0IDI0IiBmaWxsPSJub25lIiBzdHJva2U9ImN1cnJlbnRDb2xvciIgc3Ryb2tlLXdpZHRoPSIyIiBzdHJva2UtbGluZWNhcD0icm91bmQiIHN0cm9rZS1saW5lam9pbj0icm91bmQiPjxsaW5lIHgxPSI0IiB5MT0iMjEiIHgyPSI0IiB5Mj0iMTQiPjwvbGluZT48bGluZSB4MT0iNCIgeTE9IjEwIiB4Mj0iNCIgeTI9IjMiPjwvbGluZT48bGluZSB4MT0iMTIiIHkxPSIyMSIgeDI9IjEyIiB5Mj0iMTIiPjwvbGluZT48bGluZSB4MT0iMTIiIHkxPSI4IiB4Mj0iMTIiIHkyPSIzIj48L2xpbmU+PGxpbmUgeDE9IjIwIiB5MT0iMjEiIHgyPSIyMCIgeTI9IjE2Ij48L2xpbmU+PGxpbmUgeDE9IjIwIiB5MT0iMTIiIHgyPSIyMCIgeTI9IjMiPjwvbGluZT48bGluZSB4MT0iMSIgeTE9IjE4IiB4Mj0iNyIgeTI9IjE4Ij48L2xpbmU+PGxpbmUgeDE9IjkiIHkxPSI4IiB4Mj0iMTUiIHkyPSI4Ij48L2xpbmU+PGxpbmUgeDE9IjE3IiB5MT0iMTYiIHgyPSIyMyIgeTI9IjE2Ij48L2xpbmU+PC9zdmc+",
        "whiteboard": "PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHdpZHRoPSIyNCIgaGVpZHRoPSIyNCIgdmlld0JveD0iMCAwIDI0IDI0IiBmaWxsPSJub25lIiBzdHJva2U9ImN1cnJlbnRDb2xvciIgc3Ryb2tlLXdpZHRoPSIyIiBzdHJva2UtbGluZWNhcD0icm91bmQiIHN0cm9rZS1saW5lam9pbj0icm91bmQiPjxyZWN0IHg9IjMiIHk9IjMiIHdpZHRoPSIxOCIgaGVpZGh0PSIxMiIgcng9IjIiLz48cGF0aCBkPSJNMTIgMTV2NiIvPjxwYXRoIGQ9Ik04IDIxaDgiLz48L3N2Zz4=",
    }
    _cache: Dict[str, QIcon] = {}

    @classmethod
    def get_icon(cls, name: str) -> QIcon:
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


# ---------- 对话框 ----------
class PenSettingsDialog(QDialog):
    COLORS = {
        "#000000": "黑",
        "#FF0000": "红",
        "#1E90FF": "蓝",
        "#FFFF00": "黄",
        "#800080": "紫",
        "#FFFFFF": "白",
    }

    def __init__(self, parent: Optional[QWidget] = None, initial_size: int = 12, initial_color: str = "#FF0000") -> None:
        super().__init__(parent)
        self.setWindowTitle("画笔设置")
        self.setWindowFlag(Qt.WindowType.WindowStaysOnTopHint, True)
        self.pen_color = QColor(initial_color)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(8)

        size_layout = QHBoxLayout()
        size_label = QLabel("粗细:")
        self.size_slider = QSlider(Qt.Orientation.Horizontal)
        self.size_slider.setRange(2, 40)
        self.size_slider.setValue(int(initial_size))
        self.size_value = QLabel(str(initial_size))
        self.size_slider.valueChanged.connect(lambda value: self.size_value.setText(str(value)))

        size_layout.addWidget(size_label)
        size_layout.addWidget(self.size_slider, 1)
        size_layout.addWidget(self.size_value)
        layout.addLayout(size_layout)

        layout.addWidget(QLabel("颜色:"))
        color_layout = QGridLayout()
        color_layout.setSpacing(6)
        row, col = 0, 0
        for color_hex, name in self.COLORS.items():
            button = QPushButton()
            button.setFixedSize(28, 28)
            button.setStyleSheet(f"background-color: {color_hex}; border: 1px solid #888; border-radius: 4px;")
            button.setToolTip(name)
            button.clicked.connect(lambda _checked, c=color_hex: self._select_color(c))
            color_layout.addWidget(button, row, col)
            col += 1
            if col > 3:
                row += 1
                col = 0
        layout.addLayout(color_layout)

        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def _select_color(self, color_hex: str) -> None:
        self.pen_color = QColor(color_hex)

    def get_settings(self) -> tuple[int, QColor]:
        return self.size_slider.value(), self.pen_color


class ShapeSettingsDialog(QDialog):
    SHAPES = {"line": "直线", "dashed_line": "虚线", "rect": "矩形", "circle": "圆形"}

    def __init__(self, parent: Optional[QWidget] = None) -> None:
        super().__init__(parent)
        self.setWindowTitle("选择图形")
        self.setWindowFlag(Qt.WindowType.WindowStaysOnTopHint, True)
        self.selected_shape: Optional[str] = None

        layout = QGridLayout(self)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(8)

        for index, (shape_key, name) in enumerate(self.SHAPES.items()):
            button = QPushButton(name)
            button.setFixedWidth(72)
            button.clicked.connect(lambda _checked, key=shape_key: self._select_shape(key))
            layout.addWidget(button, index // 2, index % 2)

    def _select_shape(self, shape: str) -> None:
        self.selected_shape = shape
        self.accept()

    def get_shape(self) -> Optional[str]:
        return self.selected_shape


class BoardColorDialog(QDialog):
    COLORS = {"#FFFFFF": "白板", "#000000": "黑板", "#0E4020": "深绿"}

    def __init__(self, parent: Optional[QWidget] = None) -> None:
        super().__init__(parent)
        self.setWindowTitle("选择底色")
        self.setWindowFlag(Qt.WindowType.WindowStaysOnTopHint, True)
        self.selected_color: Optional[QColor] = None

        layout = QHBoxLayout(self)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(8)

        for color_hex, name in self.COLORS.items():
            button = QPushButton(name)
            button.clicked.connect(lambda _checked, c=color_hex: self._select_color(c))
            layout.addWidget(button)

    def _select_color(self, color_hex: str) -> None:
        self.selected_color = QColor(color_hex)
        self.accept()

    def get_color(self) -> Optional[QColor]:
        return self.selected_color


# ---------- 工具条 ----------
class TitleBar(QWidget):
    def __init__(self, toolbar: "FloatingToolbar") -> None:
        super().__init__(toolbar)
        self.toolbar = toolbar
        self._dragging = False
        self._drag_offset = QPoint()

        self.setAutoFillBackground(True)
        p = self.palette()
        p.setColor(self.backgroundRole(), QColor(230, 230, 230, 240))
        self.setPalette(p)
        self.setFixedHeight(26)

        layout = QHBoxLayout(self)
        layout.setContentsMargins(10, 0, 6, 0)
        title = QLabel("屏幕画笔")
        f = title.font()
        f.setBold(True)
        title.setFont(f)
        title.setStyleSheet("color: #333333;")
        layout.addWidget(title)
        layout.addStretch(1)

    def mousePressEvent(self, event) -> None:
        if event.button() == Qt.MouseButton.LeftButton:
            self._dragging = True
            self._drag_offset = event.globalPosition().toPoint() - self.toolbar.pos()
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
        event.accept()
        super().mouseReleaseEvent(event)


class FloatingToolbar(QWidget):
    def __init__(self, overlay: "OverlayWindow", settings_manager: SettingsManager) -> None:
        super().__init__(None, Qt.WindowType.Tool | Qt.WindowType.FramelessWindowHint | Qt.WindowType.WindowStaysOnTopHint)
        self.overlay = overlay
        self.settings_manager = settings_manager
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        self._build_ui()

        s = self.settings_manager.load_settings().get("Paint", {})
        self.move(int(s.get("x", "260")), int(s.get("y", "260")))

    def _build_ui(self) -> None:
        self.setStyleSheet(
            """
            #container { background-color: rgba(245, 245, 245, 240); border-radius: 10px; border: 1px solid rgba(0,0,0,60); }
            QPushButton { background: rgba(255, 255, 255, 220); border: 1px solid rgba(0,0,0,60); border-radius: 6px; padding: 4px; min-width: 34px; min-height: 34px; }
            QPushButton:hover { background: rgba(255,255,255,255); }
            QPushButton:checked { background: rgba(0, 120, 215, 70); border-color: rgba(0,120,215,160); }
            #whiteboardButtonActive { background: rgba(0, 120, 215, 70); border-color: rgba(0,120,215,160); }
            """
        )
        root = QVBoxLayout(self); root.setContentsMargins(0, 0, 0, 0)
        container = QWidget(self); container.setObjectName("container"); root.addWidget(container)

        v = QVBoxLayout(container); v.setContentsMargins(8, 8, 8, 8); v.setSpacing(6)
        self.title_bar = TitleBar(self); v.addWidget(self.title_bar)

        h = QHBoxLayout(); h.setContentsMargins(2, 2, 2, 2); h.setSpacing(6)
        self.btn_cursor = QPushButton(IconManager.get_icon("cursor"), "")
        self.btn_brush = QPushButton(IconManager.get_icon("brush"), "")
        self.btn_shape = QPushButton(IconManager.get_icon("shape"), "")
        self.btn_eraser = QPushButton(IconManager.get_icon("eraser"), "")
        self.btn_clear_all = QPushButton(IconManager.get_icon("clear"), "")
        self.btn_whiteboard = QPushButton(IconManager.get_icon("whiteboard"), "")
        self.btn_settings = QPushButton(IconManager.get_icon("settings"), "")
        for b in (self.btn_cursor, self.btn_brush, self.btn_shape, self.btn_eraser, self.btn_clear_all, self.btn_whiteboard, self.btn_settings):
            b.setIconSize(QSize(22, 22)); h.addWidget(b)
        v.addLayout(h)

        tooltip = {
            self.btn_cursor: "指针",
            self.btn_brush: "画笔",
            self.btn_shape: "图形",
            self.btn_eraser: "橡皮",
            self.btn_clear_all: "清除所有",
            self.btn_whiteboard: "白板 (单击开关 / 双击选色)",
            self.btn_settings: "画笔设置",
        }
        for b, t in tooltip.items(): b.setToolTip(t)

        self.tool_buttons = QButtonGroup(self)
        for b in (self.btn_cursor, self.btn_brush, self.btn_shape, self.btn_eraser):
            b.setCheckable(True); self.tool_buttons.addButton(b)

        self.btn_cursor.clicked.connect(lambda: self.overlay.set_mode("cursor"))
        self.btn_brush.clicked.connect(lambda: self.overlay.set_mode("brush"))
        self.btn_shape.clicked.connect(self._select_shape)
        self.btn_eraser.clicked.connect(lambda: self.overlay.set_mode("eraser"))
        self.btn_clear_all.clicked.connect(self.overlay.clear_all)
        self.btn_settings.clicked.connect(self.overlay.open_pen_settings)
        self.btn_whiteboard.clicked.connect(self._handle_whiteboard_click)

        self._wb_click_timer = QTimer(self)
        self._wb_click_timer.setInterval(QApplication.instance().doubleClickInterval())
        self._wb_click_timer.setSingleShot(True)
        self._wb_click_timer.timeout.connect(self.overlay.toggle_whiteboard)

    def _select_shape(self) -> None:
        d = ShapeSettingsDialog(self)
        if d.exec():
            s = d.get_shape()
            if s: self.overlay.set_mode("shape", shape_type=s)
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

    def enterEvent(self, e) -> None:
        self.raise_()
        super().enterEvent(e)


# ---------- 叠加层（画笔/白板） ----------
class OverlayWindow(QWidget):
    def __init__(self, settings_manager: SettingsManager) -> None:
        super().__init__(None, Qt.WindowType.FramelessWindowHint | Qt.WindowType.WindowStaysOnTopHint)
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        self.settings_manager = settings_manager
        s = self.settings_manager.load_settings().get("Paint", {})
        self.pen_size = int(s.get("brush_size", "12"))
        self.pen_color = QColor(s.get("brush_color", "#ff0000"))
        self.mode = "cursor"
        self.current_shape: Optional[str] = None
        self.shape_start_point: Optional[QPoint] = None
        self.drawing = False
        self.last_point = QPointF(); self.prev_point = QPointF()
        self.last_width = float(self.pen_size); self.last_time = time.time()
        self.whiteboard_active = False
        self.whiteboard_color = QColor(0, 0, 0, 0); self.last_board_color = QColor("#ffffff")
        self.cursor_pixmap = QPixmap()
        self.setFocusPolicy(Qt.FocusPolicy.StrongFocus)
        self.setMouseTracking(True)

        self._build_scene()
        self.toolbar = FloatingToolbar(self, self.settings_manager)
        self.set_mode("brush", initial=True)

    def _build_scene(self) -> None:
        virtual = QRect()
        for screen in QApplication.screens():
            virtual = virtual.united(screen.geometry())
        self.setGeometry(virtual)
        self.canvas = QPixmap(self.size()); self.canvas.fill(Qt.GlobalColor.transparent)
        self.temp_canvas = QPixmap(self.size()); self.temp_canvas.fill(Qt.GlobalColor.transparent)

    def show_overlay(self) -> None:
        self.show(); self.toolbar.show(); self.toolbar.raise_()
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
                # [FIX] 确保工具条在顶层可见
                self.toolbar.show()
                self.toolbar.raise_()
                self.update()

    def toggle_whiteboard(self) -> None:
        self.whiteboard_active = not self.whiteboard_active
        self.whiteboard_color = self.last_board_color if self.whiteboard_active else QColor(0, 0, 0, 0)
        self._update_visibility_for_mode(initial=False)
        self.toolbar.show(); self.toolbar.raise_()
        self.toolbar.update_whiteboard_button_state(self.whiteboard_active)
        self.update()

    def set_mode(self, mode: str, shape_type: Optional[str] = None, *, initial: bool = False) -> None:
        self.mode = mode
        if shape_type is not None or mode != "shape":
            self.current_shape = shape_type
        if mode != "shape":
            self.shape_start_point = None
        self._update_visibility_for_mode(initial=initial)
        if not initial:
            self.toolbar.show(); self.toolbar.raise_()
        self.update_toolbar_state()
        self.update_cursor()

    def update_toolbar_state(self) -> None:
        if not getattr(self, 'toolbar', None):
            return
        m = {"cursor": self.toolbar.btn_cursor, "brush": self.toolbar.btn_brush, "shape": self.toolbar.btn_shape, "eraser": self.toolbar.btn_eraser}
        for k, b in m.items():
            prev = b.blockSignals(True); b.setChecked(self.mode == k); b.blockSignals(prev)

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
        self.setAttribute(Qt.WidgetAttribute.WA_TransparentForMouseEvents, enabled)
        self.setWindowFlag(Qt.WindowType.WindowTransparentForInput, enabled)  # [FIX] OS 层穿透
        if self.isVisible():
            super().show()  # 让 flag 立即生效

    def _update_visibility_for_mode(self, *, initial: bool = False) -> None:
        passthrough = (self.mode == "cursor") and (not self.whiteboard_active)
        self._apply_input_passthrough(passthrough)
        if initial:
            return
        if not passthrough:
            self.show(); self.raise_(); self.activateWindow()
            self.setFocus(Qt.FocusReason.ActiveWindowFocusReason)
        else:
            if not self.isVisible():
                self.show()

    def clear_all(self) -> None:
        self.canvas.fill(Qt.GlobalColor.transparent)
        self.temp_canvas.fill(Qt.GlobalColor.transparent)
        self.update()

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
            self.drawing = True
            pointf = e.position(); self.last_point = pointf; self.prev_point = pointf
            self.last_width = self.pen_size * 0.4
            self.shape_start_point = e.pos() if self.mode == "shape" else None
            e.accept()
        super().mousePressEvent(e)

    def mouseMoveEvent(self, e) -> None:
        if self.drawing and self.mode != "cursor":
            p = e.pos(); pf = e.position()
            if self.mode == "brush": self._draw_brush_line(pf)
            elif self.mode == "eraser": self._erase_at(p)
            elif self.mode == "shape" and self.current_shape: self._draw_shape_preview(p)
            self.update()
        super().mouseMoveEvent(e)

    def mouseReleaseEvent(self, e) -> None:
        if e.button() == Qt.MouseButton.LeftButton and self.drawing:
            if self.mode == "shape" and self.current_shape: self._draw_shape_final(e.pos())
            self.drawing = False; self.shape_start_point = None; self.update()
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
        if isinstance(pos, QPointF): pos = pos.toPoint()
        p = QPainter(self.canvas); p.setRenderHint(QPainter.RenderHint.Antialiasing)
        p.setCompositionMode(QPainter.CompositionMode.CompositionMode_Clear)
        radius = max(6, int(self.pen_size * 3.2))
        p.setPen(QPen(QColor(255, 255, 255, 0), radius, Qt.PenStyle.SolidLine, Qt.PenCapStyle.RoundCap, Qt.PenJoinStyle.RoundJoin))
        p.drawPoint(pos); p.end()

    def _draw_shape_preview(self, end_point) -> None:
        if not self.shape_start_point: return
        self.temp_canvas.fill(Qt.GlobalColor.transparent)
        p = QPainter(self.temp_canvas); p.setRenderHint(QPainter.RenderHint.Antialiasing)
        pen = QPen(self.pen_color, self.pen_size)
        if self.current_shape and "dashed" in self.current_shape: pen.setStyle(Qt.PenStyle.DashLine)
        p.setPen(pen); self._draw_shape(p, self.shape_start_point, end_point); p.end()

    def _draw_shape_final(self, end_point) -> None:
        if not self.shape_start_point: return
        p = QPainter(self.canvas); p.setRenderHint(QPainter.RenderHint.Antialiasing)
        pen = QPen(self.pen_color, self.pen_size)
        if self.current_shape and "dashed" in self.current_shape: pen.setStyle(Qt.PenStyle.DashLine)
        p.setPen(pen); self._draw_shape(p, self.shape_start_point, end_point); p.end()
        self.temp_canvas.fill(Qt.GlobalColor.transparent)

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

    def closeEvent(self, e) -> None:
        self.save_settings(); self.save_window_position()
        super().closeEvent(e)


# ---------- 语音 ----------
class TTSManager(QObject):
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
        if voice_id in self.voice_ids: self.current_voice_id = voice_id

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
    def __init__(self, parent: Optional[QWidget], minutes: int, seconds: int) -> None:
        super().__init__(parent)
        self.setWindowTitle("设置倒计时")
        self.setWindowFlag(Qt.WindowType.WindowStaysOnTopHint, True)
        self.result: Optional[tuple[int, int]] = None

        layout = QVBoxLayout(self); layout.setContentsMargins(12, 12, 12, 12); layout.setSpacing(10)

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

    def _accept(self) -> None:
        self.result = (self.minutes_spin.value(), self.seconds_spin.value()); self.accept()


class ClickableFrame(QFrame):
    clicked = pyqtSignal()
    def mousePressEvent(self, e) -> None:
        if e.button() == Qt.MouseButton.LeftButton: self.clicked.emit()
        super().mousePressEvent(e)


class RollCallTimerWindow(QWidget):
    window_closed = pyqtSignal()
    visibility_changed = pyqtSignal(bool)

    STUDENT_FILE = "students.xlsx"
    MIN_FONT_SIZE = 9; MAX_FONT_SIZE = 220

    def __init__(self, settings_manager: SettingsManager, student_data, parent: Optional[QWidget] = None) -> None:
        super().__init__(parent)
        self.setWindowTitle("点名 / 计时")
        self.setWindowFlag(Qt.WindowType.WindowStaysOnTopHint, True)
        self.setAttribute(Qt.WidgetAttribute.WA_DeleteOnClose)  # 只在真正 close() 时删除
        self.settings_manager = settings_manager
        self.student_data = student_data

        s = self.settings_manager.load_settings().get("RollCallTimer", {})
        apply_geometry_from_text(self, s.get("geometry", "420x240+180+180"))
        self.setMinimumSize(260, 160)

        self.mode = s.get("mode", "roll_call") if s.get("mode", "roll_call") in {"roll_call", "timer"} else "roll_call"
        self.timer_modes = ["countdown", "stopwatch", "clock"]
        self.timer_mode_index = self.timer_modes.index(s.get("timer_mode", "countdown")) if s.get("timer_mode", "countdown") in self.timer_modes else 0

        self.timer_countdown_minutes = int(s.get("timer_countdown_minutes", "5"))
        self.timer_countdown_seconds = int(s.get("timer_countdown_seconds", "0"))
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

        self._shuffled_indices: List[int] = []; self.current_student_index: Optional[int] = None
        self.timer_seconds_left = max(0, self.timer_countdown_minutes * 60 + self.timer_countdown_seconds)
        self.timer_stopwatch_seconds = 0; self.timer_running = False

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

        self._build_ui()
        self._update_menu_state()
        self.update_mode_ui()
        self.on_group_change(initial=True)
        self.display_current_student()

    def _build_ui(self) -> None:
        self.setStyleSheet("background-color: #f4f5f7;")
        layout = QVBoxLayout(self); layout.setContentsMargins(10, 10, 10, 10); layout.setSpacing(8)

        top = QHBoxLayout(); top.setSpacing(6)
        self.title_label = QLabel("点名"); f = QFont("Microsoft YaHei UI", 10, QFont.Weight.Bold)
        self.title_label.setFont(f); self.title_label.setStyleSheet("color: #202124;"); top.addWidget(self.title_label)

        self.mode_button = QPushButton("切换到计时"); self.mode_button.setCursor(Qt.CursorShape.PointingHandCursor)
        self.mode_button.clicked.connect(self.toggle_mode); self.mode_button.setFixedHeight(28); top.addWidget(self.mode_button)

        self.group_combo = QComboBox(); self.group_combo.addItems(self.groups); self.group_combo.setCurrentText(self.current_group_name)
        self.group_combo.currentTextChanged.connect(self.on_group_change); self.group_combo.setFixedHeight(28); top.addWidget(self.group_combo, 1)

        self.menu_button = QToolButton(); self.menu_button.setText("..."); self.menu_button.setPopupMode(QToolButton.ToolButtonPopupMode.InstantPopup)
        self.menu_button.setFixedSize(28, 28); self.menu_button.setStyleSheet("font-size: 18px; padding-bottom: 6px;")
        self.main_menu = self._build_menu(); self.menu_button.setMenu(self.main_menu); top.addWidget(self.menu_button)
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
            b.setCursor(Qt.CursorShape.PointingHandCursor); b.setFixedHeight(32); b.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed); ctrl.addWidget(b)
        tl.addLayout(ctrl); self.stack.addWidget(self.timer_frame)

        self.roll_call_frame.clicked.connect(self.roll_student)
        self.id_label.installEventFilter(self); self.name_label.installEventFilter(self)

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
        # [FIX] 修复当两者都取消勾选时，无法保持至少一个勾选的逻辑
        if not self.show_id_action.isChecked() and not self.show_name_action.isChecked():
            # 强制保持学号为勾选状态，避免两个都取消
            self.show_id_action.setChecked(True)
        
        self.show_id = self.show_id_action.isChecked()
        self.show_name = self.show_name_action.isChecked()
        self.update_display_layout()
        self.display_current_student()

    def _toggle_speech(self, enabled: bool) -> None:
        if not self.tts_manager or not self.tts_manager.available:
            QMessageBox.information(self, "提示", "未检测到语音引擎，无法开启语音播报。"); self.speech_enabled_action.setChecked(False); return
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
        if is_roll:
            self.stack.setCurrentWidget(self.roll_call_frame); self.group_combo.show()
            self.count_timer.stop(); self.clock_timer.stop(); self.timer_running = False; self.timer_start_pause_button.setText("开始")
            self.update_display_layout(); self.display_current_student()
        else:
            self.stack.setCurrentWidget(self.timer_frame); self.group_combo.hide(); self.update_timer_mode_ui()
        self.updateGeometry()

    def update_timer_mode_ui(self) -> None:
        mode = self.timer_modes[self.timer_mode_index]
        self.clock_timer.stop()
        if mode == "countdown":
            self.timer_mode_button.setText("倒计时")
            self.timer_start_pause_button.setEnabled(True); self.timer_reset_button.setEnabled(True); self.timer_set_button.setEnabled(True)
            if self.timer_running and not self.count_timer.isActive(): self.count_timer.start()
            self.update_timer_display()
        elif mode == "stopwatch":
            self.timer_mode_button.setText("秒表")
            self.timer_start_pause_button.setEnabled(True); self.timer_reset_button.setEnabled(True); self.timer_set_button.setEnabled(False)
            if self.timer_running and not self.count_timer.isActive(): self.count_timer.start()
            self.update_timer_display()
        else:
            self.timer_mode_button.setText("时钟")
            self.timer_start_pause_button.setEnabled(False); self.timer_reset_button.setEnabled(False); self.timer_set_button.setEnabled(False)
            self.timer_running = False; self.count_timer.stop(); self._update_clock(); self.clock_timer.start()

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
        self.update_dynamic_fonts()

    def _update_clock(self) -> None:
        self.time_display_label.setText(time.strftime("%H:%M:%S")); self.update_dynamic_fonts()

    def play_timer_sound(self) -> None:
        if SOUNDDEVICE_AVAILABLE:
            def _play() -> None:
                try:
                    fs = 44100; duration = 0.5; frequency = 880
                    t = np.linspace(0, duration, int(fs * duration), endpoint=False)
                    data = 0.4 * np.sin(2 * np.pi * frequency * t)
                    sd.play(data.astype(np.float32), fs); sd.wait()
                except Exception:
                    QApplication.beep()
            threading.Thread(target=_play, daemon=True).start()
        else:
            QApplication.beep()

    def on_group_change(self, group_name: Optional[str] = None, initial: bool = False) -> None:
        if group_name is None: group_name = self.group_combo.currentText()
        self.current_group_name = group_name
        if self.student_data.empty: self._shuffled_indices = []; return
        if group_name == "全部": idx = list(self.student_data.index)
        else: idx = list(self.student_data[self.student_data["分组"].astype(str).str.upper() == group_name].index)
        random.shuffle(idx); self._shuffled_indices = idx; self.current_student_index = None
        self.display_current_student()
        if not initial: self.roll_student(speak=False)

    def roll_student(self, speak: bool = True) -> None:
        if self.mode != "roll_call": return
        if not self._shuffled_indices:
            QMessageBox.information(self, "提示", f"'{self.current_group_name}' 的同学已经全部点到，重新开始一次抽取。")
            self.on_group_change(); return
        self.current_student_index = self._shuffled_indices.pop(); self.display_current_student()
        if speak and self.speech_enabled and self.tts_manager and self.tts_manager.available:
            stu = self.student_data.loc[self.current_student_index]
            name = str(stu["姓名"]) if "姓名" in stu and pd.notna(stu["姓名"]) else ""
            if name: self.tts_manager.speak(name)

    def display_current_student(self) -> None:
        if self.current_student_index is None:
            self.id_label.setText("点击抽取"); self.name_label.setText("学生")
        else:
            stu = self.student_data.loc[self.current_student_index]
            sid = str(stu["学号"]) if pd.notna(stu["学号"]) else ""; name = str(stu["姓名"]) if pd.notna(stu["姓名"]) else ""
            self.id_label.setText(sid if self.show_id else ""); self.name_label.setText(name if self.show_name else "")
            if not self.show_id: self.id_label.setText("")
            if not self.show_name: self.name_label.setText("")
        self.update_display_layout(); self.update_dynamic_fonts()

    def update_display_layout(self) -> None:
        self.id_label.setVisible(self.show_id); self.name_label.setVisible(self.show_name)
        layout: QGridLayout = self.roll_call_frame.layout()
        layout.setColumnStretch(0, 1); layout.setColumnStretch(1, 1)
        if not self.show_id: layout.setColumnStretch(0, 0)
        if not self.show_name: layout.setColumnStretch(1, 0)
        # 延迟更新字体，以确保布局计算完成
        QTimer.singleShot(0, self.update_dynamic_fonts)

    def update_dynamic_fonts(self) -> None:
        for lab in (self.id_label, self.name_label):
            if not lab.isVisible(): continue
            w = max(40, lab.width()); h = max(40, lab.height()); text = lab.text()
            size = self._calc_font_size(w, h, text)
            lab.setFont(QFont("Microsoft YaHei UI", size, QFont.Weight.Bold))
        if self.timer_frame.isVisible():
            text = self.time_display_label.text()
            w = max(60, self.time_display_label.width())
            h = max(60, self.time_display_label.height())
            size = self._calc_font_size(w, h, text)
            self.time_display_label.setFont(QFont("Consolas", size, QFont.Weight.Bold))

    def _calc_font_size(self, w: int, h: int, text: str) -> int:
        # [FIX] 优化字体大小计算逻辑，避免字体缩小问题
        if not text or w < 20 or h < 20:
            return self.MIN_FONT_SIZE
        # 减去样式表中的内边距 (padding: 8px)
        w_eff = w - 16
        h_eff = h - 16
        if w_eff < 1 or h_eff < 1:
            return self.MIN_FONT_SIZE

        is_cjk = any("\u4e00" <= c <= "\u9fff" for c in text)
        length = max(1, len(text))
        
        # 根据宽度估算：CJK字符宽度约等于字号，西文约0.6倍
        size_by_width = w_eff / (length * (1.0 if is_cjk else 0.6))
        
        # 根据高度估算：字符视觉高度约为字号的0.75倍
        size_by_height = h_eff * 0.75
        
        # 取两者中较小的值以确保文本能完全显示
        final_size = int(min(size_by_width, size_by_height))
        
        return max(self.MIN_FONT_SIZE, min(self.MAX_FONT_SIZE, final_size))

    def showEvent(self, e) -> None:
        super().showEvent(e)
        self.visibility_changed.emit(True)
        self.update_dynamic_fonts()
        
    def resizeEvent(self, e: QResizeEvent) -> None:
        super().resizeEvent(e)
        # [FIX] 在窗口大小改变后更新字体，确保使用正确的尺寸计算
        QTimer.singleShot(0, self.update_dynamic_fonts)

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


# ---------- 数据 ----------
def load_student_data(parent: Optional[QWidget]) -> Optional[pd.DataFrame]:
    if not (PANDAS_AVAILABLE and OPENPYXL_AVAILABLE):
        QMessageBox.warning(parent, "提示", "未安装 pandas/openpyxl，点名功能不可用。")
        return None
    file_path = RollCallTimerWindow.STUDENT_FILE
    if not os.path.exists(file_path):
        try:
            df = pd.DataFrame({"学号": [101, 102, 103], "姓名": ["张三", "李四", "王五"], "分组": ["A", "B", "A"]})
            df.to_excel(file_path, index=False)
            QMessageBox.information(parent, "提示", f"未找到学生名单，已为您创建模板文件：{file_path}")
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
class LauncherWindow(QWidget):
    def __init__(self, settings_manager: SettingsManager, student_data: Optional[pd.DataFrame]) -> None:
        super().__init__(None, Qt.WindowType.Tool | Qt.WindowType.FramelessWindowHint | Qt.WindowType.WindowStaysOnTopHint)
        self.settings_manager = settings_manager
        self.student_data = student_data
        self.overlay: Optional[OverlayWindow] = None
        self.roll_call_window: Optional[RollCallTimerWindow] = None
        self._dragging = False; self._drag_offset = QPoint()

        self.setAttribute(Qt.WidgetAttribute.WA_StyledBackground, True)
        self.setStyleSheet(
            """
            QWidget#launcherContainer { background-color: rgba(32, 33, 36, 210); border-radius: 12px; border: 1px solid rgba(255,255,255,60); }
            QPushButton { color: white; background-color: rgba(95, 99, 104, 210); border: none; border-radius: 8px; padding: 6px 10px; }
            QPushButton:hover { background-color: rgba(138, 180, 248, 210); }
            QCheckBox { color: white; }
            """
        )

        container = QWidget(self); container.setObjectName("launcherContainer")
        layout = QVBoxLayout(self); layout.setContentsMargins(0, 0, 0, 0); layout.addWidget(container)
        v = QVBoxLayout(container); v.setContentsMargins(12, 12, 12, 12); v.setSpacing(8)

        row = QHBoxLayout(); row.setSpacing(6)
        self.paint_button = QPushButton("画笔"); self.paint_button.clicked.connect(self.toggle_paint); row.addWidget(self.paint_button)
        self.roll_call_button = QPushButton("点名/计时"); self.roll_call_button.clicked.connect(self.toggle_roll_call); row.addWidget(self.roll_call_button)
        v.addLayout(row)

        bottom = QHBoxLayout(); bottom.setSpacing(6)
        self.autostart_check = QCheckBox("开机启动"); self.autostart_check.stateChanged.connect(self.toggle_autostart); bottom.addWidget(self.autostart_check)

        right = QHBoxLayout(); right.setSpacing(4)
        self.about_button = QPushButton("关于"); self.about_button.setFixedWidth(48); self.about_button.clicked.connect(self.show_about); right.addWidget(self.about_button)
        self.exit_button = QPushButton("退出"); self.exit_button.setFixedWidth(48)
        self.exit_button.clicked.connect(QApplication.instance().quit)  # [FIX] 仅此处真正退出应用
        right.addWidget(self.exit_button)
        bottom.addLayout(right); v.addLayout(bottom)

        s = self.settings_manager.load_settings().get("Launcher", {})
        self.move(int(s.get("x", "120")), int(s.get("y", "120")))

        startup = self.settings_manager.load_settings().get("Startup", {})
        autostart_enabled = str_to_bool(startup.get("autostart_enabled", "False"), False)
        self.autostart_check.setChecked(autostart_enabled and WINREG_AVAILABLE)
        self.autostart_check.setEnabled(WINREG_AVAILABLE)

        if not PANDAS_AVAILABLE or not OPENPYXL_AVAILABLE or self.student_data is None:
            self.roll_call_button.setEnabled(False)

        for w in (self, container, self.paint_button, self.roll_call_button, self.autostart_check):
            w.installEventFilter(self)

    def eventFilter(self, obj, e) -> bool:
        if e.type() == QEvent.Type.MouseButtonPress and e.button() == Qt.MouseButton.LeftButton:
            self._dragging = True; self._drag_offset = e.globalPosition().toPoint() - self.pos()
        elif e.type() == QEvent.Type.MouseMove and self._dragging and e.buttons() & Qt.MouseButton.LeftButton:
            self.move(e.globalPosition().toPoint() - self._drag_offset)
        elif e.type() == QEvent.Type.MouseButtonRelease and e.button() == Qt.MouseButton.LeftButton:
            self._dragging = False; self.save_position()
        return super().eventFilter(obj, e)

    def save_position(self) -> None:
        settings = self.settings_manager.load_settings()
        launcher = settings.get("Launcher", {}); launcher["x"] = str(self.x()); launcher["y"] = str(self.y()); settings["Launcher"] = launcher
        startup = settings.get("Startup", {}); startup["autostart_enabled"] = bool_to_str(self.autostart_check.isChecked()); settings["Startup"] = startup
        self.settings_manager.save_settings(settings)

    def toggle_paint(self) -> None:
        if self.overlay is None: self.overlay = OverlayWindow(self.settings_manager)
        if self.overlay.isVisible():
            self.overlay.hide_overlay(); self.paint_button.setText("画笔")
        else:
            self.overlay.show_overlay(); self.paint_button.setText("隐藏画笔")

    def toggle_roll_call(self) -> None:
        if self.student_data is None:
            QMessageBox.warning(self, "提示", "学生数据加载失败，无法打开点名器。"); return
        if self.roll_call_window is None:
            self.roll_call_window = RollCallTimerWindow(self.settings_manager, self.student_data)
            # 当点名窗真正 close() 时，恢复按钮文本
            self.roll_call_window.window_closed.connect(self.on_roll_call_window_closed)
            # 显示/隐藏只影响可见性，不退出应用
            self.roll_call_window.visibility_changed.connect(self.on_roll_call_visibility_changed)
            self.roll_call_window.show()
            self.roll_call_button.setText("隐藏点名")
        else:
            if self.roll_call_window.isVisible():
                self.roll_call_window.hide()                   # [FIX] 隐藏而非关闭
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
        if self.roll_call_window is not None: self.roll_call_window.close()  # 这里是真关闭（退出应用时）
        if self.overlay is not None: self.overlay.close()
        super().closeEvent(e)


# ---------- 入口 ----------
def main() -> None:
    ensure_high_dpi_awareness()
    app = QApplication(sys.argv)
    app.setQuitOnLastWindowClosed(False)  # [FIX] 防止“隐藏/关闭”任意窗口导致应用退出

    settings_manager = SettingsManager()
    student_data = load_student_data(None) if PANDAS_AVAILABLE else None

    window = LauncherWindow(settings_manager, student_data)
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()