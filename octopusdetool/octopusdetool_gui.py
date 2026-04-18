#!/usr/bin/env python3
"""
Octopus Energy Germany Smart Meter Data Logger - GUI Version

A PySide6-based GUI for fetching smart meter consumption data from
Octopus Energy Germany API and saving it to CSV, Excel, JSON, or YAML.
"""

from __future__ import annotations

import argparse
import base64
import calendar
import csv
import hashlib
import io
import json
import os
import platform
import sys
import traceback
from contextlib import contextmanager, redirect_stderr, redirect_stdout
from datetime import date, datetime, time, timedelta
from pathlib import Path
from typing import TypeVar

from cryptography.hazmat.primitives.ciphers.aead import AESGCM
from PySide6.QtCore import QDate, QFile, QIODeviceBase, QObject, QSize, Qt
from PySide6.QtGui import QColor, QIcon, QPainter, QPalette, QPixmap, QStandardItem, QStandardItemModel
from PySide6.QtUiTools import QUiLoader
from PySide6.QtWidgets import (
    QAbstractItemView,
    QApplication,
    QCalendarWidget,
    QCheckBox,
    QComboBox,
    QDateEdit,
    QFileDialog,
    QFrame,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QProgressBar,
    QSizePolicy,
    QMenu,
    QScrollArea,
    QTableView,
    QTabWidget,
    QToolTip,
    QVBoxLayout,
    QWidget,
    QStyleFactory,
)

from octopusdetool.analysis_view import DisplayBucket, TariffChartView
from octopusdetool.octopusdetool import (
    APP_TIMEZONE,
    DEFAULT_MONTHLY_BASE_PRICE_EUR,
    DEFAULT_TARIFF_GO_CT,
    DEFAULT_TARIFF_HEAT_HIGH_CT,
    DEFAULT_TARIFF_HEAT_LOW_CT,
    DEFAULT_TARIFF_HEAT_STANDARD_CT,
    DEFAULT_TARIFF_STANDARD_CT,
    DISPLAY_NAME_DYNAMIC_OCTOPUS,
    DISPLAY_NAME_OCTOPUS_GO,
    DISPLAY_NAME_OCTOPUS_HEAT,
    OctopusGermanyClient,
    TARIFF_INTELLIGENT_GO,
    TARIFF_INTELLIGENT_HEAT,
    TariffAgreement,
    TariffRate,
    classify_tariff_zone,
    detect_excel_template_type,
    ensure_excel_template,
    ensure_smartmeter_data_folder,
    fill_excel_template,
    format_datetime,
    get_default_excel_path,
    get_default_tariff_settings_for_type,
    get_default_output_path,
    get_app_config_folder,
    get_smartmeter_data_folder,
    init_app_config_folder,
    load_excel_tariff_settings,
    load_existing_consumption_data,
    normalize_datetime,
    save_to_json,
    save_to_yaml,
)


CONFIG_WRITABLE = True
CONFIG_FILE = get_app_config_folder() / "config.json"
CONFIG_ENCRYPTION_VERSION = 1
CONFIG_ENCRYPTED_FIELDS = ("email", "password")
CONFIG_AES_KEY = hashlib.sha256(b"octopusdetool_rocks!").digest()
CONFIG_SAVE_FLAG = "save_config_enabled"
AUTO_OUTPUT_FLAG = "auto_output_enabled"
TARIFF_TYPE_CONFIG_KEY = "tariff_type"
LAST_TARIFF_CODE_CONFIG_KEY = "last_tariff_code"
LAST_TARIFF_DISPLAY_NAME_CONFIG_KEY = "last_tariff_display_name"
TARIFF_RATES_CONFIG_KEY = "tariff_rates"
EXCEL_EXPORT_SUPPORTED_CONFIG_KEY = "excel_export_supported"
EXCEL_EXPORT_REASON_CONFIG_KEY = "excel_export_reason"
OUTPUT_EXTENSIONS = {
    "excel": ".xlsx",
    "csv": ".csv",
    "json": ".json",
    "yaml": ".yaml",
}
def _migrate_config_from_data_folder() -> None:
    """Migrate config.json from data folder to config folder if needed."""
    data_config = get_smartmeter_data_folder() / "config.json"
    app_config = get_app_config_folder() / "config.json"
    if data_config.exists() and not app_config.exists():
        try:
            get_app_config_folder().mkdir(parents=True, exist_ok=True)
            app_config.write_text(data_config.read_text(encoding="utf-8"), encoding="utf-8")
            data_config.unlink()
        except Exception:
            pass


GERMAN_MONTH_NAMES = [
    "Januar",
    "Februar",
    "Maerz",
    "April",
    "Mai",
    "Juni",
    "Juli",
    "August",
    "September",
    "Oktober",
    "November",
    "Dezember",
]
GERMAN_MONTH_ABBR = [
    "Jan",
    "Feb",
    "Maerz",
    "Apr",
    "Mai",
    "Jun",
    "Jul",
    "Aug",
    "Sep",
    "Okt",
    "Nov",
    "Dez",
]
GERMAN_WEEKDAY_NAMES = [
    "Montag",
    "Dienstag",
    "Mittwoch",
    "Donnerstag",
    "Freitag",
    "Samstag",
    "Sonntag",
]
GERMAN_WEEKDAY_ABBR = ["Mo", "Di", "Mi", "Do", "Fr", "Sa", "So"]
STATUS_CHECK_PNG_B64 = (
    "iVBORw0KGgoAAAANSUhEUgAAADAAAAAwCAYAAABXAvmHAAAACXBIWXMAAA7EAAAOxAGVKw4bAAAB4klEQVRoge2WsUtW"
    "URjGf5oEQQS6BEFxoD+gQCFwMAicqkPSVFM5NQhBIORwKDmDQa5FReIgUfOLSxZoBC1CmIPzgVIEl1ILdLAGv0Hw+HXv"
    "e79rBO9v+85z7vM+D/e7hwOGYRiGYRjGf0vbvxjqUzgD3AbOA0eBLWBGXHxW1qujxdma4lNoB+4DD9gNvpe3Gs9DK+BT"
    "OA68AS5n5Cfi4nON76EU8Cl0ATNAd0aeBu5qvWv/BnwKncAscC4jfwb6xMWfWv9aC/gUjgHvgd6M/A24IC6uVJnRXuXh"
    "ZvgU2oAp8uE3gatVw0ONBYCHwPXM+g5wU1xcaMWQWv5CPoUrgBzgPywujrdqlqpA4zy/BtwCeoAgLk40tNPAAtCVeXRS"
    "XBzURc2jPUbfAZf2/B4BJhrFXpEP/wm4o5x3IKXfgE/hBPAjI11k94Mdy2jLQI+4uFp23t/QvIENYJv9V4HXwMnM/i1g"
    "oI7woDiFxMXfwFJGOgUcyawPiYvzZecURXuMfii4b1JcfKmcUQhtASmwZwkYUvoXRltgDvjaRN8GboiLv5T+hVEVEBd3"
    "gBdNtoyKi4u6SOWocpV4Cqxn1r8Ajyv4lqLSVcKn0A/cA84C34GPwCNxca0F2QzDMAzDMAzDqJc/AHFqjRWuBmQAAAAA"
    "SUVORK5CYII="
)
STATUS_CROSS_PNG_B64 = (
    "iVBORw0KGgoAAAANSUhEUgAAADAAAAAwCAYAAABXAvmHAAAACXBIWXMAAA7EAAAOxAGVKw4bAAACCElEQVRoge3XP2sU"
    "QRjH8c+qkFZBRFALjZ1e5RvIH0FNI1goKvcCRAthrfUVbGFvt9jEYBTLgFHEzi7+aRRB7bQXIXAWuQ3JecntzCxCYL7t"
    "znP7/fHczDxLJpPJZDKZPc6gXx74HzXj2Jf6A4N+eQ3vBv3ySEDNAj4O+uXJ1PcXKcVD+cfYjw+YLerq54SaBTzFFL5h"
    "pqirr7EO0QFG5Bt2DTEi35AUIirADvIN7zE3GmLQLy9h2Xb5hugQwQEmyDdsCzHolxfxzHj5hqgQQQEG/XIKn9Bm861h"
    "HudMlm94WNTV3RCnoFOoqKs/OI/vLZb38FZ7+UWUIT7E74FpvMLxmPoxLOJmUVfroYUpp1BXIaLlSb8HTmNVfIgnuBEr"
    "T2IANjvxGscCS5Pl6SAAm514g6MtS5ZxNVWeDmahIdM4FLj+YBcv7uIvdAHPtTsqt7KG+Umz0yRSN3GsfENyiJRjNFW+"
    "Yezs1JbYi6wr+YboEMGbOFD+BX60WHcWqyEfRQ1BAYbD3CPt5JdwBbPahTiDByE+xA1zC/g1YekSrhd1tV7U1WfMmBxi"
    "BfdCfIjfAz28xOExjzflR2p2m51WcLmoq9+hLlEXWVFXa5jzbyfGyg9rvtjoxOgoHi1P+j2wtRM7yo/UnLLRiRMS5enm"
    "Ju7hDm63nW2GIe7jVop8JpPJZDKZTGaP8xedeMcuVRjS0gAAAABJRU5ErkJggg=="
)
CALENDAR_ICON_PNG_B64 = (
    "iVBORw0KGgoAAAANSUhEUgAAADAAAAAwCAYAAABXAvmHAAAACXBIWXMAAA7EAAAOxAGVKw4bAAADzElEQVRoge2X3Usc"
    "VxjGf+/M7hpdP3BjQxvS+oF6YbZNQNpir7L2otKaRoi52UIrlmLYG8Eqza2VKrmoIAj+BXsRKAWvKm5LpREtJJrQNrXF"
    "xEaLSbbqlhRd1+3OnF5oFrea4IyDtDA/WBjOnPd9nmfnzJkZcHFxcXFx+R8jdorWXq8ulkyJBGZmHifq60tKS+0buPNg"
    "c9N/IqlVeI7nkZnZkAkyVuotB9hofO2yEq4CGciEUJ4JBV6rfbIGlPkpmnZaKS6C+mHVXHmzcuJ+6qD1HquCptAkUCyK"
    "u+hq1TRYE6Haap8dDGXKkobKB/xAQ6l54gW4/9tBG2g2hUGZH/tjtx7oSn2ye9j7wUeg5/4veVd6obhkbwvUt0UTN64V"
    "GDf6RWQJEcmzaMPyFcgicn7t3CuzSmjKafhOC8b1CcjPRyWTiNeDVl2LHjyDWvlje5KmYf76M6IkuH7u1bMbwnHgOTs2"
    "bAdQ8GGefqxNKaX/+5xWF0R7qQLz0UPEXwiA+AvxNJ0HpciMfrE9UXhedLmpUAJooMyUJ2Na8WF/CYkI4EFE0D143r4A"
    "Ph+CgFIgsv3bwfjxNlp5JdqL5Rg/3c6OK9BBnvjI+P7W0lZs2F9CuzEy6C+fxfvuRTJff4X5yx28Fy6h1WyQ+e4bzLUVS"
    "KUwbn4PpoK0JY/PxJkAwNbVXtB1MAwANi+/D+b2cebLa8jJU+inz7D1+WdOSQIOBgCy5oGs+SeotVVSvVdQD5cdlbQc4"
    "NEboa2TFeVoYv/2eRqJRILHiXW4PnPgGssBfg+9tXWqoQFd37P5HJq/7t1jaXER+vsPXOPsEtohnU7j9XoREdbX1/edo"
    "2kaBQUFh9ZyLIBhGExOThKNRtXCwoLEYjEAKisr951fX1/P2NjYoXUdCTA9PU04HFbLy8sCSG1tbfZca2trzty5uTmmp"
    "qaoqqpyQtqZAPF4nJqaGmlra1MDAwM5b7gjIyPZY6UUzc3N6LpOJBJxQvoQT+JdtLS0MD4+TkdHh4g8/Q19dnaWWCxGY"
    "2MjwWDQCWlnAhwEpRSDg4MAdHZ2Kqf6HlmA+fl5RkdHqaurIxQK2foS3I8jCaCUYmhoiHQ6TSQSweu1/QG3hyMJEI/Hi"
    "UajlJWVEQ6HHe19JAGGh4dJJpO0t7dTWFjoaG9Hn8R+v5/u7m4CgUB2zDRNAoEAPT09jm2du3E0QFFREX19fTljmqbR1"
    "dXlpEwOdgLcWlxcfM/n8zm+/BKJxLxpmgkrNXa2M4nFYkEROWaj9pmkUqm7zc3Nfzrd18XFxcXFxcXF5T/KP/cWO467"
    "9H7sAAAAAElFTkSuQmCC"
)
CALENDAR_PICK_ICON_PNG_B64 = (
    "iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAACXBIWXMAAA9hAAAPYQGoP6dpAAABUElEQVQ4ja2S"
    "P04CURDGf/NETVxJ3LgtF6AxwT8FNBzAeBML9ACeQAqvwA2MDd02WohiLPQCUhFYCD5tDG8s9rEaXIXCaWYy833ffJP3"
    "4L/DDrVhh9pYti+z4n2gB85wDBz61uMcdsfnK3FcBJF0MgFPvgZGwOoCgQ8gNI7aRiS3qb1EWzbRvqqu2URjm2j84wTf"
    "95i+TbQFUPDzkkDPjqkCTwCvI61/F1BN+3ZM1UBPoZSd8DbSO1V257f+FSLcB6HsFby6FXhwwskyZKM0VbEAJrMIk2Io"
    "cTGUWBwVcVSKocRGmOCoG/maK0xmvELeBoRzXzXdlCMRztwUgO48NF9AOc3srnDppmnOg+YKbG5Lc1YHW9LN25wt8PkF"
    "KKvq2m/AzJzqOlD2HP+MA91Xww3pT3xeoFEGQnFUg0g6BiCIpGMcNaC9yAHQnpGXwC6OT215lNxGJBV2AAAAAElFTkSu"
    "QmCC"
)
TRASH_ICON_PNG_B64 = (
    "iVBORw0KGgoAAAANSUhEUgAAADAAAAAwCAYAAABXAvmHAAAACXBIWXMAAA7EAAAOxAGVKw4bAAADyElEQVRoge3YW8hV"
    "RRQH8N/+vJJbrVNeKiIrI027QEhlRRRZWS9hdDNIwqISInsUNLBCe0i6UBKJ9RJBD5lFWZlBJYlpZamYlPR9WJiGTl5G"
    "Q1Onh9kSfFTE2afoYf9fzpm99sxaa89a/7VmaNCgQYMGNVB0crEYUh8Mx0iUGIgD2IPd2F22ithJnbUciCH1x/W4ARdj"
    "LAYgYZ9seIEhskNd2I6N+ATvlK3iyzo2tO1ADGkGZmME1uBdrMNmfF+2iiO93u+Hs3C27OwkXISvMbNsFSvbtaUd41+M"
    "IR2KIc2PIY2osc7JMaQF1VrTOmnj3ykdF0NKMaTpHVxzdgxpZztzu9qYc0r1u60dhX+BXRhSkcC/ixjSwBhSTwxpVwxp"
    "RgxpUI21yhjSQzGk/TGkl9pZo60kjiGNwTKcIdPkp/gYX2ALespWcbjXnAEYhTMxAZfjMhyHtzC1bBX7/xMHKoMexRyZ"
    "Kof2Eh9BQJTDdDBO6KXvcPXOcFxatorV7djRTg4cwwbsxW9Yicl4pZJ9hp9wamXgj/imks3DJTLlDpVrxsZ2jajjwHq5"
    "QK3HMLRwnlwPtuK5SrYFT+EQXsBVGIe+cth1d7o6/yPEkPrEkA7EkCbFkC6MIR2NIT1Wyd6PIe2JIZ0UQ7qgks2tZIti"
    "SAdjSKfFkObFkJbWsaNuK7EGmzAaH+AOvIw7q+e/yAm7BNOwGHfJBDBWDq/lZauYU8eOtlF9zRRDWlaNH6jGk2NIx1e7"
    "cEw2rZJNiSF1xZA2VOPb6tjQt6YP6+TmrH8MaRGuwC14Bj/jcVwZQ1pcyW7GE5gq79B4mQzaRp0kJjdiI/E0puM1mdO3"
    "YwyWYi7uxhtlq1iCtbgRq+TE/raOAXVzYLBcB3ZgChagH1bgQyyU+f5ezMdBOS+ex9syA51fx4ZaO1C2in3oljl/E3rk"
    "Sru5+t8XR2Uq/QgT8Z4cet1qhg/1Q4gcRqPwHfbjHMyUWekePCKHzUT5PHCfHD7n+p848JVc0LbJX79EH/yKQXI13uOP"
    "VmIfTpQr8Pq6yju1A11yY7ZWDp+5uBrPysk9EYvkUFqN2+Xwqr0DdWmU7ECBa3GNnMC3yskc5G5zgnyEXCV3oTvlA/4P"
    "HdBfHzGknVVRmlWN36yK2LAY0uiq5Xi1ki2s3l3RCd0duVaJIT0oF6+tcsN2Pz6X86KffG4YLzd4s6pnN5WtYnld3R27"
    "F4ohTcDDuE7uTP8MO/A6nixbRXcn9Hb0YgtiSAVOl6l1qHy42YstZavo5Dm6QYMGDRo0aNCgQYMG9fA7K0q4aHsy7twA"
    "AAAASUVORK5CYII="
)
_EMBEDDED_PIXMAP_CACHE: dict[str, QPixmap] = {}
WidgetType = TypeVar("WidgetType", bound=QObject)


class _TeeStream:
    """Write output to multiple streams."""

    def __init__(self, *streams):
        self.streams = [stream for stream in streams if stream is not None]

    def write(self, data):
        for stream in self.streams:
            stream.write(data)
        return len(data)

    def flush(self):
        for stream in self.streams:
            stream.flush()


def _configure_macos_qt_theme(app: QApplication) -> None:
    fusion_style = QStyleFactory.create("Fusion")
    if fusion_style is not None:
        app.setStyle(fusion_style)

    palette = QPalette()
    palette.setColor(QPalette.ColorRole.Window, QColor("#16042c"))
    palette.setColor(QPalette.ColorRole.WindowText, QColor("#f4eeff"))
    palette.setColor(QPalette.ColorRole.Base, QColor("#240748"))
    palette.setColor(QPalette.ColorRole.AlternateBase, QColor("#2e1160"))
    palette.setColor(QPalette.ColorRole.ToolTipBase, QColor("#240748"))
    palette.setColor(QPalette.ColorRole.ToolTipText, QColor("#ffffff"))
    palette.setColor(QPalette.ColorRole.Text, QColor("#f4eeff"))
    palette.setColor(QPalette.ColorRole.Button, QColor("#6f4df6"))
    palette.setColor(QPalette.ColorRole.ButtonText, QColor("#ffffff"))
    palette.setColor(QPalette.ColorRole.BrightText, QColor("#ffffff"))
    palette.setColor(QPalette.ColorRole.Highlight, QColor("#6f4df6"))
    palette.setColor(QPalette.ColorRole.HighlightedText, QColor("#ffffff"))
    palette.setColor(QPalette.ColorRole.Link, QColor("#8d71ff"))
    palette.setColor(QPalette.ColorRole.LinkVisited, QColor("#cbbcff"))
    app.setPalette(palette)


def _pixmap_from_base64(encoded_png: str) -> QPixmap:
    pixmap = _EMBEDDED_PIXMAP_CACHE.get(encoded_png)
    if pixmap is None:
        loaded = QPixmap()
        loaded.loadFromData(base64.b64decode(encoded_png), "PNG")
        _EMBEDDED_PIXMAP_CACHE[encoded_png] = loaded
        pixmap = loaded
    return pixmap


def _icon_from_base64(encoded_png: str) -> QIcon:
    return QIcon(_pixmap_from_base64(encoded_png))


class StatusIndicatorCheckBox(QCheckBox):
    def __init__(self, parent: QWidget | None = None):
        super().__init__(parent)
        self.setCursor(Qt.CursorShape.PointingHandCursor)
        self.setMinimumHeight(30)
        self.setStyleSheet("background: transparent; border: none;")

    def sizeHint(self) -> QSize:
        metrics = self.fontMetrics()
        text_width = metrics.horizontalAdvance(self.text())
        return QSize(text_width + 44, max(30, metrics.height() + 10))

    def paintEvent(self, _event) -> None:
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        painter.setRenderHint(QPainter.RenderHint.SmoothPixmapTransform)

        opacity = 1.0 if self.isEnabled() else 0.45
        painter.setOpacity(opacity)

        indicator_size = 24
        indicator_rect = self.rect().adjusted(0, 0, 0, 0)
        indicator_rect.setWidth(indicator_size)
        indicator_rect.setHeight(indicator_size)
        indicator_rect.moveTop((self.height() - indicator_size) // 2)

        painter.setPen(QColor("#6f4df6"))
        painter.setBrush(QColor("#2e1160"))
        painter.drawRoundedRect(indicator_rect, 6, 6)

        icon_source = STATUS_CHECK_PNG_B64 if self.isChecked() else STATUS_CROSS_PNG_B64
        icon_pixmap = _pixmap_from_base64(icon_source)
        icon_rect = indicator_rect.adjusted(1, 1, -1, -1)
        painter.drawPixmap(icon_rect, icon_pixmap)

        text_rect = self.rect().adjusted(indicator_size + 12, 0, 0, 0)
        text_color = QColor("#f4eeff") if self.isEnabled() else QColor("#a498cb")
        painter.setPen(text_color)
        painter.drawText(text_rect, Qt.AlignmentFlag.AlignVCenter | Qt.AlignmentFlag.AlignLeft, self.text())

        if self.hasFocus():
            painter.setPen(QColor("#bfb4ff"))
            painter.setBrush(Qt.BrushStyle.NoBrush)
            focus_rect = self.rect().adjusted(0, 1, -1, -2)
            painter.drawRoundedRect(focus_rect, 6, 6)

        painter.end()


class CurrencyToggleSwitch(QCheckBox):
    def __init__(self, parent: QWidget | None = None):
        super().__init__(parent)
        self._pressed_inside = False
        self.setCursor(Qt.CursorShape.PointingHandCursor)
        self.setFixedSize(58, 32)
        self.setText("")
        self.setStyleSheet("background: transparent; border: none;")

    def sizeHint(self) -> QSize:
        return QSize(58, 32)

    def hitButton(self, pos) -> bool:
        return self.rect().contains(pos)

    def mousePressEvent(self, event) -> None:
        if event.button() == Qt.MouseButton.LeftButton and self.rect().contains(event.position().toPoint()):
            self._pressed_inside = True
            self.setDown(True)
            event.accept()
            return
        self._pressed_inside = False
        super().mousePressEvent(event)

    def mouseMoveEvent(self, event) -> None:
        if self._pressed_inside:
            self.setDown(self.rect().contains(event.position().toPoint()))
            event.accept()
            return
        super().mouseMoveEvent(event)

    def mouseReleaseEvent(self, event) -> None:
        if event.button() == Qt.MouseButton.LeftButton:
            should_toggle = self._pressed_inside and self.rect().contains(event.position().toPoint())
            self._pressed_inside = False
            self.setDown(False)
            if should_toggle:
                self.click()
                event.accept()
                return
        self._pressed_inside = False
        super().mouseReleaseEvent(event)

    def paintEvent(self, _event) -> None:
        track_rect = self.rect().adjusted(1, 3, -1, -3)
        knob_diameter = track_rect.height() - 6
        knob_x = track_rect.left() + 3
        if self.isChecked():
            knob_x = track_rect.right() - knob_diameter - 2

        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        painter.setPen(QColor("#bfb4ff"))
        painter.setBrush(QColor("#d9d3ff"))
        painter.drawRoundedRect(track_rect, track_rect.height() / 2, track_rect.height() / 2)

        knob_rect = track_rect.adjusted(3, 3, 0, -3)
        knob_rect.setLeft(knob_x)
        knob_rect.setWidth(knob_diameter)
        painter.setPen(Qt.PenStyle.NoPen)
        painter.setBrush(QColor("#6f4df6"))
        painter.drawEllipse(knob_rect)
        painter.end()


class ViewCalendarPopup(QFrame):
    def __init__(self, parent: QWidget):
        super().__init__(parent, Qt.WindowType.Popup)
        self.setObjectName("viewCalendarPopup")
        self.setFrameShape(QFrame.Shape.StyledPanel)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(10, 10, 10, 10)

        self.calendar = QCalendarWidget(self)
        self.calendar.setFirstDayOfWeek(Qt.DayOfWeek.Monday)
        self.calendar.setGridVisible(True)
        layout.addWidget(self.calendar)


class OctopusSmartMeterGUI:
    WINDOW_SCREEN_FRACTION = 0.92
    RESIZE_STEP = 20

    def __init__(self, app: QApplication, debug_enabled: bool = False):
        self.app = app
        self._debug_enabled_from_cli = debug_enabled
        self.window = self._load_ui()
        self._bind_widgets()
        self._clear_line_edit_actions: list = []
        self._replace_line_edit_clear_buttons()
        self._replace_data_tab_checkboxes()
        self._replace_currency_toggle()
        self._setup_analysis_widgets()
        self._setup_view_calendar_popup()
        self._apply_popup_styling()
        self._configure_range_date_edits()
        self._configure_settings_fields()
        self._configure_primary_buttons()
        self._configure_tooltip_palette()
        self._configure_view_calendar_button()
        self._set_window_icon()
        self.current_tariff_type = TARIFF_INTELLIGENT_GO
        self.current_tariff_display_name = ""
        self.current_tariff_valid_from = ""
        self.current_tariff_valid_to = ""
        self.current_tariff_rates: list[TariffRate] = []
        self._excel_export_supported = True
        self._excel_export_reason = ""
        self._has_saved_base_price = False

        self.template_path = ensure_excel_template()
        self.default_tariff_settings = load_excel_tariff_settings(self.template_path)
        self.csv_path = get_default_output_path()
        self.existing_data: list[dict] = []
        self.latest_timestamp: datetime | None = None
        self.last_output_format = "excel"
        self._analysis_date_initialized = False

        self._set_initial_values()
        self._connect_signals()
        _migrate_config_from_data_folder()
        self.load_config()
        if self._debug_enabled_from_cli:
            self.debug_checkbox.setChecked(True)
        self.check_existing_data()

    def _load_ui(self) -> QWidget:
        ui_path = Path(__file__).resolve().with_name("octopusdetool_gui.ui")
        ui_file = QFile(str(ui_path))
        if not ui_file.open(QIODeviceBase.OpenModeFlag.ReadOnly):
            raise FileNotFoundError(f"Could not open UI file: {ui_path}")

        loader = QUiLoader()
        window = loader.load(ui_file)
        ui_file.close()

        if window is None:
            raise RuntimeError(f"Could not load UI file: {ui_path}")

        return window

    def _find_widget(self, widget_type: type[WidgetType], name: str) -> WidgetType:
        widget = self.window.findChild(widget_type, name)
        if widget is None:
            raise RuntimeError(f"Required UI widget '{name}' was not found")
        return widget

    def _bind_widgets(self) -> None:
        self.main_tab_widget = self._find_widget(QTabWidget, "mainTabWidget")
        self.scroll_area = self._find_widget(QScrollArea, "scrollArea")
        self.scroll_area_contents = self._find_widget(QWidget, "scrollAreaWidgetContents")
        self.email_line_edit = self._find_widget(QLineEdit, "emailLineEdit")
        self.password_line_edit = self._find_widget(QLineEdit, "passwordLineEdit")
        self.show_password_checkbox = self._find_widget(QCheckBox, "showPasswordCheckBox")
        self.save_config_checkbox = self._find_widget(QCheckBox, "saveConfigCheckBox")
        self.debug_checkbox = self._find_widget(QCheckBox, "debugCheckBox")
        self.output_format_combo = self._find_widget(QComboBox, "outputFormatComboBox")
        self.auto_output_checkbox = self._find_widget(QCheckBox, "autoOutputCheckBox")
        self.output_file_line_edit = self._find_widget(QLineEdit, "outputFileLineEdit")
        self.browse_output_button = self._find_widget(QPushButton, "browseOutputButton")
        self.from_date_edit = self._find_widget(QDateEdit, "fromDateEdit")
        self.to_date_edit = self._find_widget(QDateEdit, "toDateEdit")
        self.status_value_label = self._find_widget(QLabel, "statusValueLabel")
        self.progress_bar = self._find_widget(QProgressBar, "progressBar")
        progress_policy = self.progress_bar.sizePolicy()
        progress_policy.setRetainSizeWhenHidden(True)
        self.progress_bar.setSizePolicy(progress_policy)
        self.fetch_data_button = self._find_widget(QPushButton, "fetchDataButton")

        self.tariff_type_combo = self._find_widget(QComboBox, "tariffTypeComboBox")
        self.tariff_type_label = self._find_widget(QLabel, "tariffTypeLabel")
        self.tariff_go_label = self._find_widget(QLabel, "tariffGoLabel")
        self.tariff_go_line_edit = self._find_widget(QLineEdit, "tariffGoLineEdit")
        self.tariff_standard_label = self._find_widget(QLabel, "tariffStandardLabel")
        self.tariff_standard_line_edit = self._find_widget(QLineEdit, "tariffStandardLineEdit")
        self.tariff_high_label = self._find_widget(QLabel, "tariffHighLabel")
        self.tariff_high_line_edit = self._find_widget(QLineEdit, "tariffHighLineEdit")
        self.base_price_line_edit = self._find_widget(QLineEdit, "basePriceLineEdit")
        self.save_settings_button = self._find_widget(QPushButton, "saveSettingsButton")
        self.config_path_line_edit = self._find_widget(QLineEdit, "configPathLineEdit")

        self.view_mode_combo = self._find_widget(QComboBox, "viewModeComboBox")
        self.view_date_edit = self._find_widget(QDateEdit, "viewDateEdit")
        self.view_calendar_button = self._find_widget(QPushButton, "viewCalendarButton")
        self.view_previous_button = self._find_widget(QPushButton, "viewPreviousButton")
        self.view_next_button = self._find_widget(QPushButton, "viewNextButton")
        self.view_currency_checkbox = self._find_widget(QCheckBox, "viewCurrencyCheckBox")
        self.view_range_label = self._find_widget(QLabel, "viewRangeLabel")
        self.view_total_caption_label = self._find_widget(QLabel, "viewTotalCaptionLabel")
        self.view_total_value_label = self._find_widget(QLabel, "viewTotalValueLabel")
        self.analysis_content_tabs = self._find_widget(QTabWidget, "analysisContentTabs")
        self.analysis_table_view = self._find_widget(QTableView, "analysisTableView")
        self.chart_container = self._find_widget(QFrame, "chartContainer")

    def _replace_currency_toggle(self) -> None:
        placeholder = self.view_currency_checkbox
        parent = placeholder.parentWidget()
        layout = parent.layout() if parent is not None else None
        if layout is None:
            return

        replacement = CurrencyToggleSwitch(parent)
        replacement.setObjectName("viewCurrencyCheckBox")
        replacement.setChecked(placeholder.isChecked())
        replacement.setToolTip(placeholder.toolTip())
        replacement.setAccessibleName("Waehrungsschalter")

        layout.replaceWidget(placeholder, replacement)
        placeholder.hide()
        placeholder.setParent(None)
        placeholder.deleteLater()
        self.view_currency_checkbox = replacement

    def _replace_line_edit_clear_buttons(self) -> None:
        for line_edit in (
            self.email_line_edit,
            self.password_line_edit,
            self.output_file_line_edit,
        ):
            line_edit.setClearButtonEnabled(False)
            action = line_edit.addAction(
                _icon_from_base64(TRASH_ICON_PNG_B64),
                QLineEdit.ActionPosition.TrailingPosition,
            )
            action.setToolTip("Feld leeren")
            action.triggered.connect(line_edit.clear)
            action.setVisible(bool(line_edit.text()))
            line_edit.textChanged.connect(lambda text, clear_action=action: clear_action.setVisible(bool(text)))
            self._clear_line_edit_actions.append(action)

    def _replace_data_tab_checkboxes(self) -> None:
        for attribute_name in (
            "show_password_checkbox",
            "save_config_checkbox",
            "debug_checkbox",
            "auto_output_checkbox",
        ):
            placeholder = getattr(self, attribute_name)
            parent = placeholder.parentWidget()
            layout = parent.layout() if parent is not None else None
            if layout is None:
                continue

            replacement = StatusIndicatorCheckBox(parent)
            replacement.setObjectName(placeholder.objectName())
            replacement.setText(placeholder.text())
            replacement.setToolTip(placeholder.toolTip())
            replacement.setStatusTip(placeholder.statusTip())
            replacement.setWhatsThis(placeholder.whatsThis())
            replacement.setChecked(placeholder.isChecked())
            replacement.setEnabled(placeholder.isEnabled())
            replacement.setSizePolicy(placeholder.sizePolicy())
            replacement.setAccessibleName(placeholder.accessibleName() or placeholder.text())

            layout.replaceWidget(placeholder, replacement)
            placeholder.hide()
            placeholder.setParent(None)
            placeholder.deleteLater()
            setattr(self, attribute_name, replacement)

    def _setup_analysis_widgets(self) -> None:
        chart_layout = self.chart_container.layout()
        if chart_layout is None:
            chart_layout = QVBoxLayout(self.chart_container)
            chart_layout.setContentsMargins(12, 12, 12, 12)

        self.chart_view = TariffChartView(self.chart_container)
        chart_layout.addWidget(self.chart_view)

        self.analysis_table_model = QStandardItemModel(self.analysis_table_view)
        self.analysis_table_view.setModel(self.analysis_table_model)
        self.analysis_table_view.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.analysis_table_view.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectItems)
        self.analysis_table_view.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.analysis_table_view.setAlternatingRowColors(True)
        self.analysis_table_view.verticalHeader().hide()
        self.analysis_table_view.horizontalHeader().setStretchLastSection(True)
        self.analysis_table_view.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.analysis_table_view.customContextMenuRequested.connect(
            self._show_analysis_table_context_menu
        )

    def _setup_view_calendar_popup(self) -> None:
        self.view_calendar_popup = ViewCalendarPopup(self.window)
        self.view_calendar_popup.setStyleSheet(self.window.styleSheet())
        self.view_calendar_popup.calendar.clicked.connect(self._on_view_calendar_date_selected)

    def _apply_popup_styling(self) -> None:
        combo_stylesheet = """
QComboBox {
    background-color: #240748;
    color: #f4eeff;
    border: 1px solid #6f4df6;
    border-radius: 10px;
    padding: 8px 34px 8px 10px;
    selection-background-color: #6f4df6;
    selection-color: #ffffff;
}

QComboBox::drop-down {
    border: none;
    width: 28px;
    background-color: #240748;
    subcontrol-origin: padding;
    subcontrol-position: center right;
    border-top-right-radius: 10px;
    border-bottom-right-radius: 10px;
}

QComboBox::down-arrow {
    width: 0px;
    height: 0px;
    border-left: 6px solid #240748;
    border-right: 6px solid #240748;
    border-top: 8px solid #f4eeff;
}

QComboBox QAbstractItemView {
    selection-background-color: #6f4df6;
    selection-color: #ffffff;
}
"""
        popup_stylesheet = """
QListView {
    background-color: #240748;
    color: #f4eeff;
    border: 1px solid #6f4df6;
    outline: 0;
    padding: 0;
    margin: 0;
    show-decoration-selected: 1;
}

QListView::item {
    color: #f4eeff;
    background: transparent;
    border: none;
    padding: 10px 16px;
    margin: 0;
}

QListView::item:hover {
    background-color: #4a2388;
}

QListView::item:selected {
    background-color: #6f4df6;
    color: #ffffff;
}
"""
        for combo in (self.view_mode_combo, self.output_format_combo, self.tariff_type_combo):
            combo.setStyleSheet(combo_stylesheet)
            view = combo.view()
            view.setStyleSheet(popup_stylesheet)
            view.setFrameShape(QFrame.Shape.NoFrame)
            view.setContentsMargins(0, 0, 0, 0)
            popup_window = view.window()
            popup_window.setStyleSheet("background-color: #240748; border: 1px solid #6f4df6;")
            popup_window.setContentsMargins(0, 0, 0, 0)

        abruf_combo_stylesheet = """
QComboBox {
    background-color: #1c0638;
    color: #f4eeff;
    border: 1px solid #6f4df6;
    border-radius: 10px;
    padding: 8px 34px 8px 10px;
    selection-background-color: #6f4df6;
    selection-color: #ffffff;
}

QComboBox::drop-down {
    border: none;
    width: 28px;
    background-color: #1c0638;
    subcontrol-origin: padding;
    subcontrol-position: center right;
    border-top-right-radius: 10px;
    border-bottom-right-radius: 10px;
}

QComboBox::down-arrow {
    width: 0px;
    height: 0px;
    border-left: 6px solid #1c0638;
    border-right: 6px solid #1c0638;
    border-top: 8px solid #f4eeff;
}
"""
        self.output_format_combo.setStyleSheet(abruf_combo_stylesheet)

    def _configure_range_date_edits(self) -> None:
        icon_path = self._ensure_embedded_calendar_pick_icon_file().as_posix()
        date_edit_stylesheet = f"""
QDateEdit {{
    background-color: #1c0638;
    border: 1px solid #6f4df6;
    border-radius: 10px;
    padding: 8px 34px 8px 10px;
    color: #f4eeff;
}}

QDateEdit::drop-down {{
    subcontrol-origin: padding;
    subcontrol-position: center right;
    width: 28px;
    border: none;
    background-color: #1c0638;
    border-top-right-radius: 10px;
    border-bottom-right-radius: 10px;
}}

QDateEdit::down-arrow {{
    image: url({icon_path});
    width: 16px;
    height: 16px;
}}
"""
        for date_edit in (self.from_date_edit, self.to_date_edit):
            date_edit.setStyleSheet(date_edit_stylesheet)

    def _ensure_embedded_calendar_pick_icon_file(self) -> Path:
        icon_path = Path("/tmp/octopusdetool_calendar_pick_icon.png")
        icon_bytes = base64.b64decode(CALENDAR_PICK_ICON_PNG_B64)
        if not icon_path.exists() or icon_path.read_bytes() != icon_bytes:
            icon_path.write_bytes(icon_bytes)
        return icon_path

    def _configure_settings_fields(self) -> None:
        line_edit_stylesheet = """
QLineEdit {
    background-color: #240748;
    border: 1px solid #6f4df6;
    border-radius: 10px;
    padding: 8px 10px;
    color: #f4eeff;
    selection-background-color: #6f4df6;
    selection-color: #ffffff;
}

QLineEdit:focus {
    border: 1px solid #b79dff;
}
"""
        for line_edit in (
            self.tariff_go_line_edit,
            self.tariff_standard_line_edit,
            self.tariff_high_line_edit,
            self.base_price_line_edit,
            self.config_path_line_edit,
        ):
            line_edit.setStyleSheet(line_edit_stylesheet)

        settings_combo_stylesheet = """
QComboBox {
    background-color: #240748;
    color: #f4eeff;
    border: 1px solid #6f4df6;
    border-radius: 10px;
    padding: 8px 34px 8px 10px;
    selection-background-color: #6f4df6;
    selection-color: #ffffff;
}

QComboBox::drop-down {
    border: none;
    width: 28px;
    background-color: #240748;
    subcontrol-origin: padding;
    subcontrol-position: center right;
    border-top-right-radius: 10px;
    border-bottom-right-radius: 10px;
}

QComboBox::down-arrow {
    width: 0px;
    height: 0px;
    border-left: 6px solid #240748;
    border-right: 6px solid #240748;
    border-top: 8px solid #f4eeff;
}
"""
        self.tariff_type_combo.setStyleSheet(settings_combo_stylesheet)

    def _configure_primary_buttons(self) -> None:
        primary_button_stylesheet = """
QPushButton {
    background-color: #6f4df6;
    border: 1px solid #6f4df6;
    border-radius: 10px;
    padding: 9px 16px;
    color: #ffffff;
}

QPushButton:hover {
    background-color: #8d71ff;
    border-color: #8d71ff;
}

QPushButton:pressed {
    background-color: #5537cf;
    border-color: #5537cf;
}

QPushButton:disabled {
    background-color: #4d369f;
    border-color: #4d369f;
    color: #d8ccff;
}
"""
        for button in (self.fetch_data_button, self.save_settings_button):
            button.setStyleSheet(primary_button_stylesheet)

    def _configure_tooltip_palette(self) -> None:
        palette = QToolTip.palette()
        palette.setColor(QPalette.ColorRole.ToolTipBase, QColor("#240748"))
        palette.setColor(QPalette.ColorRole.ToolTipText, QColor("#ffffff"))
        palette.setColor(QPalette.ColorRole.Window, QColor("#240748"))
        palette.setColor(QPalette.ColorRole.WindowText, QColor("#ffffff"))
        QToolTip.setPalette(palette)

    def _configure_view_calendar_button(self) -> None:
        self.view_calendar_button.setText("")
        self.view_calendar_button.setIcon(_icon_from_base64(CALENDAR_ICON_PNG_B64))
        self.view_calendar_button.setIconSize(QSize(30, 30))
        self.view_calendar_button.setToolTip("Kalender öffnen")
        self.view_calendar_button.setAccessibleName("Kalender öffnen")

    def _copy_text_to_clipboard(self, text: str) -> None:
        QApplication.clipboard().setText(text)

    def _analysis_table_row_to_csv(self, row: int) -> str:
        output = io.StringIO()
        writer = csv.writer(output)
        values = []
        for column in range(self.analysis_table_model.columnCount()):
            item = self.analysis_table_model.item(row, column)
            values.append("" if item is None else item.text())
        writer.writerow(values)
        return output.getvalue().strip("\r\n")

    def _analysis_table_all_to_csv(self) -> str:
        output = io.StringIO()
        writer = csv.writer(output)
        headers = [
            self.analysis_table_model.headerData(column, Qt.Orientation.Horizontal)
            for column in range(self.analysis_table_model.columnCount())
        ]
        writer.writerow(headers)
        for row in range(self.analysis_table_model.rowCount()):
            writer.writerow(
                [
                    "" if self.analysis_table_model.item(row, column) is None else self.analysis_table_model.item(row, column).text()
                    for column in range(self.analysis_table_model.columnCount())
                ]
            )
        return output.getvalue()

    def _save_analysis_table_as_csv(self) -> None:
        default_path = self._ensure_output_suffix(
            self._get_default_output_path("csv").with_name("datenansicht_export.csv"),
            "csv",
        )
        filename, _ = QFileDialog.getSaveFileName(
            self.window,
            "Alle Werte als CSV speichern",
            str(default_path),
            "CSV-Dateien (*.csv)",
        )
        if not filename:
            return
        target = self._ensure_output_suffix(Path(filename), "csv")
        target.write_text(self._analysis_table_all_to_csv(), encoding="utf-8", newline="")
        self._set_status(f"Tabellenwerte als CSV gespeichert: {target}")

    def _show_analysis_table_context_menu(self, position) -> None:
        index = self.analysis_table_view.indexAt(position)
        if index.isValid():
            self.analysis_table_view.setCurrentIndex(index)

        current_index = self.analysis_table_view.currentIndex()
        has_current = current_index.isValid()
        has_rows = self.analysis_table_model.rowCount() > 0

        menu = QMenu(self.analysis_table_view)
        copy_value_action = menu.addAction("Wert kopieren")
        copy_row_action = menu.addAction("Zeile kopieren")
        copy_all_action = menu.addAction("Alle Werte kopieren")
        save_all_action = menu.addAction("Alle Werte in .csv speichern")

        copy_value_action.setEnabled(has_current)
        copy_row_action.setEnabled(has_current)
        copy_all_action.setEnabled(has_rows)
        save_all_action.setEnabled(has_rows)

        selected_action = menu.exec(self.analysis_table_view.viewport().mapToGlobal(position))
        if selected_action is None:
            return

        if selected_action is copy_value_action and has_current:
            self._copy_text_to_clipboard(current_index.data() or "")
        elif selected_action is copy_row_action and has_current:
            self._copy_text_to_clipboard(self._analysis_table_row_to_csv(current_index.row()))
        elif selected_action is copy_all_action and has_rows:
            self._copy_text_to_clipboard(self._analysis_table_all_to_csv())
        elif selected_action is save_all_action and has_rows:
            self._save_analysis_table_as_csv()

    def _set_initial_values(self) -> None:
        self.main_tab_widget.setCurrentIndex(0)
        self.analysis_content_tabs.setCurrentIndex(0)
        self.from_date_edit.setDate(QDate(2024, 1, 1))
        self.to_date_edit.setDate(QDate.currentDate())
        self.output_format_combo.setCurrentText("excel")
        self.auto_output_checkbox.setChecked(True)
        self.output_file_line_edit.setText(str(self._get_default_output_path("excel")))
        self.progress_bar.hide()
        self._toggle_password_visibility(False)
        self.tariff_type_combo.clear()
        self.tariff_type_combo.addItems([TARIFF_INTELLIGENT_GO, TARIFF_INTELLIGENT_HEAT])
        self._set_tariff_inputs(
            TARIFF_INTELLIGENT_GO,
            self.default_tariff_settings.get("tariff_go_ct", DEFAULT_TARIFF_GO_CT),
            self.default_tariff_settings.get("tariff_standard_ct", DEFAULT_TARIFF_STANDARD_CT),
            0.0,
            self.default_tariff_settings.get(
                "monthly_base_price_eur",
                DEFAULT_MONTHLY_BASE_PRICE_EUR,
            ),
        )
        self.view_mode_combo.setCurrentIndex(0)
        self.view_date_edit.setDate(QDate.currentDate())
        self.view_currency_checkbox.setChecked(False)
        self.config_path_line_edit.setText(str(get_app_config_folder()))
        self._configure_view_date_edit()
        self._refresh_analysis_view()

    def _connect_signals(self) -> None:
        self.show_password_checkbox.toggled.connect(self._toggle_password_visibility)
        self.output_format_combo.currentTextChanged.connect(self.on_format_changed)
        self.output_file_line_edit.editingFinished.connect(self._normalize_output_entry)
        self.browse_output_button.clicked.connect(self.browse_output_file)
        self.fetch_data_button.clicked.connect(self.get_data)
        self.tariff_type_combo.currentTextChanged.connect(self._on_tariff_type_changed)
        self.tariff_go_line_edit.editingFinished.connect(self._on_tariff_fields_edited)
        self.tariff_standard_line_edit.editingFinished.connect(self._on_tariff_fields_edited)
        self.tariff_high_line_edit.editingFinished.connect(self._on_tariff_fields_edited)
        self.base_price_line_edit.editingFinished.connect(self._on_tariff_fields_edited)
        self.save_settings_button.clicked.connect(self._save_settings_from_tab)
        self.view_mode_combo.currentIndexChanged.connect(self._on_view_mode_changed)
        self.view_calendar_button.clicked.connect(self._open_view_calendar_popup)
        self.view_date_edit.dateChanged.connect(lambda _date: self._refresh_analysis_view())
        self.view_currency_checkbox.toggled.connect(lambda _checked: self._refresh_analysis_view())
        self.view_previous_button.clicked.connect(lambda: self._shift_view_date(-1))
        self.view_next_button.clicked.connect(lambda: self._shift_view_date(1))

    def _set_window_icon(self) -> None:
        icon = QIcon()
        icon_dirs: list[Path] = []
        package_dir = Path(__file__).resolve().parent
        executable_dir = Path(sys.executable).resolve().parent
        for candidate in (package_dir, executable_dir):
            if candidate not in icon_dirs:
                icon_dirs.append(candidate)

        for filename in (
            "octopusdetool_gui.ico",
            "octopusdetool_gui-16.png",
            "octopusdetool_gui-32.png",
            "octopusdetool_gui-64.png",
            "octopusdetool_gui.svg",
        ):
            for directory in icon_dirs:
                icon_path = directory / filename
                if icon_path.exists():
                    icon.addFile(str(icon_path))

        if not icon.isNull():
            self.window.setWindowIcon(icon)

    def show(self) -> None:
        self.window.show()
        self.app.processEvents()
        self._fit_window_to_content(recenter=True)

    def _get_debug_log_path(self) -> Path:
        return get_app_config_folder() / "log.txt"

    def _content_height_for_window_width(self, window_width: int) -> int:
        content_layout = self.scroll_area_contents.layout()
        viewport_width = max(1, window_width - self._window_width_overhead())
        if content_layout.hasHeightForWidth():
            return content_layout.heightForWidth(viewport_width)
        return self.scroll_area_contents.sizeHint().height()

    def _window_width_overhead(self) -> int:
        return self.window.width() - self.scroll_area.viewport().width()

    def _window_height_overhead(self) -> int:
        return self.window.height() - self.scroll_area.viewport().height()

    def _fit_window_to_content(self, *, recenter: bool = False) -> None:
        screen = self.window.screen() or self.app.primaryScreen()
        if screen is None:
            return

        self.window.ensurePolished()
        self.scroll_area_contents.layout().activate()
        self.app.processEvents()

        available = screen.availableGeometry()
        max_width = int(available.width() * self.WINDOW_SCREEN_FRACTION)
        max_height = int(available.height() * self.WINDOW_SCREEN_FRACTION)
        width_overhead = self._window_width_overhead()
        height_overhead = self._window_height_overhead()

        current_width = max(self.window.width(), self.window.minimumWidth())
        required_height = self._content_height_for_window_width(current_width) + height_overhead
        target_width = current_width

        if required_height > max_height:
            best_width = current_width
            best_height = required_height
            candidate_width = current_width

            while candidate_width <= max_width:
                candidate_height = self._content_height_for_window_width(candidate_width) + height_overhead
                if candidate_height < best_height:
                    best_width = candidate_width
                    best_height = candidate_height
                if candidate_height <= max_height:
                    target_width = candidate_width
                    required_height = candidate_height
                    break
                candidate_width += self.RESIZE_STEP
            else:
                target_width = best_width
                required_height = best_height

        target_width = min(max(target_width, self.window.minimumWidth()), max_width)
        target_height = min(max(required_height, self.window.minimumHeight()), max_height)
        self.window.resize(target_width, target_height)

        if recenter:
            position_x = available.x() + max((available.width() - target_width) // 2, 0)
            position_y = available.y() + max((available.height() - target_height) // 3, 0)
            self.window.move(position_x, position_y)

    def _encrypt_config_value(self, value: str) -> str:
        if not value:
            return ""

        nonce = os.urandom(12)
        encrypted = AESGCM(CONFIG_AES_KEY).encrypt(nonce, value.encode("utf-8"), None)
        return base64.b64encode(nonce + encrypted).decode("ascii")

    def _decrypt_config_value(self, value: str) -> str:
        if not value:
            return ""

        raw = base64.b64decode(value)
        if len(raw) < 13:
            raise ValueError("Encrypted value is too short")

        nonce = raw[:12]
        ciphertext = raw[12:]
        plaintext = AESGCM(CONFIG_AES_KEY).decrypt(nonce, ciphertext, None)
        return plaintext.decode("utf-8")

    def _read_config_with_migration(self) -> tuple[dict, bool]:
        with open(CONFIG_FILE, "r", encoding="utf-8") as config_file:
            config = json.load(config_file)

        migrated = False
        encrypted_version = config.get("credential_encryption_version", 0)

        for field in CONFIG_ENCRYPTED_FIELDS:
            value = config.get(field, "")
            if not value:
                continue

            if encrypted_version >= CONFIG_ENCRYPTION_VERSION:
                config[field] = self._decrypt_config_value(value)
                continue

            config[field] = value
            migrated = True

        # Migrate old tariff type names to new ones
        old_tariff = config.get(TARIFF_TYPE_CONFIG_KEY, "")
        if old_tariff == "Intelligent Go":
            config[TARIFF_TYPE_CONFIG_KEY] = TARIFF_INTELLIGENT_GO
            migrated = True
        elif old_tariff == "Intelligent Heat":
            config[TARIFF_TYPE_CONFIG_KEY] = TARIFF_INTELLIGENT_HEAT
            migrated = True

        if migrated and CONFIG_WRITABLE:
            self._write_config(config)

        return config, migrated

    def _write_config(self, config: dict) -> None:
        if not CONFIG_WRITABLE:
            return
        config_to_save = dict(config)
        for field in CONFIG_ENCRYPTED_FIELDS:
            config_to_save[field] = self._encrypt_config_value(config.get(field, ""))
        config_to_save["credential_encryption_version"] = CONFIG_ENCRYPTION_VERSION

        with open(CONFIG_FILE, "w", encoding="utf-8") as config_file:
            json.dump(config_to_save, config_file, indent=2)

    def _serialize_tariff_rates(self) -> list[dict]:
        return [
            {
                "name": rate.name,
                "rate_ct": rate.rate_ct,
                "windows": [list(window) for window in rate.windows],
            }
            for rate in self.current_tariff_rates
        ]

    def _deserialize_tariff_rates(self, raw_rates) -> list[TariffRate]:
        if not isinstance(raw_rates, list):
            return []

        parsed_rates: list[TariffRate] = []
        for raw_rate in raw_rates:
            if not isinstance(raw_rate, dict):
                continue
            try:
                name = str(raw_rate["name"])
                rate_ct = float(raw_rate["rate_ct"])
            except (KeyError, TypeError, ValueError):
                continue

            windows_raw = raw_rate.get("windows", [])
            windows: list[tuple[str, str]] = []
            if isinstance(windows_raw, list):
                for window in windows_raw:
                    if isinstance(window, (list, tuple)) and len(window) == 2:
                        start, end = window
                        windows.append((str(start), str(end)))

            parsed_rates.append(
                TariffRate(
                    name=name,
                    rate_ct=rate_ct,
                    windows=tuple(windows),
                )
            )

        return parsed_rates

    def _persist_requested_tariff_display_name(self, display_name: str) -> None:
        if not CONFIG_WRITABLE:
            return
        if CONFIG_FILE.exists():
            try:
                config, _migrated = self._read_config_with_migration()
            except Exception:
                config = {}
        else:
            config = {}

        if config.get(LAST_TARIFF_DISPLAY_NAME_CONFIG_KEY) == display_name:
            return

        config[LAST_TARIFF_DISPLAY_NAME_CONFIG_KEY] = display_name
        self._write_config(config)

    def _persist_manual_tariff_selection(self, tariff_type: str) -> None:
        if not CONFIG_WRITABLE:
            return
        if CONFIG_FILE.exists():
            try:
                config, _migrated = self._read_config_with_migration()
            except Exception:
                config = {}
        else:
            config = {}

        if config.get(TARIFF_TYPE_CONFIG_KEY) == tariff_type:
            return

        config[TARIFF_TYPE_CONFIG_KEY] = tariff_type
        self._write_config(config)

    def _toggle_password_visibility(self, checked: bool) -> None:
        self.password_line_edit.setEchoMode(
            QLineEdit.EchoMode.Normal if checked else QLineEdit.EchoMode.Password
        )

    @contextmanager
    def _capture_debug_output(self):
        if not self.debug_checkbox.isChecked():
            yield None
            return

        log_path = self._get_debug_log_path()
        try:
            log_path.parent.mkdir(parents=True, exist_ok=True)
        except OSError:
            pass
        separator = "=" * 80

        with open(log_path, "a", encoding="utf-8") as log_file:
            log_file.write(
                f"\n{separator}\n"
                f"Debug session started: {datetime.now():%Y-%m-%d %H:%M:%S}\n"
                f"{separator}\n"
            )
            log_file.flush()

            tee_stream = _TeeStream(getattr(sys, "stdout", None), log_file)
            tee_error_stream = _TeeStream(getattr(sys, "stderr", None), log_file)

            with redirect_stdout(tee_stream), redirect_stderr(tee_error_stream):
                print(f"Debug log file: {log_path}")
                try:
                    yield log_path
                finally:
                    print(f"Debug session finished: {datetime.now():%Y-%m-%d %H:%M:%S}")

    def _set_status(self, message: str, update: bool = False) -> None:
        self.status_value_label.setText(message)
        if self.debug_checkbox.isChecked():
            print(f"[STATUS] {message}")
        if update:
            self.app.processEvents()
            if self.progress_bar.isVisible() and self.scroll_area.verticalScrollBar().maximum() > 0:
                self._fit_window_to_content()

    def _show_error(self, message: str) -> None:
        QMessageBox.critical(self.window, "Fehler", message)

    def _show_info(self, message: str) -> None:
        QMessageBox.information(self.window, "Erfolg", message)

    def _get_extension_for_format(self, format_type: str) -> str:
        return OUTPUT_EXTENSIONS.get(format_type, ".csv")

    def _get_default_output_path(self, format_type: str) -> Path:
        if format_type == "excel":
            return get_default_excel_path(self.current_tariff_type)
        return get_smartmeter_data_folder() / f"smartmeter_daten{self._get_extension_for_format(format_type)}"

    def _get_file_filter_for_format(self, format_type: str) -> str:
        filters = {
            "excel": "Excel-Dateien (*.xlsx);;Alle Dateien (*)",
            "csv": "CSV-Dateien (*.csv);;Alle Dateien (*)",
            "json": "JSON-Dateien (*.json);;Alle Dateien (*)",
            "yaml": "YAML-Dateien (*.yaml);;Alle Dateien (*)",
        }
        return filters.get(format_type, "Alle Dateien (*)")

    def _ensure_output_suffix(self, path: Path, format_type: str | None = None) -> Path:
        format_type = format_type or self.output_format_combo.currentText()
        target_suffix = self._get_extension_for_format(format_type)
        if path.suffix.lower() == target_suffix:
            return path
        return path.with_suffix(target_suffix)

    def _get_normalized_output_path(self, format_type: str | None = None) -> Path:
        format_type = format_type or self.output_format_combo.currentText()
        raw_value = self.output_file_line_edit.text().strip()
        if not raw_value:
            return self._get_default_output_path(format_type)
        return self._ensure_output_suffix(Path(raw_value).expanduser(), format_type)

    def _normalize_output_entry(self) -> None:
        self.output_file_line_edit.setText(str(self._get_normalized_output_path()))

    def _date_to_string(self, date_edit: QDateEdit) -> str:
        return date_edit.date().toString("dd.MM.yyyy")

    def _date_to_datetime(self, date_edit: QDateEdit) -> datetime:
        selected_date = date_edit.date()
        return datetime(selected_date.year(), selected_date.month(), selected_date.day())

    def _set_date_from_string(self, date_edit: QDateEdit, value: str, fallback: QDate) -> None:
        try:
            parsed = datetime.strptime(value, "%d.%m.%Y")
            date_edit.setDate(QDate(parsed.year, parsed.month, parsed.day))
        except ValueError:
            date_edit.setDate(fallback)

    def _set_tariff_inputs(
        self,
        tariff_type: str,
        tariff_go_ct: float,
        tariff_standard_ct: float,
        tariff_high_ct: float,
        base_price_eur: float,
    ) -> None:
        previous_tariff_type = self.current_tariff_type
        self.current_tariff_type = tariff_type
        self.tariff_type_combo.blockSignals(True)
        self.tariff_type_combo.setCurrentText(tariff_type)
        self.tariff_type_combo.blockSignals(False)
        self.tariff_go_line_edit.setText(f"{tariff_go_ct:.2f}")
        self.tariff_standard_line_edit.setText(f"{tariff_standard_ct:.2f}")
        self.tariff_high_line_edit.setText(f"{tariff_high_ct:.2f}")
        self.base_price_line_edit.setText(f"{base_price_eur:.2f}")
        self._apply_tariff_type_ui()
        self._sync_default_excel_output_path(previous_tariff_type)

    def _sync_default_excel_output_path(self, previous_tariff_type: str) -> None:
        if self.output_format_combo.currentText() != "excel":
            return

        current_output_path = self._get_normalized_output_path("excel")
        if current_output_path in {
            get_default_excel_path(previous_tariff_type),
            get_default_excel_path(TARIFF_INTELLIGENT_GO),
            get_default_excel_path(TARIFF_INTELLIGENT_HEAT),
        }:
            self.output_file_line_edit.setText(str(self._get_default_output_path("excel")))

    def _apply_tariff_type_ui(self) -> None:
        if self._apply_live_tariff_rate_ui():
            return

        is_heat = self.current_tariff_type == TARIFF_INTELLIGENT_HEAT
        if is_heat:
            self.tariff_go_label.setText("Tarif Niedrig 02:00-05:59, 12:00-15:59 (ct/kWh)")
            self.tariff_standard_label.setText(
                "Tarif Standard 06:00-11:59, 16:00-17:59, 21:00-01:59 (ct/kWh)"
            )
            self.tariff_high_label.setText("Tarif Hoch 18:00-20:59 (ct/kWh)")
        else:
            self.tariff_go_label.setText("Tarif Go 00:00-04:59 (ct/kWh)")
            self.tariff_standard_label.setText("Tarif Standard 05:00-23:59 (ct/kWh)")
            self.tariff_high_label.setText("Tarif Hoch (ct/kWh)")

        self.tariff_high_label.setVisible(is_heat)
        self.tariff_high_line_edit.setVisible(is_heat)

    def _on_tariff_type_changed(self, tariff_type: str) -> None:
        self.current_tariff_rates = []
        self._set_excel_export_support(True)
        defaults = get_default_tariff_settings_for_type(tariff_type)
        try:
            base_price_eur = self._parse_decimal_input(self.base_price_line_edit.text())
        except ValueError:
            base_price_eur = defaults.monthly_base_price_eur

        if not self._has_saved_base_price and not self.base_price_line_edit.text().strip():
            base_price_eur = defaults.monthly_base_price_eur

        self._set_tariff_inputs(
            tariff_type,
            defaults.low_ct,
            defaults.standard_ct,
            defaults.high_ct,
            base_price_eur,
        )
        self._refresh_analysis_view()

    def _parse_decimal_input(self, raw_value: str) -> float:
        cleaned = (
            raw_value.strip()
            .replace("ct/kWh", "")
            .replace("EUR", "")
            .replace("€", "")
            .replace(" ", "")
            .replace(",", ".")
        )
        if not cleaned:
            raise ValueError("Missing numeric input")
        return float(cleaned)

    def _get_config_decimal(self, config: dict, key: str, fallback: float) -> float:
        value = config.get(key, fallback)
        if isinstance(value, (int, float)):
            return float(value)
        if isinstance(value, str):
            try:
                return self._parse_decimal_input(value)
            except ValueError:
                return fallback
        return fallback

    def _get_tariff_values(self, *, show_error: bool) -> tuple[str, float, float, float, float] | None:
        tariff_type = self.tariff_type_combo.currentText() or TARIFF_INTELLIGENT_GO
        is_heat = tariff_type == TARIFF_INTELLIGENT_HEAT

        field_specs = [
            (
                self.tariff_go_line_edit,
                self.tariff_go_label.text(),
                self.default_tariff_settings.get(
                    "tariff_heat_low_ct" if is_heat else "tariff_go_ct",
                    DEFAULT_TARIFF_HEAT_LOW_CT if is_heat else DEFAULT_TARIFF_GO_CT,
                ),
            ),
            (
                self.tariff_standard_line_edit,
                self.tariff_standard_label.text(),
                self.default_tariff_settings.get(
                    "tariff_heat_standard_ct" if is_heat else "tariff_standard_ct",
                    DEFAULT_TARIFF_HEAT_STANDARD_CT if is_heat else DEFAULT_TARIFF_STANDARD_CT,
                ),
            ),
        ]
        if is_heat:
            field_specs.append(
                (
                    self.tariff_high_line_edit,
                    self.tariff_high_label.text(),
                    self.default_tariff_settings.get("tariff_heat_high_ct", DEFAULT_TARIFF_HEAT_HIGH_CT),
                )
            )
        field_specs.append(
            (
                self.base_price_line_edit,
                "Grundpreis pro Monat (EUR)",
                self.default_tariff_settings.get("monthly_base_price_eur", DEFAULT_MONTHLY_BASE_PRICE_EUR),
            )
        )

        parsed_values: list[float] = []
        for line_edit, label, fallback in field_specs:
            try:
                parsed_values.append(self._parse_decimal_input(line_edit.text()))
            except ValueError:
                if show_error:
                    self._show_error(f"Ungueltiger Wert fuer '{label}'. Bitte eine Zahl eingeben.")
                    line_edit.setFocus()
                    return None
                parsed_values.append(float(fallback))

        low_ct = parsed_values[0]
        standard_ct = parsed_values[1]
        high_ct = parsed_values[2] if is_heat else 0.0
        base_price_eur = parsed_values[3] if is_heat else parsed_values[2]
        return tariff_type, low_ct, standard_ct, high_ct, base_price_eur

    def _on_tariff_fields_edited(self) -> None:
        values = self._get_tariff_values(show_error=False)
        if values is None:
            return
        self.current_tariff_rates = []
        self._set_excel_export_support(True)
        self._set_tariff_inputs(*values)
        self._refresh_analysis_view()

    def _save_settings_from_tab(self) -> None:
        values = self._get_tariff_values(show_error=True)
        if values is None:
            return

        self.current_tariff_rates = []
        self._set_excel_export_support(True)
        self._set_tariff_inputs(*values)
        self._persist_manual_tariff_selection(values[0])
        self._refresh_analysis_view()
        self.save_config(force=True)

    def _resolve_tariff_type_from_display_name(self, display_name: str) -> str | None:
        normalized_display_name = display_name.strip().lower()
        if "lite" in normalized_display_name:
            return None
        if DISPLAY_NAME_DYNAMIC_OCTOPUS in normalized_display_name:
            return None
        if DISPLAY_NAME_OCTOPUS_GO in normalized_display_name:
            return TARIFF_INTELLIGENT_GO
        if DISPLAY_NAME_OCTOPUS_HEAT in normalized_display_name:
            return TARIFF_INTELLIGENT_HEAT
        return None

    def _set_excel_export_support(self, supported: bool, reason: str = "") -> None:
        self._excel_export_supported = supported
        self._excel_export_reason = reason

    def _format_rate_windows(self, rate: TariffRate) -> str:
        if not rate.windows:
            return ""
        return ", ".join(f"{start}-{end}" for start, end in rate.windows)

    def _format_rate_label(self, rate: TariffRate) -> str:
        windows = self._format_rate_windows(rate)
        if windows:
            return f"{rate.name} {windows} (ct/kWh)"
        return f"{rate.name} (ct/kWh)"

    def _apply_live_tariff_rate_ui(self) -> bool:
        if not self.current_tariff_rates:
            return False

        displayed_rates = self.current_tariff_rates[:3]
        labels = [
            self.tariff_go_label,
            self.tariff_standard_label,
            self.tariff_high_label,
        ]
        line_edits = [
            self.tariff_go_line_edit,
            self.tariff_standard_line_edit,
            self.tariff_high_line_edit,
        ]

        for index, rate in enumerate(displayed_rates):
            labels[index].setText(self._format_rate_label(rate))
            line_edits[index].setText(f"{rate.rate_ct:.2f}")

        has_third_rate = len(displayed_rates) >= 3
        self.tariff_high_label.setVisible(has_third_rate)
        self.tariff_high_line_edit.setVisible(has_third_rate)

        if len(self.current_tariff_rates) > 3:
            extra_count = len(self.current_tariff_rates) - 3
            self.tariff_high_label.setText(
                f"{self._format_rate_label(displayed_rates[2])} (+{extra_count} weitere nur im Diagramm)"
            )

        return True

    def _get_fixed_rate_labels(self, tariff_type: str) -> list[str]:
        labels = ["Tarif Niedrig" if tariff_type == TARIFF_INTELLIGENT_HEAT else "Tarif Go", "Tarif Standard"]
        if tariff_type == TARIFF_INTELLIGENT_HEAT:
            labels.append("Tarif Hoch")
        return labels

    def _get_analysis_rate_order(self, tariff_type: str) -> list[str]:
        if self.current_tariff_rates:
            return [rate.name for rate in self.current_tariff_rates]
        return self._get_fixed_rate_labels(tariff_type)

    def _get_analysis_rate_prices(self, tariff_type: str) -> dict[str, float]:
        tariff_values = self._get_tariff_values(show_error=False)
        if tariff_values is None:
            tariff_values = (
                tariff_type,
                DEFAULT_TARIFF_HEAT_LOW_CT if tariff_type == TARIFF_INTELLIGENT_HEAT else DEFAULT_TARIFF_GO_CT,
                DEFAULT_TARIFF_HEAT_STANDARD_CT if tariff_type == TARIFF_INTELLIGENT_HEAT else DEFAULT_TARIFF_STANDARD_CT,
                DEFAULT_TARIFF_HEAT_HIGH_CT if tariff_type == TARIFF_INTELLIGENT_HEAT else 0.0,
                DEFAULT_MONTHLY_BASE_PRICE_EUR,
            )

        _tariff_type, tariff_go_ct, tariff_standard_ct, tariff_high_ct, _monthly_base_price_eur = tariff_values
        if self.current_tariff_rates:
            return {rate.name: rate.rate_ct for rate in self.current_tariff_rates}

        labels = self._get_fixed_rate_labels(tariff_type)
        prices = {
            labels[0]: tariff_go_ct,
            labels[1]: tariff_standard_ct,
        }
        if tariff_type == TARIFF_INTELLIGENT_HEAT and len(labels) > 2:
            prices[labels[2]] = tariff_high_ct
        return prices

    def _reading_matches_window(self, reading_start: datetime, window: tuple[str, str]) -> bool:
        try:
            start_time = time.fromisoformat(window[0])
            end_time = time.fromisoformat(window[1])
        except ValueError:
            return False

        reading_time = normalize_datetime(reading_start).time().replace(second=0, microsecond=0)
        if start_time < end_time:
            return start_time <= reading_time < end_time
        return reading_time >= start_time or reading_time < end_time

    def _get_rate_name_for_reading(self, reading_start: datetime, tariff_type: str) -> str:
        if self.current_tariff_rates:
            for rate in self.current_tariff_rates:
                if any(self._reading_matches_window(reading_start, window) for window in rate.windows):
                    return rate.name

        fixed_labels = self._get_fixed_rate_labels(tariff_type)
        zone = classify_tariff_zone(reading_start, tariff_type)
        if zone == "low":
            return fixed_labels[0]
        if zone == "high" and len(fixed_labels) > 2:
            return fixed_labels[2]
        return fixed_labels[1]

    def _apply_tariff_agreement(
        self,
        agreement: TariffAgreement | None,
        api_tariff_settings=None,
        api_tariff_rates: list[TariffRate] | None = None,
    ) -> None:
        if agreement is None:
            self.current_tariff_display_name = "None"
            self.current_tariff_valid_from = ""
            self.current_tariff_valid_to = ""
            self.current_tariff_rates = []
            self._set_excel_export_support(True)
            self._persist_requested_tariff_display_name("None")
            self.save_config()
            return

        self.current_tariff_display_name = agreement.display_name
        self.current_tariff_valid_from = agreement.valid_from
        self.current_tariff_valid_to = agreement.valid_to or ""
        self.current_tariff_rates = list(api_tariff_rates or [])
        self._persist_requested_tariff_display_name(agreement.display_name)

        detected_type = self._resolve_tariff_type_from_display_name(agreement.display_name)
        if detected_type is None:
            self._set_excel_export_support(False, "Excel-Export wird fuer diesen Tarif aktuell nicht unterstuetzt.")
            QMessageBox.warning(
                self.window,
                "Tarif nicht unterstuetzt",
                f"Dieser Tarif wird aktuell noch nicht unterstuetzt: {agreement.display_name}",
            )
            return

        defaults = get_default_tariff_settings_for_type(detected_type)
        try:
            base_price = self._parse_decimal_input(self.base_price_line_edit.text())
        except ValueError:
            base_price = defaults.monthly_base_price_eur

        if (
            api_tariff_settings is not None
            and api_tariff_settings.tariff_type == detected_type
        ):
            defaults = api_tariff_settings
            base_price = api_tariff_settings.monthly_base_price_eur
            self._set_excel_export_support(True)
        elif self.current_tariff_rates:
            self._set_excel_export_support(
                False,
                "Excel-Export ist fuer die live abgerufene Tarifstruktur nicht kompatibel. "
                "Bitte CSV, JSON oder YAML verwenden.",
            )
        else:
            self._set_excel_export_support(True)

        self._set_tariff_inputs(
            detected_type,
            self.current_tariff_rates[0].rate_ct if self.current_tariff_rates else defaults.low_ct,
            self.current_tariff_rates[1].rate_ct if len(self.current_tariff_rates) > 1 else defaults.standard_ct,
            self.current_tariff_rates[2].rate_ct if len(self.current_tariff_rates) > 2 else defaults.high_ct,
            base_price,
        )
        self.save_config()

    def on_format_changed(self, _value: str | None = None) -> None:
        previous_format = getattr(self, "last_output_format", "excel")
        format_type = self.output_format_combo.currentText()
        current_path = self._get_normalized_output_path(previous_format)
        self.output_file_line_edit.setText(
            str(self._ensure_output_suffix(current_path, format_type))
        )
        self.last_output_format = format_type

    def browse_output_file(self) -> None:
        format_type = self.output_format_combo.currentText()
        current_path = self._get_normalized_output_path(format_type)
        title_map = {
            "excel": "Excel-Datei speichern unter",
            "csv": "CSV-Datei speichern unter",
            "json": "JSON-Datei speichern unter",
            "yaml": "YAML-Datei speichern unter",
        }
        filename, _ = QFileDialog.getSaveFileName(
            self.window,
            title_map.get(format_type, "Datei speichern unter"),
            str(current_path),
            self._get_file_filter_for_format(format_type),
        )
        if filename:
            self.output_file_line_edit.setText(
                str(self._ensure_output_suffix(Path(filename), format_type))
            )

    def _current_view_mode(self) -> str:
        mode_mapping = {
            "Tag": "day",
            "Woche": "week",
            "Monat": "month",
            "Jahr": "year",
        }
        return mode_mapping.get(self.view_mode_combo.currentText(), "day")

    def _configure_view_date_edit(self) -> None:
        mode = self._current_view_mode()
        display_format = {
            "day": "dd.MM.yyyy",
            "week": "dd.MM.yyyy",
            "month": "MM.yyyy",
            "year": "yyyy",
        }.get(mode, "dd.MM.yyyy")
        self.view_date_edit.setDisplayFormat(display_format)

        current_date = self.view_date_edit.date()
        normalized_date = current_date
        if mode == "month":
            normalized_date = QDate(current_date.year(), current_date.month(), 1)
        elif mode == "year":
            normalized_date = QDate(current_date.year(), 1, 1)

        if normalized_date != current_date:
            self.view_date_edit.blockSignals(True)
            self.view_date_edit.setDate(normalized_date)
            self.view_date_edit.blockSignals(False)

    def _on_view_mode_changed(self) -> None:
        self._configure_view_date_edit()
        self._refresh_analysis_view()

    def _open_view_calendar_popup(self) -> None:
        mode = self._current_view_mode()
        calendar = self.view_calendar_popup.calendar
        calendar.setSelectedDate(self.view_date_edit.date())
        calendar.setCurrentPage(self.view_date_edit.date().year(), self.view_date_edit.date().month())
        if mode == "week":
            calendar.setVerticalHeaderFormat(QCalendarWidget.VerticalHeaderFormat.ISOWeekNumbers)
        else:
            calendar.setVerticalHeaderFormat(QCalendarWidget.VerticalHeaderFormat.NoVerticalHeader)

        popup_pos = self.view_calendar_button.mapToGlobal(self.view_calendar_button.rect().bottomLeft())
        self.view_calendar_popup.adjustSize()
        self.view_calendar_popup.move(popup_pos)
        self.view_calendar_popup.show()
        self.view_calendar_popup.raise_()
        self.view_calendar_popup.activateWindow()

    def _on_view_calendar_date_selected(self, selected_date: QDate) -> None:
        mode = self._current_view_mode()
        normalized = selected_date
        if mode == "week":
            normalized = selected_date.addDays(1 - selected_date.dayOfWeek())
        elif mode == "month":
            normalized = QDate(selected_date.year(), selected_date.month(), 1)
        elif mode == "year":
            normalized = QDate(selected_date.year(), 1, 1)

        self.view_date_edit.setDate(normalized)
        self.view_calendar_popup.hide()

    def _shift_view_date(self, step: int) -> None:
        current_date = self.view_date_edit.date()
        mode = self._current_view_mode()
        if mode == "day":
            target_date = current_date.addDays(step)
        elif mode == "week":
            target_date = current_date.addDays(step * 7)
        elif mode == "month":
            target_date = current_date.addMonths(step)
        else:
            target_date = current_date.addYears(step)

        self.view_date_edit.setDate(target_date)

    def _qdate_to_date(self, qdate: QDate) -> date:
        return date(qdate.year(), qdate.month(), qdate.day())

    def _format_decimal(self, value: float, decimals: int) -> str:
        formatted = f"{value:,.{decimals}f}"
        return formatted.replace(",", "X").replace(".", ",").replace("X", ".")

    def _format_period_label(self, period_date: date) -> str:
        return f"{period_date.day}. {GERMAN_MONTH_NAMES[period_date.month - 1]} {period_date.year}"

    def _calculate_base_price_share(
        self,
        start_date: date,
        end_date: date,
        monthly_base_price_eur: float,
    ) -> float:
        total = 0.0
        cursor = date(start_date.year, start_date.month, 1)

        while cursor <= end_date:
            days_in_month = calendar.monthrange(cursor.year, cursor.month)[1]
            month_start = cursor
            month_end = date(cursor.year, cursor.month, days_in_month)
            overlap_start = max(start_date, month_start)
            overlap_end = min(end_date, month_end)
            if overlap_start <= overlap_end:
                covered_days = (overlap_end - overlap_start).days + 1
                total += monthly_base_price_eur * covered_days / days_in_month

            if cursor.month == 12:
                cursor = date(cursor.year + 1, 1, 1)
            else:
                cursor = date(cursor.year, cursor.month + 1, 1)

        return total

    def _set_default_analysis_date(self, *, force: bool = False) -> None:
        if not self.existing_data:
            return

        latest_start = max(normalize_datetime(reading["start"]) for reading in self.existing_data)
        target_date = QDate(latest_start.year, latest_start.month, latest_start.day)

        if force or not self._analysis_date_initialized:
            self.view_date_edit.blockSignals(True)
            self.view_date_edit.setDate(target_date)
            self.view_date_edit.blockSignals(False)
            self._analysis_date_initialized = True
            self._configure_view_date_edit()

    def _build_analysis_buckets(
        self,
        mode: str,
        selected_date: date,
    ) -> tuple[list[DisplayBucket], str, str, date, date]:
        if mode == "day":
            start_date = selected_date
            end_date = selected_date
            buckets = [
                DisplayBucket(
                    axis_label=f"{hour:02d}",
                    tooltip_label=f"{selected_date.strftime('%d.%m.%Y')} {hour:02d}:00",
                )
                for hour in range(24)
            ]
            title = self._format_period_label(selected_date)
            first_column_title = "Stunde"
        elif mode == "week":
            start_date = selected_date - timedelta(days=selected_date.weekday())
            end_date = start_date + timedelta(days=6)
            buckets = []
            for offset in range(7):
                current_day = start_date + timedelta(days=offset)
                buckets.append(
                    DisplayBucket(
                        axis_label=GERMAN_WEEKDAY_ABBR[offset],
                        tooltip_label=f"{GERMAN_WEEKDAY_NAMES[offset]}, {current_day.strftime('%d.%m.%Y')}",
                    )
                )
            iso_year, iso_week, _ = start_date.isocalendar()
            title = (
                f"Woche {iso_week}/{iso_year} "
                f"({start_date.strftime('%d.%m.%Y')} - {end_date.strftime('%d.%m.%Y')})"
            )
            first_column_title = "Tag"
        elif mode == "month":
            start_date = date(selected_date.year, selected_date.month, 1)
            days_in_month = calendar.monthrange(selected_date.year, selected_date.month)[1]
            end_date = date(selected_date.year, selected_date.month, days_in_month)
            buckets = []
            for day_index in range(days_in_month):
                current_day = start_date + timedelta(days=day_index)
                buckets.append(
                    DisplayBucket(
                        axis_label=f"{current_day.day:02d}",
                        tooltip_label=current_day.strftime("%d.%m.%Y"),
                    )
                )
            title = f"{GERMAN_MONTH_NAMES[selected_date.month - 1]} {selected_date.year}"
            first_column_title = "Tag"
        else:
            start_date = date(selected_date.year, 1, 1)
            end_date = date(selected_date.year, 12, 31)
            buckets = [
                DisplayBucket(
                    axis_label=GERMAN_MONTH_ABBR[month - 1],
                    tooltip_label=f"{GERMAN_MONTH_NAMES[month - 1]} {selected_date.year}",
                )
                for month in range(1, 13)
            ]
            title = str(selected_date.year)
            first_column_title = "Monat"

        start_dt = datetime(start_date.year, start_date.month, start_date.day)
        end_dt = datetime(end_date.year, end_date.month, end_date.day) + timedelta(days=1)

        for reading in self.existing_data:
            reading_start = normalize_datetime(reading["start"])
            if reading_start < start_dt or reading_start >= end_dt:
                continue

            if mode == "day":
                index = reading_start.hour
            elif mode in {"week", "month"}:
                index = (reading_start.date() - start_date).days
            else:
                index = reading_start.month - 1

            if not 0 <= index < len(buckets):
                continue

            rate_name = self._get_rate_name_for_reading(reading_start, self.current_tariff_type)
            buckets[index].rate_values_kwh[rate_name] = (
                buckets[index].rate_values_kwh.get(rate_name, 0.0)
                + float(reading["consumption_kwh"])
            )

        return buckets, title, first_column_title, start_date, end_date

    def _populate_analysis_table(
        self,
        buckets: list[DisplayBucket],
        first_column_title: str,
        *,
        show_currency: bool,
        tariff_type: str,
    ) -> None:
        self.analysis_table_model.clear()
        unit_title = "EUR" if show_currency else "kWh"
        rate_order = self._get_analysis_rate_order(tariff_type)
        rate_prices = self._get_analysis_rate_prices(tariff_type)
        headers = [first_column_title, *rate_order]
        headers.append(f"Gesamt ({unit_title})")
        self.analysis_table_model.setHorizontalHeaderLabels(headers)

        for bucket in buckets:
            if show_currency:
                rate_values = [
                    f"{self._format_decimal(bucket.rate_cost_eur(rate_name, rate_prices.get(rate_name, 0.0)), 2)} EUR"
                    for rate_name in rate_order
                ]
                total_value = f"{self._format_decimal(bucket.total_cost_eur(rate_prices), 2)} EUR"
            else:
                rate_values = [
                    f"{self._format_decimal(bucket.rate_kwh(rate_name), 3)} kWh"
                    for rate_name in rate_order
                ]
                total_value = f"{self._format_decimal(bucket.total_kwh, 3)} kWh"

            row_items = [QStandardItem(bucket.tooltip_label)]
            row_items.extend(QStandardItem(value) for value in rate_values)
            row_items.append(QStandardItem(total_value))

            for item in row_items[1:]:
                item.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)

            self.analysis_table_model.appendRow(row_items)

        header = self.analysis_table_view.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        for index in range(1, self.analysis_table_model.columnCount()):
            header.setSectionResizeMode(index, QHeaderView.ResizeMode.Stretch)

    def _refresh_analysis_view(self) -> None:
        tariff_values = self._get_tariff_values(show_error=False)
        if tariff_values is None:
            tariff_type = TARIFF_INTELLIGENT_GO
            tariff_go_ct = DEFAULT_TARIFF_GO_CT
            tariff_standard_ct = DEFAULT_TARIFF_STANDARD_CT
            tariff_high_ct = 0.0
            monthly_base_price_eur = DEFAULT_MONTHLY_BASE_PRICE_EUR
        else:
            tariff_type, tariff_go_ct, tariff_standard_ct, tariff_high_ct, monthly_base_price_eur = tariff_values

        selected_date = self._qdate_to_date(self.view_date_edit.date())
        mode = self._current_view_mode()
        show_currency = self.view_currency_checkbox.isChecked()
        buckets, title, first_column_title, start_date, end_date = self._build_analysis_buckets(
            mode,
            selected_date,
        )
        rate_order = self._get_analysis_rate_order(tariff_type)
        rate_prices = self._get_analysis_rate_prices(tariff_type)

        total_kwh = sum(bucket.total_kwh for bucket in buckets)
        variable_total_eur = sum(bucket.total_cost_eur(rate_prices) for bucket in buckets)
        has_readings = total_kwh > 0
        base_price_share = (
            self._calculate_base_price_share(start_date, end_date, monthly_base_price_eur)
            if has_readings
            else 0.0
        )
        base_price_label = {
            "day": "Grundpreis pro Tag",
            "week": "Grundpreis pro Woche",
            "month": "Grundpreis pro Monat",
            "year": "Grundpreis pro Jahr",
        }.get(mode, "Grundpreis")

        self.view_range_label.setText(title)
        if show_currency:
            self.view_total_caption_label.setText(
                "Gesamtkosten inkl. "
                f"{self._format_decimal(base_price_share, 2)} EUR {base_price_label}"
            )
            self.view_total_value_label.setText(
                f"{self._format_decimal(variable_total_eur + base_price_share, 2)} EUR"
            )
        else:
            self.view_total_caption_label.setText("Gesamtverbrauch")
            self.view_total_value_label.setText(f"{self._format_decimal(total_kwh, 3)} kWh")

        self.chart_view.update_buckets(
            buckets,
            show_currency=show_currency,
            rate_order=rate_order,
            rate_prices_ct=rate_prices,
            category_axis_title=first_column_title,
            value_axis_title="EUR" if show_currency else "kWh",
        )
        self._populate_analysis_table(
            buckets,
            first_column_title,
            show_currency=show_currency,
            tariff_type=tariff_type,
        )

    def load_config(self) -> None:
        if not CONFIG_FILE.exists():
            self._has_saved_base_price = False
            self._refresh_analysis_view()
            return

        try:
            config, migrated = self._read_config_with_migration()
            config_saving_enabled = config.get(CONFIG_SAVE_FLAG, "excel_file" in config)
            valid_formats = {"excel", "csv", "json", "yaml"}
            saved_format = config.get("output_format", "excel")
            if saved_format not in valid_formats:
                print(f"[DEBUG] Falsches Format in config: {saved_format}, verwende stattdessen Excel")
                saved_format = "excel"

            self.email_line_edit.setText(config.get("email", ""))
            self.password_line_edit.setText(config.get("password", ""))
            self.save_config_checkbox.setChecked(bool(config_saving_enabled))
            self.debug_checkbox.setChecked(bool(config.get("debug", False)))
            self.auto_output_checkbox.setChecked(bool(config.get(AUTO_OUTPUT_FLAG, True)))

            self.output_format_combo.blockSignals(True)
            self.output_format_combo.setCurrentText(saved_format)
            self.output_format_combo.blockSignals(False)
            self.last_output_format = saved_format
            saved_tariff_type = config.get(TARIFF_TYPE_CONFIG_KEY, TARIFF_INTELLIGENT_GO)
            saved_display_name = config.get(
                LAST_TARIFF_DISPLAY_NAME_CONFIG_KEY,
                config.get(LAST_TARIFF_CODE_CONFIG_KEY, "None"),
            )
            if saved_display_name not in {"", "None", None}:
                inferred_tariff_type = self._resolve_tariff_type_from_display_name(saved_display_name)
                if inferred_tariff_type is not None:
                    saved_tariff_type = inferred_tariff_type
            self.current_tariff_type = saved_tariff_type
            self.current_tariff_rates = self._deserialize_tariff_rates(config.get(TARIFF_RATES_CONFIG_KEY))
            self._excel_export_supported = bool(
                config.get(
                    EXCEL_EXPORT_SUPPORTED_CONFIG_KEY,
                    True if not self.current_tariff_rates else False,
                )
            )
            self._excel_export_reason = str(config.get(EXCEL_EXPORT_REASON_CONFIG_KEY, ""))

            if config_saving_enabled:
                saved_output_file = config.get("output_file") or config.get("excel_file")
                if saved_output_file:
                    self.output_file_line_edit.setText(saved_output_file)
                else:
                    self.output_file_line_edit.setText(str(self._get_default_output_path(saved_format)))
            else:
                self.output_file_line_edit.setText(str(self._get_default_output_path(saved_format)))

            self._set_date_from_string(self.from_date_edit, config.get("from_date", "01.01.2024"), QDate(2024, 1, 1))
            self.to_date_edit.setDate(QDate.currentDate())
            self._has_saved_base_price = "monthly_base_price_eur" in config
            self._set_tariff_inputs(
                saved_tariff_type,
                self._get_config_decimal(
                    config,
                    "tariff_heat_low_ct" if saved_tariff_type == TARIFF_INTELLIGENT_HEAT else "tariff_go_ct",
                    self.default_tariff_settings.get(
                        "tariff_heat_low_ct" if saved_tariff_type == TARIFF_INTELLIGENT_HEAT else "tariff_go_ct",
                        DEFAULT_TARIFF_HEAT_LOW_CT if saved_tariff_type == TARIFF_INTELLIGENT_HEAT else DEFAULT_TARIFF_GO_CT,
                    ),
                ),
                self._get_config_decimal(
                    config,
                    "tariff_heat_standard_ct" if saved_tariff_type == TARIFF_INTELLIGENT_HEAT else "tariff_standard_ct",
                    self.default_tariff_settings.get(
                        "tariff_heat_standard_ct" if saved_tariff_type == TARIFF_INTELLIGENT_HEAT else "tariff_standard_ct",
                        DEFAULT_TARIFF_HEAT_STANDARD_CT if saved_tariff_type == TARIFF_INTELLIGENT_HEAT else DEFAULT_TARIFF_STANDARD_CT,
                    ),
                ),
                self._get_config_decimal(
                    config,
                    "tariff_heat_high_ct",
                    self.default_tariff_settings.get("tariff_heat_high_ct", DEFAULT_TARIFF_HEAT_HIGH_CT),
                ) if saved_tariff_type == TARIFF_INTELLIGENT_HEAT else 0.0,
                self._get_config_decimal(
                    config,
                    "monthly_base_price_eur",
                    self.default_tariff_settings.get(
                        "monthly_base_price_eur",
                        DEFAULT_MONTHLY_BASE_PRICE_EUR,
                    ),
                ),
            )
            self.current_tariff_display_name = (
                saved_display_name if saved_display_name not in {None, ""} else "None"
            )

            self.on_format_changed(saved_format)
            self._refresh_analysis_view()

            if migrated:
                self._set_status("Konfiguration geladen und Zugangsdaten verschluesselt migriert")
            else:
                self._set_status("Konfiguration aus config.json geladen")
        except Exception as exc:
            self._has_saved_base_price = False
            self._set_status(f"Fehler beim Laden der Konfiguration: {exc}")

    def check_existing_data(self) -> None:
        try:
            self.existing_data = []
            self.latest_timestamp = None
            self.csv_path.parent.mkdir(parents=True, exist_ok=True)

            existing_data, latest_end = load_existing_consumption_data()
            if not existing_data:
                self._set_status("Keine Verbrauchsdaten gefunden. Bereit zum Abruf aller Daten.")
                self._refresh_analysis_view()
                return

            self.existing_data = existing_data
            self.latest_timestamp = latest_end

            self._set_default_analysis_date()
            self._refresh_analysis_view()

            today = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
            if self.latest_timestamp and self.latest_timestamp.date() >= (today - timedelta(days=1)).date():
                self._set_status(
                    f"Daten sind aktuell: {len(self.existing_data)} Eintraege bis {self.latest_timestamp.date()}."
                )
            else:
                self._set_status(
                    f"{len(self.existing_data)} Eintraege gefunden. Letzter: {self.latest_timestamp}. "
                    "Fehlende Daten werden abgerufen."
                )
        except Exception as exc:
            self._set_status(f"Fehler beim Lesen der Verbrauchsdaten: {exc}")

    def save_config(self, *, force: bool = False) -> None:
        if not CONFIG_WRITABLE:
            return
        if not force and not self.save_config_checkbox.isChecked():
            return

        tariff_values = self._get_tariff_values(show_error=False)
        if tariff_values is None:
            tariff_type = TARIFF_INTELLIGENT_GO
            tariff_go_ct = DEFAULT_TARIFF_GO_CT
            tariff_standard_ct = DEFAULT_TARIFF_STANDARD_CT
            tariff_high_ct = 0.0
            monthly_base_price_eur = DEFAULT_MONTHLY_BASE_PRICE_EUR
        else:
            tariff_type, tariff_go_ct, tariff_standard_ct, tariff_high_ct, monthly_base_price_eur = tariff_values
        config = {
            "email": self.email_line_edit.text(),
            "password": self.password_line_edit.text(),
            "output_format": self.output_format_combo.currentText(),
            "from_date": self._date_to_string(self.from_date_edit),
            "debug": self.debug_checkbox.isChecked(),
            AUTO_OUTPUT_FLAG: self.auto_output_checkbox.isChecked(),
            TARIFF_TYPE_CONFIG_KEY: tariff_type,
            LAST_TARIFF_DISPLAY_NAME_CONFIG_KEY: self.current_tariff_display_name or "None",
            "tariff_go_ct": tariff_go_ct,
            "tariff_standard_ct": tariff_standard_ct,
            "tariff_heat_low_ct": tariff_go_ct if tariff_type == TARIFF_INTELLIGENT_HEAT else DEFAULT_TARIFF_HEAT_LOW_CT,
            "tariff_heat_standard_ct": tariff_standard_ct if tariff_type == TARIFF_INTELLIGENT_HEAT else DEFAULT_TARIFF_HEAT_STANDARD_CT,
            "tariff_heat_high_ct": tariff_high_ct if tariff_type == TARIFF_INTELLIGENT_HEAT else DEFAULT_TARIFF_HEAT_HIGH_CT,
            "monthly_base_price_eur": monthly_base_price_eur,
            CONFIG_SAVE_FLAG: self.save_config_checkbox.isChecked(),
            "output_file": str(self._get_normalized_output_path()),
            "excel_file": str(self._get_normalized_output_path()),
        }
        config.pop(LAST_TARIFF_CODE_CONFIG_KEY, None)

        if self.current_tariff_rates:
            config[TARIFF_RATES_CONFIG_KEY] = self._serialize_tariff_rates()
            config[EXCEL_EXPORT_SUPPORTED_CONFIG_KEY] = self._excel_export_supported
            config[EXCEL_EXPORT_REASON_CONFIG_KEY] = self._excel_export_reason
        else:
            config.pop(TARIFF_RATES_CONFIG_KEY, None)
            config.pop(EXCEL_EXPORT_SUPPORTED_CONFIG_KEY, None)
            config.pop(EXCEL_EXPORT_REASON_CONFIG_KEY, None)

        try:
            self._write_config(config)
            self._has_saved_base_price = True
            self._set_status("Zugangsdaten speichern")
        except Exception as exc:
            self._set_status(f"Fehler beim Speichern der Zugangsdaten: {exc}")

    def validate_inputs(self) -> bool:
        if not self.email_line_edit.text().strip():
            self._show_error("E-Mail ist erforderlich!")
            return False

        if not self.password_line_edit.text():
            self._show_error("Passwort ist erforderlich!")
            return False

        if self.auto_output_checkbox.isChecked() and not self.output_file_line_edit.text().strip():
            self._show_error("Bitte waehlen Sie einen Dateinamen aus!")
            return False

        if self._get_tariff_values(show_error=True) is None:
            return False

        if (
            self.auto_output_checkbox.isChecked()
            and self.output_format_combo.currentText() == "excel"
            and not self._excel_export_supported
        ):
            self._show_error(self._excel_export_reason or "Excel-Export ist fuer diesen Tarif nicht verfuegbar.")
            return False

        self._normalize_output_entry()

        from_date = self._date_to_datetime(self.from_date_edit)
        to_date = self._date_to_datetime(self.to_date_edit)
        if from_date > to_date:
            self._show_error("Das Von-Datum muss vor oder gleich dem Bis-Datum sein!")
            return False

        return True

    def _write_csv_file(self, path: Path, readings: list[dict]) -> None:
        with open(path, "w", newline="", encoding="utf-8") as csv_file:
            writer = csv.writer(csv_file)
            writer.writerow(["start", "end", "consumption_kwh"])
            for reading in readings:
                writer.writerow(
                    [
                        format_datetime(reading["start"]),
                        format_datetime(reading["end"]),
                        reading["consumption_kwh"],
                    ]
                )

    def _expected_intervals_for_day(self, day: date) -> int:
        start_local = datetime(day.year, day.month, day.day, tzinfo=APP_TIMEZONE)
        end_local = datetime.combine(day + timedelta(days=1), time.min, tzinfo=APP_TIMEZONE)
        return int((end_local.astimezone(APP_TIMEZONE) - start_local).total_seconds() // 3600)

    def _get_incomplete_days(
        self,
        readings: list[dict],
        start_dt: datetime,
        end_dt: datetime,
    ) -> list[date]:
        if start_dt > end_dt:
            return []

        in_range_readings: list[datetime] = []
        for reading in readings:
            reading_start = normalize_datetime(reading["start"])
            if reading_start < start_dt or reading_start > end_dt:
                continue
            in_range_readings.append(reading_start)

        if not in_range_readings:
            return []

        counts: dict[date, set[datetime]] = {}
        cursor = min(in_range_readings).date()
        end_date = end_dt.date()
        while cursor <= end_date:
            counts[cursor] = set()
            cursor += timedelta(days=1)

        for reading_start in in_range_readings:
            counts.setdefault(reading_start.date(), set()).add(reading_start)

        incomplete_days: list[date] = []
        for day, intervals in counts.items():
            if len(intervals) < self._expected_intervals_for_day(day):
                incomplete_days.append(day)

        return incomplete_days

    def _start_progress(self) -> None:
        self.fetch_data_button.setEnabled(False)
        self.progress_bar.setRange(0, 0)
        self.progress_bar.show()
        self.app.processEvents()
        self._fit_window_to_content()

    def _stop_progress(self) -> None:
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        self.progress_bar.hide()
        self.fetch_data_button.setEnabled(True)
        self.app.processEvents()
        self._fit_window_to_content()

    def get_data(self) -> None:
        if not self.validate_inputs():
            return

        tariff_values = self._get_tariff_values(show_error=False)
        if tariff_values is None:
            tariff_go_ct = DEFAULT_TARIFF_GO_CT
            tariff_standard_ct = DEFAULT_TARIFF_STANDARD_CT
            monthly_base_price_eur = DEFAULT_MONTHLY_BASE_PRICE_EUR
        else:
            _tariff_type, tariff_go_ct, tariff_standard_ct, _tariff_high_ct, monthly_base_price_eur = tariff_values

        data_dir = get_smartmeter_data_folder()
        try:
            data_dir.mkdir(parents=True, exist_ok=True)
        except OSError:
            pass
        self.save_config()
        self._start_progress()

        try:
            with self._capture_debug_output():
                try:
                    period_from = self._date_to_datetime(self.from_date_edit)
                    period_to = self._date_to_datetime(self.to_date_edit)
                    period_to = period_to + timedelta(days=1) - timedelta(seconds=1)

                    today = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
                    yesterday_end = today - timedelta(seconds=1)

                    if period_to >= today:
                        period_to = yesterday_end

                    fetch_from = period_from
                    fetch_to = period_to
                    need_to_fetch = True

                    if self.latest_timestamp and self.latest_timestamp.date() >= (today - timedelta(days=1)).date():
                        self._set_status(
                            f"CSV ist bereits aktuell ({self.latest_timestamp.date()}). "
                            "Es werden keine Daten geladen.",
                            update=True,
                        )
                        need_to_fetch = False
                        fetch_from = None
                        fetch_to = None
                    elif self.latest_timestamp and self.latest_timestamp >= period_from:
                        fetch_from = self.latest_timestamp - timedelta(hours=1)
                        if fetch_from > yesterday_end:
                            fetch_from = yesterday_end
                            need_to_fetch = False
                        self._set_status(
                            f"Vorhandene Daten entdeckt. Lese ab {fetch_from}...",
                            update=True,
                        )

                    new_readings = []
                    client = None
                    account_number = None
                    property_id = None
                    malo_number = None

                    # Always authenticate and refresh tariff info when the user clicks the button.
                    self._set_status("Authentifizierung...", update=True)
                    client = OctopusGermanyClient(
                        self.email_line_edit.text(),
                        self.password_line_edit.text(),
                        debug=self.debug_checkbox.isChecked(),
                    )

                    if not client.authenticate():
                        if client.last_error_kind == "network":
                            raise Exception(
                                "Keine Internetverbindung erkannt oder Netzwerkfehler bei der Anmeldung. "
                                "Bitte Verbindung pruefen und erneut versuchen."
                            )
                        if client.last_error_kind == "response":
                            raise Exception(
                                "Der Octopus-Server hat unerwartet geantwortet. "
                                "Bitte spaeter erneut versuchen."
                            )
                        raise Exception(
                            "Authentifizierung fehlgeschlagen! Ueberpruefen Sie Ihre E-Mail und Ihr Passwort."
                        )

                    self._set_status("Kundennummer wird ermittelt...", update=True)
                    accounts = client.get_accounts_from_viewer()

                    if not accounts:
                        if client.last_error_kind == "network":
                            raise Exception(
                                "Keine Internetverbindung erkannt oder Netzwerkfehler beim Laden des Kontos. "
                                "Bitte Verbindung pruefen und erneut versuchen."
                            )
                        raise Exception("Kein Konto gefunden! Ueberpruefen Sie Ihre Zugangsdaten.")

                    if len(accounts) > 1:
                        account_list = "\n".join(
                            [f"  - {account.get('number', 'unknown')}" for account in accounts]
                        )
                        raise Exception(
                            f"Mehrere Konten gefunden ({len(accounts)}). Bitte waehlen Sie ein Konto aus:\n{account_list}"
                        )

                    account_number = accounts[0].get("number")
                    active_agreement = client.get_active_tariff_agreement(account_number)
                    api_tariff_settings = None
                    api_tariff_rates: list[TariffRate] = []
                    if active_agreement is not None:
                        api_tariff_rates = client.get_tariff_rates_for_agreement(active_agreement)
                        api_tariff_settings = client.get_tariff_settings_for_agreement(
                            active_agreement,
                        )
                    self._apply_tariff_agreement(
                        active_agreement,
                        api_tariff_settings,
                        api_tariff_rates,
                    )
                    tariff_values = self._get_tariff_values(show_error=False)
                    if tariff_values is None:
                        raise Exception("Tarifeinstellungen konnten nicht gelesen werden")
                    _tariff_type, tariff_go_ct, tariff_standard_ct, _tariff_high_ct, monthly_base_price_eur = tariff_values

                    if need_to_fetch:
                        self._set_status(
                            f"Kundennummer gefunden: {account_number}",
                            update=True,
                        )
                        self._set_status("Zaehler werden ermittelt...", update=True)

                        if client is None or account_number is None:
                            raise Exception("Kundendaten konnten nicht geladen werden")
                        meter_info = client.find_smart_meter(account_number)
                        if not meter_info:
                            if client.last_error_kind == "network":
                                raise Exception(
                                    "Keine Internetverbindung erkannt oder Netzwerkfehler beim Laden der Zaehlerdaten. "
                                    "Bitte Verbindung pruefen und erneut versuchen."
                                )
                            raise Exception(
                                "Kein Smart Meter fuer diesen Account gefunden!\n\n"
                                "Moegliche Gruende:\n"
                                "- Smart meter noch nicht eingerichtet\n"
                                "- Kein smart meter gefunden\n"
                                "- Falsche Kundennummer"
                            )

                        malo_number, _meter_id, property_id = meter_info
                        self._set_status(
                            f"Zaehler fuer MALO {malo_number} gefunden, Daten werden abgerufen...",
                            update=True,
                        )

                        def update_progress(count: int, page: int) -> None:
                            self._set_status(
                                f"Empfange Daten... {count} Eintraege (Seite {page})",
                                update=True,
                            )

                        new_readings = client.get_consumption_graphql(
                            property_id=property_id,
                            period_from=fetch_from,
                            period_to=fetch_to,
                            fetch_all=True,
                            progress_callback=update_progress,
                        )

                        if client.last_error_kind == "network":
                            raise Exception(
                                "Keine Internetverbindung erkannt oder Netzwerkfehler beim Abrufen der Verbrauchsdaten. "
                                "Bitte Verbindung pruefen und erneut versuchen."
                            )

                        if not new_readings and not self.existing_data:
                            raise Exception(
                                "Keine Verbrauchsdaten gefunden!\n\n"
                                "Moegliche Gruende:\n"
                                "- Smart Meter sendet noch keine Daten\n"
                                "- Keine Messwerte verfuegbar\n"
                                "- Zaehlerproblem - kontaktieren Sie Octopus"
                            )

                    all_readings = self.existing_data + new_readings
                    incomplete_days = self._get_incomplete_days(
                        all_readings,
                        normalize_datetime(period_from),
                        normalize_datetime(period_to),
                    )
                    if incomplete_days:
                        if client is None:
                            self._set_status("Authentifizierung fuer fehlende Tage...", update=True)
                            client = OctopusGermanyClient(
                                self.email_line_edit.text(),
                                self.password_line_edit.text(),
                                debug=self.debug_checkbox.isChecked(),
                            )
                            if not client.authenticate():
                                if client.last_error_kind == "network":
                                    raise Exception(
                                        "Keine Internetverbindung erkannt oder Netzwerkfehler bei der Anmeldung. "
                                        "Bitte Verbindung pruefen und erneut versuchen."
                                    )
                                raise Exception(
                                    "Authentifizierung fehlgeschlagen! Ueberpruefen Sie Ihre E-Mail und Ihr Passwort."
                                )

                        if account_number is None:
                            accounts = client.get_accounts_from_viewer()
                            if not accounts:
                                raise Exception("Kein Konto gefunden! Ueberpruefen Sie Ihre Zugangsdaten.")
                            account_number = accounts[0].get("number")

                        if property_id is None or malo_number is None:
                            meter_info = client.find_smart_meter(account_number)
                            if not meter_info:
                                raise Exception("Zaehlerdaten konnten fuer fehlende Tage nicht geladen werden.")
                            malo_number, _meter_id, property_id = meter_info

                        self._set_status(
                            f"Ergaenze {len(incomplete_days)} unvollstaendige Tage ueber GetSmartUsage...",
                            update=True,
                        )
                        fallback_readings: list[dict] = []
                        for missing_day in incomplete_days:
                            day_readings = client.get_smart_usage(
                                property_id=property_id,
                                market_supply_point_id=malo_number,
                                day=missing_day,
                            )
                            fallback_readings.extend(day_readings)
                        if fallback_readings:
                            all_readings.extend(fallback_readings)

                    if not all_readings:
                        raise Exception("Keine Daten zum Speichern!")

                    seen = {}
                    for reading in all_readings:
                        seen[reading["start"].isoformat()] = reading

                    unique_data = list(seen.values())
                    unique_data.sort(key=lambda item: normalize_datetime(item["start"]))

                    self.existing_data = unique_data
                    if unique_data:
                        self.latest_timestamp = max(
                            normalize_datetime(reading["end"]) for reading in unique_data
                        )

                    format_type = self.output_format_combo.currentText()
                    output_path = self._get_normalized_output_path().resolve()
                    try:
                        output_path.parent.mkdir(parents=True, exist_ok=True)
                    except OSError:
                        pass
                    self._set_status(
                        f"Speichere {len(unique_data)} Eintraege in consumption.csv...",
                        update=True,
                    )
                    self._write_csv_file(self.csv_path, unique_data)

                    # Always keep a JSON copy in the config folder
                    try:
                        json_config_path = get_app_config_folder() / "consumption.json"
                        save_to_json(unique_data, json_config_path)
                    except Exception:
                        pass

                    if self.auto_output_checkbox.isChecked():
                        if format_type == "excel":
                            if not self._excel_export_supported:
                                raise Exception(
                                    self._excel_export_reason
                                    or "Excel-Export ist fuer diesen Tarif nicht verfuegbar."
                                )
                            current_tariff_values = self._get_tariff_values(show_error=False)
                            if current_tariff_values is None:
                                raise Exception("Tarifeinstellungen konnten nicht gelesen werden")
                            current_tariff_type, tariff_go_ct, tariff_standard_ct, tariff_high_ct, monthly_base_price_eur = current_tariff_values

                            template_path = output_path
                            if output_path.exists():
                                existing_template_type = detect_excel_template_type(output_path)
                                if existing_template_type != current_tariff_type:
                                    raise Exception(
                                        "Die ausgewaehlte Excel-Datei passt nicht zum aktuellen Tarifmodell. "
                                        f"Datei: {existing_template_type}, aktueller Tarif: {current_tariff_type}."
                                    )
                            else:
                                template_path = ensure_excel_template(current_tariff_type)

                            self._set_status("Excel-Datei wird gefuellt...", update=True)
                            success = fill_excel_template(
                                unique_data,
                                str(template_path),
                                str(output_path),
                                tariff_go_ct=tariff_go_ct,
                                tariff_standard_ct=tariff_standard_ct,
                                tariff_high_ct=tariff_high_ct,
                                monthly_base_price_eur=monthly_base_price_eur,
                                tariff_type=current_tariff_type,
                            )
                            if not success:
                                raise Exception("Excel-Vorlage konnte nicht gefuellt werden")

                            self._show_info(
                                "Daten erfolgreich gespeichert!\n\n"
                                f"CSV: consumption.csv ({len(unique_data)} Eintraege)\n"
                                f"Excel: {output_path}"
                            )
                        elif format_type == "csv":
                            if output_path != self.csv_path.resolve():
                                self._set_status(
                                    f"Speichere {len(unique_data)} Eintraege als CSV...",
                                    update=True,
                                )
                                self._write_csv_file(output_path, unique_data)

                            self._show_info(
                                "Daten erfolgreich gespeichert!\n\n"
                                f"CSV: {output_path}\n"
                                f"Gesamteintraege: {len(unique_data)}"
                            )
                        elif format_type == "json":
                            self._set_status(
                                f"Speichere {len(unique_data)} Eintraege als JSON...",
                                update=True,
                            )
                            if not save_to_json(unique_data, output_path):
                                raise Exception("Fehler beim Speichern als JSON")

                            self._show_info(
                                "Daten erfolgreich gespeichert!\n\n"
                                f"JSON: {output_path}\n"
                                f"Gesamteintraege: {len(unique_data)}"
                            )
                        elif format_type == "yaml":
                            self._set_status(
                                f"Speichere {len(unique_data)} Eintraege als YAML...",
                                update=True,
                            )
                            if not save_to_yaml(unique_data, output_path):
                                raise Exception("Fehler beim Speichern als YAML")

                            self._show_info(
                                "Daten erfolgreich gespeichert!\n\n"
                                f"YAML: {output_path}\n"
                                f"Gesamteintraege: {len(unique_data)}"
                            )
                    else:
                        self._show_info(
                            "Daten erfolgreich gespeichert!\n\n"
                            f"CSV: consumption.csv ({len(unique_data)} Eintraege)\n"
                            "Automatische Ausgabe ist deaktiviert."
                        )

                    self._set_default_analysis_date(force=True)
                    self._refresh_analysis_view()
                    self._set_status(
                        f"Fertig! Daten gespeichert ({len(unique_data)} Eintraege)"
                    )
                except Exception:
                    if self.debug_checkbox.isChecked():
                        traceback.print_exc()
                    raise
        except Exception as exc:
            self._show_error(f"Ein Fehler ist aufgetreten:\n\n{exc}")
            self._set_status(f"Fehler: {exc}")
        finally:
            self._stop_progress()


def main() -> None:
    parser = argparse.ArgumentParser(add_help=True)
    parser.add_argument(
        "--debug",
        action="store_true",
        help="Print GraphQL requests and responses to the console and debug log",
    )
    args, qt_args = parser.parse_known_args(sys.argv[1:])

    app = QApplication.instance() or QApplication([sys.argv[0], *qt_args])
    if platform.system() == "Darwin":
        _configure_macos_qt_theme(app)
    app.setApplicationDisplayName("OctopusDETool")

    # Initialize app config folder once at program start
    ok, error = init_app_config_folder()
    if not ok:
        QMessageBox.warning(
            None,
            "Ordnerfehler",
            f"Der Anwendungsordner konnte nicht erstellt werden.\n"
            f"Einstellungen werden nicht gespeichert.\n\n{error}",
        )
    global CONFIG_WRITABLE
    CONFIG_WRITABLE = ok

    gui = OctopusSmartMeterGUI(app, debug_enabled=args.debug)
    gui.show()
    app.exec()


if __name__ == "__main__":
    main()
