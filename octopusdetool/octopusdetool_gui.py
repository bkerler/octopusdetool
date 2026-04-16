#!/usr/bin/env python3
"""
Octopus Energy Germany Smart Meter Data Logger - GUI Version

A PySide6-based GUI for fetching smart meter consumption data from
Octopus Energy Germany API and saving it to CSV, Excel, JSON, or YAML.
"""

from __future__ import annotations

import base64
import csv
import hashlib
import json
import os
import sys
import traceback
from contextlib import contextmanager, redirect_stderr, redirect_stdout
from datetime import datetime, timedelta
from pathlib import Path
from typing import TypeVar

from cryptography.hazmat.primitives.ciphers.aead import AESGCM
from PySide6.QtCore import QDate, QFile, QIODeviceBase, QObject
from PySide6.QtGui import QIcon
from PySide6.QtUiTools import QUiLoader
from PySide6.QtWidgets import (
    QApplication,
    QCheckBox,
    QComboBox,
    QDateEdit,
    QFileDialog,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QProgressBar,
    QScrollArea,
    QWidget,
)

from octopusdetool import (
    OctopusGermanyClient,
    ensure_excel_template,
    fill_excel_template,
    format_datetime,
    get_default_output_path,
    get_smartmeter_data_folder,
    normalize_datetime,
    save_to_json,
    save_to_yaml,
)


CONFIG_FILE = get_smartmeter_data_folder() / "config.json"
CONFIG_ENCRYPTION_VERSION = 1
CONFIG_ENCRYPTED_FIELDS = ("email", "password")
CONFIG_AES_KEY = hashlib.sha256(b"octopusdetool_rocks!").digest()
CONFIG_SAVE_FLAG = "save_config_enabled"
OUTPUT_EXTENSIONS = {
    "excel": ".xlsx",
    "csv": ".csv",
    "json": ".json",
    "yaml": ".yaml",
}
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


class OctopusSmartMeterGUI:
    WINDOW_SCREEN_FRACTION = 0.92
    RESIZE_STEP = 20

    def __init__(self, app: QApplication):
        self.app = app
        self.window = self._load_ui()
        self._bind_widgets()
        self._set_window_icon()

        ensure_excel_template()
        self.csv_path = get_default_output_path()
        self.existing_data: list[dict] = []
        self.latest_timestamp: datetime | None = None
        self.last_output_format = "excel"

        self._set_initial_values()
        self._connect_signals()
        self.load_config()
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
        self.scroll_area = self._find_widget(QScrollArea, "scrollArea")
        self.scroll_area_contents = self._find_widget(QWidget, "scrollAreaWidgetContents")
        self.email_line_edit = self._find_widget(QLineEdit, "emailLineEdit")
        self.password_line_edit = self._find_widget(QLineEdit, "passwordLineEdit")
        self.show_password_checkbox = self._find_widget(QCheckBox, "showPasswordCheckBox")
        self.save_config_checkbox = self._find_widget(QCheckBox, "saveConfigCheckBox")
        self.debug_checkbox = self._find_widget(QCheckBox, "debugCheckBox")
        self.output_format_combo = self._find_widget(QComboBox, "outputFormatComboBox")
        self.output_file_line_edit = self._find_widget(QLineEdit, "outputFileLineEdit")
        self.browse_output_button = self._find_widget(QPushButton, "browseOutputButton")
        self.from_date_edit = self._find_widget(QDateEdit, "fromDateEdit")
        self.to_date_edit = self._find_widget(QDateEdit, "toDateEdit")
        self.status_value_label = self._find_widget(QLabel, "statusValueLabel")
        self.progress_bar = self._find_widget(QProgressBar, "progressBar")
        self.fetch_data_button = self._find_widget(QPushButton, "fetchDataButton")

    def _set_initial_values(self) -> None:
        self.from_date_edit.setDate(QDate(2024, 1, 1))
        self.to_date_edit.setDate(QDate.currentDate())
        self.output_format_combo.setCurrentText("excel")
        self.output_file_line_edit.setText(str(self._get_default_output_path("excel")))
        self.progress_bar.hide()
        self._toggle_password_visibility(False)

    def _connect_signals(self) -> None:
        self.show_password_checkbox.toggled.connect(self._toggle_password_visibility)
        self.output_format_combo.currentTextChanged.connect(self.on_format_changed)
        self.output_file_line_edit.editingFinished.connect(self._normalize_output_entry)
        self.browse_output_button.clicked.connect(self.browse_output_file)
        self.fetch_data_button.clicked.connect(self.get_data)

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
        self._fit_window_to_content()

    def _get_debug_log_path(self) -> Path:
        return get_smartmeter_data_folder() / "log.txt"

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

    def _fit_window_to_content(self) -> None:
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

        if migrated:
            self._write_config(config)

        return config, migrated

    def _write_config(self, config: dict) -> None:
        config_to_save = dict(config)
        for field in CONFIG_ENCRYPTED_FIELDS:
            config_to_save[field] = self._encrypt_config_value(config.get(field, ""))
        config_to_save["credential_encryption_version"] = CONFIG_ENCRYPTION_VERSION

        with open(CONFIG_FILE, "w", encoding="utf-8") as config_file:
            json.dump(config_to_save, config_file, indent=2)

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
        log_path.parent.mkdir(parents=True, exist_ok=True)
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

    def _show_error(self, message: str) -> None:
        QMessageBox.critical(self.window, "Fehler", message)

    def _show_info(self, message: str) -> None:
        QMessageBox.information(self.window, "Erfolg", message)

    def _get_extension_for_format(self, format_type: str) -> str:
        return OUTPUT_EXTENSIONS.get(format_type, ".csv")

    def _get_default_output_path(self, format_type: str) -> Path:
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

    def load_config(self) -> None:
        get_smartmeter_data_folder().mkdir(parents=True, exist_ok=True)

        if not CONFIG_FILE.exists():
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

            self.output_format_combo.blockSignals(True)
            self.output_format_combo.setCurrentText(saved_format)
            self.output_format_combo.blockSignals(False)
            self.last_output_format = saved_format

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

            self.on_format_changed(saved_format)

            if migrated:
                self._set_status("Konfiguration geladen und Zugangsdaten verschlüsselt migriert")
            else:
                self._set_status("Konfiguration aus config.json geladen")
        except Exception as exc:
            self._set_status(f"Fehler beim Laden der Konfiguration: {exc}")

    def check_existing_data(self) -> None:
        try:
            self.existing_data = []
            self.latest_timestamp = None
            self.csv_path.parent.mkdir(parents=True, exist_ok=True)

            if not self.csv_path.exists():
                self._set_status("Keine consumption.csv gefunden. Bereit zum Abruf aller Daten.")
                return

            with open(self.csv_path, "r", newline="", encoding="utf-8") as csv_file:
                reader = csv.DictReader(csv_file)
                for row in reader:
                    try:
                        try:
                            start = datetime.strptime(row["start"], "%d.%m.%Y %H:%M:%S")
                            end = datetime.strptime(row["end"], "%d.%m.%Y %H:%M:%S")
                        except ValueError:
                            start = normalize_datetime(datetime.fromisoformat(row["start"]))
                            end = normalize_datetime(datetime.fromisoformat(row["end"]))

                        consumption = float(row["consumption_kwh"])
                        self.existing_data.append(
                            {
                                "start": start,
                                "end": end,
                                "consumption_kwh": consumption,
                            }
                        )

                        if self.latest_timestamp is None or end > self.latest_timestamp:
                            self.latest_timestamp = end
                    except Exception:
                        continue

            if not self.existing_data:
                self._set_status("Keine bestehenden Daten gefunden. Bereit zum Abruf.")
                return

            today = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
            if self.latest_timestamp and self.latest_timestamp.date() >= (today - timedelta(days=1)).date():
                self._set_status(
                    f"CSV ist aktuell: {len(self.existing_data)} Einträge bis {self.latest_timestamp.date()}."
                )
            else:
                self._set_status(
                    f"{len(self.existing_data)} Einträge gefunden. Letzter: {self.latest_timestamp}. "
                    "Fehlende Daten werden abgerufen."
                )
        except Exception as exc:
            self._set_status(f"Fehler beim Lesen der CSV: {exc}")

    def save_config(self) -> None:
        if not self.save_config_checkbox.isChecked():
            return

        get_smartmeter_data_folder().mkdir(parents=True, exist_ok=True)

        config = {
            "email": self.email_line_edit.text(),
            "password": self.password_line_edit.text(),
            "output_format": self.output_format_combo.currentText(),
            "from_date": self._date_to_string(self.from_date_edit),
            "debug": self.debug_checkbox.isChecked(),
            CONFIG_SAVE_FLAG: True,
            "output_file": str(self._get_normalized_output_path()),
            "excel_file": str(self._get_normalized_output_path()),
        }

        try:
            self._write_config(config)
            self._set_status("Konfiguration in config.json gespeichert")
        except Exception as exc:
            self._set_status(f"Fehler beim Speichern der Konfiguration: {exc}")

    def validate_inputs(self) -> bool:
        if not self.email_line_edit.text().strip():
            self._show_error("E-Mail ist erforderlich!")
            return False

        if not self.password_line_edit.text():
            self._show_error("Passwort ist erforderlich!")
            return False

        if not self.output_file_line_edit.text().strip():
            self._show_error("Bitte wählen Sie einen Dateinamen aus!")
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

    def _start_progress(self) -> None:
        self.fetch_data_button.setEnabled(False)
        self.progress_bar.setRange(0, 0)
        self.progress_bar.show()
        self.app.processEvents()

    def _stop_progress(self) -> None:
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        self.progress_bar.hide()
        self.fetch_data_button.setEnabled(True)

    def get_data(self) -> None:
        if not self.validate_inputs():
            return

        data_dir = get_smartmeter_data_folder()
        data_dir.mkdir(parents=True, exist_ok=True)
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

                    if need_to_fetch:
                        self._set_status("Authentifizierung...", update=True)
                        client = OctopusGermanyClient(
                            self.email_line_edit.text(),
                            self.password_line_edit.text(),
                            debug=self.debug_checkbox.isChecked(),
                        )

                        if not client.authenticate():
                            raise Exception(
                                "Authentifizierung fehlgeschlagen! Überprüfen Sie Ihre E-Mail und Ihr Passwort."
                            )

                        self._set_status("Kundennummer wird ermittelt...", update=True)
                        accounts = client.get_accounts_from_viewer()

                        if not accounts:
                            raise Exception("Kein Konto gefunden! Überprüfen Sie Ihre Zugangsdaten.")

                        if len(accounts) > 1:
                            account_list = "\n".join(
                                [f"  - {account.get('number', 'unknown')}" for account in accounts]
                            )
                            raise Exception(
                                f"Mehrere Konten gefunden ({len(accounts)}). Bitte wählen Sie ein Konto aus:\n{account_list}"
                            )

                        account_number = accounts[0].get("number")
                        self._set_status(
                            f"Kundennummer gefunden: {account_number}",
                            update=True,
                        )
                        self._set_status("Zähler werden ermittelt...", update=True)

                        meter_info = client.find_smart_meter(account_number)
                        if not meter_info:
                            raise Exception(
                                "Kein Smart Meter für diesen Account gefunden!\n\n"
                                "Mögliche Gründe:\n"
                                "- Smart meter noch nicht eingerichtet\n"
                                "- Kein smart meter gefunden\n"
                                "- Falsche Kundennummer"
                            )

                        malo_number, _meter_id, property_id = meter_info
                        self._set_status(
                            f"Zähler für MALO {malo_number} gefunden, Daten werden abgerufen...",
                            update=True,
                        )

                        def update_progress(count: int, page: int) -> None:
                            self._set_status(
                                f"Empfange Daten... {count} Einträge (Seite {page})",
                                update=True,
                            )

                        new_readings = client.get_consumption_graphql(
                            property_id=property_id,
                            period_from=fetch_from,
                            period_to=fetch_to,
                            fetch_all=True,
                            progress_callback=update_progress,
                        )

                        if not new_readings and not self.existing_data:
                            raise Exception(
                                "Keine Verbrauchsdaten gefunden!\n\n"
                                "Mögliche Gründe:\n"
                                "- Smart Meter sendet noch keine Daten\n"
                                "- Keine Messwerte verfügbar\n"
                                "- Zählerproblem - kontaktieren Sie Octopus"
                            )

                    all_readings = self.existing_data + new_readings
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
                    output_path.parent.mkdir(parents=True, exist_ok=True)

                    self._set_status(
                        f"Speichere {len(unique_data)} Einträge in consumption.csv...",
                        update=True,
                    )
                    self._write_csv_file(self.csv_path, unique_data)

                    if format_type == "excel":
                        self._set_status("Excel-Datei wird gefüllt...", update=True)
                        success = fill_excel_template(unique_data, str(output_path), str(output_path))
                        if not success:
                            raise Exception("Excel-Vorlage konnte nicht gefüllt werden")

                        self._show_info(
                            "Daten erfolgreich gespeichert!\n\n"
                            f"CSV: consumption.csv ({len(unique_data)} Einträge)\n"
                            f"Excel: {output_path}"
                        )
                    elif format_type == "csv":
                        if output_path != self.csv_path.resolve():
                            self._set_status(
                                f"Speichere {len(unique_data)} Einträge als CSV...",
                                update=True,
                            )
                            self._write_csv_file(output_path, unique_data)

                        self._show_info(
                            "Daten erfolgreich gespeichert!\n\n"
                            f"CSV: {output_path}\n"
                            f"Gesamteinträge: {len(unique_data)}"
                        )
                    elif format_type == "json":
                        self._set_status(
                            f"Speichere {len(unique_data)} Einträge als JSON...",
                            update=True,
                        )
                        if not save_to_json(unique_data, output_path):
                            raise Exception("Fehler beim Speichern als JSON")

                        self._show_info(
                            "Daten erfolgreich gespeichert!\n\n"
                            f"JSON: {output_path}\n"
                            f"Gesamteinträge: {len(unique_data)}"
                        )
                    elif format_type == "yaml":
                        self._set_status(
                            f"Speichere {len(unique_data)} Einträge als YAML...",
                            update=True,
                        )
                        if not save_to_yaml(unique_data, output_path):
                            raise Exception("Fehler beim Speichern als YAML")

                        self._show_info(
                            "Daten erfolgreich gespeichert!\n\n"
                            f"YAML: {output_path}\n"
                            f"Gesamteinträge: {len(unique_data)}"
                        )

                    self._set_status(
                        f"Fertig! Daten in Documents/smartmeter_data/ ({len(unique_data)} Einträge)"
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
    app = QApplication.instance() or QApplication(sys.argv)
    app.setApplicationDisplayName("OctopusDETool")
    gui = OctopusSmartMeterGUI(app)
    gui.show()
    app.exec()


if __name__ == "__main__":
    main()
