#!/usr/bin/env python3
"""
Octopus Energy Germany Smart Meter Data Logger - GUI Version

A tkinter-based GUI for fetching smart meter consumption data from 
Octopus Energy Germany API and saving it to CSV or Excel.
"""

import base64
import csv
import hashlib
import json
import os
import platform
import shutil
import sys
import traceback
from contextlib import contextmanager, redirect_stderr, redirect_stdout
from datetime import datetime, timedelta
from io import BytesIO
from pathlib import Path

from cryptography.hazmat.primitives.ciphers.aead import AESGCM


def _configure_windows_tk_runtime():
    """Point bundled Windows builds at the copied Tcl/Tk runtime."""
    if sys.platform != "win32":
        return

    executable_dir = Path(sys.executable).resolve().parent
    bundled_tcl_root = executable_dir / "tcl"
    bundled_tcl_lib = bundled_tcl_root / "tcl8.6"
    bundled_tk_lib = bundled_tcl_root / "tk8.6"

    if bundled_tcl_lib.exists():
        os.environ.setdefault("TCL_LIBRARY", str(bundled_tcl_lib))
    if bundled_tk_lib.exists():
        os.environ.setdefault("TK_LIBRARY", str(bundled_tk_lib))
    if bundled_tcl_root.exists() and hasattr(os, "add_dll_directory"):
        os.add_dll_directory(str(executable_dir))


_configure_windows_tk_runtime()

import tkinter as tk
import tkinter.font as tkfont
from tkinter import ttk, messagebox, filedialog

# Import from the same package
from octopusdetool import (
    OctopusGermanyClient, 
    fill_excel_template, 
    format_datetime,
    normalize_datetime,
    get_documents_folder,
    get_smartmeter_data_folder,
    ensure_excel_template,
    get_default_output_path,
    get_default_excel_path,
    save_to_json,
    save_to_yaml
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

# Embedded calendar icon (PNG, 32x32)
CALENDAR_ICON_BASE64 = "iVBORw0KGgoAAAANSUhEUgAAACAAAAAgCAYAAABzenr0AAAAuUlEQVR4nO1Wyw3FIAxLn7oTC8EGjMEGsFC2YYO+W1XR0ET9uQdygmIcy0lQp1rrQsD4IZMTEc0WUIxxXaeUbsWrDmzJpP1V/NT2gHbhjti6IpbAe7+uSymH521o+PZcLUGb7Cj5GbypCTWSK3hRgGTjU/HddyDnbCJgZnLOmb6HEHY4uANDwBDQnQJmNpP0sBaOrgBptHpJrGMoPXDwEsAFjB6AlwAuYPQAvATw3/LvOfB2wB2AC/gDw6NqeR/bFyoAAAAASUVORK5CYII="


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
    BASE_WINDOW_WIDTH = 620
    BASE_WINDOW_HEIGHT = 500
    WINDOW_SCREEN_FRACTION = 0.92
    WINDOW_EXTRA_WIDTH = 32
    WINDOW_EXTRA_HEIGHT = 20
    BASE_FONT_SIZE = 10
    HEADER_FONT_SIZE = 12
    STATUS_FONT_SIZE = 9
    BASE_PADDING = 12
    BASE_CALENDAR_WIDTH = 300
    BASE_CALENDAR_HEIGHT = 280
    MAX_UI_SCALE = 1.5
    COLOR_BG = "#f3f7fb"
    COLOR_SURFACE = "#ffffff"
    COLOR_SURFACE_ALT = "#eef4f9"
    COLOR_BORDER = "#d7e2ec"
    COLOR_TEXT = "#152230"
    COLOR_MUTED = "#61758a"
    COLOR_ACCENT = "#0f766e"
    COLOR_ACCENT_HOVER = "#115e59"
    COLOR_ACCENT_SOFT = "#d7f3ee"
    COLOR_WARNING = "#f59e0b"

    def __init__(self, root):
        self.root = root
        self.root.title("Octopus Energy Germany - Smartmeter Datenexport")
        self.ui_scale = self._detect_ui_scale()
        self.font_family = "Segoe UI"
        self.default_font = (self.font_family, self._scaled_font_size(self.BASE_FONT_SIZE))
        self.header_font = (self.font_family, self._scaled_font_size(self.HEADER_FONT_SIZE + 4), "bold")
        self.section_font = (self.font_family, self._scaled_font_size(self.HEADER_FONT_SIZE), "bold")
        self.status_font = (self.font_family, self._scaled_font_size(self.STATUS_FONT_SIZE))
        self.caption_font = (self.font_family, self._scaled_font_size(self.STATUS_FONT_SIZE))
        self.calendar_button_size = self._scaled(32)
        self._configure_window()
        self._set_window_icon()
        self.root.configure(background=self.COLOR_BG)
        
        # Style configuration
        self.style = ttk.Style()
        self._configure_styles()
        
        # Create a canvas with scrollbar for resizable content
        self.canvas = tk.Canvas(
            root,
            background=self.COLOR_BG,
            highlightthickness=0,
            borderwidth=0
        )
        self.scrollbar = ttk.Scrollbar(root, orient="vertical", command=self.canvas.yview, style="Vertical.TScrollbar")
        self.main_frame = ttk.Frame(self.canvas, padding=self._scaled(self.BASE_PADDING), style="App.TFrame")
        
        self.main_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )
        
        self.canvas_window = self.canvas.create_window((0, 0), window=self.main_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        self.canvas.bind("<Configure>", self._on_canvas_configure)
        
        self.canvas.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        self.scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        # Configure grid weights
        root.columnconfigure(0, weight=1)
        root.rowconfigure(0, weight=1)
        self.main_frame.columnconfigure(1, weight=1)
        
        # Mouse wheel scrolling
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)
        
        # Data storage
        self.existing_data = []
        self.latest_timestamp = None
        
        # Ensure Excel template exists and get paths
        ensure_excel_template()
        self.csv_path = get_default_output_path()
        self.excel_path = get_default_excel_path()
        self.last_output_format = "excel"
        
        self.create_widgets()
        self._fit_window_to_content()
        self.load_config()
        self.check_existing_data()
        self._fit_window_to_content()

    def _set_window_icon(self):
        """Set the application icon so Tk does not fall back to the default feather."""
        icon_dirs = []
        package_dir = Path(__file__).resolve().parent
        executable_dir = Path(sys.executable).resolve().parent
        for candidate in (package_dir, executable_dir):
            if candidate not in icon_dirs:
                icon_dirs.append(candidate)

        try:
            icon_ico = next((path / "octopusdetool_gui.ico" for path in icon_dirs if (path / "octopusdetool_gui.ico").exists()), None)
            png_paths = [
                path / filename
                for path in icon_dirs
                for filename in ("octopusdetool_gui-16.png", "octopusdetool_gui-32.png", "octopusdetool_gui-64.png")
                if (path / filename).exists()
            ]

            if sys.platform == "win32" and icon_ico is not None:
                self.root.iconbitmap(str(icon_ico))

            if png_paths:
                self.window_icons = [tk.PhotoImage(file=str(path)) for path in png_paths]
                self.root.iconphoto(True, *self.window_icons)
        except Exception as e:
            print(f"Warning: Could not set application icon: {e}")

    def _detect_ui_scale(self):
        """Scale the GUI based on actual display DPI when available."""
        try:
            pixels_per_inch = float(self.root.winfo_fpixels('1i'))
        except tk.TclError:
            return 1.0

        if pixels_per_inch <= 0:
            return 1.0

        # Tk uses 72 DPI as its logical baseline for scaling.
        dpi_scale = pixels_per_inch / 96.0
        if dpi_scale < 1.1:
            return 1.0
        return min(dpi_scale, self.MAX_UI_SCALE)

    def _scaled(self, value):
        return max(1, int(round(value * self.ui_scale)))

    def _scaled_font_size(self, value):
        return max(1, int(round(value * self.ui_scale)))

    def _configure_window(self):
        """Set a larger default window on high-resolution displays."""
        max_width, max_height = self._get_max_window_size()
        window_width = min(self._scaled(self.BASE_WINDOW_WIDTH), max_width)
        window_height = min(self._scaled(self.BASE_WINDOW_HEIGHT), max_height)
        min_width = min(self._scaled(self.BASE_WINDOW_WIDTH), window_width)
        min_height = min(self._scaled(self.BASE_WINDOW_HEIGHT), window_height)

        self._set_window_geometry(window_width, window_height)
        self.root.minsize(min_width, min_height)
        self.root.resizable(True, True)

    def _get_max_window_size(self):
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        return (
            int(screen_width * self.WINDOW_SCREEN_FRACTION),
            int(screen_height * self.WINDOW_SCREEN_FRACTION),
        )

    def _set_window_geometry(self, window_width, window_height):
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        position_x = max((screen_width - window_width) // 2, 0)
        position_y = max((screen_height - window_height) // 3, 0)
        self.root.geometry(f"{window_width}x{window_height}+{position_x}+{position_y}")

    def _fit_window_to_content(self):
        """Resize the initial window so the full form is visible when possible."""
        self.root.update_idletasks()

        content_width = (
            self.main_frame.winfo_reqwidth()
            + self.scrollbar.winfo_reqwidth()
            + self._scaled(self.WINDOW_EXTRA_WIDTH)
        )
        content_height = self.main_frame.winfo_reqheight() + self._scaled(self.WINDOW_EXTRA_HEIGHT)

        max_width, max_height = self._get_max_window_size()
        window_width = min(max(self.root.winfo_width(), content_width), max_width)
        window_height = min(max(self.root.winfo_height(), content_height), max_height)

        self._set_window_geometry(window_width, window_height)
        min_width = min(self._scaled(self.BASE_WINDOW_WIDTH), max_width)
        min_height = min(self._scaled(self.BASE_WINDOW_HEIGHT), max_height)
        self.root.minsize(min_width, min_height)

    def _configure_styles(self):
        """Apply a more modern visual system across ttk widgets."""
        try:
            self.style.theme_use("clam")
        except tk.TclError:
            pass

        named_font_overrides = {
            "TkDefaultFont": self.default_font,
            "TkTextFont": self.default_font,
            "TkMenuFont": self.default_font,
            "TkHeadingFont": self.section_font,
            "TkCaptionFont": self.default_font,
        }
        for font_name, font_config in named_font_overrides.items():
            try:
                tkfont.nametofont(font_name).configure(
                    family=font_config[0],
                    size=font_config[1],
                    weight="bold" if "bold" in font_config else "normal",
                    slant="italic" if "italic" in font_config else "roman",
                )
            except tk.TclError:
                continue

        self.style.configure(".", background=self.COLOR_BG, foreground=self.COLOR_TEXT, font=self.default_font)
        self.style.configure("App.TFrame", background=self.COLOR_BG)
        self.style.configure("Surface.TFrame", background=self.COLOR_SURFACE)
        self.style.configure("Hero.TFrame", background=self.COLOR_SURFACE)
        self.style.configure("TFrame", background=self.COLOR_BG)
        self.style.configure("TLabel", background=self.COLOR_BG, foreground=self.COLOR_TEXT, font=self.default_font)
        self.style.configure("Surface.TLabel", background=self.COLOR_SURFACE, foreground=self.COLOR_TEXT, font=self.default_font)
        self.style.configure("Muted.TLabel", background=self.COLOR_SURFACE, foreground=self.COLOR_MUTED, font=self.caption_font)
        self.style.configure("HeroBadge.TLabel", background=self.COLOR_ACCENT_SOFT, foreground=self.COLOR_ACCENT, font=self.caption_font)
        self.style.configure("Header.TLabel", background=self.COLOR_SURFACE, foreground=self.COLOR_TEXT, font=self.header_font)
        self.style.configure("Section.TLabel", background=self.COLOR_SURFACE, foreground=self.COLOR_TEXT, font=self.section_font)
        self.style.configure("Status.TLabel", background=self.COLOR_SURFACE_ALT, foreground=self.COLOR_ACCENT, font=self.status_font)

        self.style.configure(
            "TButton",
            font=self.default_font,
            padding=(14, 9),
            borderwidth=0,
            relief="flat",
            background=self.COLOR_SURFACE_ALT,
            foreground=self.COLOR_TEXT,
        )
        self.style.map(
            "TButton",
            background=[("active", self.COLOR_BORDER)],
            foreground=[("disabled", self.COLOR_MUTED)],
        )
        self.style.configure(
            "Primary.TButton",
            background=self.COLOR_ACCENT,
            foreground="#ffffff",
            padding=(18, 11),
            font=self.section_font,
        )
        self.style.map(
            "Primary.TButton",
            background=[("active", self.COLOR_ACCENT_HOVER), ("disabled", self.COLOR_BORDER)],
            foreground=[("disabled", "#ffffff")],
        )

        field_padding = (10, 8)
        self.style.configure(
            "TEntry",
            fieldbackground=self.COLOR_SURFACE,
            foreground=self.COLOR_TEXT,
            bordercolor=self.COLOR_BORDER,
            lightcolor=self.COLOR_BORDER,
            darkcolor=self.COLOR_BORDER,
            insertcolor=self.COLOR_TEXT,
            padding=field_padding,
        )
        self.style.configure(
            "TCombobox",
            fieldbackground=self.COLOR_SURFACE,
            foreground=self.COLOR_TEXT,
            bordercolor=self.COLOR_BORDER,
            lightcolor=self.COLOR_BORDER,
            darkcolor=self.COLOR_BORDER,
            arrowsize=self._scaled(14),
            padding=field_padding,
        )
        self.style.map(
            "TCombobox",
            fieldbackground=[("readonly", self.COLOR_SURFACE)],
            background=[("readonly", self.COLOR_SURFACE)],
            foreground=[("readonly", self.COLOR_TEXT)],
        )
        self.style.configure("TSpinbox", arrowsize=self._scaled(14), padding=(8, 6))
        self.style.configure(
            "TCheckbutton",
            background=self.COLOR_SURFACE,
            foreground=self.COLOR_TEXT,
            font=self.default_font,
        )
        self.style.map(
            "TCheckbutton",
            background=[("active", self.COLOR_SURFACE)],
            foreground=[("disabled", self.COLOR_MUTED)],
        )

        self.style.configure(
            "Card.TLabelframe",
            background=self.COLOR_SURFACE,
            bordercolor=self.COLOR_BORDER,
            lightcolor=self.COLOR_BORDER,
            darkcolor=self.COLOR_BORDER,
            relief="solid",
            borderwidth=1,
        )
        self.style.configure(
            "Card.TLabelframe.Label",
            background=self.COLOR_SURFACE,
            foreground=self.COLOR_TEXT,
            font=self.section_font,
        )
        self.style.configure(
            "TSeparator",
            background=self.COLOR_BORDER,
        )
        self.style.configure(
            "Vertical.TScrollbar",
            background=self.COLOR_SURFACE_ALT,
            troughcolor=self.COLOR_BG,
            bordercolor=self.COLOR_BG,
            arrowcolor=self.COLOR_MUTED,
        )
        self.style.configure(
            "Modern.Horizontal.TProgressbar",
            troughcolor=self.COLOR_SURFACE_ALT,
            background=self.COLOR_ACCENT,
            bordercolor=self.COLOR_SURFACE_ALT,
            lightcolor=self.COLOR_ACCENT,
            darkcolor=self.COLOR_ACCENT,
        )
    
    def _on_mousewheel(self, event):
        self.canvas.yview_scroll(int(-1*(event.delta/120)), "units")

    def _on_canvas_configure(self, event):
        self.canvas.itemconfigure(self.canvas_window, width=event.width)

    def _get_debug_log_path(self):
        return get_smartmeter_data_folder() / "log.txt"

    def _encrypt_config_value(self, value):
        """Encrypt config values with AES-256-GCM and store them as base64."""
        if not value:
            return ""

        nonce = os.urandom(12)
        encrypted = AESGCM(CONFIG_AES_KEY).encrypt(nonce, value.encode("utf-8"), None)
        return base64.b64encode(nonce + encrypted).decode("ascii")

    def _decrypt_config_value(self, value):
        """Decrypt AES-256-GCM config values stored as base64."""
        if not value:
            return ""

        raw = base64.b64decode(value)
        if len(raw) < 13:
            raise ValueError("Encrypted value is too short")

        nonce = raw[:12]
        ciphertext = raw[12:]
        plaintext = AESGCM(CONFIG_AES_KEY).decrypt(nonce, ciphertext, None)
        return plaintext.decode("utf-8")

    def _read_config_with_migration(self):
        """Load config and upgrade plaintext credentials to encrypted storage."""
        with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
            config = json.load(f)

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

    def _write_config(self, config):
        """Persist config with encrypted credentials."""
        config_to_save = dict(config)
        for field in CONFIG_ENCRYPTED_FIELDS:
            config_to_save[field] = self._encrypt_config_value(config.get(field, ""))
        config_to_save["credential_encryption_version"] = CONFIG_ENCRYPTION_VERSION

        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(config_to_save, f, indent=2)

    def _toggle_password_visibility(self):
        self.password_entry.config(show="" if self.show_password_var.get() else "*")

    @contextmanager
    def _capture_debug_output(self):
        """Persist GUI debug output to log.txt while the fetch is running."""
        if not self.debug_var.get():
            yield None
            return

        log_path = self._get_debug_log_path()
        log_path.parent.mkdir(parents=True, exist_ok=True)
        separator = "=" * 80

        with open(log_path, 'a', encoding='utf-8') as log_file:
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

    def _set_status(self, message, update=False):
        self.status_var.set(message)
        if self.debug_var.get():
            print(f"[STATUS] {message}")
        if update:
            self.root.update()
    
    def create_widgets(self):
        """Create all GUI widgets."""
        row = 0
        pad_small = self._scaled(4)
        pad_medium = self._scaled(8)
        pad_large = self._scaled(14)
        pad_xlarge = self._scaled(18)

        hero_frame = ttk.Frame(self.main_frame, style="Hero.TFrame", padding=(pad_large, pad_large, pad_large, pad_xlarge))
        hero_frame.grid(row=row, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, pad_medium))
        hero_frame.columnconfigure(0, weight=1)

        ttk.Label(
            hero_frame,
            text="Octopus Energy Deutschland Smartmeter Datenexport",
            style="Header.TLabel"
        ).grid(row=1, column=0, sticky=tk.W)
        ttk.Label(
            hero_frame,
            text="(c) B.Kerler / S. Stahl 2026",
            style="Muted.TLabel",
            wraplength=self._scaled(560),
            justify=tk.LEFT
        ).grid(row=2, column=0, sticky=tk.W, pady=(pad_small, 0))
        row += 1

        account_frame = ttk.LabelFrame(
            self.main_frame,
            text="Zugang & Einstellungen",
            style="Card.TLabelframe",
            padding=pad_large
        )
        account_frame.grid(row=row, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, pad_small))
        account_frame.columnconfigure(1, weight=1)
        row += 1

        account_row = 0
        ttk.Label(account_frame, text="E-Mail:", style="Surface.TLabel").grid(
            row=account_row, column=0, sticky=tk.W, pady=pad_small
        )
        self.email_var = tk.StringVar()
        self.email_entry = ttk.Entry(account_frame, textvariable=self.email_var, width=60)
        self.email_entry.grid(row=account_row, column=1, sticky=(tk.W, tk.E), pady=pad_small, padx=(pad_medium, 0))
        self.email_entry.config(state='normal')
        account_row += 1

        ttk.Label(account_frame, text="Passwort:", style="Surface.TLabel").grid(
            row=account_row, column=0, sticky=tk.W, pady=pad_small
        )
        self.password_var = tk.StringVar()
        self.password_entry = ttk.Entry(
            account_frame, textvariable=self.password_var, width=60, show="*"
        )
        self.password_entry.grid(row=account_row, column=1, sticky=(tk.W, tk.E), pady=pad_small, padx=(pad_medium, 0))
        self.password_entry.config(state='normal')

        self.show_password_var = tk.BooleanVar(value=False)
        self.show_password_checkbox = ttk.Checkbutton(
            account_frame,
            text="Passwort anzeigen",
            variable=self.show_password_var,
            command=self._toggle_password_visibility
        )
        self.show_password_checkbox.grid(row=account_row, column=2, sticky=tk.W, pady=pad_small, padx=(pad_medium, 0))
        account_row += 1

        self.save_config_var = tk.BooleanVar(value=False)
        self.save_config_checkbox = ttk.Checkbutton(
            account_frame, text="Konfiguration in config.json speichern",
            variable=self.save_config_var
        )
        self.save_config_checkbox.grid(row=account_row, column=0, columnspan=3, sticky=tk.W, pady=(pad_small, 0))
        account_row += 1

        self.debug_var = tk.BooleanVar(value=False)
        self.debug_check = ttk.Checkbutton(
            account_frame,
            text="Debug-Ausgabe aktivieren (zeigt alle API-Anfragen, wird in Dokumente/smartmeter_daten/log.txt gespeichert)",
            variable=self.debug_var
        )
        self.debug_check.grid(row=account_row, column=0, columnspan=3, sticky=tk.W, pady=(pad_small, 0))

        output_frame = ttk.LabelFrame(
            self.main_frame,
            text="Ausgabeoptionen",
            style="Card.TLabelframe",
            padding=pad_large
        )
        output_frame.grid(row=row, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, pad_small))
        output_frame.columnconfigure(1, weight=1)
        row += 1

        ttk.Label(output_frame, text="Format:", style="Surface.TLabel").grid(row=0, column=0, sticky=tk.W, pady=pad_small)
        self.output_format_var = tk.StringVar(value="excel")

        format_combo = ttk.Combobox(
            output_frame,
            textvariable=self.output_format_var,
            values=["excel", "csv", "json", "yaml"],
            state="readonly",
            width=20
        )
        format_combo.grid(row=0, column=1, sticky=(tk.W, tk.E), pady=pad_small, padx=(pad_medium, 0))
        format_combo.bind("<<ComboboxSelected>>", self.on_format_changed)

        ttk.Label(
            output_frame,
            text="Dateiname:",
            style="Surface.TLabel"
        ).grid(row=1, column=0, sticky=tk.W, pady=pad_small)

        output_file_frame = ttk.Frame(output_frame, style="Surface.TFrame")
        output_file_frame.grid(row=1, column=1, sticky=(tk.W, tk.E), pady=pad_small, padx=(pad_medium, 0))
        output_file_frame.columnconfigure(0, weight=1)

        self.output_file_var = tk.StringVar(value=str(self._get_default_output_path("excel")))
        self.output_file_entry = ttk.Entry(output_file_frame, textvariable=self.output_file_var, width=50)
        self.output_file_entry.grid(row=0, column=0, sticky=(tk.W, tk.E))
        self.output_file_entry.bind("<FocusOut>", self._normalize_output_entry)
        self.output_file_entry.bind("<Return>", self._normalize_output_entry)

        self.browse_btn = ttk.Button(
            output_file_frame, text="Speichern unter", command=self.browse_output_file, width=14
        )
        self.browse_btn.grid(row=0, column=1, padx=(pad_medium, 0))

        ttk.Label(
            output_frame,
            text="Der Dateiname wird automatisch auf das gewählte Format abgestimmt.",
            style="Muted.TLabel",
            wraplength=self._scaled(520),
            justify=tk.LEFT
        ).grid(row=2, column=0, columnspan=2, sticky=tk.W, pady=(2, 0))

        self.date_frame = ttk.LabelFrame(
            self.main_frame,
            text="Datumsbereich",
            style="Card.TLabelframe",
            padding=pad_large
        )
        self.date_frame.grid(row=row, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, pad_small))
        self.date_frame.columnconfigure(1, weight=1)
        self.date_frame.columnconfigure(3, weight=1)
        row += 1

        ttk.Label(self.date_frame, text="Von:", style="Surface.TLabel").grid(row=0, column=0, sticky=tk.W, padx=(0, pad_small))
        self.from_date_var = tk.StringVar(value="01.01.2024")
        self.from_date_entry = ttk.Entry(self.date_frame, textvariable=self.from_date_var, width=12)
        self.from_date_entry.grid(row=0, column=1, sticky=tk.W, padx=0)

        try:
            icon_data = base64.b64decode(CALENDAR_ICON_BASE64)
            icon_image = tk.PhotoImage(data=icon_data)
            if self.ui_scale >= 1.4:
                icon_image = icon_image.zoom(2, 2)
            self.calendar_icon = icon_image
            self.calendar_button_size = max(self.calendar_button_size, icon_image.width())
        except Exception as e:
            print(f"Fehler: Konnte das Kalender icon nicht laden: {e}")
            self.calendar_icon = None
        
        self.from_calendar_btn = tk.Button(
            self.date_frame,
            image=self.calendar_icon,
            width=self.calendar_button_size,
            height=self.calendar_button_size,
            relief=tk.FLAT,
            borderwidth=0,
            highlightthickness=0,
            bg=self.COLOR_SURFACE_ALT,
            activebackground=self.COLOR_ACCENT_SOFT,
            cursor="hand2",
            command=lambda: self.show_calendar(self.from_date_var)
        )
        self.from_calendar_btn.grid(row=0, column=2, sticky=tk.W, padx=(0, pad_medium))

        ttk.Label(self.date_frame, text="Bis:", style="Surface.TLabel").grid(row=0, column=3, sticky=tk.W, padx=(pad_medium, pad_small))
        yesterday = (datetime.now() - timedelta(days=1)).strftime("%d.%m.%Y")
        self.to_date_var = tk.StringVar(value=yesterday)
        self.to_date_entry = ttk.Entry(self.date_frame, textvariable=self.to_date_var, width=12)
        self.to_date_entry.grid(row=0, column=4, sticky=tk.W, padx=0)
        self.to_calendar_btn = tk.Button(
            self.date_frame,
            image=self.calendar_icon,
            width=self.calendar_button_size,
            height=self.calendar_button_size,
            relief=tk.FLAT,
            borderwidth=0,
            highlightthickness=0,
            bg=self.COLOR_SURFACE_ALT,
            activebackground=self.COLOR_ACCENT_SOFT,
            cursor="hand2",
            command=lambda: self.show_calendar(self.to_date_var)
        )
        self.to_calendar_btn.grid(row=0, column=5, sticky=tk.W, padx=0)

        self.progress_var = tk.DoubleVar(value=0)
        self.progress_bar = ttk.Progressbar(
            self.main_frame,
            variable=self.progress_var,
            maximum=100,
            mode='indeterminate',
            style='Modern.Horizontal.TProgressbar'
        )

        status_frame = ttk.Frame(self.main_frame, style="Surface.TFrame", padding=(pad_large, pad_medium))
        status_frame.grid(row=row, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, pad_small))
        status_frame.columnconfigure(0, weight=1)
        self.status_var = tk.StringVar(value="Bereit")
        self.status_label = ttk.Label(
            status_frame,
            textvariable=self.status_var,
            style='Status.TLabel'
        )
        self.status_label.grid(row=0, column=0, sticky=(tk.W, tk.E))
        row += 1

        self.get_data_btn = ttk.Button(
            self.main_frame,
            text="Daten vom Server abrufen",
            command=self.get_data,
            width=25,
            style="Primary.TButton"
        )
        self.get_data_btn.grid(row=row, column=2, sticky=tk.E, pady=0)

        self.on_format_changed()
    
    def on_format_changed(self, event=None):
        """Handle output format change."""
        previous_format = getattr(self, "last_output_format", "excel")
        format_type = self.output_format_var.get()
        current_path = self._get_normalized_output_path(previous_format)
        self.output_file_var.set(str(self._ensure_output_suffix(current_path, format_type)))
        self.last_output_format = format_type

        self.output_file_entry.config(state='normal')
        self.browse_btn.config(state='normal')
    
    def browse_output_file(self):
        """Open a save dialog for the selected output file."""
        format_type = self.output_format_var.get()
        current_path = self._get_normalized_output_path(format_type)
        title_map = {
            "excel": "Excel-Datei speichern unter",
            "csv": "CSV-Datei speichern unter",
            "json": "JSON-Datei speichern unter",
            "yaml": "YAML-Datei speichern unter",
        }
        filename = filedialog.asksaveasfilename(
            title=title_map.get(format_type, "Datei speichern unter"),
            defaultextension=self._get_extension_for_format(format_type),
            initialdir=str(current_path.parent),
            initialfile=current_path.name,
            filetypes=self._get_filetypes_for_format(format_type)
        )
        if filename:
            self.output_file_var.set(str(self._ensure_output_suffix(Path(filename), format_type)))

    def _get_extension_for_format(self, format_type):
        return OUTPUT_EXTENSIONS.get(format_type, ".csv")

    def _get_default_output_path(self, format_type):
        return get_smartmeter_data_folder() / f"smartmeter_daten{self._get_extension_for_format(format_type)}"

    def _get_filetypes_for_format(self, format_type):
        filetypes = {
            "excel": [("Excel-Dateien", "*.xlsx")],
            "csv": [("CSV-Dateien", "*.csv")],
            "json": [("JSON-Dateien", "*.json")],
            "yaml": [("YAML-Dateien", "*.yaml")],
        }
        return filetypes.get(format_type, [("Alle Dateien", "*.*")]) + [("Alle Dateien", "*.*")]

    def _ensure_output_suffix(self, path, format_type=None):
        format_type = format_type or self.output_format_var.get()
        target_suffix = self._get_extension_for_format(format_type)
        if path.suffix.lower() == target_suffix:
            return path
        return path.with_suffix(target_suffix)

    def _get_normalized_output_path(self, format_type=None):
        format_type = format_type or self.output_format_var.get()
        raw_value = self.output_file_var.get().strip()
        if not raw_value:
            return self._get_default_output_path(format_type)
        return self._ensure_output_suffix(Path(raw_value).expanduser(), format_type)

    def _normalize_output_entry(self, event=None):
        self.output_file_var.set(str(self._get_normalized_output_path()))
    
    def show_calendar(self, target_var):
        """Show a simple calendar dialog."""
        top = tk.Toplevel(self.root)
        top.configure(background=self.COLOR_BG)
        top.title("Datum auswählen")
        calendar_width = self._scaled(self.BASE_CALENDAR_WIDTH)
        calendar_height = self._scaled(self.BASE_CALENDAR_HEIGHT)
        top.geometry(f"{calendar_width}x{calendar_height}")
        top.minsize(calendar_width, calendar_height)
        top.transient(self.root)
        top.grab_set()
        
        # Parse current date (European format: DD.MM.YYYY)
        try:
            current_date = datetime.strptime(target_var.get(), "%d.%m.%Y")
        except:
            current_date = datetime.now()
        
        selected_year = tk.IntVar(value=current_date.year)
        selected_month = tk.IntVar(value=current_date.month)
        selected_day = tk.IntVar(value=current_date.day)
        
        # Year and Month selection
        header_frame = ttk.Frame(top, style="App.TFrame")
        header_frame.pack(pady=self._scaled(10))
        
        ttk.Spinbox(
            header_frame, from_=2020, to=2030, width=6,
            textvariable=selected_year
        ).pack(side=tk.LEFT, padx=self._scaled(5))
        
        months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
                  "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
        month_combo = ttk.Combobox(
            header_frame, values=months, width=6, state='readonly'
        )
        month_combo.set(months[current_date.month - 1])
        month_combo.pack(side=tk.LEFT, padx=self._scaled(5))
        
        # Calendar frame
        cal_frame = ttk.Frame(top, style="App.TFrame")
        cal_frame.pack(pady=self._scaled(10))
        
        # Day buttons frame
        days_frame = ttk.Frame(cal_frame, style="App.TFrame")
        days_frame.pack()
        
        def select_day(day):
            month_idx = months.index(month_combo.get()) + 1
            # European format: DD.MM.YYYY
            date_str = f"{day:02d}.{month_idx:02d}.{selected_year.get():04d}"
            target_var.set(date_str)
            top.destroy()
        
        def update_calendar():
            # Clear existing buttons
            for widget in days_frame.winfo_children():
                widget.destroy()
            
            # Get days in month
            import calendar
            year = selected_year.get()
            month = months.index(month_combo.get()) + 1
            _, days_in_month = calendar.monthrange(year, month)
            
            # Day labels
            for i, day_name in enumerate(["Mo", "Tu", "We", "Th", "Fr", "Sa", "Su"]):
                ttk.Label(days_frame, text=day_name, width=4, style="Muted.TLabel").grid(row=0, column=i)
            
            # Day buttons
            first_weekday, _ = calendar.monthrange(year, month)
            day = 1
            for week in range(1, 7):
                for weekday in range(7):
                    if week == 1 and weekday < first_weekday:
                        ttk.Label(days_frame, text="", width=4).grid(row=week, column=weekday)
                    elif day <= days_in_month:
                        btn = tk.Button(
                            days_frame,
                            text=str(day),
                            width=4,
                            font=self.default_font,
                            relief=tk.FLAT,
                            borderwidth=0,
                            highlightthickness=0,
                            bg=self.COLOR_SURFACE,
                            fg=self.COLOR_TEXT,
                            activebackground=self.COLOR_ACCENT_SOFT,
                            activeforeground=self.COLOR_TEXT,
                            cursor="hand2",
                            command=lambda d=day: select_day(d)
                        )
                        if (day == current_date.day and 
                            month == current_date.month and 
                            year == current_date.year):
                            btn.config(bg=self.COLOR_WARNING, fg="#ffffff", activebackground=self.COLOR_WARNING)
                        btn.grid(row=week, column=weekday)
                        day += 1
        
        # Update button
        update_btn = ttk.Button(top, text="Kalender aktualisieren", command=update_calendar)
        update_btn.pack(pady=self._scaled(5))
        
        # Initial calendar display
        update_calendar()
    
    def load_config(self):
        """Load configuration from config.json."""
        # Ensure smartmeter_data folder exists
        get_smartmeter_data_folder().mkdir(parents=True, exist_ok=True)
        
        if CONFIG_FILE.exists():
            try:
                config, migrated = self._read_config_with_migration()
                config_saving_enabled = config.get(CONFIG_SAVE_FLAG, 'excel_file' in config)
                
                self.email_var.set(config.get('email', ''))
                self.password_var.set(config.get('password', ''))
                if config_saving_enabled:
                    saved_output_file = config.get('output_file') or config.get('excel_file')
                    if saved_output_file:
                        self.output_file_var.set(saved_output_file)
                    else:
                        self.output_file_var.set(str(self._get_default_output_path('excel')))
                    self.save_config_var.set(True)
                else:
                    self.output_file_var.set(str(self._get_default_output_path('excel')))
                    self.save_config_var.set(False)
                # Validate output format, default to excel if invalid
                valid_formats = ['excel', 'csv', 'json', 'yaml']
                saved_format = config.get('output_format', 'excel')
                if saved_format not in valid_formats:
                    print(f"[DEBUG] Falsches Format in config: {saved_format}, verwende stattdessen Excel")
                    saved_format = 'excel'
                self.output_format_var.set(saved_format)
                self.last_output_format = saved_format
                # Default from date: 01.01.2024 or from config
                default_from = config.get('from_date', '01.01.2024')
                self.from_date_var.set(default_from)
                # To date is always yesterday (last complete day)
                yesterday = (datetime.now() - timedelta(days=1)).strftime("%d.%m.%Y")
                self.to_date_var.set(yesterday)
                self.debug_var.set(config.get('debug', False))
                
                # Update UI based on loaded format
                self.on_format_changed()
                
                if migrated:
                    self.status_var.set("Konfiguration geladen und Zugangsdaten verschlüsselt migriert")
                else:
                    self.status_var.set("Konfiguration aus config.json geladen")
            except Exception as e:
                self.status_var.set(f"Fehler beim Laden der Konfiguration: {e}")
    
    def check_existing_data(self):
        """Read existing CSV and update status."""
        try:
            # Create directory if needed
            self.csv_path.parent.mkdir(parents=True, exist_ok=True)
            
            if self.csv_path.exists():
                with open(self.csv_path, 'r', newline='', encoding='utf-8') as f:
                    reader = csv.DictReader(f)
                    for row in reader:
                        try:
                            # Try European format first (DD.MM.YYYY HH:MM:SS), then ISO
                            try:
                                start = datetime.strptime(row['start'], "%d.%m.%Y %H:%M:%S")
                                end = datetime.strptime(row['end'], "%d.%m.%Y %H:%M:%S")
                            except ValueError:
                                # Fallback to ISO format for backwards compatibility
                                start = normalize_datetime(datetime.fromisoformat(row['start']))
                                end = normalize_datetime(datetime.fromisoformat(row['end']))
                            consumption = float(row['consumption_kwh'])
                            
                            self.existing_data.append({
                                'start': start,
                                'end': end,
                                'consumption_kwh': consumption
                            })
                            
                            if self.latest_timestamp is None or end > self.latest_timestamp:
                                self.latest_timestamp = end
                        except:
                            continue
                
                if self.existing_data:
                    today = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
                    if self.latest_timestamp and self.latest_timestamp.date() >= (today - timedelta(days=1)).date():
                        self.status_var.set(f"CSV ist aktuell: {len(self.existing_data)} Einträge bis {self.latest_timestamp.date()}.")
                    else:
                        self.status_var.set(f"{len(self.existing_data)} Einträge gefunden. Letzter: {self.latest_timestamp}. Fehlende Daten werden abgerufen.")
                else:
                    self.status_var.set("Keine bestehenden Daten gefunden. Bereit zum Abruf.")
            else:
                self.status_var.set("Keine consumption.csv gefunden. Bereit zum Abruf aller Daten.")
        except Exception as e:
            self.status_var.set(f"Fehler beim Lesen der CSV: {e}")
    
    def save_config(self):
        """Save configuration to config.json."""
        if not self.save_config_var.get():
            return
        
        # Ensure smartmeter_data folder exists
        get_smartmeter_data_folder().mkdir(parents=True, exist_ok=True)
        
        config = {
            'email': self.email_var.get(),
            'password': self.password_var.get(),
            'output_format': self.output_format_var.get(),
            'from_date': self.from_date_var.get(),  # Store the from date
            'debug': self.debug_var.get(),
            CONFIG_SAVE_FLAG: True,
        }

        if self.save_config_var.get():
            normalized_output_path = str(self._get_normalized_output_path())
            config['output_file'] = normalized_output_path
            config['excel_file'] = normalized_output_path
        
        try:
            self._write_config(config)
            self.status_var.set("Konfiguration in config.json gespeichert")
        except Exception as e:
            self.status_var.set(f"Fehler beim Speichern der Konfiguration: {e}")
    
    def validate_inputs(self):
        """Validate user inputs."""
        if not self.email_var.get():
            messagebox.showerror("Fehler", "E-Mail ist erforderlich!")
            return False
        if not self.password_var.get():
            messagebox.showerror("Fehler", "Passwort ist erforderlich!")
            return False
        # Kundennummer ist optional - wird automatisch ermittelt wenn nicht angegeben
        
        format_type = self.output_format_var.get()
        if format_type == "excel" and not self.excel_var.get():
            messagebox.showerror("Fehler", "Bitte wählen Sie eine Excel-Datei aus!")
            return False
        if format_type == "excel":
            self._normalize_excel_entry()
        
        try:
            from_date = datetime.strptime(self.from_date_var.get(), "%d.%m.%Y")
            to_date = datetime.strptime(self.to_date_var.get(), "%d.%m.%Y")
            if from_date > to_date:
                messagebox.showerror("Fehler", "Das Von-Datum muss vor oder gleich dem Bis-Datum sein!")
                return False
        except ValueError:
            messagebox.showerror("Fehler", "Ungültiges Datumsformat! Verwenden Sie TT.MM.JJJJ (z.B. 01.01.2024)")
            return False
        
        return True
    
    def get_data(self):
        """Fetch data from Octopus Energy server - only fetch missing data."""
        if not self.validate_inputs():
            return
        
        # Ensure smartmeter_data folder exists (in Documents)
        data_dir = get_smartmeter_data_folder()
        data_dir.mkdir(parents=True, exist_ok=True)
        
        # Save config if checkbox is checked
        self.save_config()
        
        # Disable button and show progress
        self.get_data_btn.config(state='disabled')
        self.progress_bar.grid(row=17, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)
        self.progress_bar.start(10)
        self.root.update()
        
        try:
            with self._capture_debug_output():
                try:
                    # Parse date range from UI (European format: DD.MM.YYYY)
                    period_from = datetime.strptime(self.from_date_var.get(), "%d.%m.%Y")
                    period_to = datetime.strptime(self.to_date_var.get(), "%d.%m.%Y")
                    period_to = period_to + timedelta(days=1) - timedelta(seconds=1)  # End of day
                    
                    # Safety: Never fetch data for current day - data may be incomplete
                    today = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
                    yesterday_end = today - timedelta(seconds=1)
                    
                    if period_to >= today:
                        period_to = yesterday_end
                    
                    # Check if we need to fetch data
                    fetch_from = period_from
                    fetch_to = period_to
                    need_to_fetch = True
                    
                    # Check if we already have data up to yesterday
                    if self.latest_timestamp and self.latest_timestamp.date() >= (today - timedelta(days=1)).date():
                        self._set_status(f"CSV ist bereits aktuell ({self.latest_timestamp.date()}). " +
                                         "Es werden keine Daten geladen.", update=True)
                        need_to_fetch = False
                        fetch_from = None
                        fetch_to = None
                    elif self.latest_timestamp and self.latest_timestamp >= period_from:
                        # We have some data, fetch only what's missing
                        fetch_from = self.latest_timestamp - timedelta(hours=1)
                        if fetch_from > yesterday_end:
                            fetch_from = yesterday_end
                            need_to_fetch = False
                        self._set_status(f"Vorhandene Daten entdeckt. Lese ab {fetch_from}...", update=True)
                    
                    new_readings = []
                    
                    if need_to_fetch:
                        self._set_status("Authentifizierung...", update=True)
                        
                        # Initialize client
                        client = OctopusGermanyClient(
                            self.email_var.get(),
                            self.password_var.get(),
                            debug=self.debug_var.get()
                        )
                        
                        if not client.authenticate():
                            raise Exception("Authentifizierung fehlgeschlagen! Überprüfen Sie Ihre E-Mail und Ihr Passwort.")
                        
                        # Auto-discover account number
                        self._set_status("Kundennummer wird ermittelt...", update=True)
                        
                        accounts = client.get_accounts_from_viewer()
                        if not accounts:
                            raise Exception("Kein Konto gefunden! Überprüfen Sie Ihre Zugangsdaten.")
                        if len(accounts) > 1:
                            account_list = "\n".join([f"  - {acc.get('number', 'unknown')}" for acc in accounts])
                            raise Exception(f"Mehrere Konten gefunden ({len(accounts)}). Bitte wählen Sie ein Konto aus:\n{account_list}")
                        
                        account_number = accounts[0].get('number')
                        self._set_status(f"Kundennummer gefunden: {account_number}", update=True)
                        
                        self._set_status("Zähler werden ermittelt...", update=True)
                        
                        # Discover meters
                        meter_info = client.find_smart_meter(account_number)
                        
                        if not meter_info:
                            raise Exception("No smart meter found for this account!\n\nPossible reasons:\n- Smart meter not yet commissioned\n- No electricity meter found\n- Check account number")
                        
                        malo_number, meter_id, property_id = meter_info
                        
                        self._set_status(f"Zähler für MALO {malo_number} gefunden, Daten werden abgerufen...", update=True)
                        
                        # Progress callback function
                        def update_progress(count, page):
                            self._set_status(f"Empfange Daten... {count} Einträge (Seite {page})", update=True)
                        
                        # Fetch consumption data with progress updates
                        new_readings = client.get_consumption_graphql(
                            property_id=property_id,
                            period_from=fetch_from,
                            period_to=fetch_to,
                            fetch_all=True,
                            progress_callback=update_progress
                        )
                        
                        if not new_readings and not self.existing_data:
                            raise Exception("Keine Verbrauchsdaten gefunden!\n\nMögliche Gründe:\n- Smart Meter sendet noch keine Daten\n- Keine Messwerte verfügbar\n- Zählerproblem - kontaktieren Sie Octopus")
                    
                    # Merge existing and new data
                    all_readings = self.existing_data + new_readings
                    
                    if not all_readings:
                        raise Exception("Keine Daten zum Speichern!")
                    
                    # Remove duplicates based on start time
                    seen = {}
                    for reading in all_readings:
                        key = reading['start'].isoformat()
                        seen[key] = reading
                    
                    unique_data = list(seen.values())
                    unique_data.sort(key=lambda x: normalize_datetime(x['start']))
                    
                    # Update our data
                    self.existing_data = unique_data
                    if unique_data:
                        self.latest_timestamp = max(normalize_datetime(r['end']) for r in unique_data)
                    
                    # Save based on selected format
                    format_type = self.output_format_var.get()
                    data_folder = get_smartmeter_data_folder()
                    
                    if format_type == "excel":
                        # Save to CSV and fill Excel
                        self._set_status(f"Speichere {len(unique_data)} Einträge in consumption.csv...", update=True)
                        
                        with open(self.csv_path, 'w', newline='', encoding='utf-8') as f:
                            writer = csv.writer(f)
                            writer.writerow(['start', 'end', 'consumption_kwh'])
                            for reading in unique_data:
                                writer.writerow([
                                    format_datetime(reading['start']),
                                    format_datetime(reading['end']),
                                    reading['consumption_kwh']
                                ])
                        
                        excel_path = self._get_normalized_excel_path().resolve()
                        self._set_status("Excel-Datei wird gefüllt...", update=True)
                        
                        success = fill_excel_template(unique_data, str(excel_path), str(excel_path))
                        if success:
                            messagebox.showinfo(
                                "Erfolg", 
                                f"Daten erfolgreich gespeichert!\n\n"
                                f"CSV: consumption.csv ({len(unique_data)} Einträge)\n"
                                f"Excel: {excel_path}"
                            )
                        else:
                            raise Exception("Excel-Vorlage konnte nicht gefüllt werden")
                            
                    elif format_type == "csv":
                        # Save to CSV only
                        self._set_status(f"Speichere {len(unique_data)} Einträge in consumption.csv...", update=True)
                        
                        with open(self.csv_path, 'w', newline='', encoding='utf-8') as f:
                            writer = csv.writer(f)
                            writer.writerow(['start', 'end', 'consumption_kwh'])
                            for reading in unique_data:
                                writer.writerow([
                                    format_datetime(reading['start']),
                                    format_datetime(reading['end']),
                                    reading['consumption_kwh']
                                ])
                        
                        messagebox.showinfo(
                            "Erfolg",
                            f"Daten erfolgreich gespeichert!\n\n"
                            f"CSV: consumption.csv\n"
                            f"Gesamteinträge: {len(unique_data)}"
                        )
                        
                    elif format_type == "json":
                        # Save to JSON
                        json_path = data_folder / "consumption.json"
                        self._set_status(f"Speichere {len(unique_data)} Einträge als JSON...", update=True)
                        
                        if save_to_json(unique_data, json_path):
                            messagebox.showinfo(
                                "Erfolg",
                                f"Daten erfolgreich gespeichert!\n\n"
                                f"JSON: consumption.json\n"
                                f"Gesamteinträge: {len(unique_data)}"
                            )
                        else:
                            raise Exception("Fehler beim Speichern als JSON")
                            
                    elif format_type == "yaml":
                        # Save to YAML
                        yaml_path = data_folder / "consumption.yaml"
                        self._set_status(f"Speichere {len(unique_data)} Einträge als YAML...", update=True)
                        
                        if save_to_yaml(unique_data, yaml_path):
                            messagebox.showinfo(
                                "Erfolg",
                                f"Daten erfolgreich gespeichert!\n\n"
                                f"YAML: consumption.yaml\n"
                                f"Gesamteinträge: {len(unique_data)}"
                            )
                        else:
                            raise Exception("Fehler beim Speichern als YAML")
                    
                    # Show completion status
                    self._set_status(f"Fertig! Daten in Documents/smartmeter_data/ ({len(unique_data)} Einträge)")
                except Exception:
                    if self.debug_var.get():
                        traceback.print_exc()
                    raise
            
        except Exception as e:
            messagebox.showerror("Fehler", f"Ein Fehler ist aufgetreten:\n\n{str(e)}")
            self.status_var.set(f"Fehler: {str(e)}")
        finally:
            self.progress_bar.stop()
            self.progress_bar.grid_remove()
            self.get_data_btn.config(state='normal')

    def _write_csv_file(self, path, readings):
        """Write readings to a CSV file."""
        with open(path, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            writer.writerow(['start', 'end', 'consumption_kwh'])
            for reading in readings:
                writer.writerow([
                    format_datetime(reading['start']),
                    format_datetime(reading['end']),
                    reading['consumption_kwh']
                ])

    def validate_inputs(self):
        """Validate user inputs."""
        if not self.email_var.get():
            messagebox.showerror("Fehler", "E-Mail ist erforderlich!")
            return False
        if not self.password_var.get():
            messagebox.showerror("Fehler", "Passwort ist erforderlich!")
            return False
        if not self.output_file_var.get():
            messagebox.showerror("Fehler", "Bitte wählen Sie einen Dateinamen aus!")
            return False

        self._normalize_output_entry()

        try:
            from_date = datetime.strptime(self.from_date_var.get(), "%d.%m.%Y")
            to_date = datetime.strptime(self.to_date_var.get(), "%d.%m.%Y")
            if from_date > to_date:
                messagebox.showerror("Fehler", "Das Von-Datum muss vor oder gleich dem Bis-Datum sein!")
                return False
        except ValueError:
            messagebox.showerror("Fehler", "Ungültiges Datumsformat! Verwenden Sie TT.MM.JJJJ (z.B. 01.01.2024)")
            return False

        return True

    def get_data(self):
        """Fetch data from Octopus Energy server - only fetch missing data."""
        if not self.validate_inputs():
            return

        data_dir = get_smartmeter_data_folder()
        data_dir.mkdir(parents=True, exist_ok=True)

        self.save_config()

        self.get_data_btn.config(state='disabled')
        self.progress_bar.grid(row=17, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)
        self.progress_bar.start(10)
        self.root.update()

        try:
            with self._capture_debug_output():
                try:
                    period_from = datetime.strptime(self.from_date_var.get(), "%d.%m.%Y")
                    period_to = datetime.strptime(self.to_date_var.get(), "%d.%m.%Y")
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
                            f"CSV ist bereits aktuell ({self.latest_timestamp.date()}). Es werden keine Daten geladen.",
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
                        self._set_status(f"Vorhandene Daten entdeckt. Lese ab {fetch_from}...", update=True)

                    new_readings = []

                    if need_to_fetch:
                        self._set_status("Authentifizierung...", update=True)

                        client = OctopusGermanyClient(
                            self.email_var.get(),
                            self.password_var.get(),
                            debug=self.debug_var.get()
                        )

                        if not client.authenticate():
                            raise Exception("Authentifizierung fehlgeschlagen! Überprüfen Sie Ihre E-Mail und Ihr Passwort.")

                        self._set_status("Kundennummer wird ermittelt...", update=True)

                        accounts = client.get_accounts_from_viewer()
                        if not accounts:
                            raise Exception("Kein Konto gefunden! Überprüfen Sie Ihre Zugangsdaten.")
                        if len(accounts) > 1:
                            account_list = "\n".join([f"  - {acc.get('number', 'unknown')}" for acc in accounts])
                            raise Exception(f"Mehrere Konten gefunden ({len(accounts)}). Bitte wählen Sie ein Konto aus:\n{account_list}")

                        account_number = accounts[0].get('number')
                        self._set_status(f"Kundennummer gefunden: {account_number}", update=True)

                        self._set_status("Zähler werden ermittelt...", update=True)

                        meter_info = client.find_smart_meter(account_number)
                        if not meter_info:
                            raise Exception(
                                "No smart meter found for this account!\n\n"
                                "Possible reasons:\n"
                                "- Smart meter not yet commissioned\n"
                                "- No electricity meter found\n"
                                "- Check account number"
                            )

                        malo_number, meter_id, property_id = meter_info
                        self._set_status(f"Zähler für MALO {malo_number} gefunden, Daten werden abgerufen...", update=True)

                        def update_progress(count, page):
                            self._set_status(f"Empfange Daten... {count} Einträge (Seite {page})", update=True)

                        new_readings = client.get_consumption_graphql(
                            property_id=property_id,
                            period_from=fetch_from,
                            period_to=fetch_to,
                            fetch_all=True,
                            progress_callback=update_progress
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
                        seen[reading['start'].isoformat()] = reading

                    unique_data = list(seen.values())
                    unique_data.sort(key=lambda x: normalize_datetime(x['start']))

                    self.existing_data = unique_data
                    if unique_data:
                        self.latest_timestamp = max(normalize_datetime(r['end']) for r in unique_data)

                    format_type = self.output_format_var.get()
                    output_path = self._get_normalized_output_path().resolve()
                    output_path.parent.mkdir(parents=True, exist_ok=True)

                    self._set_status(f"Speichere {len(unique_data)} Einträge in consumption.csv...", update=True)
                    self._write_csv_file(self.csv_path, unique_data)

                    if format_type == "excel":
                        self._set_status("Excel-Datei wird gefüllt...", update=True)
                        success = fill_excel_template(unique_data, str(output_path), str(output_path))
                        if success:
                            messagebox.showinfo(
                                "Erfolg",
                                f"Daten erfolgreich gespeichert!\n\n"
                                f"CSV: consumption.csv ({len(unique_data)} Einträge)\n"
                                f"Excel: {output_path}"
                            )
                        else:
                            raise Exception("Excel-Vorlage konnte nicht gefüllt werden")

                    elif format_type == "csv":
                        if output_path != self.csv_path.resolve():
                            self._set_status(f"Speichere {len(unique_data)} Einträge als CSV...", update=True)
                            self._write_csv_file(output_path, unique_data)

                        messagebox.showinfo(
                            "Erfolg",
                            f"Daten erfolgreich gespeichert!\n\n"
                            f"CSV: {output_path}\n"
                            f"Gesamteinträge: {len(unique_data)}"
                        )

                    elif format_type == "json":
                        self._set_status(f"Speichere {len(unique_data)} Einträge als JSON...", update=True)
                        if save_to_json(unique_data, output_path):
                            messagebox.showinfo(
                                "Erfolg",
                                f"Daten erfolgreich gespeichert!\n\n"
                                f"JSON: {output_path}\n"
                                f"Gesamteinträge: {len(unique_data)}"
                            )
                        else:
                            raise Exception("Fehler beim Speichern als JSON")

                    elif format_type == "yaml":
                        self._set_status(f"Speichere {len(unique_data)} Einträge als YAML...", update=True)
                        if save_to_yaml(unique_data, output_path):
                            messagebox.showinfo(
                                "Erfolg",
                                f"Daten erfolgreich gespeichert!\n\n"
                                f"YAML: {output_path}\n"
                                f"Gesamteinträge: {len(unique_data)}"
                            )
                        else:
                            raise Exception("Fehler beim Speichern als YAML")

                    self._set_status(f"Fertig! Daten in Documents/smartmeter_data/ ({len(unique_data)} Einträge)")
                except Exception:
                    if self.debug_var.get():
                        traceback.print_exc()
                    raise

        except Exception as e:
            messagebox.showerror("Fehler", f"Ein Fehler ist aufgetreten:\n\n{str(e)}")
            self.status_var.set(f"Fehler: {str(e)}")
        finally:
            self.progress_bar.stop()
            self.progress_bar.grid_remove()
            self.get_data_btn.config(state='normal')


def main():
    root = tk.Tk()
    app = OctopusSmartMeterGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
