#!/usr/bin/env python3
"""
Octopus Energy Germany Smart Meter Data Logger - GUI Version

A tkinter-based GUI for fetching smart meter consumption data from 
Octopus Energy Germany API and saving it to CSV or Excel.
"""

import base64
import csv
import json
import os
import platform
import shutil
import sys
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import datetime, timedelta
from io import BytesIO
from pathlib import Path

# Import from the same package
from octopusdetool import (
    OctopusGermanyClient, 
    fill_excel_template, 
    format_datetime,
    get_documents_folder,
    get_smartmeter_data_folder,
    ensure_excel_template,
    get_default_output_path,
    get_default_excel_path
)


CONFIG_FILE = get_smartmeter_data_folder() / "config.json"

# Embedded calendar icon (PNG, 32x32)
CALENDAR_ICON_BASE64 = "iVBORw0KGgoAAAANSUhEUgAAACAAAAAgCAYAAABzenr0AAAAuUlEQVR4nO1Wyw3FIAxLn7oTC8EGjMEGsFC2YYO+W1XR0ET9uQdygmIcy0lQp1rrQsD4IZMTEc0WUIxxXaeUbsWrDmzJpP1V/NT2gHbhjti6IpbAe7+uSymH521o+PZcLUGb7Cj5GbypCTWSK3hRgGTjU/HddyDnbCJgZnLOmb6HEHY4uANDwBDQnQJmNpP0sBaOrgBptHpJrGMoPXDwEsAFjB6AlwAuYPQAvATw3/LvOfB2wB2AC/gDw6NqeR/bFyoAAAAASUVORK5CYII="


class OctopusSmartMeterGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Octopus Energy Germany - Smart Meter Data Logger")
        # Set size large enough to show all elements without scrolling
        self.root.geometry("1600x950")
        self.root.minsize(1400, 900)
        self.root.resizable(True, True)
        
        # Style configuration
        self.style = ttk.Style()
        self.style.configure('TFrame', background='#f0f0f0')
        self.style.configure('TLabel', background='#f0f0f0', font=('Arial', 10))
        self.style.configure('TButton', font=('Arial', 10))
        self.style.configure('Header.TLabel', font=('Arial', 12, 'bold'))
        
        # Create a canvas with scrollbar for resizable content
        self.canvas = tk.Canvas(root, background='#f0f0f0')
        self.scrollbar = ttk.Scrollbar(root, orient="vertical", command=self.canvas.yview)
        self.main_frame = ttk.Frame(self.canvas, padding="20")
        
        self.main_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )
        
        self.canvas.create_window((0, 0), window=self.main_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        
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
        
        self.create_widgets()
        self.load_config()
        self.check_existing_data()
    
    def _on_mousewheel(self, event):
        self.canvas.yview_scroll(int(-1*(event.delta/120)), "units")
    
    def create_widgets(self):
        """Create all GUI widgets."""
        row = 0
        
        # Title
        title_label = ttk.Label(
            self.main_frame, 
            text="Octopus Energy Deutschland\nSmart Meter Daten-Logger",
            style='Header.TLabel',
            justify='center'
        )
        title_label.grid(row=row, column=0, columnspan=3, pady=(0, 20))
        row += 1
        
        # Separator
        ttk.Separator(self.main_frame, orient='horizontal').grid(
            row=row, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10)
        )
        row += 1
        
        # Email
        ttk.Label(self.main_frame, text="E-Mail:").grid(
            row=row, column=0, sticky=tk.W, pady=5
        )
        self.email_var = tk.StringVar()
        self.email_entry = ttk.Entry(self.main_frame, textvariable=self.email_var, width=60)
        self.email_entry.grid(row=row, column=1, sticky=(tk.W, tk.E), pady=5, padx=5)
        self.email_entry.config(state='normal')
        row += 1
        
        # Password
        ttk.Label(self.main_frame, text="Passwort:").grid(
            row=row, column=0, sticky=tk.W, pady=5
        )
        self.password_var = tk.StringVar()
        self.password_entry = ttk.Entry(
            self.main_frame, textvariable=self.password_var, width=60, show="*"
        )
        self.password_entry.grid(row=row, column=1, sticky=(tk.W, tk.E), pady=5, padx=5)
        self.password_entry.config(state='normal')
        row += 1
        
        # Account Number
        ttk.Label(self.main_frame, text="Kundennummer:").grid(
            row=row, column=0, sticky=tk.W, pady=5
        )
        self.account_var = tk.StringVar()
        self.account_entry = ttk.Entry(self.main_frame, textvariable=self.account_var, width=60)
        self.account_entry.grid(row=row, column=1, sticky=(tk.W, tk.E), pady=5, padx=5)
        self.account_entry.config(state='normal')
        row += 1
        
        # Save Configuration Checkbox - right under account number
        self.save_config_var = tk.BooleanVar(value=False)
        self.save_config_checkbox = ttk.Checkbutton(
            self.main_frame, text="Konfiguration in config.json speichern",
            variable=self.save_config_var
        )
        self.save_config_checkbox.grid(row=row, column=0, columnspan=3, sticky=tk.W, pady=5)
        row += 1
        
        # Debug Mode Checkbox - right under save config
        self.debug_var = tk.BooleanVar(value=False)
        self.debug_check = ttk.Checkbutton(
            self.main_frame, 
            text="Debug-Ausgabe aktivieren (zeigt alle API-Anfragen)",
            variable=self.debug_var
        )
        self.debug_check.grid(row=row, column=0, columnspan=3, sticky=tk.W, pady=5)
        row += 1
        
        # Separator
        ttk.Separator(self.main_frame, orient='horizontal').grid(
            row=row, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=10
        )
        row += 1
        
        # Output Options Frame
        output_frame = ttk.LabelFrame(self.main_frame, text="Ausgabeoptionen", padding="10")
        output_frame.grid(row=row, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=10, padx=5)
        output_frame.columnconfigure(1, weight=1)
        row += 1
        
        # Output Format
        ttk.Label(output_frame, text="Format:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.output_format_var = tk.StringVar(value="csv")
        
        format_combo = ttk.Combobox(
            output_frame, 
            textvariable=self.output_format_var,
            values=["csv", "hourly_csv", "excel"],
            state="readonly",
            width=20
        )
        format_combo.grid(row=0, column=1, sticky=tk.W, pady=5, padx=5)
        format_combo.bind("<<ComboboxSelected>>", self.on_format_changed)
        
        # Excel File Selection
        ttk.Label(output_frame, text="Excel-Vorlage:").grid(row=1, column=0, sticky=tk.W, pady=5)
        
        excel_frame = ttk.Frame(output_frame)
        excel_frame.grid(row=1, column=1, sticky=(tk.W, tk.E), pady=5)
        excel_frame.columnconfigure(0, weight=1)
        
        # Default to the Excel path in Documents folder
        self.excel_var = tk.StringVar(value=str(get_default_excel_path()))
        self.excel_entry = ttk.Entry(excel_frame, textvariable=self.excel_var, width=50)
        self.excel_entry.grid(row=0, column=0, sticky=(tk.W, tk.E))
        
        self.browse_btn = ttk.Button(
            excel_frame, text="Durchsuchen...", command=self.browse_excel, width=10
        )
        self.browse_btn.grid(row=0, column=1, padx=(5, 0))
        
        # Date Range Frame
        self.date_frame = ttk.LabelFrame(self.main_frame, text="Datumsbereich", padding="10")
        self.date_frame.grid(row=row, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=10, padx=5)
        self.date_frame.columnconfigure(1, weight=1)
        self.date_frame.columnconfigure(3, weight=1)
        row += 1
        
        # From Date - with calendar button directly adjacent
        ttk.Label(self.date_frame, text="Von:").grid(row=0, column=0, sticky=tk.W, padx=(5, 2))
        # Default from date is 01.01.2024 (European format)
        self.from_date_var = tk.StringVar(value="01.01.2024")
        self.from_date_entry = ttk.Entry(self.date_frame, textvariable=self.from_date_var, width=12)
        self.from_date_entry.grid(row=0, column=1, sticky=tk.W, padx=0)
        # Load embedded calendar icon
        try:
            icon_data = base64.b64decode(CALENDAR_ICON_BASE64)
            icon_image = tk.PhotoImage(data=icon_data)
            self.calendar_icon = icon_image
        except Exception as e:
            print(f"Warning: Could not load calendar icon: {e}")
            self.calendar_icon = None
        
        self.from_calendar_btn = tk.Button(
            self.date_frame, image=self.calendar_icon, width=32, height=32,
            command=lambda: self.show_calendar(self.from_date_var)
        )
        self.from_calendar_btn.grid(row=0, column=2, sticky=tk.W, padx=(0, 10))
        
        # To Date - with calendar button directly adjacent
        ttk.Label(self.date_frame, text="Bis:").grid(row=0, column=3, sticky=tk.W, padx=(5, 2))
        # To date is always yesterday (last complete day), format DD.MM.YYYY
        yesterday = (datetime.now() - timedelta(days=1)).strftime("%d.%m.%Y")
        self.to_date_var = tk.StringVar(value=yesterday)
        self.to_date_entry = ttk.Entry(self.date_frame, textvariable=self.to_date_var, width=12)
        self.to_date_entry.grid(row=0, column=4, sticky=tk.W, padx=0)
        self.to_calendar_btn = tk.Button(
            self.date_frame, image=self.calendar_icon, width=32, height=32,
            command=lambda: self.show_calendar(self.to_date_var)
        )
        self.to_calendar_btn.grid(row=0, column=5, sticky=tk.W, padx=0)
        
        # Progress Bar (initially hidden)
        self.progress_var = tk.DoubleVar(value=0)
        self.progress_bar = ttk.Progressbar(
            self.main_frame, variable=self.progress_var, maximum=100, mode='indeterminate'
        )
        
        # Status Label
        self.status_var = tk.StringVar(value="Bereit")
        self.status_label = ttk.Label(
            self.main_frame, textvariable=self.status_var, 
            foreground='blue', font=('Arial', 9, 'italic')
        )
        self.status_label.grid(row=row, column=0, columnspan=3, pady=5)
        row += 1
        
        # Get Data Button (bottom right)
        self.get_data_btn = ttk.Button(
            self.main_frame, text="Daten vom Server abrufen", 
            command=self.get_data, width=25
        )
        self.get_data_btn.grid(row=row, column=2, sticky=tk.E, pady=20)
        
        # Update UI state
        self.on_format_changed()
    
    def on_format_changed(self, event=None):
        """Handle output format change."""
        format_type = self.output_format_var.get()
        
        if format_type == "excel":
            self.excel_entry.config(state='normal')
            self.browse_btn.config(state='normal')
        else:
            self.excel_entry.config(state='disabled')
            self.browse_btn.config(state='disabled')
    
    def browse_excel(self):
        """Open file dialog to select Excel file."""
        filename = filedialog.askopenfilename(
            title="Excel-Vorlage auswählen",
            filetypes=[("Excel-Dateien", "*.xlsx"), ("Alle Dateien", "*.*")]
        )
        if filename:
            self.excel_var.set(filename)
    
    def show_calendar(self, target_var):
        """Show a simple calendar dialog."""
        top = tk.Toplevel(self.root)
        top.title("Datum auswählen")
        top.geometry("300x280")
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
        header_frame = ttk.Frame(top)
        header_frame.pack(pady=10)
        
        ttk.Spinbox(
            header_frame, from_=2020, to=2030, width=6,
            textvariable=selected_year
        ).pack(side=tk.LEFT, padx=5)
        
        months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
                  "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
        month_combo = ttk.Combobox(
            header_frame, values=months, width=6, state='readonly'
        )
        month_combo.set(months[current_date.month - 1])
        month_combo.pack(side=tk.LEFT, padx=5)
        
        # Calendar frame
        cal_frame = ttk.Frame(top)
        cal_frame.pack(pady=10)
        
        # Day buttons frame
        days_frame = ttk.Frame(cal_frame)
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
                ttk.Label(days_frame, text=day_name, width=4).grid(row=0, column=i)
            
            # Day buttons
            first_weekday, _ = calendar.monthrange(year, month)
            day = 1
            for week in range(1, 7):
                for weekday in range(7):
                    if week == 1 and weekday < first_weekday:
                        ttk.Label(days_frame, text="", width=4).grid(row=week, column=weekday)
                    elif day <= days_in_month:
                        btn = tk.Button(
                            days_frame, text=str(day), width=4,
                            command=lambda d=day: select_day(d)
                        )
                        if (day == current_date.day and 
                            month == current_date.month and 
                            year == current_date.year):
                            btn.config(bg='#4CAF50', fg='white')
                        btn.grid(row=week, column=weekday)
                        day += 1
        
        # Update button
        update_btn = ttk.Button(top, text="Update", command=update_calendar)
        update_btn.pack(pady=5)
        
        # Initial calendar display
        update_calendar()
    
    def load_config(self):
        """Load configuration from config.json."""
        # Ensure smartmeter_data folder exists
        get_smartmeter_data_folder().mkdir(parents=True, exist_ok=True)
        
        if CONFIG_FILE.exists():
            try:
                with open(CONFIG_FILE, 'r') as f:
                    config = json.load(f)
                
                self.email_var.set(config.get('email', ''))
                self.password_var.set(config.get('password', ''))
                self.account_var.set(config.get('account_number', ''))
                self.excel_var.set(config.get('excel_file', str(get_default_excel_path())))
                self.output_format_var.set(config.get('output_format', 'csv'))
                # Default from date: 01.01.2024 or from config
                default_from = config.get('from_date', '01.01.2024')
                self.from_date_var.set(default_from)
                # To date is always yesterday (last complete day)
                yesterday = (datetime.now() - timedelta(days=1)).strftime("%d.%m.%Y")
                self.to_date_var.set(yesterday)
                self.debug_var.set(config.get('debug', False))
                
                # Update UI based on loaded format
                self.on_format_changed()
                
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
                                start = datetime.fromisoformat(row['start'])
                                end = datetime.fromisoformat(row['end'])
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
            'account_number': self.account_var.get(),
            'excel_file': self.excel_var.get(),
            'output_format': self.output_format_var.get(),
            'from_date': self.from_date_var.get(),  # Store the from date
            'debug': self.debug_var.get(),
        }
        
        try:
            with open(CONFIG_FILE, 'w') as f:
                json.dump(config, f, indent=2)
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
        if not self.account_var.get():
            messagebox.showerror("Fehler", "Kundennummer ist erforderlich!")
            return False
        
        format_type = self.output_format_var.get()
        if format_type == "excel" and not self.excel_var.get():
            messagebox.showerror("Fehler", "Bitte wählen Sie eine Excel-Vorlagendatei aus!")
            return False
        
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
            account_number = self.account_var.get()
            
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
                self.status_var.set(f"CSV already up to date ({self.latest_timestamp.date()}). No fetch needed.")
                self.root.update()
                need_to_fetch = False
                fetch_from = None
                fetch_to = None
            elif self.latest_timestamp and self.latest_timestamp >= period_from:
                # We have some data, fetch only what's missing
                fetch_from = self.latest_timestamp - timedelta(hours=1)
                if fetch_from > yesterday_end:
                    fetch_from = yesterday_end
                    need_to_fetch = False
                self.status_var.set(f"Found existing data. Fetching from {fetch_from}...")
                self.root.update()
            
            new_readings = []
            
            if need_to_fetch:
                self.status_var.set("Authentifizierung...")
                self.root.update()
                
                # Initialize client
                client = OctopusGermanyClient(
                    self.email_var.get(),
                    self.password_var.get(),
                    debug=self.debug_var.get()
                )
                
                if not client.authenticate():
                    raise Exception("Authentifizierung fehlgeschlagen! Überprüfen Sie Ihre E-Mail und Ihr Passwort.")
                
                self.status_var.set("Zähler werden ermittelt...")
                self.root.update()
                
                # Discover meters
                meter_info = client.find_smart_meter(account_number)
                
                if not meter_info:
                    raise Exception("No smart meter found for this account!\n\nPossible reasons:\n- Smart meter not yet commissioned\n- No electricity meter found\n- Check account number")
                
                malo_number, meter_id, property_id = meter_info
                
                self.status_var.set(f"Zähler für MALO {malo_number} gefunden, Daten werden abgerufen...")
                self.root.update()
                
                # Fetch consumption data
                new_readings = client.get_consumption_graphql(
                    property_id=property_id,
                    period_from=fetch_from,
                    period_to=fetch_to,
                    fetch_all=True
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
            unique_data.sort(key=lambda x: x['start'])
            
            # Save to single consumption.csv
            self.status_var.set(f"Speichere {len(unique_data)} Einträge in consumption.csv...")
            self.root.update()
            
            with open(self.csv_path, 'w', newline='', encoding='utf-8') as f:
                writer = csv.writer(f)
                writer.writerow(['start', 'end', 'consumption_kwh'])
                for reading in unique_data:
                    writer.writerow([
                        format_datetime(reading['start']),
                        format_datetime(reading['end']),
                        reading['consumption_kwh']
                    ])
            
            # Update our data
            self.existing_data = unique_data
            if unique_data:
                self.latest_timestamp = max(r['end'] for r in unique_data)
            
            # Fill Excel if requested
            format_type = self.output_format_var.get()
            if format_type == "excel":
                excel_path = Path(self.excel_var.get()).resolve()
                self.status_var.set("Excel-Vorlage wird gefüllt...")
                self.root.update()
                
                success = fill_excel_template(unique_data, str(excel_path), str(excel_path))
                docs_folder = get_documents_folder()
                if success:
                    messagebox.showinfo(
                        "Erfolg", 
                        f"Daten erfolgreich in Documents/smartmeter_data/ gespeichert!\n\n"
                        f"CSV: consumption.csv ({len(unique_data)} Einträge)\n"
                        f"Excel: {excel_path}"
                    )
                else:
                    raise Exception("Excel-Vorlage konnte nicht gefüllt werden")
            else:
                docs_folder = get_documents_folder()
                messagebox.showinfo(
                    "Erfolg",
                    f"Daten erfolgreich in Documents/smartmeter_data/ gespeichert!\n\n"
                    f"CSV: consumption.csv\n"
                    f"Gesamteinträge: {len(unique_data)}"
                )
            
            # Show completion status
            self.status_var.set(f"Fertig! Daten in Documents/smartmeter_data/ ({len(unique_data)} Einträge)")
            
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
