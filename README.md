# Octopus Energy Deutschland Smart Meter Daten-Logger

Ein Python-Tool zum Abrufen von Smart Meter Verbrauchsdaten über die **Octopus Energy Deutschland** API.

(c) B.Kerler (App) und B.Stahl (xlsx) 2026

## GUI-Version

### Screenshots
<img width="1750" height="1445" alt="image" src="https://github.com/user-attachments/assets/8b62973e-7b0b-40e0-8373-6dba8df87129" />
<img width="1744" height="1442" alt="image" src="https://github.com/user-attachments/assets/38e0a22b-09e2-4d90-82bd-0d739248b25a" />
<img width="1737" height="1441" alt="image" src="https://github.com/user-attachments/assets/b5ef988c-c02a-4577-bb14-8a3f354c7d47" />
<img width="1737" height="1440" alt="image" src="https://github.com/user-attachments/assets/1280178e-f33c-4044-b4cf-797a516406c3" />

## Funktionen
- 📊 **Stündliche Verbrauchsdaten** - Import und Export
- 🔄 **Inkrementelle Aktualisierung** - Nur neue Daten abrufen
- 📁 **Mehrere Ausgabeformate** - CSV, JSON, YAML oder Excel
- 📈 **Excel-Erstellung** - Erstellt eine Excel-Datei mit allen Daten 
- 🔍 **Automatische Kontoerkennung** - Kundennummer wird automatisch ermittelt
- 🔄 **Automatische Token-Aktualisierung** - Behandelt 60-minütige Token-Gültigkeit
- 🖥️ **GUI-Version** - Benutzerfreundliche grafische Oberfläche

## Voraussetzungen

1. Ein Octopus Energy **Deutschland** Konto
2. Ihre Octopus Energy Zugangsdaten (E-Mail und Passwort)
3. Ein Octopus Intelligent Go oder Octopus Intelligent Heat Tarif (Light Tarife werden wegen fehlender Smartmeter-Übertragung nicht unterstützt)
Die Kundennummer wird automatisch ermittelt. Bei mehreren Konten können Sie mit `--list-accounts` alle Konten anzeigen und mit `--account-number` eines auswählen.

## Installation

```bash
# Repository klonen oder herunterladen
cd octopusdetool

# Abhängigkeiten installieren
pip install -r requirements.txt
```

## Verwendung

### GUI starten (Empfohlen)

```bash
python -m octopusdetool.octopusdetool_gui
```

Oder nach Installation:

```bash
octopusdetool_gui
```

### Kommandozeile (CLI)

#### Alle verfügbaren Daten abrufen (einfachst - nur E-Mail und Passwort)

```bash
python -m octopusdetool.octopusdetool \
    --email user@example.com \
    --password ihr_passwort
```

#### Als JSON speichern

```bash
python -m octopusdetool.octopusdetool \
    --email user@example.com \
    --password ihr_passwort \
    --output-format json
```

#### Als YAML speichern

```bash
python -m octopusdetool.octopusdetool \
    --email user@example.com \
    --password ihr_passwort \
    --output-format yaml
```

#### Konten auflisten (nur E-Mail und Passwort nötig)

```bash
python -m octopusdetool.octopusdetool \
    --email user@example.com \
    --password ihr_passwort \
    --list-accounts
```

#### Daten für bestimmtes Konto abrufen (bei mehreren Konten)

```bash
python -m octopusdetool.octopusdetool \
    --email user@example.com \
    --password ihr_passwort \
    --account-number A-12345678
```

#### Bestimmten Zeitraum abrufen (DD.MM.YYYY Format)

```bash
python -m octopusdetool.octopusdetool \
    --email user@example.com \
    --password ihr_passwort \
    --period-from 01.01.2024 \
    --period-to 31.01.2024
```

#### Excel-Vorlage mit Verbrauchsdaten füllen

```bash
python -m octopusdetool.octopusdetool \
    --email user@example.com \
    --password ihr_passwort \
    --fill-excel ~/Documents/smartmeter_data/smartmeter_daten.xlsx
```

#### Debug-Modus aktivieren

```bash
python -m octopusdetool.octopusdetool \
    --email user@example.com \
    --password ihr_passwort \
    --debug
```

## Kommandozeilenoptionen

| Option | Beschreibung |
|--------|--------------|
| `--email` | **Erforderlich.** Ihre Octopus Energy E-Mail |
| `--password` | **Erforderlich.** Ihr Octopus Energy Passwort |
| `--account-number` | Ihre Kundennummer (z.B. `A-12345678`). Wird automatisch ermittelt. |
| `--list-accounts` | Listet alle Konten auf (nützlich bei mehreren Konten) |
| `--output` | Ausgabepfad (Standard: `~/Documents/smartmeter_data/consumption`) |
| `--output-format` | Dateiformat: `csv`, `json`, `yaml` (Standard: csv) |
| `--period-from` | Startdatum (DD.MM.YYYY) |
| `--period-to` | Enddatum (DD.MM.YYYY) |
| `--format` | Datenausgabe: `csv`, `hourly`, `all` (veraltet, nutzen Sie `--output-format`) |
| `--fill-excel` | Excel-Vorlage mit Verbrauchsdaten füllen |
| `--debug` | Debug-Ausgabe für alle API-Anfragen aktivieren |

## Ausgabe

Die Daten werden in `~/Documents/smartmeter_data/` gespeichert:

```
~/Documents/smartmeter_data/
├── consumption.csv              # Alle Verbrauchsdaten (CSV)
├── consumption.json             # Alle Verbrauchsdaten (JSON)
├── consumption.yaml             # Alle Verbrauchsdaten (YAML)
├── log.txt                      # GUI-Debug-Log mit API-Requests/Responses, Statusmeldungen und Fehlern
├── smartmeter_daten.xlsx  # Excel-Vorlage
└── config.json                  # GUI-Konfiguration (optional)
```

### CSV-Format

| Spalte | Beschreibung |
|--------|--------------|
| `start` | Startzeitpunkt (DD.MM.YYYY HH:MM:SS) |
| `end` | Endzeitpunkt (DD.MM.YYYY HH:MM:SS) |
| `consumption_kwh` | Energieverbrauch in kWh |

### JSON-Format

```json
{
  "metadata": {
    "export_date": "2024-01-15T10:30:00",
    "total_readings": 3600,
    "source": "Octopus Energy Germany Smart Meter"
  },
  "readings": [
    {
      "start": "2024-01-01T00:00:00+00:00",
      "end": "2024-01-01T01:00:00+00:00",
      "consumption_kwh": 0.5,
      "duration_seconds": 3600,
      "unit": "kWh"
    }
  ]
}
```

### YAML-Format

```yaml
metadata:
  export_date: '2024-01-15T10:30:00'
  total_readings: 3600
  source: Octopus Energy Germany Smart Meter
readings:
  - start: '2024-01-01T00:00:00+00:00'
    end: '2024-01-01T01:00:00+00:00'
    consumption_kwh: 0.5
    duration_seconds: 3600
    unit: kWh
```

### Excel-Vorlage

Das Tool füllt die deutsche Stromtarif-Vorlage:

- **Spalte A (Datum)**: Formelbasiert aus Einstellungen
- **Spalte B (Stunde)**: Formelbasiert (0-23)
- **Spalte C (Verbrauch)**: Wird mit Ihren Smart Meter Daten gefüllt
- **Einstellungen B5/B6**: Zeitraum wird automatisch aktualisiert

## API-Hinweise

- **Deutsche API Basis-URL:** `https://api.oeg-kraken.energy/v1/`
- **GraphQL URL:** `https://api.oeg-kraken.energy/v1/graphql/`
- **GraphQL Dokumentation:** `https://docs.oeg-kraken.energy/`
- **Authentifizierung:** E-Mail/Passwort via GraphQL `obtainKrakenToken`
- **Token-Gültigkeit:** 60 Minuten (automatisch aktualisiert)
- **Aktualisierungstoken:** 7 Tage gültig
- **Datenverzögerung:** Typischerweise bis zu 24 Stunden
- **Paginierung:** 100 Datensätze pro Anfrage

### Unterschiede zwischen deutscher und britischer API

| Merkmal | Deutschland | UK |
|---------|-------------|-----|
| Basis-URL | `api.oeg-kraken.energy` | `api.octopus.energy` |
| Auth-Methode | E-Mail/Passwort | API-Schlüssel |
| Zähler-ID | MALO (Marktlokation) | MPAN |

## Offizielle Dokumentation

- **Startseite:** https://docs.oeg-kraken.energy/
- **GraphQL Guides:** https://docs.oeg-kraken.energy/graphql/guides/basics
- **GraphQL Referenz:** https://docs.oeg-kraken.energy/graphql/reference/

## Fehlerbehebung

### "Authentifizierung fehlgeschlagen"
- Überprüfen Sie E-Mail und Passwort
- Stellen Sie sicher, dass das Konto aktiv ist
- Testen Sie die Anmeldung über die Webseite

### "Kein Konto gefunden"
- Überprüfen Sie Ihre Zugangsdaten
- Verwenden Sie `--list-accounts` um verfügbare Konten anzuzeigen

### "Kein Smart Meter gefunden"
- Prüfen Sie, ob ein Smart Meter installiert ist
- Vergewissern Sie sich, dass der Zähler für Smart-Readings freigeschaltet ist
- Kontaktieren Sie Octopus Energy bei anhaltenden Problemen

### "Keine Verbrauchsdaten gefunden"
- Smart Meter Daten sind typischerweise 24-48 Stunden verzögert
- Versuchen Sie, ältere Daten abzurufen
- Neue Smart Meter können mehrere Tage brauchen

### Excel-Vorlage nicht gefunden
- Die Vorlage wird automatisch aus dem Paket nach `~/Documents/smartmeter_data/` kopiert
- Stellen Sie sicher, dass der Documents-Ordner existiert

## Sicherheitshinweise

- Das CLI-Skript speichert Ihr Passwort nicht dauerhaft
- Tokens sind 60 Minuten gültig
- Die GUI kann Zugangsdaten optional in `config.json` speichern
- `email` und `password` werden dabei mit AES-256-GCM verschlüsselt und als Base64 abgelegt
- Vorhandene Klartext-Einträge in älteren `config.json`-Dateien werden beim Laden automatisch migriert
- Diese Datei wird im Documents-Ordner gespeichert

## Lizenz

MIT Lizenz - Frei zu verwenden und zu modifizieren.
