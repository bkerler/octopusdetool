# Octopus Energy Deutschland Smart Meter Daten-Logger

Ein Python-Tool zum Abrufen von Smart Meter Verbrauchsdaten über die **Octopus Energy Deutschland** API.

(c) B.Kerler (App) und B.Stahl (xlsx) 2026

<img width="578" height="496" alt="image" src="https://github.com/user-attachments/assets/262456e5-1821-470a-bada-801520b5e9c9" />

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
    --fill-excel ~/Documents/smartmeter_data/stromtarif_verbrauch_bis_2027_mit_grundpreis_blanko.xlsx
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
├── stromtarif_verbrauch_bis_2027_mit_grundpreis_blanko.xlsx  # Excel-Vorlage
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

## GUI-Version

Die GUI bietet eine benutzerfreundliche Oberfläche:

### Funktionen

- **Eingabefelder** für E-Mail und Passwort
- **Passwortanzeige umschaltbar** - Passwort ist standardmäßig maskiert; `Show password` blendet es ein
- **Ausgabeformatauswahl**: Excel (Standard), CSV, JSON oder YAML
- **Excel-Dateiauswahl** mit Speichern-unter-Dialog für neue oder bestehende `.xlsx`-Dateien
- **Gezielter Excel-Speicherort** - Die Excel-Datei wird immer genau unter dem im Eingabefeld angegebenen Pfad gespeichert
- **Datumsbereichsauswahl** mit Kalender-Buttons
- **Fortschrittsanzeige** - Zeigt "Empfange Daten... X Einträge (Seite Y)"
- **Konfigurationsspeicherung** - Speichert E-Mail/Passwort verschlüsselt in `config.json` (AES-256-GCM, Base64)
- **Debug-Ausgabe** - Wenn aktiviert, wird der vollständige GUI-Debug-Log nach `~/Documents/smartmeter_data/log.txt` geschrieben
- **Automatisches Laden** - Lädt gespeicherte Konfiguration beim Start
- **Automatische Migration** - Bereits vorhandene Klartext-Zugangsdaten in `config.json` werden beim Laden automatisch verschlüsselt
- **Automatische Excel-Vorlage** - Existiert die angegebene Excel-Datei noch nicht, wird die mitgelieferte Blanko-Vorlage unter genau diesem Dateinamen erstellt

### Screenshot

```
┌─────────────────────────────────────────────┐
│  Octopus Energy Deutschland                 │
│  Smart Meter Daten-Logger                   │
├─────────────────────────────────────────────┤
│  E-Mail:        [____________________]      │
│  Passwort:      [***************     ] [ ]  │
│                           Passwort anzeigen │
│                                             │
│  [✓] Konfiguration in config.json speichern │
│  [ ] Debug-Ausgabe aktivieren               │
├─────────────────────────────────────────────┤
│  Ausgabeoptionen                            │
│  Format: (*) Excel  ( ) CSV  ( ) JSON/YAML  │
│  Excel-Vorlage: [____________] [Speichern unter]│
├─────────────────────────────────────────────┤
│  Datumsbereich                              │
│  Von: [01.01.2024    ] [📅]                 │
│  Bis: [07.01.2026    ] [📅]                 │
├─────────────────────────────────────────────┤
│  Bereit                                     │
│  [Daten vom Server abrufen]                 │
└─────────────────────────────────────────────┘
```

## API-Hinweise

- **Deutsche API Basis-URL:** `https://api.oeg-kraken.energy/v1/`
- **GraphQL URL:** `https://api.oeg-kraken.energy/v1/graphql/`
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
