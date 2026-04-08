# Octopus Energy Deutschland Smart Meter Daten-Logger

Ein Python-Tool zum Abrufen von Smart Meter Verbrauchsdaten über die **Octopus Energy Deutschland** API.

## Funktionen

- 📊 **Stündliche Verbrauchsdaten** - Import und Export
- 🔄 **Inkrementelle Aktualisierung** - Nur neue Daten abrufen, bestehende CSV beibehalten
- 📁 **Einzelne CSV-Datei** - Alle Daten in `consumption.csv`
- 📈 **Excel-Vorlagenfüllung** - Füllt die deutsche Stromtarif-Vorlage
- 🔍 **Automatische Zählererkennung** über GraphQL mit Kundennummer
- 🔐 **E-Mail/Passwort-Authentifizierung** über GraphQL
- 🔄 **Automatische Token-Aktualisierung** - Behandelt 60-minütige Token-Gültigkeit
- 🖥️ **GUI-Version** - Benutzerfreundliche grafische Oberfläche

## Voraussetzungen

1. Ein Octopus Energy **Deutschland** Konto
2. Ihre Octopus Energy Zugangsdaten (E-Mail und Passwort)
3. Ihre Kundennummer (z.B. `A-12345678`)

### Ihre Kundennummer finden

- Melden Sie sich bei Ihrem [Octopus Energy Deutschland Dashboard](https://octopus.energy/de/) an
- Ihre Kundennummer wird im Dashboard angezeigt (z.B. `A-12345678`)
- Sie finden sie auch auf Ihren Rechnungen

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

#### Alle verfügbaren Daten abrufen

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
    --account-number A-12345678 \
    --period-from 01.01.2024 \
    --period-to 31.01.2024
```

#### Excel-Vorlage mit Verbrauchsdaten füllen

```bash
python -m octopusdetool.octopusdetool \
    --email user@example.com \
    --password ihr_passwort \
    --account-number A-12345678 \
    --fill-excel ~/Documents/smartmeter_data/stromtarif_verbrauch_bis_2027_mit_grundpreis_blanko.xlsx
```

#### Debug-Modus aktivieren

```bash
python -m octopusdetool.octopusdetool \
    --email user@example.com \
    --password ihr_passwort \
    --account-number A-12345678 \
    --debug
```

## Kommandozeilenoptionen

| Option | Beschreibung |
|--------|--------------|
| `--email` | **Erforderlich.** Ihre Octopus Energy E-Mail |
| `--password` | **Erforderlich.** Ihr Octopus Energy Passwort |
| `--account-number` | **Erforderlich.** Ihre Kundennummer (z.B. `A-12345678`) |
| `--meter-id` | Zähler-ID (optional - wird automatisch ermittelt) |
| `--property-id` | Eigenschafts-ID (optional - wird automatisch ermittelt) |
| `--output` | Ausgabepfad für CSV (Standard: `~/Documents/smartmeter_data/consumption.csv`) |
| `--period-from` | Startdatum (DD.MM.YYYY) |
| `--period-to` | Enddatum (DD.MM.YYYY) |
| `--format` | Ausgabeformat: `csv`, `hourly`, oder `all` |
| `--fill-excel` | Excel-Vorlage mit Verbrauchsdaten füllen |
| `--debug` | Debug-Ausgabe für alle API-Anfragen aktivieren |

## Ausgabe

Die Daten werden in `~/Documents/smartmeter_data/` gespeichert:

```
~/Documents/smartmeter_data/
├── consumption.csv              # Alle Verbrauchsdaten
├── stromtarif_verbrauch_bis_2027_mit_grundpreis_blanko.xlsx  # Excel-Vorlage
└── config.json                  # GUI-Konfiguration (optional)
```

### CSV-Format

| Spalte | Beschreibung |
|--------|--------------|
| `start` | Startzeitpunkt (DD.MM.YYYY HH:MM:SS) |
| `end` | Endzeitpunkt (DD.MM.YYYY HH:MM:SS) |
| `consumption_kwh` | Energieverbrauch in kWh |

### Excel-Vorlage

Das Tool füllt die deutsche Stromtarif-Vorlage:

- **Spalte A (Datum)**: Formelbasiert aus Einstellungen
- **Spalte B (Stunde)**: Formelbasiert (0-23)
- **Spalte C (Verbrauch)**: Wird mit Ihren Smart Meter Daten gefüllt
- **Einstellungen B5/B6**: Zeitraum wird automatisch aktualisiert

## GUI-Version

Die GUI bietet eine benutzerfreundliche Oberfläche:

### Funktionen

- **Eingabefelder** für E-Mail, Passwort und Kundennummer
- **Excel-Dateiauswahl** mit Dateibrowser
- **Ausgabeformatauswahl**: CSV oder Excel
- **Datumsbereichsauswahl** mit Kalender-Buttons
- **Konfigurationsspeicherung** - Speichert Einstellungen in `config.json`
- **Automatisches Laden** - Lädt gespeicherte Konfiguration beim Start
- **Fortschrittsanzeige** - Zeigt aktuellen Status
- **Intelligentes Abrufen** - Liest erst CSV, authentifiziert nur wenn nötig

### Screenshot

```
┌─────────────────────────────────────────────┐
│  Octopus Energy Deutschland                 │
│  Smart Meter Daten-Logger                   │
├─────────────────────────────────────────────┤
│  E-Mail:        [____________________]      │
│  Passwort:      [____________________]      │
│  Kundennummer:  [____________________]      │
│                                             │
│  [✓] Konfiguration in config.json speichern │
│  [ ] Debug-Ausgabe aktivieren               │
├─────────────────────────────────────────────┤
│  Ausgabeoptionen                            │
│  Format: (*) CSV  ( ) Excel                 │
│  Excel-Vorlage: [____________] [Durchsuchen]│
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

- Das Skript speichert Ihr Passwort nicht
- Tokens sind 60 Minuten gültig
- Die GUI kann Zugangsdaten optional in `config.json` speichern
- Diese Datei wird im Documents-Ordner gespeichert

## Lizenz

MIT Lizenz - Frei zu verwenden und zu modifizieren.
