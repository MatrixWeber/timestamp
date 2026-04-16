# stamp ⏱️ — Interaktives Zeiterfassungs-Tool

CLI-basierte Zeiterfassung mit SQLite-Backend, Rich-TUI und Excel-Export.

## Quickstart

```bash
# Venv aktivieren
source .venv/bin/activate

# Einstempeln
stamp in

# Ausstempeln
stamp out

# Status heute
stamp status
```

## Alle Commands

```
stamp in [--time HH:MM]              Einstempeln
stamp out [--time HH:MM]             Ausstempeln
stamp status                          Heutiger Status
stamp today                           Tagesdetails
stamp week                            Wochenübersicht
stamp month [MONAT]                   Monatsübersicht
stamp overtime                        Überstunden-Stand
stamp vacation <datum> [--to <datum>] Urlaub eintragen
stamp vacation-left                   Resturlaub anzeigen
stamp sick <datum> [--to <datum>]     Kranktag eintragen
stamp flex <datum>                    Gleittag eintragen
stamp travel <datum> [--to <datum>]   Dienstreise eintragen
stamp cancel <datum>                  Eintrag löschen
stamp edit <datum> [--in] [--out]     Tag bearbeiten
stamp pause <minuten>                 Pause setzen
stamp check                           Fehlende Einträge prüfen
stamp config show                     Config anzeigen
stamp config set <key> <value>        Config ändern
```

## Datumsformate

- `DD.MM.YYYY` (z.B. `26.03.2026`)
- `YYYY-MM-DD` (z.B. `2026-03-26`)
- `DD.MM.` (aktuelles Jahr, z.B. `26.03.`)

## Konfiguration

| Schlüssel | Default | Beschreibung |
|---|---|---|
| `target_hours` | 8.0 | Soll-Stunden/Tag |
| `default_pause` | 0.75 | Pause in Stunden |
| `vacation_days` | 30 | Urlaubstage/Jahr |
| `federal_state` | BY | Bundesland (Feiertage) |
| `overtime_carryover` | 0.0 | Überstunden-Übertrag VJ |
| `vacation_carryover` | 0 | Urlaubs-Übertrag VJ |

## Tech-Stack

Python 3.10+ | Typer | Rich | SQLAlchemy | SQLite | openpyxl