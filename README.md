# stamp ⏱️ — Interaktives Zeiterfassungs-Tool

CLI-basierte Zeiterfassung mit SQLite-Backend, Rich-TUI und Excel-Export.

## Mehrere Tage in der Web-GUI erfassen

Über **Mehrere Tage erfassen** im Dashboard oder in der Monatsansicht lässt sich
ein Zeitraum **von/bis einschließlich** eintragen. Unterstützt werden Arbeit,
Urlaub, Krank, Gleittag und Dienstreise. Wochenenden und konfigurierte Feiertage
werden automatisch ausgelassen, auch über Monats- und Jahresgrenzen hinweg.

Für einzelne, nicht zusammenhängende Tage in der Monatsansicht **Mehrere Tage
auswählen** aktivieren und die gewünschten Checkboxen markieren. Die Checkbox im
Tabellenkopf wählt alle Arbeitstage des angezeigten Monats. Die Auswahl bleibt
beim Monatswechsel erhalten. **Ausgewählte Tage erfassen** öffnet den gemeinsamen
Erfassungsdialog; dort lassen sich einzelne Daten wieder entfernen.

Tagesart und Hinweis gelten für alle ausgewählten Tage. Bei **Arbeit** gelten
auch Kommen, Gehen und Pause gemeinsam; vorbelegt werden die konfigurierten
Standardwerte. Die Erfassung unterstützt Arbeitszeiten innerhalb eines Tages.
Gleittage werden mit den konfigurierten Sollstunden vom Überstundenkonto abgezogen.

Vor dem Speichern zeigt **Vorschau anzeigen** die betroffenen und ausgelassenen
Tage. Vorhandene Einträge müssen ausdrücklich mit **Vorhandene Einträge
überschreiben** bestätigt werden. Dabei werden die Tageswerte vollständig
ersetzt; ein leerer Hinweis entfernt einen vorhandenen Hinweis. Erst **Jetzt
speichern** schreibt die Auswahl gemeinsam. Abbrechen verändert keine Einträge.
Ändert sich der Datenbestand zwischen Vorschau und Speichern, ist eine neue
Vorschau erforderlich. Sind Feiertagsdaten nicht verfügbar, wird nicht gespeichert.
Die bisherige Einzelbearbeitung bleibt außerhalb des Auswahlmodus unverändert.

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