"""Konfiguration für das Zeiterfassungs-Tool."""

import os
from pathlib import Path

# Pfad zur Datenbank — im gleichen Verzeichnis wie das Script
BASE_DIR = Path(__file__).resolve().parent.parent
DB_PATH = BASE_DIR / "stamp.db"
DB_URL = f"sqlite:///{DB_PATH}"

# Defaults
DEFAULTS = {
    "target_hours": "8.0",
    "default_pause": "0.75",
    "vacation_days": "30",
    "federal_state": "BY",
    "holiday_filter": "08-08,08-15,11-19",
    "default_start": "07:30",
    "default_end": "16:15",
    "warn_max_hours": "10.0",
    "year": str(__import__("datetime").date.today().year),
    "vacation_carryover": "0",
    "overtime_carryover": "0.0",
}
