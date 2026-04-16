"""Feiertage-API Integration mit SQLite-Caching."""

from datetime import date
import requests

from stamp.db import get_session, Holiday, get_config, get_config_int


def fetch_holidays(year: int | None = None) -> list[Holiday]:
    """Holt Feiertage aus der API oder dem Cache.

    Wird nur 1x pro Jahr von der API abgerufen, danach aus SQLite.
    """
    if year is None:
        year = get_config_int("year")

    with get_session() as session:
        cached = session.query(Holiday).filter_by(year=year).all()
        if cached:
            return cached

    state = get_config("federal_state", "BY")
    filter_str = get_config("holiday_filter", "")
    filter_dates = {f.strip() for f in filter_str.split(",") if f.strip()}

    try:
        url = f"https://feiertage-api.de/api/?jahr={year}&nur_land={state}"
        response = requests.get(url, timeout=10)
        response.raise_for_status()
        data = response.json()
    except (requests.RequestException, ValueError):
        return []

    holidays = []
    with get_session() as session:
        for name, info in data.items():
            datum_str = info.get("datum", "")
            if not datum_str:
                continue
            # Filter anwenden (MM-DD Format)
            mm_dd = datum_str[5:]  # "2026-01-01" -> "01-01"
            if mm_dd in filter_dates:
                continue
            holiday_date = date.fromisoformat(datum_str)
            holiday = Holiday(date=holiday_date, name=name, year=year)
            session.merge(holiday)
            holidays.append(holiday)
        session.commit()

    return holidays


def is_holiday(check_date: date) -> tuple[bool, str | None]:
    """Prüft ob ein Datum ein Feiertag ist. Gibt (True, Name) oder (False, None) zurück."""
    holidays = fetch_holidays(check_date.year)
    for h in holidays:
        if h.date == check_date:
            return True, h.name
    return False, None


def get_holidays_for_year(year: int | None = None) -> dict[date, str]:
    """Gibt alle Feiertage eines Jahres als Dict {datum: name} zurück."""
    holidays = fetch_holidays(year)
    return {h.date: h.name for h in holidays}
