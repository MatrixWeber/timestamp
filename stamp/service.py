"""Business-Logik für die Zeiterfassung."""

from datetime import date, time, datetime, timedelta
from typing import Optional
import hashlib
import json
import math

from sqlalchemy import text

from stamp.db import get_session, Stamp, get_config_float, get_config_int
from stamp.holidays import is_holiday, fetch_holidays


# --- Berechnungen ---

def calculate_work_hours(stamp_in: time, stamp_out: time, pause: float) -> float:
    """Berechnet Arbeitszeit in Stunden (Gehzeit - Kommzeit - Pause)."""
    dt_in = datetime.combine(date.today(), stamp_in)
    dt_out = datetime.combine(date.today(), stamp_out)
    diff = (dt_out - dt_in).total_seconds() / 3600
    return round(diff - pause, 2)


def calculate_overtime(work_hours: float, target: float | None = None) -> float:
    """Berechnet Überstunden (Arbeitszeit - Soll)."""
    if target is None:
        target = get_config_float("target_hours")
    return round(work_hours - target, 2)


def is_workday(check_date: date) -> bool:
    """Prüft ob ein Datum ein Arbeitstag ist (kein WE, kein Feiertag)."""
    if check_date.weekday() >= 5:  # Samstag=5, Sonntag=6
        return False
    holiday, _ = is_holiday(check_date)
    return not holiday


# --- Stempel-Operationen ---

def stamp_in(stamp_time: time | None = None, for_date: date | None = None) -> Stamp:
    """Einstempeln. Verwendet aktuelle Zeit wenn keine angegeben."""
    if stamp_time is None:
        stamp_time = datetime.now().time().replace(microsecond=0, second=0)
    if for_date is None:
        for_date = date.today()

    default_pause = get_config_float("default_pause")

    with get_session() as session:
        existing = session.query(Stamp).filter_by(date=for_date).first()
        if existing:
            if existing.stamp_in and existing.type == "WORK":
                raise ValueError(
                    f"Bereits eingestempelt um {existing.stamp_in.strftime('%H:%M')}. "
                    f"Nutze 'stamp edit {for_date}' zum Ändern."
                )
            # Abwesenheit überschreiben falls nötig
            existing.stamp_in = stamp_time
            existing.type = "WORK"
            existing.pause = default_pause
            existing.updated_at = datetime.now()
            session.commit()
            session.refresh(existing)
            return existing

        entry = Stamp(
            date=for_date,
            stamp_in=stamp_time,
            pause=default_pause,
            type="WORK",
        )
        session.add(entry)
        session.commit()
        session.refresh(entry)
        return entry


def stamp_out(stamp_time: time | None = None, for_date: date | None = None) -> Stamp:
    """Ausstempeln. Berechnet automatisch Arbeitszeit und Überstunden."""
    if stamp_time is None:
        stamp_time = datetime.now().time().replace(microsecond=0, second=0)
    if for_date is None:
        for_date = date.today()

    with get_session() as session:
        existing = session.query(Stamp).filter_by(date=for_date).first()
        if not existing or not existing.stamp_in:
            raise ValueError(
                f"Kein Einstempel-Eintrag für {for_date}. Erst 'stamp in' ausführen."
            )
        if existing.stamp_out:
            raise ValueError(
                f"Bereits ausgestempelt um {existing.stamp_out.strftime('%H:%M')}. "
                f"Nutze 'stamp edit {for_date}' zum Ändern."
            )

        existing.stamp_out = stamp_time
        existing.work_hours = calculate_work_hours(
            existing.stamp_in, stamp_time, existing.pause
        )
        existing.overtime = calculate_overtime(existing.work_hours)
        existing.updated_at = datetime.now()
        session.commit()
        session.refresh(existing)
        return existing


def set_pause(minutes: float, for_date: date | None = None) -> Stamp:
    """Setzt die Pausendauer für einen Tag. Aktualisiert Arbeitszeit/Überstunden."""
    if for_date is None:
        for_date = date.today()

    pause_hours = round(minutes / 60, 4)

    with get_session() as session:
        existing = session.query(Stamp).filter_by(date=for_date).first()
        if not existing:
            raise ValueError(f"Kein Eintrag für {for_date}.")

        existing.pause = pause_hours
        if existing.stamp_in and existing.stamp_out:
            existing.work_hours = calculate_work_hours(
                existing.stamp_in, existing.stamp_out, pause_hours
            )
            existing.overtime = calculate_overtime(existing.work_hours)
        existing.updated_at = datetime.now()
        session.commit()
        session.refresh(existing)
        return existing


def edit_stamp(
    for_date: date,
    stamp_in_time: time | None = None,
    stamp_out_time: time | None = None,
    pause: float | None = None,
    entry_type: str | None = None,
    note: str | None = None,
) -> Stamp:
    """Bearbeitet einen bestehenden oder erstellt einen neuen Eintrag."""
    with get_session() as session:
        existing = session.query(Stamp).filter_by(date=for_date).first()
        if not existing:
            existing = Stamp(date=for_date, pause=get_config_float("default_pause"))
            session.add(existing)

        if stamp_in_time is not None:
            existing.stamp_in = stamp_in_time
        if stamp_out_time is not None:
            existing.stamp_out = stamp_out_time
        if pause is not None:
            existing.pause = pause
        if entry_type is not None:
            existing.type = entry_type.upper()
        if note is not None:
            existing.note = note

        # Arbeitszeit neu berechnen wenn beide Zeiten vorhanden
        if existing.stamp_in and existing.stamp_out and existing.type == "WORK":
            existing.work_hours = calculate_work_hours(
                existing.stamp_in, existing.stamp_out, existing.pause
            )
            existing.overtime = calculate_overtime(existing.work_hours)
        elif existing.type in ("VACATION", "SICK", "FLEX", "TRAVEL"):
            target = get_config_float("target_hours")
            existing.work_hours = target
            existing.overtime = 0.0

        existing.updated_at = datetime.now()
        session.commit()
        session.refresh(existing)
        return existing


def _apply_absence(entry: Stamp, absence_type: str, target: float, note: str | None) -> None:
    entry.type = absence_type
    entry.stamp_in = None
    entry.stamp_out = None
    entry.pause = 0.0
    entry.work_hours = target
    entry.overtime = -target if absence_type == "FLEX" else 0.0
    entry.note = note
    entry.updated_at = datetime.now()


def add_absence(
    absence_type: str,
    start_date: date,
    end_date: date | None = None,
    note: str | None = None,
) -> list[Stamp]:
    """Trägt Abwesenheit ein (Urlaub, Krank, Gleittag, Dienstreise)."""
    if end_date is None:
        end_date = start_date

    valid_types = {"VACATION", "SICK", "FLEX", "TRAVEL"}
    absence_type = absence_type.upper()
    if absence_type not in valid_types:
        raise ValueError(f"Ungültiger Typ: {absence_type}. Erlaubt: {valid_types}")

    target = get_config_float("target_hours")
    entries = []
    current = start_date

    with get_session() as session:
        while current <= end_date:
            if is_workday(current):
                existing = session.query(Stamp).filter_by(date=current).first()
                if not existing:
                    existing = Stamp(date=current)
                    session.add(existing)
                _apply_absence(existing, absence_type, target, note)
                entries.append(existing)
            current += timedelta(days=1)
        session.commit()
        for e in entries:
            session.refresh(e)

    return entries


class BatchConflictError(ValueError):
    """Die bestaetigte Vorschau stimmt nicht mehr mit dem Datenbestand ueberein."""


def _prepare_batch(
    entry_type: str,
    start_date: date | None = None,
    end_date: date | None = None,
    dates: list[date] | None = None,
    stamp_in_time: time | None = None,
    stamp_out_time: time | None = None,
    pause: float = 0,
    note: str | None = None,
) -> dict:
    if entry_type not in {"WORK", "VACATION", "SICK", "FLEX", "TRAVEL"}:
        raise ValueError("Ungueltige Tagesart.")
    if dates is not None:
        if start_date is not None or end_date is not None:
            raise ValueError("Entweder Zeitraum oder einzelne Daten angeben.")
        selected = sorted(set(dates))
    else:
        if start_date is None or end_date is None or end_date < start_date:
            raise ValueError("Gueltigen Zeitraum von/bis angeben.")
        selected = [
            start_date + timedelta(days=i)
            for i in range((end_date - start_date).days + 1)
        ]
    if not selected:
        raise ValueError("Mindestens einen Tag auswaehlen.")
    if entry_type == "WORK":
        if stamp_in_time is None or stamp_out_time is None or stamp_out_time <= stamp_in_time:
            raise ValueError("Kommen und Gehen angeben; Gehen muss nach Kommen liegen.")
        if not math.isfinite(pause) or pause < 0:
            raise ValueError("Die Pause muss eine nichtnegative Minutenzahl sein.")
        duration = (
            datetime.combine(date.today(), stamp_out_time)
            - datetime.combine(date.today(), stamp_in_time)
        ).total_seconds() / 60
        if pause >= duration:
            raise ValueError("Die Pause muss kuerzer als die Anwesenheit sein.")
    else:
        stamp_in_time, stamp_out_time, pause = None, None, 0

    holidays = {
        h.date: h.name
        for year in sorted({d.year for d in selected})
        for h in fetch_holidays(year, strict=True)
    }
    applicable, skipped = [], []
    for day in selected:
        if day.weekday() >= 5 or day in holidays:
            skipped.append({
                "date": day,
                "reason": holidays.get(day, "Wochenende"),
            })
        else:
            applicable.append(day)
    if not applicable:
        raise ValueError("Die Auswahl enthaelt keine Arbeitstage.")
    return {
        "dates": applicable, "skipped": skipped, "type": entry_type,
        "stamp_in": stamp_in_time, "stamp_out": stamp_out_time,
        "pause": pause / 60, "note": note or None,
        "target": get_config_float("target_hours"),
    }


def _batch_preview(session, batch: dict) -> tuple[dict, dict[date, Stamp]]:
    # Query by interval to avoid SQLite's parameter limit for long selections.
    selected = set(batch["dates"])
    existing = {
        e.date: e for e in session.query(Stamp).filter(
            Stamp.date >= batch["dates"][0], Stamp.date <= batch["dates"][-1]
        ).order_by(Stamp.date).all() if e.date in selected
    }
    snapshot = [
        {column.name: getattr(e, column.name) for column in Stamp.__table__.columns}
        for e in existing.values()
    ]
    token = hashlib.sha256(json.dumps(
        {"batch": batch, "existing": snapshot}, default=str, sort_keys=True
    ).encode()).hexdigest()
    return {
        "dates": batch["dates"], "skipped": batch["skipped"],
        "conflicts": list(existing.values()), "preview_token": token,
    }, existing


def preview_batch(**kwargs) -> dict:
    batch = _prepare_batch(**kwargs)
    with get_session() as session:
        preview, _ = _batch_preview(session, batch)
        return preview


def save_batch(*, preview_token: str, overwrite: bool = False, **kwargs) -> dict:
    batch = _prepare_batch(**kwargs)
    with get_session() as session:
        # SQLite write reservation prevents a change between conflict check and write.
        session.execute(text("BEGIN IMMEDIATE"))
        preview, existing = _batch_preview(session, batch)
        if preview["preview_token"] != preview_token:
            raise BatchConflictError("Eintraege oder Eingaben wurden geaendert. Bitte Vorschau erneut laden.")
        if existing and not overwrite:
            raise BatchConflictError("Vorhandene Eintraege muessen ausdruecklich zum Ueberschreiben bestaetigt werden.")
        entries = []
        for day in batch["dates"]:
            entry = existing.get(day)
            if entry is None:
                entry = Stamp(date=day)
                session.add(entry)
            if batch["type"] == "WORK":
                entry.type = "WORK"
                entry.stamp_in = batch["stamp_in"]
                entry.stamp_out = batch["stamp_out"]
                entry.pause = batch["pause"]
                entry.work_hours = calculate_work_hours(entry.stamp_in, entry.stamp_out, entry.pause)
                entry.overtime = calculate_overtime(entry.work_hours, batch["target"])
                entry.note = batch["note"]
                entry.updated_at = datetime.now()
            else:
                _apply_absence(entry, batch["type"], batch["target"], batch["note"])
            entries.append(entry)
        session.flush()
        # Keep the committed response independent of subsequent writers.
        session.expunge_all()
        session.commit()
        return {
            "created": len(entries) - len(existing), "replaced": len(existing),
            "skipped": batch["skipped"], "entries": entries,
        }


def cancel_entry(for_date: date) -> bool:
    """Löscht einen Eintrag."""
    with get_session() as session:
        existing = session.query(Stamp).filter_by(date=for_date).first()
        if not existing:
            return False
        session.delete(existing)
        session.commit()
        return True


# --- Abfragen ---

def get_today() -> Stamp | None:
    """Gibt den heutigen Eintrag zurück."""
    with get_session() as session:
        return session.query(Stamp).filter_by(date=date.today()).first()


def get_day(for_date: date) -> Stamp | None:
    """Gibt einen bestimmten Tag zurück."""
    with get_session() as session:
        return session.query(Stamp).filter_by(date=for_date).first()


def get_week(ref_date: date | None = None) -> list[Stamp]:
    """Gibt alle Einträge der aktuellen Woche zurück."""
    if ref_date is None:
        ref_date = date.today()
    start = ref_date - timedelta(days=ref_date.weekday())  # Montag
    end = start + timedelta(days=4)  # Freitag
    with get_session() as session:
        return (
            session.query(Stamp)
            .filter(Stamp.date >= start, Stamp.date <= end)
            .order_by(Stamp.date)
            .all()
        )


def get_month(month: int | None = None, year: int | None = None) -> list[Stamp]:
    """Gibt alle Einträge eines Monats zurück."""
    if month is None:
        month = date.today().month
    if year is None:
        year = date.today().year
    start = date(year, month, 1)
    if month == 12:
        end = date(year + 1, 1, 1) - timedelta(days=1)
    else:
        end = date(year, month + 1, 1) - timedelta(days=1)
    with get_session() as session:
        return (
            session.query(Stamp)
            .filter(Stamp.date >= start, Stamp.date <= end)
            .order_by(Stamp.date)
            .all()
        )


def get_year(year: int | None = None) -> list[Stamp]:
    """Gibt alle Einträge eines Jahres zurück."""
    if year is None:
        year = date.today().year
    start = date(year, 1, 1)
    end = date(year, 12, 31)
    with get_session() as session:
        return (
            session.query(Stamp)
            .filter(Stamp.date >= start, Stamp.date <= end)
            .order_by(Stamp.date)
            .all()
        )


def get_overtime_total(year: int | None = None) -> dict:
    """Berechnet den kumulierten Überstunden-Stand (nur bis einschl. heute)."""
    if year is None:
        year = date.today().year

    today = date.today()
    entries = get_year(year)
    carryover = float(get_config_float("overtime_carryover"))

    # Only count entries up to and including today
    past_entries = [e for e in entries if e.date <= today]

    total = carryover
    month_totals = {}

    for entry in past_entries:
        if entry.overtime is not None:
            total += entry.overtime
            m = entry.date.month
            month_totals[m] = month_totals.get(m, 0) + entry.overtime

    # Heute berechnen
    today_entry = get_today()
    today_overtime = today_entry.overtime if today_entry and today_entry.overtime else 0

    # Diese Woche
    week_entries = get_week()
    week_overtime = sum(e.overtime for e in week_entries if e.overtime and e.date <= today)

    # Dieser Monat
    current_month = today.month
    month_overtime = month_totals.get(current_month, 0)

    return {
        "today": today_overtime,
        "week": week_overtime,
        "month": month_overtime,
        "year": total - carryover,
        "carryover": carryover,
        "total": total,
    }


def get_vacation_balance(year: int | None = None) -> dict:
    """Berechnet den Resturlaub mit Trennung genommen/geplant."""
    if year is None:
        year = date.today().year

    today = date.today()
    total_days = get_config_int("vacation_days")
    carryover = get_config_int("vacation_carryover")
    entries = get_year(year)

    # Split into past (taken) and future (planned)
    def count_type(typ: str, past_only: bool = False, future_only: bool = False):
        return sum(1 for e in entries
                   if e.type == typ
                   and (not past_only or e.date <= today)
                   and (not future_only or e.date > today))

    vacation_taken = count_type("VACATION", past_only=True)
    vacation_planned = count_type("VACATION", future_only=True)
    vacation_total_used = vacation_taken + vacation_planned

    flex_taken = count_type("FLEX", past_only=True)
    flex_planned = count_type("FLEX", future_only=True)

    sick_days = sum(1 for e in entries if e.type == "SICK")
    travel_taken = count_type("TRAVEL", past_only=True)
    travel_planned = count_type("TRAVEL", future_only=True)

    remaining = total_days + carryover - vacation_total_used

    return {
        "total": total_days,
        "carryover": carryover,
        "taken": vacation_taken,
        "planned": vacation_planned,
        "remaining": remaining,
        "flex_taken": flex_taken,
        "flex_planned": flex_planned,
        "sick_days": sick_days,
        "travel_taken": travel_taken,
        "travel_planned": travel_planned,
    }


def get_missing_days(year: int | None = None, month: int | None = None) -> list[date]:
    """Findet Arbeitstage ohne Stempel-Eintrag."""
    if year is None:
        year = date.today().year

    today = date.today()
    start = date(year, month or 1, 1)
    if month:
        import calendar
        last_day = calendar.monthrange(year, month)[1]
        end = min(date(year, month, last_day), today - timedelta(days=1))
    else:
        end = today - timedelta(days=1)

    with get_session() as session:
        existing_dates = {
            s.date for s in session.query(Stamp.date).filter(
                Stamp.date >= start, Stamp.date <= end
            ).all()
        }

    missing = []
    current = start
    while current <= end:
        if is_workday(current) and current not in existing_dates:
            missing.append(current)
        current += timedelta(days=1)

    return missing
