"""API-Routen für das Zeiterfassungs-Tool."""

from datetime import date, time, datetime, timedelta
import calendar
import shutil
from pathlib import Path

from fastapi import APIRouter, HTTPException, UploadFile, File
from fastapi.responses import FileResponse

from stamp import service
from stamp.db import get_session, Config, Stamp
from stamp.holidays import is_holiday, get_holidays_for_year
from stamp.api.schemas import (
    StampResponse, StampInRequest, StampOutRequest, AbsenceRequest,
    EditRequest, OvertimeResponse, VacationResponse, ConfigEntry, DayInfo,
)

router = APIRouter(prefix="/api")

WEEKDAYS_DE = ["Montag", "Dienstag", "Mittwoch", "Donnerstag", "Freitag", "Samstag", "Sonntag"]


def _parse_time(time_str: str | None) -> time | None:
    if not time_str:
        return None
    return datetime.strptime(time_str, "%H:%M").time()


def _stamp_to_response(s: Stamp) -> StampResponse:
    return StampResponse(
        id=s.id, date=s.date, stamp_in=s.stamp_in, stamp_out=s.stamp_out,
        pause=s.pause, work_hours=s.work_hours, overtime=s.overtime,
        type=s.type, note=s.note,
    )


# --- Stempel ---

@router.post("/stamp/in", response_model=StampResponse)
def api_stamp_in(req: StampInRequest):
    try:
        entry = service.stamp_in(_parse_time(req.time))
        return _stamp_to_response(entry)
    except ValueError as e:
        raise HTTPException(400, str(e))


@router.post("/stamp/out", response_model=StampResponse)
def api_stamp_out(req: StampOutRequest):
    try:
        entry = service.stamp_out(_parse_time(req.time))
        return _stamp_to_response(entry)
    except ValueError as e:
        raise HTTPException(400, str(e))


# --- Abfragen ---

@router.get("/today", response_model=DayInfo)
def api_today():
    today = date.today()
    holiday, hname = is_holiday(today)
    entry = service.get_today()
    return DayInfo(
        date=today,
        weekday=WEEKDAYS_DE[today.weekday()],
        is_holiday=holiday,
        holiday_name=hname,
        is_weekend=today.weekday() >= 5,
        stamp=_stamp_to_response(entry) if entry else None,
    )


@router.get("/week")
def api_week():
    today = date.today()
    monday = today - timedelta(days=today.weekday())
    entries = service.get_week()
    entry_map = {e.date: e for e in entries}
    holidays = get_holidays_for_year(today.year)

    days = []
    for i in range(5):
        day = monday + timedelta(days=i)
        holiday = day in holidays
        entry = entry_map.get(day)
        days.append(DayInfo(
            date=day,
            weekday=WEEKDAYS_DE[day.weekday()],
            is_holiday=holiday,
            holiday_name=holidays.get(day),
            is_weekend=False,
            stamp=_stamp_to_response(entry) if entry else None,
        ))

    week_entries = [e for e in entries if e.overtime is not None]
    return {
        "kw": today.isocalendar()[1],
        "start": monday.isoformat(),
        "end": (monday + timedelta(days=4)).isoformat(),
        "days": days,
        "total_work": sum(e.work_hours for e in entries if e.work_hours),
        "total_overtime": sum(e.overtime for e in week_entries),
    }


@router.get("/month/{month}")
def api_month(month: int, year: int | None = None):
    if year is None:
        year = date.today().year
    if not 1 <= month <= 12:
        raise HTTPException(400, "Monat muss zwischen 1 und 12 sein")

    num_days = calendar.monthrange(year, month)[1]
    entries = service.get_month(month, year)
    entry_map = {e.date: e for e in entries}
    holidays = get_holidays_for_year(year)

    days = []
    for day_num in range(1, num_days + 1):
        day = date(year, month, day_num)
        holiday = day in holidays
        entry = entry_map.get(day)
        days.append(DayInfo(
            date=day,
            weekday=WEEKDAYS_DE[day.weekday()],
            is_holiday=holiday,
            holiday_name=holidays.get(day),
            is_weekend=day.weekday() >= 5,
            stamp=_stamp_to_response(entry) if entry else None,
        ))

    work_entries = [e for e in entries if e.work_hours]
    ot_entries = [e for e in entries if e.overtime is not None]

    # Month-level absence counts
    vacation_days = sum(1 for e in entries if e.type == "VACATION")
    flex_days = sum(1 for e in entries if e.type == "FLEX")
    sick_days = sum(1 for e in entries if e.type == "SICK")
    travel_days = sum(1 for e in entries if e.type == "TRAVEL")
    work_days = sum(1 for e in entries if e.type == "WORK" and e.work_hours)

    return {
        "month": month,
        "year": year,
        "days": days,
        "total_work": sum(e.work_hours for e in work_entries),
        "total_overtime": sum(e.overtime for e in ot_entries),
        "stats": {
            "work_days": work_days,
            "vacation_days": vacation_days,
            "flex_days": flex_days,
            "sick_days": sick_days,
            "travel_days": travel_days,
        },
    }


@router.get("/overtime", response_model=OvertimeResponse)
def api_overtime():
    return service.get_overtime_total()


@router.get("/vacation", response_model=VacationResponse)
def api_vacation():
    return service.get_vacation_balance()


@router.get("/missing")
def api_missing():
    missing = service.get_missing_days()
    return [{"date": d.isoformat(), "weekday": WEEKDAYS_DE[d.weekday()]} for d in missing]


# --- Abwesenheiten ---

@router.post("/absence")
def api_absence(req: AbsenceRequest):
    try:
        entries = service.add_absence(req.type, req.start_date, req.end_date, req.note)
        return {"count": len(entries), "entries": [_stamp_to_response(e) for e in entries]}
    except ValueError as e:
        raise HTTPException(400, str(e))


# --- Edit / Cancel ---

@router.put("/stamp/{stamp_date}")
def api_edit(stamp_date: date, req: EditRequest):
    try:
        entry = service.edit_stamp(
            stamp_date,
            stamp_in_time=_parse_time(req.stamp_in),
            stamp_out_time=_parse_time(req.stamp_out),
            pause=req.pause / 60 if req.pause is not None else None,
            entry_type=req.type,
            note=req.note,
        )
        return _stamp_to_response(entry)
    except ValueError as e:
        raise HTTPException(400, str(e))


@router.delete("/stamp/{stamp_date}")
def api_cancel(stamp_date: date):
    if service.cancel_entry(stamp_date):
        return {"deleted": True}
    raise HTTPException(404, "Kein Eintrag gefunden")


# --- Config ---

@router.get("/config")
def api_config():
    with get_session() as session:
        entries = session.query(Config).order_by(Config.key).all()
        return [ConfigEntry(key=e.key, value=e.value) for e in entries]


# --- Export ---

@router.get("/export/excel")
def api_export_excel(year: int | None = None):
    from stamp.excel_export import export_excel
    if year is None:
        year = date.today().year
    path = export_excel(year, f"/tmp/Stundenschreibung{year}.xlsx")
    return FileResponse(str(path), filename=f"Stundenschreibung{year}.xlsx",
                        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


# --- Import ---

@router.post("/import/excel")
async def api_import_excel(file: UploadFile = File(...), year: int | None = None):
    from stamp.data_migration import import_excel
    if year is None:
        year = date.today().year
    tmp_path = Path(f"/tmp/stamp_import_{year}.xlsx")
    try:
        with open(tmp_path, "wb") as f:
            shutil.copyfileobj(file.file, f)
        stats = import_excel(str(tmp_path), year)
        return {"success": True, **stats}
    except Exception as e:
        raise HTTPException(400, f"Import fehlgeschlagen: {e}")
    finally:
        tmp_path.unlink(missing_ok=True)
