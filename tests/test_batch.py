from datetime import date, time
from unittest.mock import Mock

import pytest
import requests
from fastapi import FastAPI
from fastapi.testclient import TestClient
from sqlalchemy import create_engine, event
from sqlalchemy.exc import OperationalError
from sqlalchemy.orm import sessionmaker

from stamp import db, service, holidays
from stamp.api.routes import router
from stamp.config import DEFAULTS


@pytest.fixture
def database(tmp_path, monkeypatch):
    engine = create_engine(f"sqlite:///{tmp_path / 'test.db'}")
    factory = sessionmaker(bind=engine)
    monkeypatch.setattr(db, "SessionLocal", factory)
    db.Base.metadata.create_all(engine)
    with factory() as session:
        session.add_all(db.Config(key=key, value=value) for key, value in DEFAULTS.items())
        session.add_all([
            db.Holiday(date=date(2026, 1, 1), name="Neujahr", year=2026),
            db.Holiday(date=date(2026, 10, 26), name="Testfeiertag", year=2026),
            db.Holiday(date=date(2027, 1, 1), name="Neujahr", year=2027),
        ])
        session.commit()
    yield factory
    engine.dispose()


@pytest.fixture
def client(database):
    app = FastAPI()
    app.include_router(router)
    with TestClient(app) as client:
        yield client


def values(entry_type="VACATION", **extra):
    return {"entry_type": entry_type, "dates": [date(2026, 10, 20), date(2026, 10, 21)], **extra}


def write(data, **extra):
    preview = service.preview_batch(**data)
    return service.save_batch(preview_token=preview["preview_token"], **data, **extra)


def test_range_filters_weekends_and_holidays_inclusively(database):
    data = {
        "entry_type": "VACATION", "start_date": date(2026, 10, 23),
        "end_date": date(2026, 10, 27),
    }
    preview = service.preview_batch(**data)
    assert preview["dates"] == [date(2026, 10, 23), date(2026, 10, 27)]
    assert [day["reason"] for day in preview["skipped"]] == ["Wochenende", "Wochenende", "Testfeiertag"]
    result = service.save_batch(preview_token=preview["preview_token"], **data)
    assert result["created"] == 2
    assert result["replaced"] == 0
    with database() as session:
        assert session.query(db.Stamp).count() == 2


def test_explicit_dates_are_deduplicated_and_sorted(database):
    data = values(dates=[date(2026, 10, 21), date(2026, 10, 20), date(2026, 10, 21)])
    assert [entry.date for entry in write(data)["entries"]] == [date(2026, 10, 20), date(2026, 10, 21)]


def test_month_and_year_boundaries(database):
    preview = service.preview_batch(
        entry_type="SICK", start_date=date(2026, 12, 31), end_date=date(2027, 1, 4)
    )
    assert preview["dates"] == [date(2026, 12, 31), date(2027, 1, 4)]
    assert len(preview["skipped"]) == 3


@pytest.mark.parametrize("entry_type", ["WORK", "VACATION", "SICK", "FLEX", "TRAVEL"])
def test_all_types_and_configured_calculations(database, entry_type):
    db.set_config("target_hours", "7.5")
    data = values(entry_type, stamp_in_time=time(8), stamp_out_time=time(17), pause=30, note="Gemeinsam")
    for entry in write(data)["entries"]:
        assert entry.type == entry_type
        assert entry.note == "Gemeinsam"
        assert entry.work_hours == (8.5 if entry_type == "WORK" else 7.5)
        assert entry.overtime == (1 if entry_type == "WORK" else -7.5 if entry_type == "FLEX" else 0)
        assert entry.pause == (0.5 if entry_type == "WORK" else 0)
        assert entry.stamp_in == (time(8) if entry_type == "WORK" else None)
        assert entry.stamp_out == (time(17) if entry_type == "WORK" else None)


def test_conflicts_require_confirmation_and_replace_all_fields(database):
    write(values("WORK", stamp_in_time=time(8), stamp_out_time=time(17), pause=45, note="Alt"))
    data = values("FLEX", note="")
    preview = service.preview_batch(**data)
    assert len(preview["conflicts"]) == 2
    with pytest.raises(service.BatchConflictError, match="ausdruecklich"):
        service.save_batch(preview_token=preview["preview_token"], **data)
    assert service.get_day(date(2026, 10, 20)).type == "WORK"
    result = service.save_batch(preview_token=preview["preview_token"], overwrite=True, **data)
    assert result["created"] == 0
    assert result["replaced"] == 2
    entry = service.get_day(date(2026, 10, 20))
    assert entry.type == "FLEX"
    assert entry.stamp_in is None and entry.stamp_out is None and entry.note is None
    assert entry.pause == 0 and entry.overtime == -8


@pytest.mark.parametrize("change", ["insert", "update", "delete", "payload"])
def test_stale_preview_writes_nothing(database, change):
    data = values()
    if change in {"update", "delete"}:
        write(values("SICK"))
    preview = service.preview_batch(**data)
    if change == "insert":
        service.edit_stamp(date(2026, 10, 20), note="Zwischenzeitlich")
    elif change == "update":
        service.edit_stamp(date(2026, 10, 20), note="Zwischenzeitlich")
    elif change == "delete":
        service.cancel_entry(date(2026, 10, 20))
    else:
        data["note"] = "Andere Eingabe"
    with database() as session:
        before = [(e.date, e.type, e.note) for e in session.query(db.Stamp).order_by(db.Stamp.date)]
    with pytest.raises(service.BatchConflictError, match="Vorschau"):
        service.save_batch(preview_token=preview["preview_token"], overwrite=True, **data)
    with database() as session:
        assert [(e.date, e.type, e.note) for e in session.query(db.Stamp).order_by(db.Stamp.date)] == before


def test_preview_and_cancel_do_not_write(database):
    service.preview_batch(**values())
    with database() as session:
        assert session.query(db.Stamp).count() == 0


@pytest.mark.parametrize("data", [
    values("INVALID"),
    values(dates=[]),
    values(dates=[date(2026, 10, 24), date(2026, 10, 26)]),
    values(start_date=date(2026, 10, 20)),
    {"entry_type": "SICK", "start_date": date(2026, 10, 21), "end_date": date(2026, 10, 20)},
    {"entry_type": "SICK", "start_date": date(2026, 10, 20)},
    values("WORK"),
    values("WORK", stamp_in_time=time(9), stamp_out_time=time(8)),
    values("WORK", stamp_in_time=time(8), stamp_out_time=time(9), pause=60),
    values("WORK", stamp_in_time=time(8), stamp_out_time=time(9), pause=-1),
    values("WORK", stamp_in_time=time(8), stamp_out_time=time(9), pause=float("nan")),
])
def test_invalid_requests_write_nothing(database, data):
    with pytest.raises(ValueError):
        service.preview_batch(**data)
    with database() as session:
        assert session.query(db.Stamp).count() == 0


def test_rollback_on_flush_failure(database):
    write(values("SICK", dates=[date(2026, 10, 20)]))
    data = values("TRAVEL")
    preview = service.preview_batch(**data)

    def fail_flush(*args):
        raise RuntimeError("Simulierter Speicherfehler")

    event.listen(database.class_, "before_flush", fail_flush)
    try:
        with pytest.raises(RuntimeError, match="Speicherfehler"):
            service.save_batch(preview_token=preview["preview_token"], overwrite=True, **data)
    finally:
        event.remove(database.class_, "before_flush", fail_flush)
    with database() as session:
        entries = session.query(db.Stamp).all()
        assert len(entries) == 1 and entries[0].type == "SICK"


def test_write_transaction_blocks_concurrent_writer(database, monkeypatch):
    data = values()
    preview = service.preview_batch(**data)
    original_preview = service._batch_preview
    concurrent = create_engine(str(database.kw["bind"].url), connect_args={"timeout": 0})

    def check_lock(session, batch):
        result = original_preview(session, batch)
        with concurrent.connect() as connection:
            with pytest.raises(OperationalError, match="locked"):
                connection.execute(db.Stamp.__table__.insert().values(date=date(2026, 10, 20), type="SICK"))
        return result

    monkeypatch.setattr(service, "_batch_preview", check_lock)
    try:
        result = service.save_batch(preview_token=preview["preview_token"], **data)
        assert result["created"] == 2
    finally:
        concurrent.dispose()


@pytest.mark.parametrize("change", ["target", "holiday"])
def test_changed_workday_rules_invalidate_preview(database, change):
    data = values()
    preview = service.preview_batch(**data)
    if change == "target":
        db.set_config("target_hours", "7.5")
    else:
        with database() as session:
            session.add(db.Holiday(date=date(2026, 10, 20), year=2026, name="Neuer Feiertag"))
            session.commit()
    with pytest.raises(service.BatchConflictError):
        service.save_batch(preview_token=preview["preview_token"], **data)
    with database() as session:
        assert session.query(db.Stamp).count() == 0


def test_legacy_absence_calls_remain_compatible(database):
    entries = service.add_absence("FLEX", date(2026, 10, 23), date(2026, 10, 27), "CLI")
    assert len(entries) == 2 and all(entry.overtime == -8 for entry in entries)
    assert service.add_absence("SICK", date(2026, 10, 23))[0].type == "SICK"


def test_holiday_failure_is_explicit_for_batch_only(database, monkeypatch):
    monkeypatch.setattr(holidays.requests, "get", Mock(side_effect=requests.ConnectionError("offline")))
    data = values(dates=[date(2028, 1, 3)])
    with pytest.raises(holidays.HolidayUnavailableError):
        service.preview_batch(**data)
    assert holidays.fetch_holidays(2028) == []
    with database() as session:
        assert session.query(db.Stamp).count() == 0


@pytest.mark.parametrize("response", [{}, [], {"Feiertag": {}}, {"Feiertag": {"datum": 12}},
                                      {"Feiertag": {"datum": "kaputt"}}, {"Feiertag": {"datum": "2027-01-01"}}])
def test_bad_holiday_response_is_not_used(database, monkeypatch, response):
    monkeypatch.setattr(holidays.requests, "get", Mock(return_value=Mock(json=lambda: response)))
    with pytest.raises(holidays.HolidayUnavailableError):
        holidays.fetch_holidays(2028, strict=True)


def test_api_preview_commit_conflict_and_validation(client):
    data = {"type": "WORK", "dates": ["2026-10-20", "2026-10-21"],
            "stamp_in": "08:00", "stamp_out": "17:00", "pause": 45, "note": "API"}
    response = client.post("/api/stamps/batch/preview", json=data)
    assert response.status_code == 200
    preview = response.json()
    assert preview["conflicts"] == [] and len(preview["dates"]) == 2
    response = client.post("/api/stamps/batch", json={**data, "preview_token": preview["preview_token"]})
    assert response.status_code == 200
    assert response.json()["entries"][0]["work_hours"] == 8.25
    response = client.post("/api/stamps/batch", json={**data, "preview_token": preview["preview_token"], "overwrite": True})
    assert response.status_code == 409
    assert response.json()["detail"]["code"] == "batch_conflict"
    assert client.post("/api/stamps/batch/preview", json={**data, "type": "INVALID"}).status_code == 422
    assert client.post("/api/stamps/batch/preview", json={**data, "stamp_in": "bad"}).status_code == 422
    assert client.post("/api/stamps/batch/preview", json={**data, "dates": []}).status_code == 400
    assert client.post("/api/stamps/batch", json=data).status_code == 422


def test_api_holiday_failure(client, monkeypatch):
    monkeypatch.setattr(holidays.requests, "get", Mock(side_effect=requests.ConnectionError()))
    response = client.post("/api/stamps/batch/preview", json={"type": "SICK", "dates": ["2028-01-03"]})
    assert response.status_code == 503
    assert "Feiertage" in response.json()["detail"]


def test_api_database_failure_is_explicit(client, monkeypatch):
    def fail(**kwargs):
        raise OperationalError("test", {}, Exception("database unavailable"))

    monkeypatch.setattr(service, "preview_batch", fail)
    response = client.post("/api/stamps/batch/preview", json={"type": "SICK", "dates": ["2026-10-20"]})
    assert response.status_code == 503
    assert "Datenbank" in response.json()["detail"]
