"""Pydantic-Schemas für die API."""

from datetime import date, time
from typing import Optional
from pydantic import BaseModel


class StampResponse(BaseModel):
    id: int
    date: date
    stamp_in: Optional[time] = None
    stamp_out: Optional[time] = None
    pause: float
    work_hours: Optional[float] = None
    overtime: Optional[float] = None
    type: str
    note: Optional[str] = None

    class Config:
        from_attributes = True


class StampInRequest(BaseModel):
    time: Optional[str] = None  # HH:MM


class StampOutRequest(BaseModel):
    time: Optional[str] = None  # HH:MM


class AbsenceRequest(BaseModel):
    type: str  # VACATION, SICK, FLEX, TRAVEL
    start_date: date
    end_date: Optional[date] = None
    note: Optional[str] = None


class EditRequest(BaseModel):
    stamp_in: Optional[str] = None  # HH:MM
    stamp_out: Optional[str] = None  # HH:MM
    pause: Optional[float] = None  # Minuten
    type: Optional[str] = None
    note: Optional[str] = None


class OvertimeResponse(BaseModel):
    today: float
    week: float
    month: float
    year: float
    carryover: float
    total: float


class VacationResponse(BaseModel):
    total: int
    carryover: int
    taken: int
    planned: int
    remaining: int
    flex_taken: int
    flex_planned: int
    sick_days: int
    travel_taken: int
    travel_planned: int


class ConfigEntry(BaseModel):
    key: str
    value: str


class DayInfo(BaseModel):
    date: date
    weekday: str
    is_holiday: bool
    holiday_name: Optional[str] = None
    is_weekend: bool
    stamp: Optional[StampResponse] = None
