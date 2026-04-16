"""SQLAlchemy-Datenbank-Models und Setup."""

from datetime import date, time, datetime
from sqlalchemy import create_engine, Column, Integer, String, Float, Date, Time, DateTime, UniqueConstraint
from sqlalchemy.orm import declarative_base, sessionmaker

from stamp.config import DB_URL, DEFAULTS

Base = declarative_base()


class Stamp(Base):
    """Ein Tageseintrag (Stempelung oder Abwesenheit)."""
    __tablename__ = "stamps"

    id = Column(Integer, primary_key=True, autoincrement=True)
    date = Column(Date, nullable=False, unique=True, index=True)
    stamp_in = Column(Time, nullable=True)
    stamp_out = Column(Time, nullable=True)
    pause = Column(Float, default=0.75)
    work_hours = Column(Float, nullable=True)
    overtime = Column(Float, nullable=True)
    type = Column(String, default="WORK")  # WORK, VACATION, SICK, FLEX, TRAVEL
    note = Column(String, nullable=True)
    created_at = Column(DateTime, default=datetime.now)
    updated_at = Column(DateTime, default=datetime.now, onupdate=datetime.now)

    def __repr__(self):
        return f"<Stamp {self.date} {self.type} in={self.stamp_in} out={self.stamp_out}>"


class Holiday(Base):
    """Feiertag (gecacht aus der API)."""
    __tablename__ = "holidays"

    id = Column(Integer, primary_key=True, autoincrement=True)
    date = Column(Date, nullable=False, unique=True)
    name = Column(String, nullable=False)
    year = Column(Integer, nullable=False)

    def __repr__(self):
        return f"<Holiday {self.date} {self.name}>"


class Config(Base):
    """Key-Value-Konfiguration."""
    __tablename__ = "config"

    key = Column(String, primary_key=True)
    value = Column(String, nullable=False)


# Engine und Session
engine = create_engine(DB_URL, echo=False)
SessionLocal = sessionmaker(bind=engine)


def init_db():
    """Erstellt alle Tabellen und setzt Default-Config."""
    Base.metadata.create_all(engine)
    with SessionLocal() as session:
        for key, value in DEFAULTS.items():
            existing = session.query(Config).filter_by(key=key).first()
            if not existing:
                session.add(Config(key=key, value=value))
        session.commit()


def get_session():
    """Gibt eine neue DB-Session zurück."""
    return SessionLocal()


def get_config(key: str, default: str | None = None) -> str | None:
    """Liest einen Config-Wert aus der DB."""
    with get_session() as session:
        entry = session.query(Config).filter_by(key=key).first()
        return entry.value if entry else (default or DEFAULTS.get(key))


def set_config(key: str, value: str):
    """Setzt einen Config-Wert in der DB."""
    with get_session() as session:
        entry = session.query(Config).filter_by(key=key).first()
        if entry:
            entry.value = value
        else:
            session.add(Config(key=key, value=value))
        session.commit()


def get_config_float(key: str) -> float:
    """Liest einen Float-Config-Wert."""
    return float(get_config(key, DEFAULTS.get(key, "0")))


def get_config_int(key: str) -> int:
    """Liest einen Int-Config-Wert."""
    return int(get_config(key, DEFAULTS.get(key, "0")))
