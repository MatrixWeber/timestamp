"""FastAPI Web-Application."""

from pathlib import Path

from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles

from stamp.db import init_db
from stamp.api.routes import router

app = FastAPI(title="stamp", description="⏱️ Zeiterfassung", version="0.1.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

app.include_router(router)

# Statische Dateien für das Web-Frontend (falls vorhanden)
STATIC_DIR = Path(__file__).parent.parent.parent / "web" / "dist"
if STATIC_DIR.exists():
    app.mount("/", StaticFiles(directory=str(STATIC_DIR), html=True), name="static")


@app.on_event("startup")
def startup():
    init_db()
