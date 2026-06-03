"""Точка входа FastAPI-приложения «Образовательный центр»."""
from __future__ import annotations

from pathlib import Path

from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles

from .database import init_schema, seed_data
from .routers import courses_router, feedback_router, teachers_router

BASE_DIR = Path(__file__).resolve().parent
FRONTEND_DIR = BASE_DIR.parent / "frontend"

app = FastAPI(
    title="Образовательный центр — REST API",
    description="ВКР. Сайт образовательного центра с формой обратной связи.",
    version="1.0.0",
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)


@app.on_event("startup")
def on_startup() -> None:
    init_schema()
    seed_data()


app.include_router(courses_router.router)
app.include_router(teachers_router.router)
app.include_router(feedback_router.router)

app.mount("/static", StaticFiles(directory=FRONTEND_DIR), name="static")


@app.get("/")
def root_index() -> FileResponse:
    return FileResponse(FRONTEND_DIR / "index.html")
