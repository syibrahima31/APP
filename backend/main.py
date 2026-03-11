"""
Backend FastAPI — Dashboard ISI v3
Lancement : uvicorn backend.main:app --reload --port 8000
"""
from __future__ import annotations

import time
from contextlib import asynccontextmanager

from fastapi import FastAPI, Request
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse

from backend.core.config import settings
from db.database import init_db
from backend.routers import classes, enseignements, pointage
from backend.routers import auth as auth_router
from backend.routers import analytics, alertes as alertes_router, excel_import


@asynccontextmanager
async def lifespan(app: FastAPI):
    init_db()
    yield


app = FastAPI(
    title=settings.APP_NAME,
    description=(
        "API REST pour le suivi mensuel des enseignements, le systeme de pointage, "
        "les alertes intelligentes et les analytics avances."
    ),
    version=settings.APP_VERSION,
    lifespan=lifespan,
    docs_url="/api/docs",
    redoc_url="/api/redoc",
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=settings.ALLOWED_ORIGINS,
    allow_credentials=True,
    allow_methods=["GET", "POST", "PUT", "DELETE", "OPTIONS"],
    allow_headers=["Authorization", "Content-Type", "Accept"],
)


@app.middleware("http")
async def add_process_time(request: Request, call_next):
    start = time.perf_counter()
    response = await call_next(request)
    response.headers["X-Process-Time"] = f"{(time.perf_counter()-start)*1000:.1f}ms"
    return response


@app.exception_handler(Exception)
async def global_exception_handler(request: Request, exc: Exception):
    return JSONResponse(
        status_code=500,
        content={"detail": "Erreur interne du serveur", "type": type(exc).__name__},
    )


app.include_router(auth_router.router)
app.include_router(classes.router,           prefix="/api/classes",       tags=["Classes"])
app.include_router(enseignements.router,     prefix="/api/enseignements", tags=["Enseignements"])
app.include_router(pointage.router,          prefix="/api/pointages",     tags=["Pointage"])
app.include_router(analytics.router,         prefix="/api/analytics",     tags=["Analytics"])
app.include_router(alertes_router.router,    prefix="/api/alertes",       tags=["Alertes"])
app.include_router(excel_import.router,      prefix="/api/import",        tags=["Import Excel"])


@app.get("/api/health", tags=["Sante"])
def health():
    return {
        "status": "ok",
        "version": settings.APP_VERSION,
        "environment": "production" if settings.is_prod() else "development",
    }
