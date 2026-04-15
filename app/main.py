"""
Excel to PowerPoint Web Application — v6.0.0 (Refactored)

Slim entry point: creates the FastAPI app, mounts middleware and static files,
and includes the API router.  All business logic lives in services/.
"""
import asyncio
from contextlib import asynccontextmanager

from fastapi import FastAPI
from fastapi.responses import HTMLResponse
from fastapi.staticfiles import StaticFiles
from fastapi.middleware.cors import CORSMiddleware

from app.config import (
    logger,
    APP_TITLE,
    APP_DESCRIPTION,
    APP_VERSION,
    STATIC_DIR,
    FILE_CLEANUP_MAX_AGE,
)
from app.routers.api import router as api_router
from app.services.file_manager import file_manager


# ── Lifespan: startup / shutdown hooks ────────────────────────────────
@asynccontextmanager
async def lifespan(app: FastAPI):
    """Application lifespan events."""
    logger.info("Starting %s v%s", APP_TITLE, APP_VERSION)

    # Start periodic cleanup task
    cleanup_task = asyncio.create_task(_periodic_cleanup())

    yield

    # Shutdown
    cleanup_task.cancel()
    try:
        await cleanup_task
    except asyncio.CancelledError:
        pass
    logger.info("Shutting down %s", APP_TITLE)


async def _periodic_cleanup(interval: int = 3600):
    """Run file cleanup every *interval* seconds."""
    while True:
        await asyncio.sleep(interval)
        try:
            file_manager.cleanup_old_files(FILE_CLEANUP_MAX_AGE)
        except Exception as e:
            logger.warning("Cleanup task error: %s", e)


# ── Create app ────────────────────────────────────────────────────────
app = FastAPI(
    title=APP_TITLE,
    description=APP_DESCRIPTION,
    version=APP_VERSION,
    lifespan=lifespan,
)

# CORS middleware
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Mount static files
app.mount("/static", StaticFiles(directory=str(STATIC_DIR)), name="static")

# Include API router
app.include_router(api_router)


# ── Serve the frontend SPA ───────────────────────────────────────────
@app.get("/", response_class=HTMLResponse)
async def home():
    """Serve the main HTML page."""
    html_path = STATIC_DIR / "index.html"
    if html_path.exists():
        return HTMLResponse(
            content=html_path.read_text(encoding="utf-8"),
            headers={"Cache-Control": "no-cache, no-store, must-revalidate"},
        )
    return HTMLResponse(
        "<html><body><h1>Please add index.html</h1></body></html>",
        status_code=503,
    )


# ── Dev server ────────────────────────────────────────────────────────
if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="127.0.0.1", port=8000)
