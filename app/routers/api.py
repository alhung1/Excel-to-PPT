"""
API router — all REST endpoints for the Excel-to-PPT application.
"""
import uuid
import shutil
from pathlib import Path

from fastapi import APIRouter, UploadFile, File, HTTPException
from fastapi.responses import FileResponse

from app.config import (
    logger,
    UPLOAD_DIR,
    OUTPUT_DIR,
    APP_VERSION,
    ALLOWED_EXCEL_EXTENSIONS,
    ALLOWED_PPT_EXTENSIONS,
)
from app.models.schemas import GenerateRequest, HealthResponse
from app.services.excel_service import get_excel_info
from app.services.ppt_service import (
    get_ppt_info,
    get_ppt_slide_titles,
    process_image_mappings,
    process_embedded_mappings,
)
from app.services.file_manager import file_manager, get_directory_size_mb

from pptx import Presentation

router = APIRouter(prefix="/api")


# ============================================================
# Upload Excel
# ============================================================
@router.post("/upload-excel")
async def upload_excel(file: UploadFile = File(...)):
    """Upload an Excel file and return sheet/chart info."""
    if not file.filename.endswith(ALLOWED_EXCEL_EXTENSIONS):
        raise HTTPException(400, "必須是 Excel 檔案 (.xlsx, .xlsm, .xls)")

    file_id = uuid.uuid4().hex[:8]
    file_path = UPLOAD_DIR / f"{file_id}_{file.filename}"

    with open(file_path, "wb") as f:
        content = await file.read()
        f.write(content)

    try:
        info = get_excel_info(str(file_path))
        file_manager.register(file_id, "excel", str(file_path), file.filename)

        return {
            "status": "success",
            "file_id": file_id,
            "filename": file.filename,
            "worksheets": info["worksheets"],
            "chartsheets": info["chartsheets"],
        }
    except Exception as e:
        if file_path.exists():
            file_path.unlink()
        logger.error("Failed to read Excel: %s", e, exc_info=True)
        raise HTTPException(500, f"讀取 Excel 失敗: {e}")


# ============================================================
# Upload PPT
# ============================================================
@router.post("/upload-ppt")
async def upload_ppt(file: UploadFile = File(...)):
    """Upload a PPT template and return slide info."""
    if not file.filename.endswith(ALLOWED_PPT_EXTENSIONS):
        raise HTTPException(400, "必須是 PowerPoint 檔案 (.pptx, .ppt)")

    file_id = uuid.uuid4().hex[:8]
    file_path = UPLOAD_DIR / f"{file_id}_{file.filename}"

    with open(file_path, "wb") as f:
        content = await file.read()
        f.write(content)

    try:
        info = get_ppt_info(str(file_path))
        file_manager.register(file_id, "ppt", str(file_path), file.filename)

        return {"status": "success", "file_id": file_id, "filename": file.filename, **info}
    except Exception as e:
        if file_path.exists():
            file_path.unlink()
        logger.error("Failed to read PPT: %s", e, exc_info=True)
        raise HTTPException(500, f"讀取 PPT 失敗: {e}")


# ============================================================
# Remove file
# ============================================================
@router.delete("/remove-file/{file_id}")
async def remove_file(file_id: str):
    """Remove an uploaded file."""
    if file_manager.remove(file_id):
        return {"status": "success"}
    raise HTTPException(404, "檔案不存在")


# ============================================================
# Generate PPT  (sync — FastAPI runs it in a thread pool)
# ============================================================
@router.post("/generate")
def generate_ppt(request: GenerateRequest):
    """Generate a PowerPoint with chart mappings.

    NOTE: This is intentionally a **sync** function (``def``, not ``async def``)
    so that FastAPI automatically runs it in a thread pool, preventing the
    long-running COM operations from blocking the event loop.
    """
    template_info = file_manager.get(request.template_id)
    if not template_info:
        raise HTTPException(404, "PPT 模板不存在，請重新上傳")
    template_path = template_info["path"]

    # Build a lookup for uploaded files needed by mappings
    uploaded_files: dict = {}
    for m in request.mappings:
        info = file_manager.get(m.excel_id)
        if not info:
            raise HTTPException(404, f"Excel 檔案不存在: {m.excel_id}")
        uploaded_files[m.excel_id] = info

    job_id = uuid.uuid4().hex[:8]
    job_dir = OUTPUT_DIR / job_id
    job_dir.mkdir(exist_ok=True)

    try:
        slide_titles = get_ppt_slide_titles(template_path)

        image_mappings = [m for m in request.mappings if m.chart_mode == "image"]
        embedded_mappings = [m for m in request.mappings if m.chart_mode == "embedded"]

        logger.info(
            "[Generate] Image mappings: %d, Embedded mappings: %d",
            len(image_mappings),
            len(embedded_mappings),
        )

        output_filename = (
            f"{request.output_name}.pptx"
            if not request.output_name.endswith(".pptx")
            else request.output_name
        )
        output_path = job_dir / output_filename

        all_results = []

        # Step 1: image mode (python-pptx)
        if image_mappings:
            logger.info("[Generate] Processing image mode mappings...")
            prs = Presentation(template_path)
            image_results = process_image_mappings(
                image_mappings, prs, request, job_dir, slide_titles, uploaded_files
            )
            all_results.extend(image_results)
            prs.save(str(output_path))
        else:
            shutil.copy(template_path, str(output_path))

        # Step 2: embedded mode (COM)
        if embedded_mappings:
            logger.info("[Generate] Processing embedded mode mappings...")
            embedded_results = process_embedded_mappings(
                embedded_mappings,
                str(output_path.resolve()),
                request,
                slide_titles,
                uploaded_files,
            )
            all_results.extend(embedded_results)

        # Determine mode string
        if image_mappings and embedded_mappings:
            mode_str = "mixed"
        elif embedded_mappings:
            mode_str = "embedded"
        else:
            mode_str = "image"

        return {
            "status": "success",
            "job_id": job_id,
            "download_url": f"/api/download/{job_id}/{output_filename}",
            "results": all_results,
            "output_file": str(output_path),
            "mode": mode_str,
        }

    except Exception as e:
        logger.error("Generate PPT failed: %s", e, exc_info=True)
        raise HTTPException(500, f"產生 PPT 失敗: {e}")


# ============================================================
# Download
# ============================================================
@router.get("/download/{job_id}/{filename}")
async def download_file(job_id: str, filename: str):
    """Download a generated file."""
    file_path = OUTPUT_DIR / job_id / filename
    if not file_path.exists():
        raise HTTPException(404, "檔案不存在或已過期")
    return FileResponse(
        file_path,
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        filename=filename,
    )


# ============================================================
# Health check
# ============================================================
@router.get("/health", response_model=HealthResponse)
async def health():
    """Return service health and status."""
    return HealthResponse(
        status="ok",
        version=APP_VERSION,
        uploads_count=file_manager.count,
        outputs_dir_size_mb=round(get_directory_size_mb(OUTPUT_DIR), 2),
    )
