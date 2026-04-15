"""
Pydantic data models for API request/response schemas.
"""
from typing import List, Optional, Dict
from pydantic import BaseModel, Field


class ChartMapping(BaseModel):
    """A single mapping from an Excel item to a PPT slide."""
    excel_id: str = Field(..., description="Uploaded Excel file ID")
    name: str = Field(..., description="Sheet/Chart name in Excel")
    page: int = Field(..., ge=1, description="Target PPT page number (1-based)")
    type: str = Field(..., description="'worksheet' or 'chartsheet'")
    chart_mode: str = Field(
        default="image",
        description="'image' for static PNG or 'embedded' for editable chart",
    )


class GenerateRequest(BaseModel):
    """Request body for the /api/generate endpoint."""
    template_id: str
    output_name: str
    mappings: List[ChartMapping]
    chart_mode: str = "image"  # global default (overridden by per-mapping)
    img_left: float = 0.423
    img_top: float = 1.4
    img_width: float = 12.0
    img_height: float = 5.6


class FileInfo(BaseModel):
    """Metadata for an uploaded file stored in memory."""
    type: str  # "excel" or "ppt"
    path: str
    filename: str


class UploadExcelResponse(BaseModel):
    status: str = "success"
    file_id: str
    filename: str
    worksheets: List[Dict]
    chartsheets: List[Dict]


class UploadPptResponse(BaseModel):
    status: str = "success"
    file_id: str
    filename: str
    total_slides: int
    slides: List[Dict]
    width: float
    height: float


class GenerateResult(BaseModel):
    name: str
    excel: str
    status: str
    page: Optional[int] = None
    mode: Optional[str] = None
    reason: Optional[str] = None
    mesh_layout: Optional[bool] = None


class GenerateResponse(BaseModel):
    status: str = "success"
    job_id: str
    download_url: str
    results: List[Dict]
    output_file: str
    mode: str


class HealthResponse(BaseModel):
    status: str = "ok"
    version: str
    uploads_count: int
    outputs_dir_size_mb: float
