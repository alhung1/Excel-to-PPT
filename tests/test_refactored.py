"""
Comprehensive test suite for the refactored Excel-to-PPT application.

Tests cover:
1. Module import integrity
2. Config values
3. Pydantic model validation
4. Image validator (with real test images)
5. File manager operations
6. PPT service helpers
7. FastAPI app routes (TestClient)
"""
import os
import sys
import time
import tempfile
import shutil

# Ensure project root is on the path
PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, PROJECT_ROOT)

PASS = 0
FAIL = 0
ERRORS = []


def test(name):
    """Decorator that runs a test function and tracks pass/fail."""
    def decorator(fn):
        global PASS, FAIL
        try:
            fn()
            PASS += 1
            print(f"  [PASS] {name}")
        except Exception as e:
            FAIL += 1
            ERRORS.append((name, str(e)))
            print(f"  [FAIL] {name}: {e}")
        return fn
    return decorator


# =====================================================================
# 1. Module imports
# =====================================================================
print("\n=== 1. Module Import Tests ===")

@test("import app.config")
def _():
    from app.config import logger, BASE_DIR, APP_VERSION
    assert APP_VERSION == "6.0.0"

@test("import app.models.schemas")
def _():
    from app.models.schemas import ChartMapping, GenerateRequest, HealthResponse
    assert ChartMapping is not None

@test("import app.utils.image_validator")
def _():
    from app.utils.image_validator import validate_image
    assert callable(validate_image)

@test("import app.utils.clipboard")
def _():
    from app.utils.clipboard import clear_clipboard
    assert callable(clear_clipboard)

@test("import app.services.file_manager")
def _():
    from app.services.file_manager import file_manager, FileManager
    assert isinstance(file_manager, FileManager)

@test("import app.services.ppt_service")
def _():
    from app.services.ppt_service import get_ppt_info, get_effective_layout, is_mesh_slide_title
    assert callable(get_ppt_info)

@test("import app.services.excel_service")
def _():
    from app.services.excel_service import get_excel_info, capture_item, ExcelCOM
    assert callable(get_excel_info)

@test("import app.routers.api")
def _():
    from app.routers.api import router
    assert router is not None

@test("import app.main (FastAPI app)")
def _():
    from app.main import app
    assert app is not None
    assert app.title == "Excel to PowerPoint Generator"

@test("import cli.report_cli")
def _():
    from cli.report_cli import parse_args, load_config
    assert callable(parse_args)


# =====================================================================
# 2. Config values
# =====================================================================
print("\n=== 2. Config Value Tests ===")

@test("config paths exist")
def _():
    from app.config import BASE_DIR, UPLOAD_DIR, OUTPUT_DIR, STATIC_DIR
    assert BASE_DIR.exists()
    # These are created by config import
    assert UPLOAD_DIR.exists()
    assert OUTPUT_DIR.exists()

@test("config constants are reasonable")
def _():
    from app.config import (
        COM_MAX_RETRIES, COM_RETRY_DELAY, IMAGE_MIN_SIZE_BYTES,
        MAX_UPLOAD_SIZE_MB, FILE_CLEANUP_MAX_AGE,
    )
    assert COM_MAX_RETRIES >= 1
    assert COM_RETRY_DELAY > 0
    assert IMAGE_MIN_SIZE_BYTES > 0
    assert MAX_UPLOAD_SIZE_MB > 0
    assert FILE_CLEANUP_MAX_AGE > 0

@test("mesh layout presets have required keys")
def _():
    from app.config import MESH_BACKHAUL_LAYOUT, MESH_FRONTHAUL_LAYOUT
    for layout in [MESH_BACKHAUL_LAYOUT, MESH_FRONTHAUL_LAYOUT]:
        assert "left" in layout
        assert "top" in layout
        assert "width" in layout
        assert "height" in layout


# =====================================================================
# 3. Pydantic model validation
# =====================================================================
print("\n=== 3. Pydantic Model Tests ===")

@test("ChartMapping valid data")
def _():
    from app.models.schemas import ChartMapping
    m = ChartMapping(excel_id="abc123", name="Sheet1", page=1, type="worksheet")
    assert m.chart_mode == "image"  # default

@test("ChartMapping embedded mode")
def _():
    from app.models.schemas import ChartMapping
    m = ChartMapping(excel_id="x", name="C1", page=2, type="chartsheet", chart_mode="embedded")
    assert m.chart_mode == "embedded"

@test("GenerateRequest defaults")
def _():
    from app.models.schemas import GenerateRequest, ChartMapping
    r = GenerateRequest(
        template_id="tpl1",
        output_name="Report",
        mappings=[ChartMapping(excel_id="e1", name="S1", page=1, type="worksheet")],
    )
    assert r.img_left == 0.423
    assert r.img_width == 12.0

@test("HealthResponse")
def _():
    from app.models.schemas import HealthResponse
    h = HealthResponse(status="ok", version="6.0.0", uploads_count=3, outputs_dir_size_mb=12.5)
    assert h.version == "6.0.0"


# =====================================================================
# 4. Image validator
# =====================================================================
print("\n=== 4. Image Validator Tests ===")

@test("validate_image: non-existent file returns False")
def _():
    from app.utils.image_validator import validate_image
    assert validate_image("/tmp/does_not_exist_12345.png") is False

@test("validate_image: empty file returns False")
def _():
    from app.utils.image_validator import validate_image
    tmp = tempfile.NamedTemporaryFile(suffix=".png", delete=False)
    tmp.write(b"x" * 100)
    tmp.close()
    assert validate_image(tmp.name, min_size=500) is False
    os.unlink(tmp.name)

@test("validate_image: valid color image returns True")
def _():
    from app.utils.image_validator import validate_image
    from PIL import Image
    import random
    # Create a colorful test image
    img = Image.new("RGB", (200, 200))
    pixels = img.load()
    for x in range(200):
        for y in range(200):
            pixels[x, y] = (x % 256, y % 256, (x + y) % 256)
    tmp = tempfile.NamedTemporaryFile(suffix=".png", delete=False)
    img.save(tmp.name)
    img.close()
    assert validate_image(tmp.name) is True
    os.unlink(tmp.name)

@test("validate_image: solid color image returns False")
def _():
    from app.utils.image_validator import validate_image
    from PIL import Image
    img = Image.new("RGB", (200, 200), color=(255, 255, 255))
    tmp = tempfile.NamedTemporaryFile(suffix=".png", delete=False)
    img.save(tmp.name)
    img.close()
    result = validate_image(tmp.name)
    os.unlink(tmp.name)
    assert result is False


# =====================================================================
# 5. File manager
# =====================================================================
print("\n=== 5. File Manager Tests ===")

@test("FileManager register and get")
def _():
    from app.services.file_manager import FileManager
    fm = FileManager()
    fm.register("f1", "excel", "/tmp/test.xlsx", "test.xlsx")
    info = fm.get("f1")
    assert info is not None
    assert info["type"] == "excel"
    assert info["filename"] == "test.xlsx"
    assert "created_at" in info

@test("FileManager remove")
def _():
    from app.services.file_manager import FileManager
    fm = FileManager()
    # Create a real temp file
    tmp = tempfile.NamedTemporaryFile(delete=False)
    tmp.close()
    fm.register("f2", "ppt", tmp.name, "temp.pptx")
    assert fm.remove("f2") is True
    assert not os.path.exists(tmp.name)
    assert fm.get("f2") is None

@test("FileManager remove non-existent returns False")
def _():
    from app.services.file_manager import FileManager
    fm = FileManager()
    assert fm.remove("no_such_id") is False

@test("FileManager count")
def _():
    from app.services.file_manager import FileManager
    fm = FileManager()
    assert fm.count == 0
    fm.register("a", "excel", "/tmp/a", "a.xlsx")
    fm.register("b", "ppt", "/tmp/b", "b.pptx")
    assert fm.count == 2

@test("FileManager cleanup_old_files")
def _():
    from app.services.file_manager import FileManager
    fm = FileManager()
    fm.register("old", "excel", "/tmp/nonexistent", "old.xlsx")
    # Manually set old timestamp
    fm._files["old"]["created_at"] = time.time() - 100000
    fm.cleanup_old_files(max_age=1)
    assert fm.get("old") is None


# =====================================================================
# 6. PPT service helpers
# =====================================================================
print("\n=== 6. PPT Service Helper Tests ===")

@test("get_mesh_layout_type: backhaul")
def _():
    from app.services.ppt_service import get_mesh_layout_type
    assert get_mesh_layout_type("Mesh Backhaul Test") == "backhaul"

@test("get_mesh_layout_type: fronthaul")
def _():
    from app.services.ppt_service import get_mesh_layout_type
    assert get_mesh_layout_type("MESH Fronthaul 5GHz") == "fronthaul"

@test("get_mesh_layout_type: non-mesh")
def _():
    from app.services.ppt_service import get_mesh_layout_type
    assert get_mesh_layout_type("Summary") is None

@test("is_mesh_slide_title")
def _():
    from app.services.ppt_service import is_mesh_slide_title
    assert is_mesh_slide_title("Mesh Backhaul") is True
    assert is_mesh_slide_title("Regular Slide") is False

@test("get_effective_layout: uses mesh layout for backhaul")
def _():
    from app.services.ppt_service import get_effective_layout
    from app.models.schemas import GenerateRequest, ChartMapping
    from app.config import MESH_BACKHAUL_LAYOUT

    req = GenerateRequest(
        template_id="t", output_name="o",
        mappings=[ChartMapping(excel_id="e", name="s", page=1, type="worksheet")],
    )
    layout = get_effective_layout(req, "Mesh Backhaul DL")
    assert layout["left"] == MESH_BACKHAUL_LAYOUT["left"]
    assert layout["height"] == MESH_BACKHAUL_LAYOUT["height"]

@test("get_effective_layout: uses request values for non-mesh")
def _():
    from app.services.ppt_service import get_effective_layout
    from app.models.schemas import GenerateRequest, ChartMapping
    req = GenerateRequest(
        template_id="t", output_name="o",
        mappings=[ChartMapping(excel_id="e", name="s", page=1, type="worksheet")],
        img_left=1.0, img_top=2.0, img_width=10.0, img_height=6.0,
    )
    layout = get_effective_layout(req, "Summary")
    assert layout["left"] == 1.0
    assert layout["top"] == 2.0

@test("get_ppt_info on template file")
def _():
    from app.services.ppt_service import get_ppt_info
    template = os.path.join(PROJECT_ROOT, "templates", "Mesh.pptx")
    if os.path.exists(template):
        info = get_ppt_info(template)
        assert info["total_slides"] > 0
        assert "slides" in info
        assert "width" in info
    else:
        # Skip if template not available
        pass


# =====================================================================
# 7. FastAPI TestClient
# =====================================================================
print("\n=== 7. FastAPI TestClient Tests ===")

@test("TestClient: GET / returns HTML")
def _():
    from fastapi.testclient import TestClient
    from app.main import app
    client = TestClient(app)
    resp = client.get("/")
    assert resp.status_code == 200
    assert "html" in resp.headers.get("content-type", "").lower()

@test("TestClient: GET /api/health returns status")
def _():
    from fastapi.testclient import TestClient
    from app.main import app
    client = TestClient(app)
    resp = client.get("/api/health")
    assert resp.status_code == 200
    data = resp.json()
    assert data["status"] == "ok"
    assert data["version"] == "6.0.0"

@test("TestClient: POST /api/upload-excel rejects non-Excel")
def _():
    from fastapi.testclient import TestClient
    from app.main import app
    client = TestClient(app)
    resp = client.post("/api/upload-excel", files={"file": ("test.txt", b"hello", "text/plain")})
    assert resp.status_code == 400

@test("TestClient: POST /api/upload-ppt rejects non-PPT")
def _():
    from fastapi.testclient import TestClient
    from app.main import app
    client = TestClient(app)
    resp = client.post("/api/upload-ppt", files={"file": ("test.txt", b"hello", "text/plain")})
    assert resp.status_code == 400

@test("TestClient: DELETE /api/remove-file/unknown returns 404")
def _():
    from fastapi.testclient import TestClient
    from app.main import app
    client = TestClient(app)
    resp = client.delete("/api/remove-file/nonexistent")
    assert resp.status_code == 404

@test("TestClient: GET /api/download/bad/bad returns 404")
def _():
    from fastapi.testclient import TestClient
    from app.main import app
    client = TestClient(app)
    resp = client.get("/api/download/bad_job/bad_file.pptx")
    assert resp.status_code == 404

@test("TestClient: POST /api/generate with bad template_id returns 404")
def _():
    from fastapi.testclient import TestClient
    from app.main import app
    client = TestClient(app)
    resp = client.post("/api/generate", json={
        "template_id": "nonexistent",
        "output_name": "test",
        "mappings": [],
    })
    assert resp.status_code == 404


# =====================================================================
# Summary
# =====================================================================
print("\n" + "=" * 60)
total = PASS + FAIL
print(f"  Results: {PASS}/{total} passed, {FAIL} failed")
if ERRORS:
    print("\n  Failed tests:")
    for name, err in ERRORS:
        print(f"    - {name}: {err}")
print("=" * 60)

sys.exit(0 if FAIL == 0 else 1)
