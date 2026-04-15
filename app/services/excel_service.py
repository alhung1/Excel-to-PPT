"""
Excel COM automation service.

Provides context-managed access to Excel via Windows COM,
chart/worksheet capture, and info extraction.
"""
import os
import time
from typing import Dict, List, Optional

from app.config import logger, COM_MAX_RETRIES, COM_RETRY_DELAY, COM_CLIPBOARD_DELAY
from app.utils.image_validator import validate_image
from app.utils.clipboard import clear_clipboard


# ---------------------------------------------------------------------------
# Lazy imports for COM — only available on Windows
# ---------------------------------------------------------------------------
def _init_com():
    """Import and initialise COM libraries (Windows only)."""
    import pythoncom
    import win32com.client as win32
    return pythoncom, win32


# ---------------------------------------------------------------------------
# Context manager for Excel COM
# ---------------------------------------------------------------------------
class ExcelCOM:
    """Context manager that guarantees Excel COM cleanup.

    Usage::

        with ExcelCOM() as (excel_app, pythoncom_mod):
            wb = excel_app.Workbooks.Open(path)
            ...
    """

    def __init__(self, visible: bool = False):
        self._visible = visible
        self._excel_app = None
        self._pythoncom = None

    def __enter__(self):
        self._pythoncom, win32 = _init_com()
        self._pythoncom.CoInitialize()
        self._excel_app = win32.DispatchEx("Excel.Application")
        self._excel_app.Visible = self._visible
        self._excel_app.DisplayAlerts = False
        logger.info("Excel COM started (visible=%s)", self._visible)
        return self._excel_app, self._pythoncom

    def __exit__(self, exc_type, exc_val, exc_tb):
        if self._excel_app:
            try:
                self._excel_app.Quit()
                logger.info("Excel COM quit successfully")
            except Exception as e:
                logger.warning("Failed to quit Excel COM: %s", e)
        if self._pythoncom:
            try:
                self._pythoncom.CoUninitialize()
            except Exception:
                pass
        return False  # do not suppress exceptions


# ---------------------------------------------------------------------------
# Context manager for PowerPoint COM
# ---------------------------------------------------------------------------
class PowerPointCOM:
    """Context manager that guarantees PowerPoint COM cleanup."""

    def __init__(self, visible: bool = True):
        self._visible = visible
        self._ppt_app = None
        self._pythoncom = None

    def __enter__(self):
        self._pythoncom, win32 = _init_com()
        self._pythoncom.CoInitialize()
        self._ppt_app = win32.DispatchEx("PowerPoint.Application")
        self._ppt_app.Visible = self._visible
        logger.info("PowerPoint COM started (visible=%s)", self._visible)
        return self._ppt_app, self._pythoncom

    def __exit__(self, exc_type, exc_val, exc_tb):
        if self._ppt_app:
            try:
                self._ppt_app.Quit()
                logger.info("PowerPoint COM quit successfully")
            except Exception as e:
                logger.warning("Failed to quit PowerPoint COM: %s", e)
        if self._pythoncom:
            try:
                self._pythoncom.CoUninitialize()
            except Exception:
                pass
        return False


# ---------------------------------------------------------------------------
# Excel info extraction
# ---------------------------------------------------------------------------
def get_excel_info(excel_path: str) -> dict:
    """Return worksheet and chart sheet metadata from an Excel file."""
    with ExcelCOM() as (excel_app, _pythoncom):
        workbook = excel_app.Workbooks.Open(excel_path)

        worksheets = []
        for sheet in workbook.Worksheets:
            chart_count = 0
            try:
                chart_count = sheet.ChartObjects().Count
            except Exception:
                pass
            worksheets.append({
                "name": sheet.Name,
                "type": "worksheet",
                "has_charts": chart_count > 0,
                "chart_count": chart_count,
            })

        chartsheets = []
        try:
            for chart in workbook.Charts:
                chartsheets.append({"name": chart.Name, "type": "chartsheet"})
        except Exception:
            pass

        workbook.Close(SaveChanges=False)

    return {"worksheets": worksheets, "chartsheets": chartsheets}


# ---------------------------------------------------------------------------
# Chart / worksheet capture
# ---------------------------------------------------------------------------
def _export_via_copypicture(
    workbook, excel_app, source_obj, output_path: str
) -> bool:
    """Fallback: export an object using CopyPicture + temp chart sheet."""
    try:
        clear_clipboard()
        time.sleep(COM_CLIPBOARD_DELAY / 2)

        source_obj.CopyPicture(Appearance=1, Format=2)  # xlScreen, xlBitmap
        time.sleep(COM_CLIPBOARD_DELAY)

        temp_chart_sheet = workbook.Charts.Add()
        time.sleep(0.1)
        try:
            temp_chart_sheet.Paste()
            time.sleep(COM_CLIPBOARD_DELAY)
            temp_chart_sheet.Export(output_path, "PNG")
        finally:
            excel_app.DisplayAlerts = False
            temp_chart_sheet.Delete()

        return os.path.exists(output_path)
    except Exception as e:
        logger.warning("CopyPicture fallback failed: %s", e)
        return False


def capture_item(
    excel_app,
    workbook,
    name: str,
    item_type: str,
    output_path: str,
    max_retries: int = None,
) -> bool:
    """Capture a worksheet or chart sheet as a PNG image.

    Args:
        excel_app: Active Excel COM application instance.
        workbook: Open workbook COM object.
        name: Sheet or chart name.
        item_type: ``'chartsheet'`` or ``'worksheet'``.
        output_path: Destination PNG file path.
        max_retries: Override for retry count.

    Returns:
        ``True`` if capture succeeded, ``False`` otherwise.
    """
    if max_retries is None:
        max_retries = COM_MAX_RETRIES

    try:
        if item_type == "chartsheet":
            return _capture_chartsheet(excel_app, workbook, name, output_path, max_retries)
        else:
            return _capture_worksheet(excel_app, workbook, name, output_path, max_retries)
    except Exception as e:
        logger.error("Capturing '%s' failed: %s", name, e, exc_info=True)
        return False


def _capture_chartsheet(
    excel_app, workbook, name: str, output_path: str, max_retries: int
) -> bool:
    """Handle chart-sheet capture with fallback."""
    logger.info("  [ChartSheet] Exporting '%s' directly...", name)
    chart_sheet = workbook.Charts(name)
    chart_sheet.Export(output_path, "PNG")

    if validate_image(output_path):
        return True

    logger.info("  [ChartSheet] Direct export invalid, trying CopyPicture fallback...")
    if os.path.exists(output_path):
        os.remove(output_path)

    chart_sheet.ChartArea.CopyPicture(Appearance=1, Format=2)
    time.sleep(COM_CLIPBOARD_DELAY)
    temp_chart_sheet = workbook.Charts.Add()
    try:
        temp_chart_sheet.Paste()
        time.sleep(COM_CLIPBOARD_DELAY)
        temp_chart_sheet.Export(output_path, "PNG")
    finally:
        excel_app.DisplayAlerts = False
        temp_chart_sheet.Delete()

    if validate_image(output_path):
        return True

    logger.warning("  [ChartSheet] All methods failed for '%s'", name)
    return False


def _capture_worksheet(
    excel_app, workbook, name: str, output_path: str, max_retries: int
) -> bool:
    """Handle worksheet capture (with or without embedded charts)."""
    sheet = workbook.Worksheets(name)
    chart_count = 0
    try:
        chart_count = sheet.ChartObjects().Count
    except Exception:
        pass

    logger.info("  [Worksheet] '%s' has %d embedded chart(s)", name, chart_count)

    if chart_count > 0:
        return _capture_worksheet_chart(
            excel_app, workbook, sheet, name, output_path, max_retries
        )
    else:
        return _capture_worksheet_range(
            excel_app, workbook, sheet, name, output_path, max_retries
        )


def _capture_worksheet_chart(
    excel_app, workbook, sheet, name: str, output_path: str, max_retries: int
) -> bool:
    """Capture the first embedded chart from a worksheet."""
    chart_obj = sheet.ChartObjects(1)

    # Try 1: direct export
    logger.info("  [Worksheet] Trying direct Chart.Export()...")
    chart_obj.Chart.Export(output_path, "PNG")
    if validate_image(output_path):
        return True

    logger.info("  [Worksheet] Direct export invalid, trying CopyPicture on ChartObject...")
    if os.path.exists(output_path):
        os.remove(output_path)

    # Try 2: CopyPicture on the chart object
    for attempt in range(max_retries):
        try:
            clear_clipboard()
            time.sleep(COM_CLIPBOARD_DELAY / 2)

            chart_obj.CopyPicture(Appearance=1, Format=2)
            time.sleep(COM_CLIPBOARD_DELAY)

            temp_chart_sheet = workbook.Charts.Add()
            time.sleep(0.1)
            try:
                temp_chart_sheet.Paste()
                time.sleep(COM_CLIPBOARD_DELAY)
                temp_chart_sheet.Export(output_path, "PNG")
            finally:
                excel_app.DisplayAlerts = False
                temp_chart_sheet.Delete()

            if validate_image(output_path, min_size=500):
                logger.info(
                    "  [Worksheet] CopyPicture succeeded on attempt %d", attempt + 1
                )
                return True

            logger.info(
                "  [Worksheet] Attempt %d: CopyPicture validation failed", attempt + 1
            )
            if os.path.exists(output_path):
                os.remove(output_path)

        except Exception as e:
            logger.warning("  [Worksheet] Attempt %d CopyPicture failed: %s", attempt + 1, e)
            time.sleep(COM_RETRY_DELAY)

    # Try 3: UsedRange fallback
    logger.info("  [Worksheet] Trying UsedRange CopyPicture as last resort...")
    try:
        used_range = sheet.UsedRange
        if _export_via_copypicture(workbook, excel_app, used_range, output_path):
            if validate_image(output_path, min_size=500):
                logger.info("  [Worksheet] UsedRange fallback succeeded")
                return True
    except Exception as e:
        logger.warning("  [Worksheet] UsedRange fallback failed: %s", e)

    logger.warning("  [Worksheet] All methods failed for '%s'", name)
    return False


def _capture_worksheet_range(
    excel_app, workbook, sheet, name: str, output_path: str, max_retries: int
) -> bool:
    """Capture the UsedRange of a worksheet without charts."""
    logger.info("  [Worksheet] No charts found, capturing UsedRange...")

    for attempt in range(max_retries):
        try:
            clear_clipboard()
            time.sleep(COM_CLIPBOARD_DELAY / 2)

            used_range = sheet.UsedRange
            if used_range.Rows.Count == 0 or used_range.Columns.Count == 0:
                logger.warning("  [Worksheet] Sheet '%s' appears empty", name)
                return False

            logger.info(
                "  [Worksheet] Attempt %d: UsedRange = %d rows x %d cols",
                attempt + 1,
                used_range.Rows.Count,
                used_range.Columns.Count,
            )

            used_range.CopyPicture(Appearance=1, Format=2)
            time.sleep(COM_CLIPBOARD_DELAY)

            temp_chart_sheet = workbook.Charts.Add()
            time.sleep(0.1)
            try:
                temp_chart_sheet.Paste()
                time.sleep(COM_CLIPBOARD_DELAY)
                temp_chart_sheet.Export(output_path, "PNG")
            finally:
                excel_app.DisplayAlerts = False
                temp_chart_sheet.Delete()

            if validate_image(output_path, min_size=500):
                return True

            logger.info("  [Worksheet] Attempt %d: validation failed", attempt + 1)
            if os.path.exists(output_path):
                os.remove(output_path)

        except Exception as e:
            logger.warning("  [Worksheet] Attempt %d failed: %s", attempt + 1, e)
            time.sleep(COM_RETRY_DELAY)

    logger.warning(
        "  [Worksheet] Failed to capture '%s' after %d attempts", name, max_retries
    )
    return False
