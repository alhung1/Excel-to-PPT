"""
Excel Chart Extractor using Windows COM Automation
Captures charts from Excel files as images
"""
import os
import tempfile
import pythoncom
import win32com.client as win32
from pathlib import Path
from typing import List, Dict, Optional
import shutil


class ExcelChartExtractor:
    """Extract charts from Excel files using Windows COM automation"""
    
    def __init__(self):
        self.excel_app = None
        
    def __enter__(self):
        """Initialize Excel application"""
        pythoncom.CoInitialize()
        self.excel_app = win32.DispatchEx('Excel.Application')
        self.excel_app.Visible = False
        self.excel_app.DisplayAlerts = False
        return self
    
    def __exit__(self, exc_type, exc_val, exc_tb):
        """Clean up Excel application"""
        if self.excel_app:
            try:
                self.excel_app.Quit()
            except:
                pass
        pythoncom.CoUninitialize()
    
    def extract_charts_from_file(
        self, 
        excel_path: str, 
        output_dir: str,
        sheet_names: Optional[List[str]] = None
    ) -> List[Dict]:
        """
        Extract all charts from an Excel file
        
        Args:
            excel_path: Path to the Excel file
            output_dir: Directory to save chart images
            sheet_names: Optional list of specific sheet names to process
            
        Returns:
            List of dicts with chart info: {name, sheet, image_path}
        """
        excel_path = os.path.abspath(excel_path)
        output_dir = os.path.abspath(output_dir)
        os.makedirs(output_dir, exist_ok=True)
        
        extracted_charts = []
        workbook = None
        
        try:
            workbook = self.excel_app.Workbooks.Open(excel_path)
            
            for sheet in workbook.Worksheets:
                # Skip if specific sheets requested and this isn't one
                if sheet_names and sheet.Name not in sheet_names:
                    continue
                
                # Extract embedded charts (ChartObjects)
                for i, chart_obj in enumerate(sheet.ChartObjects(), 1):
                    chart_name = chart_obj.Name or f"Chart_{i}"
                    safe_name = self._safe_filename(f"{sheet.Name}_{chart_name}")
                    image_path = os.path.join(output_dir, f"{safe_name}.png")
                    
                    # Export chart as image
                    chart_obj.Chart.Export(image_path, "PNG")
                    
                    extracted_charts.append({
                        'name': chart_name,
                        'sheet': sheet.Name,
                        'image_path': image_path,
                        'type': 'embedded'
                    })
                
            # Also check for chart sheets (entire sheet is a chart)
            for sheet in workbook.Charts:
                if sheet_names and sheet.Name not in sheet_names:
                    continue
                    
                safe_name = self._safe_filename(f"ChartSheet_{sheet.Name}")
                image_path = os.path.join(output_dir, f"{safe_name}.png")
                
                sheet.Export(image_path, "PNG")
                
                extracted_charts.append({
                    'name': sheet.Name,
                    'sheet': sheet.Name,
                    'image_path': image_path,
                    'type': 'chart_sheet'
                })
                
        finally:
            if workbook:
                workbook.Close(SaveChanges=False)
        
        return extracted_charts
    
    def capture_sheet_as_image(
        self,
        excel_path: str,
        sheet_name: str,
        output_path: str,
        range_address: Optional[str] = None
    ) -> str:
        """
        Capture an entire sheet or specific range as an image
        
        Args:
            excel_path: Path to Excel file
            sheet_name: Name of sheet to capture
            output_path: Where to save the image
            range_address: Optional range like "A1:H20", captures used range if None
            
        Returns:
            Path to saved image
        """
        excel_path = os.path.abspath(excel_path)
        output_path = os.path.abspath(output_path)
        os.makedirs(os.path.dirname(output_path), exist_ok=True)
        
        workbook = None
        
        try:
            workbook = self.excel_app.Workbooks.Open(excel_path)
            sheet = workbook.Worksheets(sheet_name)
            
            # Get range to capture
            if range_address:
                capture_range = sheet.Range(range_address)
            else:
                capture_range = sheet.UsedRange
            
            # Copy range as picture and export
            capture_range.CopyPicture(Format=2)  # xlBitmap = 2
            
            # Create a temporary chart to paste and export
            temp_chart = sheet.ChartObjects().Add(0, 0, capture_range.Width, capture_range.Height)
            temp_chart.Chart.Paste()
            temp_chart.Chart.Export(output_path, "PNG")
            temp_chart.Delete()
            
        finally:
            if workbook:
                workbook.Close(SaveChanges=False)
        
        return output_path
    
    def get_sheet_names(self, excel_path: str) -> List[str]:
        """Get list of all sheet names in an Excel file"""
        excel_path = os.path.abspath(excel_path)
        workbook = None
        
        try:
            workbook = self.excel_app.Workbooks.Open(excel_path)
            sheets = [sheet.Name for sheet in workbook.Worksheets]
            chart_sheets = [sheet.Name for sheet in workbook.Charts]
            return sheets + chart_sheets
        finally:
            if workbook:
                workbook.Close(SaveChanges=False)
    
    def get_charts_info(self, excel_path: str) -> List[Dict]:
        """Get info about all charts in an Excel file without exporting"""
        excel_path = os.path.abspath(excel_path)
        workbook = None
        charts_info = []
        
        try:
            workbook = self.excel_app.Workbooks.Open(excel_path)
            
            for sheet in workbook.Worksheets:
                for chart_obj in sheet.ChartObjects():
                    charts_info.append({
                        'name': chart_obj.Name,
                        'sheet': sheet.Name,
                        'type': 'embedded',
                        'width': chart_obj.Width,
                        'height': chart_obj.Height
                    })
            
            for sheet in workbook.Charts:
                charts_info.append({
                    'name': sheet.Name,
                    'sheet': sheet.Name,
                    'type': 'chart_sheet'
                })
                
        finally:
            if workbook:
                workbook.Close(SaveChanges=False)
        
        return charts_info
    
    @staticmethod
    def _safe_filename(name: str) -> str:
        """Convert string to safe filename"""
        invalid_chars = '<>:"/\\|?*'
        for char in invalid_chars:
            name = name.replace(char, '_')
        return name


def extract_all_charts(excel_files: List[str], output_dir: str) -> Dict[str, List[Dict]]:
    """
    Extract charts from multiple Excel files
    
    Args:
        excel_files: List of Excel file paths
        output_dir: Directory to save extracted chart images
        
    Returns:
        Dict mapping filename to list of extracted chart info
    """
    results = {}
    
    with ExcelChartExtractor() as extractor:
        for excel_file in excel_files:
            filename = os.path.basename(excel_file)
            file_output_dir = os.path.join(output_dir, Path(filename).stem)
            
            try:
                charts = extractor.extract_charts_from_file(excel_file, file_output_dir)
                results[filename] = charts
            except Exception as e:
                results[filename] = {'error': str(e)}
    
    return results

