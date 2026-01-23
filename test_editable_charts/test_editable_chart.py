"""
Test Script: Editable Charts from Excel to PowerPoint

This script tests copying Excel charts to PowerPoint as editable objects
instead of static images.

Three paste modes are tested:
1. Embedded (ppPasteOLEObject) - Chart data stored in PPT, fully editable
2. Linked (ppPasteLink) - Chart linked to source Excel file
3. Image (for comparison) - Static picture, not editable

Requirements:
- Windows OS
- Microsoft Excel installed
- Microsoft PowerPoint installed
- pywin32 package

Usage:
    python test_editable_chart.py [--excel path_to_excel.xlsx]
"""

import os
import sys
import time
import argparse
from pathlib import Path

# Check for pywin32
try:
    import pythoncom
    import win32com.client as win32
except ImportError:
    print("ERROR: pywin32 is required. Install with: pip install pywin32")
    sys.exit(1)

# Check for openpyxl (for creating sample Excel)
try:
    import openpyxl
    from openpyxl.chart import BarChart, Reference
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False
    print("WARNING: openpyxl not found. Cannot create sample Excel file.")


# PowerPoint paste types
PP_PASTE_DEFAULT = 0
PP_PASTE_BITMAP = 1
PP_PASTE_LINK = 2  # Linked to source
PP_PASTE_METAFILE = 3
PP_PASTE_OLE_OBJECT = 10  # Embedded, editable


def create_sample_excel(output_path: str) -> str:
    """Create a sample Excel file with a bar chart for testing"""
    if not HAS_OPENPYXL:
        raise RuntimeError("openpyxl is required to create sample Excel")
    
    print(f"Creating sample Excel file: {output_path}")
    
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sales Data"
    
    # Add sample data
    data = [
        ["Product", "Q1", "Q2", "Q3", "Q4"],
        ["Product A", 150, 200, 180, 220],
        ["Product B", 100, 150, 170, 190],
        ["Product C", 80, 120, 140, 160],
        ["Product D", 200, 180, 210, 250],
    ]
    
    for row_idx, row_data in enumerate(data, 1):
        for col_idx, value in enumerate(row_data, 1):
            ws.cell(row=row_idx, column=col_idx, value=value)
    
    # Create a bar chart
    chart = BarChart()
    chart.type = "col"
    chart.grouping = "clustered"
    chart.title = "Quarterly Sales by Product"
    chart.y_axis.title = "Sales ($)"
    chart.x_axis.title = "Product"
    
    # Data references
    data_ref = Reference(ws, min_col=2, min_row=1, max_col=5, max_row=5)
    categories = Reference(ws, min_col=1, min_row=2, max_row=5)
    
    chart.add_data(data_ref, titles_from_data=True)
    chart.set_categories(categories)
    chart.shape = 4  # Rectangle shape
    
    # Position and size
    chart.width = 15
    chart.height = 10
    
    ws.add_chart(chart, "G2")
    
    wb.save(output_path)
    print(f"  Created Excel with bar chart at {output_path}")
    
    return output_path


def clear_clipboard():
    """Clear the Windows clipboard"""
    try:
        import win32clipboard
        win32clipboard.OpenClipboard()
        win32clipboard.EmptyClipboard()
        win32clipboard.CloseClipboard()
    except:
        pass


def test_editable_charts(excel_path: str, output_dir: str):
    """
    Test copying Excel chart to PowerPoint in different modes
    
    Args:
        excel_path: Path to Excel file with chart
        output_dir: Directory to save output PowerPoint files
    """
    excel_path = os.path.abspath(excel_path)
    output_dir = os.path.abspath(output_dir)
    os.makedirs(output_dir, exist_ok=True)
    
    print("\n" + "=" * 60)
    print("Testing Editable Charts: Excel to PowerPoint")
    print("=" * 60)
    print(f"Excel source: {excel_path}")
    print(f"Output dir: {output_dir}")
    
    # Initialize COM
    pythoncom.CoInitialize()
    
    excel_app = None
    ppt_app = None
    workbook = None
    presentation = None
    
    try:
        # Start Excel - MUST be visible for clipboard operations to work properly
        print("\n[1] Starting Excel...")
        excel_app = win32.DispatchEx('Excel.Application')
        excel_app.Visible = True  # Must be visible for copy to work
        excel_app.DisplayAlerts = False
        
        # Open workbook
        print(f"[2] Opening workbook: {excel_path}")
        workbook = excel_app.Workbooks.Open(excel_path)
        
        # Find charts
        chart_found = False
        chart_obj = None
        chart_sheet = None
        source_sheet = None
        
        # Check worksheets for embedded charts
        for sheet in workbook.Worksheets:
            try:
                count = sheet.ChartObjects().Count
                if count > 0:
                    chart_obj = sheet.ChartObjects(1)
                    source_sheet = sheet
                    print(f"[3] Found embedded chart '{chart_obj.Name}' in sheet '{sheet.Name}'")
                    chart_found = True
                    break
            except:
                pass
        
        # Check for chart sheets
        if not chart_found:
            try:
                for cs in workbook.Charts:
                    chart_sheet = cs
                    print(f"[3] Found chart sheet: '{cs.Name}'")
                    chart_found = True
                    break
            except:
                pass
        
        if not chart_found:
            print("ERROR: No charts found in the Excel file!")
            return False
        
        # Start PowerPoint - MUST be visible for paste operations
        print("\n[4] Starting PowerPoint...")
        ppt_app = win32.DispatchEx('PowerPoint.Application')
        ppt_app.Visible = True  # Required for paste operations
        
        results = []
        
        # ============================================================
        # Test 1: Embedded Chart (using Paste, not PasteSpecial)
        # ============================================================
        print("\n" + "-" * 40)
        print("Test 1: EMBEDDED Chart (Paste as Microsoft Office Graphic Object)")
        print("-" * 40)
        
        try:
            # Create new presentation
            presentation = ppt_app.Presentations.Add()
            slide = presentation.Slides.Add(1, 12)  # ppLayoutBlank = 12
            
            # Clear clipboard first
            clear_clipboard()
            time.sleep(0.2)
            
            # Activate the sheet and select the chart
            if chart_obj:
                source_sheet.Activate()
                chart_obj.Select()
                time.sleep(0.2)
                # Copy using Chart.ChartArea.Copy() for better clipboard handling
                chart_obj.Chart.ChartArea.Copy()
            else:
                chart_sheet.Activate()
                chart_sheet.ChartArea.Copy()
            
            time.sleep(0.5)  # Wait for clipboard
            
            # Use Paste() instead of PasteSpecial - this pastes as editable chart
            shape = slide.Shapes.Paste()
            time.sleep(0.3)
            
            # Position the chart (Paste returns a ShapeRange, get the first shape)
            if hasattr(shape, 'Item'):
                shape = shape.Item(1)
            shape.Left = 50
            shape.Top = 100
            shape.Width = 600
            shape.Height = 400
            
            # Save
            output_path = os.path.join(output_dir, "test_embedded_chart.pptx")
            presentation.SaveAs(output_path)
            presentation.Close()
            presentation = None
            
            file_size = os.path.getsize(output_path)
            print(f"  SUCCESS: Saved to {output_path}")
            print(f"  File size: {file_size:,} bytes")
            results.append(("Embedded", output_path, file_size, True))
            
        except Exception as e:
            print(f"  FAILED: {e}")
            import traceback
            traceback.print_exc()
            results.append(("Embedded", None, 0, False))
            if presentation:
                try:
                    presentation.Close()
                except:
                    pass
                presentation = None
        
        # ============================================================
        # Test 2: Linked Chart (using PasteSpecial with Link=True)
        # ============================================================
        print("\n" + "-" * 40)
        print("Test 2: LINKED Chart (PasteSpecial with Link)")
        print("-" * 40)
        
        try:
            # Create new presentation
            presentation = ppt_app.Presentations.Add()
            slide = presentation.Slides.Add(1, 12)  # ppLayoutBlank = 12
            
            # Clear clipboard first
            clear_clipboard()
            time.sleep(0.2)
            
            # Select and copy chart
            if chart_obj:
                source_sheet.Activate()
                chart_obj.Select()
                time.sleep(0.2)
                chart_obj.Copy()  # Use ChartObject.Copy() for linked
            else:
                chart_sheet.Activate()
                chart_sheet.Copy()
            
            time.sleep(0.5)
            
            # Try PasteSpecial with Link - use ppPasteMetafilePicture with Link for best compatibility
            # Link=True creates a link to source, Link=False embeds
            try:
                shape = slide.Shapes.PasteSpecial(DataType=PP_PASTE_DEFAULT, Link=-1)  # -1 = msoTrue
            except:
                # Fallback: try without DataType
                shape = slide.Shapes.Paste()
            
            time.sleep(0.3)
            
            # Position the chart
            if hasattr(shape, 'Item'):
                shape = shape.Item(1)
            shape.Left = 50
            shape.Top = 100
            shape.Width = 600
            shape.Height = 400
            
            # Save
            output_path = os.path.join(output_dir, "test_linked_chart.pptx")
            presentation.SaveAs(output_path)
            presentation.Close()
            presentation = None
            
            file_size = os.path.getsize(output_path)
            print(f"  SUCCESS: Saved to {output_path}")
            print(f"  File size: {file_size:,} bytes")
            print(f"  NOTE: This chart is linked to: {excel_path}")
            results.append(("Linked", output_path, file_size, True))
            
        except Exception as e:
            print(f"  FAILED: {e}")
            import traceback
            traceback.print_exc()
            results.append(("Linked", None, 0, False))
            if presentation:
                try:
                    presentation.Close()
                except:
                    pass
                presentation = None
        
        # ============================================================
        # Test 3: Image (for comparison)
        # ============================================================
        print("\n" + "-" * 40)
        print("Test 3: IMAGE (static picture, for comparison)")
        print("-" * 40)
        
        try:
            # Create new presentation
            presentation = ppt_app.Presentations.Add()
            slide = presentation.Slides.Add(1, 12)
            
            # Export chart as image first
            temp_image = os.path.join(output_dir, "temp_chart.png")
            if chart_obj:
                chart_obj.Chart.Export(temp_image, "PNG")
            else:
                chart_sheet.Export(temp_image, "PNG")
            
            time.sleep(0.3)
            
            # Add as picture
            shape = slide.Shapes.AddPicture(
                temp_image,
                LinkToFile=False,
                SaveWithDocument=True,
                Left=50,
                Top=100,
                Width=600,
                Height=400
            )
            
            # Save
            output_path = os.path.join(output_dir, "test_image_chart.pptx")
            presentation.SaveAs(output_path)
            
            # Close with error handling
            try:
                presentation.Close()
            except:
                pass  # Ignore close errors, file is already saved
            presentation = None
            
            time.sleep(0.3)  # Wait for file system
            
            # Clean up temp image
            try:
                if os.path.exists(temp_image):
                    os.remove(temp_image)
            except:
                pass  # Ignore cleanup errors
            
            file_size = os.path.getsize(output_path) if os.path.exists(output_path) else 0
            if file_size > 0:
                print(f"  SUCCESS: Saved to {output_path}")
                print(f"  File size: {file_size:,} bytes")
                results.append(("Image", output_path, file_size, True))
            else:
                print(f"  FAILED: File not created or empty")
                results.append(("Image", None, 0, False))
            
        except Exception as e:
            print(f"  FAILED: {e}")
            results.append(("Image", None, 0, False))
            if presentation:
                try:
                    presentation.Close()
                except:
                    pass
                presentation = None
        
        # ============================================================
        # Summary
        # ============================================================
        print("\n" + "=" * 60)
        print("SUMMARY")
        print("=" * 60)
        print(f"{'Mode':<12} {'Status':<10} {'File Size':>15} {'Editable':<10}")
        print("-" * 60)
        
        for mode, path, size, success in results:
            status = "SUCCESS" if success else "FAILED"
            size_str = f"{size:,} bytes" if size > 0 else "N/A"
            editable = "YES" if mode in ["Embedded", "Linked"] and success else "NO"
            print(f"{mode:<12} {status:<10} {size_str:>15} {editable:<10}")
        
        print("\n" + "=" * 60)
        print("VERIFICATION STEPS")
        print("=" * 60)
        print("""
1. Open each generated PowerPoint file
2. Click on the chart
3. For EMBEDDED/LINKED: You should see chart editing options
   - Right-click should show 'Edit Data' option
   - Double-click should open chart for editing
4. For IMAGE: Chart should NOT be editable (just a picture)

Key differences:
- EMBEDDED: Data is stored in PPT, no Excel link needed
- LINKED: Changes in Excel will update the chart in PPT
- IMAGE: Static picture, cannot modify data or formatting
""")
        
        return True
        
    except Exception as e:
        print(f"\nERROR: {e}")
        import traceback
        traceback.print_exc()
        return False
        
    finally:
        # Cleanup
        if presentation:
            try:
                presentation.Close()
            except:
                pass
        
        if workbook:
            try:
                workbook.Close(SaveChanges=False)
            except:
                pass
        
        if excel_app:
            try:
                excel_app.Quit()
            except:
                pass
        
        if ppt_app:
            try:
                ppt_app.Quit()
            except:
                pass
        
        pythoncom.CoUninitialize()


def main():
    parser = argparse.ArgumentParser(
        description="Test editable charts from Excel to PowerPoint"
    )
    parser.add_argument(
        "--excel", "-e",
        type=str,
        help="Path to Excel file with chart. If not provided, a sample will be created."
    )
    parser.add_argument(
        "--output", "-o",
        type=str,
        default=".",
        help="Output directory for generated files (default: current directory)"
    )
    
    args = parser.parse_args()
    
    # Determine script directory
    script_dir = Path(__file__).parent
    output_dir = Path(args.output).resolve() if args.output != "." else script_dir
    
    # Determine Excel file path
    if args.excel:
        excel_path = Path(args.excel).resolve()
        if not excel_path.exists():
            print(f"ERROR: Excel file not found: {excel_path}")
            sys.exit(1)
    else:
        # Create sample Excel
        excel_path = script_dir / "sample_chart.xlsx"
        if not excel_path.exists():
            if not HAS_OPENPYXL:
                print("ERROR: No Excel file provided and openpyxl not available to create one.")
                print("Please provide an Excel file with --excel option or install openpyxl.")
                sys.exit(1)
            create_sample_excel(str(excel_path))
        else:
            print(f"Using existing sample Excel: {excel_path}")
    
    # Run tests
    success = test_editable_charts(str(excel_path), str(output_dir))
    
    if success:
        print("\n[DONE] Test completed. Check the generated PowerPoint files.")
    else:
        print("\n[ERROR] Test failed. Check the error messages above.")
        sys.exit(1)


if __name__ == "__main__":
    main()
