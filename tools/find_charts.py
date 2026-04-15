"""
Find all charts in Excel workbook
"""
import pythoncom
import win32com.client as win32

EXCEL_FILE = r"C:\Netgear Projects\NRH testing\RBE770v2\NRH Test Result\RBE773v2 vs RBE773 vs Eero Pro 7_3Pack Test\NRH_RBE773v2 vs RBE773 vs Eero Pro 7_auto channle_Intel BE200_20251016.xlsm"

pythoncom.CoInitialize()

excel_app = win32.DispatchEx('Excel.Application')
excel_app.Visible = False
excel_app.DisplayAlerts = False

workbook = excel_app.Workbooks.Open(EXCEL_FILE)

print("=" * 60)
print("All Sheets and Charts in Workbook")
print("=" * 60)

# Check all worksheets
for sheet in workbook.Worksheets:
    print(f"\nSheet: {sheet.Name}")
    print(f"  Visible: {sheet.Visible}")
    
    # Check for chart objects
    try:
        chart_count = sheet.ChartObjects().Count
        if chart_count > 0:
            print(f"  Charts ({chart_count}):")
            for i in range(1, chart_count + 1):
                chart_obj = sheet.ChartObjects(i)
                print(f"    - Name: '{chart_obj.Name}'")
                print(f"      Size: {chart_obj.Width:.0f} x {chart_obj.Height:.0f}")
    except Exception as e:
        print(f"  Error reading charts: {e}")

# Check for chart sheets (standalone chart sheets)
print("\n" + "=" * 60)
print("Chart Sheets (standalone)")
print("=" * 60)
try:
    for chart_sheet in workbook.Charts:
        print(f"  Chart Sheet: {chart_sheet.Name}")
except Exception as e:
    print(f"  No chart sheets or error: {e}")

workbook.Close(SaveChanges=False)
excel_app.Quit()
pythoncom.CoUninitialize()


