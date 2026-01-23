# Editable Charts Test

This folder contains a test script to verify that Excel charts can be inserted into PowerPoint as **editable objects** instead of static images.

## Problem

The current implementation exports Excel charts as PNG images and inserts them using `slide.shapes.add_picture()`. This creates static pictures that **cannot be edited** in PowerPoint - you can't change data, delete bars, or modify formatting.

## Solution

Use Windows COM automation for both Excel AND PowerPoint to copy/paste charts as editable objects.

## Requirements

- Windows OS
- Microsoft Excel (installed and licensed)
- Microsoft PowerPoint (installed and licensed)
- Python 3.8+
- pywin32 (`pip install pywin32`)
- openpyxl (`pip install openpyxl`) - for creating sample Excel

## Usage

### Basic Usage (creates sample Excel automatically)

```bash
cd test_editable_charts
python test_editable_chart.py
```

### Using Your Own Excel File

```bash
python test_editable_chart.py --excel "C:\path\to\your\file.xlsx"
```

### Specify Output Directory

```bash
python test_editable_chart.py --output "C:\path\to\output"
```

## What the Test Does

1. **Creates a sample Excel file** with a bar chart (if not provided)
2. **Tests 3 paste modes:**
   - **Embedded (ppPasteOLEObject)** - Chart data stored in PPT, fully editable
   - **Linked (ppPasteLink)** - Chart linked to source Excel file
   - **Image (for comparison)** - Static picture, not editable

3. **Generates 3 PowerPoint files:**
   - `test_embedded_chart.pptx`
   - `test_linked_chart.pptx`
   - `test_image_chart.pptx`

## How to Verify Results

1. Open each generated PowerPoint file
2. Click on the chart
3. **For EMBEDDED/LINKED:**
   - Right-click should show 'Edit Data' option
   - Double-click should open chart for editing
   - You can modify colors, labels, add/remove data
4. **For IMAGE:**
   - Chart should NOT be editable (just a picture)
   - No 'Edit Data' option available

## Key Differences

| Mode | Editable | File Size | Excel Required |
|------|----------|-----------|----------------|
| Embedded | Yes | Larger | No (data in PPT) |
| Linked | Yes | Smaller | Yes (linked to source) |
| Image | No | Smallest | No |

### When to Use Each Mode

- **Embedded**: When you want editable charts that work standalone without the Excel file
- **Linked**: When you want charts that auto-update when the Excel source changes
- **Image**: When you just need a static snapshot (current behavior)

## Technical Details

### PowerPoint Paste Types

```python
PP_PASTE_DEFAULT = 0       # Default paste
PP_PASTE_BITMAP = 1        # Bitmap image
PP_PASTE_LINK = 2          # Linked to source
PP_PASTE_METAFILE = 3      # Enhanced metafile
PP_PASTE_OLE_OBJECT = 10   # Embedded OLE object
```

### Key API Calls

**Excel - Copy Chart:**
```python
chart_obj.Copy()  # For embedded charts
chart_sheet.ChartArea.Copy()  # For chart sheets
```

**PowerPoint - Paste as Editable:**
```python
# Embedded (data stored in PPT)
slide.Shapes.PasteSpecial(DataType=10)

# Linked (connected to Excel)
slide.Shapes.PasteSpecial(DataType=0, Link=True)
```

## Troubleshooting

### "No charts found in the Excel file"
- Make sure your Excel file contains at least one chart
- The chart can be embedded in a worksheet or a standalone chart sheet

### "PowerPoint paste failed"
- Make sure PowerPoint is not already open with the same file
- Try closing all Excel and PowerPoint windows and run again
- Check that both Excel and PowerPoint are properly licensed

### "pywin32 not found"
```bash
pip install pywin32
```

### "openpyxl not found" (for sample Excel creation)
```bash
pip install openpyxl
```

## Next Steps

After verifying this test works, the functionality can be integrated into the main application by:

1. Modifying `app/main.py` to add paste type option
2. Updating the frontend (`static/index.html`) to allow users to choose between Image/Embedded/Linked
