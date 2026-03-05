# XLSM Technical Reference

## XLSM = ZIP Archive

An `.xlsm` file is a standard ZIP archive identical to `.xlsx` except for VBA support.

## ZIP Contents

```
[Content_Types].xml          ← declares all part types
_rels/.rels                  ← top-level relationships
docProps/
  app.xml                    ← application metadata
  core.xml                   ← author, dates
xl/
  workbook.xml               ← sheet list, defined names
  styles.xml                 ← cell formats, fonts, fills
  sharedStrings.xml          ← deduplicated string table
  theme/theme1.xml           ← color/font theme
  vbaProject.bin             ← OLE binary with VBA code  ★ XLSM-SPECIFIC
  calcChain.xml              ← formula calculation order
  _rels/workbook.xml.rels    ← workbook relationships
  worksheets/
    sheet1.xml ... sheetN.xml
    _rels/
      sheet1.xml.rels ...    ← per-sheet relationships (printerSettings)
  printerSettings/
    printerSettings1.bin ...
```

## XLSX → XLSM: 3 Changes

### 1. Add `xl/vbaProject.bin`

OLE Compound Document containing VBA source code.

### 2. Update `[Content_Types].xml`

```xml
<!-- Change workbook ContentType -->
<Override PartName="/xl/workbook.xml"
  ContentType="application/vnd.ms-excel.sheet.macroEnabled.main+xml"/>
<!--                              ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^ was: sheet.main+xml -->

<!-- Add vbaProject override -->
<Override PartName="/xl/vbaProject.bin"
  ContentType="application/vnd.ms-office.vbaProject"/>
```

### 3. Add Relationship in `xl/_rels/workbook.xml.rels`

```xml
<Relationship Id="rIdN"
  Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject"
  Target="vbaProject.bin"/>
```

## vbaProject.bin OLE Structure

```
Root Entry
├── PROJECT             (text) — INI-like project description
├── PROJECTwm           (binary) — unicode module name mapping
└── VBA/                (storage)
    ├── _VBA_PROJECT    (binary) — version header
    ├── dir             (compressed) — project metadata + module list
    ├── ЭтаКнига        (compressed) — ThisWorkbook source
    ├── Лист1           (compressed) — Sheet1 source
    └── Module1         (compressed) — standard module source
```

### PROJECT Stream Format (text)

```ini
ID="{GUID}"
Document=ThisWorkbook/&H00000000
Document=Sheet1/&H00000000
Module=Module1
Name="VBAProject"
VersionCompatible32="393222000"
CMG="..."
DPB="..."
GC="..."

[Host Extender Info]
&H00000001={3832D640-CF90-11CF-8E43-00A0C911005A};VBE;&H00000000

[Workspace]
ThisWorkbook=0, 0, 0, 0,
Sheet1=0, 0, 0, 0, C
```

Entry types:
- `Document=Name/&H00000000` — DocModule (ThisWorkbook, sheets)
- `Module=Name` — StdModule (regular .bas)
- `Class=Name` — ClassModule (.cls)

### Module Source Compression

All module streams use MS-OVBA compression (RLE variant). Python: `ms-ovba-compression` package.

## openpyxl VBA Handling

```python
# Read xlsm preserving VBA
from openpyxl import load_workbook
wb = load_workbook('file.xlsm', keep_vba=True)
# wb.vba_archive contains the vbaProject.bin

# Modify data
ws = wb.active
ws['A1'] = 'Modified'

# Save as xlsm (VBA preserved automatically)
wb.save('output.xlsm')
```

Key points:
- `keep_vba=True` stores entire vbaProject.bin in `wb.vba_archive`
- Cannot modify VBA code through openpyxl — only preserves it
- Must save as `.xlsm` (not `.xlsx`) to keep VBA

## olevba VBA Extraction

```bash
# Extract all VBA modules from xlsm
olevba file.xlsm

# Decode hex strings
olevba --decode file.xlsm

# Output only VBA code (no analysis)
olevba --code file.xlsm
```

## Standard VBA References

Most Excel VBA projects need these library references:

| Library | CLSID | Version |
|---------|-------|---------|
| stdole (OLE Automation) | `{00020430-0000-0000-C000-000000000046}` | 2.0 |
| Office (MSO) | `{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}` | 2.8 |
| Excel | `{00020813-0000-0000-C000-000000000046}` | 1.9 |

## Common VBA Patterns for Excel

### Performance Optimization
```vba
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
' ... work ...
Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True
```

### Error Handling
```vba
On Error Resume Next
' risky operation
If Err.Number <> 0 Then
    ' handle
    Err.Clear
End If
On Error GoTo 0
```

### Find Sheet by Name Pattern
```vba
Function FindSheetByKeyword(wb As Workbook, keyword As String) As Worksheet
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If InStr(1, LCase(ws.Name), LCase(keyword)) > 0 Then
            Set FindSheetByKeyword = ws
            Exit Function
        End If
    Next ws
    Set FindSheetByKeyword = Nothing
End Function
```

### Color Constants
```vba
Dim clrGreen As Long:  clrGreen = RGB(198, 239, 206)  ' light green
Dim clrYellow As Long: clrYellow = RGB(255, 255, 153)  ' light yellow
Dim clrBlue As Long:   clrBlue = RGB(189, 215, 238)    ' light blue
Dim clrRed As Long:    clrRed = RGB(255, 199, 206)     ' light red/pink
```
