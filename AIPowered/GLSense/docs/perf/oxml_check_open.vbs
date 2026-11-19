' check_open.vbs <file.xlsx>  - opens in a hidden Excel, reports repair + cell contents, closes without saving.
Dim fso, path, xl, wb, ws, tmp, before, f, repaired, cap
Set fso = CreateObject("Scripting.FileSystemObject")
path = WScript.Arguments(0)
tmp = CreateObject("WScript.Shell").ExpandEnvironmentStrings("%TEMP%")
Set before = CreateObject("Scripting.Dictionary")
For Each f In fso.GetFolder(tmp).Files
  If LCase(Left(f.Name, 5)) = "error" And LCase(Right(f.Name, 4)) = ".xml" Then before(f.Name) = f.DateLastModified
Next
Set xl = CreateObject("Excel.Application")
xl.DisplayAlerts = False
On Error Resume Next
Set wb = xl.Workbooks.Open(path)
If Err.Number <> 0 Then WScript.Echo "OPEN FAILED: " & Err.Description : xl.Quit : WScript.Quit 2
On Error GoTo 0
cap = wb.Windows(1).Caption
repaired = (InStr(1, cap, "Repaired", 1) > 0)
For Each f In fso.GetFolder(tmp).Files
  If LCase(Left(f.Name, 5)) = "error" And LCase(Right(f.Name, 4)) = ".xml" Then
    If Not before.Exists(f.Name) Or before(f.Name) <> f.DateLastModified Then
      repaired = True
      WScript.Echo "REPAIR LOG " & f.Name & ": " & Replace(fso.OpenTextFile(f.Path).ReadAll, vbCrLf, " ")
    End If
  End If
Next
WScript.Echo "caption=" & cap & "  repaired=" & repaired & "  sheets=" & wb.Worksheets.Count & "  calcMode=" & xl.Calculation
For Each ws In wb.Worksheets
  Dim glsense, formulas, c
  glsense = 0 : formulas = 0
  For Each c In ws.UsedRange.Cells
    If c.HasFormula Then
      formulas = formulas + 1
      If InStr(1, c.Formula, "GLSense", 1) > 0 Then glsense = glsense + 1
    End If
  Next
  WScript.Echo "  [" & ws.Name & "] formulas=" & formulas & " glsense=" & glsense
  For Each a In Array("B3", "C3", "D3", "E3", "F3", "C7", "D7", "E7")
    Set c = ws.Range(a)
    WScript.Echo "     " & a & ": formula=" & c.Formula & " | value=" & c.Text & " | type=" & TypeName(c.Value) & " | numfmt=" & c.NumberFormat
  Next
Next
wb.Close False
xl.Quit
