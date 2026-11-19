' open_save_time.vbs <file.xlsx> <check|save> : times a hidden Excel open (+save), reports repair and formula counts.
Dim fso, path, mode, xl, wb, ws, t0, tOpen, tSave, tmp, before, f, repaired, fc, gl, total, cells
Set fso = CreateObject("Scripting.FileSystemObject")
path = WScript.Arguments(0) : mode = WScript.Arguments(1)
tmp = CreateObject("WScript.Shell").ExpandEnvironmentStrings("%TEMP%")
Set before = CreateObject("Scripting.Dictionary")
For Each f In fso.GetFolder(tmp).Files
  If LCase(Left(f.Name, 5)) = "error" And LCase(Right(f.Name, 4)) = ".xml" Then before(f.Name) = f.DateLastModified
Next
t0 = Timer
Set xl = CreateObject("Excel.Application")
xl.DisplayAlerts = False
On Error Resume Next
Set wb = xl.Workbooks.Open(path)
If Err.Number <> 0 Then WScript.Echo "OPEN FAILED (repair needed): " & Err.Description : xl.Quit : WScript.Quit 2
On Error GoTo 0
xl.Calculation = -4135 ' manual, as the snapshot runner
tOpen = Timer - t0
repaired = (InStr(1, wb.Windows(1).Caption, "Repaired", 1) > 0)
For Each f In fso.GetFolder(tmp).Files
  If LCase(Left(f.Name, 5)) = "error" And LCase(Right(f.Name, 4)) = ".xml" Then
    If Not before.Exists(f.Name) Or before(f.Name) <> f.DateLastModified Then repaired = True : WScript.Echo "REPAIR LOG: " & Left(Replace(fso.OpenTextFile(f.Path).ReadAll, vbCrLf, " "), 600)
  End If
Next
If mode = "check" Then
  total = 0
  For Each ws In wb.Worksheets
    fc = 0 : gl = 0
    On Error Resume Next
    Set cells = Nothing
    Set cells = ws.Cells.SpecialCells(-4123)
    If Not cells Is Nothing Then fc = cells.Count
    On Error GoTo 0
    total = total + fc
  Next
  WScript.Echo "open " & FormatNumber(tOpen, 1) & " s, repaired=" & repaired & ", sheets=" & wb.Worksheets.Count & ", formula cells left=" & total & ", sample Sheet1!J5=" & wb.Worksheets(1).Range("J5").Text & " (" & TypeName(wb.Worksheets(1).Range("J5").Value) & ")"
  wb.Close False
Else
  t0 = Timer : wb.Save : tSave = Timer - t0
  t0 = Timer : wb.Close False : xl.Quit : Set wb = Nothing : Set xl = Nothing
  WScript.Echo "hidden Excel: start+open " & FormatNumber(tOpen, 1) & " s, save " & FormatNumber(tSave, 1) & " s (repaired=" & repaired & ")"
  WScript.Quit 0
End If
xl.Quit
