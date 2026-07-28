' ===================================================================
'  Headless build driver (Windows).
'
'  Usage:
'    cscript //nologo build.vbs <TR-Analyzer.xlsm> <inputs.xlsx> <output.xlsx>
'
'    <TR-Analyzer.xlsm>  the macro template (modules from this folder)
'    <inputs.xlsx>       any workbook containing a "Deal Level Inputs"
'                        sheet (e.g. an output of the Python app)
'    <output.xlsx>       where to save the finished analysis (plain
'                        .xlsx - macros are stripped from the result)
'
'  Excel runs invisibly. A nonzero exit code means failure.
'  ONE-TIME SETUP: add the template's folder to Excel's Trusted
'  Locations (File > Options > Trust Center) or macros will not run
'  when opened by automation.
' ===================================================================
Option Explicit

Const SHEET_NAME = "Deal Level Inputs"
Const XL_XLSX = 51                       ' xlOpenXMLWorkbook (no macros)

Dim args: Set args = WScript.Arguments
If args.Count < 3 Then
    WScript.Echo "Usage: cscript //nologo build.vbs <TR-Analyzer.xlsm> <inputs.xlsx> <output.xlsx>"
    WScript.Quit 2
End If

Dim fso: Set fso = CreateObject("Scripting.FileSystemObject")
Dim tplPath, inPath, outPath
tplPath = fso.GetAbsolutePathName(args(0))
inPath = fso.GetAbsolutePathName(args(1))
outPath = fso.GetAbsolutePathName(args(2))
If Not fso.FileExists(tplPath) Then Fail Nothing, "template not found: " & tplPath
If Not fso.FileExists(inPath) Then Fail Nothing, "inputs workbook not found: " & inPath

Dim xl: Set xl = CreateObject("Excel.Application")
xl.Visible = False
xl.DisplayAlerts = False

On Error Resume Next

Dim wbT: Set wbT = xl.Workbooks.Open(tplPath)
If Err.Number <> 0 Then Fail xl, "could not open template: " & Err.Description

Dim wbI: Set wbI = xl.Workbooks.Open(inPath, True, True)   ' update-links no, read-only
If Err.Number <> 0 Then Fail xl, "could not open inputs: " & Err.Description

' replace the template's Deal Level Inputs with the fresh one
Dim ws: Set ws = Nothing
Set ws = wbT.Worksheets(SHEET_NAME)
If Not ws Is Nothing Then ws.Delete
Err.Clear
wbI.Worksheets(SHEET_NAME).Copy , wbT.Worksheets(wbT.Worksheets.Count)
If Err.Number <> 0 Then Fail xl, "could not copy '" & SHEET_NAME & "' from inputs: " & Err.Description
wbI.Close False

' run the macro (errors inside VBA surface here)
xl.Run "'" & wbT.Name & "'!modMain.BuildHeadless"
If Err.Number <> 0 Then Fail xl, "macro failed: " & Err.Description & _
    " (if this says macros are disabled, add the template folder to Trusted Locations)"

wbT.SaveAs outPath, XL_XLSX
If Err.Number <> 0 Then Fail xl, "could not save output: " & Err.Description
wbT.Close False
xl.Quit

WScript.Echo "OK: " & outPath
WScript.Quit 0

Sub Fail(app, msg)
    WScript.Echo "BUILD FAILED: " & msg
    If Not (app Is Nothing) Then
        app.DisplayAlerts = False
        app.Quit
    End If
    WScript.Quit 1
End Sub
