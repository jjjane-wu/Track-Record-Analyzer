Attribute VB_Name = "modMain"
' ===================================================================
'  Entry points + orchestration for the full 8-tab build.
'
'  Prerequisite: this workbook contains a "Deal Level Inputs" sheet in
'  the standard 28-column layout (from the Python app, pasted in, or
'  injected by build.vbs).
'
'  Interactive:  Alt+F8 -> BuildAnalysisWorkbook  (dialogs)
'  Scripted:     build.bat / build.vbs -> BuildHeadless (no dialogs;
'                any tab failure aborts with a nonzero exit)
'
'  Tab order produced: Table of Contents, Deal Level Inputs,
'  Deal List, Return & Loss Ratios, Return Dispersion, Portfolio
'  Construction, Vintage Perf by Sector, Deployment & Exits.
'  Each analysis tab builds inside its own error scope, so one bad
'  tab reports instead of killing the whole run.
' ===================================================================
Option Explicit

Public Sub BuildAnalysisWorkbook()
    Dim t0 As Double, problems As String
    t0 = Timer
    problems = BuildCore()
    If Len(problems) = 0 Then
        MsgBox "Built all 8 tabs (" & modBuild.DealCount() & " deals) in " & _
               Format(Timer - t0, "0.0") & "s.", vbInformation, "TR Analyzer (VBA)"
    Else
        MsgBox "Build finished with problems:" & vbCrLf & problems, _
               vbExclamation, "TR Analyzer (VBA)"
    End If
End Sub

Public Sub BuildHeadless()
    Dim problems As String
    problems = BuildCore()
    If Len(problems) > 0 Then
        Err.Raise vbObjectError + 99, , "Build problems: " & Replace(problems, vbCrLf, " | ")
    End If
End Sub

Private Function BuildCore() As String
    Dim problems As String
    Application.ScreenUpdating = False
    On Error GoTo hardFail

    If Not SheetExists("Deal Level Inputs") Then
        Err.Raise vbObjectError + 10, , _
            "Sheet 'Deal Level Inputs' not found - paste or import it first."
    End If

    modUtil.ResetCache
    modCharts.ResetStage

    ' Deal List is the foundation - if it fails, nothing else can build.
    modBuild.BuildDealList
    Application.Calculate

    RunTab problems, "Return & Loss Ratios", "modPivots.BuildReturnLossRatios"
    RunTab problems, "Return Dispersion", "modDispersion.BuildReturnDispersion"
    RunTab problems, "Portfolio Construction", "modConstruction.BuildPortfolioConstruction"
    RunTab problems, "Vintage Perf by Sector", "modVintage.BuildVintagePerf"
    RunTab problems, "Deployment & Exits", "modDeployment.BuildDeployment"
    RunTab problems, "Table of Contents", "modToc.BuildTOC"

    ArrangeSheets
    On Error Resume Next
    ThisWorkbook.Worksheets("_ChartData").Visible = xlSheetHidden
    On Error GoTo hardFail

    Application.ScreenUpdating = True
    BuildCore = problems
    Exit Function
hardFail:
    Dim n As Long, d As String
    n = Err.Number: d = Err.Description
    Application.ScreenUpdating = True
    Err.Raise n, , d
End Function

' Run one tab builder inside its own error scope.
Private Sub RunTab(ByRef problems As String, ByVal label As String, ByVal proc As String)
    On Error GoTo eh
    Application.Run proc
    Exit Sub
eh:
    problems = problems & "- " & label & ": " & Err.Description & vbCrLf
End Sub

Private Sub ArrangeSheets()
    Dim order As Variant, i As Long
    order = Array("Table of Contents", "Deal Level Inputs", "Deal List", _
                  "Return & Loss Ratios", "Return Dispersion", _
                  "Portfolio Construction", "Vintage Perf by Sector", _
                  "Deployment & Exits")
    On Error Resume Next
    For i = UBound(order) To LBound(order) Step -1
        ThisWorkbook.Worksheets(CStr(order(i))).Move Before:=ThisWorkbook.Worksheets(1)
    Next i
    On Error GoTo 0
End Sub

Private Function SheetExists(ByVal name As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(name)
    On Error GoTo 0
    SheetExists = Not ws Is Nothing
End Function
