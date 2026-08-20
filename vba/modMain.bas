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

Private mDealCount As Long        ' captured before the inputs sheet is removed

Public Sub BuildAnalysisWorkbook()
    Dim t0 As Double, problems As String
    t0 = Timer
    problems = BuildCore()
    If Len(problems) = 0 Then
        MsgBox "Built all 7 tabs (" & mDealCount & " deals) in " & _
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
    mDealCount = modBuild.DealCount()
    modBuild.BuildDealList
    Application.Calculate

    RunTab problems, "Return & Loss Ratios", "modPivots.BuildReturnLossRatios"
    RunTab problems, "Return Dispersion", "modDispersion.BuildReturnDispersion"
    RunTab problems, "Portfolio Construction", "modConstruction.BuildPortfolioConstruction"
    RunTab problems, "Vintage Perf by Sector", "modVintage.BuildVintagePerf"
    RunTab problems, "Deployment & Exits", "modDeployment.BuildDeployment"

    ' Transient sheets go BEFORE the TOC builds, so the contents list shows
    ' exactly the finished tabs (no dead rows): the inputs sheet is a
    ' consumed import vehicle, the Start instructions page ships only in
    ' the pristine workbook, and _ChartData is hidden staging.
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Worksheets("Deal Level Inputs").Delete
    ThisWorkbook.Worksheets("Start").Delete
    Application.DisplayAlerts = True
    ThisWorkbook.Worksheets("_ChartData").Visible = xlSheetHidden
    On Error GoTo hardFail

    RunTab problems, "Table of Contents", "modToc.BuildTOC"

    ArrangeSheets

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
    order = Array("Table of Contents", "Deal List", _
                  "Return & Loss Ratios", "Return Dispersion", _
                  "Portfolio Construction", "Vintage Perf by Sector", _
                  "Deployment & Exits")
    On Error Resume Next
    For i = UBound(order) To LBound(order) Step -1
        ThisWorkbook.Worksheets(CStr(order(i))).Move Before:=ThisWorkbook.Worksheets(1)
    Next i
    ThisWorkbook.Worksheets("_ChartData").Move _
        After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)
    On Error GoTo 0
End Sub

' Replace the Deal Level Inputs sheet from another workbook. With no
' argument, prompts for the file (the Streamlit app's "Gross Deal Level
' Input" download). Rename-copy-delete order so a same-name sheet can
' never collide; the fresh sheet keeps the old one's position.
Public Sub ReplaceInputs(Optional ByVal p As String = "")
    If Len(p) = 0 Then
        Dim f As Variant
        f = Application.GetOpenFilename()
        If VarType(f) = vbBoolean Then Exit Sub    ' cancelled
        p = CStr(f)
    End If
    Dim wbI As Workbook, old As Worksheet, pos As Long
    On Error Resume Next
    GrantAccessToMultipleFiles Array(p)
    On Error GoTo 0
    Set wbI = Workbooks.Open(p, 0, True)
    Set old = Nothing
    pos = 0
    On Error Resume Next
    Set old = ThisWorkbook.Worksheets("Deal Level Inputs")
    pos = old.Index
    On Error GoTo 0
    If Not old Is Nothing Then old.Name = "__old_inputs__"
    wbI.Worksheets("Deal Level Inputs").Copy _
        After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)
    wbI.Close False
    If Not old Is Nothing Then
        Application.DisplayAlerts = False
        old.Delete
        Application.DisplayAlerts = True
    End If
    If pos > 0 And pos <= ThisWorkbook.Worksheets.Count Then
        ThisWorkbook.Worksheets("Deal Level Inputs").Move _
            Before:=ThisWorkbook.Worksheets(pos)
    End If
End Sub

' One click: pick the inputs file, then build all 8 tabs.
Public Sub ImportInputsAndBuild()
    ReplaceInputs
    If SheetExists("Deal Level Inputs") Then BuildAnalysisWorkbook
End Sub

Private Function SheetExists(ByVal name As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(name)
    On Error GoTo 0
    SheetExists = Not ws Is Nothing
End Function
