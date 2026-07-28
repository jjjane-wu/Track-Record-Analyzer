Attribute VB_Name = "modCharts"
' ===================================================================
'  Static snapshot charts.
'
'  Rule carried over from the Python build: charts NEVER change when a
'  pivot is filtered. Native PivotCharts would, so instead each chart's
'  data is snapshotted right after its pivot is built (still unfiltered
'  = full portfolio) into a hidden "_ChartData" sheet via GetPivotData,
'  and a plain chart is drawn over that staging block.
' ===================================================================
Option Explicit

Private Const STAGE_NAME As String = "_ChartData"

Public Function StageSheet() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(STAGE_NAME)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add( _
            After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = STAGE_NAME
    End If
    Set StageSheet = ws
End Function

Public Sub ResetStage()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(STAGE_NAME)
    On Error GoTo 0
    If Not ws Is Nothing Then ws.Cells.Clear
End Sub

Private Function NextStageRow(ws As Worksheet) As Long
    Dim r As Long
    r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If r = 1 And Len(CStr(ws.Cells(1, 1).Value)) = 0 Then
        NextStageRow = 1
    Else
        NextStageRow = r + 2
    End If
End Function

' Visible, non-blank axis items of a built (still unfiltered) pivot.
Public Function AxisItems(pt As PivotTable, ByVal fieldName As String) As Collection
    Dim out As New Collection, pi As PivotItem
    For Each pi In pt.PivotFields(fieldName).PivotItems
        If pi.Visible And pi.Name <> "(blank)" Then out.Add pi.Name
    Next pi
    Set AxisItems = out
End Function

' One data value from a pivot, by captions; Empty when the cell is empty.
Public Function PivotVal(pt As PivotTable, ByVal dataCaption As String, _
                         ByVal fld As String, ByVal item As String, _
                         Optional ByVal fld2 As String = "", _
                         Optional ByVal item2 As String = "") As Variant
    On Error Resume Next
    If Len(fld2) = 0 Then
        PivotVal = pt.GetPivotData(dataCaption, fld, item).Value
    Else
        PivotVal = pt.GetPivotData(dataCaption, fld, item, fld2, item2).Value
    End If
    On Error GoTo 0
End Function

' Snapshot cats + up to two series into the staging sheet.
' Returns the first staged row; cats in col A, series in cols B / C.
Public Function Stage(cats As Collection, v1 As Variant, v2 As Variant) As Long
    Dim ws As Worksheet, r0 As Long, i As Long
    Set ws = StageSheet()
    r0 = NextStageRow(ws)
    For i = 1 To cats.Count
        ws.Cells(r0 + i - 1, 1).Value = cats(i)
        ws.Cells(r0 + i - 1, 2).Value = v1(i)
        If Not IsEmpty(v2) Then ws.Cells(r0 + i - 1, 3).Value = v2(i)
    Next i
    Stage = r0
End Function

Private Function StageRange(ByVal col As Long, ByVal r0 As Long, ByVal n As Long) As Range
    Dim ws As Worksheet
    Set ws = StageSheet()
    Set StageRange = ws.Range(ws.Cells(r0, col), ws.Cells(r0 + n - 1, col))
End Function

Private Function NewChart(ws As Worksheet, ByVal anchorRow As Long, _
                          ByVal anchorCol As Long, ByVal title As String) As Chart
    Dim co As ChartObject
    Set co = ws.ChartObjects.Add( _
        Left:=ws.Cells(anchorRow, anchorCol).Left, _
        Top:=ws.Cells(anchorRow, anchorCol).Top, _
        Width:=430, Height:=230)
    co.Chart.HasTitle = True
    co.Chart.ChartTitle.Text = title
    co.Chart.ChartTitle.Font.Size = 11
    Set NewChart = co.Chart
End Function

' Combo: columns (s1, primary axis) + line with markers (s2, secondary).
Public Sub ComboChart(ws As Worksheet, ByVal anchorRow As Long, ByVal anchorCol As Long, _
                      ByVal title As String, cats As Collection, _
                      v1 As Variant, ByVal n1 As String, ByVal f1 As String, _
                      v2 As Variant, ByVal n2 As String, ByVal f2 As String)
    Dim ch As Chart, s As Series, r0 As Long
    r0 = Stage(cats, v1, v2)
    Set ch = NewChart(ws, anchorRow, anchorCol, title)
    ch.ChartType = xlColumnClustered
    Set s = ch.SeriesCollection.NewSeries
    s.XValues = StageRange(1, r0, cats.Count)
    s.Values = StageRange(2, r0, cats.Count)
    s.Name = n1
    Set s = ch.SeriesCollection.NewSeries
    s.XValues = StageRange(1, r0, cats.Count)
    s.Values = StageRange(3, r0, cats.Count)
    s.Name = n2
    s.ChartType = xlLineMarkers
    s.AxisGroup = xlSecondary
    ch.HasLegend = False
    On Error Resume Next
    ch.Axes(xlValue, xlPrimary).TickLabels.NumberFormat = f1
    ch.Axes(xlValue, xlSecondary).TickLabels.NumberFormat = f2
    On Error GoTo 0
End Sub

' Simple column chart of one series.
Public Sub ColumnChart(ws As Worksheet, ByVal anchorRow As Long, ByVal anchorCol As Long, _
                       ByVal title As String, cats As Collection, _
                       v1 As Variant, ByVal n1 As String, ByVal f1 As String)
    Dim ch As Chart, s As Series, r0 As Long
    r0 = Stage(cats, v1, Empty)
    Set ch = NewChart(ws, anchorRow, anchorCol, title)
    ch.ChartType = xlColumnClustered
    Set s = ch.SeriesCollection.NewSeries
    s.XValues = StageRange(1, r0, cats.Count)
    s.Values = StageRange(2, r0, cats.Count)
    s.Name = n1
    ch.HasLegend = False
    On Error Resume Next
    ch.Axes(xlValue).TickLabels.NumberFormat = f1
    On Error GoTo 0
End Sub

' Pie chart of one series.
Public Sub PieChart(ws As Worksheet, ByVal anchorRow As Long, ByVal anchorCol As Long, _
                    ByVal title As String, cats As Collection, v1 As Variant)
    Dim ch As Chart, s As Series, r0 As Long
    r0 = Stage(cats, v1, Empty)
    Set ch = NewChart(ws, anchorRow, anchorCol, title)
    ch.ChartType = xlPie
    Set s = ch.SeriesCollection.NewSeries
    s.XValues = StageRange(1, r0, cats.Count)
    s.Values = StageRange(2, r0, cats.Count)
    ch.HasLegend = True
    ch.Legend.Position = xlLegendPositionRight
End Sub

' 100%-stacked columns: cats = funds, one series per dimension item.
' vals is a 2-D Variant (1..nSeries, 1..nCats).
Public Sub StackedChart(ws As Worksheet, ByVal anchorRow As Long, ByVal anchorCol As Long, _
                        ByVal title As String, cats As Collection, _
                        seriesNames As Collection, vals As Variant)
    Dim ch As Chart, s As Series, i As Long, j As Long
    Dim stg As Worksheet, r0 As Long
    Set stg = StageSheet()
    r0 = NextStageRow(stg)
    For j = 1 To cats.Count
        stg.Cells(r0 + j - 1, 1).Value = cats(j)
        For i = 1 To seriesNames.Count
            stg.Cells(r0 + j - 1, 1 + i).Value = vals(i, j)
        Next i
    Next j
    Set ch = NewChart(ws, anchorRow, anchorCol, title)
    ch.ChartType = xlColumnStacked100
    For i = 1 To seriesNames.Count
        Set s = ch.SeriesCollection.NewSeries
        s.XValues = StageRange(1, r0, cats.Count)
        s.Values = StageRange(1 + i, r0, cats.Count)
        s.Name = seriesNames(i)
    Next i
    ch.HasLegend = True
    ch.Legend.Position = xlLegendPositionBottom
End Sub
