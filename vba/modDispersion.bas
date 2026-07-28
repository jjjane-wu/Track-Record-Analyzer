Attribute VB_Name = "modDispersion"
' ===================================================================
'  Return Dispersion -- Gross MOIC and Gross IRR bucket pivots, each
'  with Count / % IC (percent of column) / actual average, plus a
'  static column chart of the % IC distribution (n/a bucket excluded
'  from the chart, kept in the pivot -- same rules as Python).
' ===================================================================
Option Explicit

Private Const TITLE_COL As Long = 3
Private Const PIVOT_COL As Long = 3
Private Const CHART_COL As Long = 9
Private Const GAP As Long = 7

Private Function BuildSection(ws As Worksheet, ByVal anchor As Long, _
                              ByVal title As String, ByVal bucketField As String, _
                              ByVal metricField As String, ByVal avgCaption As String, _
                              ByVal avgFmt As String, ByVal ptName As String) As Long
    On Error GoTo eh
    Dim pt As PivotTable, df As PivotField
    Set pt = modUtil.Cache().CreatePivotTable( _
        TableDestination:=ws.Cells(anchor + GAP, PIVOT_COL), TableName:=ptName)
    With pt
        .PivotFields(bucketField).Orientation = xlRowField
        modUtil.AddPageFilters pt, bucketField
        modUtil.AddData pt, "Company", "Count", xlCount, "0"
        Set df = .AddDataField(.PivotFields("Total Invested Capital (mlns)"), "% IC", xlSum)
        df.Calculation = xlPercentOfColumn
        df.NumberFormat = "0%"
        modUtil.AddData pt, metricField, avgCaption, xlAverage, avgFmt
        .ColumnGrand = False
        .RowGrand = True
    End With
    modUtil.HideBlank pt, bucketField
    modUtil.ApplyCanonicalOrder pt, bucketField

    ' chart: % IC per bucket, n/a excluded
    Dim cats As New Collection, allCats As Collection, v() As Variant, i As Long, n As Long
    Set allCats = modCharts.AxisItems(pt, bucketField)
    For i = 1 To allCats.Count
        If allCats(i) <> "n/a" Then cats.Add allCats(i)
    Next i
    If cats.Count > 0 Then
        ReDim v(1 To cats.Count)
        For i = 1 To cats.Count
            v(i) = modCharts.PivotVal(pt, "% IC", bucketField, CStr(cats(i)))
        Next i
        modCharts.ColumnChart ws, anchor + GAP, CHART_COL, title & " - % of Invested Capital", _
            cats, v, "% IC", "0%"
    End If
    BuildSection = pt.TableRange2.Rows.Count
    Exit Function
eh:
    BuildSection = 0
End Function

Public Sub BuildReturnDispersion()
    Dim ws As Worksheet, anchor As Long, used As Long
    Dim titles(1 To 2) As String, anchors(1 To 2) As Long
    titles(1) = "Gross MOIC": titles(2) = "Gross IRR"

    Set ws = modBuild.FreshSheet("Return Dispersion")
    modUtil.MetaBlock ws, "Return Dispersion"

    anchor = 8 + 2 + 3
    anchors(1) = anchor
    modUtil.SectionTitle ws, anchor, TITLE_COL, titles(1)
    used = BuildSection(ws, anchor, titles(1), "MOIC Buckets", _
                        "Gross" & Chr(10) & "MOIC", "Avg Gross MOIC", "0.0\x", "RD_MOIC")
    If used = 0 Then ws.Cells(anchor + GAP, PIVOT_COL).Value = "(no data)": used = 2
    If used < 12 Then used = 12
    anchor = anchor + GAP + used + 6

    anchors(2) = anchor
    modUtil.SectionTitle ws, anchor, TITLE_COL, titles(2)
    used = BuildSection(ws, anchor, titles(2), "IRR Buckets", _
                        "Gross" & Chr(10) & "IRR", "Avg Gross IRR", "0.0%", "RD_IRR")
    If used = 0 Then ws.Cells(anchor + GAP, PIVOT_COL).Value = "(no data)"

    modUtil.MiniToc ws, titles, anchors, 2
End Sub
