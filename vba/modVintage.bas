Attribute VB_Name = "modVintage"
' ===================================================================
'  Vintage Perf by Sector -- the vintage performance pivot (Count /
'  Invested Capital / MOIC / Loss Ratio, four report filters, combo
'  chart) plus three Vintage x Sector matrices: deal count, pooled
'  MOIC (calculated field = correct pooled math per cell), and deal
'  count with a Fund filter. Mirrors plan_extra's VS sections.
' ===================================================================
Option Explicit

Private Const TITLE_COL As Long = 3
Private Const PIVOT_COL As Long = 3
Private Const CHART_COL As Long = 10
Private Const GAP As Long = 7

Private Function BuildVPivot(ws As Worksheet, ByVal anchor As Long) As Long
    On Error GoTo eh
    Dim pt As PivotTable
    Set pt = modUtil.Cache().CreatePivotTable( _
        TableDestination:=ws.Cells(anchor + GAP + 1, PIVOT_COL), TableName:="VS_MAIN")
    modUtil.AddCalcFields pt
    With pt
        .PivotFields("Vintage").Orientation = xlRowField
        ' four report filters (template: Fund / Sector / Status / HP buckets)
        .PivotFields("Fund").Orientation = xlPageField
        .PivotFields("Sector").Orientation = xlPageField
        .PivotFields("Status").Orientation = xlPageField
        .PivotFields("Hold Period Buckets").Orientation = xlPageField
        modUtil.AddData pt, "Company", "Count", xlCount, "0"
        modUtil.AddData pt, "Total Invested Capital (mlns)", "Invested Capital", xlSum, "#,##0"
        modUtil.AddData pt, "CalcMOIC", "MOIC", xlSum, "0.0\x;(0.0\x)"
        modUtil.AddData pt, "CalcLossRatio", "Loss Ratio", xlSum, "0.0%"
        .ColumnGrand = False
        .RowGrand = True
    End With

    ' combo chart: invested capital columns + MOIC line
    Dim cats As Collection, vIC() As Variant, vM() As Variant, i As Long
    Set cats = modCharts.AxisItems(pt, "Vintage")
    If cats.Count > 0 Then
        ReDim vIC(1 To cats.Count): ReDim vM(1 To cats.Count)
        For i = 1 To cats.Count
            vIC(i) = modCharts.PivotVal(pt, "Invested Capital", "Vintage", CStr(cats(i)))
            vM(i) = modCharts.PivotVal(pt, "MOIC", "Vintage", CStr(cats(i)))
        Next i
        modCharts.ComboChart ws, anchor + GAP + 1, CHART_COL, _
            "Invested Capital & MOIC by Vintage", cats, _
            vIC, "Invested Capital", "#,##0", vM, "MOIC", "0.0""x"""
    End If
    BuildVPivot = pt.TableRange2.Rows.Count + 1
    Exit Function
eh:
    BuildVPivot = 0
End Function

Private Function BuildMatrix(ws As Worksheet, ByVal anchor As Long, ByVal ptName As String, _
                             ByVal countValue As Boolean, ByVal fundPage As Boolean) As Long
    On Error GoTo eh
    Dim pt As PivotTable
    Set pt = modUtil.Cache().CreatePivotTable( _
        TableDestination:=ws.Cells(anchor + IIf(fundPage, GAP, 2), PIVOT_COL), _
        TableName:=ptName)
    modUtil.AddCalcFields pt
    With pt
        .PivotFields("Vintage").Orientation = xlRowField
        .PivotFields("Sector").Orientation = xlColumnField
        If fundPage Then .PivotFields("Fund").Orientation = xlPageField
        If countValue Then
            modUtil.AddData pt, "Company", "Count of Company", xlCount, "0"
        Else
            modUtil.AddData pt, "CalcMOIC", "Pooled MOIC", xlSum, "0.0\x;(0.0\x)"
        End If
        .ColumnGrand = True
        .RowGrand = True
    End With
    modUtil.HideBlank pt, "Sector"
    BuildMatrix = pt.TableRange2.Rows.Count + 1
    Exit Function
eh:
    BuildMatrix = 0
End Function

Public Sub BuildVintagePerf()
    Dim ws As Worksheet, anchor As Long, used As Long
    Dim titles(1 To 4) As String, anchors(1 To 4) As Long
    titles(1) = "Invested Capital & MOIC by Vintage"
    titles(2) = "Deal Count by Vintage and Sector"
    titles(3) = "Pooled MOIC by Vintage and Sector"
    titles(4) = "Deal Count by Vintage and Sector (filter by Fund)"

    Set ws = modBuild.FreshSheet("Vintage Perf by Sector")
    modUtil.MetaBlock ws, "Vintage Performance by Sector"

    anchor = 8 + 4 + 3
    anchors(1) = anchor
    modUtil.SectionTitle ws, anchor, TITLE_COL, titles(1)
    used = BuildVPivot(ws, anchor)
    If used = 0 Then ws.Cells(anchor + GAP, PIVOT_COL).Value = "(no data)": used = 2
    If used < 13 Then used = 13
    anchor = anchor + GAP + used + 6

    Dim i As Long
    For i = 2 To 4
        anchors(i) = anchor
        modUtil.SectionTitle ws, anchor, TITLE_COL, titles(i)
        used = BuildMatrix(ws, anchor, "VS_M" & i, (i <> 3), (i = 4))
        If used = 0 Then ws.Cells(anchor + 2, PIVOT_COL).Value = "(no data)": used = 2
        anchor = anchor + IIf(i = 4, GAP, 2) + used + 5
    Next i
    modUtil.MiniToc ws, titles, anchors, 4
End Sub
