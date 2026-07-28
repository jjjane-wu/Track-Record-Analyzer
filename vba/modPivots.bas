Attribute VB_Name = "modPivots"
' ===================================================================
'  Return & Loss Ratios -- 15 native pivots (PIVOT_SPECS order) with
'  Count / MOIC / Loss-Ratio calculated fields, report filters, bucket
'  ordering, hidden blanks, and a static combo chart per breakdown
'  (columns = MOIC, line = Loss Ratio; snapshot, never filter-linked).
' ===================================================================
Option Explicit

Private Const TITLE_COL As Long = 3
Private Const PIVOT_COL As Long = 3
Private Const CHART_COL As Long = 10
Private Const GAP As Long = 7

' Build one breakdown pivot; returns rows occupied (0 = no data / failed).
Private Function BuildOnePivot(ws As Worksheet, ByVal anchor As Long, ByVal idx As Long, _
                               ByVal axisField As String, ByVal variant_ As String) As Long
    On Error GoTo eh
    Dim pt As PivotTable
    Set pt = modUtil.Cache().CreatePivotTable( _
        TableDestination:=ws.Cells(anchor + GAP, PIVOT_COL), _
        TableName:="RLR_" & idx)
    modUtil.AddCalcFields pt

    With pt
        .PivotFields(axisField).Orientation = xlRowField
        modUtil.AddPageFilters pt, axisField
        If variant_ = "with_ic" Then
            modUtil.AddData pt, "Total Invested Capital (mlns)", _
                    "Total Invested Capital", xlSum, "#,##0"
        End If
        modUtil.AddData pt, "Company", "Count", xlCount, "0"
        modUtil.AddData pt, "CalcMOIC", "MOIC", xlSum, "0.0\x;(0.0\x)"
        If variant_ = "impaired" Then
            modUtil.AddData pt, "CalcImpairedLossRatio", "Impaired Loss Ratio", xlSum, "0.0%"
        Else
            modUtil.AddData pt, "CalcLossRatio", "Loss Ratio", xlSum, "0.0%"
        End If
        .ColumnGrand = False
        .RowGrand = True
    End With
    modUtil.HideBlank pt, axisField
    modUtil.ApplyCanonicalOrder pt, axisField

    ' static snapshot chart (full portfolio, before any user filtering)
    Dim cats As Collection, i As Long
    Dim vMoic() As Variant, vLoss() As Variant
    Set cats = modCharts.AxisItems(pt, axisField)
    If cats.Count > 0 Then
        ReDim vMoic(1 To cats.Count): ReDim vLoss(1 To cats.Count)
        For i = 1 To cats.Count
            vMoic(i) = modCharts.PivotVal(pt, "MOIC", axisField, CStr(cats(i)))
            If variant_ = "impaired" Then
                vLoss(i) = modCharts.PivotVal(pt, "Impaired Loss Ratio", axisField, CStr(cats(i)))
            Else
                vLoss(i) = modCharts.PivotVal(pt, "Loss Ratio", axisField, CStr(cats(i)))
            End If
        Next i
        modCharts.ComboChart ws, anchor + GAP, CHART_COL, ws.Cells(anchor, TITLE_COL).Value, _
            cats, vMoic, "MOIC", "0.0""x""", vLoss, "Loss Ratio", "0%"
    End If

    BuildOnePivot = pt.TableRange2.Rows.Count
    Exit Function
eh:
    BuildOnePivot = 0
End Function

Public Sub BuildReturnLossRatios()
    Dim ws As Worksheet
    Dim title() As String, fld() As String, variant_() As String
    Dim i As Long, anchor As Long, used As Long
    Dim anchors() As Long

    modSpec.LoadPivotSpecs title, fld, variant_
    ReDim anchors(1 To modSpec.RLR_NPIVOTS)

    Set ws = modBuild.FreshSheet("Return & Loss Ratios")
    modUtil.MetaBlock ws, "Gross Returns & Loss Ratios"

    ' first section sits below the mini contents list (rows 8..)
    anchor = 8 + modSpec.RLR_NPIVOTS + 3
    For i = 1 To modSpec.RLR_NPIVOTS
        anchors(i) = anchor
        modUtil.SectionTitle ws, anchor, TITLE_COL, title(i)
        used = BuildOnePivot(ws, anchor, i, fld(i), variant_(i))
        If used = 0 Then
            ws.Cells(anchor + GAP, PIVOT_COL).Value = "(no data for this breakdown)"
            used = 2
        End If
        If used < 12 Then used = 12            ' keep room for the chart
        anchor = anchor + GAP + used + 6
    Next i
    modUtil.MiniToc ws, title, anchors, modSpec.RLR_NPIVOTS
End Sub
