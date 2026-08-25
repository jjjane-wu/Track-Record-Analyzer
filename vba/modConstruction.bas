Attribute VB_Name = "modConstruction"
' ===================================================================
'  Portfolio Construction -- two Fund x dimension capital-mix matrices
'  (Sum of Total Invested Capital, % of row, Status filter, stacked
'  chart) and the Deal Count Attributes section: five Count-of-Company
'  pivots (desc order, pie chart each) plus the bare total count.
'  Mirrors PC_MATRIX_SPECS / PC_COUNT_ORDER in build_output.py.
' ===================================================================
Option Explicit

Private Const TITLE_COL As Long = 3
Private Const PIVOT_COL As Long = 3
Private Const CHART_COL As Long = 10
Private Const GAP As Long = 7

Private Function BuildMatrix(ws As Worksheet, ByVal anchor As Long, _
                             ByVal title As String, ByVal colField As String, _
                             ByVal ptName As String) As Long
    On Error GoTo eh
    Dim pt As PivotTable, df As PivotField
    Set pt = modUtil.Cache().CreatePivotTable( _
        TableDestination:=ws.Cells(anchor + GAP, PIVOT_COL), TableName:=ptName)
    With pt
        .PivotFields("Fund").Orientation = xlRowField
        .PivotFields(colField).Orientation = xlColumnField
        .PivotFields("Status").Orientation = xlPageField
        Set df = .AddDataField(.PivotFields("Total Invested Capital (mlns)"), _
                               "% of Fund Capital", xlSum)
        df.Calculation = xlPercentOfRow
        df.NumberFormat = "0%;(0%);""-"""
        .ColumnGrand = True                     ' bottom Grand Total row
        .RowGrand = False                       ' no 100% right column
    End With
    modUtil.HideBlank pt, "Fund"
    modUtil.HideBlank pt, colField

    ' stacked chart: funds as categories, one series per dimension item
    Dim funds As Collection, dims As Collection, vals() As Variant, i As Long, j As Long
    Set funds = modCharts.AxisItems(pt, "Fund")
    Set dims = modCharts.AxisItems(pt, colField)
    If funds.Count > 0 And dims.Count > 0 Then
        ReDim vals(1 To dims.Count, 1 To funds.Count)
        For i = 1 To dims.Count
            For j = 1 To funds.Count
                vals(i, j) = modCharts.PivotVal(pt, "% of Fund Capital", _
                    "Fund", CStr(funds(j)), colField, CStr(dims(i)))
            Next j
        Next i
        modCharts.StackedChart ws, anchor + GAP, CHART_COL, title, funds, dims, vals
    End If
    BuildMatrix = pt.TableRange2.Rows.Count
    Exit Function
eh:
    BuildMatrix = 0
End Function

Private Function BuildCount(ws As Worksheet, ByVal anchor As Long, _
                            ByVal dimField As String, ByVal title As String, _
                            ByVal ptName As String) As Long
    On Error GoTo eh
    Dim pt As PivotTable
    Set pt = modUtil.Cache().CreatePivotTable( _
        TableDestination:=ws.Cells(anchor + GAP, PIVOT_COL), TableName:=ptName)
    With pt
        .PivotFields(dimField).Orientation = xlRowField
        modUtil.AddData pt, "Company", "Count of Company", xlCount, "0"
        .PivotFields(dimField).AutoSort xlDescending, "Count of Company"
        .ColumnGrand = True     ' bottom Grand Total row
        .RowGrand = False
    End With
    modUtil.HideBlank pt, dimField

    Dim cats As Collection, v() As Variant, i As Long
    Set cats = modCharts.AxisItems(pt, dimField)
    If cats.Count > 0 Then
        ReDim v(1 To cats.Count)
        For i = 1 To cats.Count
            v(i) = modCharts.PivotVal(pt, "Count of Company", dimField, CStr(cats(i)))
        Next i
        modCharts.PieChart ws, anchor + GAP, CHART_COL, title, cats, v
    End If
    BuildCount = pt.TableRange2.Rows.Count
    Exit Function
eh:
    BuildCount = 0
End Function

Public Sub BuildPortfolioConstruction()
    Dim ws As Worksheet, anchor As Long, used As Long, i As Long
    Dim titles(1 To 8) As String, anchors(1 To 8) As Long, n As Long

    Set ws = modBuild.FreshSheet("Portfolio Construction")
    modUtil.MetaBlock ws, "Portfolio Construction"

    anchor = 8 + 8 + 3

    ' -- capital-mix matrices (PC_MATRIX_SPECS) -----------------------
    Dim mdims(1 To 2) As String, mtitles(1 To 2) As String
    mdims(1) = "Sector": mtitles(1) = "Invested Capital by Fund and Sector"
    mdims(2) = "Geography": mtitles(2) = "Invested Capital by Fund and Geography"
    For i = 1 To 2
        n = n + 1: titles(n) = mtitles(i): anchors(n) = anchor
        modUtil.SectionTitle ws, anchor, TITLE_COL, mtitles(i)
        used = BuildMatrix(ws, anchor, mtitles(i), mdims(i), "PC_M" & i)
        If used = 0 Then ws.Cells(anchor + GAP, PIVOT_COL).Value = "(no data)": used = 2
        If used < 12 Then used = 12
        anchor = anchor + GAP + used + 6
    Next i

    ' -- deal-count attributes (PC_COUNT_ORDER; empty entry = total) --
    Dim cdims(1 To 6) As String
    cdims(1) = "Sector": cdims(2) = "Geography": cdims(3) = ""
    cdims(4) = "Transaction Type": cdims(5) = "GP Role": cdims(6) = "Process Type"
    For i = 1 To 6
        If Len(cdims(i)) = 0 Then
            n = n + 1: titles(n) = "Total Deal Count": anchors(n) = anchor
            modUtil.SectionTitle ws, anchor, TITLE_COL, "Total Deal Count"
            On Error Resume Next
            Dim pt As PivotTable
            Set pt = modUtil.Cache().CreatePivotTable( _
                TableDestination:=ws.Cells(anchor + 2, PIVOT_COL), TableName:="PC_TOTAL")
            modUtil.AddData pt, "Company", "Count of Company", xlCount, "0"
            pt.ColumnGrand = False: pt.RowGrand = False
            On Error GoTo 0
            anchor = anchor + 2 + 4 + 3
        Else
            n = n + 1: titles(n) = "Deal Count by " & cdims(i): anchors(n) = anchor
            modUtil.SectionTitle ws, anchor, TITLE_COL, "Deal Count by " & cdims(i)
            used = BuildCount(ws, anchor, cdims(i), "Deal Count by " & cdims(i), "PC_C" & i)
            If used = 0 Then ws.Cells(anchor + GAP, PIVOT_COL).Value = "(no data)": used = 2
            If used < 12 Then used = 12
            anchor = anchor + GAP + used + 6
        End If
    Next i
    modUtil.MiniToc ws, titles, anchors, n
End Sub
