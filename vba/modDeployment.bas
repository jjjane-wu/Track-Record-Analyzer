Attribute VB_Name = "modDeployment"
' ===================================================================
'  Deployment & Exits -- template-verbatim four sections:
'    1. InvCap %       vintage rows x fund cols, % of column total
'    2. Deal Count     vintage rows x fund cols, with grand column
'    3. Exits % of IC  fund rows x exit-year cols, % of row,
'                      Status pre-selected to Realized
'    4. Exits by Year  fund rows x exit-year cols, deal count,
'                      Status pre-selected to Realized
'  Native pivots handle the empty-dimension cases (no realized exits)
'  gracefully -- the pivot just shows no rows.
' ===================================================================
Option Explicit

Private Const TITLE_COL As Long = 2
Private Const PIVOT_COL As Long = 3
Private Const GAP As Long = 2

Private Function BuildDE(ws As Worksheet, ByVal anchor As Long, ByVal ptName As String, _
                         ByVal rowField As String, ByVal colField As String, _
                         ByVal dataField As String, ByVal caption As String, _
                         ByVal how As Long, ByVal calc As Long, ByVal fmt As String, _
                         ByVal grandCol As Boolean, ByVal realizedOnly As Boolean) As Long
    On Error GoTo eh
    Dim pt As PivotTable, df As PivotField, off As Long
    off = IIf(realizedOnly, 4, GAP)            ' room for the Status filter row
    Set pt = modUtil.Cache().CreatePivotTable( _
        TableDestination:=ws.Cells(anchor + off, PIVOT_COL), TableName:=ptName)
    With pt
        .PivotFields(rowField).Orientation = xlRowField
        .PivotFields(colField).Orientation = xlColumnField
        Set df = .AddDataField(.PivotFields(dataField), caption, how)
        If calc <> 0 Then df.Calculation = calc
        If Len(fmt) > 0 Then df.NumberFormat = fmt
        .ColumnGrand = True     ' bottom Grand Total row
        .RowGrand = grandCol    ' right-hand grand column where wanted
    End With
    If realizedOnly Then modUtil.SelectPage pt, "Status", "Realized"
    modUtil.HideBlank pt, rowField
    modUtil.HideBlank pt, colField
    ' the "n/a" exit-year column is hidden like the Python build
    If colField = "Exit Year" Then
        On Error Resume Next
        pt.PivotFields("Exit Year").PivotItems("n/a").Visible = False
        On Error GoTo 0
    End If
    BuildDE = pt.TableRange2.Rows.Count + IIf(realizedOnly, 2, 0)
    Exit Function
eh:
    BuildDE = 0
End Function

Public Sub BuildDeployment()
    Dim ws As Worksheet, anchor As Long, used As Long
    Dim titles(1 To 4) As String, anchors(1 To 4) As Long

    titles(1) = "InvCap %"
    titles(2) = "Deal Count"
    titles(3) = "Exits % of IC by Fund"
    titles(4) = "Exits by Year"

    Set ws = modBuild.FreshSheet("Deployment & Exits")
    modUtil.MetaBlock ws, "Deployment/Pacing"

    anchor = 8 + 4 + 2
    anchors(1) = anchor
    modUtil.SectionTitle ws, anchor, TITLE_COL, titles(1)
    used = BuildDE(ws, anchor, "DE_PCT", "Vintage", "Fund", _
                   "Total Invested Capital (mlns)", "Sum of Total Invested Capital (mlns)", _
                   xlSum, xlPercentOfColumn, "0%;-0%;", False, False)
    If used = 0 Then ws.Cells(anchor + GAP, PIVOT_COL).Value = "(no data)": used = 3
    anchor = anchor + GAP + used + 4

    anchors(2) = anchor
    modUtil.SectionTitle ws, anchor, TITLE_COL, titles(2)
    used = BuildDE(ws, anchor, "DE_CNT", "Vintage", "Fund", _
                   "Company", "Count of Company", xlCount, 0, "0", True, False)
    If used = 0 Then ws.Cells(anchor + GAP, PIVOT_COL).Value = "(no data)": used = 3
    anchor = anchor + GAP + used + 4

    anchors(3) = anchor
    modUtil.SectionTitle ws, anchor, TITLE_COL, titles(3)
    used = BuildDE(ws, anchor, "DE_XPCT", "Fund", "Exit Year", _
                   "Total IC mlns for Buckets", "Sum of Total IC mlns for Buckets", _
                   xlSum, xlPercentOfRow, "0%;-0%;", False, True)
    If used = 0 Then ws.Cells(anchor + 4, PIVOT_COL).Value = "(no realized exits in the data)": used = 5
    anchor = anchor + 4 + used + 4

    anchors(4) = anchor
    modUtil.SectionTitle ws, anchor, TITLE_COL, titles(4)
    used = BuildDE(ws, anchor, "DE_XCNT", "Fund", "Exit Year", _
                   "Company", "Count of Company", xlCount, 0, "0", False, True)
    If used = 0 Then ws.Cells(anchor + 4, PIVOT_COL).Value = "(no realized exits in the data)"

    modUtil.MiniToc ws, titles, anchors, 4
End Sub
