Attribute VB_Name = "modPivots"
' ===================================================================
'  Return & Loss Ratios -- NATIVE PivotTables over the DealLevelInput
'  table. This is the part where VBA beats the Python OOXML approach:
'  Excel builds and owns the pivot cache, so the whole corruption
'  class (removed views, cache repairs, 255-char items, refresh
'  drift) cannot occur.
'
'  Per pivot (same 15 cuts as PIVOT_SPECS in build_output.py):
'    rows  = the breakdown field
'    data  = Count of Company, MOIC, Loss Ratio        (base variant)
'            + Impaired Loss Ratio            (variant "impaired")
'            + Sum of Total Invested Capital  (variant "with_ic")
'    page  = Fund / Status / Hold Period Buckets (minus the axis field)
'  MOIC / loss ratios are pivot CALCULATED FIELDS, so they re-aggregate
'  correctly under any filter (pooled, not averaged).
' ===================================================================
Option Explicit

Private Const TITLE_COL As Long = 3           ' section titles in column C
Private Const PIVOT_COL As Long = 3           ' pivot top-left column
Private Const GAP As Long = 7                 ' rows between title and pivot

Private Function NL() As String
    NL = Chr(10)
End Function

' calculated-field formulas -- template semantics (same as _CALC_FIELDS)
Private Sub AddCalcFields(pt As PivotTable)
    On Error Resume Next                       ' already defined on this cache
    pt.CalculatedFields.Add "CalcMOIC", _
        "='Total" & NL() & "Value'/'Total Invested Capital (mlns)'", True
    pt.CalculatedFields.Add "CalcLossRatio", _
        "='InvCapital in Loss Position'/'Total Invested Capital (mlns)'", True
    pt.CalculatedFields.Add "CalcImpairedLossRatio", _
        "='Impaired" & NL() & "Value'/'Total Invested Capital (mlns)'", True
    On Error GoTo 0
End Sub

Private Sub AddData(pt As PivotTable, ByVal fieldName As String, _
                    ByVal caption As String, ByVal how As Long, _
                    ByVal numFmt As String)
    Dim df As PivotField
    Set df = pt.AddDataField(pt.PivotFields(fieldName), caption, how)
    If Len(numFmt) > 0 Then df.NumberFormat = numFmt
End Sub

' Build one breakdown pivot. Returns the rows it occupies (0 = failed,
' caller writes a "(no data)" note instead). Own error scope so one bad
' breakdown never derails the rest of the sheet.
Private Function BuildOnePivot(ws As Worksheet, pc As PivotCache, _
                               ByVal anchor As Long, ByVal idx As Long, _
                               ByVal axisField As String, ByVal variant_ As String, _
                               pages() As String) As Long
    On Error GoTo eh
    Dim pt As PivotTable, k As Long
    Set pt = pc.CreatePivotTable( _
        TableDestination:=ws.Cells(anchor + GAP, PIVOT_COL), _
        TableName:="RLR_" & idx)
    AddCalcFields pt

    With pt
        .PivotFields(axisField).Orientation = xlRowField
        For k = LBound(pages) To UBound(pages)
            If pages(k) <> axisField Then
                .PivotFields(pages(k)).Orientation = xlPageField
            End If
        Next k
        If variant_ = "with_ic" Then
            AddData pt, "Total Invested Capital (mlns)", _
                    "Total Invested Capital", xlSum, "#,##0"
        End If
        AddData pt, "Company", "Count", xlCount, "0"
        AddData pt, "CalcMOIC", "MOIC", xlSum, "0.0\x;(0.0\x)"
        If variant_ = "impaired" Then
            AddData pt, "CalcImpairedLossRatio", "Impaired Loss Ratio", xlSum, "0.0%"
        Else
            AddData pt, "CalcLossRatio", "Loss Ratio", xlSum, "0.0%"
        End If
        .ColumnGrand = False
        .RowGrand = True
    End With
    BuildOnePivot = pt.TableRange2.Rows.Count
    Exit Function
eh:
    BuildOnePivot = 0
End Function

Public Sub BuildReturnLossRatios()
    Dim ws As Worksheet, pc As PivotCache
    Dim title() As String, fld() As String, variant_() As String
    Dim pages() As String
    Dim i As Long, anchor As Long, used As Long

    modSpec.LoadPivotSpecs title, fld, variant_
    modSpec.LoadPageFields pages

    Set ws = modBuild.FreshSheet("Return & Loss Ratios")
    ws.Range("B2").Value = "Gross Returns & Loss Ratios"
    ws.Range("B2").Font.Bold = True: ws.Range("B2").Font.Size = 16
    ws.Range("B4").Value = "Sponsor/GP:": ws.Range("C4").Formula = "='Deal List'!$C$4"
    ws.Range("B5").Value = "As of Date:": ws.Range("C5").Formula = "='Deal List'!$C$5"
    ws.Range("C5").NumberFormat = "d-mmm-yy"
    ws.Range("B6").Value = "Currency:":   ws.Range("C6").Formula = "='Deal List'!$C$6"

    Set pc = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, SourceData:="DealLevelInput")

    anchor = 9
    For i = 1 To modSpec.RLR_NPIVOTS
        ws.Cells(anchor, TITLE_COL).Value = title(i)
        ws.Cells(anchor, TITLE_COL).Font.Bold = True
        ws.Cells(anchor, TITLE_COL).Font.Color = RGB(31, 78, 120)

        used = BuildOnePivot(ws, pc, anchor, i, fld(i), variant_(i), pages)
        If used = 0 Then
            ws.Cells(anchor + GAP, PIVOT_COL).Value = "(no data for this breakdown)"
            used = 2
        End If
        anchor = anchor + GAP + used + 6
    Next i
End Sub
