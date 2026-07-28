Attribute VB_Name = "modUtil"
' ===================================================================
'  Shared plumbing for every tab builder: the one pivot cache, the
'  calculated fields, meta blocks, page filters, item ordering, blank
'  hiding, and in-sheet mini contents lists.
' ===================================================================
Option Explicit

Private mCache As PivotCache

' One shared cache over the DealLevelInput table (like the Python build).
Public Function Cache() As PivotCache
    If mCache Is Nothing Then
        Set mCache = ThisWorkbook.PivotCaches.Create( _
            SourceType:=xlDatabase, SourceData:="DealLevelInput")
    End If
    Set Cache = mCache
End Function

Public Sub ResetCache()
    Set mCache = Nothing
End Sub

Private Function NL() As String
    NL = Chr(10)
End Function

' Calculated fields, template semantics (same as _CALC_FIELDS in Python).
' Cache-scoped: the first pivot defines them, later Adds fail silently.
Public Sub AddCalcFields(pt As PivotTable)
    On Error Resume Next
    pt.CalculatedFields.Add "CalcMOIC", _
        "='Total" & NL() & "Value'/'Total Invested Capital (mlns)'", True
    pt.CalculatedFields.Add "CalcLossRatio", _
        "='InvCapital in Loss Position'/'Total Invested Capital (mlns)'", True
    pt.CalculatedFields.Add "CalcImpairedLossRatio", _
        "='Impaired" & NL() & "Value'/'Total Invested Capital (mlns)'", True
    On Error GoTo 0
End Sub

Public Sub AddData(pt As PivotTable, ByVal fieldName As String, _
                   ByVal caption As String, ByVal how As Long, _
                   ByVal numFmt As String)
    Dim df As PivotField
    Set df = pt.AddDataField(pt.PivotFields(fieldName), caption, how)
    If Len(numFmt) > 0 Then df.NumberFormat = numFmt
End Sub

' Standard meta block (title + GP / as-of / currency linked to Deal List).
Public Sub MetaBlock(ws As Worksheet, ByVal title As String)
    ws.Range("B2").Value = title
    ws.Range("B2").Font.Bold = True: ws.Range("B2").Font.Size = 16
    ws.Range("B4").Value = "Sponsor/GP:": ws.Range("C4").Formula = "='Deal List'!$C$4"
    ws.Range("B5").Value = "As of Date:": ws.Range("C5").Formula = "='Deal List'!$C$5"
    ws.Range("C5").NumberFormat = "d-mmm-yy"
    ws.Range("B6").Value = "Currency:":   ws.Range("C6").Formula = "='Deal List'!$C$6"
End Sub

Public Sub SectionTitle(ws As Worksheet, ByVal r As Long, ByVal c As Long, _
                        ByVal title As String)
    ws.Cells(r, c).Value = title
    ws.Cells(r, c).Font.Bold = True
    ws.Cells(r, c).Font.Color = RGB(31, 78, 120)
End Sub

' Report filters: every page header except the pivot's own axis field.
Public Sub AddPageFilters(pt As PivotTable, ByVal axisField As String)
    Dim pages() As String, k As Long
    modSpec.LoadPageFields pages
    For k = LBound(pages) To UBound(pages)
        If pages(k) <> axisField Then
            On Error Resume Next               ' field may equal a col axis
            pt.PivotFields(pages(k)).Orientation = xlPageField
            On Error GoTo 0
        End If
    Next k
End Sub

' Pre-select a single page item (e.g. Status = Realized). Silently keeps
' (All) when the item does not exist in this GP's data.
Public Sub SelectPage(pt As PivotTable, ByVal fieldName As String, _
                      ByVal item As String)
    On Error Resume Next
    pt.PivotFields(fieldName).Orientation = xlPageField
    pt.PivotFields(fieldName).CurrentPage = item
    On Error GoTo 0
End Sub

' Hide the "(blank)" item of an axis field (the Python build's rule:
' unlabelled deals are hidden from the pivot and its totals).
Public Sub HideBlank(pt As PivotTable, ByVal fieldName As String)
    ' Mac Excel's PivotItems enumeration SKIPS the blank item, so a For Each
    ' never sees it -- but indexing it by name works. Try both spellings.
    On Error Resume Next
    pt.PivotFields(fieldName).PivotItems("(blank)").Visible = False
    pt.PivotFields(fieldName).PivotItems("").Visible = False
    On Error GoTo 0
End Sub

' Bucket axes display in threshold order, not alphabetical order.
Public Sub ApplyCanonicalOrder(pt As PivotTable, ByVal fieldName As String)
    Dim order As Variant, k As Long, pos As Long
    order = modSpec.CanonicalOrder(fieldName)
    If IsEmpty(order) Then Exit Sub
    pos = 1
    On Error Resume Next                       ' labels absent in this data
    For k = LBound(order) To UBound(order)
        pt.PivotFields(fieldName).PivotItems(CStr(order(k))).Position = pos
        If Err.Number = 0 Then pos = pos + 1
        Err.Clear
    Next k
    On Error GoTo 0
End Sub

' In-sheet mini contents list (numbered internal links), rows startRow..
Public Sub MiniToc(ws As Worksheet, titles() As String, anchors() As Long, _
                   ByVal n As Long, Optional ByVal startRow As Long = 8)
    Dim i As Long, c As Range
    For i = 1 To n
        Set c = ws.Cells(startRow + i - 1, 2)
        c.Value = i
        c.Font.Bold = True
        c.Interior.Color = RGB(180, 198, 231)
        Set c = ws.Cells(startRow + i - 1, 3)
        ws.Hyperlinks.Add Anchor:=c, Address:="", _
            SubAddress:="'" & ws.Name & "'!B" & anchors(i), _
            TextToDisplay:=titles(i)
        c.Font.Size = 10
    Next i
End Sub
