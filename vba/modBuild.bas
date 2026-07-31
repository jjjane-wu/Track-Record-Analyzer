Attribute VB_Name = "modBuild"
' ===================================================================
'  Deal List builder -- native VBA counterpart of build_output.py's
'  _write_deal_list. Reads the "Deal Level Inputs" sheet (produced by
'  the Python parser or pasted in) and rebuilds the "Deal List" sheet:
'  header block, tag row, 74-column table `DealLevelInput` with the
'  template formulas from modSpec (same in:/in0:/F:/FT: notation).
' ===================================================================
Option Explicit

Public Const DL_HDR_ROW As Long = 13          ' Deal List header row
Public Const DL_DATA_ROW As Long = 14         ' first data row
Public Const IN_DATA_ROW As Long = 7          ' first input data row (headers row 6)
Public Const FIRST_COL As Long = 2            ' both tables start at column B

' Number of deal rows on the inputs sheet. A row counts while ANY of the
' 28 input columns holds a value -- a deal with a blank Company (possible
' after imperfect parses) must not truncate the rows that follow it.
Public Function DealCount() As Long
    Dim ws As Worksheet, r As Long, c As Long, has As Boolean
    Set ws = ThisWorkbook.Worksheets("Deal Level Inputs")
    r = IN_DATA_ROW
    Do
        has = False
        For c = FIRST_COL To FIRST_COL + modSpec.IN_NCOLS - 1
            If Len(Trim$(CStr(ws.Cells(r, c).Value))) > 0 Then
                has = True
                Exit For
            End If
        Next c
        If Not has Then Exit Do
        r = r + 1
    Loop
    DealCount = r - IN_DATA_ROW
End Function

' 1-based position of an input-tab header (for copying its values).
Private Function InputColIndex(ByVal header As String) As Long
    Static loaded As Boolean
    Static hdrs() As String
    Dim i As Long
    If Not loaded Then
        modSpec.LoadInputCols hdrs
        loaded = True
    End If
    For i = 1 To modSpec.IN_NCOLS
        If hdrs(i) = header Then
            InputColIndex = i
            Exit Function
        End If
    Next i
    Err.Raise vbObjectError + 1, , "Input column not found: " & header
End Function

' same prefix semantics as Python _dl_src_formula
' Formula specs only; in:/in0: columns are copied as VALUES (the final
' workbook carries no Deal Level Inputs sheet to link to).
Private Function ResolveSpec(ByVal spec As String, ByVal r As Long) As String
    If Left$(spec, 3) = "FT:" Then
        ResolveSpec = "=" & Mid$(spec, 4)
    ElseIf Left$(spec, 2) = "F:" Then
        ResolveSpec = "=" & Replace(Mid$(spec, 3), "{r}", CStr(r))
    Else
        ResolveSpec = ""
    End If
End Function

Public Function FreshSheet(ByVal name As String) As Worksheet
    Dim ws As Worksheet, pos As Long
    pos = 0
    On Error Resume Next
    pos = ThisWorkbook.Worksheets(name).Index
    On Error GoTo 0
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Worksheets(name).Delete
    On Error GoTo 0
    Application.DisplayAlerts = True
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    ws.name = name
    If pos > 0 And pos <= ThisWorkbook.Worksheets.Count Then
        ws.Move Before:=ThisWorkbook.Worksheets(pos)   ' keep original position
    End If
    Set FreshSheet = ws
End Function

Public Sub BuildDealList()
    Dim ws As Worksheet, n As Long, i As Long, r As Long
    Dim h() As String, f() As String, fmt() As String, tag() As String
    Dim refs() As String, vals() As String, isNum() As Boolean, numv() As Double

    n = DealCount()
    If n = 0 Then Err.Raise vbObjectError + 2, , "No deals found on 'Deal Level Inputs'"
    modSpec.LoadDealListSpec h, f, fmt, tag

    Set ws = FreshSheet("Deal List")
    Dim lastRow As Long
    lastRow = DL_DATA_ROW + n - 1

    ' -- read the imported inputs ONCE (sheet is deleted after the build) --
    Dim inWs As Worksheet, data As Variant
    Set inWs = ThisWorkbook.Worksheets("Deal Level Inputs")
    data = inWs.Range(inWs.Cells(IN_DATA_ROW, FIRST_COL), _
                      inWs.Cells(IN_DATA_ROW + n - 1, FIRST_COL + modSpec.IN_NCOLS - 1)).Value

    ' -- meta block: VALUES (no cross-sheet links anywhere) ---------------
    ws.Range("B2").Value = "Deal List": ws.Range("B2").Font.Bold = True
    ws.Range("B2").Font.Size = 16
    ws.Range("B4").Value = "Sponsor/GP:":  ws.Range("C4").Value = inWs.Range("C3").Value
    ws.Range("B5").Value = "As of Date:":  ws.Range("C5").Value = inWs.Range("C4").Value
    ws.Range("C5").NumberFormat = "d-mmm-yy"
    ws.Range("B6").Value = "Currency:":    ws.Range("C6").Value = inWs.Range("C5").Value

    ' -- tag row 12 + header row 13 ---------------------------------------
    For i = 1 To modSpec.DL_NCOLS
        If Len(tag(i)) > 0 Then ws.Cells(DL_HDR_ROW - 1, FIRST_COL + i - 1).Value = tag(i)
        ws.Cells(DL_HDR_ROW, FIRST_COL + i - 1).Value = h(i)
        ws.Cells(DL_HDR_ROW, FIRST_COL + i - 1).Font.Bold = True
    Next i
    ws.Rows(DL_HDR_ROW).RowHeight = 35.5

    ' -- table FIRST (structured refs need it) ----------------------------
    Dim tbl As ListObject
    Set tbl = ws.ListObjects.Add(xlSrcRange, _
        ws.Range(ws.Cells(DL_HDR_ROW, FIRST_COL), ws.Cells(lastRow, FIRST_COL + modSpec.DL_NCOLS - 1)), , xlYes)
    tbl.Name = "DealLevelInput"
    tbl.TableStyle = ""

    ' -- header block AFTER the table exists: its formulas reference the
    '    DealLevelInput table and .Formula parses references immediately --
    modSpec.LoadHeaderBlock refs, vals, isNum, numv
    For i = LBound(refs) To UBound(refs)
        If isNum(i) Then
            ws.Range(refs(i)).Value = numv(i)
            ws.Range(refs(i)).Font.Color = RGB(0, 0, 255)          ' editable threshold
            ws.Range(refs(i)).Interior.Color = RGB(222, 235, 247)  ' light blue
        ElseIf Left$(vals(i), 1) = "=" Then
            ws.Range(refs(i)).Formula = Replace(vals(i), "{LAST}", CStr(lastRow))
        Else
            ws.Range(refs(i)).Value = vals(i)
        End If
    Next i

    ' -- columns: input specs as VALUES, F:/FT: specs as formulas ---------
    Dim colVals() As Variant, j As Long, zeroDefault As Boolean, v As Variant
    For i = 1 To modSpec.DL_NCOLS
        If Left$(f(i), 3) = "in:" Or Left$(f(i), 4) = "in0:" Then
            zeroDefault = (Left$(f(i), 4) = "in0:")
            j = InputColIndex(Mid$(f(i), IIf(zeroDefault, 5, 4)))
            ReDim colVals(1 To n, 1 To 1)
            For r = 1 To n
                v = data(r, j)
                If zeroDefault And (IsEmpty(v) Or Trim$(CStr(v)) = "") Then v = 0
                colVals(r, 1) = v
            Next r
            ws.Range(ws.Cells(DL_DATA_ROW, FIRST_COL + i - 1), _
                     ws.Cells(lastRow, FIRST_COL + i - 1)).Value = colVals
        ElseIf Len(f(i)) > 0 Then
            ReDim colVals(1 To n, 1 To 1)
            For r = DL_DATA_ROW To lastRow
                colVals(r - DL_DATA_ROW + 1, 1) = ResolveSpec(f(i), r)
            Next r
            ws.Range(ws.Cells(DL_DATA_ROW, FIRST_COL + i - 1), _
                     ws.Cells(lastRow, FIRST_COL + i - 1)).Formula = colVals
        End If
        If Len(fmt(i)) > 0 Then
            ws.Range(ws.Cells(DL_DATA_ROW, FIRST_COL + i - 1), _
                     ws.Cells(lastRow, FIRST_COL + i - 1)).NumberFormat = fmt(i)
        End If
    Next i

    ' -- readability styling (mirrors the Python build) -------------------
    Dim hdr As Range, body As Range
    Set hdr = ws.Range(ws.Cells(DL_HDR_ROW, FIRST_COL), _
                       ws.Cells(DL_HDR_ROW, FIRST_COL + modSpec.DL_NCOLS - 1))
    hdr.Interior.Color = RGB(31, 78, 120)
    hdr.Font.Color = RGB(255, 255, 255)
    hdr.Font.Bold = True
    hdr.WrapText = True
    hdr.VerticalAlignment = xlVAlignCenter
    Set body = ws.Range(ws.Cells(DL_HDR_ROW, FIRST_COL), _
                        ws.Cells(lastRow, FIRST_COL + modSpec.DL_NCOLS - 1))
    With body.Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(208, 208, 208)
    End With
    For i = 1 To modSpec.DL_NCOLS
        If h(i) = "Total" & Chr(10) & "Value" _
           Or h(i) = "Gross" & Chr(10) & "MOIC" _
           Or h(i) = "Performing" & Chr(10) & "(1=Underperform)" _
           Or h(i) = "InvCapital in Loss Position" _
           Or h(i) = "Impaired" & Chr(10) & "Value" Then
            ws.Range(ws.Cells(DL_DATA_ROW, FIRST_COL + i - 1), _
                     ws.Cells(lastRow, FIRST_COL + i - 1)).Interior.Color = RGB(242, 242, 242)
        End If
    Next i
    ws.Columns(2).ColumnWidth = 26
    ws.Range(ws.Columns(3), ws.Columns(FIRST_COL + modSpec.DL_NCOLS - 1)).ColumnWidth = 13
    On Error Resume Next
    ws.Activate
    ws.Range("C" & DL_DATA_ROW).Select
    ActiveWindow.FreezePanes = True
    On Error GoTo 0
End Sub
