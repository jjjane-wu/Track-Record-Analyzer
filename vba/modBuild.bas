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

' column letter of an input-tab header (for in:-link resolution).
' Called once per formula cell, so the spec array is loaded once and cached.
Private Function InputColLetter(ByVal header As String) As String
    Static loaded As Boolean
    Static hdrs() As String
    Dim i As Long
    If Not loaded Then
        modSpec.LoadInputCols hdrs
        loaded = True
    End If
    For i = 1 To modSpec.IN_NCOLS
        If hdrs(i) = header Then
            InputColLetter = Split(Cells(1, FIRST_COL + i - 1).Address, "$")(1)
            Exit Function
        End If
    Next i
    Err.Raise vbObjectError + 1, , "Input column not found: " & header
End Function

' same prefix semantics as Python _dl_src_formula
Private Function ResolveSpec(ByVal spec As String, ByVal r As Long) As String
    Dim inputRow As Long, col As String, src As String
    inputRow = IN_DATA_ROW + (r - DL_DATA_ROW)
    If Len(spec) = 0 Then
        ResolveSpec = ""
    ElseIf Left$(spec, 4) = "in0:" Then
        col = InputColLetter(Mid$(spec, 5))
        src = "'Deal Level Inputs'!" & col & inputRow
        ResolveSpec = "=IF(" & src & "="""",0," & src & ")"
    ElseIf Left$(spec, 3) = "in:" Then
        col = InputColLetter(Mid$(spec, 4))
        src = "'Deal Level Inputs'!" & col & inputRow
        ResolveSpec = "=IF(" & src & "="""",""""," & src & ")"
    ElseIf Left$(spec, 3) = "FT:" Then
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

    ' -- meta block ---------------------------------------------------
    ws.Range("B2").Value = "Deal List": ws.Range("B2").Font.Bold = True
    ws.Range("B2").Font.Size = 16
    ws.Range("B4").Value = "Sponsor/GP:":  ws.Range("C4").Formula = "='Deal Level Inputs'!C3"
    ws.Range("B5").Value = "As of Date:":  ws.Range("C5").Formula = "='Deal Level Inputs'!C4"
    ws.Range("C5").NumberFormat = "d-mmm-yy"
    ws.Range("B6").Value = "Currency:":    ws.Range("C6").Formula = "='Deal Level Inputs'!C5"

    Dim lastRow As Long: lastRow = DL_DATA_ROW + n - 1

    ' -- tag row 12 + header row 13 -----------------------------------
    For i = 1 To modSpec.DL_NCOLS
        If Len(tag(i)) > 0 Then ws.Cells(DL_HDR_ROW - 1, FIRST_COL + i - 1).Value = tag(i)
        ws.Cells(DL_HDR_ROW, FIRST_COL + i - 1).Value = h(i)
        ws.Cells(DL_HDR_ROW, FIRST_COL + i - 1).Font.Bold = True
    Next i
    ws.Rows(DL_HDR_ROW).RowHeight = 35.5

    ' -- table FIRST (structured refs need it), then formulas ----------
    Dim tbl As ListObject
    Set tbl = ws.ListObjects.Add(xlSrcRange, _
        ws.Range(ws.Cells(DL_HDR_ROW, FIRST_COL), ws.Cells(lastRow, FIRST_COL + modSpec.DL_NCOLS - 1)), , xlYes)
    tbl.name = "DealLevelInput"
    tbl.TableStyle = ""

    ' -- header block AFTER the table exists: its formulas reference the
    '    DealLevelInput table and .Formula parses references immediately
    '    (the Python XML writer does not care about order; VBA does)
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

    Dim colFormulas() As Variant
    For i = 1 To modSpec.DL_NCOLS
        If Len(f(i)) > 0 Then
            ReDim colFormulas(1 To n, 1 To 1)
            For r = DL_DATA_ROW To lastRow
                colFormulas(r - DL_DATA_ROW + 1, 1) = ResolveSpec(f(i), r)
            Next r
            ws.Range(ws.Cells(DL_DATA_ROW, FIRST_COL + i - 1), _
                     ws.Cells(lastRow, FIRST_COL + i - 1)).Formula = colFormulas
        End If
        If Len(fmt(i)) > 0 Then
            ws.Range(ws.Cells(DL_DATA_ROW, FIRST_COL + i - 1), _
                     ws.Cells(lastRow, FIRST_COL + i - 1)).NumberFormat = fmt(i)
        End If
    Next i

    ' -- readability styling (mirrors the Python build's conventions) --
    Dim hdr As Range, body As Range
    Set hdr = ws.Range(ws.Cells(DL_HDR_ROW, FIRST_COL), _
                       ws.Cells(DL_HDR_ROW, FIRST_COL + modSpec.DL_NCOLS - 1))
    hdr.Interior.Color = RGB(31, 78, 120)         ' dark blue banner
    hdr.Font.Color = RGB(255, 255, 255)
    hdr.Font.Bold = True
    hdr.WrapText = True
    hdr.VerticalAlignment = xlVAlignCenter

    Set body = ws.Range(ws.Cells(DL_HDR_ROW, FIRST_COL), _
                        ws.Cells(lastRow, FIRST_COL + modSpec.DL_NCOLS - 1))
    With body.Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(208, 208, 208)               ' light grid
    End With

    ' five key computed columns get the template's light-grey fill
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

    ' widths + frozen header row / company column
    ws.Columns(2).ColumnWidth = 26                ' Company
    ws.Range(ws.Columns(3), ws.Columns(FIRST_COL + modSpec.DL_NCOLS - 1)).ColumnWidth = 13
    On Error Resume Next                          ' needs a visible window
    ws.Activate
    ws.Range("C" & DL_DATA_ROW).Select
    ActiveWindow.FreezePanes = True
    On Error GoTo 0
End Sub
