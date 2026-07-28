Attribute VB_Name = "modToc"
' ===================================================================
'  Table of Contents -- numbered, banded list of internal links to
'  every visible tab, placed as the workbook's first sheet.
' ===================================================================
Option Explicit

Public Sub BuildTOC()
    Dim ws As Worksheet, sh As Worksheet, r As Long, n As Long, c As Range

    Set ws = modBuild.FreshSheet("Table of Contents")
    ws.Range("B2").Value = "Table of Contents"
    ws.Range("B2").Font.Bold = True: ws.Range("B2").Font.Size = 16

    r = 4
    For Each sh In ThisWorkbook.Worksheets
        If sh.Name <> ws.Name And sh.Visible = xlSheetVisible Then
            n = n + 1
            Set c = ws.Cells(r, 2)
            c.Value = n
            c.Font.Bold = True: c.Font.Color = RGB(255, 255, 255)
            c.Interior.Color = RGB(31, 78, 120)
            c.HorizontalAlignment = xlCenter
            Set c = ws.Cells(r, 3)
            ws.Hyperlinks.Add Anchor:=c, Address:="", _
                SubAddress:="'" & sh.Name & "'!A1", TextToDisplay:=sh.Name
            If n Mod 2 = 0 Then c.Interior.Color = RGB(242, 242, 242)
            r = r + 1
        End If
    Next sh
    ws.Columns("B").ColumnWidth = 4.5
    ws.Columns("C").ColumnWidth = 42
    ws.Move Before:=ThisWorkbook.Worksheets(1)
End Sub
