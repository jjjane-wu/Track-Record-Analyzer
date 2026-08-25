Attribute VB_Name = "modToc"
' ===================================================================
'  Table of Contents -- numbered, banded list of internal links to
'  every visible tab, placed as the workbook's first sheet. Each row
'  carries a one-line description of what the tab shows.
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
            Set c = ws.Cells(r, 4)
            c.Value = TabBlurb(sh.Name)
            c.Font.Size = 10
            c.Font.Color = RGB(89, 89, 89)
            If n Mod 2 = 0 Then c.Interior.Color = RGB(242, 242, 242)
            r = r + 1
        End If
    Next sh
    ws.Columns("B").ColumnWidth = 4.5
    ws.Columns("C").ColumnWidth = 42
    ws.Columns("D").ColumnWidth = 78
    ws.Move Before:=ThisWorkbook.Worksheets(1)
End Sub

' One-line description per known tab; unknown (future) tabs stay blank.
Private Function TabBlurb(nm As String) As String
    Select Case nm
        Case "Deal List"
            TabBlurb = "Every deal as plain values plus the full per-deal analytics; the blue threshold tables set the bucket boundaries"
        Case "Return & Loss Ratios"
            TabBlurb = "Pooled MOIC, Loss Ratio and Impaired Invested Capital across 15 cuts - sector, geography, vintage, fund, entry size, exit type - with a chart per cut"
        Case "Return Dispersion"
            TabBlurb = "MOIC and IRR distributions: deal count, % of invested capital and average return per bucket"
        Case "Portfolio Construction"
            TabBlurb = "Capital mix by fund x sector / geography, plus deal-count attribute breakdowns"
        Case "Vintage Perf by Sector"
            TabBlurb = "Invested capital, MOIC and loss ratios by vintage, with vintage x sector count and MOIC matrices"
        Case "Deployment & Exits"
            TabBlurb = "Capital deployment pacing (vintage x fund) and realization pacing (fund x exit year, realized deals)"
        Case "Underperforming Assets"
            TabBlurb = "Deals below the performance threshold, with their share of capital and value"
        Case "Partner Attribution"
            TabBlurb = "Returns and capital by sourcing partner"
        Case "Op Performance"
            TabBlurb = "Operating metrics (revenue / EBITDA growth, margins) for realized deals"
        Case "Op Performance - Unrealized"
            TabBlurb = "Operating metrics for the unrealized portfolio"
        Case Else
            TabBlurb = ""
    End Select
End Function
