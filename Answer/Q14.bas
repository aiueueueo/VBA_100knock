Attribute VB_Name = "Q14"
Option Explicit

'要件未定なため保留
Public Sub Question14()

    Dim ws As Worksheet
    For Each ws In Worksheets
        If InStr(ws.Name, "社外秘") <> 0 Then 
            ws.Delete
        Else
            ws.Cells.Copy
            ws.PasteSpecial xlPasteValues
        End If
    Next

End Sub