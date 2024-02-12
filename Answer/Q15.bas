Attribute VB_Name = "Q15"
Option Explicit

Public Sub Question15()

    Dim i As Long, j As Long

    For i = 1 To Sheets.Count -1
        For j = i + 1 To Sheets.Count
            If Sheets(i).Name > Sheets(j).Name Then
                Sheets(j).Move Before:=Sheets(i)
            End If
        Next
    Next

    Worksheets(1).Activate

End Sub