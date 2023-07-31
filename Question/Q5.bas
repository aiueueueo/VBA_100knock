Option Explicit

Public Sub Q5()

    '最終行の取得
    Dim r As Long
    r = Range("B2").CurrentRegion.Rows.Count + 1

    Dim i As Long
    For i = 3 To r

        If Cells(i, "B").Value <> "" AND Cells(i, "C").Value <> "" Then

            Cells(i, "D").Value = Cells(i, "B").Value * Cells(i, "C").Value
            Cells(i, "D").NumberFormatLocal = "\#,##0"

        EndIf

    Next

End Sub