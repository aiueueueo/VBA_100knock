Attribute VB_Name = "Q5"
Option Explicit

Public Sub Question5()

    Dim i As Long
    Dim r As Long

    r = Range("B2").CurrentRegion.Rows.Count + 1

    For i = 3 To r

        If Range("B" & i).Value = "" Or Range("C" & i).Value = "" Then

            GoTo CONTINUE

        Else

            Range("D" & i).Value = Range("B" & i).Value * Range("C" & i).Value

        End If

CONTINUE:
    Next

    Columns("D").NumberFormatLocal = "\#,##0"

End Sub
