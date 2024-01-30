Attribute VB_Name = "Q6"
Option Explicit

Public Sub Question6()

    Dim i As Long
    Dim r As Long

    r = Range("A1").CurrentRegion.Rows.Count

    For i = 2 To r

        If InStr(Range("A" & i).Value, "-") = 0 Then

            Range("D" & i).FormulaR1C1 = "=RC[-2]*RC[-1]"

        End If

    Next

End Sub