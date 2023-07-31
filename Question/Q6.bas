Option Explicit

Public Sub Q6_1()

    '最終行の取得
    Dim r As Long
    r = Range("A1").CurrentRegion.Rows.Count

    Dim i As Long
    For i = 2 To r

        Dim s As String
        s = Cells(i, "A").Value

        If Instr(s, "-") = 0 Then

            Cells(i, "D").Formula = "=B" & i & "*C" & i

        End If

    Next

End Sub

Public Sub Q6_2()

    '最終行の取得
    Dim r As Long
    r = Cells(Rows.Count, 1).End(xlUp).Row

    For i = 2 To r

        If Not Cells(i, 1).Value Like "*-*" Then

            Cells(i, "D").FormulaR1C1 = "=RC[-2]*RC[-1]"

        End If

    Next

End Sub