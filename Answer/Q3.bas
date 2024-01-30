Attribute VB_Name = "Q3"
Option Explicit

Public Sub Question3()

    Dim rng As Range
    Set rng = ActiveSheet.Range("A1").CurrentRegion.Offset(1, 1)

    If Application.WorksheetFunction.CountA(rng) = 0 Then

        MsgBox "データが存在しません。" & vbCrLf & _
                "処理を終了します。", vbExclamation

    End If

    rng.ClearContents

End Sub

