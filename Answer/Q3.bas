Attribute VB_Name = "Q3"
Option Explicit

Public Sub Question3()

    Dim rng As Range
    Set rng = ActiveSheet.Range("A1").CurrentRegion.Offset(1, 1)

    If Application.WorksheetFunction.CountA(rng) = 0 Then

        MsgBox "�f�[�^�����݂��܂���B" & vbCrLf & _
                "�������I�����܂��B", vbExclamation

    End If

    rng.ClearContents

End Sub

