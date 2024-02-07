Attribute VB_Name = "Q12"
Option Explicit

Public Sub Question12()

    Dim i As Long
    Dim firstNum As Long, fusoku As Long
    Dim rng As Range
    Dim r As Long

    r = Cells(Rows.Count, "C").End(xlUp).Row


    For i = 2 To r
        Set rng = Cells(i, "C").MergeArea
        If rng.Count > 1 Then
            rng.UnMerge
            firstNum = rng(1).Value
            rng.Value = Int(rng(1).Value / rng.Count)

            'いくつ足りないのかの計算
            fusoku = firstNum - (rng(1).Value * rng.Count)
            '1を分配
            If fusoku <> 0 Then rng.Resize(fusoku).Value = rng(1).Value + 1
        End If
    Next

End Sub