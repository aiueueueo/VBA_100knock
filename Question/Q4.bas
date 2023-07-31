Option Explicit

'範囲の最下行と一番右の列に計算式が入っている場合
Public Sub Q4_1()

    With Range("A1").CurrentRegion

        .Offset(1, 1).Resize(.Rows.Count - 2, .Columns.Count - 2).ClearContents

    End With

End Sub

Public Sub Q4_2()

    With Range("A1").CurrentRegion.Offset(1, 1)

        On Error Resume Next
        '定数が含まれているセルを削除
        .SpecialCells(xlCellTypeConstants).ClearContents

    End With

End Sub