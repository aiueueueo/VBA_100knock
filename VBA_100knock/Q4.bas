Attribute Value = "Q4"
Option Explicit

'#VBA100本ノック 4本目
'画像のように1行目に見出し、A列に№が入っています。
'この表範囲の一部には計算式が入っています。
'（画像の最下行とD列には数式が入っています。）
'データ行数は毎回変化します。
'見出し行とA列№と計算式は残し、定数値だけを消去してください。

'範囲の最下行と一番右の列に計算式が入っている場合
Public Sub Q4-1()

    With Range("A1").CurrentRegion

        .Offset(1, 1).Resize(.Rows.Count - 2, .Columns.Count - 2).ClearContents

    End With

End Sub

Public Sub Q4-2()

    With Range("A1").CurrentRegion.Offset(1, 1)

        On Error Resume Next
        '定数が含まれているセルを削除
        .SpecialCells(xlCellTypeConstants).ClearContents

    End With

End Sub