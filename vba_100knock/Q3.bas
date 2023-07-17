Attribute VB_Name = "Q3"
Option Explicit

'#VBA100本ノック 3本目
'画像のように1行目に見出し、A列に№が入っています。
'№行数およびデータ行数は毎回変化します。
'この表の見出し（1行目）と№（A列）を残して、データ部分のみ値を消去してください。
'※シートはアクティブシート

Public Sub Q3_1()

    'Offsetでずらしているのではみ出る(この問題ではこのやり方でも問題ない)
    Range("A1").CurrentRegion.Offset(1, 1).ClearContents

End Sub

'Resizeではみ出した部分を消去
Public Sub Q3_2()

    With Range("A1").CurrentRegion

        .Offset(1, 1).Resize(.Rows.Count - 1, .Columns.Count - 1).ClearContents

    End With

End Sub

'Intersect(解説サイトに掲載されている方法)
Public Sub Q3_3()

    With Range("A1").CurrentRegion

        'Intersectで重なっている範囲を判定し消去
        .intersect(.Cells, .Offset(1, 1)).ClearContents

    End With

End Sub