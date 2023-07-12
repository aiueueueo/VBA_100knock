Attribute VB_Name = "Module1"
Attribute Value = "Q3"
Option Explicit

'#VBA100本ノック 3本目
'画像のように1行目に見出し、A列に№が入っています。
'№行数およびデータ行数は毎回変化します。
'この表の見出し（1行目）と№（A列）を残して、データ部分のみ値を消去してください。
'※シートはアクティブシート

Public Sub Q3-1()

    'Offsetでずらしているのではみ出る(この問題ではこのやり方でも問題ない)
    Range("A1").CurrentRegion.Offset(1, 1).ClearContents

End Sub

Public Sub Q3-2()

    '

End Sub
