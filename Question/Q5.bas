Option Explicit

'#VBA100本ノック 5本目
'画像のようにB2から始まる表があります。
'B列×C列を計算した値をD列に入れ、通貨\のカンマ編集で表示してください。
'ただしB列またはC列が空欄の場合は空欄表示にしてください。
'例.D3にはB3×C3の計算結果の値を「\234,099」で表示、D5は空欄
'※ブック・シートは任意

Public Sub Q5()

    '最終行の取得
    Dim r As Long
    r = Range("B2").CurrentRegion.Rows.Count + 1

    Dim i As Long
    For i = 3 To r

        If Cells(i, "B").Value <> "" AND Cells(i, "C").Value <> "" Then

            Cells(i, "D").Value = Cells(i, "B").Value * Cells(i, "C").Value
            Cells(i, "D").NumberFormatLocal = "\#,##0"

        EndIf

    Next

End Sub