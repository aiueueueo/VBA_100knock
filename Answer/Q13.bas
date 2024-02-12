Attribute VB_Name = "Q13"
Option Explicit

Public Sub Question13()
    Dim strWord As String   '検索文字列
    Dim intLen  As Long     '検索文字列の長さ
    Dim intLeft As Long     '検索文字列開始位置
    Dim rng     As Range

    strWord = "注意"
    intLen = Len(strWord)

    For Each rng In Selection
        If TypeName(rng) = "Range" Then
            intLeft = Instr(rng, strWord)
            If intLeft <> 0 Then rng.Characters(intLeft, intLen).Font.Color = vbRed
        Else
            MsgBox "セルを選択してください"
        End If
    Next

End Sub


'別解
Public Sub Question13_2()

    Dim target As Range
    On Error Resume Next
    'SpecialCellsで該当セルだけ絞り込み
    '1セルしか選択していない場合は、SpecialCellsは全セル対象になってしまうので、Intersectで本来の選択範囲を絞り込む
    Set target = Intersect(Selection, Selection.SpecialCells(xlCellTypeConstants, xlTextValues))
    If target Is Nothing Then Exit Sub

    Const cns注意 = "注意"

    Dim reg As Object
    Set reg = CreateObject("VBScript.RegExp")

    Dim rng As Range
    For Each rng In target
        Call CharactersFont(rng, cns注意, reg)
    Next

    Set reg = Nothing

End Sub

Sub CharactersFont(rng As Range, argStr As String, reg As Object)

    Dim mc As Object
    Dim m As Object

    With reg
        .Pattern = argStr               '検索パターンを設定
        .Global = True                  '文字列全体を検索
        Set mc = .Execute(rng.Value)    'マッチングの結果をMatchesコレクションで返す
    End With

    For Each m In mc
        With rng.Characters(m.FirstIndex + 1, m.Length)
            .Font.Bold = True
            .Font.Color = vbRed
        End With
    Next

End Sub