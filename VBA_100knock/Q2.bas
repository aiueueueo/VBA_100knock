Attribute VB_Name = "Q2"
Option Explicit

'#VBA100本ノック 2本目
'「Sheet1」のA1:C5のセル範囲を、「Sheet2」のA1:C5にコピーしてください。
'数式は消して値でコピー、書式もコピーしてください。
'※書式は「セルの書式設定」で設定可能なもの（ロックは除く）。
'入力規則やメモ（旧コメント）は書式ではありません。
'「ふりがな」は任意で

Public Sub Q2()    

    With Worksheets("Sheet2").Range("A1:C5")

        Worksheets("Sheet1").Range("A1:C5").Copy

        '値のコピー
        .PasteSpecial(XlPasteValues)
        '.PasteSpecial Paste:=xlPasteValues
        '.Value = Worksheets("Sheet1").Value

        '書式のコピー
        .PasteSpecial(XlPasteFormats)
        '.PasteSpecial Paste:=xlPasteFormats

    End With

    'クリップボードのクリア
    Application.CutCopyMode = False

End Sub