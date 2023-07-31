Option Explicit

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