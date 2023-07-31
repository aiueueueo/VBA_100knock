Option Explicit

Public Sub Q1()

    Worksheets("Sheet1").Range("A1:C5").Copy Worksheets("Sheet2").Range("A1:C5")
    'valueは値のみコピー

End Sub
