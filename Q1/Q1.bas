Attribute VB_Name = "Q1"
Option Explicit

Sub Q1()

    '#VBA100ノック 1本目
    '「Sheet1」のA1:C5のセル範囲を、「Sheet2」のA1:C5にコピーしてください。
    '値も数式も書式も全てコピーしてください。
    'ただしSelectメソッドは使用禁止
    '※行高と列幅の設定はしなくて良い。

    Worksheets("Sheet1").Range("A1:C5").Copy _
        Worksheets("Sheet2").Range("A1:C5")

End Sub
