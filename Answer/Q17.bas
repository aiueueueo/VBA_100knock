Attribute VB_Name = "Q17"
Option Explicit

Public Sub Question17()

    Dim ws_社員 As Worksheet
    Set ws_社員 = Worksheets("社員")
    Dim ws_マスタ As Worksheet
    Set ws_マスタ = Worksheets("部・課マスタ")
    
    '部・課マスタの2行目以降を消去
    ws_マスタ.Range("A1").CurrentRegion.Offset(1).ClearContents
    
    '社員シートの最終行を取得
    Dim long_社員lastR As Long
    long_社員lastR = ws_社員.Range("A1").CurrentRegion.Rows.Count
    
    With ws_社員
        Dim dic_マスタ As New Dictionary
        Dim i As Long
        For i = 2 To long_社員lastR
            Dim tmp As Variant
            tmp = .Cells(i, "C").Value & ":" & .Cells(i, "D").Value
            If Not dic_マスタ.Exists(tmp) Then
                dic_マスタ.Add tmp, .Cells(i, "C").Resize(, 4).Value
            End If
        Next
    End With
    
    Dim item As Variant
    i = 2
    For Each item In dic_マスタ.Items
        ws_マスタ.Cells(i, "A").Resize(, 4).Value = item
        i = i + 1
    Next
    
    With ws_マスタ
        .Range("A1").CurrentRegion.Sort Key1:=.Range("A1"), order1:=xlAscending, _
                                        key2:=.Range("B1"), order2:=xlAscending, _
                                        Header:=xlYes
    End With
    
End Sub
