Option Explicit

Public Sub Q3_1()

    Range("A1").CurrentRegion.Offset(1, 1).ClearContents

End Sub

Public Sub Q3_2()

    With Range("A1").CurrentRegion
    
        Intersect(.Cells, .Offset(1, 1)).ClearContents

    End With

End Sub
