Option Explicit

Sub Q8()

    Dim r As Long
    r = Cells(Rows.Count, 1).End(xlUp).Row
    
    Dim i As Long, j As Long, sum As Long
    For i = 2 To r
    
        sum = 0
        
        For j = 2 To 6
        
            If Cells(i, j).Value < 50 Then
        
                Exit For
            
            End If
            
            sum = sum + Cells(i, j).Value
        
        Next
    
        If sum >= 350 Then
    
            Cells(i, "G").Value = "合格"
    
        End If
    
    Next

End Sub

Sub Q8_2()

    Dim ws As Worksheet
    Set ws = Worksheets("成績表")
  
    Dim rng As Range
    Set rng = ws.Range("A1").CurrentRegion
    Set rng = Intersect(rng, rng.Offset(1))
    rng.Columns("G").ClearContents
  
    Dim r As Range
    For Each r In rng.Rows
        With WorksheetFunction
            If .sum(r.Offset(, 1).Resize(, 5)) >= 350 And _
               .CountIf(r.Offset(, 1).Resize(, 5), ">=50") = 5 Then
                r.Columns("G") = "合格"
            End If
        End With
    Next
End Sub

