Attribute VB_Name = "Q8"
Option Explicit

Public Sub Question8()

    Dim ws As Worksheet
    Dim r As Long
    Dim i As Long, j As Long
    Dim sum As Long

    Set ws = Worksheets("ê¨ê—ï\")

    r = Range("A1").CurrentRegion.Rows.Count

    With ws

        For i = 2 To r

            sum = 0
            For j = 2 To 6
                If .Cells(i, j).Value < 50 Then
                    GoTo CONTINUE
                End If
                sum = sum + .Cells(i, j).Value
            Next

            If sum >= 350 Then .Range("G" & i).Value = "çáäi"

CONTINUE:
        Next

    End With

End Sub

Public Sub Question8_1()

    Dim ws As Worksheet
    Set ws = Worksheets("ê¨ê—ï\")

    Dim rng As Range
    Set rng = ws.Range("A1").CurrentRegion
    Set rng = Intersect(rng, rng.Offset(1))
    rng.Columns("G").ClearContents

    Dim ws As Worksheet
    Set ws = Worksheets("ê¨ê—ï\")

    Dim rng As Range
    Set rng = ws.Range("A1").CurrentRegion
    Set rng = Intersect(rng, rng.Offset(1))
    rng.Columns("G").ClearContents

   Dim r As Range
    For Each r In rng.Rows
        With WorksheetFunction
            If .sum(r.Offset(, 1).Resize(, 5)) >= 350 And _
               .CountIf(r.Offset(, 1).Resize(, 5), ">=50") = 5 Then
                r.Columns("G") = "çáäi"
            End If
        End With
    Next

End SUb