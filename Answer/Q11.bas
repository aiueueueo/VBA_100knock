Attribute VB_Name = "Q11"
Option Explicit

Public Sub Question11()

    Dim i As Long
    Dim rng As Range

    For Each rng In Range("A1").CurrentRegion
        If rng.MergeCells Then
            If rng.Address = rng.MergeArea.Item(1) Then
                If Not rng.AddComment Is Nothing Then
                    rng.ClearComments
                End If
                rng.AddComment "!"
            End If
        End If
    Next

End Sub