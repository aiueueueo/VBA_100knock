Attribute VB_Name = "Q10"
Option Explicit

Public Sub Question10()

    Dim ws As Worksheet
    Set ws = Worksheets("受注")

    ws.AutoFilterMode = False

    Dim rng As Range
    With ws.Range("A1").CurrentRegion
        .AutoFilter 3, ""
        .AutoFilter 4, "*削除*", xlOr, "*不要*"

        'フィルターされている行がある場合だけ削除
        Set rng = .Offset(1, 0).Resize(.Rows.Count-1)
        If Not rng Is Nothing Then rng.EntireRow.Delete
    End With

    ws.AutoFilterMode = False

End Sub