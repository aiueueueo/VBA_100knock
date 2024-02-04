Attribute VB_Name = "Q9"
Option Explicit

Public Sub Question9()

    Dim wsSeiseki As Worksheet
    Dim wsGoukaku As Worksheet

    Set wsSeiseki = Worksheets("ê¨ê—ï\")

    If Not SheetCheck("çáäié“") Then
        Set wsGoukaku = Worksheets.Add(After:=Worksheets("ê¨ê—ï\"))
        wsGoukaku.Name = "çáäié“"
    Else
        Set wsGoukaku = Worksheets("çáäié“")
        wsGoukaku.Cells.ClearContents
    End If

    With wsSeiseki

        .AutoFilterMode = False
        .Range("A1").AutoFilter Field:=7, Criteria1:="çáäi"
        .Columns("A").Copy
        wsGoukaku.Range("A1").PasteSpecial
        .AutoFilterMode = False

    End With

End Sub

Public Function SheetCheck(wsName As String) As Boolean

    Dim ws As Worksheet

    On Error Resume Next
    Set ws = Worksheets(wsName)
    On Error GoTo 0

    SheetCheck = Not ws Is Nothing

End Function