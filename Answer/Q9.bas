Attribute VB_Name = "Q9"
Option Explicit

Public Sub Question9()

    Dim wsSeiseki As Worksheet
    Dim wsGoukaku As Worksheet

    Set wsSeiseki = Worksheets("���ѕ\")

    If Not SheetCheck("���i��") Then
        Set wsGoukaku = Worksheets.Add(After:=Worksheets("���ѕ\"))
        wsGoukaku.Name = "���i��"
    Else
        Set wsGoukaku = Worksheets("���i��")
        wsGoukaku.Cells.ClearContents
    End If

    With wsSeiseki

        .AutoFilterMode = False
        .Range("A1").AutoFilter Field:=7, Criteria1:="���i"
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