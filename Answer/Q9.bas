Attribute VB_Name = "Q9"
Option Explicit

Public Sub Question9()

    Dim wsSeiseki As Worksheet
    Dim wsGoukaku As Worksheet

    Set wsSeiseki = Worksheets("成績表")

    If Not SheetCheck("合格者") Then
        Set wsGoukaku = Worksheets.Add(After:=Worksheets("成績表"))
        wsGoukaku.Name = "合格者"
    Else
        Set wsGoukaku = Worksheets("合格者")
        wsGoukaku.Cells.ClearContents
    End If

    With wsSeiseki

        .AutoFilterMode = False
        .Range("A1").AutoFilter Field:=7, Criteria1:="合格"
        .Columns("A").Copy
        wsGoukaku.Range("A1").PasteSpecial
        .AutoFilterMode = False

    End With

End Sub

'シートチェック
Public Function SheetCheck(wsName As String) As Boolean

    Dim ws As Worksheet

    On Error Resume Next
    Set ws = Worksheets(wsName)
    On Error GoTo 0

    SheetCheck = Not ws Is Nothing

End Function