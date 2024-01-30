Attribute VB_Name = "Q1"
Option Explicit

'メイン処理
Public Sub Question1()

    'Sheet1が存在するか確認
    If Not SheetCheck("Sheet1") Then

        MsgBox "Sheet1が存在しません。" & vbCrLf & _
                "処理を終了します。", vbExclamation
        Exit Sub

    End If

    'Sheet2が存在するか確認
    If Not SheetCheck("Sheet2") Then

        Worksheets.Add(After:=Worksheets(Worksheets.Count)).Name = "Sheet2"

    End If

    Worksheets("Sheet1").Range("A1:C5").Copy
    Worksheets("Sheet2").Range("A1").PasteSpecial

    Application.CutCopyMode = False

End Sub

'該当するシートが存在するか確認するプロシージャ
'   シートが存在する場合はTrue、しない場合はFalseを返す
Public Function SheetCheck(wsName As String) As Boolean

    Dim ws As Worksheet

    On Error Resume Next
    Set ws = Worksheets(wsName)
    On Error GoTo 0

    SheetCheck = Not ws Is Nothing

End Function
