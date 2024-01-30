Attribute VB_Name = "Q1"
Option Explicit

'���C������
Public Sub Question1()

    'Sheet1�����݂��邩�m�F
    If Not SheetCheck("Sheet1") Then

        MsgBox "Sheet1�����݂��܂���B" & vbCrLf & _
                "�������I�����܂��B", vbExclamation
        Exit Sub

    End If

    'Sheet2�����݂��邩�m�F
    If Not SheetCheck("Sheet2") Then

        Worksheets.Add(After:=Worksheets(Worksheets.Count)).Name = "Sheet2"

    End If

    Worksheets("Sheet1").Range("A1:C5").Copy
    Worksheets("Sheet2").Range("A1").PasteSpecial

    Application.CutCopyMode = False

End Sub

'�Y������V�[�g�����݂��邩�m�F����v���V�[�W��
'   �V�[�g�����݂���ꍇ��True�A���Ȃ��ꍇ��False��Ԃ�
Public Function SheetCheck(wsName As String) As Boolean

    Dim ws As Worksheet

    On Error Resume Next
    Set ws = Worksheets(wsName)
    On Error GoTo 0

    SheetCheck = Not ws Is Nothing

End Function
