Attribute VB_Name = "Q4"
Option Explicit

Public Sub Question4()

    Dim rng As Range
    Set rng = ActiveSheet.Range("A1").CurrentRegion.Offset(1, 1)

    '�Y������Z���������ꍇ�̓G���[�ɂȂ��Ă��܂��̂őΏ�����K�v������
    On Error Resume Next
    rng.SpecialCells(xlCellTypeConstants).ClearContents

End Sub
