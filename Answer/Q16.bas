Attribute VB_Name = "Q16"
Option Explicit

Public Function Question16(str As String) As String
    
    With CreateObject("VBScript.RegExp")
        .Global = True
        .Pattern = "^\n+|\n+$|\n+(?=\n)"
        Question16 = .Replace(Replace(str, vbCrLf, vbLf), "")
    End With

End Function
