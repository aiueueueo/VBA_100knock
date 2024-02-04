Attribute VB_Name = "Q7"
Option Explicit

Public Sub Question7()

    Dim i As Long
    Dim r As Long
    Dim d As Variant, ld As Variant

    r = Cells(Rows.Count, "A").End(xlUp).Row

    For i = 2 To r 

        d = Replace(Range("A" & i).Value, ".", "/")

        If IsDate(d) Then

            '––“úŽæ“¾
            ld = DateSerial(Year(d), Month(d) + 1, 0)
            '0–„‚ß‚Ì‚½‚ß‚É•¶Žš—ñ•\Ž¦
            Range("B" & i).Value = Format(ld, "'mmdd")

        End If

    Next

End Sub