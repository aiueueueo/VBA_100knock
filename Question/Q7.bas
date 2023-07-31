Option Explicit

Public Sub Q7(

    '最終行の取得
    Dim r As Long
    r = Cells(Rows.Count, 1).End(xlUp).Row

    Dim i As Long
    Dim d As Variant
    For i = 2 To r 

        d = Raplace(Cells(i, "A").Value, ".", "/")

        If IsDate(d) Then

            d = CDate(d)
            Cells(i, "B") = Format(DateSerial(Year(d), Month(d) + 1, 0), "'mmdd")

        EndIf

    Next

)