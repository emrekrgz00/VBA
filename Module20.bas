Attribute VB_Name = "Module20"
Sub makro_34()
Dim mail_adresi(1 To 15, 1 To 500)
Dim mailler
Dim satir As Long
satir = 0

For i = 1 To 15
    mailler = Split(Range("A" & i).Value, ";") ' yazýyý dizye atýyor ayraç ;
    For j = 0 To UBound(mailler)
        mail_adresi(i, j + 1) = mailler(j)
    Next j
Next i

For i = 1 To 15
    For j = 1 To 500
        If mail_adresi(i, j) <> "" Then
            satir = satir + 1
            Range("B" & satir).Value = mail_adresi(i, j)
         
        End If
    
    Next j


Next i
End Sub


