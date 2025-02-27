Attribute VB_Name = "Module10"
Function iki_sayi_topla(sayi1, sayi2)

iki_sayi_topla = sayi1 + sayi2

End Function

Function us_hesapla(sayi1, us)
us_hesapla = sayi1 ^ us
End Function

Function faktoriyel_hesapla(sayi1)
Dim carpim
Dim i

carpim = 1
For i = 1 To sayi1
carpim = carpim * i
Next
faktoriyel_hesapla = carpim
End Function

Sub makro_18()
    MsgBox faktoriyel_hesapla(5)
End Sub
