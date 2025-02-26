Attribute VB_Name = "Module2"
Sub net_maas_hesapla()
Dim brut_maas As Currency
Dim ssk_kesintili_maas As Currency
Dim net_maas  As Currency
Dim i As Integer

For i = 2 To 18
    brut_maas = Cells(i, "C").Value
    ssk_kesintili_maas = brut_maas * 0.85
    net_maas = ssk_kesintili_maas * 0.8
    Cells(i, "D").Value = net_maas
Next i

End Sub


Sub faktoriyel()
Dim carpim
carpim = 1

For i = 1 To 6
    carpim = carpim * i
Next

MsgBox "6 faktoriyel = " & carpim
End Sub
