Attribute VB_Name = "Module9"
'' Net maaþ hesaplama
Sub makro_16()
Dim brut_maas  As Currency
Dim ssk_kes As Currency
Dim vergi_oncesi As Currency
Dim net_maas As Currency
Dim i As Integer

For i = 2 To 21
    brut_maas = Cells(i, "C").Value
    ssk_kes = brut_maas * 0.85
    net_maas = ssk_kes * 0.8
    Cells(i, "D").Value = net_maas
    
Next i

End Sub

'' FAktoriyel

Sub makro_17()
Dim carpým As Long
carpým = 1

For i = 1 To 10
    carpým = carpým * i
Next
MsgBox "10 Faktoriyel =" & carpým
End Sub
