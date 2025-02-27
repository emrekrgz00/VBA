Attribute VB_Name = "Module11"
Function makine_adi(urun_kodu As String) As String
    Dim ilk_karakter As String
    ilk_karakter = Left(urun_kodu, 1)
    
    Select Case ilk_karakter
            Case Is = "A"
            makine_adi = "Haddeeleme"
            Case Is = "B"
            makine_adi = "Torna"
            Case Is = "C"
            makine_adi = "Freeze"
            Case Is = "D"
            makine_adi = "TamamlanmýþÜrün"
            Case Else
            makine_adi = "HaddeelemeBöyle Bir ürün yoktur."

    End Select
    

End Function


'' Urun capý
Function makine_adi2(urun_kodu As String) As String
    Dim ikinci_karakter As Long
    ikinci_karakter = Right(urun_kodu, 2)
    makine_adi2 = ikinci_karakter
End Function

'' urun capý
Function urun_capi(urun_kodu As String) As Integer
urun_capi = Mid(urun_kodu, 2, 5)

End Function

''' -----------

Function net_maas(brut_maas As Currency) As Currency

Dim ssk_maas As Currency
Dim vergi_maas As Currency

ssk_maas = brut_maas * 0.85
vergi_maas = ssk_maas * 0.8
net_maas = vergi_maas

End Function
'''----
Function sayi_degeri(sayi As Integer) As String
    If sayi > 50 Then
        sayi_degeri = "Bu sayi 50'den büyüktür"
    ElseIf sayi < 50 Then
        sayi_degeri = "Bu sayi 50'den küçüktür"
    Else
        sayi_degeri = "Bu sayi 50'ye eþittir."
    End If
End Function
