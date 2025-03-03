Attribute VB_Name = "Module12"
Function ulke_sehir_toplam(ulke_adi As String, sehir_adi As String)
Dim Toplam_tutar As Currency
Toplam_tutar = 0
For i = 2 To 52
    If Range("A" & i).Value = ulke_adi And Range("B" & i).Value = sehir_adi Then
        Toplam_tutar = Toplam_tutar + Range("E" & i).Value
    End If
    
Next i

ulke_sehir_toplam = Toplam_tutar

End Function


Function ulke_sehir_ortalama(ulke_adi As String, sehir_adi As String)
Dim Toplam_tutar As Currency
Dim Toplam_adet As Integer
Toplam_tutar = 0
Toplam_adet = 0
For i = 2 To 52
    If Range("A" & i).Value = ulke_adi And Range("B" & i).Value = sehir_adi Then
        Toplam_tutar = Toplam_tutar + Range("E" & i).Value
        Toplam_adet = Toplam_adet + 1
    End If
    
Next i
If adet = 0 Then
    ulke_sehir_ortalama = 0
Else
End If

ulke_sehir_ortalama = Toplam_tutar / Toplam_adet

End Function
Function ulke_sehir_urunler(ulke_adi As String, sehir_adi As String)
Dim urun_listesi As String
For i = 2 To 52
    If Range("A" & i).Value = ulke_adi And Range("B" & i).Value = sehir_adi Then
        urun_listesi = urun_listesi & ";" & Chr(10) & Range("D" & i).Value
        '' Chr(10) bir alt satýra geçirir,
    End If
    
Next i

ulke_sehir_urunler = Trim(Mid(urun_listesi, 2, 999))


End Function

