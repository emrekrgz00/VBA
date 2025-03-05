Attribute VB_Name = "Module21"
' Option Private Module '' Module makrolarý ana ekranda göstermez

'Private Function kelime_anlami(ingilizce_kelime As String) As String
'Özel fonksiyon, Bu fonksiyon sadece bu modülde çaðýrýlýr
'Public Function kelime_anlami(ingilizce_kelime As String) As String
'Özel fonksiyon, Bu fonksiyon her modülde çaðrýlýr, default hali

Function kelime_anlami(ingilizce_kelime As String) As String
    On Error GoTo hata
    ingilizce_kelime = Application.WorksheetFunction.Trim(ingilizce_kelime) '' Baþka sayfada çalýþtýrmak için sheets("SayfaAdi").Range("AlanAdi")
    'kelime_anlami = Application.WorksheetFunction.VLookup(ingilizce_kelime, Sheets("ing_kelimeler").Range("A2:B36587"), 2, 0)
    kelime_anlami = Application.WorksheetFunction.VLookup(ingilizce_kelime, Veri_Sayfasý.Range("A2:B36587"), 2, 0)
                                                                            'VBa adý sayfa5 idi Veri_sayfasý yaptým artýk excel sayfa adý deðiþse de sorun olmaz
                                                                            'VBA project sayfayý bul --  (Name) deðiþitir kalýcý isim.




hata:
    If Err Then
    kelime_anlami = "-- Böyle Bir Kelime Yoktur -- "
    End If
End Function

' for ile dolaþtým
Function kelime_anlami2(ingilizce_kelime As String) As String
    On Error GoTo hata ' Hata Analizi
    ' Bu tarihten sonra çalýþmaz.
'    If Date > DateValue("04.03.2025") Then
'        Exit Function
'    End If
    For i = 2 To 36587
        If Range("A" & i).Value = ingilizce_kelime Then
            kelime_anlami2 = Range("B" & i)
            Exit Function  ' Sonucu bulursan eðer burada bitirmek içindir.
        End If
    Next i
        
hata:
    If Err Then
    kelime_anlami2 = "-- Böyle Bir Kelime Yoktur -- "
    End If

End Function


Function çoklu_kelime_anlami(ingilizce_kelime As String) As String
Dim kelimenin_anlamlari()
Dim sayi As Integer

For i = 2 To 36587
    If ingilizce_kelime = Range("A" & i).Value Then
        sayi = sayi + 1
        ReDim Preserve kelimenin_anlamlari(1 To sayi) ' Preserve yazmadan düz yazarsam her boyutlama yaptýðýnda siler
        kelimenin_anlamlari(sayi) = Range("B" & i).Value
    End If

Next i

a = UBound(kelimenin_anlamlari)
b = LBound(kelimenin_anlamlari)

çoklu_kelime_anlami = Join(kelimenin_anlamlari, " ; " & Chr(10))  ' Chr(10) alt alta kaydýrma ama Metni Kaydýr özelliði aktif olmalýdýr.

End Function

Sub hata_bul()
MsgBox çoklu_kelime_anlami("ab")


End Sub


