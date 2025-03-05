Attribute VB_Name = "Module21"
' Option Private Module '' Module makrolar� ana ekranda g�stermez

'Private Function kelime_anlami(ingilizce_kelime As String) As String
'�zel fonksiyon, Bu fonksiyon sadece bu mod�lde �a��r�l�r
'Public Function kelime_anlami(ingilizce_kelime As String) As String
'�zel fonksiyon, Bu fonksiyon her mod�lde �a�r�l�r, default hali

Function kelime_anlami(ingilizce_kelime As String) As String
    On Error GoTo hata
    ingilizce_kelime = Application.WorksheetFunction.Trim(ingilizce_kelime) '' Ba�ka sayfada �al��t�rmak i�in sheets("SayfaAdi").Range("AlanAdi")
    'kelime_anlami = Application.WorksheetFunction.VLookup(ingilizce_kelime, Sheets("ing_kelimeler").Range("A2:B36587"), 2, 0)
    kelime_anlami = Application.WorksheetFunction.VLookup(ingilizce_kelime, Veri_Sayfas�.Range("A2:B36587"), 2, 0)
                                                                            'VBa ad� sayfa5 idi Veri_sayfas� yapt�m art�k excel sayfa ad� de�i�se de sorun olmaz
                                                                            'VBA project sayfay� bul --  (Name) de�i�itir kal�c� isim.




hata:
    If Err Then
    kelime_anlami = "-- B�yle Bir Kelime Yoktur -- "
    End If
End Function

' for ile dola�t�m
Function kelime_anlami2(ingilizce_kelime As String) As String
    On Error GoTo hata ' Hata Analizi
    ' Bu tarihten sonra �al��maz.
'    If Date > DateValue("04.03.2025") Then
'        Exit Function
'    End If
    For i = 2 To 36587
        If Range("A" & i).Value = ingilizce_kelime Then
            kelime_anlami2 = Range("B" & i)
            Exit Function  ' Sonucu bulursan e�er burada bitirmek i�indir.
        End If
    Next i
        
hata:
    If Err Then
    kelime_anlami2 = "-- B�yle Bir Kelime Yoktur -- "
    End If

End Function


Function �oklu_kelime_anlami(ingilizce_kelime As String) As String
Dim kelimenin_anlamlari()
Dim sayi As Integer

For i = 2 To 36587
    If ingilizce_kelime = Range("A" & i).Value Then
        sayi = sayi + 1
        ReDim Preserve kelimenin_anlamlari(1 To sayi) ' Preserve yazmadan d�z yazarsam her boyutlama yapt���nda siler
        kelimenin_anlamlari(sayi) = Range("B" & i).Value
    End If

Next i

a = UBound(kelimenin_anlamlari)
b = LBound(kelimenin_anlamlari)

�oklu_kelime_anlami = Join(kelimenin_anlamlari, " ; " & Chr(10))  ' Chr(10) alt alta kayd�rma ama Metni Kayd�r �zelli�i aktif olmal�d�r.

End Function

Sub hata_bul()
MsgBox �oklu_kelime_anlami("ab")


End Sub


