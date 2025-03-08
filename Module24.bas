Attribute VB_Name = "Module24"
Sub veri_giris_formu_ac()
VeriGirisFormu.Show vbModeless
End Sub

Private Sub IlkVeriGetir_Click()
Range("A2").Select
Call veri_getir
End Sub

Private Sub ÖncekiVeriGetir_Click()
If Selection.Row = 2 Then
    Exit Sub
End If

Selection.Offset(-1, 0).Select
Call veri_getir
End Sub

Private Sub SonrakiVeriGetir_Click()
Dim son_satir As Long
son_satir = Range("A50000").End(xlUp).Row
If Selection.Row > son_satir Then
    Exit Sub
End If

'MsgBox son_satir
'MsgBox Selection.Row
Selection.Offset(1, 0).Select
Call veri_getir
End Sub

Private Sub SonVeriGetir_Click()
Range("A50000").End(xlUp).Select
Call veri_getir
End Sub


Sub veri_getir()
'Veri getirme Makrosu
Dim secili_satir As Long
secili_satir = Selection.Row

AdiveSoyadi.Text = Range("A" & secili_satir).Value
Mezuniyet.Text = Range("B" & secili_satir).Value
DogumYeri.Text = Range("C" & secili_satir).Value
Adres.Text = Range("D" & secili_satir).Value
Departman.Text = Range("E" & secili_satir).Value

If Range("F" & secili_satir).Value = "Erkek" Then
    Erkek.Value = True
Else
    Kadin.Value = True
End If

If Range("G" & secili_satir).Value Like "*Ýngilizce*" Then
    Ýngilizce.Value = True
Else
    Ýngilizce.Value = False
End If

If Range("G" & secili_satir).Value Like "*Almanca*" Then
    Almanca.Value = True
Else
    Almanca.Value = False
End If
    
If Range("G" & secili_satir).Value Like "*Fransýzca*" Then
    Fransýzca.Value = True
Else
    Fransýzca.Value = False
End If



'AdiveSoyadi.Text = Selection.Value
'Mezuniyet.Text = Selection.Offset(0, 1).Value
'DogumYeri.Text = Selection.Offset(0, 2).Value

End Sub

' Kapatýldýuðýnda
Private Sub Kapat_Click()
Unload VeriGirisFormu
End Sub

'Kaydet butonunda kaydetme iþlemleri
Private Sub Kaydet_Click()
Dim son_satir As Long
Dim bildigi_diller As String
son_satir = Range("A50000").End(xlUp).Row + 1  'Son satir +1 (boþ satir)
Range("A" & son_satir).Value = AdiveSoyadi.Text
Range("B" & son_satir).Value = Mezuniyet.Text
Range("C" & son_satir).Value = DogumYeri.Text
Range("D" & son_satir).Value = Adres.Text
Range("E" & son_satir).Value = Departman.Text
If Erkek.Value = True Then
    Range("F" & son_satir).Value = "Erkek"
Else
    Range("F" & son_satir).Value = "Kadýn"
End If

If Ýngilizce = True Then
    bildigi_diller = bildigi_diller & " " & "Ýngilizce"
End If

If Almanca = True Then
    bildigi_diller = bildigi_diller & " " & "Almanca"
End If

If Fransýzca = True Then
    bildigi_diller = bildigi_diller & " " & "Fransýzca"
End If

Range("G" & son_satir).Value = bildigi_diller


AdiveSoyadi.Text = ""
Mezuniyet.Text = ""
DogumYeri.Text = ""
Adres.Text = ""
Departman.Text = ""
Departman.Text = ""
Erkek.Value = False
Kadin.Value = False
Ýngilizce.Value = False
Fransýzca.Value = False
Almanca.Value = False
AdiveSoyadi.SetFocus


End Sub

'' Açýlýr Ekrana Veri ekledi

Private Sub UserForm_Activate()
Departman.AddItem "Yönetim"
Departman.AddItem "Muhasebe"
Departman.AddItem "Üretim"
Departman.AddItem "Pazarlama"
Departman.AddItem "Ýnsan Kaynaklarý"

'AdiveSoyadi.Text = "Adý ve Soyadý Giriniz"
'VeriGirisFormu.BackColor = vbRed
End Sub

Sub veri_giris_formu_ac()
VeriGirisFormu.Show vbModeless
End Sub
