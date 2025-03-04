Attribute VB_Name = "Module18"
Sub makro_29()
Dim sayi  '' As demezse variant olur

End Sub



Sub makro_30()
Dim sayi1 As Byte
Dim sayi2 As Integer
Dim sayi3 As Long
Dim sayi4 As Single
Dim karar As Boolean
Dim tarih As Date
Dim isim As String * 5
karar = True

isim = "Ali Veli Mehmet"
MsgBox isim

sayi2 = 5.3
MsgBox (sayi2)
sayi1 = 255   ' 256 hata verir
MsgBox (sayi1)
sayi3 = 55
MsgBox sayi3
tarih = 1
MsgBox tarih
tarih = DateValue("12.11.2017")
MsgBox (tarih)

tarih = "12.11.2017"
MsgBox Day(tarih)





End Sub
