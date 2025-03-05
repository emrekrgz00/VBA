Attribute VB_Name = "Module22"
Sub kelime_arat()
MsgBox kelime_anlami("find")
End Sub

Sub makro_35()
On Error Resume Next 'Hata verse de devam eder.
Dim dizi(3)

dizi(0) = 433
dizi(1) = 43
dizi(2) = 0
dizi(3) = 41

For i = 0 To UBound(dizi)

MsgBox 5 / dizi(i)


Next i
End Sub

Sub makro_36()
Call makro_35
Call kelime_arat  ' Baþka makroyu çaðýrýr
Call Makro1
MsgBox "bitti"

'' Call Module2.A1C12boya ' module adý. orda bulunan makrolar
End Sub


