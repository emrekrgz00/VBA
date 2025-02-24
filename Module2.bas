Attribute VB_Name = "Module2"
' Kod yazma Mantýk Örgüsü -- Adreslemeler ve Özellik Tanýmlama
'F5 Makro otomatik çalýþtýr.
'F8 Makro satýr satýr çalýþtýr.

Sub A1Hücreyaz()
Range("A1").Value = 12  'RAnge ile hücre tanýtýyorum.
Range("A2").Value = Range("A1").Value ^ 2
End Sub

Sub A1C12boya()
Range("A1:C12").Interior.Color = vbRed
Range("A1:C12").Value = "Merhaba"
Range("A1:C12").Font.Color = vbWhite
End Sub

