Attribute VB_Name = "Module2"
' Kod yazma Mant�k �rg�s� -- Adreslemeler ve �zellik Tan�mlama
'F5 Makro otomatik �al��t�r.
'F8 Makro sat�r sat�r �al��t�r.

Sub A1H�creyaz()
Range("A1").Value = 12  'RAnge ile h�cre tan�t�yorum.
Range("A2").Value = Range("A1").Value ^ 2
End Sub

Sub A1C12boya()
Range("A1:C12").Interior.Color = vbRed
Range("A1:C12").Value = "Merhaba"
Range("A1:C12").Font.Color = vbWhite
End Sub

