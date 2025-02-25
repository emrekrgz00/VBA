Attribute VB_Name = "Module5"
Sub Makro_4()
Range("A1:C1").EntireColumn.AutoFit  'Bu aralýðý seç oto fit yap
Cells(1, 1).Value = UCase(Cells(1, 1).Value) ' Ucase büyük harf yapar
Cells(1, 2).Value = UCase(Cells(1, 2).Value) ' Ucase büyük harf yapar
Cells(1, 3).Value = UCase(Cells(1, 3).Value) ' Ucase büyük harf yapar
                                             'LCASE küçük harf yapar
Cells(1, 3).Value = Application.WorksheetFunction.Proper(Cells(1, 3).Value)  'Ýlk harfleri büyük yap.
Range("A8:C8").Copy
'Range("A60000").End (xlUp) ' 60000.Satýrdan yukarý ilk dolu olan hücre
Range("A60000").End(xlUp).Offset(1, 0).PasteSpecial 'Ýlk dolu olan hücrenin bir satýr altýna
Range("A60000").End(xlUp).EntireRow.Delete 'En altta dolu olan ilk satýr sil


End Sub
