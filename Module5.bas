Attribute VB_Name = "Module5"
Sub Makro_4()
Range("A1:C1").EntireColumn.AutoFit  'Bu aral��� se� oto fit yap
Cells(1, 1).Value = UCase(Cells(1, 1).Value) ' Ucase b�y�k harf yapar
Cells(1, 2).Value = UCase(Cells(1, 2).Value) ' Ucase b�y�k harf yapar
Cells(1, 3).Value = UCase(Cells(1, 3).Value) ' Ucase b�y�k harf yapar
                                             'LCASE k���k harf yapar
Cells(1, 3).Value = Application.WorksheetFunction.Proper(Cells(1, 3).Value)  '�lk harfleri b�y�k yap.
Range("A8:C8").Copy
'Range("A60000").End (xlUp) ' 60000.Sat�rdan yukar� ilk dolu olan h�cre
Range("A60000").End(xlUp).Offset(1, 0).PasteSpecial '�lk dolu olan h�crenin bir sat�r alt�na
Range("A60000").End(xlUp).EntireRow.Delete 'En altta dolu olan ilk sat�r sil


End Sub
