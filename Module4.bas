Attribute VB_Name = "Module4"
Sub makro_3()
Cells.Clear 'T�m h�cre sil
Range("A1").Select 'H�cre se�
Selection.Interior.Color = vbRed 'se�ilen h�cre boyama
Selection.Value = "�u anda buradas�n."
Selection.Offset(4, 5).Activate  '4 sat�r a�a��s� 5 s�tun yan�
Selection.Value = "Kral Emre"  'Select birden fazla h�creyi se�er activate tek h�creyi aktif eder.
ActiveCell.Value = "B�yle de olur."
Range("A1").Copy  ' H�creyi se�medi�im halde kopyalad�m
Range("A1").Offset(2, 0).PasteSpecial ' Offset ile hareket et kopyalad���n� yap��t�r.
Range("A3").EntireRow.Delete 'B�t�n sat�r� sil
Rows(3).Delete  'Ayn� bir �st�n
Range("A3").EntireColumn.Delete
Columns(3).Delete
' cells(3,5) == Range("E3")
Cells(3, "D").Cut
Cells(1, 1).Select
ActiveSheet.Paste 'Cut yaparsak Pastespecial olmaz, Activesheet.Paste olmas� gerekiyor.
Range("A1").AddComment "Merhaba, Emre"
'MsgBox Range("A1").Comment.Text 'Mesaj verdi yorumu
Range("A1").Comment.Visible = True
Rows(5).Interior.Color = vbGreen
Rows(5).Hidden = True 'Ya sat�r ya s�tun gizlenir, h�cre gizlenmez
Range("A1").Comment.Visible = False
Rows.Hidden = False
Rows(5).Hidden = False ' Sadece 5.Sat�r

End Sub

