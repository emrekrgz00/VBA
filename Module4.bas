Attribute VB_Name = "Module4"
Sub makro_3()
Cells.Clear 'Tüm hücre sil
Range("A1").Select 'Hücre seç
Selection.Interior.Color = vbRed 'seçilen hücre boyama
Selection.Value = "Þu anda buradasýn."
Selection.Offset(4, 5).Activate  '4 satýr aþaðýsý 5 sütun yaný
Selection.Value = "Kral Emre"  'Select birden fazla hücreyi seçer activate tek hücreyi aktif eder.
ActiveCell.Value = "Böyle de olur."
Range("A1").Copy  ' Hücreyi seçmediðim halde kopyaladým
Range("A1").Offset(2, 0).PasteSpecial ' Offset ile hareket et kopyaladýðýný yapýþtýr.
Range("A3").EntireRow.Delete 'Bütün satýrý sil
Rows(3).Delete  'Ayný bir üstün
Range("A3").EntireColumn.Delete
Columns(3).Delete
' cells(3,5) == Range("E3")
Cells(3, "D").Cut
Cells(1, 1).Select
ActiveSheet.Paste 'Cut yaparsak Pastespecial olmaz, Activesheet.Paste olmasý gerekiyor.
Range("A1").AddComment "Merhaba, Emre"
'MsgBox Range("A1").Comment.Text 'Mesaj verdi yorumu
Range("A1").Comment.Visible = True
Rows(5).Interior.Color = vbGreen
Rows(5).Hidden = True 'Ya satýr ya sütun gizlenir, hücre gizlenmez
Range("A1").Comment.Visible = False
Rows.Hidden = False
Rows(5).Hidden = False ' Sadece 5.Satýr

End Sub

