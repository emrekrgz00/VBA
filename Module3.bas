Attribute VB_Name = "Module3"
'Hücrelere Veri Girme ve Hücrelerde Deðiþiklik Yapma Kodlarý
Sub Makro1()
Range("A1").Value = 1
Cells(2, 1).Value = 2
Cells(2, "A").Value = 3
Cells.Interior.Color = vbBlue ' Bütün hücreler yeþil
Cells(1, "A").Interior.Color = vbBlack
Range("A2").Interior.Color = vbRed
'Rows(1), column(1)
End Sub
'-------------------
Sub makro2()
Rows(3).Interior.Color = vbYellow
Rows(3).Value = 3
Columns(5).Interior.Color = vbGreen
Union(Rows(1), Columns(3), Range("A1:C2")).Value = 55
End Sub
'--------------------------------
'.Activate = ilgili hücreyi seçer,Aktif hale getirir,
'.Select = Hücreyi seçer,
'.Select alaný seçer, .Active bir hücreyi aktif eder,
'.Address  -- Hücrenin adresini verir,
'.Clear -- Hücrenin için temizler,
'.Coloumn -- Nesnenin kolonunu verir,
'.Row -- Nesnenin satýrýný verir,
'.Copy -- nesneyi kopyalar,
'.Cut -- nesneyi kese,
'.PasteSpecial -- nesneyi özel yapýþtýr, bir boþluk ile diðer yapýþtýrma seçenekleri açýlýr,
'.Comment -- hücrelere yazýlan yorumlar ile ilgili açýklamalarý
'.Delete -- siler Satýr ve sütunu siler,
'.End -- tablonun uçuna gider,
'.Fond.Bold - yazýnýn fontonu bold yapar
'.Hidden Satýr ve sütun gizlenemsi
'.Interior.Color  - içini boyar,
'.Formula -- formül yazma iþlemi içn kullanýlýr,
'.Offset  -- Belirli hir hücre referans alýnýr, saðýna soluna üstüne




