Attribute VB_Name = "Module3"
'H�crelere Veri Girme ve H�crelerde De�i�iklik Yapma Kodlar�
Sub Makro1()
Range("A1").Value = 1
Cells(2, 1).Value = 2
Cells(2, "A").Value = 3
Cells.Interior.Color = vbBlue ' B�t�n h�creler ye�il
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
'.Activate = ilgili h�creyi se�er,Aktif hale getirir,
'.Select = H�creyi se�er,
'.Select alan� se�er, .Active bir h�creyi aktif eder,
'.Address  -- H�crenin adresini verir,
'.Clear -- H�crenin i�in temizler,
'.Coloumn -- Nesnenin kolonunu verir,
'.Row -- Nesnenin sat�r�n� verir,
'.Copy -- nesneyi kopyalar,
'.Cut -- nesneyi kese,
'.PasteSpecial -- nesneyi �zel yap��t�r, bir bo�luk ile di�er yap��t�rma se�enekleri a��l�r,
'.Comment -- h�crelere yaz�lan yorumlar ile ilgili a��klamalar�
'.Delete -- siler Sat�r ve s�tunu siler,
'.End -- tablonun u�una gider,
'.Fond.Bold - yaz�n�n fontonu bold yapar
'.Hidden Sat�r ve s�tun gizlenemsi
'.Interior.Color  - i�ini boyar,
'.Formula -- form�l yazma i�lemi i�n kullan�l�r,
'.Offset  -- Belirli hir h�cre referans al�n�r, sa��na soluna �st�ne




