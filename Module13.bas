Attribute VB_Name = "Module13"
'Daha �nceki derslerde h�creleri tan�mlayan nesne ve �zelliklerden bahsedildi. _
�imdiki b�l�mde excel sayfalar� ile i�lem yapmay� ��renece�iz.
'Excelde sayfa tan�mlayan ifade "Sheets" ifadesidir. A�a��daki �rnekleri inceleyiniz:

Sub makro_18()
    Cells.Clear
    Sheets(1).Select
    Range("A1").Value = "�u anda birinci sayfada ve A1 h�cresindesiniz."
    Sheets("Makro2").Select
    Range("A2").Value = "�u anda birinci sayfada ve A2 h�cresindesiniz."
End Sub

'G�r�ld��� gibi sayfa tan�mlama i�leminde �nce "Sheets" yaz�ld� ve parantez i�erisine bir say� giridi.
'Burada parantez i�erisine say� grilmesi, soldan sa�a do�ru ka��nc� sayfan�n kullan�laca�� anla�na gelir. _
Her zaman en soldaki sayfa ilk sayfad�r ve sa�a do�ru di�er sayfalar gelir.

'Parantez i�erisine sadece sayfa numaras� de�il, ayn� zamanda sayfan�n ad� da yaz�labilir.
Sub sayfa_islemleri2()
    Sheets("Sayfa1").Select
End Sub

'"Sheets" ifadesi tek ba��na yaz�l�p yan�na herhangi bir parantez getirilmezse, t�m sayfalar kastedilmi� olur.

'"Sheets" ile birlikte kullan�labilecek �zellikler: _
.select: �lgili sayfay� se�er. _
.Add: Yeni bir sayfa ekler. Sadece "Sheets.Add" �eklinde kullan�l�r ve parametreleri de vard�r. _
.Count: Mevcut durumda ka� adet sayfan�n var oldu�unu g�sterir. _
.Delete: �lgili sayfay� siler. _
.Visible: �lgili sayfan�n g�r�nmesini veya gizlenmesini sa�lar. �rne�in Sheets("Sayfa1").Visible=False _
ifadesi "Sayfa1" ad�ndaki sayfay� gizlerken Sheets("Sayfa1").Visible=True ifadesi ise sayfan�n g�r�n�r olmas�n� sa�lar. _
.Name: �lgili sayfan�n ad�n� getirir veya sayfan�n ad�n�n de�i�tirilmesini sa�lar.

'Bu parametrelerin yan�nda ".Range()", ".Cells()" gibi parametreler de kullan�labilir. Bu parametreler, _
ilgili sayfan�n i�erisindeki h�creleri kastetmektedir.

'ActiveSheet ifadesi mevcut i�inde bulundu�umuz sayfay� kasteder.
'Sayfalara ayr�ca isim de verilebilir. Project penceresindeyken sayfay� se�ip Properties penceresinden ad� de�i�tirilebilir. _
B�ylece sayfa ad� kullan�larak do�rudan i�lemler yap�labilir.
'Sayfalara isim vermenin bir di�er avantaj� ise, baz� kullan�c�lar�n Excel dosyas�ndayken sayfa ad�n� de�i�tirebilmesidir. _
Bu durumda biz sayfaya makro ekran�ndayken farkl� bir isim verirsek bu durumda Excel dosyas�nda g�r�nen ismi de�i�se de problem ��kmayacakt�r. Konu ile ilgili olarak ilerleyen �rnekler incelenebilir.
Sub makro_19()
MsgBox "Sheet Say�s�:" & Sheets.Count & " adet sayfa var"

    For i = 1 To Sheets.Count
        Sheets(i).Select
        Sheets(i).Name = "Sayfa" & i
    Next i
End Sub
Sub makro_20()
'.ADD - before - after - adet
Sheets.Add , Sheets(1), 20
End Sub
Sub makro_21()
Application.DisplayAlerts = False
For i = 2 To Sheets.Count
    Sheets(2).Select
    Sheets(2).Delete

Next i
Application.DisplayAlerts = True
End Sub
' --------------------------------------------------------------------------

'''''''''''''''''' ----------------------------------------------------------
Sub sayfa_islemleri_3()
MsgBox "�u anda var olan sayfa say�s�: " & Sheets.Count & " ayr�ca ilk sayfan�n ad� ise: " & Sheets(1).Name
Sheets.Add , Sheets(Sheets.Count)
Sheets(Sheets.Count).Name = "Yeni bir sayfa"
Sheets(1).Select
Sheets(1).Name = "INDEX"
For i = 1 To 30
    Sheets.Add , Sheets(Sheets.Count)
Next i
MsgBox "�u anda toplam " & Sheets.Count & " adet sayfa var!"

'Application.DisplayAlerts = False
For i = Sheets.Count To 25 Step -1
    Sheets(i).Visible = False
Next i

For i = 1 To 31
    Sheets(2).Delete
Next i
'Application.DisplayAlerts = True
0
End Sub


