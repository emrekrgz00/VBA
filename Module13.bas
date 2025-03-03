Attribute VB_Name = "Module13"
'Daha önceki derslerde hücreleri tanýmlayan nesne ve özelliklerden bahsedildi. _
Þimdiki bölümde excel sayfalarý ile iþlem yapmayý öðreneceðiz.
'Excelde sayfa tanýmlayan ifade "Sheets" ifadesidir. Aþaðýdaki örnekleri inceleyiniz:

Sub makro_18()
    Cells.Clear
    Sheets(1).Select
    Range("A1").Value = "Þu anda birinci sayfada ve A1 hücresindesiniz."
    Sheets("Makro2").Select
    Range("A2").Value = "Þu anda birinci sayfada ve A2 hücresindesiniz."
End Sub

'Görüldüðü gibi sayfa tanýmlama iþleminde önce "Sheets" yazýldý ve parantez içerisine bir sayý giridi.
'Burada parantez içerisine sayý grilmesi, soldan saða doðru kaçýncý sayfanýn kullanýlacaðý anlaýna gelir. _
Her zaman en soldaki sayfa ilk sayfadýr ve saða doðru diðer sayfalar gelir.

'Parantez içerisine sadece sayfa numarasý deðil, ayný zamanda sayfanýn adý da yazýlabilir.
Sub sayfa_islemleri2()
    Sheets("Sayfa1").Select
End Sub

'"Sheets" ifadesi tek baþýna yazýlýp yanýna herhangi bir parantez getirilmezse, tüm sayfalar kastedilmiþ olur.

'"Sheets" ile birlikte kullanýlabilecek özellikler: _
.select: Ýlgili sayfayý seçer. _
.Add: Yeni bir sayfa ekler. Sadece "Sheets.Add" þeklinde kullanýlýr ve parametreleri de vardýr. _
.Count: Mevcut durumda kaç adet sayfanýn var olduðunu gösterir. _
.Delete: Ýlgili sayfayý siler. _
.Visible: Ýlgili sayfanýn görünmesini veya gizlenmesini saðlar. Örneðin Sheets("Sayfa1").Visible=False _
ifadesi "Sayfa1" adýndaki sayfayý gizlerken Sheets("Sayfa1").Visible=True ifadesi ise sayfanýn görünür olmasýný saðlar. _
.Name: Ýlgili sayfanýn adýný getirir veya sayfanýn adýnýn deðiþtirilmesini saðlar.

'Bu parametrelerin yanýnda ".Range()", ".Cells()" gibi parametreler de kullanýlabilir. Bu parametreler, _
ilgili sayfanýn içerisindeki hücreleri kastetmektedir.

'ActiveSheet ifadesi mevcut içinde bulunduðumuz sayfayý kasteder.
'Sayfalara ayrýca isim de verilebilir. Project penceresindeyken sayfayý seçip Properties penceresinden adý deðiþtirilebilir. _
Böylece sayfa adý kullanýlarak doðrudan iþlemler yapýlabilir.
'Sayfalara isim vermenin bir diðer avantajý ise, bazý kullanýcýlarýn Excel dosyasýndayken sayfa adýný deðiþtirebilmesidir. _
Bu durumda biz sayfaya makro ekranýndayken farklý bir isim verirsek bu durumda Excel dosyasýnda görünen ismi deðiþse de problem çýkmayacaktýr. Konu ile ilgili olarak ilerleyen örnekler incelenebilir.
Sub makro_19()
MsgBox "Sheet Sayýsý:" & Sheets.Count & " adet sayfa var"

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
MsgBox "Þu anda var olan sayfa sayýsý: " & Sheets.Count & " ayrýca ilk sayfanýn adý ise: " & Sheets(1).Name
Sheets.Add , Sheets(Sheets.Count)
Sheets(Sheets.Count).Name = "Yeni bir sayfa"
Sheets(1).Select
Sheets(1).Name = "INDEX"
For i = 1 To 30
    Sheets.Add , Sheets(Sheets.Count)
Next i
MsgBox "Þu anda toplam " & Sheets.Count & " adet sayfa var!"

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


