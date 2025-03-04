Attribute VB_Name = "Module16"
'Excel �al��ma kitab� i�in "Workbooks" nesnesi kullan�l�r. _
Bu nesne ile birlikte kullan�lan parametreler �u �ekildedir: _
1. ".Add" ile yeni bir �al��ma kitab� a��l�r. Sadece "Workbooks.Add" �eklinde bullan�l�r. _
2. ".Open" ile mevcut �al��ma kitaplar�ndan bir tanesi a��l�r. �rne�in: _
Workbooks.Open "1_Genel Bilgiler.xlsm" gibi. E�er a�aca��n�z ��al��ma kitab� ba�ka bir klas�r�n i�erisinde ise klas�r�n yolunu da g�steriniz. �rne�in: _
Workbooks.Open "C:\Users\ABC\1_Genel Bilgiler.xlsm" gibi. _
3. ".Save" �al��ma kitab�n� kaydeder. �rne�in: Workbooks("1_Genel Bilgiler.xlsm").Save veya Workbooks(3).Save, ThisWorkBook.Save gibi. Burada dikkat edilmesi gereken husus ise, _
kaydediyor oldu�unuz �al��ma kitab�n�n o an a��k olmas� gerekmektedir. _
4. ".SaveAs" ise yine ".Save" gibi �al��makta olup �al��ma kitab�n�n farkl� bir �ekilde kaydedilmesi i�in kullan�l�r. _
5. ".SaveCopyAs" ifadesi ise, �al��ma kitab�n�n bir kopyas�n�n al�nmas�n� sa�lar. ".SaveAs" den farkl� olarak sadece bir kitab�n copy-paste i�lemi olarak da d���n�lebilir. _
6. ".Count" mevcut durumda ka� tane �al��ma kitab�n�n a��k oldu�unu g�sterir. _
7. ".Close" a��k olan bir �al��ma kitab�n�n kapat�lmas�n� sa�lar. �e�itli parametreleri de vard�r. _
8. ".Name" bir �al��ma kitab�n�n ad�n� verir. _
9. ".Activate" ilgili kitab� se�er. ".Select" komutu kullan�lamaz.


'NOT: "ThisWorkbook" nesnesi, mevcut durumda �zerinde �al���yor oldu�umuz �al��ma kitab� anlam�na gelir. Yani, makroyu hangi �al��ma kitab�n�n i�erisine yazd�ysak, o �al��ma kitab�d�r.
'NOT: "ActiveWorkbook" ifadesi, �u anda hangi kitapta i�lem yap�yorsak o kitab� kasteder.


'Set ifadesi do�rudan atama i�lemleri i�in yap�l�r.

Sub makro_24()

'MsgBox Workbooks(1).Name
'MsgBox ThisWorkbook.Name

'Workbooks.Add
MsgBox "�u anda a��k olan Excel Sayfa say�s�" & " " & Workbooks.Count
MsgBox Workbooks(2).Name

End Sub
Sub makro_25()
    For i = 1 To 100
        Workbooks.Add
    Next i
End Sub

Sub makro_26()
    For i = 2 To Workbooks.Count
        Workbooks(2).Close
    Next i

End Sub

Sub makro_27()
    Dim kitap As Workbook
    Set kitap = Workbooks.Add
    MsgBox kitap.Name
    kitap.SaveAs "C:\Users\emre.karagoz\Desktop\yeni kitap.xlsm", 52
                                                        ' 52    .xlsm
                                                        ' 51    .xlsx
                                                        ' 50    .xlsb
                                                        ' 56    .xls (97-2003 format)
   kitap.Close True
   Set kitap = Workbooks.Open("C:\Users\emre.karagoz\Desktop\yeni kitap.xlsm")
   kitap.SaveCopyAs "C:\Users\emre.karagoz\Desktop\yeni kitap1.xlsm"
   


End Sub

