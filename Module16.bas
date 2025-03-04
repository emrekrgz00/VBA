Attribute VB_Name = "Module16"
'Excel çalýþma kitabý için "Workbooks" nesnesi kullanýlýr. _
Bu nesne ile birlikte kullanýlan parametreler þu þekildedir: _
1. ".Add" ile yeni bir çalýþma kitabý açýlýr. Sadece "Workbooks.Add" þeklinde bullanýlýr. _
2. ".Open" ile mevcut çalýþma kitaplarýndan bir tanesi açýlýr. Örneðin: _
Workbooks.Open "1_Genel Bilgiler.xlsm" gibi. Eðer açacaðýnýz öçalýþma kitabý baþka bir klasörün içerisinde ise klasörün yolunu da gösteriniz. Örneðin: _
Workbooks.Open "C:\Users\ABC\1_Genel Bilgiler.xlsm" gibi. _
3. ".Save" çalýþma kitabýný kaydeder. Örneðin: Workbooks("1_Genel Bilgiler.xlsm").Save veya Workbooks(3).Save, ThisWorkBook.Save gibi. Burada dikkat edilmesi gereken husus ise, _
kaydediyor olduðunuz çalýþma kitabýnýn o an açýk olmasý gerekmektedir. _
4. ".SaveAs" ise yine ".Save" gibi çalýþmakta olup çalýþma kitabýnýn farklý bir þekilde kaydedilmesi için kullanýlýr. _
5. ".SaveCopyAs" ifadesi ise, çalýþma kitabýnýn bir kopyasýnýn alýnmasýný saðlar. ".SaveAs" den farklý olarak sadece bir kitabýn copy-paste iþlemi olarak da düþünülebilir. _
6. ".Count" mevcut durumda kaç tane çalýþma kitabýnýn açýk olduðunu gösterir. _
7. ".Close" açýk olan bir çalýþma kitabýnýn kapatýlmasýný saðlar. Çeþitli parametreleri de vardýr. _
8. ".Name" bir çalýþma kitabýnýn adýný verir. _
9. ".Activate" ilgili kitabý seçer. ".Select" komutu kullanýlamaz.


'NOT: "ThisWorkbook" nesnesi, mevcut durumda üzerinde çalýþýyor olduðumuz çalýþma kitabý anlamýna gelir. Yani, makroyu hangi çalýþma kitabýnýn içerisine yazdýysak, o çalýþma kitabýdýr.
'NOT: "ActiveWorkbook" ifadesi, þu anda hangi kitapta iþlem yapýyorsak o kitabý kasteder.


'Set ifadesi doðrudan atama iþlemleri için yapýlýr.

Sub makro_24()

'MsgBox Workbooks(1).Name
'MsgBox ThisWorkbook.Name

'Workbooks.Add
MsgBox "Þu anda açýk olan Excel Sayfa sayýsý" & " " & Workbooks.Count
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

