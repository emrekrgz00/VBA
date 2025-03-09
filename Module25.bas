Attribute VB_Name = "Module25"
Sub word_olustur()
    Dim calisan_adi As String
    Dim sicil_no As String
    Dim yillik_prim As Currency
    Dim objWord
    Dim objDoc
    Dim objSelection
    Set objWord = CreateObject("Word.Application")
   
    For i = 2 To 4744
        calisan_adi = Range("B" & i).Value
        sicil_no = Range("A" & i).Value
        yillik_prim = Range("D" & i).Value
        
        ' Word uygulamasýný baþlat
        Set objWord = CreateObject("Word.Application")
        objWord.Visible = True
            
        ' Yeni bir Word belgesi oluþtur
        Set objDoc = objWord.Documents.Add
        
       ' Seçim nesnesini al
        Set objSelection = objWord.Selection
        objSelection.TypeText "Sayýn " & calisan_adi & ", bu yýlki priminiz " & yillik_prim & " TL'dir."
        
        
        ' Dosyayý kaydet
        objDoc.SaveAs "C:\Users\Emre Karagöz\Desktop\Deneme\" & sicil_no & ".docx"
        
        ' Belgeyi kapat
        objDoc.Close False
        
        ' Word uygulamasýný kapat
        objWord.Quit
                
         ' Nesneleri serbest býrak
        Set objSelection = Nothing
        Set objDoc = Nothing
        Set objWord = Nothing
        
        
        
        'MsgBox calisan_adi
        'MsgBox sicil_no
        'MsgBox yillik_prim
        
        'Set objWord = CreateObject("Word.Application")
        'Set objDoc = objWord.Documents.Add
        'objDoc = objWord.Documents.Add
        
        'objWord.Visible = True
        'Set objSelection = objWord.Selection
        'objSelection.TypeText ("Sayin" & " " & calisan_adi & " " & ", bu yýlki priminiz" & " " & yillik_prim & " Tl'dir.")
        'objDoc.SaveAs ("C:\Users\Emre Karagöz\Desktop\Deneme\" & sicil_no & ".doc")
        'SetobjDoc.Close
        
    Next i
End Sub

' i = 50'de bitti.
