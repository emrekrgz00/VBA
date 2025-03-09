Attribute VB_Name = "Module26"
'MAil gönderme

Sub Mail_send()
    Dim OutApp As Object
    Dim OutMail As Object
    Set OutApp = CreateObject("Outlook.Application")
    For i = 2 To 45
        Set OutMail = OutApp.CreateItem(0)
        On Error Resume Next
        With OutMail
            .To = Range("E" & i).Value
            .CC = ""
            .BCC = ""
            .Subject = "Yýllýk Hakediþiniz"
            .Body = "Sayýn" & " " & Range("B" & i).Value & " " & ", bu yýlki hakediþiniz" & " " & Range("D" & i) & " " & " Tl'dir. !"
            .Attachments.Add "C:\Users\Emre Karagöz\Desktop\Deneme\" & Range("A" & i).Value & ".docx"
            .Display
            .Send
        End With
        On Error GoTo 0
        Set OutMail = Nothing
        Set OutApp = Nothing
    Next i
End Sub

