Attribute VB_Name = "Module26"
'MAil g�nderme

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
            .Subject = "Y�ll�k Hakedi�iniz"
            .Body = "Say�n" & " " & Range("B" & i).Value & " " & ", bu y�lki hakedi�iniz" & " " & Range("D" & i) & " " & " Tl'dir. !"
            .Attachments.Add "C:\Users\Emre Karag�z\Desktop\Deneme\" & Range("A" & i).Value & ".docx"
            .Display
            .Send
        End With
        On Error GoTo 0
        Set OutMail = Nothing
        Set OutApp = Nothing
    Next i
End Sub

