Attribute VB_Name = "Module23"
Sub makro_37()  '' uyar�

MsgBox "Mail G�ndermeyi unutma"


End Sub

Sub alarm() ' alarm makrosu
Application.OnTime TimeValue("12:49:00"), "makro_37"
End Sub

'Sub otomatik_kaydetmeti_cagir()
'Application.OnTime Now + TimeValue("00:00:10"), "beni_kaydet"
'End Sub

'
' Otomatik Kaydet
'Sub beni_kaydet()
'ThisWorkbook.Save
'MsgBox "Kaydedildi"
'Application.OnTime Now + TimeValue("00:00:10"), "beni_kaydet"
'End Sub

Sub saat() ' Saat yapma
Range("A1").Value = Time
'Application.OnTime Now + TimeValue("00:00:01"), "saat"
End Sub


Sub dialoglar()
'' Pencere men�ler -- Dialog
'Application.Dialogs(xlDialogAlignment).Show
'Application.Dialogs(xlDialogSaveAs).Show
Application.Dialogs(xlDialogPasteSpecial).Show

End Sub

