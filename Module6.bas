Attribute VB_Name = "Module6"
''''''' Koþullar
Sub Makro_5()
If Range("B1").Value > 50 Then
    MsgBox "Girilen Sayý 50'den büyük"
ElseIf Range("B1").Value < 50 Then
    MsgBox "Girilen sayý 50'den küçük"
'ElseIf Range("B1").Value = 50 Then
'    MsgBox "Girilen sayý 50'ye eþit"
Else
    MsgBox "50'ye eþit."

End If
End Sub

Sub Makro_6()
If Range("B1").Value < 0 Then
    MsgBox "Girilen Sayý 0'dan Küçük"
ElseIf Range("B1").Value < 50 Then
    MsgBox "Girilen sayý 50'den küçük"
ElseIf Range("B1").Value < 100 Then
    MsgBox "Girilen sayý 100'den küçük"
ElseIf Range("B1").Value < 1000 Then
    MsgBox "Girilen sayý 1000'den küçük"
Else
    MsgBox "1000'den büyük veya eþit."

End If
End Sub

Sub Makro_7()
Select Case Range("B1").Value
Case Is < 0
    MsgBox "Girilen Sayý 0'dan Küçük"
Case Is < 50
    MsgBox "Girilen sayý 50'den küçük"
Case Is < 100
    MsgBox "Girilen sayý 100'den küçük"
Case Is < 1000
    MsgBox "Girilen sayý 1000'den küçük"
Case Else
    MsgBox "1000'den büyük veya eþit."
End Select

End Sub

