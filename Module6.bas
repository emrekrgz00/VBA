Attribute VB_Name = "Module6"
''''''' Ko�ullar
Sub Makro_5()
If Range("B1").Value > 50 Then
    MsgBox "Girilen Say� 50'den b�y�k"
ElseIf Range("B1").Value < 50 Then
    MsgBox "Girilen say� 50'den k���k"
'ElseIf Range("B1").Value = 50 Then
'    MsgBox "Girilen say� 50'ye e�it"
Else
    MsgBox "50'ye e�it."

End If
End Sub

Sub Makro_6()
If Range("B1").Value < 0 Then
    MsgBox "Girilen Say� 0'dan K���k"
ElseIf Range("B1").Value < 50 Then
    MsgBox "Girilen say� 50'den k���k"
ElseIf Range("B1").Value < 100 Then
    MsgBox "Girilen say� 100'den k���k"
ElseIf Range("B1").Value < 1000 Then
    MsgBox "Girilen say� 1000'den k���k"
Else
    MsgBox "1000'den b�y�k veya e�it."

End If
End Sub

Sub Makro_7()
Select Case Range("B1").Value
Case Is < 0
    MsgBox "Girilen Say� 0'dan K���k"
Case Is < 50
    MsgBox "Girilen say� 50'den k���k"
Case Is < 100
    MsgBox "Girilen say� 100'den k���k"
Case Is < 1000
    MsgBox "Girilen say� 1000'den k���k"
Case Else
    MsgBox "1000'den b�y�k veya e�it."
End Select

End Sub

