Attribute VB_Name = "Module8"
'Do Loop - sonsuza kadar döner
'While ile sýnýr koymalýyým
Sub makro_12()
    Dim i As Integer
    i = 1
    Cells.Clear
'    Do While i <= 15
'        Cells(i, 1).Value = i
'        i = i + 1
'        Loop

    Do
        Cells(i, 1).Value = i
        i = i + 1
    Loop While i <= 10

End Sub
'''''''''''''''''''''''''''''''''''''''''''''''
Sub makro_13()
    Dim i As Integer
    i = 1
    Cells.Clear
    Do Until i > 10
        Cells(i, 1).Value = i
        i = i + 1
    Loop


End Sub
''''''''''''''''''''''''''''''''''''''''''''''
Sub makro_14()
    Dim i As Integer
    Cells.Clear
    i = 1
    While i <= 10
        Cells(i, 1).Value = i
        i = i + 1
    Wend
End Sub
'''''''''''''''''
Sub makro_15()
Cells.Clear
Dim hucre As Range
Dim i As Integer
For Each hucre In Range("A1:C10")
    i = i + 1
    hucre.Value = i
    hucre.Interior.Color = RGB(i * 4, 0, i * 5 * 3)
Next hucre

End Sub































