Attribute VB_Name = "Module7"
Sub makro_8()
    Dim sayi As Integer '+-32565
    Cells.Clear
'    For sayi = 1 To 20
    For sayi = 1 To 20 Step 1 '2þer 2 þer --- -1 negatif olarak da çalýþýr
        Cells(sayi, 1).Select
        Cells(sayi, 1).Value = sayi ^ 2
    Next sayi
End Sub
'--------------------------------------------------
Sub makro_9()
For i = 20 To 1 Step -1
    Cells(i, 1).Select
    Cells(i, 1).EntireRow.Delete
Next
End Sub
'----------------------------
Sub makro_10()
For i = 15 To 1 Step -1
    Range("A" & i).Select
    Range("A" & i).EntireRow.Delete
Next
End Sub
'-------------------------
Sub makro_11()
Dim i As Integer
Dim j As Integer
Cells.Clear
For i = 1 To 20
    For j = 1 To 10
        Cells(i, j) = "Satýr" & i & "Sütun" & j
    Next j
Next i
End Sub
