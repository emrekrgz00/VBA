Attribute VB_Name = "Module1"
Sub makro1()
    Dim sayi As Integer
    Cells.Clear
    For sayi = 50 To 1 Step -1
        Cells(sayi, 1).Select
        Cells(sayi, 1).Value = sayi ^ 2
    Next sayi
End Sub
Sub makro2()
For i = 50 To 1 Step -1
    Range("A" & i).Select
    Range("A" & i).EntireRow.Delete
    
Next
End Sub

Sub makro3()
Dim i As Integer
Dim j As Integer
Cells.Clear
For i = 1 To 50
    For j = 1 To 10
        Cells(i, j).Value = "Bu hücre-> Satýr " & i & " Sütun " & j
    Next j
Next i


End Sub





Sub fordongusu()
    Dim i As Integer
    Cells.Clear
    For i = 1 To 10 Step 1
        Cells(i, 1) = i 'i nin her yeni deðerinde ilgili hücreye i deðerini yazdýr
        Next i 'i nin deðerini 2 arttýr taki i deðeri 10 olana kadar
End Sub
Sub icicefordongusu()
    Dim i As Integer
    Dim s As Integer
    Cells.Clear
    For s = 1 To 3 's nin deðerini 1 ver
        For i = 1 To 10
            Cells(i, s) = i 'i nin her yeni deðerinde ilgili hücreye i deðerini yazdýr, sütun numaram s=1 (ilk döngüde)
        Next i 'i nin deðerini 1 arttýr taki i deðeri 10 olana kadar
    Next s 's'nin deðerini 1 arttýr
End Sub
Sub foreachdongusu()
    Dim hucre As Range
    Dim i As Integer
    Cells.Clear
    For Each hucre In Range("A1:D20")
        i = i + 1
        hucre.Value = i
    Next hucre
End Sub
Sub makro4()
    Dim i As Integer
    i = 1
    Cells.Clear
    Do
        Cells(i, 1).Value = i
        i = i + 1
    Loop Until i > 10


End Sub




Sub dowhiledongusu()
    Dim i As Integer
    Cells.Clear
    i = 1 'baslangýc i degeri verilmeli
    Do While i < 11 ' i 11 den küçük olduðu sürece döngüyü çalýþtýr
        Cells(i, 1).Value = i ' hücreye i deðerini yazdýr
        i = i + 1 'i yi 1 arttýr
    Loop
End Sub
Sub dountildongusu()
    Dim i As Integer
    Cells.Clear
    i = 1 'baslangýc i degeri verilmeli
    Do Until i >= 10   ' i 10 dan büyük olana kadar döngüyü çalýþtýr.
        Cells(i, 1).Value = i ' hücreye i deðerini yazdýr
        i = i + 1 'i yi 1 arttýr
    Loop
End Sub
Sub dowhiledongusuu()
    Dim i As Integer
    Cells.Clear
    i = 1
    Do
        Cells(i, 1).Value = i
        i = i + 1
    Loop While i < 11
End Sub
Sub dountildongusu1()
    Dim i As Integer
    Cells.Clear
    i = 1
    Do
        Cells(i, 1).Value = i
        i = i + 1
    Loop Until i > 10
End Sub

Sub while_wend()
    Dim i As Integer
    Cells.Clear
    i = 1
    While i <= 10
        Cells(i, 1).Value = i
        i = i + 1
    Wend
End Sub
 


Sub makro6()
Cells.Clear
Dim hucre As Range
Dim i As Integer
For Each hucre In Range("A1:E10")
    i = i + 1
    hucre.Value = i
    hucre.Interior.Color = RGB(i * 1, 0, i * 5.3)
Next hucre
End Sub





























