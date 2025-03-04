Attribute VB_Name = "Module17"
Sub makro_28()
Dim i As Integer
Dim kitap As Workbook
Dim son_satir As Long
Application.DisplayAlerts = False
For i = 1 To 10
    Application.Calculation = xlCalculationManual
    Application.StatusBar = i
    Set kitap = Workbooks.Open("C:\Users\emre.karagoz\Desktop\Kurslarým\veriler" & i & ".xls")
    kitap.Activate
    Sheets(1).Select
    son_satir = Range("A60000").End(xlUp).Row
    Range("A2:D" & son_satir).Copy
    ThisWorkbook.Activate
    Range("A1000000").End(xlUp).Offset(1, 0).PasteSpecial
    Range("A:D").EntireColumn.AutoFit
    Range("A1").Select
    
    kitap.Close
    


Next i
Application.DisplayAlerts = True
Application.Calculation = xlCalculationAutomatic

End Sub
