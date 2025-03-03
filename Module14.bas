Attribute VB_Name = "Module14"
Sub makro_23()
Dim adi_soyadi As String
Dim department As String
Dim brut_maas As Currency
Dim net_maas As Currency
For i = 2 To 192
    adi_soyadi = Sheets("Sayfa212").Range("A" & i).Value
    department = Sheets("Sayfa212").Range("B" & i).Value
    brut_maas = Sheets("Sayfa212").Range("C" & i).Value
    net_maas = brut_maas * 0.85
    
    Sheets.Add , Sheets(Sheets.Count)
    Sheets(Sheets.Count).Select
    Sheets(Sheets.Count).Name = adi_soyadi
    Range("A1").Value = "ÇalýþanAdi :"
    Range("B1").Value = "Department :"
    Range("C1").Value = "Brut Maas :"
    Range("D1").Value = "Net Maas :"
    

    Range("A2").Value = adi_soyadi
    Range("B2").Value = department
    Range("C2").Value = brut_maas
    Range("D2").Value = net_maas
    
    Range("A1:D2").EntireColumn.AutoFit
    
    
Next i
End Sub
