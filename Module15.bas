Attribute VB_Name = "Module15"
Sub link_oluþtur()
Attribute link_oluþtur.VB_ProcData.VB_Invoke_Func = " \n14"
'
' link_oluþtur Makro
'

'
For i = 2 To 192
    Range("A" & i).Select
    adi = Range("A" & i).Value
    
    
    ActiveSheet.Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:= _
        "'" & adi & "'!A1", TextToDisplay:=adi
Next i
End Sub
'---------------------------------------------------------------
'Sub Buton_kopyala()
'
' Buton_kopyala Makro
'
'
'    ActiveSheet.Shapes.Range(Array("Rectangle 1")).Select
'    Selection.Copy
'    Selection.Copy
'    Sheets("Abdullah Kaya").Select
'    Range("E1").Select
'    ActiveSheet.Paste
'    Range("F7").Select
'End Sub

' -----------------------
Sub Buton_kopyala()

    ActiveSheet.Shapes.Range(Array("Rectangle 1")).Select
    Selection.Copy
    
    For i = 2 To Sheets.Count
    
    Sheets(i).Select
    Range("G1").Select
    ActiveSheet.Paste
    
    Next i
End Sub
