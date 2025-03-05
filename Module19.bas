Attribute VB_Name = "Module19"
Sub makro_31()
Dim sayi(5) As Integer  '' sayi(5) dizi olduðunu excel anlar,Dizinin 6 elemaný var.
sayi(0) = 11
sayi(1) = 32
sayi(2) = 332
sayi(3) = 12
sayi(4) = 45
sayi(5) = 22

'MsgBox "Bu sayilarin toplamý:" & " " & Application.WorksheetFunction.Sum(sayi)
'MsgBox "Bu sayilarin ortalamasi:" & " " & Application.WorksheetFunction.Average(sayi)
'MsgBox "Bu sayilarin en küçüðü:" & " " & Application.WorksheetFunction.Min(sayi)
'MsgBox "Bu sayilarin en büyüðü:" & " " & Application.WorksheetFunction.Max(sayi)
'MsgBox "Dizi elemanlarnýn sonuncusu:" & " " & UBound(sayi)
'MsgBox "Dizi elemanlarnýn ilki:" & " " & LBound(sayi)

Dim yazi(1 To 3) As String
MsgBox UBound(yazi) ' Dizi hangi indekten bitiyor
MsgBox LBound(yazi) 'Dizi hangi indekste baþlýyor.
yazi(1) = "merhaba"
yazi(2) = "herkese"
yazi(3) = "Günaydýn"

MsgBox Join(yazi, "-")
End Sub

Sub makro_32()
Dim yazi As String
Dim sayilar() As String
Dim eleman

yazi = "33;22;113;555;333;222"
sayilar = Split(yazi, ";")
MsgBox sayilar(2)
MsgBox UBound(sayilar) ' Son indeks sayýsý
MsgBox "Bu dizideki eleman sayisi" & UBound(sayilar) + 1

For i = 0 To UBound(sayilar)
       MsgBox sayilar(i)
Next i

For Each eleman In sayilar
    MsgBox eleman
Next


End Sub


Sub makro_33()
Dim dizi(10, 15)
dizi(0, 0) = 3444
dizi(0, 1) = "merhaba"

For i = 0 To 10
    For j = 0 To 15
        MsgBox dizi(i, j)
        
    
    Next j
Next i
End Sub

Function TUFE_Verisi(tarih As Date)
Dim TUFE_rakami(1 To 168)
Dim TUFE_tarihi(1 To 168)

On Error Resume Next
On Error GoTo devam

TUFE_tarihi(1) = 38353
TUFE_tarihi(2) = 38384
TUFE_tarihi(3) = 38412
TUFE_tarihi(4) = 38443
TUFE_tarihi(5) = 38473
TUFE_tarihi(6) = 38504
TUFE_tarihi(7) = 38534
TUFE_tarihi(8) = 38565
TUFE_tarihi(9) = 38596
TUFE_tarihi(10) = 38626
TUFE_tarihi(11) = 38657
TUFE_tarihi(12) = 38687
TUFE_tarihi(13) = 38718
TUFE_tarihi(14) = 38749
TUFE_tarihi(15) = 38777
TUFE_tarihi(16) = 38808
TUFE_tarihi(17) = 38838
TUFE_tarihi(18) = 38869
TUFE_tarihi(19) = 38899
TUFE_tarihi(20) = 38930
TUFE_tarihi(21) = 38961
TUFE_tarihi(22) = 38991
TUFE_tarihi(23) = 39022
TUFE_tarihi(24) = 39052
TUFE_tarihi(25) = 39083
TUFE_tarihi(26) = 39114
TUFE_tarihi(27) = 39142
TUFE_tarihi(28) = 39173
TUFE_tarihi(29) = 39203
TUFE_tarihi(30) = 39234
TUFE_tarihi(31) = 39264
TUFE_tarihi(32) = 39295
TUFE_tarihi(33) = 39326
TUFE_tarihi(34) = 39356
TUFE_tarihi(35) = 39387
TUFE_tarihi(36) = 39417
TUFE_tarihi(37) = 39448
TUFE_tarihi(38) = 39479
TUFE_tarihi(39) = 39508
TUFE_tarihi(40) = 39539
TUFE_tarihi(41) = 39569
TUFE_tarihi(42) = 39600
TUFE_tarihi(43) = 39630
TUFE_tarihi(44) = 39661
TUFE_tarihi(45) = 39692
TUFE_tarihi(46) = 39722
TUFE_tarihi(47) = 39753
TUFE_tarihi(48) = 39783
TUFE_tarihi(49) = 39814
TUFE_tarihi(50) = 39845
TUFE_tarihi(51) = 39873
TUFE_tarihi(52) = 39904
TUFE_tarihi(53) = 39934
TUFE_tarihi(54) = 39965
TUFE_tarihi(55) = 39995
TUFE_tarihi(56) = 40026
TUFE_tarihi(57) = 40057
TUFE_tarihi(58) = 40087
TUFE_tarihi(59) = 40118
TUFE_tarihi(60) = 40148
TUFE_tarihi(61) = 40179
TUFE_tarihi(62) = 40210
TUFE_tarihi(63) = 40238
TUFE_tarihi(64) = 40269
TUFE_tarihi(65) = 40299
TUFE_tarihi(66) = 40330
TUFE_tarihi(67) = 40360
TUFE_tarihi(68) = 40391
TUFE_tarihi(69) = 40422
TUFE_tarihi(70) = 40452
TUFE_tarihi(71) = 40483
TUFE_tarihi(72) = 40513
TUFE_tarihi(73) = 40544
TUFE_tarihi(74) = 40575
TUFE_tarihi(75) = 40603
TUFE_tarihi(76) = 40634
TUFE_tarihi(77) = 40664
TUFE_tarihi(78) = 40695
TUFE_tarihi(79) = 40725
TUFE_tarihi(80) = 40756
TUFE_tarihi(81) = 40787
TUFE_tarihi(82) = 40817
TUFE_tarihi(83) = 40848
TUFE_tarihi(84) = 40878
TUFE_tarihi(85) = 40909
TUFE_tarihi(86) = 40940
TUFE_tarihi(87) = 40969
TUFE_tarihi(88) = 41000
TUFE_tarihi(89) = 41030
TUFE_tarihi(90) = 41061
TUFE_tarihi(91) = 41091
TUFE_tarihi(92) = 41122
TUFE_tarihi(93) = 41153
TUFE_tarihi(94) = 41183
TUFE_tarihi(95) = 41214
TUFE_tarihi(96) = 41244
TUFE_tarihi(97) = 41275
TUFE_tarihi(98) = 41306
TUFE_tarihi(99) = 41334
TUFE_tarihi(100) = 41365
TUFE_tarihi(101) = 41395
TUFE_tarihi(102) = 41426
TUFE_tarihi(103) = 41456
TUFE_tarihi(104) = 41487
TUFE_tarihi(105) = 41518
TUFE_tarihi(106) = 41548
TUFE_tarihi(107) = 41579
TUFE_tarihi(108) = 41609
TUFE_tarihi(109) = 41640
TUFE_tarihi(110) = 41671
TUFE_tarihi(111) = 41699
TUFE_tarihi(112) = 41730
TUFE_tarihi(113) = 41760
TUFE_tarihi(114) = 41791
TUFE_tarihi(115) = 41821
TUFE_tarihi(116) = 41852
TUFE_tarihi(117) = 41883
TUFE_tarihi(118) = 41913
TUFE_tarihi(119) = 41944
TUFE_tarihi(120) = 41974
TUFE_tarihi(121) = 42005
TUFE_tarihi(122) = 42036
TUFE_tarihi(123) = 42064
TUFE_tarihi(124) = 42095
TUFE_tarihi(125) = 42125
TUFE_tarihi(126) = 42156
TUFE_tarihi(127) = 42186
TUFE_tarihi(128) = 42217
TUFE_tarihi(129) = 42248
TUFE_tarihi(130) = 42278
TUFE_tarihi(131) = 42309
TUFE_tarihi(132) = 42339
TUFE_tarihi(133) = 42370
TUFE_tarihi(134) = 42401
TUFE_tarihi(135) = 42430
TUFE_tarihi(136) = 42461
TUFE_tarihi(137) = 42491
TUFE_tarihi(138) = 42522
TUFE_tarihi(139) = 42552
TUFE_tarihi(140) = 42583
TUFE_tarihi(141) = 42614
TUFE_tarihi(142) = 42644
TUFE_tarihi(143) = 42675
TUFE_tarihi(144) = 42705
TUFE_tarihi(145) = 42736
TUFE_tarihi(146) = 42767
TUFE_tarihi(147) = 42795
TUFE_tarihi(148) = 42826
TUFE_tarihi(149) = 42856
TUFE_tarihi(150) = 42887
TUFE_tarihi(151) = 42917
TUFE_tarihi(152) = 42948
TUFE_tarihi(153) = 42979
TUFE_tarihi(154) = 43009
TUFE_tarihi(155) = 43040
TUFE_tarihi(156) = 43070
TUFE_tarihi(157) = 43101
TUFE_tarihi(158) = 43132
TUFE_tarihi(159) = 43160
TUFE_tarihi(160) = 43191
TUFE_tarihi(161) = 43221
TUFE_tarihi(162) = 43252
TUFE_tarihi(163) = 43282
TUFE_tarihi(164) = 43313
TUFE_tarihi(165) = 43344
TUFE_tarihi(166) = 43374
TUFE_tarihi(167) = 43405
TUFE_tarihi(168) = 43435

TUFE_rakami(1) = 9.24
TUFE_rakami(2) = 8.69
TUFE_rakami(3) = 7.94
TUFE_rakami(4) = 8.18
TUFE_rakami(5) = 8.7
TUFE_rakami(6) = 8.95
TUFE_rakami(7) = 7.82
TUFE_rakami(8) = 7.91
TUFE_rakami(9) = 7.99
TUFE_rakami(10) = 7.52
TUFE_rakami(11) = 7.61
TUFE_rakami(12) = 7.72
TUFE_rakami(13) = 7.93
TUFE_rakami(14) = 8.15
TUFE_rakami(15) = 8.16
TUFE_rakami(16) = 8.83
TUFE_rakami(17) = 9.86
TUFE_rakami(18) = 10.12
TUFE_rakami(19) = 11.69
TUFE_rakami(20) = 10.26
TUFE_rakami(21) = 10.55
TUFE_rakami(22) = 9.98
TUFE_rakami(23) = 9.86
TUFE_rakami(24) = 9.65
TUFE_rakami(25) = 9.93
TUFE_rakami(26) = 10.16
TUFE_rakami(27) = 10.86
TUFE_rakami(28) = 10.72
TUFE_rakami(29) = 9.23
TUFE_rakami(30) = 8.6
TUFE_rakami(31) = 6.9
TUFE_rakami(32) = 7.39
TUFE_rakami(33) = 7.12
TUFE_rakami(34) = 7.7
TUFE_rakami(35) = 8.4
TUFE_rakami(36) = 8.39
TUFE_rakami(37) = 8.17
TUFE_rakami(38) = 9.1
TUFE_rakami(39) = 9.15
TUFE_rakami(40) = 9.66
TUFE_rakami(41) = 10.74
TUFE_rakami(42) = 10.61
TUFE_rakami(43) = 12.06
TUFE_rakami(44) = 11.77
TUFE_rakami(45) = 11.13
TUFE_rakami(46) = 11.99
TUFE_rakami(47) = 10.76
TUFE_rakami(48) = 10.06
TUFE_rakami(49) = 9.5
TUFE_rakami(50) = 7.73
TUFE_rakami(51) = 7.89
TUFE_rakami(52) = 6.13
TUFE_rakami(53) = 5.24
TUFE_rakami(54) = 5.73
TUFE_rakami(55) = 5.39
TUFE_rakami(56) = 5.33
TUFE_rakami(57) = 5.27
TUFE_rakami(58) = 5.08
TUFE_rakami(59) = 5.53
TUFE_rakami(60) = 6.53
TUFE_rakami(61) = 8.19
TUFE_rakami(62) = 10.13
TUFE_rakami(63) = 9.56
TUFE_rakami(64) = 10.19
TUFE_rakami(65) = 9.1
TUFE_rakami(66) = 8.37
TUFE_rakami(67) = 7.58
TUFE_rakami(68) = 8.33
TUFE_rakami(69) = 9.24
TUFE_rakami(70) = 8.62
TUFE_rakami(71) = 7.29
TUFE_rakami(72) = 6.4
TUFE_rakami(73) = 4.9
TUFE_rakami(74) = 4.16
TUFE_rakami(75) = 3.99
TUFE_rakami(76) = 4.26
TUFE_rakami(77) = 7.17
TUFE_rakami(78) = 6.24
TUFE_rakami(79) = 6.31
TUFE_rakami(80) = 6.65
TUFE_rakami(81) = 6.15
TUFE_rakami(82) = 7.66
TUFE_rakami(83) = 9.48
TUFE_rakami(84) = 10.45
TUFE_rakami(85) = 10.61
TUFE_rakami(86) = 10.43
TUFE_rakami(87) = 10.43
TUFE_rakami(88) = 11.14
TUFE_rakami(89) = 8.28
TUFE_rakami(90) = 8.87
TUFE_rakami(91) = 9.07
TUFE_rakami(92) = 8.88
TUFE_rakami(93) = 9.19
TUFE_rakami(94) = 7.8
TUFE_rakami(95) = 6.37
TUFE_rakami(96) = 6.16
TUFE_rakami(97) = 7.31
TUFE_rakami(98) = 7.03
TUFE_rakami(99) = 7.29
TUFE_rakami(100) = 6.13
TUFE_rakami(101) = 6.51
TUFE_rakami(102) = 8.3
TUFE_rakami(103) = 8.88
TUFE_rakami(104) = 8.17
TUFE_rakami(105) = 7.88
TUFE_rakami(106) = 7.71
TUFE_rakami(107) = 7.32
TUFE_rakami(108) = 7.4
TUFE_rakami(109) = 7.75
TUFE_rakami(110) = 7.89
TUFE_rakami(111) = 8.39
TUFE_rakami(112) = 9.38
TUFE_rakami(113) = 9.66
TUFE_rakami(114) = 9.16
TUFE_rakami(115) = 9.32
TUFE_rakami(116) = 9.54
TUFE_rakami(117) = 8.86
TUFE_rakami(118) = 8.96
TUFE_rakami(119) = 9.15
TUFE_rakami(120) = 8.17
TUFE_rakami(121) = 7.24
TUFE_rakami(122) = 7.55
TUFE_rakami(123) = 7.61
TUFE_rakami(124) = 7.91
TUFE_rakami(125) = 8.09
TUFE_rakami(126) = 7.2
TUFE_rakami(127) = 6.81
TUFE_rakami(128) = 7.14
TUFE_rakami(129) = 7.95
TUFE_rakami(130) = 7.58
TUFE_rakami(131) = 8.1
TUFE_rakami(132) = 8.81
TUFE_rakami(133) = 9.58
TUFE_rakami(134) = 8.78
TUFE_rakami(135) = 7.46
TUFE_rakami(136) = 6.57
TUFE_rakami(137) = 6.58
TUFE_rakami(138) = 7.64
TUFE_rakami(139) = 8.79
TUFE_rakami(140) = 8.05
TUFE_rakami(141) = 7.28
TUFE_rakami(142) = 7.16
TUFE_rakami(143) = 7
TUFE_rakami(144) = 8.53
TUFE_rakami(145) = 9.22
TUFE_rakami(146) = 10.13
TUFE_rakami(147) = 11.29
TUFE_rakami(148) = 11.87
TUFE_rakami(149) = 11.72
TUFE_rakami(150) = 10.9
TUFE_rakami(151) = 9.79
TUFE_rakami(152) = 10.68
TUFE_rakami(153) = 11.2
TUFE_rakami(154) = 11.9
TUFE_rakami(155) = 12.98
TUFE_rakami(156) = 11.92
TUFE_rakami(157) = 10.35
TUFE_rakami(158) = 10.26
TUFE_rakami(159) = 10.23
TUFE_rakami(160) = 10.85
TUFE_rakami(161) = 12.15
TUFE_rakami(162) = 15.39
TUFE_rakami(163) = 15.85
TUFE_rakami(164) = 17.9
TUFE_rakami(165) = 24.52
TUFE_rakami(166) = 25.24
TUFE_rakami(167) = 21.62
TUFE_rakami(168) = 20.3
If Year(tarih) < 2005 Then
    TUFE_Verisi = "Böyle bir tarih yoktur."
    
End If

ay = Month(tarih)
yil = Year(tarih)

For i = 1 To 168
    If yil = Year(TUFE_tarihi(i)) And ay = Month(TUFE_tarihi(i)) Then
        TUFE_Verisi = TUFE_rakami(i)
    End If

Next i

devam:
If Err Then
    MsgBox "Bu tarihlerde veriye sahip deðilim"
    TUFE_Verisi = "--"
End If
End Function




