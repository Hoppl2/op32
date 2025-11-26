Attribute VB_Name = "modEan"
Option Explicit

Dim EanA(10) As String
Dim EanB(10) As String
Dim EanC(10) As String
Dim EanTab(10, 10) As Integer
Dim CodeTab As String

Dim SyncMark As String
Dim CenterMark As String

Private Const DefErrModul = "EAN.BAS"

Sub EanDruck(Picture1 As PictureBox, Stellen As Integer, Nummer As String)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("EanDruck")
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
If (Err.Number = (999 + vbObjectError)) Then End
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
'On Error GoTo 0
'Err.Raise 999 + vbObjectError, "DSK.VBP", "Fehler in DLL"
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim CodeAnf As Integer
Dim Pruefsumme As Integer
Dim tpx As Long
Dim x As Long
Dim EanCode As String

If CodeTab = "" Then
  CodeTab = "111111112122112212112221121122122112122211121212121221122121"
  For i = 0 To 9
    For j = 1 To 6
      EanTab(i, j) = Val(Mid$(CodeTab, i * 6 + j, 1))
    Next j
  Next i
  SyncMark = "101"
  CenterMark = "01010"
  EanA(0) = "0001101"
  EanA(1) = "0011001"
  EanA(2) = "0010011"
  EanA(3) = "0111101"
  EanA(4) = "0100011"
  EanA(5) = "0110001"
  EanA(6) = "0101111"
  EanA(7) = "0111011"
  EanA(8) = "0110111"
  EanA(9) = "0001011"
  
  EanB(0) = "0100111"
  EanB(1) = "0110011"
  EanB(2) = "0011011"
  EanB(3) = "0100001"
  EanB(4) = "0011101"
  EanB(5) = "0111001"
  EanB(6) = "0000101"
  EanB(7) = "0010001"
  EanB(8) = "0001001"
  EanB(9) = "0010111"
  
  EanC(0) = "1110010"
  EanC(1) = "1100110"
  EanC(2) = "1101100"
  EanC(3) = "1000010"
  EanC(4) = "1011100"
  EanC(5) = "1001110"
  EanC(6) = "1010000"
  EanC(7) = "1000100"
  EanC(8) = "1001000"
  EanC(9) = "1110100"
End If

If Len(Nummer) = 8 Then Nummer = Left(Nummer, 7)
If Len(Nummer) = 13 Then Nummer = Left(Nummer, 12)
Nummer = Right(String$(12, "0") + CStr(Val(Nummer)), 12)
If Stellen <= 8 Then
  Stellen = 8
  CodeAnf = 6
Else
  Stellen = 12
  CodeAnf = 1
End If


Pruefsumme = 0
For i = CodeAnf To 12
  Pruefsumme = Pruefsumme + Val(Mid$(Nummer, i, 1)) * (1 + 2 * ((i + 1) Mod 2))
Next i
Nummer = Nummer + Right$(CStr(10 - (Pruefsumme Mod 10)), 1)

If Stellen = 8 Then Nummer = Right(Nummer, Stellen)

EanCode = ""

EanCode = EanCode + SyncMark
j = Val(Left$(Nummer, 1))
For i = 1 To Stellen / 2
  k = Val(Mid$(Nummer, i - (Stellen <> 8), 1))
  If EanTab(j, i) = 1 Or Stellen = 8 Then
    EanCode = EanCode + EanA(k)
  Else
    EanCode = EanCode + EanB(k)
  End If
Next i
EanCode = EanCode + CenterMark
For i = 1 To Stellen / 2
  k = Val(Mid$(Nummer, Stellen / 2 + i - (Stellen <> 8), 1))
  EanCode = EanCode + EanC(k)
Next i
EanCode = EanCode + SyncMark
  
tpx = Screen.TwipsPerPixelX
Picture1.Width = (2 + Len(EanCode)) * tpx + (Picture1.Width - Picture1.ScaleWidth)
  
Picture1.BackColor = vbWhite
Picture1.Cls
Picture1.DrawWidth = 1
x = 2 * tpx '10 * tpx
For i = 1 To Len(EanCode)
  If Mid(EanCode, i, 1) = "1" Then
    Picture1.Line (x, 0)-(x, Picture1.ScaleHeight), vbBlack
  End If
  x = x + Picture1.DrawWidth * tpx
Next i

'With Printer
'    tpx = Screen.TwipsPerPixelX
'    .DrawWidth = 1
'    x = 2 * tpx '10 * tpx
'    For i = 1 To Len(EanCode)
'      If Mid(EanCode, i, 1) = "1" Then
'        For j = 0 To 3
'            Printer.Line (x, 0)-(x, 600), vbBlack
'            x = x + .DrawWidth * tpx
'        Next j
'      Else
'          x = x + 4 * .DrawWidth * tpx
'      End If
'    Next i
'End With

Call DefErrPop
End Sub

