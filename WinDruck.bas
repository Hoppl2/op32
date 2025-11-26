Attribute VB_Name = "modWinDruck"
Option Explicit

Const LF_FACESIZE = 32

'Type LOGFONT
'   lfHeight As Long
'   lfWidth As Long
'   lfEscapement As Long
'   lfOrientation As Long
'   lfWeight As Long
'   lfItalic As Byte
'   lfUnderline As Byte
'   lfStrikeOut As Byte
'   lfCharSet As Byte
'   lfOutPrecision As Byte
'   lfClipPrecision As Byte
'   lfQuality As Byte
'   lfPitchAndFamily As Byte
'   lfFaceName As String * LF_FACESIZE
'End Type

Type DOCINFO
   cbSize As Long
   lpszDocName As String
   lpszOutput As String
   lpszDatatype As String
   fwType As Long
End Type

Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Declare Function CreateDC Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, ByVal lpDeviceName As String, _
ByVal lpOutput As Long, ByVal lpInitData As Long) As Long

Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long

Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, _
ByVal lpString As String, ByVal nCount As Long) As Long ' or Boolean

Declare Function StartDoc Lib "gdi32" Alias "StartDocA" (ByVal hdc As Long, lpdi As DOCINFO) As Long
Declare Function EndDoc Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function StartPage Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function EndPage Lib "gdi32" (ByVal hdc As Long) As Long

Public Const DESIREDFONTSIZE = 10     ' Could use variable, TextBox, etc.


Public RezDruckZeilen$(40)

Public RezeptVersatzX%, RezeptVersatzY%, DatumVersatzX%, DatumVersatzY%, RezeptNrVersatzY%
Public PrivatRezeptVersatzY%
Public RezeptFont$

Public Code128VersatzX%, Code128VersatzY%
Public Code128Font$
Public Code128FontSize%
Public Code128Flag%

Public RezDruckInd%

Public Tm290WaitErg%
Public LetztDruckZeit&

Public DruckStr$

Public AnzDruckSpalten%
Public DruckSpalte() As DruckSpalteStruct
Public DruckFontSize%

Public NurHashCodeDruck%
Public HochFormatDruck%
Public PrivRezDruckHoch%
Public RezeptNrPositionAlt%

Public Pzn8Test%

Private Const DefErrModul = "WINDRUCK.BAS"

Sub InitDruckZeile(Optional ZentrierX% = False)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("InitDruckZeile")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim i%, j%, zentr%
Dim GesBreite&
Dim h$
        
For j% = 12 To 5 Step -1
    Printer.ScaleMode = vbTwips
    Printer.Font.Name = "Arial"
    Printer.Font.Size = 18  'nötig wegen Canon-BJ; sonst ab 2.Ausdruck falsch
    Printer.Font.Size = j%
        
    DruckSpalte(0).StartX = 0
    For i% = 0 To (AnzDruckSpalten% - 1)
        If (i% = 0) Then
            DruckSpalte(0).StartX = 0
        Else
            DruckSpalte(i%).StartX = DruckSpalte(i% - 1).StartX + DruckSpalte(i% - 1).BreiteX + Printer.TextWidth("  ")
        End If
        DruckSpalte(i%).BreiteX = Printer.TextWidth(RTrim(DruckSpalte(i%).TypStr))
    Next i%
    
    GesBreite& = DruckSpalte(AnzDruckSpalten% - 1).StartX + DruckSpalte(AnzDruckSpalten% - 1).BreiteX
    If (GesBreite& < Printer.ScaleWidth) Then Exit For
Next j%

DruckFontSize% = j%

If (ZentrierX%) Then
    zentr% = (Printer.ScaleWidth - GesBreite&) / 2
    For i% = 0 To (AnzDruckSpalten% - 1)
        DruckSpalte(i%).StartX = DruckSpalte(i%).StartX + zentr%
    Next i%
End If

Call DefErrPop
End Sub


Sub DruckZeile(ZeilenText$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("DruckZeile")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim i%, ind%, x%, iItalic%, iBold%
Dim h$, tx$

h$ = ZeilenText$

For i% = 0 To (AnzDruckSpalten% - 1)
    ind% = InStr(h$, vbTab)
    tx$ = Left$(h$, ind% - 1)
    h$ = Mid$(h$, ind% + 1)
    
    If (DruckSpalte(i%).Ausrichtung = "L") Then
        Do
            If (Len(tx$) <= 1) Then Exit Do
            If (Printer.TextWidth(tx$) <= DruckSpalte(i%).BreiteX) Then Exit Do
            tx$ = Left$(tx$, Len(tx$) - 1)
        Loop
        x% = DruckSpalte(i%).StartX
    Else
        x% = DruckSpalte(i%).StartX + DruckSpalte(i%).BreiteX - Printer.TextWidth(tx$)
    End If
    Printer.CurrentX = x%
    
    iItalic% = Printer.Font.Italic
    iBold% = Printer.Font.Bold
    If (DruckSpalte(i%).Attrib = 1) Then
        Printer.Font.Name = "Times New Roman"
        Printer.Font.Italic = True
    End If
    If (DruckSpalte(i%).Attrib = 2) Then
        Printer.Font.Name = "Symbol"
        Printer.Font.Bold = True
    End If
    
    Printer.Print tx$;
    
    If (DruckSpalte(i%).Attrib) Then
        Printer.Font.Name = "Arial"
        Printer.Font.Bold = iBold%
        Printer.Font.Italic = iItalic%
    End If
Next i%
Printer.Print " "
                
Call DefErrPop
End Sub

Sub RezeptDruck()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("RezeptDruck")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim i%, StartX%, StartY%, x%, y%, TaxierungX%, TaxierungY%, aFontSize%
Dim Preis&
Dim tx$, sPzn$, sIk$

'Call GetPrinterParam

RezDruckInd% = 0

Printer.ScaleMode = vbTwips
Printer.Font.Name = RezeptFont$ ' "Lucida Console"
Printer.Font.Size = 11
Printer.Font.Size = 12
StartY% = 0

x% = RezeptVersatzX% * 567& * 0.01
StartX% = Printer.TextWidth(Space$(52)) + x%
If (StartX% > Printer.ScaleWidth) Then StartX% = Printer.ScaleWidth - 567& * 0.1
StartY% = RezeptVersatzY% * 567& * 0.01

If (SonderBelegRezept%) Then
    Printer.CurrentX = x% + 20
    Printer.CurrentY = StartY% + 40
    tx$ = HoleDruckZeile$
    Printer.Print tx$;
    Printer.CurrentX = x% + 20
    Printer.CurrentY = StartY% + 500
    tx$ = HoleDruckZeile$
    Printer.Print tx$;
    Printer.CurrentX = x% + 20
    Printer.CurrentY = StartY% + 900
    tx$ = HoleDruckZeile$
    Printer.Print tx$;
    Printer.CurrentX = x% + 20
    Printer.CurrentY = StartY% + 1550
    tx$ = HoleDruckZeile$
    Printer.Print tx$;
    Printer.CurrentX = x% + 1200
    Printer.CurrentY = StartY% + 1550
    tx$ = HoleDruckZeile$
    Printer.Print tx$;
    Printer.CurrentX = x% + 3300
    Printer.CurrentY = StartY% + 1550
    tx$ = HoleDruckZeile$
    Printer.Print tx$;
    
    Printer.CurrentX = x% + 20
    Printer.CurrentY = StartY% + 2050
    tx$ = HoleDruckZeile$
    Printer.Print tx$;
    Printer.CurrentX = x% + 1800
    Printer.CurrentY = StartY% + 2050
    tx$ = HoleDruckZeile$
    Printer.Print tx$;
    Printer.CurrentX = x% + 3300
    Printer.CurrentY = StartY% + 2050
    tx$ = HoleDruckZeile$
    Printer.Print tx$;
End If

If (NurHashCodeDruck) Then
    Printer.CurrentX = StartX%
    Printer.CurrentY = StartY%
    
    If (AvpTeilnahme%) Or (FiveRxFlag) Then
        tx$ = HoleDruckZeile$
    End If
    
    tx$ = HoleDruckZeile$
    Printer.CurrentY = Printer.CurrentY + 567
    tx$ = HoleDruckZeile$
    tx$ = HoleDruckZeile$
    
    Printer.CurrentY = Printer.CurrentY + 567 * 1
    For i% = 1 To 1
        tx$ = HoleDruckZeile$
        tx$ = Trim(HoleDruckZeile$)
        tx$ = HoleDruckZeile$
        Printer.CurrentY = Printer.CurrentY + 567 * 0.85
    Next i%
    
    For i% = 2 To 3
        tx$ = HoleDruckZeile$
        If (Left$(tx$, 5) = "@BTM@") Then   'And (i% = 4) Then
            tx$ = Mid$(tx$, 6)
            If (Len(tx$) > 38) Then Printer.Font.Size = 10
            x% = Printer.TextWidth(Space$(10)) + (RezeptVersatzX% + DatumVersatzX%) * 567& * 0.01
            Printer.CurrentX = x%
        Else
            x% = StartX% - Printer.TextWidth(Space$(10))
            If (Len(tx$) > 8) Then Printer.Font.Size = 10
            Printer.CurrentX = x% - Printer.TextWidth(tx$)
        End If
        Printer.Print tx$;
        Printer.Font.Size = 12
        
        x% = StartX% - Printer.TextWidth(Space$(6))
        tx$ = Trim(HoleDruckZeile$)
        Printer.CurrentX = x% - Printer.TextWidth(tx$)
        Printer.Print tx$;
        
        x% = StartX%
        tx$ = HoleDruckZeile$
        Preis& = Val(tx$)
        If (Preis& > 99900) Then Printer.Font.Size = 10
        Printer.CurrentX = x% - Printer.TextWidth(tx$)
        Printer.Print tx$;
        Printer.Font.Size = 12
        
        If (i% <= 3) Then
            Printer.CurrentY = Printer.CurrentY + 567 * 0.85
        Else
            Printer.CurrentY = Printer.CurrentY + 567 * 0.7
        End If
    Next i%
    
ElseIf (HochFormatDruck) Then
    If (AvpTeilnahme%) Or (FiveRxFlag) Then
        tx$ = HoleDruckZeile$
    End If
    
    StartY% = PrivatRezeptVersatzY * 567& * 0.01
    Printer.CurrentX = StartX%
    Printer.CurrentY = StartY%
    
    sIk$ = HoleDruckZeile$
    tx$ = HoleDruckZeile$
    
    x% = StartX% - Printer.TextWidth(Space$(1))
    tx$ = HoleDruckZeile$
    Printer.CurrentX = x% - Printer.TextWidth(tx$) - 567 * 0.1
    Printer.Print tx$;
    
    Printer.CurrentY = Printer.CurrentY + 567 * 1
    For i% = 1 To 7
        tx$ = HoleDruckZeile$
        
        If (Left$(tx$, 9) = "@ABDATUM@") Then
            tx$ = HoleDruckZeile$
        End If
        
        If (Left$(tx$, 7) = "@REZNR@") Then
            tx$ = HoleDruckZeile$
        End If
        
        If (Left$(tx$, 6) = "@BTM2@") Then   'And (i% = 4) Then
            tx$ = HoleDruckZeile$
        End If
        If (Left$(tx$, 5) = "@BTM@") Then   'And (i% = 4) Then
            tx$ = Mid$(tx$, 6)
            If (Len(tx$) > 38) Then Printer.Font.Size = 10
            x% = Printer.TextWidth(Space$(10)) + (RezeptVersatzX% + DatumVersatzX%) * 567& * 0.01
            Printer.CurrentX = x%
        Else
            If (Pzn8Test = 0) And (Year(Now) = 2012) And (Len(tx) = 8) Then
                tx = Mid(tx, 2)
            End If
            sPzn = tx$
            x% = StartX% - Printer.TextWidth(Space$(10))
            If (Len(tx$) > 8) Then Printer.Font.Size = 10
            Printer.CurrentX = x% - Printer.TextWidth(tx$)
        End If
        Printer.Print tx$;
        Printer.Font.Size = 12
        
        x% = StartX% - Printer.TextWidth(Space$(6))
        tx$ = Trim(HoleDruckZeile$)
        Printer.CurrentX = x% - Printer.TextWidth(tx$)
        Printer.Print tx$;
        
        x% = StartX%
        tx$ = HoleDruckZeile$
        Preis& = Val(tx$)
        If (Preis& > 99900) Then Printer.Font.Size = 10
        Printer.CurrentX = x% - Printer.TextWidth(tx$)
        Printer.Print tx$;
        Printer.Font.Size = 12
        
        If (Trim(sPzn) <> "") Then
            Printer.CurrentY = Printer.CurrentY + 567 * 0.85
        End If
    Next i%
        
'    For i% = 4 To 7
'        tx$ = HoleDruckZeile$
'        tx$ = HoleDruckZeile$
'        tx$ = HoleDruckZeile$
'    Next i
    'tx$ = HoleDruckZeile$
    tx$ = HoleDruckZeile$
    
    x% = StartX%
    tx$ = Trim(sIk$)
    Printer.CurrentX = x% - Printer.TextWidth(tx$)
    Printer.Print tx$;
    
    Printer.CurrentY = Printer.CurrentY + 567 * 0.85
    x% = StartX%
    tx$ = HoleDruckZeile$
    Printer.CurrentX = x% - Printer.TextWidth(tx$)
    Printer.Print tx$;
        
    Printer.CurrentY = Printer.CurrentY + 567 * 0.85
    x% = StartX%
    tx$ = HoleDruckZeile$
    Printer.CurrentX = x% - Printer.TextWidth(tx$)
    Printer.Print tx$;
Else
    If (AvpTeilnahme%) Or (FiveRxFlag) Or (Code128Flag) Then
    'If (AvpTeilnahme%) Or (ParenteralRezept >= 0) Then
    '    Printer.Font.Size = 10
        x% = StartX%
        tx$ = HoleDruckZeile$
        If (tx = "X" + "Selbsterklaerung" + "X") Then
            tx = " "
        End If
        Printer.CurrentX = x% - Printer.TextWidth(tx$)
        y% = StartY% - Printer.TextHeight(tx$) - (567 * 0.1)
        If (y% >= 0) Then
            Printer.CurrentY = y%
            Printer.Print tx$;
        End If
    '    Printer.Font.Size = 12
    
'        If (FiveRxFlag) And (Code128Flag) Then
        If (Code128Flag) Then
            
            x% = StartX% - Printer.TextWidth(Space$(22))
            tx = code128(tx)
            
'            With frmAction.picTemp
'                .Cls
'                .BackColor = vbWhite
'                .ForeColor = vbBlack
'                .Font.Name = Code128Font
'                .Font.Size = Code128FontSize
'                .Width = .TextWidth(tx$) + 30
'                .Height = .TextHeight(tx) + 30
'                MsgBox (Str(.TextWidth(tx)) + Str(.TextHeight(tx)))
'                .CurrentX = 15
'                .CurrentY = 15
'                frmAction.picTemp.Print tx$;
'                frmAction.picTemp.PaintPicture .Image, 0, 0, 2 * .Width, .Height
'                .Visible = True
'            End With
'
'            BitBlt Printer.hdc, 0, 0, frmAction.picTemp.ScaleWidth / Printer.TwipsPerPixelX, frmAction.picTemp.ScaleHeight / Printer.TwipsPerPixelY, frmAction.picTemp.hdc, 0, 0, SRCCOPY
''            StretchBlt Printer.hdc, 0, 0, frmAction.picTemp.ScaleWidth / Printer.TwipsPerPixelX, frmAction.picTemp.ScaleHeight / Printer.TwipsPerPixelY, frmAction.picTemp.hdc, 0, 0, frmAction.picTemp.ScaleWidth / Screen.TwipsPerPixelX, frmAction.picTemp.ScaleHeight / Screen.TwipsPerPixelY, SRCCOPY
            
            aFontSize = Printer.Font.Size
            Printer.Font.Name = Code128Font
            Printer.Font.Size = Code128FontSize
'                MsgBox (Str(Printer.TextWidth(tx)) + Str(Printer.TextHeight(tx)))
            Printer.CurrentX = x% - Printer.TextWidth(tx$) + Code128VersatzX * 567& * 0.01
            Printer.CurrentY = Printer.CurrentY + Code128VersatzY * 567& * 0.01
            
            Printer.Print tx$;
            Printer.Font.Name = RezeptFont$ ' "Lucida Console"
            Printer.Font.Size = aFontSize
        End If
    End If
    
    Printer.CurrentX = StartX%
    Printer.CurrentY = StartY%
    
    x% = StartX%
    tx$ = HoleDruckZeile$
    Printer.CurrentX = x% - Printer.TextWidth(tx$)
    Printer.Print tx$;
    
    x% = StartX% - Printer.TextWidth(Space$(14))
    tx$ = HoleDruckZeile$
    If (tx$ = "0") Then tx$ = "0   "
    Printer.CurrentX = x% - Printer.TextWidth(tx$) - 567 * 0.1
    Printer.CurrentY = Printer.CurrentY + 567
    Printer.Print tx$;
    
    x% = StartX% - Printer.TextWidth(Space$(1))
    tx$ = HoleDruckZeile$
    Printer.CurrentX = x% - Printer.TextWidth(tx$) - 567 * 0.1
    Printer.Print tx$;
    
    Printer.CurrentY = Printer.CurrentY + 567 * 1
    For i% = 1 To 7
        If (i% = 5) Then
            TaxierungX% = Printer.TextWidth(Space$(10)) + (RezeptVersatzX% + DatumVersatzX%) * 567& * 0.01
            TaxierungY% = Printer.CurrentY
        End If
        
        tx$ = HoleDruckZeile$
        If (Left$(tx$, 9) = "@ABDATUM@") Then
            tx$ = Mid$(tx$, 10)
            
            x% = StartX% - Printer.TextWidth(Space$(30))
            Printer.CurrentX = x%
'            y = Printer.CurrentY
'            Printer.CurrentY = y - 567 * 0.35 + (RezeptNrVersatzY% * 567& * 0.01)
            Printer.Print tx$;
            Printer.CurrentY = y
            tx$ = HoleDruckZeile$
        End If
        
'        tx$ = HoleDruckZeile$
        If (Left$(tx$, 14) = "@SCHUTZMASKEN@") Then
            tx$ = Mid$(tx$, 15)
            x% = Printer.TextWidth(Space$(12)) + (RezeptVersatzX% + DatumVersatzX%) * 567& * 0.01
            Printer.CurrentX = x%
            y = Printer.CurrentY
            Printer.CurrentY = y - 567 * 0.35 + (RezeptNrVersatzY% * 567& * 0.01)
            Printer.Font.Size = 14
            Printer.Print tx$;
            Printer.CurrentY = y
            Printer.Font.Size = 12
            tx$ = HoleDruckZeile$
        End If
        
        If (Left$(tx$, 7) = "@REZNR@") Then
            tx$ = Mid$(tx$, 8)
            x% = Printer.TextWidth(Space$(10)) + (RezeptVersatzX% + DatumVersatzX%) * 567& * 0.01
            Printer.CurrentX = x%
            y = Printer.CurrentY
            Printer.CurrentY = y - 567 * 0.35 + (RezeptNrVersatzY% * 567& * 0.01)
            If (Len(tx) >= 18) Then
                Printer.Font.Size = 9
            ElseIf (Len(tx) >= 15) Then
                Printer.Font.Size = 10
            End If
            Printer.Print tx$;
            Printer.CurrentY = y
            Printer.Font.Size = 12
            tx$ = HoleDruckZeile$
        End If
        
        If (Left$(tx$, 6) = "@BTM2@") Then   'And (i% = 4) Then
            tx$ = Mid$(tx$, 7)
            Printer.Font.Size = 10
            x% = Printer.TextWidth(Space$(10)) + (RezeptVersatzX% + DatumVersatzX%) * 567& * 0.01
            Printer.CurrentX = x%
            
            Printer.Font.Size = 12
            x% = StartX% - Printer.TextWidth(Space$(10)) - Printer.TextWidth(String(8, "9")) - x - 100
            Printer.Font.Size = 10
            Do
                If (Len(tx) < 5) Then
                    Exit Do
                End If
                If (Printer.TextWidth(tx) < x) Then
                    Exit Do
                End If
                tx = Left(tx, Len(tx) - 1)
            Loop
            
            y = Printer.CurrentY
            Printer.CurrentY = y - 567 * 0.35 + (RezeptNrVersatzY% * 567& * 0.01)
            Printer.Print tx$;
            Printer.CurrentY = y
            Printer.Font.Size = 12
            tx$ = HoleDruckZeile$
        End If
        
        If (Left$(tx$, 5) = "@BTM@") Then   'And (i% = 4) Then
            tx$ = Mid$(tx$, 6)
            If (Len(tx$) > 38) Then Printer.Font.Size = 10
            x% = Printer.TextWidth(Space$(10)) + (RezeptVersatzX% + DatumVersatzX%) * 567& * 0.01
            Printer.CurrentX = x%
            y = Printer.CurrentY
            Printer.CurrentY = y - 567 * 0.35 + (RezeptNrVersatzY% * 567& * 0.01)
            Printer.Print tx$;
            Printer.CurrentY = y
        Else
            If (Pzn8Test = 0) And (Year(Now) = 2012) And (Len(tx) = 8) Then
                tx = Mid(tx, 2)
            End If
            x% = StartX% - Printer.TextWidth(Space$(10))
            If (Len(tx$) > 8) Then Printer.Font.Size = 10
            Printer.CurrentX = x% - Printer.TextWidth(tx$)
            Printer.Print tx$;
        End If
'        Printer.Print tx$;
        Printer.Font.Size = 12
        
        x% = StartX% - Printer.TextWidth(Space$(6))
        tx$ = Trim(HoleDruckZeile$)
        If (Val(tx) > 999) Then Printer.Font.Size = 9
        Printer.CurrentX = x% - Printer.TextWidth(tx$)
        Printer.Print tx$;
        Printer.Font.Size = 12
        
        x% = StartX%
        tx$ = HoleDruckZeile$
        Preis& = Val(tx$)
        If (Preis& > 99900) Then Printer.Font.Size = 10
        If (Preis& > 999900) Then Printer.Font.Size = 9
        If (Preis& > 9999900) Then Printer.Font.Size = 8
        If (Len(tx) = 7) Then Printer.Font.Size = 9 'Hashcode
        Printer.CurrentX = x% - Printer.TextWidth(tx$)
        Printer.Print tx$;
        Printer.Font.Size = 12
        
        If (i% <= 3) Then
            Printer.CurrentY = Printer.CurrentY + 567 * 0.85
        Else
            Printer.CurrentY = Printer.CurrentY + 567 * 0.7
        End If
    Next i%
    
    x% = StartX%
    tx$ = HoleDruckZeile$
    Call OemToChar(tx$, tx$)
    Printer.CurrentY = Printer.CurrentY - 567 * 0.7
    Printer.CurrentX = x% - Printer.TextWidth(tx$)
    Printer.Print tx$;
    
    Printer.CurrentY = StartY% + 567 * 7.3 + (DatumVersatzY% * 567& * 0.01)
    x% = Printer.TextWidth(Space$(10)) + (RezeptVersatzX% + DatumVersatzX%) * 567& * 0.01
    If (x% < 0) Then x% = 0
    tx$ = HoleDruckZeile$
    Call OemToChar(tx$, tx$)
    If (Left$(tx$, 11) = "@PRIVDATUM@") Then
        tx$ = Mid$(tx$, 12)
        x% = StartX% - Printer.TextWidth(Space$(10))
'        If (Len(tx$) > 8) Then Printer.Font.Size = 10
        Printer.CurrentX = x% - Printer.TextWidth("12345678")
        Dim OrgY%
        OrgY = Printer.CurrentY
        Printer.CurrentY = StartY
        Printer.Print tx$;
'        Printer.Font.Size = 12
        Printer.CurrentY = OrgY
    Else
        Printer.CurrentX = x%
        Printer.Print tx$;
    End If
    
    Printer.Font.Size = 10
    x% = StartX%    ' - Printer.TextWidth(Space$(2))
    tx$ = HoleDruckZeile$
    Call OemToChar(tx$, tx$)
    Printer.CurrentX = x% - Printer.TextWidth(tx$)
    Printer.Print tx$
    
'    If (MagSpeicherIndex > 0) Then  'And (ParenteralRezept < 0) Then
    If (MagSpeicherIndex > 0) And (RezepturDruck) Then  'neu in 4.0.65
        Call ActProgram.DruckTaxierung(TaxierungX%, TaxierungY%)
    End If
End If

Printer.EndDoc

Beep

Call DefErrPop
End Sub

Function HoleDruckZeile$()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("HoleDruckZeile$")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim ret$

ret$ = RTrim(RezDruckZeilen$(RezDruckInd%))
RezDruckInd% = RezDruckInd% + 1

HoleDruckZeile$ = ret$

Call DefErrPop
End Function

Sub SeriellSend(SendStr$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("SeriellSend")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim ret%
Dim l$, deb$

ret% = True

frmAction.comSenden.Output = SendStr$
'deb$ = "> " + SendStr$: Call StatusZeile(deb$)

'SeriellSend% = ret%


Call DefErrPop
End Sub

Function SeriellReceive$()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("SeriellReceive$")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim ret$
Dim ch As Variant
Dim chByte() As Byte

ret$ = ""

If (frmAction.comSenden.InBufferCount > 0) Then
'   char$ = SendeForm.comSenden.Input
    ch = frmAction.comSenden.Input
    chByte = ch
    ret$ = Chr$(chByte(0))
End If

SeriellReceive$ = ret$

Call DefErrPop
End Function

Function OpenDruckerCom%(sForm As Object, xFilePara$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("OpenDruckerCom%")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim i%, ret%, fehler%, CommPort%, ind%, ind2%, stByte%
Dim h$, Settings$

ret% = True

fehler% = 0
' COM einsetzen.
ind% = InStr(xFilePara$, "COM")
If (ind% > 0) Then
    CommPort% = Val(Mid$(xFilePara$, ind% + 3, 1))
    Settings$ = Mid$(xFilePara$, ind% + 5)
    
    stByte% = 1
    For i% = 1 To 4
        ind2% = InStr(stByte%, Settings$, ",")
        If (ind2% > 0) Then
            If (i% = 4) Then
                Settings$ = Left$(Settings$, ind2% - 1)
            Else
                stByte% = ind2% + 1
            End If
        Else
            Exit For
        End If
    Next i%
End If

With sForm
    .comSenden.CommPort = CommPort%
    .comSenden.Settings = Settings$
'    .comSenden.InputMode = comInputModeText
    .comSenden.InputMode = comInputModeBinary
    .comSenden.Handshaking = 0  ' comRTSXOnXOff
    .comSenden.InputLen = 1
    
    ' _Anschluß öffnen.
    On Error GoTo ErrorHandler
    .comSenden.PortOpen = True
End With

If (fehler%) Then ret% = False

OpenDruckerCom% = ret%

Call DefErrPop: Exit Function

ErrorHandler:
    fehler% = Err
    Err = 0
    Resume Next
    Return

End Function

Function WarteAufRezept%()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("WarteAufRezept%")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

Tm290WaitErg% = True

frmTm290Wait.Show 1

LetztDruckZeit& = Timer

WarteAufRezept% = Tm290WaitErg%
  
Call DefErrPop
End Function

Sub SeriellPause()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("SeriellPause")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim StartSek, IstSek

StartSek = Timer
Do
    If ((Timer - StartSek) > 1) Then Exit Do
'    DoEvents
Loop

Call DefErrPop
End Sub

Sub GetPrinterParam()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("GetPrinterParam")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim dpiX As Long, dpiY As Long
Dim MarginLeft As Long, MarginRight As Long
Dim MarginTop As Long, MarginBottom As Long
Dim PrintAreaHorz As Long, PrintAreaVert As Long
Dim PhysHeight As Long, PhysWidth As Long
Dim Info As String

dpiX = GetDeviceCaps(Printer.hdc, LOGPIXELSX)
Info = "Pixels X: " & dpiX & " dpi"

dpiY = GetDeviceCaps(Printer.hdc, LOGPIXELSY)
Info = Info & vbCrLf & "Pixels Y: " & dpiY & " dpi"

MarginLeft = GetDeviceCaps(Printer.hdc, PHYSICALOFFSETX)
Info = Info & vbCrLf & "Unprintable space on left: " & _
MarginLeft & " pixels = " & MarginLeft / dpiX & " inches"

MarginTop = GetDeviceCaps(Printer.hdc, PHYSICALOFFSETY)
Info = Info & vbCrLf & "Unprintable space on top: " & _
MarginTop & " pixels = " & MarginTop / dpiY & " inches"

PrintAreaHorz = GetDeviceCaps(Printer.hdc, HORZRES)
Info = Info & vbCrLf & "Printable space (Horizontal): " & _
PrintAreaHorz & " pixels = " & PrintAreaHorz / dpiX & " inches"

PrintAreaVert = GetDeviceCaps(Printer.hdc, VERTRES)
Info = Info & vbCrLf & "Printable space (Vertical): " & _
PrintAreaVert & " pixels = " & PrintAreaVert / dpiY & " inches"

PhysWidth = GetDeviceCaps(Printer.hdc, PHYSICALWIDTH)
Info = Info & vbCrLf & "Total space (Horizontal): " & _
PhysWidth & " pixels = " & PhysWidth / dpiX & " inches"

MarginRight = PhysWidth - PrintAreaHorz - MarginLeft
Info = Info & vbCrLf & "Unprintable space on right: " & _
MarginRight & " pixels = " & MarginRight / dpiX & " inches"

PhysHeight = GetDeviceCaps(Printer.hdc, PHYSICALHEIGHT)
Info = Info & vbCrLf & "Total space (Vertical): " & _
PhysHeight & " pixels = " & PhysHeight / dpiY & " inches"

MarginBottom = PhysHeight - PrintAreaVert - MarginTop
Info = Info & vbCrLf & "Unprintable space on bottom: " & _
MarginBottom & " pixels = " & MarginBottom / dpiY & " inches"

MessageBox Info, , "GetDeviceCaps Returned the Following:"

Call DefErrPop
End Sub

Function PrinterScaleX#(dVal#, ScaleOld%, ScaleNew%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("PrinterScaleX#")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim OrgScaleMode%
Dim OldScaleHeight&, NewScaleHeight&

With Printer
    OrgScaleMode% = .ScaleMode
    .ScaleMode = ScaleOld%
    OldScaleHeight& = .ScaleWidth
    .ScaleMode = ScaleNew%
    NewScaleHeight& = .ScaleWidth
    .ScaleMode = OrgScaleMode%
End With
PrinterScaleX# = dVal# * (NewScaleHeight& / OldScaleHeight&)

Call DefErrPop
End Function

Function PrinterScaleY#(dVal#, ScaleOld%, ScaleNew%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("PrinterScaleY#")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim OrgScaleMode%
Dim OldScaleHeight&, NewScaleHeight&

With Printer
    OrgScaleMode% = .ScaleMode
    .ScaleMode = ScaleOld%
    OldScaleHeight& = .ScaleHeight
    .ScaleMode = ScaleNew%
    NewScaleHeight& = .ScaleHeight
    .ScaleMode = OrgScaleMode%
End With
PrinterScaleY# = dVal# * (NewScaleHeight& / OldScaleHeight&)

Call DefErrPop
End Function

Public Function code128$(chaine$)
  'Cette fonction est régie par la Licence Générale Publique Amoindrie GNU (GNU LGPL)
  'This function is governed by the GNU Lesser General Public License (GNU LGPL)
  'V 2.0.0
  'Paramètres : une chaine
  'Parameters : a string
  'Retour : * une chaine qui, affichée avec la police CODE128.TTF, donne le code barre
  '         * une chaine vide si paramètre fourni incorrect
  'Return : * a string which give the bar code when it is dispayed with CODE128.TTF font
  '         * an empty string if the supplied parameter is no good
  Dim i%, checksum&, mini%, dummy%, tableB As Boolean
  code128$ = ""
  If Len(chaine$) > 0 Then
  'Vérifier si caractères valides
  'Check for valid characters
    For i% = 1 To Len(chaine$)
      Select Case Asc(Mid$(chaine$, i%, 1))
      Case 32 To 126, 203
      Case Else
        i% = 0
        Exit For
      End Select
    Next
    'Calculer la chaine de code en optimisant l'usage des tables B et C
    'Calculation of the code string with optimized use of tables B and C
    code128$ = ""
    tableB = True
    If i% > 0 Then
      i% = 1 'i% devient l'index sur la chaine / i% become the string index
      Do While i% <= Len(chaine$)
        If tableB Then
          'Voir si intéressant de passer en table C / See if interesting to switch to table C
          'Oui pour 4 chiffres au début ou à la fin, sinon pour 6 chiffres / yes for 4 digits at start or end, else if 6 digits
          mini% = IIf(i% = 1 Or i% + 3 = Len(chaine$), 4, 6)
          GoSub testnum
          If mini% < 0 Then 'Choix table C / Choice of table C
            If i% = 1 Then 'Débuter sur table C / Starting with table C
              code128$ = Chr$(210)
            Else 'Commuter sur table C / Switch to table C
              code128$ = code128$ & Chr$(204)
            End If
            tableB = False
          Else
            If i% = 1 Then code128$ = Chr$(209) 'Débuter sur table B / Starting with table B
          End If
        End If
        If Not tableB Then
          'On est sur la table C, essayer de traiter 2 chiffres / We are on table C, try to process 2 digits
          mini% = 2
          GoSub testnum
          If mini% < 0 Then 'OK pour 2 chiffres, les traiter / OK for 2 digits, process it
            dummy% = Val(Mid$(chaine$, i%, 2))
            dummy% = IIf(dummy% < 95, dummy% + 32, dummy% + 105)
            code128$ = code128$ & Chr$(dummy%)
            i% = i% + 2
          Else 'On n'a pas 2 chiffres, repasser en table B / We haven't 2 digits, switch to table B
            code128$ = code128$ & Chr$(205)
            tableB = True
          End If
        End If
        If tableB Then
          'Traiter 1 caractère en table B / Process 1 digit with table B
          code128$ = code128$ & Mid$(chaine$, i%, 1)
          i% = i% + 1
        End If
      Loop
      'Calcul de la clé de contrôle / Calculation of the checksum
      For i% = 1 To Len(code128$)
        dummy% = Asc(Mid$(code128$, i%, 1))
        dummy% = IIf(dummy% < 127, dummy% - 32, dummy% - 105)
        If i% = 1 Then checksum& = dummy%
        checksum& = (checksum& + (i% - 1) * dummy%) Mod 103
      Next
      'Calcul du code ASCII de la clé / Calculation of the checksum ASCII code
      checksum& = IIf(checksum& < 95, checksum& + 32, checksum& + 105)
      'Ajout de la clé et du STOP / Add the checksum and the STOP
      code128$ = code128$ & Chr$(checksum&) & Chr$(211)
    End If
  End If
  Exit Function
testnum:
  'si les mini% caractères à partir de i% sont numériques, alors mini%=0
  'if the mini% characters from i% are numeric, then mini%=0
  mini% = mini% - 1
  If i% + mini% <= Len(chaine$) Then
    Do While mini% >= 0
      If Asc(Mid$(chaine$, i% + mini%, 1)) < 48 Or Asc(Mid$(chaine$, i% + mini%, 1)) > 57 Then Exit Do
      mini% = mini% - 1
    Loop
  End If
Return
End Function

