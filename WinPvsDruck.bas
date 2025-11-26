Attribute VB_Name = "modDruck"
Option Explicit

Private Const DefErrModul = "WINPVSDRUCK.BAS"

Public Const RASTERCAPS = 38        '  Bitblt capabilities
Public Const RC_PALETTE = &H100                 '  supports a palette
Public Const SIZEPALETTE = 104      '  Number of entries in physical palette

Type PALETTEENTRY
    peRed As Byte
    peGreen As Byte
    peBlue As Byte
    peFlags As Byte
End Type

Type LOGPALETTE
    palVersion As Integer
    palNumEntries As Integer
    palPalEntry(1) As PALETTEENTRY
End Type

Private Type PicBmp
    Size As Long
    Type As Long
    hBmp As Long
    hPal As Long
    Reserved As Long
End Type
Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type


Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Declare Function GetSystemPaletteEntries Lib "gdi32" (ByVal hdc As Long, ByVal wStartIndex As Long, ByVal wNumEntries As Long, lpPaletteEntries As PALETTEENTRY) As Long
Declare Function CreatePalette Lib "gdi32" (lpLogPalette As LOGPALETTE) As Long
Declare Function SelectPalette Lib "gdi32" (ByVal hdc As Long, ByVal hPalette As Long, ByVal bForceBackground As Long) As Long
Declare Function RealizePalette Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (PicDesc As PicBmp, RefIID As GUID, ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long


Public AnzDruckSpalten%
Public DruckSpalte() As DruckSpalteStruct
Public DruckFontSize%

Sub WinPvsAusdruck()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("WinPvsAusdruck")
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
Dim i%, j%, k%, ret%, rInd%, Y%, sp%(5), anz%, ind%, AktLief%, Erst%, AnzRetourArtikel%, Handle%
Dim ZentrierX%, DruckerWechsel%, MitArbInd%, GesBreite&, x%
Dim RetourWert#
Dim tx$, h$, AktDruckerName$, titel$, h2$

With frmAction.cboMitarbeiter
  For i% = 1 To UBound(MitArb$)
    If Trim(Left(MitArb$(i%), 20)) = .List(.ListIndex) Then
      MitArbInd% = Val(Mid(MitArb$(i%), 22, 2))
      Exit For
    End If
  Next i%
End With
If frmAction.cboMitarbeiter.ListIndex = 0 Then
    titel$ = "Personal-Verkaufs-Statistik: Gesamtsummen"
Else
    titel$ = "Personal-Verkaufs-Statistik: " + Str$(MitArbInd%) + " - " + para.Personal(MitArbInd%)
End If

With frmAction.flxarbeit(0)

    AnzDruckSpalten% = 15
    ReDim DruckSpalte(AnzDruckSpalten% - 1)
    
    With DruckSpalte(0)
        .titel = "Tag"
        .TypStr = String$(3, "9")
        .Ausrichtung = "R"
    End With
    With DruckSpalte(1)
        .titel = "PrämBasis"
        .TypStr = String$(9, "9")
        .Ausrichtung = "R"
    End With
    With DruckSpalte(2)
        .titel = "Zus.Basis"
        .TypStr = String$(9, "9")
        .Ausrichtung = "R"
    End With
    With DruckSpalte(3)
        .titel = "Snd.Basis"
        .TypStr = String$(9, "9")
        .Ausrichtung = "R"
    End With
    With DruckSpalte(4)
        .titel = "AnzKd"
        .TypStr = String$(5, "9")
        .Ausrichtung = "R"
    End With
    With DruckSpalte(5)
        .titel = "AnzRez"
        .TypStr = String$(6, "9")
        .Ausrichtung = "R"
    End With
    With DruckSpalte(6)
        .titel = "%PräKd"
        .TypStr = String$(7, "9")
        .Ausrichtung = "R"
    End With
    With DruckSpalte(7)
        .titel = "%PrivKd"
        .TypStr = String$(7, "9")
        .Ausrichtung = "R"
    End With
    With DruckSpalte(8)
        .titel = "%RezKd"
        .TypStr = String$(7, "9")
        .Ausrichtung = "R"
    End With
    With DruckSpalte(9)
        .titel = "Erster"
        .TypStr = "999:99"
        .Ausrichtung = "R"
    End With
    With DruckSpalte(10)
        .titel = "Letzter"
        .TypStr = "999:99"
        .Ausrichtung = "R"
    End With
    With DruckSpalte(11)
        .titel = "kleine"
        .TypStr = String$(5, "9")
        .Ausrichtung = "R"
    End With
    With DruckSpalte(12)
        .titel = "Pausen"
        .TypStr = String$(6, "9")
        .Ausrichtung = "R"
    End With
    With DruckSpalte(13)
        .titel = "große"
        .TypStr = String$(5, "9")
        .Ausrichtung = "R"
    End With
    With DruckSpalte(14)
        .titel = "Pausen"
        .TypStr = String$(6, "9")
        .Ausrichtung = "R"
    End With
    
    
    Call InitDruckZeile(True)
    
    Call DruckKopf(titel$)
    
    For i% = 1 To (.Rows - 1)
        h$ = ""
        For j% = 1 To AnzDruckSpalten%
            h$ = h$ + .TextMatrix(i%, j% - 1) + vbTab
        Next j%
        
        Call DruckZeile(h$)
        
        If (Printer.CurrentY > Printer.ScaleHeight - 1000) Then
            Call DruckFuss
            Call DruckKopf(titel$)
        End If
    Next i%
    
    Y% = Printer.CurrentY
    GesBreite& = DruckSpalte(AnzDruckSpalten% - 1).StartX + DruckSpalte(AnzDruckSpalten% - 1).BreiteX
    Printer.Line (DruckSpalte(0).StartX, Y%)-(GesBreite&, Y%)
    Y% = Printer.CurrentY
    Printer.CurrentY = Y% + 30
    
    h$ = ""
    For j% = 1 To AnzDruckSpalten%
        h$ = h$ + aDetail$(j% - 1, MitArbInd%, aTage%) + vbTab
    Next j%
    Call DruckZeile(h$)
    Printer.CurrentY = Printer.CurrentY + 750
End With

    
With frmAction.flxInfo(0)

    AnzDruckSpalten% = 6
    ReDim DruckSpalte(AnzDruckSpalten% - 1)
    
    With DruckSpalte(0)
        .titel = ""
        .TypStr = "Zusatzverk.Kunden/Rezeptkun"
        .Ausrichtung = "L"
    End With
    With DruckSpalte(1)
        .titel = ""
        .TypStr = String$(9, "9") + "  ."
        .Ausrichtung = "R"
    End With
    For i% = 2 To 5
        DruckSpalte(i%) = DruckSpalte(i% - 2)
    Next i%
    
    
    Call InitDruckZeile(True)
    
    For i% = 0 To (.Rows - 1)
        h$ = ""
        For j% = 1 To AnzDruckSpalten%
            h$ = h$ + .TextMatrix(i%, j%)
            If (j% Mod 2 = 0) Then h$ = h$ + Space$(3)
            h$ = h$ + vbTab
        Next j%
        
        Call DruckZeile(h$)
        Printer.CurrentY = Printer.CurrentY + 60
        
        If (Printer.CurrentY > Printer.ScaleHeight - 1000) Then
            Call DruckFuss
            Call DruckKopf(titel$)
        End If
    Next i%
    
    Y% = Printer.CurrentY
    Printer.CurrentY = Y% + 750
End With

AnzDruckSpalten% = 4
ReDim DruckSpalte(AnzDruckSpalten% - 1)

With DruckSpalte(0)
    .titel = ""
    .TypStr = "   Prämiensumme"
    .Ausrichtung = "L"
End With
With DruckSpalte(1)
    .titel = "Prämienbasis"
    .TypStr = .titel
    .Ausrichtung = "R"
End With
With DruckSpalte(2)
    .titel = "Faktor"
    .TypStr = String(12, "9")
    .Ausrichtung = "R"
End With
With DruckSpalte(3)
    .titel = "Prämie"
    .TypStr = String(15, "9")
    .Ausrichtung = "R"
End With


Call InitDruckZeile(True)
    
For i% = 0 To (AnzDruckSpalten% - 1)
    h2$ = RTrim(DruckSpalte(i%).titel)
    If (DruckSpalte(i%).Ausrichtung = "L") Then
        x% = DruckSpalte(i%).StartX
    Else
        x% = DruckSpalte(i%).StartX + DruckSpalte(i%).BreiteX - Printer.TextWidth(h2$)
    End If
    Printer.CurrentX = x%
    Printer.Print h2$;
Next i%

Printer.Print " "

Y% = Printer.CurrentY
GesBreite& = DruckSpalte(AnzDruckSpalten% - 1).StartX + DruckSpalte(AnzDruckSpalten% - 1).BreiteX
Printer.Line (DruckSpalte(0).StartX, Y%)-(GesBreite&, Y%)

Y% = Printer.CurrentY
Printer.CurrentY = Y% + 30

With frmAction.flxInfo(0)
    h$ = "Grundprämie" + vbTab + aDetail$(1, MitArbInd%, aTage%) + vbTab + Format(pb#, "0.00") + vbTab + .TextMatrix(3, 6) + vbTab
    Call DruckZeile(h$)
    h$ = "Zusatzprämie" + vbTab + aDetail$(2, MitArbInd%, aTage%) + vbTab + Format(zpb#, "0.00") + vbTab + .TextMatrix(4, 6) + vbTab
    Call DruckZeile(h$)
    h$ = "Sonderprämie" + vbTab + aDetail$(3, MitArbInd%, aTage%) + vbTab + Format(spb#, "0.00") + vbTab + .TextMatrix(5, 6) + vbTab
    Call DruckZeile(h$)
    
    Y% = Printer.CurrentY
    GesBreite& = DruckSpalte(AnzDruckSpalten% - 1).StartX + DruckSpalte(AnzDruckSpalten% - 1).BreiteX
    Printer.Line (DruckSpalte(0).StartX, Y%)-(GesBreite&, Y%)
    
    Y% = Printer.CurrentY
    Printer.CurrentY = Y% + 30
    
    h$ = "Prämiensumme" + vbTab + vbTab + vbTab + .TextMatrix(6, 6) + vbTab
    Call DruckZeile(h$)
End With

Call DruckFuss(False)
Printer.EndDoc

Call DefErrPop
End Sub

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
        DruckSpalte(i%).BreiteX = Printer.TextWidth(RTrim$(DruckSpalte(i%).TypStr))
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

Sub DruckKopf(titel$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("DruckKopf")
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
Dim l&, i%, pos%, h&, x%, Y%, heute$, SeitenNr$, GesBreite&, ind%
Dim h2$

With Printer
    .CurrentX = 0: .CurrentY = 0
    .Font.Size = 14
    Printer.Print titel$
    
    .CurrentY = .CurrentY + 150
    
    h2$ = frmAction.Caption
    ind% = InStr(h2$, ":")
    If (ind% > 0) Then
        h2$ = Mid$(h2$, ind% + 2)
    End If
    .Font.Size = 11
    Printer.Print h2$

    h2$ = frmAction.lblPrmBasis.Caption
    Printer.Print h2$

    .CurrentY = .CurrentY + 450
    
    If (DruckFontSize% > 0) Then
        .Font.Size = DruckFontSize%
    Else
        .Font.Size = 12
    End If

    For i% = 0 To (AnzDruckSpalten% - 1)
        h2$ = RTrim(DruckSpalte(i%).titel)
        If (DruckSpalte(i%).Ausrichtung = "L") Then
            x% = DruckSpalte(i%).StartX
        Else
            x% = DruckSpalte(i%).StartX + DruckSpalte(i%).BreiteX - Printer.TextWidth(h2$)
        End If
        .CurrentX = x%
        Printer.Print h2$;
    Next i%
    
    Printer.Print " "
    
    Y% = Printer.CurrentY
    GesBreite& = DruckSpalte(AnzDruckSpalten% - 1).StartX + DruckSpalte(AnzDruckSpalten% - 1).BreiteX
    Printer.Line (DruckSpalte(0).StartX, Y%)-(GesBreite&, Y%)

    Y% = Printer.CurrentY
    Printer.CurrentY = Y% + 30
End With

Call DefErrPop
End Sub

Sub DruckFuss(Optional NewPage% = True)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("DruckFuss")
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
Dim l&, h&, Y%

With Printer
    .Font.Bold = False
    .Font.Size = 14
    l& = .TextWidth(para.FISTAM(0))
    h& = .TextHeight("A")
    
    .CurrentX = 0
    .CurrentY = .ScaleHeight - 3 * h&
    Y% = .CurrentY
    Printer.Line (0, Y%)-(.ScaleWidth, Y%)
    
    .CurrentX = 0
    .CurrentY = .CurrentY + 200
    Printer.Print para.FISTAM(0)
    .CurrentX = 0
    Printer.Print para.FISTAM(1);
    
    If (NewPage% = True) Then .NewPage
End With
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
Dim i%, ind%, x%
Dim h$, tx$

h$ = ZeilenText$

For i% = 0 To (AnzDruckSpalten% - 1)
    ind% = InStr(h$, vbTab)
    tx$ = Left$(h$, ind% - 1)
    h$ = Mid$(h$, ind% + 1)
    
    If (tx$ = Chr$(214)) Then
        Printer.Font.Name = "Symbol"
    End If
    
    If (DruckSpalte(i%).Ausrichtung = "L") Then
        x% = DruckSpalte(i%).StartX
    Else
        x% = DruckSpalte(i%).StartX + DruckSpalte(i%).BreiteX - Printer.TextWidth(tx$)
    End If
    Printer.CurrentX = x%
    Printer.Print tx$;
    
    If (tx$ = Chr$(214)) Then
        Printer.Font.Name = "Arial"
    End If
Next i%
Printer.Print " "
   
Call DefErrPop
End Sub

Sub DruckeWindow(frmSrc As Form)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("DruckeWindow")
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
Dim border, faktor
Dim i%, MaxWi%, MaxHe%, wi%, he%, TextHe%, CurrX%, CurrY%
Dim OrgBackColor&
Dim h$

OrgBackColor& = frmSrc.BackColor
frmSrc.BackColor = vbWhite
DoEvents

With Printer
    .Orientation = vbPRORLandscape
    .ScaleMode = vbTwips
    .Font.Name = "Arial"

    border = .ScaleWidth / 20
    
    faktor = frmSrc.ScaleWidth / frmSrc.ScaleHeight
    MaxWi% = .ScaleWidth - 2 * border
    MaxHe% = .ScaleHeight - 2 * border

    wi% = MaxWi%
    he% = wi% / faktor
    If (he% > MaxHe%) Then
        he% = MaxHe%
        wi% = he% * faktor
    End If
        
    Set frmSrc.picAusdruck.Picture = CaptureClient(frmSrc)
    Printer.PaintPicture frmSrc.picAusdruck.Picture, border, border, wi%, he%
    
    .EndDoc
End With

frmSrc.BackColor = OrgBackColor&

Call DefErrPop
End Sub

Public Function CaptureClient(frmSrc As Form) As Picture
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("CaptureClient")
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
Dim he As Single
   
On Error Resume Next
he = frmSrc.cmdEsc.Top
If he = 0 Then he = frmSrc.ScaleHeight
On Error GoTo DefErr

' Call CaptureWindow to capture the client area of the form given
' its window handle and return the resulting Picture object.
Set CaptureClient = CaptureWindow(frmSrc.hwnd, 0, 0, _
   frmSrc.ScaleX(frmSrc.ScaleWidth, frmSrc.ScaleMode, vbPixels), _
   frmSrc.ScaleY(he, frmSrc.ScaleMode, vbPixels))

Call DefErrPop
End Function

' CaptureWindow
'    - Captures any portion of a window.
'
' hWndSrc
'    - Handle to the window to be captured.
'
' Client
'    - If True CaptureWindow captures from the client area of the
'      window.
'    - If False CaptureWindow captures from the entire window.
'
' LeftSrc, TopSrc, WidthSrc, HeightSrc
'    - Specify the portion of the window to capture.
'    - Dimensions need to be specified in pixels.
'
' Returns
'    - Returns a Picture object containing a bitmap of the specified
'      portion of the window that was captured.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''
'
Public Function CaptureWindow(ByVal hWndSrc&, ByVal LeftSrc&, ByVal TopSrc&, ByVal WidthSrc&, ByVal HeightSrc&) As Picture
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("CaptureWindow")
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

Dim hDCMemory As Long
Dim hBmp As Long
Dim hBmpPrev As Long
Dim r As Long
Dim hDCSrc As Long
Dim hPal As Long
Dim hPalPrev As Long
Dim RasterCapsScrn As Long
Dim HasPaletteScrn As Long
Dim PaletteSizeScrn As Long
            
Dim LogPal As LOGPALETTE

hDCSrc = GetDC(hWndSrc) ' Get device context for client area.

' Create a memory device context for the copy process.
hDCMemory = CreateCompatibleDC(hDCSrc)
' Create a bitmap and place it in the memory DC.
hBmp = CreateCompatibleBitmap(hDCSrc, WidthSrc, HeightSrc)
hBmpPrev = SelectObject(hDCMemory, hBmp)

' Get screen properties.
RasterCapsScrn = GetDeviceCaps(hDCSrc, RASTERCAPS) ' Raster
                                                   ' capabilities.
HasPaletteScrn = RasterCapsScrn And RC_PALETTE       ' Palette
                                                     ' support.
PaletteSizeScrn = GetDeviceCaps(hDCSrc, SIZEPALETTE) ' Size of
                                                     ' palette.

' If the screen has a palette make a copy and realize it.
If HasPaletteScrn And (PaletteSizeScrn = 256) Then
   ' Create a copy of the system palette.
   LogPal.palVersion = &H300
   LogPal.palNumEntries = 256
   r = GetSystemPaletteEntries(hDCSrc, 0, 256, _
       LogPal.palPalEntry(0))
   hPal = CreatePalette(LogPal)
   ' Select the new palette into the memory DC and realize it.
   hPalPrev = SelectPalette(hDCMemory, hPal, 0)
   r = RealizePalette(hDCMemory)
End If

' Copy the on-screen image into the memory DC.
r = BitBlt(hDCMemory, 0, 0, WidthSrc, HeightSrc, hDCSrc, _
   LeftSrc, TopSrc, vbSrcCopy)

' Remove the new copy of the  on-screen image.
   hBmp = SelectObject(hDCMemory, hBmpPrev)

   ' If the screen has a palette get back the palette that was
   ' selected in previously.
   If HasPaletteScrn And (PaletteSizeScrn = 256) Then
      hPal = SelectPalette(hDCMemory, hPalPrev, 0)
   End If

   ' Release the device context resources back to the system.
   r = DeleteDC(hDCMemory)
   r = ReleaseDC(hWndSrc, hDCSrc)

   ' Call CreateBitmapPicture to create a picture object from the
   ' bitmap and palette handles. Then return the resulting picture
   ' object.
   Set CaptureWindow = CreateBitmapPicture(hBmp, hPal)
   
Call DefErrPop
End Function

Public Function CreateBitmapPicture(ByVal hBmp&, ByVal hPal&) As Picture
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("CaptureWindow")
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

Dim r As Long
Dim Pic As PicBmp
' IPicture requires a reference to "Standard OLE Types."
Dim IPic As IPicture
Dim IID_IDispatch As GUID

' Fill in with IDispatch Interface ID.
With IID_IDispatch
   .Data1 = &H20400
   .Data4(0) = &HC0
   .Data4(7) = &H46
End With

' Fill Pic with necessary parts.
With Pic
   .Size = Len(Pic)          ' Length of structure.
   .Type = vbPicTypeBitmap   ' Type of Picture (bitmap).
   .hBmp = hBmp              ' Handle to bitmap.
   .hPal = hPal              ' Handle to palette (may be null).
End With

' Create Picture object.
r = OleCreatePictureIndirect(Pic, IID_IDispatch, 1, IPic)

' Return the new Picture object.
Set CreateBitmapPicture = IPic

Call DefErrPop
End Function



