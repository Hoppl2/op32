VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsWinPara"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim iWorkAreaHeight&, iWorkAreaWidth&
Dim iFrmCaptionHeight&, iFrmMenuHeight&, iFrmBorderHeight&, iFrmScrollHeight&, iFrmCaptionFontSize&
Dim iFrmMenuColor&


Dim iBildFaktor!

Dim iLinksX%, iTitelY%, iFlexY%, iUntenFreiY%, iOhneInfoY%
Dim iButtonX%, iButtonY%
    
Dim iNlCaptionY%, iNlFlexBackY%

Dim iFontName$(1)
Dim iFontSize%(1)

Dim iFarbeArbeit&, iFarbeInfo&
Dim iFarbeRahmen%
Dim iFarbeAktZeile&
Dim iFarbeDunklerBereich&

Dim iIniSection$
Dim iIniDatei$

Dim iProgName$

Dim iOptipharmRot&

Dim inlFlexBackColor&, inlFlexBackColorSel&, inlFlexBackColorFixed&, inlFlexBackColorInfo&, inlRoundRectColor&
Dim iEntwicklungsUmgebung%

Dim iAnzPrinter%
Dim iPrinterName$()

Private Const DefErrModul = "WINPARA.CLS"

Private Sub Class_Initialize()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("Class_Initialize")
Call clsError.DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case clsError.DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call ProjektForm.EndeDll
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Set wPara1 = Me
Call clsError.DefErrPop
End Sub

Public Property Get WorkAreaHeight&()
WorkAreaHeight& = iWorkAreaHeight&
End Property
Public Property Let WorkAreaHeight(ByVal vNewValue&)
iWorkAreaHeight& = vNewValue&
End Property

Public Property Get WorkAreaWidth&()
WorkAreaWidth& = iWorkAreaWidth&
End Property
Public Property Let WorkAreaWidth(ByVal vNewValue&)
iWorkAreaWidth& = vNewValue&
End Property

Public Property Get FrmCaptionHeight&()
FrmCaptionHeight& = iFrmCaptionHeight&
End Property
Public Property Let FrmCaptionHeight(ByVal vNewValue&)
iFrmCaptionHeight& = vNewValue&
End Property

Public Property Get FrmMenuHeight&()
FrmMenuHeight& = iFrmMenuHeight&
End Property
Public Property Let FrmMenuHeight(ByVal vNewValue&)
iFrmMenuHeight& = vNewValue&
End Property

Public Property Get FrmBorderHeight&()
FrmBorderHeight& = iFrmBorderHeight&
End Property
Public Property Let FrmBorderHeight(ByVal vNewValue&)
iFrmBorderHeight& = vNewValue&
End Property

Public Property Get FrmScrollHeight&()
FrmScrollHeight& = iFrmScrollHeight&
End Property
Public Property Let FrmScrollHeight(ByVal vNewValue&)
iFrmScrollHeight& = vNewValue&
End Property

Public Property Get BildFaktor!()
BildFaktor! = iBildFaktor!
End Property
Public Property Let BildFaktor(ByVal vNewValue!)
iBildFaktor! = vNewValue!
End Property

Public Property Get LinksX%()
LinksX% = iLinksX%
End Property
Public Property Let LinksX(ByVal vNewValue%)
iLinksX% = vNewValue%
End Property

Public Property Get TitelY%()
TitelY% = iTitelY%
End Property
Public Property Let TitelY(ByVal vNewValue%)
iTitelY% = vNewValue%
End Property

Public Property Get FlexY%()
FlexY% = iFlexY%
End Property
Public Property Let FlexY(ByVal vNewValue%)
iFlexY% = vNewValue%
End Property

Public Property Get UntenFreiY%()
UntenFreiY% = iUntenFreiY%
End Property
Public Property Let UntenFreiY(ByVal vNewValue%)
iUntenFreiY% = vNewValue%
End Property

Public Property Get OhneInfoY%()
OhneInfoY% = iOhneInfoY%
End Property
Public Property Let OhneInfoY(ByVal vNewValue%)
iOhneInfoY% = vNewValue%
End Property

Public Property Get ButtonX%()
ButtonX% = iButtonX%
End Property
Public Property Let ButtonX(ByVal vNewValue%)
iButtonX% = vNewValue%
End Property

Public Property Get ButtonY%()
ButtonY% = iButtonY%
End Property
Public Property Let ButtonY(ByVal vNewValue%)
iButtonY% = vNewValue%
End Property

Public Property Get NlCaptionY%()
NlCaptionY% = iNlCaptionY%
End Property
Public Property Let NlCaptionY(ByVal vNewValue%)
iNlCaptionY% = vNewValue%
End Property

Public Property Get NlFlexBackY%()
NlFlexBackY% = iNlFlexBackY%
End Property
Public Property Let NlFlexBackY(ByVal vNewValue%)
iNlFlexBackY% = vNewValue%
End Property

Public Property Get FarbeArbeit&()
FarbeArbeit& = iFarbeArbeit&
End Property

Public Property Get FarbeInfo&()
FarbeInfo& = iFarbeInfo&
End Property

Public Property Get FarbeRahmen%()
FarbeRahmen% = iFarbeRahmen%
End Property
Public Property Let FarbeRahmen(ByVal vNewValue%)
Dim l&
iFarbeRahmen% = vNewValue%
l& = WritePrivateProfileString(iIniSection$, "FarbeRahmen", Str$(iFarbeRahmen%), iIniDatei$)
End Property

Public Property Get FarbeAktZeile&()
FarbeAktZeile& = iFarbeAktZeile&
End Property
'Public Property Let FarbeAktZeile(ByVal vNewValue%)
'Dim l&
'iFarbeAktZeile% = vNewValue%
'l& = WritePrivateProfileString(iIniSection$, "FarbeAktZeile", Str$(iFarbeAktZeile%), iIniDatei$)
'End Property

Public Property Get FarbeDunklerBereich&()
FarbeDunklerBereich& = iFarbeDunklerBereich&
End Property

Public Property Get FontName$(ind%)
FontName$ = iFontName$(ind%)
End Property
Public Property Get FontSize%(ind%)
FontSize% = iFontSize%(ind%)
End Property

Public Property Get FaktorX!()
FaktorX! = iFaktorX!
End Property
Public Property Let FaktorX(ByVal vNewValue!)
iFaktorX! = vNewValue!
End Property

Public Property Get FaktorY!()
FaktorY! = iFaktorY!
End Property
Public Property Let FaktorY(ByVal vNewValue!)
iFaktorY! = vNewValue!
End Property

Public Property Get OptipharmRot&()
OptipharmRot& = iOptipharmRot&
End Property

Public Property Get nlRed&()
nlRed& = RGB(135, 61, 52)
End Property
Public Property Get nlGreen&()
nlGreen& = RGB(206, 185, 23)
End Property
Public Property Get nlYellow&()
nlYellow& = RGB(201, 123, 58)
End Property
Public Property Get nlFlexBackColor&()
nlFlexBackColor& = inlFlexBackColor& ' RGB(255, 246, 217)
End Property
Public Property Get nlFlexBackColorSel&()
nlFlexBackColorSel& = inlFlexBackColorSel&   ' RGB(217, 130, 24)
End Property
Public Property Get nlFlexBackColorFixed&()
nlFlexBackColorFixed& = inlFlexBackColorFixed&    ' RGB(199, 176, 123)
End Property
Public Property Get nlFlexBackColorInfo&()
nlFlexBackColorInfo& = inlFlexBackColorInfo& ' vbGreen       ' RGB(232, 217, 172)
End Property
Public Property Get nlRoundRectColor&()
nlRoundRectColor& = inlRoundRectColor&   ' vbYellow ' RGB(240, 240, 240)
End Property

Public Property Get EntwicklungsUmgebung%()
EntwicklungsUmgebung% = iEntwicklungsUmgebung%
End Property

Sub HoleWindowsParameter()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("HoleWindowsParameter")
Call clsError.DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case clsError.DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call ProjektForm.EndeDll
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim l&
Dim WorkRect As RECT
Dim iNonClientMetrics As NONCLIENTMETRICS

l& = SystemParametersInfo(SPI_GETWORKAREA, vbNull, WorkRect, 0)
iWorkAreaHeight& = (WorkRect.Bottom - WorkRect.Top - 10) * Screen.TwipsPerPixelY
iWorkAreaWidth& = (WorkRect.Right - WorkRect.Left) * Screen.TwipsPerPixelY

iFrmScrollHeight& = Abs(QueryRegistryValue&(".DEFAULT\Control Panel\Desktop\WindowMetrics", "ScrollHeight"))
iFrmScrollHeight& = GetSystemMetrics(SM_CXVSCROLL) * 15

iFrmCaptionHeight& = GetSystemMetrics(SM_CYCAPTION) * 15
iFrmMenuHeight& = GetSystemMetrics(SM_CYMENU) * 15
iFrmBorderHeight& = GetSystemMetrics(SM_CYFRAME) * 15

'iNonClientMetrics
iNonClientMetrics.cbSize = Len(iNonClientMetrics)
l& = SystemParametersInfo(SPI_GETNONCLIENTMETRICS, Len(iNonClientMetrics), iNonClientMetrics, 0)
iFrmCaptionFontSize& = Abs(iNonClientMetrics.lfCaptionFont.lfHeight)

iFrmMenuColor& = GetSysColor(4)

' Verhältnis der Bildschirmauflösungen
If (iNewLine) Then
    iBildFaktor! = Screen.Width / 25200
Else
    iBildFaktor! = Screen.Width / 12000
End If

iLinksX% = 150 * iBildFaktor!
iTitelY% = 150 * iBildFaktor!   '195
iFlexY% = 495 * iBildFaktor!   '690
iUntenFreiY% = 60 * iBildFaktor
iOhneInfoY% = 600 * iBildFaktor!

iButtonX% = 1200 * iBildFaktor!
iButtonY% = 390 * iBildFaktor!

iOptipharmRot = vbRed

If (iNewLine) Then
    iLinksX% = 210 * iFaktorX   '300
    iTitelY% = 300 * iFaktorY
    iFlexY% = 600 * iFaktorY
    iNlCaptionY% = 450 * iFaktorY
    iNlFlexBackY = 210 * iFaktorY

    iOptipharmRot = RGB(255, 131, 62)
    
    Call HoleNewLineParameter
End If

'iEntwicklungsUmgebung% = (FindWindow(vbNullString, App.EXEName + " - Microsoft Visual Basic [Ausführen]") <> 0)
iEntwicklungsUmgebung% = CheckEntwicklungsUmgebung%

Call clsError.DefErrPop
End Sub

Sub HoleNewLineParameter()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("HoleNewLineParameter")
Call clsError.DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case clsError.DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call ProjektForm.EndeDll
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim i%, j%, ind%, iRGB%(2)
Dim l&, lColor&
Dim h$, h2$, sIniDatei$, sIniSection$, sKey$(4)
    
sKey(0) = "FlexBackColor 255,246,217"
sKey(1) = "FlexBackColorSel 217,130,24"
sKey(2) = "FlexBackColorFixed 199,176,123"
sKey(3) = "FlexBackColorInfo 232,217,172"
sKey(4) = "RoundRectColor 240,240,240"

sIniDatei = "\user\NewLine.op"
sIniSection = "Allgemein"

For j = 0 To 4
    For i = 0 To 2
        iRGB(i) = 255
    Next i
    
    ind = InStr(sKey(j), " ")
    h$ = Space$(20)
    h2$ = Left$(Mid(sKey(j), ind + 1) + Space$(20), 20)
    l& = GetPrivateProfileString(sIniSection$, Left$(sKey(j), ind - 1), h2$, h$, 21, sIniDatei)
    h$ = Trim(Left$(h$, l&))
    ind% = InStr(h$, ",")
    If (ind > 0) Then
        h = h + ","
        For i = 0 To 2
            ind% = InStr(h$, ",")
            If (ind > 0) Then
                iRGB(i) = Val(Left(h, ind - 1))
                h = Mid(h, ind + 1)
            Else
                Exit For
            End If
        Next i
    End If
    
    lColor = RGB(iRGB(0), iRGB(1), iRGB(2))
    If (j = 0) Then
        inlFlexBackColor& = lColor
    ElseIf (j = 1) Then
        inlFlexBackColorSel& = lColor
    ElseIf (j = 2) Then
        inlFlexBackColorFixed& = lColor
    ElseIf (j = 3) Then
        inlFlexBackColorInfo& = lColor
    Else
        inlRoundRectColor& = lColor
    End If
Next j%


Call clsError.DefErrPop
End Sub

Function CheckEntwicklungsUmgebung%()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("CheckEntwicklungsUmgebung%")
Call clsError.DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
'If (Err.Number = (999 + vbObjectError)) Then End
Select Case clsError.DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call ProjektForm.EndeDll
'On Error GoTo 0
'Err.Raise 999 + vbObjectError, "OPWMENU.VBP", "Fehler in DLL"
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim ret%
Dim CurrWnd As Long
Dim Length As Integer
Dim ListItem As String
Dim x As Long
Dim ThreadID As Long

ret = 0
CurrWnd = GetWindow(GetForegroundWindow, GW_HWNDFIRST)
Do While CurrWnd <> 0
    Length = GetWindowTextLength(CurrWnd)
    ListItem = Space(Length + 1)
    Length = GetWindowText(CurrWnd, ListItem, Length + 1)
    If (Length > 0) Then
        If (InStr(ListItem, " - Microsoft Visual Basic [Ausführen]") > 0) Then
            ret = True
            Exit Do
        End If
    End If
    CurrWnd = GetWindow(CurrWnd, GW_HWNDNEXT)
Loop

CheckEntwicklungsUmgebung% = ret

Call clsError.DefErrPop
End Function

Private Function MulMul(arg1 As Long, arg2 As Long, arg3 As Long) As Integer
 
  'A weird name for a function;
  'actually, it is based on the reverse of MulDiv
  '(the multiple divide C macro) and since it
  'returns the opposite data, I named it MulMul
  '(though there is no multiplication in it!).
  'I have no idea what its real corresponding
  'name would be in C.
   Dim tmp As Single
   
   tmp = arg2 / arg3
   tmp = arg1 / tmp
   
   MulMul = CInt(tmp)
   
End Function

Function QueryRegistryValue&(sKeyName$, sValueName$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("QueryRegistryValue&")
Call clsError.DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case clsError.DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call ProjektForm.EndeDll
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim lRetVal&         'result of the API functions
Dim hKey&         'handle of opened key
Dim vValue As Variant      'setting of queried value

lRetVal& = RegOpenKeyEx(HKEY_USERS, sKeyName$, 0, KEY_ALL_ACCESS, hKey&)
lRetVal& = QueryValueEx(hKey&, sValueName$, vValue)
RegCloseKey (hKey&)

QueryRegistryValue& = vValue
Call clsError.DefErrPop
End Function

Function QueryValueEx&(ByVal lhKey&, ByVal szValueName$, vValue As Variant)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("QueryValueEx&")
Call clsError.DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case clsError.DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call ProjektForm.EndeDll
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim cch&, lrc&, lType&, lValue&
Dim sValue$

'On Error GoTo QueryValueExError

' Determine the size and type of data to be read
lrc& = RegQueryValueExNULL(lhKey&, szValueName$, 0&, lType&, 0&, cch&)
If (lrc& = ERROR_NONE) Then

    Select Case lType&
        ' For strings
        Case REG_SZ:
            sValue$ = String(cch&, 0)
            lrc& = RegQueryValueExString(lhKey&, szValueName$, 0&, lType&, sValue$, cch&)
            If (lrc& = ERROR_NONE) Then
                vValue = Left$(sValue$, cch& - 1)
            Else
                vValue = Empty
            End If
            
        ' For DWORDS
        Case REG_DWORD:
            lrc& = RegQueryValueExLong(lhKey&, szValueName$, 0&, lType&, lValue&, cch&)
            If (lrc& = ERROR_NONE) Then
                vValue = lValue&
            End If
            
        Case Else
            'all other data types not supported
            lrc& = -1
    End Select
End If

QueryValueEx& = lrc&

Call clsError.DefErrPop
End Function

Sub HoleGlobalIniWerte(IniSection$, IniDatei$, ProgName$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("HoleGlobalIniWerte")
Call clsError.DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case clsError.DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call ProjektForm.EndeDll
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim ind%
Dim l&
Dim h$, h2$

iIniSection$ = IniSection$
iIniDatei$ = IniDatei$

iProgName$ = ProgName$

h2$ = Left$("Arial,8" + Space$(20), 20)
h$ = Space$(20)
l& = GetPrivateProfileString(iIniSection$, "FontInformation", h2$, h$, 21, iIniDatei$)
h$ = Left$(h$, l&)
ind% = InStr(h$, ",")
iFontName$(0) = RTrim$(Left$(h$, ind% - 1))
iFontSize%(0) = Val(Mid$(h$, ind% + 1))

h2$ = Left$("Arial,8" + Space$(20), 20)
h$ = Space$(20)
l& = GetPrivateProfileString(iIniSection$, "FontBeschriftung", h2$, h$, 21, iIniDatei$)
h$ = Left$(h$, l&)
ind% = InStr(h$, ",")
iFontName$(1) = RTrim$(Left$(h$, ind% - 1))
iFontSize%(1) = Val(Mid$(h$, ind% + 1))
 
If (iNewLine) Then
    Call HoleProgFont
End If
  
' Call InitFont(Me)
  
h2$ = "C0C0C0"
h$ = Space$(8)
l& = GetPrivateProfileString(iIniSection$, "FarbeArbeit", h2$, h$, 9, iIniDatei$)
h$ = Left$(h$, l&)
iFarbeArbeit& = BerechneFarbWert&(h$)

h2$ = "80FFFF"
h$ = Space$(8)
l& = GetPrivateProfileString(iIniSection$, "FarbeInfo", h2$, h$, 9, iIniDatei$)
h$ = Left$(h$, l&)
iFarbeInfo& = BerechneFarbWert&(h$)
   
h$ = "0"
l& = GetPrivateProfileString(iIniSection$, "FarbeRahmen", "0", h$, 2, iIniDatei$)
h$ = Left$(h$, l&)
iFarbeRahmen% = Val(h$)
    
'h$ = "0"
'l& = GetPrivateProfileString(iIniSection$, "FarbeAktZeile", "0", h$, 2, iIniDatei$)
'h$ = Left$(h$, l&)
'iFarbeAktZeile% = val(h$)

h$ = Space$(8)
l& = GetPrivateProfileString(iIniSection$, "FarbeAktuelleZeile", h$, h$, 9, iIniDatei$)
h$ = Left$(h$, l&)
If (Trim(h$) = "") Then
    iFarbeAktZeile& = vbHighlightText
Else
    iFarbeAktZeile& = BerechneFarbWert&(h$)
End If
   
h$ = Space$(8)
l& = GetPrivateProfileString(iIniSection$, "FarbeDunklerBereich", h$, h$, 9, iIniDatei$)
h$ = Left$(h$, l&)
If (Trim(h$) = "") Then
    iFarbeDunklerBereich& = vbGrayText
Else
    iFarbeDunklerBereich& = BerechneFarbWert&(h$)
End If
   
If (iNewLine) Then
    iFarbeArbeit = iFrmMenuColor&
End If


'Call InitAlleBereichsFarben

Call clsError.DefErrPop
End Sub

Function BerechneFarbWert&(Farb$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("BerechneFarbWert&")
Call clsError.DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case clsError.DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call ProjektForm.EndeDll
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim i%, l%
Dim tetrade&, Wert&
Dim ch$

Wert& = 0
l% = Len(Farb$)
For i% = 1 To l%
    Wert& = Wert& * 16
    ch$ = Mid$(Farb$, i%, 1)
    tetrade& = Val("&H" + ch$)
    Wert& = Wert& + tetrade&
Next i%
    
BerechneFarbWert& = Wert&

Call clsError.DefErrPop
End Function

Function EditFont%(CommDlg As Object, Index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("EditFont%")
Call clsError.DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case clsError.DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call ProjektForm.EndeDll
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim ret%
Dim l&
Dim h$

'On Error Resume Next

ret% = False

CommDlg.FontName = RTrim$(iFontName$(Index))
CommDlg.FontSize = iFontSize%(Index)
CommDlg.FontBold = False
CommDlg.FontItalic = False
CommDlg.FontStrikethru = False
CommDlg.FontUnderline = False
CommDlg.flags = cdlCFBoth
CommDlg.CancelError = False
Call CommDlg.ShowFont
If (Err = 0) Then
    If ((iFontName$(Index) <> CommDlg.FontName) Or (iFontSize%(Index) <> CommDlg.FontSize)) Then
        iFontName$(Index) = CommDlg.FontName
        iFontSize%(Index) = CommDlg.FontSize
'        h$ = RTrim$(iFontName$(index)) + "," + Mid$(Str$(iFontSize%(index)), 2)
'        If (index = 0) Then
'            l& = WritePrivateProfileString(iIniSection$, "FontInformation", h$, iIniDatei$)
'        Else
'            l& = WritePrivateProfileString(iIniSection$, "FontBeschriftung", h$, iIniDatei$)
'        End If
        
        Call SpeicherProgFont(Index)
        ret% = True
    End If
End If

EditFont% = ret%

Call clsError.DefErrPop
End Function

Function EditFarbe%(CommDlg As Object, Index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("EditFarbe%")
Call clsError.DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case clsError.DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call ProjektForm.EndeDll
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim ret%
Dim l&

On Error Resume Next

ret% = False

If (Index = 0) Then
    CommDlg.Color = iFarbeArbeit&
ElseIf (Index = 1) Then
    CommDlg.Color = iFarbeInfo&
ElseIf (Index = 2) Then
    CommDlg.Color = iFarbeDunklerBereich&
Else
    CommDlg.Color = iFarbeAktZeile&
End If
CommDlg.CancelError = True
CommDlg.flags = cdlCCFullOpen + cdlCCRGBInit
Call CommDlg.ShowColor
If (Err = 0) Then
    If (Index = 0) Then
        iFarbeArbeit& = CommDlg.Color
        l& = WritePrivateProfileString(iIniSection$, "FarbeArbeit", Hex$(iFarbeArbeit&), iIniDatei$)
    ElseIf (Index = 1) Then
        iFarbeInfo& = CommDlg.Color
        l& = WritePrivateProfileString(iIniSection$, "FarbeInfo", Hex$(iFarbeInfo&), iIniDatei$)
    ElseIf (Index = 2) Then
        iFarbeDunklerBereich& = CommDlg.Color
        l& = WritePrivateProfileString(iIniSection$, "FarbeDunklerBereich", Hex$(iFarbeDunklerBereich&), iIniDatei$)
    Else
        iFarbeAktZeile& = CommDlg.Color
        l& = WritePrivateProfileString(iIniSection$, "FarbeAktuelleZeile", Hex$(iFarbeAktZeile&), iIniDatei$)
    End If
'    Call InitAlleBereichsFarben
    ret% = True
End If

EditFarbe% = ret%

Call clsError.DefErrPop
End Function

Sub InitFont(frm As Object)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("InitFont")
Call clsError.DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case clsError.DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call ProjektForm.EndeDll
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim c As Control

On Error Resume Next

For Each c In frm.Controls
    If (c.Name = "picToolbar") Or (c.Name = "cmdToolbar") Then
        If (iNewLine) Then
            c.Font.Name = "Arial"   'RTrim$(iFontName$(0))
            c.Font.Size = iFontSize%(0)
        Else
            c.Font.Name = "Small Fonts"
            c.Font.Size = 7
        End If
    ElseIf (c.Container.Name = "picStartBack") Then
    Else
'        If (TypeOf c Is MSFlexGrid) Or (TypeOf c Is ListBox) Or (TypeOf c Is TextBox) Or (TypeOf c Is ComboBox) Then
        If (TypeOf c Is MSFlexGrid) Or (TypeOf c Is ListBox) Or (TypeOf c Is TextBox) Or (TypeOf c Is ComboBox) Then
            c.Font.Name = RTrim$(iFontName$(0))
            c.Font.Size = iFontSize%(0)
        Else
            c.Font.Name = RTrim$(iFontName$(1))
            c.Font.Size = iFontSize%(1)
        End If
        
        If (iNewLine) Then
            If (TypeOf c Is TextBox) Or (TypeOf c Is ComboBox) Then
                c.Font.Size = iFontSize%(0) + 2
            End If
        End If
'        Else
'            c.Font.Name = RTrim$(iFontName$(1))
'            c.Font.Size = iFontSize%(1)
'        End If
        If (TypeOf c Is TextBox) Or (TypeOf c Is Label) Or (TypeOf c Is CheckBox) Or (TypeOf c Is OptionButton) Then
            frm.Font.Name = c.Font.Name
            frm.Font.Size = c.Font.Size
            frm.Font.Bold = c.Font.Bold
            c.Height = frm.TextHeight("Äg") + 60
            If (TypeOf c Is TextBox) Then
            ElseIf (TypeOf c Is Label) Then
                c.Width = frm.TextWidth(c.Caption) + 90
            Else
                c.Width = frm.TextWidth(c.Caption) + 600
            End If
        End If
    End If
Next

frm.Font.Name = RTrim$(iFontName$(0))
frm.Font.Size = iFontSize%(0)
frm.Font.Bold = False

Call clsError.DefErrPop
End Sub

Function CreateDirectory%(DirName$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("CreateDirectory%")
Call clsError.DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case clsError.DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call ProjektForm.EndeDll
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim fTmp%, ret%
Dim TestDateiName$

On Error Resume Next

ret% = True

DirName$ = UCase$(RTrim$(DirName$))
TestDateiName$ = DirName$ + "\TEST.$$$"

fTmp% = FreeFile
Open TestDateiName$ For Random As #fTmp%

If Err <> 0 Then
    Err = 0
    MkDir DirName$
    If Err <> 0 Then
        Call clsDialog.MessageBox("Problem: Das Verzeichnis " + DirName$ + " kann nicht angelegt werden!" + vbCrLf + "Einspielvorgang wird abgebrochen !", vbExclamation)
        ret% = False
    End If
Else
    Close #fTmp%
    Kill TestDateiName$
End If

CreateDirectory% = ret%

Call clsError.DefErrPop
End Function

Function CalcFileSize&(FileName$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("CalcFileSize&")
Call clsError.DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case clsError.DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call ProjektForm.EndeDll
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim erg%
Dim SearchHandle&
Dim FindDataRec As WIN32_FIND_DATA

SearchHandle& = FindFirstFile(FileName$, FindDataRec)
If (SearchHandle& = INVALID_HANDLE_VALUE) Then
    CalcFileSize& = 0&
Else
    CalcFileSize& = FindDataRec.nFileSizeHigh * MAXDWORD + FindDataRec.nFileSizeLow
End If
erg% = FindClose(SearchHandle&)

Call clsError.DefErrPop
End Function

Sub InitEndSub(iForm As Object)
Set ProjektForm = iForm
End Sub

Sub ExitEndSub()
Set ProjektForm = Nothing
End Sub

'Sub FillGradient(FillControl As Object, x1%, y1%, x2%, y2%, color1&, color2&)
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Call clsError.DefErrFnc("FillGradient")
'Call clsError.DefErrMod(DefErrModul)
'On Error GoTo DefErr
'GoTo DefErrEnd
'DefErr:
'Select Case clsError.DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
'Case vbRetry
'  Resume
'Case vbIgnore
'  Resume Next
'End Select
'Call ProjektForm.EndeDll
'DefErrEnd:
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Dim R#
'Dim R1 As Integer
'Dim R2 As Integer
'Dim DeltaR#
'Dim G#
'Dim G1 As Integer
'Dim G2 As Integer
'Dim DeltaG#
'Dim b#
'Dim B1 As Integer
'Dim B2 As Integer
'Dim DeltaB#
'Dim N As Integer
'Dim Color As Long
'Dim XN As Integer
'Dim YN As Integer
'Dim Size%
'
'R1 = color1 And &HFF
'G1 = (color1 \ &H100) And &HFF
'B1 = (color1 \ &H10000) And &HFF
'
'Size = y2 - y1
'
''With ParentForm.picToolbar
'With FillControl
''    SetPixel .hdc, x, y, RGB(R, G, B)
'    R2 = color2 And &HFF
'    G2 = (color2 \ &H100) And &HFF
'    B2 = (color2 \ &H10000) And &HFF
'    DeltaR = (R2 - R1) / Size
'    DeltaG = (G2 - G1) / Size
'    DeltaB = (B2 - B1) / Size
'    R = R1: G = G1: b = B1
'    Do
'        N = N + 1
'        R = R + DeltaR: G = G + DeltaG: b = b + DeltaB
''        For XN = x1 To x2
''            SetPixel .hdc, XN, y1 + N, RGB(Abs(R), Abs(G), Abs(b))
''        Next XN
'        FillControl.Line (x1 * Screen.TwipsPerPixelX, (y1 + N) * Screen.TwipsPerPixelY)-(x2 * Screen.TwipsPerPixelX, (y1 + N + 1) * Screen.TwipsPerPixelY), RGB(Abs(R), Abs(G), Abs(b)), BF
'    Loop Until N = Size
''    .Refresh
'End With
'
'Call clsError.DefErrPop
'End Sub
'
'Function GetMainWnd(ByVal TaskName As String) As Long
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Call clsError.DefErrFnc("GetMainWnd")
'Call clsError.DefErrMod(DefErrModul)
'On Error GoTo DefErr
'GoTo DefErrEnd
'DefErr:
''If (Err.Number = (999 + vbObjectError)) Then End
'Select Case clsError.DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
'Case vbRetry
'  Resume
'Case vbIgnore
'  Resume Next
'End Select
'Call ProjektForm.EndeDll
''On Error GoTo 0
''Err.Raise 999 + vbObjectError, "OPWMENU.VBP", "Fehler in DLL"
'DefErrEnd:
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Dim CurrWnd As Long
'Dim Length As Integer
'Dim ListItem As String
'Dim x As Long
'Dim ThreadID As Long
'
'TaskName = UCase(TaskName)
'
'GetMainWnd = 0
'CurrWnd = GetWindow(ProjektForm.hWnd, GW_HWNDFIRST)
'Do While CurrWnd <> 0
'  Length = GetWindowTextLength(CurrWnd)
'  ListItem = Space(Length + 1)
'  Length = GetWindowText(CurrWnd, ListItem, Length + 1)
'  If Length > 0 Then
'    If UCase(Left(ListItem, Len(TaskName))) = TaskName Then
'      GetMainWnd = CurrWnd
'      Exit Do
'    End If
'  End If
'  CurrWnd = GetWindow(CurrWnd, GW_HWNDNEXT)
'Loop
'
'Call clsError.DefErrPop
'End Function
'
''Public Sub ControlBorderless(hwnd As Long, Optional ByVal IsWindow As Boolean, Optional ByVal DeflateX As Long = 2, Optional ByVal DeflateY As Long = 2)
'Public Sub ControlBorderless(iControl As Object, Optional ByVal DeflateX As Long = 2, Optional ByVal DeflateY As Long = 2)
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Call clsError.DefErrFnc("ControlBorderless")
'Call clsError.DefErrMod(DefErrModul)
'On Error GoTo DefErr
'GoTo DefErrEnd
'DefErr:
''If (Err.Number = (999 + vbObjectError)) Then End
'Select Case clsError.DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
'Case vbRetry
'  Resume
'Case vbIgnore
'  Resume Next
'End Select
'Call ProjektForm.EndeDll
''On Error GoTo 0
''Err.Raise 999 + vbObjectError, "OPWMENU.VBP", "Fehler in DLL"
'DefErrEnd:
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Dim nRect As RECT
'Dim nRgn As Long
'
'If (DeflateX < 0) Then
'    'rounded
'    If (TypeOf iControl Is Form) Then
'        nRect.Left = Abs(DeflateX)
'        nRect.Top = Abs(DeflateY)
'        nRect.Right = iControl.Width / Screen.TwipsPerPixelX - Abs(DeflateX)
'        nRect.Bottom = iControl.Height / Screen.TwipsPerPixelY - Abs(DeflateY)
'    Else
'        GetClientRect iControl.hWnd, nRect
'        InflateRect nRect, DeflateX, DeflateY
'    End If
'    With nRect
'        nRgn = CreateRoundRectRgn(.Left, .Top, .Right, .Bottom, 20, 20)
'    End With
'Else
'    If (TypeOf iControl Is Form) Then
'        nRect.Left = 0
'        nRect.Top = 0
'        nRect.Right = iControl.Width / Screen.TwipsPerPixelX
'        nRect.Bottom = iControl.Height / Screen.TwipsPerPixelY
'    Else
'        GetClientRect iControl.hWnd, nRect
'    End If
'
'    '        Die API-Funktion InflateRect verkleinert das Rechteck jeweils links und rechts und oben und unten um den gewünschten Betrag:
'    InflateRect nRect, -DeflateX, -DeflateY
'
'    '        Nun erstellen wir anhand dieses verkleinerten Rechtecks eine Region
'    nRgn = CreateRectRgnIndirect(nRect)
'End If
'
''        und weisen diese dem Fenster der ComboBox zu:
'SetWindowRgn iControl.hWnd, nRgn, True
'
'Call clsError.DefErrPop
'End Sub

Sub HoleProgFont()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("HoleProgFont")
Call clsError.DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case clsError.DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call ProjektForm.EndeDll
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim i%, ind%
Dim l&
Dim h$, key$, prog$

'For i% = 1 To 50
'    h$ = Space$(100)
'    key$ = "Font" + Format(i%, "00")
'    l& = GetPrivateProfileString(iIniSection$, key$, " ", h$, 101, iIniDatei$)
'    h$ = Trim$(UCase(h$))
'    If (Len(h$) <= 1) Then Exit For
'
'    prog$ = ""
'    h$ = Left$(h$, Len(h$) - 1)
'    ind% = InStr(h$, ",")
'    If (ind% > 0) Then
'        prog$ = RTrim$(Left$(h$, ind% - 1))
'        h$ = Mid$(h$, ind% + 1)
'    End If
'    If (prog$ <> "") Then
'        If (prog$ <> iProgName$) Then
'            prog$ = ""
'        End If
'    End If
'
'    If (prog$ <> "") Then
'        ind% = InStr(h$, ",")
'        iFontName$(0) = RTrim$(Left$(h$, ind% - 1))
'        iFontSize%(0) = Val(Mid$(h$, ind% + 1))
'        Exit For
'    End If
'Next i%

'If (iNewLine) Then
'    Dim hdc&
'    Dim tmp$
'    Dim iNonClientMetrics As NONCLIENTMETRICS
'
'    hdc = GetDC(0)
'    iNonClientMetrics.cbSize = Len(iNonClientMetrics)
'    l& = SystemParametersInfo(SPI_GETNONCLIENTMETRICS, Len(iNonClientMetrics), iNonClientMetrics, 0)
'    tmp = StrConv(iNonClientMetrics.lfMenuFont.lfFaceName, vbUnicode)
'    iFontName$(0) = Left$(tmp, InStr(tmp, Chr$(0)) - 1)
'    iFontSize%(0) = -MulMul(iNonClientMetrics.lfMenuFont.lfHeight, GetDeviceCaps(hdc, LOGPIXELSY), 72) '+ 2
'    ReleaseDC 0, hdc
'
'    h$ = Space$(100)
'    l& = GetPrivateProfileString("Allgemein", "Font", h$, h$, 101, CurDir + "\newline.ini")
'    h$ = Trim$(UCase(h$))
'    If (Len(h$) > 1) Then
'        ind% = InStr(h$, ",")
'        iFontName$(0) = RTrim$(Left$(h$, ind% - 1))
'        iFontSize%(0) = Val(Mid$(h$, ind% + 1))
'    End If
'Else
'    h$ = Space$(100)
'    key$ = "Font_" + iProgName$
'    l& = GetPrivateProfileString(iIniSection$, key$, " ", h$, 101, iIniDatei$)
'    h$ = Trim$(UCase(h$))
'    If (Len(h$) > 1) Then
'        ind% = InStr(h$, ",")
'        iFontName$(0) = RTrim$(Left$(h$, ind% - 1))
'        iFontSize%(0) = Val(Mid$(h$, ind% + 1))
'    End If
'End If

h$ = Space$(100)
key$ = "Font_" + iProgName$
If (iNewLine) Then
    l& = GetPrivateProfileString(iIniSection$, key$, " ", h$, 101, CurDir + "\newline.ini")
Else
    l& = GetPrivateProfileString(iIniSection$, key$, " ", h$, 101, iIniDatei$)
End If
h$ = Trim$(UCase(h$))
If (Len(h$) > 1) Then
    ind% = InStr(h$, ",")
    iFontName$(0) = RTrim$(Left$(h$, ind% - 1))
    iFontSize%(0) = Val(Mid$(h$, ind% + 1))
ElseIf (iNewLine) Then
    h$ = Space$(100)
    key$ = "Font"
    l& = GetPrivateProfileString(iIniSection$, key$, " ", h$, 101, CurDir + "\newline.ini")
    h$ = Trim$(UCase(h$))
    If (Len(h$) > 1) Then
        ind% = InStr(h$, ",")
        iFontName$(0) = RTrim$(Left$(h$, ind% - 1))
        iFontSize%(0) = Val(Mid$(h$, ind% + 1))
    End If
End If

iFontName$(1) = iFontName$(0)
iFontSize%(1) = iFontSize%(0)

Call clsError.DefErrPop
End Sub

Sub SpeicherProgFont(Index%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("SpeicherProgFont")
Call clsError.DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case clsError.DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call ProjektForm.EndeDll
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim l&
Dim h$, key$, sValue$

sValue$ = RTrim$(iFontName$(0)) + "," + Mid$(Str$(iFontSize%(0)), 2)
key$ = "Font_" + iProgName$
If (iNewLine) Then
    l& = WritePrivateProfileString(iIniSection$, key$, sValue$, CurDir + "\newline.ini")

    h$ = Space$(100)
    key$ = "Font"
    l& = GetPrivateProfileString(iIniSection$, key$, " ", h$, 101, CurDir + "\newline.ini")
    h$ = Trim$(UCase(h$))
    If (Len(h$) <= 1) Then
        l& = WritePrivateProfileString(iIniSection$, key$, sValue$, CurDir + "\newline.ini")
    End If
Else
    h$ = RTrim$(iFontName$(Index)) + "," + Mid$(Str$(iFontSize%(Index)), 2)
    If (Index = 0) Then
        l& = WritePrivateProfileString(iIniSection$, "FontInformation", h$, iIniDatei$)
    Else
        l& = WritePrivateProfileString(iIniSection$, "FontBeschriftung", h$, iIniDatei$)
    End If
'    l& = WritePrivateProfileString(iIniSection$, key$, sValue$, iIniDatei$)
End If

Call clsError.DefErrPop
End Sub

'Sub NewLineWindow(frm As Object, iButtonY%)
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Call clsError.DefErrFnc("NewLineWindow")
'Call clsError.DefErrMod(DefErrModul)
'On Error GoTo DefErr
'GoTo DefErrEnd
'DefErr:
'Select Case clsError.DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
'Case vbRetry
'  Resume
'Case vbIgnore
'  Resume Next
'End Select
'Call ProjektForm.EndeDll
'DefErrEnd:
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Dim x%, y%, iAdd2%, wi%
'
'iAdd2 = wPara1.NlCaptionY
'With frm
'    .BackColor = vbGreen ' RGB(180, 180, 180)
'
'    frm.Line (15, 15)-(.ScaleWidth - 30, 180), RGB(247, 247, 247), BF
'
'    .ForeColor = RGB(177, 177, 177)
'    .FillStyle = vbSolid
'    .FillColor = RGB(247, 247, 247)
'    RoundRect .hdc, 0, 0, .ScaleWidth / Screen.TwipsPerPixelX, .ScaleHeight / Screen.TwipsPerPixelY, 20, 20
'
'    Call wPara1.FillGradient(frm, 1, (180 * iFaktorY) / Screen.TwipsPerPixelY, .ScaleWidth / Screen.TwipsPerPixelX - 2, iAdd2 / Screen.TwipsPerPixelY, RGB(230, 230, 230), RGB(191, 191, 191))
'
'    y = iButtonY + 210 * iFaktorY
'    frm.Line (15, y)-(.ScaleWidth - 30, .ScaleHeight - 15), RGB(177, 177, 177), BF
'
'    Call wPara1.FillGradient(frm, 1, iAdd2 / Screen.TwipsPerPixelY, .ScaleWidth / Screen.TwipsPerPixelX - 2, y / Screen.TwipsPerPixelY, RGB(177, 177, 177), RGB(225, 225, 225))
'
'    y = iButtonY - 300 * iFaktorY
'    .ForeColor = RGB(177, 177, 177)
'    frm.Line (15, y)-(.ScaleWidth - 15, y)
'
'    wi = 16 * Screen.TwipsPerPixelX
'    x = .ScaleWidth - wi - 210 * iFaktorX
'    y = (iAdd2 - wi) / 2
'    If (.MinButton) Then
'        With .picControlBox(0)
'            .Left = x - 2 * (wi + 45)
'            .Top = y
'            .Visible = True
'        End With
'    End If
'    If (.MaxButton) Then
'        With .picControlBox(1)
'            .Left = x - (wi + 45)
'            .Top = y
'            .Visible = True
'        End With
'    End If
'    If (.ControlBox) Then
'        With .picControlBox(2)
'            .Left = x
'            .Top = y
'            .Visible = True
'        End With
'    End If
'
'    .Height = .Height + wPara1.FrmCaptionHeight + 1
'    Call wPara1.ControlBorderless(frm, -3, wPara1.FrmCaptionHeight / Screen.TwipsPerPixelY + 4)
'End With
'
'Call clsError.DefErrPop
'End Sub
'
'Sub InitForm(frm As Object, sUserSection$, sIniDatei$)
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Call clsError.DefErrFnc("InitForm")
'Call clsError.DefErrMod(DefErrModul)
'On Error GoTo DefErr
'GoTo DefErrEnd
'DefErr:
'Select Case clsError.DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
'Case vbRetry
'  Resume
'Case vbIgnore
'  Resume Next
'End Select
'Call ProjektForm.EndeDll
'DefErrEnd:
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Dim i%, ind%, iVal%
'Dim l&, f&, lVal&, MainWnd&
'Dim h$, h2$, key$
'Dim wRect As RECT
'
'With frm
'    'Optipharm NewLine
'    MainWnd = wPara1.GetMainWnd("Optipharm NewLine")
'    If (MainWnd <> 0) Then
'    '    MsgBox (wpara.FrmCaptionHeight)
'        l = GetWindowRect(MainWnd, wRect)
''        If (wpara.BildFaktor >= 1) Then
''            Me.Left = (wRect.Left + 10 + 32 + 5 - 15) * Screen.TwipsPerPixelX
''            Me.Top = (wRect.Top + 10 + 84 + 5 - 20) * Screen.TwipsPerPixelY - wpara.FrmCaptionHeight
''            Me.Width = (wRect.Left + (1517 - 32 - 10) + 30) * Screen.TwipsPerPixelX
''            Me.Height = (wRect.Top + (955 - 84 - 5) + 30) * Screen.TwipsPerPixelY + 2 * wpara.FrmCaptionHeight
''            Call wpara.ControlBorderless(Me, 3, wpara.FrmCaptionHeight / Screen.TwipsPerPixelY + 4)
''        Else
''            Me.Left = (wRect.Left + 2 + 15) * Screen.TwipsPerPixelX
''            Me.Top = (wRect.Top + 47) * Screen.TwipsPerPixelY - wpara.FrmCaptionHeight
''            Me.Width = (wRect.Left + 996) * Screen.TwipsPerPixelX
''            Me.Height = (wRect.Top + 598) * Screen.TwipsPerPixelY + 2 * wpara.FrmCaptionHeight
''            Call wpara.ControlBorderless(Me, 3, wpara.FrmCaptionHeight / Screen.TwipsPerPixelY + 3)
''        End If
'        .Left = (wRect.Left + 10) * Screen.TwipsPerPixelX
'        .Top = (wRect.Top + 10 + 33) * Screen.TwipsPerPixelX + wPara1.NlCaptionY
'        .Width = (wRect.Right - wRect.Left - 2 * (10)) * Screen.TwipsPerPixelX
'        .Height = (wRect.Bottom - 25) * Screen.TwipsPerPixelX - .Top
'
'        .Left = .Left - 3 * Screen.TwipsPerPixelX
'        .Width = .Width + 2 * 3 * Screen.TwipsPerPixelX
'        .Top = .Top - (3 * Screen.TwipsPerPixelX + wPara1.FrmCaptionHeight)
'        .Height = .Height + 2 * (3 * Screen.TwipsPerPixelX + wPara1.FrmCaptionHeight)
'        Call wPara1.ControlBorderless(frm, 3, wPara1.FrmCaptionHeight / Screen.TwipsPerPixelY + 3)
'
'        Call SetWindowPos(MainWnd, .hWnd, 0, 0, 0, 0, 3)
'
'        .mnuDatei.Caption = Space(2) + .mnuDatei.Caption
'    Else
'        .Height = wPara1.WorkAreaHeight
'        If (wPara1.BildFaktor = 0.8!) Then
'            .Width = 10980 * wPara1.BildFaktor '10200
'        Else
'            .Width = .Width * wPara1.BildFaktor
'        End If
'
'        iVal% = (Screen.Width - .Width) / 2
'        If (iVal% < 0) Then
'            iVal% = 0
'        End If
'        h$ = Format(iVal%, "00000")
'        l& = GetPrivateProfileString(sUserSection$, "StartX", h$, h$, 6, sIniDatei)
'        h$ = Left$(h$, l&)
'        lVal& = Val(h$)
'        If (lVal& > 30000) Then
'            iVal% = -9999
'        Else
'            iVal% = lVal&
'        End If
'        If (iVal% = -9999) Then
'            .WindowState = vbMaximized
'        Else
'            If (iVal% < 0) Then
'                iVal% = 0
'            End If
'            .Left = iVal%
'
'            iVal% = 75
'            h$ = Format(iVal%, "00000")
'            l& = GetPrivateProfileString(sUserSection$, "StartY", h$, h$, 6, sIniDatei)
'            h$ = Left$(h$, l&)
'            iVal% = Val(h$)
'            If (iVal% < 0) Then
'                iVal% = 0
'            End If
'            .Top = iVal%
'
'            iVal% = .Width
'            h$ = Format(iVal%, "00000")
'            l& = GetPrivateProfileString(sUserSection$, "BreiteX", h$, h$, 6, sIniDatei)
'            h$ = Left$(h$, l&)
'            iVal% = Val(h$)
'            If (iVal% < 0) Then
'                iVal% = 0
'            End If
'            If (.Left + iVal% > wPara1.WorkAreaWidth) Then
'                .WindowState = vbMaximized
'            Else
'                .Width = iVal%
'
'                iVal% = .Height
'                h$ = Format(iVal%, "00000")
'                l& = GetPrivateProfileString(sUserSection$, "HoeheY", h$, h$, 6, sIniDatei)
'                h$ = Left$(h$, l&)
'                iVal% = Val(h$)
'                If (iVal% < 0) Then
'                    iVal% = 0
'                End If
'                If (.Top + iVal% > wPara1.WorkAreaHeight) Then
'                    .WindowState = vbMaximized
'                Else
'                    .Height = iVal%
'                End If
'            End If
'        End If
'    End If
'End With
'
'Call clsError.DefErrPop
'End Sub

Sub FillGradient(FillControl As Object, x1%, y1%, x2%, Y2%, color1&, color2&)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("FillGradient")
Call clsError.DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case clsError.DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call ProjektForm.EndeDll
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim R#
Dim R1 As Integer
Dim R2 As Integer
Dim DeltaR#
Dim G#
Dim G1 As Integer
Dim G2 As Integer
Dim DeltaG#
Dim b#
Dim B1 As Integer
Dim B2 As Integer
Dim DeltaB#
Dim N As Integer
Dim Color As Long
Dim XN As Integer
Dim YN As Integer
Dim Size%

R1 = color1 And &HFF
G1 = (color1 \ &H100) And &HFF
B1 = (color1 \ &H10000) And &HFF

Size = Y2 - y1

'With ParentForm.picToolbar
With FillControl
'    SetPixel .hdc, x, y, RGB(R, G, B)
    R2 = color2 And &HFF
    G2 = (color2 \ &H100) And &HFF
    B2 = (color2 \ &H10000) And &HFF
    DeltaR = (R2 - R1) / Size
    DeltaG = (G2 - G1) / Size
    DeltaB = (B2 - B1) / Size
    R = R1: G = G1: b = B1
    N = -1
    Do
        N = N + 1
        R = R + DeltaR: G = G + DeltaG: b = b + DeltaB
'        For XN = x1 To x2
'            SetPixel .hdc, XN, y1 + N, RGB(Abs(R), Abs(G), Abs(b))
'        Next XN
        FillControl.Line (x1 * Screen.TwipsPerPixelX, (y1 + N) * Screen.TwipsPerPixelY)-(x2 * Screen.TwipsPerPixelX, (y1 + N + 1) * Screen.TwipsPerPixelY), RGB(Abs(R), Abs(G), Abs(b)), BF
    Loop Until N = Size
'    .Refresh
End With

Call clsError.DefErrPop
End Sub

Function GetMainWnd(ByVal TaskName As String) As Long
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("GetMainWnd")
Call clsError.DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
'If (Err.Number = (999 + vbObjectError)) Then End
Select Case clsError.DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call ProjektForm.EndeDll
'On Error GoTo 0
'Err.Raise 999 + vbObjectError, "OPWMENU.VBP", "Fehler in DLL"
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim CurrWnd As Long
Dim Length As Integer
Dim ListItem As String
Dim x As Long
Dim ThreadID As Long

TaskName = UCase(TaskName)

GetMainWnd = 0
'CurrWnd = GetWindow(ProjektForm.hWnd, GW_HWNDFIRST)
CurrWnd = GetWindow(GetForegroundWindow, GW_HWNDFIRST)
Do While CurrWnd <> 0
  Length = GetWindowTextLength(CurrWnd)
  ListItem = Space(Length + 1)
  Length = GetWindowText(CurrWnd, ListItem, Length + 1)
  If Length > 0 Then
    If UCase(Left(ListItem, Len(TaskName))) = TaskName Then
      GetMainWnd = CurrWnd
      Exit Do
    End If
  End If
  CurrWnd = GetWindow(CurrWnd, GW_HWNDNEXT)
Loop

Call clsError.DefErrPop
End Function

'Public Sub ControlBorderless(hwnd As Long, Optional ByVal IsWindow As Boolean, Optional ByVal DeflateX As Long = 2, Optional ByVal DeflateY As Long = 2)
Public Sub ControlBorderless(iControl As Object, Optional ByVal DeflateX As Long = 2, Optional ByVal DeflateY As Long = 2)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("ControlBorderless")
Call clsError.DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
'If (Err.Number = (999 + vbObjectError)) Then End
Select Case clsError.DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call ProjektForm.EndeDll
'On Error GoTo 0
'Err.Raise 999 + vbObjectError, "OPWMENU.VBP", "Fehler in DLL"
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim nRect As RECT
Dim nRgn As Long
  
If (DeflateX < 0) Then
    'rounded
    If (TypeOf iControl Is Form) Then
        nRect.Left = Abs(DeflateX)
        nRect.Top = Abs(DeflateY)
        nRect.Right = iControl.Width / Screen.TwipsPerPixelX - Abs(DeflateX)
        nRect.Bottom = iControl.Height / Screen.TwipsPerPixelY - Abs(DeflateY)
    Else
        GetClientRect iControl.hWnd, nRect
    End If
    With nRect
        nRgn = CreateRoundRectRgn(.Left, .Top, .Right, .Bottom, 20, 20)
    End With
Else
    If (TypeOf iControl Is Form) Then
        nRect.Left = 0
        nRect.Top = 0
        nRect.Right = iControl.Width / Screen.TwipsPerPixelX
        nRect.Bottom = iControl.Height / Screen.TwipsPerPixelY
    Else
        GetClientRect iControl.hWnd, nRect
    End If
        
    '        Die API-Funktion InflateRect verkleinert das Rechteck jeweils links und rechts und oben und unten um den gewünschten Betrag:
    InflateRect nRect, -DeflateX, -DeflateY
        
    '        Nun erstellen wir anhand dieses verkleinerten Rechtecks eine Region
    nRgn = CreateRectRgnIndirect(nRect)
End If
    
'        und weisen diese dem Fenster der ComboBox zu:
SetWindowRgn iControl.hWnd, nRgn, True

Call clsError.DefErrPop
End Sub

'Sub NewLineWindow(frm As Object, iButtonY%)
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Call clsError.DefErrFnc("NewLineWindow")
'Call clsError.DefErrMod(DefErrModul)
'On Error GoTo DefErr
'GoTo DefErrEnd
'DefErr:
'Select Case clsError.DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
'Case vbRetry
'  Resume
'Case vbIgnore
'  Resume Next
'End Select
'Call ProjektForm.EndeDll
'DefErrEnd:
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Dim x%, y%, iAdd2%, wi%
'
'iAdd2 = wPara1.NlCaptionY
'With frm
''    .BackColor = vbGreen ' RGB(180, 180, 180)
'
'    frm.Line (15, 15)-(.ScaleWidth - 30, 180), RGB(247, 247, 247), BF
'
'    .ForeColor = RGB(177, 177, 177)
'    .FillStyle = vbSolid
'    .FillColor = RGB(247, 247, 247)
'    RoundRect .hdc, 0, 0, .ScaleWidth / Screen.TwipsPerPixelX, .ScaleHeight / Screen.TwipsPerPixelY, 20, 20
'    .Refresh
'
'    Call wPara1.FillGradient(frm, 1, (180 * iFaktorY) / Screen.TwipsPerPixelY, .ScaleWidth / Screen.TwipsPerPixelX - 2, iAdd2 / Screen.TwipsPerPixelY, RGB(230, 230, 230), RGB(191, 191, 191))
'    .Refresh
'
'    y = iButtonY + 210 * iFaktorY
'    frm.Line (15, y)-(.ScaleWidth - 30, .ScaleHeight - 15), RGB(177, 177, 177), BF
'
'    Call wPara1.FillGradient(frm, 1, iAdd2 / Screen.TwipsPerPixelY, .ScaleWidth / Screen.TwipsPerPixelX - 2, y / Screen.TwipsPerPixelY, RGB(177, 177, 177), RGB(225, 225, 225))
'    .Refresh
'
'    y = iButtonY - 300 * iFaktorY
'    .ForeColor = RGB(177, 177, 177)
'    frm.Line (15, y)-(.ScaleWidth - 15, y)
'
'    If (.ControlBox) Then
'        wi = 16 * Screen.TwipsPerPixelX
'        x = .ScaleWidth - wi - 210 * iFaktorX
'        y = (iAdd2 - wi) / 2
'        If (.MinButton) Then
'            With .picControlBox(0)
'                .Left = x - 2 * (wi + 45)
'                .Top = y
'                .Visible = True
'            End With
'        End If
'        If (.MaxButton) Then
'            With .picControlBox(1)
'                .Left = x - (wi + 45)
'                .Top = y
'                .Visible = True
'            End With
'        End If
'        With .picControlBox(2)
'            .Left = x
'            .Top = y
'            .Visible = True
'        End With
'    End If
'
'    .Height = .Height + wPara1.FrmCaptionHeight + 1
'    Call wPara1.ControlBorderless(frm, -3, wPara1.FrmCaptionHeight / Screen.TwipsPerPixelY + 4)
'    .Refresh
'
'    Call NewLineFrame(frm)
'    .Refresh
'End With
'
'Call clsError.DefErrPop
'End Sub

Sub NewLineWindow(frm As Object, iButtonY%, Optional FormLoad% = True)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("NewLineWindow")
Call clsError.DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case clsError.DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call ProjektForm.EndeDll
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim x%, y%, iAdd2%, wi%, ind%
Dim lColor&(1)
Dim h$
Dim c As Control

iAdd2 = wPara1.NlCaptionY
With frm
'    .BackColor = vbGreen ' RGB(180, 180, 180)

    frm.Line (15, 15)-(.ScaleWidth - 30, 180), RGB(247, 247, 247), BF
    
    .ForeColor = RGB(177, 177, 177)
    .FillStyle = vbSolid
    .FillColor = RGB(247, 247, 247)
'    RoundRect .hdc, 0, 0, .ScaleWidth / Screen.TwipsPerPixelX, .ScaleHeight / Screen.TwipsPerPixelY, 20, 20
    .DrawWidth = 2
    frm.Line (60, 0)-(.ScaleWidth - 60, 0)
    frm.Line (0, 60)-(0, .ScaleHeight - 60)
    frm.Line (.ScaleWidth, 60)-(.ScaleWidth, .ScaleHeight - 60)
    .DrawWidth = 1

    Call wPara1.FillGradient(frm, 1, (180 * iFaktorY) / Screen.TwipsPerPixelY, .ScaleWidth / Screen.TwipsPerPixelX - 2, iAdd2 / Screen.TwipsPerPixelY, RGB(230, 230, 230), RGB(191, 191, 191))

    y = iButtonY + 210 * iFaktorY
    frm.Line (15, y)-(.ScaleWidth - 30, .ScaleHeight - 15), RGB(177, 177, 177), BF

    Call wPara1.FillGradient(frm, 1, iAdd2 / Screen.TwipsPerPixelY, .ScaleWidth / Screen.TwipsPerPixelX - 2, y / Screen.TwipsPerPixelY, RGB(177, 177, 177), RGB(225, 225, 225))

    y = iButtonY - 300 * iFaktorY
    .ForeColor = RGB(177, 177, 177)
    frm.Line (15, y)-(.ScaleWidth - 15, y)

    If (.ControlBox) Then
        wi = 16 * Screen.TwipsPerPixelX
        x = .ScaleWidth - wi - 210 * iFaktorX
        y = (iAdd2 - wi) / 2
        If (.MinButton) Then
            With .picControlBox(0)
                .Left = x - 2 * (wi + 45)
                .Top = y
                .Visible = True
            End With
        End If
        If (.MaxButton) Then
            With .picControlBox(1)
                .Left = x - (wi + 45)
                .Top = y
                .Visible = True
            End With
        End If
        With .picControlBox(2)
            .Left = x
            .Top = y
            .Visible = True
        End With
    End If

    If (FormLoad) Then
        .Height = .Height + wPara1.FrmCaptionHeight + 1
        Call wPara1.ControlBorderless(frm, -7, wPara1.FrmCaptionHeight / Screen.TwipsPerPixelY + 8) '+ 4
'        MsgBox ("BEEEEEEEEP: " + frm.Name)
    End If
    
    Call NewLineFrame(frm)

    On Error Resume Next
    For Each c In frm.Controls
        If (c.Tag <> "0") Then
            If (TypeOf c Is Label) And (FormLoad) Then
                c.BackStyle = 0 'duchsichtig
            ElseIf (TypeOf c Is TextBox) Or (TypeOf c Is ComboBox) Then
                If (TypeOf c Is ComboBox) Then
                    Call wPara1.ControlBorderless(c)
                ElseIf (c.Appearance = 1) Then
                    Call wPara1.ControlBorderless(c, 2, 2)
                Else
                    Call wPara1.ControlBorderless(c, 1, 1)
                End If
                
                If (c.Enabled) Then
                    c.BackColor = vbWhite
                Else
                    c.BackColor = frm.BackColor
                End If
                
'                If (c.Visible) Then
                    With c.Container
                        .ForeColor = RGB(180, 180, 180) ' vbWhite
                        .FillStyle = vbSolid
                        .FillColor = c.BackColor
        
                        RoundRect .hdc, (c.Left - 60) / Screen.TwipsPerPixelX, (c.Top - 30) / Screen.TwipsPerPixelY, (c.Left + c.Width + 60) / Screen.TwipsPerPixelX, (c.Top + c.Height + 15) / Screen.TwipsPerPixelY, 10, 10
                    End With
'                End If
            ElseIf (TypeOf c Is CheckBox) Or (TypeOf c Is OptionButton) Then
                With c
                    .BackColor = GetPixel(.Container.hdc, .Left / Screen.TwipsPerPixelX - 2, .Top / Screen.TwipsPerPixelY)
                    .Height = 0
                    .Width = .Height * 3 / 4 '- 30
                End With
            End If
        End If
    Next
    On Error GoTo DefErr
    
    .FillColor = nlRoundRectColor
End With

Call clsError.DefErrPop
End Sub

Sub NewLineFrame(frm As Object)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("NewLineFrame")
Call clsError.DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case clsError.DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call ProjektForm.EndeDll
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim y%, ind%
Dim lColor&(1)
Dim h$
Dim c As Control
    
For Each c In frm.Controls
    If (Left(c.Name, 6) = "picfme") Then
        With c
            If (iNewLine) Then
                lColor(0) = GetPixel(.Container.hdc, c.Left / Screen.TwipsPerPixelX - 2, c.Top / Screen.TwipsPerPixelY)
                lColor(1) = GetPixel(.Container.hdc, c.Left / Screen.TwipsPerPixelX - 2, (c.Top + c.Height) / Screen.TwipsPerPixelY)
    '                Call wpara.FillGradient(c, 0, 0, .ScaleWidth / Screen.TwipsPerPixelX, .ScaleHeight / Screen.TwipsPerPixelY, RGB(177, 177, 177), RGB(225, 225, 225))
                If (lColor(1) <> -1) Then
                    Call wPara1.FillGradient(c, 0, 0, .ScaleWidth / Screen.TwipsPerPixelX, .ScaleHeight / Screen.TwipsPerPixelY, lColor(0), lColor(1))
                End If
'                If (c.Name = "picfmeLieferant") Then
'                    Debug.Print c.Name + Str(Timer) + Str(lColor(0)) + Str(lColor(1))
'                End If
            Else
                .BackColor = frm.BackColor
            End If
            
            .ForeColor = vbBlack ' RGB(150, 150, 150)
            .FillStyle = vbFSTransparent ' vbSolid
            .FillColor = vbWhite
        
            y = 30 + .TextHeight(.Tag) / 2
            If (iNewLine) Then
                RoundRect .hdc, 30 / Screen.TwipsPerPixelX, y / Screen.TwipsPerPixelY, (.Width - 30) / Screen.TwipsPerPixelX, (.Height - 30) / Screen.TwipsPerPixelY, 20, 20
            Else
                RoundRect .hdc, 30 / Screen.TwipsPerPixelX, y / Screen.TwipsPerPixelY, (.Width - 30) / Screen.TwipsPerPixelX, (.Height - 30) / Screen.TwipsPerPixelY, 0, 0
            End If
            
            h = .Tag
            If (h <> "") Then
                ind = InStr(h, "&")
                If (ind > 0) Then
                    h = Left(h, ind - 1) + Mid(h, ind + 1)
                End If
                c.Line (240, y)-(300 + .TextWidth(h) + 90, y), RGB(180, 180, 180)
                
                .TabStop = False
                
                .CurrentX = 300
                .CurrentY = 30
    '                If (ind > 0) Then
    '                    c.Print Left(h$, ind - 1);
    '                    .FontUnderline = True
    '                    c.Print Mid(h$, ind, 1);
    '                    .FontUnderline = False
    '                    c.Print Mid(h$, ind + 1)
    '                Else
                    c.Print h$
    '                End If
            End If
    
            .FillColor = nlRoundRectColor
        End With
    End If
Next

Call clsError.DefErrPop
End Sub

Sub picBackPaint(nlPicBack As Object, Optional iInfoZusatz% = 0)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("picBackPaint")
Call clsError.DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case clsError.DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call ProjektForm.EndeDll
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim x1&, x2&, y1&, Y2&, iAdd&, iAdd2&

iAdd = wPara1.NlFlexBackY
iAdd2 = 120 * wPara1.FaktorY

With nlPicBack.Container
    x1 = .flxarbeit(0).Left - iAdd
    x2 = .flxarbeit(0).Left + .flxarbeit(0).Width + iAdd
    y1 = .flxInfo(0).Top - iAdd
'    y2 = ParentForm.flxInfo(0).Top + ParentForm.flxInfo(0).Height
    
    Y2 = nlPicBack.Height - (70 * wPara1.FaktorY - 10) * Screen.TwipsPerPixelY
    nlPicBack.Line (0, Y2)-(nlPicBack.Width, nlPicBack.Height), RGB(180, 180, 180), BF
    nlPicBack.ForeColor = RGB(180, 180, 180) ' vbWhite
    nlPicBack.FillStyle = vbSolid
    nlPicBack.FillColor = nlRoundRectColor  ' vbWhite
    If (.flxInfo(0).Tag <> "0") Then
        RoundRect nlPicBack.hdc, x1 / Screen.TwipsPerPixelX, (.flxInfo(0).Top - iAdd) / Screen.TwipsPerPixelY, x2 / Screen.TwipsPerPixelX, (.flxInfo(0).Top + .flxInfo(0).Height + iAdd) / Screen.TwipsPerPixelY, 20, 20
    End If
    If (iInfoZusatz) Then
        RoundRect nlPicBack.hdc, x1 / Screen.TwipsPerPixelX, (.flxInfoZusatz(0).Top - iAdd) / Screen.TwipsPerPixelY, x2 / Screen.TwipsPerPixelX, (.flxInfoZusatz(0).Top + .flxInfoZusatz(0).Height + iAdd) / Screen.TwipsPerPixelY, 20, 20
        y1 = .flxInfoZusatz(0).Top - iAdd
    End If
    If (.flxarbeit(0).Tag <> "0") Then
        RoundRect nlPicBack.hdc, x1 / Screen.TwipsPerPixelX, (.flxarbeit(0).Top - iAdd) / Screen.TwipsPerPixelY, x2 / Screen.TwipsPerPixelX, (y1 - wPara1.UntenFreiY) / Screen.TwipsPerPixelY, 20, 20
    End If
End With

Call clsError.DefErrPop
End Sub

Sub NewLineRefresh(nlForm As Object, iButtonY%, nlContainer As Object)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("NewLineRefresh")
Call clsError.DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case clsError.DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call ProjektForm.EndeDll
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim y%, iAdd%, iAdd2%, ok%
Dim lColor&(1)
Dim c As Control

With nlForm
    If (TypeOf nlContainer Is Form) Then
        y = iButtonY + 210 * iFaktorY
        iAdd2 = wPara1.NlCaptionY
        nlForm.Line (15, y)-(.ScaleWidth - 30, .ScaleHeight - 15), RGB(177, 177, 177), BF
        Call wPara1.FillGradient(nlForm, 1, iAdd2 / Screen.TwipsPerPixelY, .ScaleWidth / Screen.TwipsPerPixelX - 2, y / Screen.TwipsPerPixelY, RGB(177, 177, 177), RGB(225, 225, 225))
    Else
        nlContainer.Cls
    End If

    On Error Resume Next
    For Each c In nlForm.Controls
        If (Left(c.Name, 6) = "picfme") Then
            With c
                lColor(0) = GetPixel(.Container.hdc, c.Left / Screen.TwipsPerPixelX - 2, c.Top / Screen.TwipsPerPixelY)
                lColor(1) = GetPixel(.Container.hdc, c.Left / Screen.TwipsPerPixelX - 2, (c.Top + c.Height) / Screen.TwipsPerPixelY)
'                Call wpara.FillGradient(c, 0, 0, .ScaleWidth / Screen.TwipsPerPixelX, .ScaleHeight / Screen.TwipsPerPixelY, RGB(177, 177, 177), RGB(225, 225, 225))
                Call wPara1.FillGradient(c, 0, 0, .ScaleWidth / Screen.TwipsPerPixelX, .ScaleHeight / Screen.TwipsPerPixelY, lColor(0), lColor(1))
                
                .ForeColor = vbBlack ' RGB(150, 150, 150)
                .FillStyle = vbFSTransparent ' vbSolid
                .FillColor = vbWhite
            
                y = 30 + .TextHeight(.Tag) / 2
                RoundRect .hdc, 30 / Screen.TwipsPerPixelX, y / Screen.TwipsPerPixelY, (.Width - 30) / Screen.TwipsPerPixelX, (.Height - 30) / Screen.TwipsPerPixelY, 20, 20
                
                c.Line (240, y)-(300 + .TextWidth(.Tag) + 90, y), RGB(180, 180, 180)
                
                .CurrentX = 300
                .CurrentY = 30
                c.Print .Tag
            End With
        End If
    Next
    
    For Each c In nlForm.Controls
        If (c.Tag <> "0") Then
            ok = False
            If (TypeOf nlContainer Is Form) Then
                ok = True
            Else
                ok = (c.Container Is nlContainer)
            End If
            If (ok) Then
                If (TypeOf c Is TextBox) Or (TypeOf c Is ComboBox) Or (TypeOf c Is ListBox) Then
                    If (TypeOf c Is ComboBox) Or (TypeOf c Is ListBox) Then
                        Call wPara1.ControlBorderless(c)
                    ElseIf (c.Appearance = 1) Then
                        Call wPara1.ControlBorderless(c, 2, 2)
                    Else
                        Call wPara1.ControlBorderless(c, 1, 1)
                    End If
    
                    If (c.Enabled) And (c.Locked = 0) Then
                        c.BackColor = vbWhite
                    Else
                        c.BackColor = nlForm.BackColor
                    End If
    
                    With c.Container
                        .ForeColor = RGB(180, 180, 180) ' vbWhite
                        .FillStyle = vbSolid
                        .FillColor = c.BackColor
    
                        If (TypeOf c Is ListBox) Then
                            RoundRect .hdc, (c.Left - 60) / Screen.TwipsPerPixelX, (c.Top - 60) / Screen.TwipsPerPixelY, (c.Left + c.Width + 60) / Screen.TwipsPerPixelX, (c.Top + c.Height + 45) / Screen.TwipsPerPixelY, 10, 10
                        Else
                            RoundRect .hdc, (c.Left - 60) / Screen.TwipsPerPixelX, (c.Top - 30) / Screen.TwipsPerPixelY, (c.Left + c.Width + 60) / Screen.TwipsPerPixelX, (c.Top + c.Height + 15) / Screen.TwipsPerPixelY, 10, 10
                        End If
                    End With
                ElseIf (TypeOf c Is MSFlexGrid) Then
                    If (c.Tag = "1") Then
                        iAdd = 120   'wpara.NlFlexBackY
                        With c
                            RoundRect .Container.hdc, (.Left - iAdd) / Screen.TwipsPerPixelX, (.Top - iAdd) / Screen.TwipsPerPixelY, (.Left + .Width + iAdd) / Screen.TwipsPerPixelX, (.Top + .Height + iAdd) / Screen.TwipsPerPixelY, 20, 20
                        End With
                    End If
                End If
            End If
        End If
    Next
    nlContainer.Refresh
    On Error GoTo DefErr
End With

Call clsError.DefErrPop
End Sub

Sub InitForm(frm As Object, sUserSection$, sIniDatei$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("InitForm")
Call clsError.DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case clsError.DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call ProjektForm.EndeDll
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim i%, ind%, iVal%
Dim l&, f&, lVal&, MainWnd&
Dim h$, h2$, key$
Dim wRect As RECT
Dim wRect2 As RECT

With frm
    'Optipharm NewLine
    MainWnd = wPara1.GetMainWnd("Optipharm NewLine")
    If (MainWnd <> 0) Then
    '    MsgBox (wpara.FrmCaptionHeight)
        l = GetWindowRect(MainWnd, wRect)
'        If (wpara.BildFaktor >= 1) Then
'            Me.Left = (wRect.Left + 10 + 32 + 5 - 15) * Screen.TwipsPerPixelX
'            Me.Top = (wRect.Top + 10 + 84 + 5 - 20) * Screen.TwipsPerPixelY - wpara.FrmCaptionHeight
'            Me.Width = (wRect.Left + (1517 - 32 - 10) + 30) * Screen.TwipsPerPixelX
'            Me.Height = (wRect.Top + (955 - 84 - 5) + 30) * Screen.TwipsPerPixelY + 2 * wpara.FrmCaptionHeight
'            Call wpara.ControlBorderless(Me, 3, wpara.FrmCaptionHeight / Screen.TwipsPerPixelY + 4)
'        Else
'            Me.Left = (wRect.Left + 2 + 15) * Screen.TwipsPerPixelX
'            Me.Top = (wRect.Top + 47) * Screen.TwipsPerPixelY - wpara.FrmCaptionHeight
'            Me.Width = (wRect.Left + 996) * Screen.TwipsPerPixelX
'            Me.Height = (wRect.Top + 598) * Screen.TwipsPerPixelY + 2 * wpara.FrmCaptionHeight
'            Call wpara.ControlBorderless(Me, 3, wpara.FrmCaptionHeight / Screen.TwipsPerPixelY + 3)
'        End If
        
'        .Left = (wRect.Left + 10) * Screen.TwipsPerPixelX
'        .Top = (wRect.Top + 10 + 33) * Screen.TwipsPerPixelX + wPara1.NlCaptionY
'        .Width = (wRect.Right - wRect.Left - 2 * (10)) * Screen.TwipsPerPixelX
'        .Height = (wRect.Bottom - 25) * Screen.TwipsPerPixelX - .Top
'
'        .Left = .Left - 3 * Screen.TwipsPerPixelX
'        .Width = .Width + 2 * 3 * Screen.TwipsPerPixelX
'        .Top = .Top - (3 * Screen.TwipsPerPixelX + wPara1.FrmCaptionHeight)
'        .Height = .Height + 2 * (3 * Screen.TwipsPerPixelX + wPara1.FrmCaptionHeight)
'        Call wPara1.ControlBorderless(frm, 3, wPara1.FrmCaptionHeight / Screen.TwipsPerPixelY + 3)
        
        .Left = (wRect.Left) * Screen.TwipsPerPixelX
        .Top = (wRect.Top + 33) * Screen.TwipsPerPixelX + wPara1.NlCaptionY
        .Width = (wRect.Right - wRect.Left) * Screen.TwipsPerPixelX
        .Height = (wRect.Bottom - 13) * Screen.TwipsPerPixelX - .Top

        l& = SystemParametersInfo(SPI_GETWORKAREA, vbNull, wRect2, 0)
        If (wRect2.Bottom < (wRect.Bottom - 5)) Then
            .Height = (wRect2.Bottom) * Screen.TwipsPerPixelX - .Top
        End If
        
        Dim iAddX%, iAddY%
        iAddX% = ((.Width - .ScaleWidth) / 2) / Screen.TwipsPerPixelY - 1
        iAddY% = iAddX  '7 '3
        .Left = .Left - iAddX * Screen.TwipsPerPixelX
        .Width = .Width + 2 * iAddX * Screen.TwipsPerPixelX
        .Top = .Top - (iAddY * Screen.TwipsPerPixelX + wPara1.FrmCaptionHeight)
        .Height = .Height + 2 * (iAddY * Screen.TwipsPerPixelX + wPara1.FrmCaptionHeight)
        Call wPara1.ControlBorderless(frm, iAddX, wPara1.FrmCaptionHeight / Screen.TwipsPerPixelY + iAddY) '+ 3
        
        Call SetWindowPos(MainWnd, .hWnd, 0, 0, 0, 0, 3)
        
        .mnuDatei.Caption = Space(2) + .mnuDatei.Caption
    Else
        .Height = wPara1.WorkAreaHeight
        If (wPara1.BildFaktor = 0.8!) Then
            .Width = 10980 * wPara1.BildFaktor '10200
        Else
            .Width = .Width * wPara1.BildFaktor
        End If
        
        iVal% = (Screen.Width - .Width) / 2
        If (iVal% < 0) Then
            iVal% = 0
        End If
        h$ = Format(iVal%, "00000")
        l& = GetPrivateProfileString(sUserSection$, "StartX", h$, h$, 6, sIniDatei)
        h$ = Left$(h$, l&)
        lVal& = Val(h$)
        If (lVal& > 30000) Then
            iVal% = -9999
        Else
            iVal% = lVal&
        End If
        If (iVal% = -9999) Then
            .WindowState = vbMaximized
        Else
            If (iVal% < 0) Then
                iVal% = 0
            End If
            .Left = iVal%
            
            iVal% = 75
            h$ = Format(iVal%, "00000")
            l& = GetPrivateProfileString(sUserSection$, "StartY", h$, h$, 6, sIniDatei)
            h$ = Left$(h$, l&)
            iVal% = Val(h$)
            If (iVal% < 0) Then
                iVal% = 0
            End If
            .Top = iVal%
            
            iVal% = .Width
            h$ = Format(iVal%, "00000")
            l& = GetPrivateProfileString(sUserSection$, "BreiteX", h$, h$, 6, sIniDatei)
            h$ = Left$(h$, l&)
            iVal% = Val(h$)
            If (iVal% < 0) Then
                iVal% = 0
            End If
            If (.Left + iVal% > wPara1.WorkAreaWidth) Then
                .WindowState = vbMaximized
            Else
                .Width = iVal%
            
                iVal% = .Height
                h$ = Format(iVal%, "00000")
                l& = GetPrivateProfileString(sUserSection$, "HoeheY", h$, h$, 6, sIniDatei)
                h$ = Left$(h$, l&)
                iVal% = Val(h$)
                If (iVal% < 0) Then
                    iVal% = 0
                End If
                If (.Top + iVal% > wPara1.WorkAreaHeight) Then
                    .WindowState = vbMaximized
                Else
                    .Height = iVal%
                End If
            End If
        End If
    End If
End With

Call clsError.DefErrPop
End Sub

Function TextBoxWidth&(frm As Object, s$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("TextBoxWidth&")
Call clsError.DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case clsError.DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call ProjektForm.EndeDll
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim ret&

With frm
    If (iNewLine) Then
        .Font.Size = iFontSize(0) + 2
        ret = .TextWidth(s)
        .Font.Size = iFontSize(0)
    Else
        ret = .TextWidth(s)
    End If
End With

TextBoxWidth = ret

Call clsError.DefErrPop
End Function

Sub EnumeratePrinters1(p As Long)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("EnumeratePrinters1")
Call clsError.DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case clsError.DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call ProjektForm.EndeDll
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

Dim Success As Boolean, cbRequired As Long, cbBuffer As Long
Dim Buffer() As Long, nEntries As Long
Dim i As Long, PFlags As Long, PDesc As String, PName As String
Dim PComment As String, temp As Long

cbBuffer = 3072
ReDim Buffer((cbBuffer \ 4) - 1) As Long

Success = EnumPrinters(p, vbNullString, 1, Buffer(0), cbBuffer, cbRequired, nEntries)
If Success Then
    If cbRequired > cbBuffer Then
        cbBuffer = cbRequired
        ReDim Buffer(cbBuffer \ 4) As Long
        Success = EnumPrinters(p, vbNullString, 1, Buffer(0), cbBuffer, cbRequired, nEntries)
        If Not Success Then
           Call clsDialog.MessageBox("Fehler bei der Druckerauflistung.")
           Call clsError.DefErrPop: Exit Sub
        End If
    End If
    For i = 0 To nEntries - 1
        PFlags = Buffer(4 * i)
        PDesc = Space$(StrLen(Buffer(i * 4 + 1)))
        temp = PtrToStr(PDesc, Buffer(i * 4 + 1))
        PName = Space$(StrLen(Buffer(i * 4 + 2)))
        temp = PtrToStr(PName, Buffer(i * 4 + 2))
        PComment = Space$(StrLen(Buffer(i * 4 + 2)))
        temp = PtrToStr(PComment, Buffer(i * 4 + 2))
        
        iAnzPrinter = iAnzPrinter + 1
        ReDim Preserve iPrinterName(iAnzPrinter)
        iPrinterName(iAnzPrinter) = PName
    Next i
Else
    Call clsDialog.MessageBox("Fehler bei der Druckerauflistung.")
End If

Call clsError.DefErrPop
End Sub


Sub EnumeratePrinters4(p As Long)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("EnumeratePrinters4")
Call clsError.DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case clsError.DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call ProjektForm.EndeDll
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

Dim Success As Boolean, cbRequired As Long, cbBuffer As Long
Dim Buffer() As Long, nEntries As Long
Dim i As Long, PName As String, sName As String
Dim Attrib As Long, temp As Long
cbBuffer = 3072
ReDim Buffer((cbBuffer \ 4) - 1) As Long

Success = EnumPrinters(p, vbNullString, 4, Buffer(0), cbBuffer, cbRequired, nEntries)
If Success Then
    If cbRequired > cbBuffer Then
        cbBuffer = cbRequired
        ReDim Buffer(cbBuffer \ 4) As Long
        Success = EnumPrinters(p, vbNullString, 4, Buffer(0), cbBuffer, cbRequired, nEntries)
        If Not Success Then
            Call clsDialog.MessageBox("Fehler bei der Druckerauflistung.")
            clsError.DefErrPop: Exit Sub
        End If
    End If
    For i = 0 To nEntries - 1
        PName = Space$(StrLen(Buffer(i * 3)))
        temp = PtrToStr(PName, Buffer(i * 3))
        sName = Space$(StrLen(Buffer(i * 3 + 1)))
        temp = PtrToStr(sName, Buffer(i * 3 + 1))
        Attrib = Buffer(i * 3 + 2)
        
        iAnzPrinter = iAnzPrinter + 1
        ReDim Preserve iPrinterName(iAnzPrinter)
        iPrinterName(iAnzPrinter) = PName
    Next i
Else
    Call clsDialog.MessageBox("Fehler bei der Druckerauflistung.")
End If

Call clsError.DefErrPop
End Sub

Sub InstalledPrinters()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("InstalledPrinters")
Call clsError.DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case clsError.DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call ProjektForm.EndeDll
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Dim strComputer As String
'Dim objWMIService As Object
'Dim objPrinter As Object
'Dim colInstalledPrinters As Object
'
'strComputer = "."
'Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
'
'Set colInstalledPrinters = objWMIService.ExecQuery("Select * from Win32_Printer")
'
'iAnzPrinter = 0
'ReDim iPrinterName(iAnzPrinter)
'For Each objPrinter In colInstalledPrinters
'    iAnzPrinter = iAnzPrinter + 1
'    ReDim Preserve iPrinterName(iAnzPrinter)
'    iPrinterName(iAnzPrinter) = PrinterNameOP(objPrinter.Name)
'Next

iAnzPrinter = 0
ReDim iPrinterName(iAnzPrinter)
Call EnumeratePrinters1(PRINTER_ENUM_CONNECTIONS)
Call EnumeratePrinters4(PRINTER_ENUM_LOCAL)

Call clsError.DefErrPop
End Sub

Function PrinterCount%()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("PrinterCount%")
Call clsError.DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case clsError.DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call ProjektForm.EndeDll
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

PrinterCount = iAnzPrinter

Call clsError.DefErrPop
End Function


Function PrinterName$(Index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("PrinterName$")
Call clsError.DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case clsError.DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call ProjektForm.EndeDll
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim ret$

ret$ = ""
If (Index <= UBound(iPrinterName)) Then
    ret = iPrinterName(Index)
End If

PrinterName = ret

Call clsError.DefErrPop
End Function

Function PrinterNameOP$(sPrinterNameTS$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("PrinterNameOP$")
Call clsError.DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case clsError.DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call ProjektForm.EndeDll
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim ind%
Dim ret$

ret = sPrinterNameTS
ind = InStr(UCase(sPrinterNameTS), "(UMGELEITET")
If (ind > 0) Then
    ret = Trim(Left(sPrinterNameTS, ind - 1))
End If

PrinterNameOP = ret

Call clsError.DefErrPop
End Function

Function ServerRechnerName$()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("ServerRechnerName$")
Call clsError.DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case clsError.DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call ProjektForm.EndeDll
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim i%
Dim lErg&
Dim h2$, ret$

h2 = Left(CurDir, 2)

ret = String(MAX_PATH, 0)
lErg = WNetGetConnection(h2, ret, Len(ret))
If (lErg = NO_ERROR) Then
    Do
        If (Right(ret, 1) = Chr(0)) Then
            ret = Left(ret, Len(ret) - 1)
        Else
            Exit Do
        End If
    Loop
'    ret = Chr0RechtsWeg(ret)
    If (Left(ret, 2) = "\\") Then
        ret = Mid(ret, 3)
    End If
    i = InStr(ret, "\")
    If (i > 0) Then
        ret = UCase(Left(ret, i - 1))
    End If
Else
    ret = UCase(Trim(Environ("COMPUTERNAME")))
End If

ServerRechnerName = ret

Call clsError.DefErrPop
End Function

Function LwZ$()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("LwZ$")
Call clsError.DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case clsError.DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call ProjektForm.EndeDll
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim i&
Dim lErg&
Dim s$, ret$, IniFile$

ret = "Z:"
IniFile = Environ("OPINI")
If (IniFile = "") Then IniFile = Environ("ProgramData")
If (IniFile = "") Then
  If (Dir(IniFile + "\optipharm\opsystem.ini") <> "") Then IniFile = IniFile + "\optipharm"
End If
If (IniFile = "") Then
  If (Dir("c:\programme\optipharm\opsystem.ini") <> "") Then IniFile = "c:\programme\optipharm"
End If

If (IniFile <> "") Then
    IniFile = IniFile + "\opsystem.ini"
    
    ' aus historischen Gründen kann es sein, dass beim Laufwerk z.B. „ZK“ oder „Z:K:“ drinsteht, darum gilt K nur, wenn es die einzige Angabe ist
    s = Space(30)
    i = GetPrivateProfileString("Network", "OPDrives", "", s, 30, IniFile)
    If (i > 0) Then
        s = UCase(Trim(Left(s, i)))
        If (Left(s, 1) <> "K") Or (Len(s) = 1) Then
            ret = Left(s, 1) + ":"
        Else
            s = Mid(s, 2)
            If Left(s, 1) = ":" Then s = Mid(s, 2)
            If Len(s) = 0 Then
                ret = "K:"
            Else
                ret = Left(s, 1) + ":"
            End If
        End If
    End If
End If
  
LwZ = ret

Call clsError.DefErrPop
End Function

Function HavarieRechnerName$()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("HavarieRechnerName$")
Call clsError.DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case clsError.DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call ProjektForm.EndeDll
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim i&
Dim lErg&
Dim s$, ret$, IniFile$

ret = ""
IniFile = Environ("OPINI")
If (IniFile = "") Then IniFile = Environ("ProgramData")
If (IniFile = "") Then
    If (Dir(IniFile + "\optipharm\opsystem.ini") <> "") Then IniFile = IniFile + "\optipharm"
End If
If (IniFile = "") Then
    If (Dir("c:\programme\optipharm\opsystem.ini") <> "") Then IniFile = "c:\programme\optipharm"
End If

If (IniFile <> "") Then
    IniFile = IniFile + "\opsystem.ini"
    
    s = Space(30)
    i = GetPrivateProfileString("Network", "OPHvrServer", "", s, 30, IniFile)
    If (i > 0) Then
        ret = UCase$(Trim$(Left(s, i)))
        Do While Left(ret, 1) = "\"
            ret = Mid(ret, 2)
        Loop
        i = InStr(ret, "\")
        If (i > 0) Then
            ret = Left(ret, i - 1)
        End If
    End If
End If
  
HavarieRechnerName = ret

Call clsError.DefErrPop
End Function

Function HavarieBetrieb%()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("HavarieBetrieb%")
Call clsError.DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case clsError.DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call ProjektForm.EndeDll
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim ret%
Dim h$

h = wPara1.HavarieRechnerName
If (h <> "") Then
    ret = (wPara1.ServerRechnerName = h)
End If

HavarieBetrieb = ret

Call clsError.DefErrPop
End Function
 

