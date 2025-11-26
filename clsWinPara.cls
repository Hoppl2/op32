VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsWinPara"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim iWorkAreaHeight&, iWorkAreaWidth&
Dim iFrmCaptionHeight&, iFrmMenuHeight&, iFrmBorderHeight&, iFrmScrollHeight&

Dim iBildFaktor!

Dim iLinksX%, iTitelY%, iFlexY%, iUntenFreiY%, iOhneInfoY%
Dim iButtonX%, iButtonY%

Dim iFontName$(1)
Dim iFontSize%(1)

Dim iFarbeArbeit&, iFarbeInfo&
Dim iFarbeRahmen%

Dim iIniSection$
Dim iIniDatei$

Private Const DefErrModul = "clsWinPara.bas"

Private Sub Class_Initialize()
Set wPara1 = Me
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

Public Property Get FontName$(ind%)
FontName$ = iFontName$(ind%)
End Property
Public Property Get FontSize%(ind%)
FontSize% = iFontSize%(ind%)
End Property


Sub HoleWindowsParameter()
Dim l&
Dim WorkRect As RECT

l& = SystemParametersInfo(SPI_GETWORKAREA, vbNull, WorkRect, 0)
iWorkAreaHeight& = (WorkRect.Bottom - WorkRect.Top - 10) * Screen.TwipsPerPixelY
iWorkAreaWidth& = (WorkRect.Right - WorkRect.Left) * Screen.TwipsPerPixelY

iFrmScrollHeight& = Abs(QueryRegistryValue&(".DEFAULT\Control Panel\Desktop\WindowMetrics", "ScrollHeight"))

iFrmCaptionHeight& = GetSystemMetrics(SM_CYCAPTION) * 15
iFrmMenuHeight& = GetSystemMetrics(SM_CYMENU) * 15
iFrmBorderHeight& = GetSystemMetrics(SM_CYFRAME) * 15

' Verhältnis der Bildschirmauflösungen
iBildFaktor! = Screen.Width / 12000

iLinksX% = 150 * iBildFaktor!
iTitelY% = 195 * iBildFaktor!
iFlexY% = 690 * iBildFaktor!
iUntenFreiY% = 60 * iBildFaktor
iOhneInfoY% = 600 * iBildFaktor!

iButtonX% = 1200 * iBildFaktor!
iButtonY% = 390 * iBildFaktor!

End Sub

Function QueryRegistryValue&(sKeyName$, sValueName$)
Dim lRetVal&         'result of the API functions
Dim hKey&         'handle of opened key
Dim vValue As Variant      'setting of queried value

lRetVal& = RegOpenKeyEx(HKEY_USERS, sKeyName$, 0, KEY_ALL_ACCESS, hKey&)
lRetVal& = QueryValueEx(hKey&, sValueName$, vValue)
RegCloseKey (hKey&)

QueryRegistryValue& = vValue
End Function

Function QueryValueEx&(ByVal lhKey&, ByVal szValueName$, vValue As Variant)
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

End Function

Sub HoleGlobalIniWerte(IniSection$, IniDatei$)
Dim ind%
Dim l&
Dim h$, h2$

iIniSection$ = IniSection$
iIniDatei$ = IniDatei$

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
    
'Call InitAlleBereichsFarben

End Sub

Function BerechneFarbWert&(Farb$)
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Call DefErrFnc("BerechneFarbWert&")
'On Error GoTo DefErr
'GoTo DefErrEnd
'DefErr:
'Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
'Case vbRetry
'  Resume
'Case vbIgnore
'  Resume Next
'End Select
'Call DefErrAbort
'DefErrEnd:
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim i%, l%
Dim tetrade&, wert&
Dim ch$

wert& = 0
l% = Len(Farb$)
For i% = 1 To l%
    wert& = wert& * 16
    ch$ = Mid$(Farb$, i%, 1)
    tetrade& = Val("&H" + ch$)
    wert& = wert& + tetrade&
Next i%
    
BerechneFarbWert& = wert&

'Call DefErrPop
End Function

Function EditFont%(CommDlg As Object, index As Integer)
'Call DefErrFnc("EditFont")
'On Error GoTo DefErr
'GoTo DefErrEnd
'DefErr:
'Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
'Case vbRetry
'  Resume
'Case vbIgnore
'  Resume Next
'End Select
'Call DefErrAbort
'DefErrEnd:
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim ret%
Dim l&
Dim h$

'On Error Resume Next

ret% = False

CommDlg.FontName = RTrim$(iFontName$(index))
CommDlg.FontSize = iFontSize%(index)
CommDlg.FontBold = False
CommDlg.FontItalic = False
CommDlg.FontStrikethru = False
CommDlg.FontUnderline = False
CommDlg.Flags = cdlCFBoth
CommDlg.CancelError = False
Call CommDlg.ShowFont
If (Err = 0) Then
    If ((iFontName$(index) <> CommDlg.FontName) Or (iFontSize%(index) <> CommDlg.FontSize)) Then
        iFontName$(index) = CommDlg.FontName
        iFontSize%(index) = CommDlg.FontSize
        h$ = RTrim$(iFontName$(index)) + "," + Mid$(Str$(iFontSize%(index)), 2)
        If (index = 0) Then
            l& = WritePrivateProfileString(iIniSection$, "FontInformation", h$, iIniDatei$)
        Else
            l& = WritePrivateProfileString(iIniSection$, "FontBeschriftung", h$, iIniDatei$)
        End If
        ret% = True
    End If
End If

EditFont% = ret%

'Call DefErrPop

End Function

Function EditFarbe%(CommDlg As Object, index As Integer)
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Call DefErrFnc("EditBereichsFarbe")
'On Error GoTo DefErr
'GoTo DefErrEnd
'DefErr:
'Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
'Case vbRetry
'  Resume
'Case vbIgnore
'  Resume Next
'End Select
'Call DefErrAbort
'DefErrEnd:
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim ret%
Dim l&

On Error Resume Next

ret% = False

If (index = 0) Then
    CommDlg.Color = iFarbeArbeit&
Else
    CommDlg.Color = iFarbeInfo&
End If
CommDlg.CancelError = True
CommDlg.Flags = cdlCCFullOpen + cdlCCRGBInit
Call CommDlg.ShowColor
If (Err = 0) Then
    If (index = 0) Then
        iFarbeArbeit& = CommDlg.Color
        l& = WritePrivateProfileString(iIniSection$, "FarbeArbeit", Hex$(iFarbeArbeit&), iIniDatei$)
    Else
        iFarbeInfo& = CommDlg.Color
        l& = WritePrivateProfileString(iIniSection$, "FarbeInfo", Hex$(iFarbeInfo&), iIniDatei$)
    End If
'    Call InitAlleBereichsFarben
    ret% = True
End If

EditFarbe% = ret%

'Call DefErrPop
End Function

Sub InitFont(frm As Object)
Dim c As Control

On Error Resume Next

For Each c In frm.Controls
    If (c.Name = "picToolbar") Then
        c.Font.Name = "Small Fonts"
        c.Font.Size = 7
    ElseIf (c.Container.Name = "picStartBack") Then
    Else
        If (TypeOf c Is MSFlexGrid) Or (TypeOf c Is ListBox) Or (TypeOf c Is TextBox) Then
            c.Font.Name = RTrim$(iFontName$(0))
            c.Font.Size = iFontSize%(0)
        Else
            c.Font.Name = RTrim$(iFontName$(1))
            c.Font.Size = iFontSize%(1)
        End If
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

End Sub


