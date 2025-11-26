VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMisc"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim iFontName$(1)
Dim iFontSize%(1)

Dim iFarbeArbeit&, iFarbeInfo&
Dim iFarbeRahmen%

Dim iIniSection$
Dim iIniDatei$

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

Function EditFont%(CommDlg As Object, Index As Integer)
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

CommDlg.FontName = RTrim$(iFontName$(Index))
CommDlg.FontSize = iFontSize%(Index)
CommDlg.FontBold = False
CommDlg.FontItalic = False
CommDlg.FontStrikethru = False
CommDlg.FontUnderline = False
CommDlg.Flags = cdlCFBoth
CommDlg.CancelError = False
Call CommDlg.ShowFont
If (Err = 0) Then
    If ((iFontName$(Index) <> CommDlg.FontName) Or (iFontSize%(Index) <> CommDlg.FontSize)) Then
        iFontName$(Index) = CommDlg.FontName
        iFontSize%(Index) = CommDlg.FontSize
        h$ = RTrim$(iFontName$(Index)) + "," + Mid$(Str$(iFontSize%(Index)), 2)
        If (Index = 0) Then
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

Function EditFarbe%(CommDlg As Object, Index As Integer)
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

If (Index = 0) Then
    CommDlg.Color = iFarbeArbeit&
Else
    CommDlg.Color = iFarbeInfo&
End If
CommDlg.CancelError = True
CommDlg.Flags = cdlCCFullOpen + cdlCCRGBInit
Call CommDlg.ShowColor
If (Err = 0) Then
    If (Index = 0) Then
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

Public Property Get FarbeArbeit&()

FarbeArbeit& = iFarbeArbeit&

End Property

Public Property Get FarbeInfo&()

FarbeInfo& = iFarbeInfo&

End Property

Public Property Get FarbeRahmen%()

FarbeRahmen% = iFarbeRahmen%

End Property



