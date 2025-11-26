Attribute VB_Name = "modToolbar"
Option Explicit

Type ToolbarStruct
    Visible As Integer
    Labels As Integer
    Position As Integer
    BigSymbols As Integer
End Type

Public Const TOOLBAR_BUTTONSIZE = 330
Public Const TOOLBAR_BUTTONSIZE_BIGSYMBOLS = TOOLBAR_BUTTONSIZE + 240
Public Const TOOLBAR_HEIGHT = 390
Public Const TOOLBAR_HEIGHT_BIGSYMBOLS = TOOLBAR_HEIGHT + 240
Public Const TOOLBAR_WIDTH = 390
Public Const TOOLBAR_WIDTH_BIGSYMBOLS = TOOLBAR_WIDTH + 240
Public Const TOOLBAR_LABELS_HEIGHT = 240
Public Const TOOLBAR_LABELS_WIDTH = 330

Private Const DefErrModul = "toolbar.bas"

Sub InitToolbar(iForm As Form, iToolbar As ToolbarStruct, IniDatei$, IniSection$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("InitToolbar")
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call DefErrAbort
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim i%, X%, loch%, imgInd%

With iForm
    .picToolbar.AutoRedraw = True
    
    For i% = 1 To 19
        Load .cmdToolbar(i%)
        Load .lblTasten(i%)
    Next i%
    
    .cmdToolbar(0).ToolTipText = "ESC Zurück: Zurückschalten auf vorige Bildschirmmaske"
    .cmdToolbar(1).ToolTipText = "F2 Erfassen zusätzlicher Artikel"
    .cmdToolbar(2).ToolTipText = "F3 Auswahl eines Lieferanten"
    'cmdToolbar(2).ToolTipText = "F3 Fixierung des Artikels für einen Lieferanten"
    .cmdToolbar(3).ToolTipText = "F4 Editieren des Infobereichs"
    .cmdToolbar(4).ToolTipText = "F5 Entfernen eines Artikels oder einer Bedingung"
    .cmdToolbar(5).ToolTipText = "F6 Drucken: Drucken der Bestellung"
    .cmdToolbar(6).ToolTipText = "F7 BM: Korrektur der Bestellmenge"
    .cmdToolbar(7).ToolTipText = "F8 Bestellvorschlag"
    .cmdToolbar(8).ToolTipText = "F9 NR: Korrektur des Naturalrabatts"
    .cmdToolbar(9).ToolTipText = "shift+F2 Bestell-Status"
    .cmdToolbar(10).ToolTipText = "shift+F3 Lieferant zuordnen: dem aktuellen Artikel einen Lieferanten zuordnen"
    .cmdToolbar(11).ToolTipText = "shift+F4 "
    .cmdToolbar(12).ToolTipText = "shift+F5 Statistik: Durchgriff auf Statistik-Anzeige"
    .cmdToolbar(13).ToolTipText = "shift+F6 Senden: Datenübertragung zum GH"
    .cmdToolbar(19).ToolTipText = "Programm beenden"
    
    For i% = 0 To 19
        If (i% = 0) Then
            .lblTasten(i%).Caption = "Esc"
        ElseIf (i% <= 8) Then
            .lblTasten(i%).Caption = "F" + Format(i% + 1, "0")
        ElseIf (i% <= 16) Then
            .lblTasten(i%).Caption = "sF" + Format(i% - 7, "0")
        Else
            .lblTasten(i%).Caption = "a" + Format(i% - 13, "0")
        End If
    Next i%
End With

Call HoleIniToolbar(iForm, iToolbar, IniDatei$, IniSection$)

Call DefErrPop
End Sub

Sub HoleIniToolbar(iForm As Form, iToolbar As ToolbarStruct, IniDatei$, IniSection$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("HoleIniToolbar")
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call DefErrAbort
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim l&
Dim h$
    
h$ = "J"
l& = GetPrivateProfileString(IniSection$, "Toolbar", "J", h$, 2, IniDatei$)
h$ = Left$(h$, l&)
If (h$ = "J") Then
    Call OpToolbarVisible(iForm, iToolbar, True)
Else
    Call OpToolbarVisible(iForm, iToolbar, False)
End If

h$ = "J"
l& = GetPrivateProfileString(IniSection$, "ToolbarTasten", "J", h$, 2, IniDatei$)
h$ = Left$(h$, l&)
If (h$ = "J") Then
    Call OpToolbarLabels(iForm, iToolbar, True)
Else
    Call OpToolbarLabels(iForm, iToolbar, False)
End If

h$ = "N"
l& = GetPrivateProfileString(IniSection$, "ToolbarGross", "N", h$, 2, IniDatei$)
h$ = Left$(h$, l&)
If (h$ = "J") Then
    Call OpToolbarBigSymbols(iForm, iToolbar, True)
Else
    Call OpToolbarBigSymbols(iForm, iToolbar, False)
End If

h$ = "0"
l& = GetPrivateProfileString(IniSection$, "ToolbarPosition", "0", h$, 2, IniDatei$)
Call OpToolbarPosition(iForm, iToolbar, Val(Left$(h$, l&)))

Call DefErrPop
End Sub

Sub SpeicherIniToolbar(iToolbar As ToolbarStruct, IniDatei$, IniSection$)
Dim l&
Dim h$

h$ = "N"
If (iToolbar.Visible) Then
    h$ = "J"
End If
l& = WritePrivateProfileString(IniSection$, "Toolbar", h$, IniDatei$)

h$ = "N"
If (iToolbar.Labels) Then
    h$ = "J"
End If
l& = WritePrivateProfileString(IniSection$, "ToolbarTasten", h$, IniDatei$)

h$ = "N"
If (iToolbar.BigSymbols) Then
    h$ = "J"
End If
l& = WritePrivateProfileString(IniSection$, "ToolbarGross", h$, IniDatei$)

l& = WritePrivateProfileString(IniSection$, "ToolbarPosition", Str$(iToolbar.Position), IniDatei$)

End Sub

Sub ResizeToolbar(iForm As Form, iToolbar As ToolbarStruct)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("ResizeToolbar")
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call DefErrAbort
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim i%, X%, loch%, ToolbarPos%, cmdToolbarSize%
Dim h$

On Error Resume Next


ToolbarPos% = iToolbar.Position

With iForm

    .mnuToolbarGross.Checked = iToolbar.BigSymbols
    .mnuToolbarLabels.Checked = iToolbar.Labels
    For i% = 0 To 3
        .mnuToolbarPositionInd(i%).Checked = False
    Next i%
    If (ToolbarPos% < 4) Then
        .mnuToolbarPositionInd(ToolbarPos%).Checked = True
    End If
    
    If (iToolbar.Visible) Then
        .mnuToolbarVisible.Caption = "Ausblenden"
    Else
        .mnuToolbarVisible.Caption = "Einblenden"
    End If
    
    .mnuToolbarGross.Enabled = iToolbar.Visible
    .mnuToolbarLabels.Enabled = iToolbar.Visible
    .mnuToolbarPosition.Enabled = iToolbar.Visible
  
    
    
    If (iToolbar.BigSymbols) Then
        cmdToolbarSize% = TOOLBAR_BUTTONSIZE_BIGSYMBOLS
    Else
        cmdToolbarSize% = TOOLBAR_BUTTONSIZE
    End If
    
    If (ToolbarPos% = 1) Or (ToolbarPos% = 3) Then
        If (iToolbar.BigSymbols) Then
            .picToolbar.Width = TOOLBAR_WIDTH_BIGSYMBOLS
        Else
            .picToolbar.Width = TOOLBAR_WIDTH
        End If
        If (iToolbar.Labels) Then
            .picToolbar.Width = .picToolbar.Width + TOOLBAR_LABELS_WIDTH
        End If
        
        .picToolbar.Top = 0
        .picToolbar.Height = .ScaleHeight
        If (ToolbarPos% = 1) Then
            .picToolbar.Left = .ScaleWidth - .picToolbar.Width
        Else
            .picToolbar.Left = 0
        End If
    Else
        If (iToolbar.BigSymbols) Then
            .picToolbar.Height = TOOLBAR_HEIGHT_BIGSYMBOLS
        Else
            .picToolbar.Height = TOOLBAR_HEIGHT
        End If
        If (iToolbar.Labels) Then
            .picToolbar.Height = .picToolbar.Height + TOOLBAR_LABELS_HEIGHT
        End If
        
        .picToolbar.Width = .ScaleWidth
        If (ToolbarPos% = 0) Then
            .picToolbar.Top = 0
            .picToolbar.Left = 0
        ElseIf (ToolbarPos% = 2) Then
            .picToolbar.Top = .ScaleHeight - .picToolbar.Height
            .picToolbar.Left = 0
        End If
    End If
        
    
    If (ToolbarPos% > 3) Then
        loch% = 75
    Else
        If (ToolbarPos% = 0) Or (ToolbarPos% = 2) Then
            loch% = (.picToolbar.ScaleWidth - (20 * (cmdToolbarSize% + 15))) / 3
        Else
            loch% = (.picToolbar.ScaleHeight - (20 * (cmdToolbarSize% + 15))) / 3
        End If
    End If
    
'    .picToolbar.Font.Name = .lblTasten(0).Font.Name
'    .picToolbar.Font.Size = .lblTasten(0).Font.Size
    .picToolbar.Cls

    X% = 0
    For i% = 0 To 19
        If ((i% = 1) Or (i% = 9) Or (i% = 17)) Then X% = X% + loch% '105
    '    If (i% = 19) And (ToolbarPos% < 3) Then
    '        If (ToolbarPos% = 0) Or (ToolbarPos% = 2) Then
    '            X% = picToolbar.ScaleWidth - cmdToolbarSize%
    '        Else
    '            X% = picToolbar.ScaleHeight - cmdToolbarSize%
    '        End If
    '    End If
        
        If (i% = 0) Then
            h$ = "Esc"
        ElseIf (i% <= 8) Then
            h$ = "F" + Format(i% + 1, "0")
        ElseIf (i% <= 16) Then
            h$ = "sF" + Format(i% - 7, "0")
        Else
            h$ = "a" + Format(i% - 16, "0")
        End If
        
        If (ToolbarPos% = 1) Or (ToolbarPos% = 3) Then
            .cmdToolbar(i%).Top = X%
            .cmdToolbar(i%).Left = 0
            .cmdToolbar(i%).Width = cmdToolbarSize%
            .cmdToolbar(i%).Height = cmdToolbarSize%
            .cmdToolbar(i%).Visible = True
            
            .lblTasten(i%).Top = X% + 30
            .lblTasten(i%).Left = cmdToolbarSize% + 30
            .lblTasten(i%).Width = 240
            .lblTasten(i%).Height = cmdToolbarSize%
            .lblTasten(i%).Visible = False  'True
            
            .picToolbar.CurrentX = cmdToolbarSize% + 30
            .picToolbar.CurrentY = X% + 30
        Else
            .cmdToolbar(i%).Top = 0
            .cmdToolbar(i%).Left = X%
            .cmdToolbar(i%).Width = cmdToolbarSize%
            .cmdToolbar(i%).Height = cmdToolbarSize%
            .cmdToolbar(i%).Visible = True
            
            .lblTasten(i%).Top = cmdToolbarSize% + 45
            .lblTasten(i%).Left = X%
            .lblTasten(i%).Width = cmdToolbarSize%
            .lblTasten(i%).Height = 240
            .lblTasten(i%).Visible = False  'True
        
            .picToolbar.CurrentX = X% + (cmdToolbarSize% - .picToolbar.TextWidth(h$)) / 2
            .picToolbar.CurrentY = cmdToolbarSize% + 45
        End If
        
'        .picToolbar.CurrentX = X%
'        .picToolbar.CurrentY = cmdToolbarSize% + 45
        .picToolbar.Print h$
        
        
        X% = X% + cmdToolbarSize% + 15
    Next i%
    
    If (ToolbarPos% = 1) Or (ToolbarPos% = 3) Then
        .picToolbar.Height = X% + 30
    Else
        .picToolbar.Width = X% + 30
    End If
    
    If (iToolbar.Visible) Then
        .picToolbar.Visible = True
    Else
        .picToolbar.Visible = False
    End If
End With

Call DefErrPop
End Sub

Sub OpToolbarBigSymbols(iForm As Form, iToolbar As ToolbarStruct, NeuerWert%)
Dim i%, ToolbarImageInd%

iToolbar.BigSymbols = NeuerWert%

If (NeuerWert%) Then
    ToolbarImageInd% = 1
Else
    ToolbarImageInd% = 0
End If

With iForm
    For i% = 0 To 19
        .cmdToolbar(i%).Picture = frmAction.imgToolbar(ToolbarImageInd%).ListImages(i% + 1).ExtractIcon
    Next i%
    
'    .cmdToolbar(11).Picture = LoadPicture("")
'    For i% = 14 To 18
'        .cmdToolbar(i%).Picture = LoadPicture("")
'    Next i%
    
    Call .OpToolbarResized
End With

End Sub

Sub OpToolbarLabels(iForm As Form, iToolbar As ToolbarStruct, NeuerWert%)

iToolbar.Labels = NeuerWert%

Call iForm.OpToolbarResized

End Sub

Sub OpToolbarVisible(iForm As Form, iToolbar As ToolbarStruct, NeuerWert%)

iToolbar.Visible = NeuerWert%

Call iForm.OpToolbarResized

End Sub

Sub OpToolbarPosition(iForm As Form, iToolbar As ToolbarStruct, NeuerWert%)

iToolbar.Position = NeuerWert%

Call iForm.OpToolbarResized

End Sub



