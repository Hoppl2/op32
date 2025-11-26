VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsToolbar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Event Resized()

Dim ParentForm As Form

Dim iVisible%
Dim iLabels%
Dim iPosition%
Dim iBigSymbols%

Dim iDragX!
Dim iDragY!

Dim iIniDatei$
Dim iIniSection$

Const TOOLBAR_BUTTONSIZE = 330
Const TOOLBAR_BUTTONSIZE_BIGSYMBOLS = TOOLBAR_BUTTONSIZE + 240
Const TOOLBAR_HEIGHT = 390
Const TOOLBAR_HEIGHT_BIGSYMBOLS = TOOLBAR_HEIGHT + 240
Const TOOLBAR_WIDTH = 390
Const TOOLBAR_WIDTH_BIGSYMBOLS = TOOLBAR_WIDTH + 240
Const TOOLBAR_LABELS_HEIGHT = 240
Const TOOLBAR_LABELS_WIDTH = 330

Private Const DefErrModul = "toolbar.cls"

Sub InitToolbar(iForm As Object, IniDatei$, IniSection$)
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Call DefErrFnc("InitToolbar")
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
Dim i%, X%, loch%, imgInd%

Set ParentForm = iForm
iIniDatei$ = IniDatei$
iIniSection$ = IniSection$

Call HoleIniToolbar

With ParentForm
    .picToolbar.AutoRedraw = True
    
    For i% = 1 To 19
        Load .cmdToolbar(i%)
    Next i%
    
'    .cmdToolbar(0).ToolTipText = "ESC Zurück: Zurückschalten auf vorige Bildschirmmaske"
'    .cmdToolbar(1).ToolTipText = "F2 Erfassen zusätzlicher Artikel"
'    .cmdToolbar(2).ToolTipText = "F3 Auswahl eines Lieferanten"
'    'cmdToolbar(2).ToolTipText = "F3 Fixierung des Artikels für einen Lieferanten"
'    .cmdToolbar(3).ToolTipText = "F4 Editieren des Infobereichs"
'    .cmdToolbar(4).ToolTipText = "F5 Entfernen eines Artikels oder einer Bedingung"
'    .cmdToolbar(5).ToolTipText = "F6 Drucken: Drucken der Bestellung"
'    .cmdToolbar(6).ToolTipText = "F7 BM: Korrektur der Bestellmenge"
'    .cmdToolbar(7).ToolTipText = "F8 Bestellvorschlag"
'    .cmdToolbar(8).ToolTipText = "F9 NR: Korrektur des Naturalrabatts"
'    .cmdToolbar(9).ToolTipText = "shift+F2 Bestell-Status"
'    .cmdToolbar(10).ToolTipText = "shift+F3 Lieferant zuordnen: dem aktuellen Artikel einen Lieferanten zuordnen"
'    .cmdToolbar(11).ToolTipText = "shift+F4 "
'    .cmdToolbar(12).ToolTipText = "shift+F5 Statistik: Durchgriff auf Statistik-Anzeige"
'    .cmdToolbar(13).ToolTipText = "shift+F6 Senden: Datenübertragung zum GH"
'    .cmdToolbar(19).ToolTipText = "Programm beenden"
End With

Call LoadSymbols

'Call DefErrPop
End Sub

Sub HoleIniToolbar()
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Call DefErrFnc("HoleIniToolbar")
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
Dim l&
Dim h$
    
h$ = "J"
l& = GetPrivateProfileString(iIniSection$, "Toolbar", "J", h$, 2, iIniDatei$)
h$ = Left$(h$, l&)
If (h$ = "J") Then
    iVisible% = True
Else
    iVisible% = False
End If

h$ = "J"
l& = GetPrivateProfileString(iIniSection$, "ToolbarTasten", "J", h$, 2, iIniDatei$)
h$ = Left$(h$, l&)
If (h$ = "J") Then
    iLabels% = True
Else
    iLabels% = False
End If

h$ = "N"
l& = GetPrivateProfileString(iIniSection$, "ToolbarGross", "N", h$, 2, iIniDatei$)
h$ = Left$(h$, l&)
If (h$ = "J") Then
    iBigSymbols% = True
Else
    iBigSymbols% = False
End If

h$ = "0"
l& = GetPrivateProfileString(iIniSection$, "ToolbarPosition", "0", h$, 2, iIniDatei$)
iPosition% = Val(Left$(h$, l&))

'Call OpToolbarPosition(iForm, iToolbar, Val(Left$(h$, l&)))

'Call DefErrPop
End Sub

Sub SpeicherIniToolbar()
Dim l&
Dim h$

h$ = "N"
If (iVisible%) Then
    h$ = "J"
End If
l& = WritePrivateProfileString(iIniSection$, "Toolbar", h$, iIniDatei$)

h$ = "N"
If (iLabels%) Then
    h$ = "J"
End If
l& = WritePrivateProfileString(iIniSection$, "ToolbarTasten", h$, iIniDatei$)

h$ = "N"
If (iBigSymbols%) Then
    h$ = "J"
End If
l& = WritePrivateProfileString(iIniSection$, "ToolbarGross", h$, iIniDatei$)

l& = WritePrivateProfileString(iIniSection$, "ToolbarPosition", Str$(iPosition%), iIniDatei$)

Set ParentForm = Nothing

End Sub

Sub ResizeToolbar()
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Call DefErrFnc("ResizeToolbar")
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
Dim i%, X%, loch%, ToolbarPos%, cmdToolbarSize%
Dim h$

On Error Resume Next


With ParentForm

    .mnuToolbarGross.Checked = iBigSymbols%
    .mnuToolbarLabels.Checked = iLabels%
    For i% = 0 To 3
        .mnuToolbarPositionInd(i%).Checked = False
    Next i%
    If (iPosition% < 4) Then
        .mnuToolbarPositionInd(iPosition%).Checked = True
    End If
    
    If (iVisible%) Then
        .mnuToolbarVisible.Caption = "Ausblenden"
    Else
        .mnuToolbarVisible.Caption = "Einblenden"
    End If
    
    .mnuToolbarGross.Enabled = iVisible%
    .mnuToolbarLabels.Enabled = iVisible%
    .mnuToolbarPosition.Enabled = iVisible%
  
    
    
    If (iBigSymbols%) Then
        cmdToolbarSize% = TOOLBAR_BUTTONSIZE_BIGSYMBOLS
    Else
        cmdToolbarSize% = TOOLBAR_BUTTONSIZE
    End If
    
    If (iPosition% = 1) Or (iPosition% = 3) Then
        If (iBigSymbols%) Then
            .picToolbar.Width = TOOLBAR_WIDTH_BIGSYMBOLS
        Else
            .picToolbar.Width = TOOLBAR_WIDTH
        End If
        If (iLabels%) Then
            .picToolbar.Width = .picToolbar.Width + TOOLBAR_LABELS_WIDTH
        End If
        
        .picToolbar.Top = 0
        .picToolbar.Height = .ScaleHeight
        If (iPosition% = 1) Then
            .picToolbar.Left = .ScaleWidth - .picToolbar.Width
        Else
            .picToolbar.Left = 0
        End If
    Else
        If (iBigSymbols%) Then
            .picToolbar.Height = TOOLBAR_HEIGHT_BIGSYMBOLS
        Else
            .picToolbar.Height = TOOLBAR_HEIGHT
        End If
        If (iLabels%) Then
            .picToolbar.Height = .picToolbar.Height + TOOLBAR_LABELS_HEIGHT
        End If
        
        .picToolbar.Width = .ScaleWidth
        If (iPosition% = 0) Then
            .picToolbar.Top = 0
            .picToolbar.Left = 0
        ElseIf (iPosition% = 2) Then
            .picToolbar.Top = .ScaleHeight - .picToolbar.Height
            .picToolbar.Left = 0
        End If
    End If
        
    
    If (iPosition% > 3) Then
        loch% = 75
    Else
        If (iPosition% = 0) Or (iPosition% = 2) Then
            loch% = (.picToolbar.ScaleWidth - (20 * (cmdToolbarSize% + 15))) / 3
        Else
            loch% = (.picToolbar.ScaleHeight - (20 * (cmdToolbarSize% + 15))) / 3
        End If
    End If
    
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
        
        If (iPosition% = 1) Or (iPosition% = 3) Then
            .cmdToolbar(i%).Top = X%
            .cmdToolbar(i%).Left = 0
            .cmdToolbar(i%).Width = cmdToolbarSize%
            .cmdToolbar(i%).Height = cmdToolbarSize%
            .cmdToolbar(i%).Visible = True
            
            .picToolbar.CurrentX = cmdToolbarSize% + 30
            .picToolbar.CurrentY = X% + 30
        Else
            .cmdToolbar(i%).Top = 0
            .cmdToolbar(i%).Left = X%
            .cmdToolbar(i%).Width = cmdToolbarSize%
            .cmdToolbar(i%).Height = cmdToolbarSize%
            .cmdToolbar(i%).Visible = True
            
            .picToolbar.CurrentX = X% + (cmdToolbarSize% - .picToolbar.TextWidth(h$)) / 2
            .picToolbar.CurrentY = cmdToolbarSize% + 45
        End If
        
        .picToolbar.Print h$
        
        
        X% = X% + cmdToolbarSize% + 15
    Next i%
    
    If (iPosition% = 1) Or (iPosition% = 3) Then
        .picToolbar.Height = X% + 30
    Else
        .picToolbar.Width = X% + 30
    End If
    
    If (iVisible%) Then
        .picToolbar.Visible = True
    Else
        .picToolbar.Visible = False
    End If
End With

'Call DefErrPop
End Sub

Public Property Get BigSymbols%()
BigSymbols% = iBigSymbols%
End Property
Public Property Let BigSymbols(ByVal NeuerWert%)
iBigSymbols% = NeuerWert%
Call LoadSymbols
RaiseEvent Resized
End Property

Public Property Get Labels%()
Labels% = iLabels%
End Property
Public Property Let Labels(ByVal NeuerWert%)
iLabels% = NeuerWert%
RaiseEvent Resized
End Property

Public Property Get Visible%()
Visible% = iVisible%
End Property
Public Property Let Visible(ByVal NeuerWert%)
iVisible% = NeuerWert%
RaiseEvent Resized
End Property

Public Property Get Position%()
Position% = iPosition%
End Property
Public Property Let Position(ByVal NeuerWert%)
iPosition% = NeuerWert%
RaiseEvent Resized
End Property

Public Property Get DragX!()
DragX! = iDragX!
End Property
Public Property Let DragX(ByVal NeuerWert!)
iDragX! = NeuerWert!
End Property

Public Property Get DragY!()
DragY! = iDragY!
End Property
Public Property Let DragY(ByVal NeuerWert!)
iDragY! = NeuerWert!
End Property

Private Sub LoadSymbols()
Dim i%, ToolbarImageInd%

If (iBigSymbols%) Then
    ToolbarImageInd% = 1
Else
    ToolbarImageInd% = 0
End If

With ParentForm
    For i% = 0 To 19
        .cmdToolbar(i%).Picture = .imgToolbar(ToolbarImageInd%).ListImages(i% + 1).ExtractIcon
    Next i%
    
'    .cmdToolbar(11).Picture = LoadPicture("")
'    For i% = 14 To 18
'        .cmdToolbar(i%).Picture = LoadPicture("")
'    Next i%
    
'    Call .OpToolbarResized
End With

End Sub

Sub Move(dest As Object, pBack As Object, Source As Object, X!, Y!)
Dim i%
Dim Suc&

With Source
    .Top = pBack.Top + dest.Top + Y! - iDragY!
    .Left = pBack.Left + dest.Left + X! - iDragX!
    If (.Top <= 0) Then
        Position = 0
    ElseIf (.Top + .Height >= ParentForm.ScaleHeight) Then
        Position = 2
    ElseIf (.Top > 0) Then
        Position = 4
'        Call OpToolbarPosition(4)
    End If
End With

End Sub

Sub UpdateFlag(status%)
Dim ToolbarImageInd%

If (iBigSymbols%) Then
    ToolbarImageInd% = 1
Else
    ToolbarImageInd% = 0
End If

With ParentForm
    If (status%) Then
        .cmdToolbar(6).Picture = .imgToolbar(ToolbarImageInd%).ListImages(22).ExtractIcon
    Else
        .cmdToolbar(6).Picture = .imgToolbar(ToolbarImageInd%).ListImages(7).ExtractIcon
    End If
End With

End Sub


