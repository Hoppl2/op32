VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
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

Dim iSize%
Dim IconPixel%, ButtonWidth%, SpaceWidth%
Dim FaktorY!

Dim iDragX!
Dim iDragY!

Dim iIniDatei$
Dim iIniSection$
Dim iErsatzBilder$
Dim ImgInd%(19)
Dim iAktIndex%

Const TOOLBAR_BUTTONSIZE = 330
Const TOOLBAR_BUTTONSIZE_BIGSYMBOLS = TOOLBAR_BUTTONSIZE + 240
Const TOOLBAR_HEIGHT = 390
Const TOOLBAR_HEIGHT_BIGSYMBOLS = TOOLBAR_HEIGHT + 240
Const TOOLBAR_WIDTH = 390
Const TOOLBAR_WIDTH_BIGSYMBOLS = TOOLBAR_WIDTH + 240
Const TOOLBAR_LABELS_HEIGHT = 240
Const TOOLBAR_LABELS_WIDTH = 330


Private Const DefErrModul = "TOOLBAR.CLS"

Sub InitToolbar(iForm As Object, IniDatei$, IniSection$, Optional ErsatzBilder$ = "")
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("InitToolbar")
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

Set ParentForm = iForm
iIniDatei$ = IniDatei$
iIniSection$ = IniSection$

iErsatzBilder$ = ErsatzBilder$
ind = InStr(iErsatzBilder, ",")
If (ind = 0) Then
    iErsatzBilder = iErsatzBilder + ","
End If
ind = InStr(iErsatzBilder, ",")
If (iNewLine) Then
    iErsatzBilder = Mid$(iErsatzBilder, ind + 1)
    If (iErsatzBilder <> "") Then
        iErsatzBilder = "," + iErsatzBilder
    End If
Else
    iErsatzBilder = Left$(iErsatzBilder, ind - 1)
End If

Call HoleIniToolbar

With ParentForm
    .picToolbar.AutoRedraw = True
    
    On Error Resume Next
    For i% = 1 To 19
        Load .cmdToolbar(i%)
    Next i%
    On Error GoTo DefErr
    
End With

If (iNewLine) Then
    Call LoadIcons
    Call LoadToolbar
Else
    Call LoadSymbols
End If

Call clsError.DefErrPop
End Sub

Sub HoleIniToolbar()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("HoleIniToolbar")
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
Dim h$
    
If (iNewLine) Then
    If (PixelY >= 1200) Then
        h = "2"
    ElseIf (PixelY >= 1024) Then
        h = "1"
    Else
        h = "0"
    End If
    l& = GetPrivateProfileString("Allgemein", "Toolbar", h$, h$, 2, CurDir + "\newline.ini")
    iSize% = Val(Left$(h$, l&))
    If (iSize < 0) Or (iSize > 2) Then
        iSize = 0
    End If
Else
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
    If (iPosition < 0) Or (iPosition > 3) Then
        iPosition = 0
    End If
    
    'Call OpToolbarPosition(iForm, iToolbar, Val(Left$(h$, l&)))
End If

Call clsError.DefErrPop
End Sub

Sub SpeicherIniToolbar()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("SpeicherIniToolbar")
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
Dim h$

If (iNewLine) Then
    l& = WritePrivateProfileString("Allgemein", "Toolbar", Str$(iSize%), CurDir + "\newline.ini")
Else
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
End If

Set ParentForm = Nothing

Call clsError.DefErrPop
End Sub

Sub ResizeToolbar()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("ResizeToolbar")
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
Dim i%, x%, loch%, ToolbarPos%, cmdToolbarSize%
Dim h$

On Error Resume Next


With ParentForm

    If (iNewLine) Then
        Call ShowToolbar
    Else
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
    '            loch% = (.picToolbar.ScaleWidth - (20 * (cmdToolbarSize% + 15))) / 3
                loch% = (.picToolbar.ScaleWidth - (19 * (cmdToolbarSize% + 15))) / 2
            Else
    '            loch% = (.picToolbar.ScaleHeight - (20 * (cmdToolbarSize% + 15))) / 3
                loch% = (.picToolbar.ScaleHeight - (19 * (cmdToolbarSize% + 15))) / 2
            End If
        End If
        
        .picToolbar.Cls
    
        x% = 0
    '    For i% = 0 To 19
        For i% = 1 To 19
    '        If ((i% = 1) Or (i% = 9) Or (i% = 17)) Then X% = X% + loch% '105
            If ((i% = 9) Or (i% = 17)) Then x% = x% + loch% '105
            
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
                .cmdToolbar(i%).Top = x%
                .cmdToolbar(i%).Left = 0
                .cmdToolbar(i%).Width = cmdToolbarSize%
                .cmdToolbar(i%).Height = cmdToolbarSize%
                .cmdToolbar(i%).Visible = True
                
                .picToolbar.CurrentX = cmdToolbarSize% + 30
                .picToolbar.CurrentY = x% + 30
            Else
                .cmdToolbar(i%).Top = 15 ' 0
                .cmdToolbar(i%).Left = x%
                .cmdToolbar(i%).Width = cmdToolbarSize%
                .cmdToolbar(i%).Height = cmdToolbarSize%
                .cmdToolbar(i%).Visible = True
                
                .picToolbar.CurrentX = x% + (cmdToolbarSize% - .picToolbar.TextWidth(h$)) / 2
                .picToolbar.CurrentY = cmdToolbarSize% + 45
            End If
            
            .picToolbar.Print h$
            
            
            x% = x% + cmdToolbarSize% + 15
        Next i%
        
        If (iPosition% = 1) Or (iPosition% = 3) Then
            .picToolbar.Height = x% + 30
        Else
            .picToolbar.Width = x% + 30
        End If
        
        With .picToolbar
            .DrawWidth = 2
            ParentForm.picToolbar.Line (0, 0)-(.Width, 0), vbBlack, BF
            ParentForm.picToolbar.Line (0, .Height)-(.Width, .Height), vbWhite, BF
        End With
        
        
        If (iVisible%) Then
            .picToolbar.Visible = True
        Else
            .picToolbar.Visible = False
        End If
    
        .cmdToolbar(0).Visible = False
    End If

End With

Call clsError.DefErrPop
End Sub

Public Property Get Size%()
Size% = iSize%
End Property
Public Property Let Size(ByVal NeuerWert%)
If (iSize >= 0) And (iSize <= 2) Then
    iSize% = NeuerWert%
    Call LoadToolbar
    RaiseEvent Resized
End If
End Property

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

Public Property Get ImageInd%(ByVal ind%)
ImageInd = ImgInd(ind)
End Property
Public Property Let ImageInd(ByVal ind%, ByVal vNewValue%)
ImgInd(ind%) = vNewValue
End Property


Private Sub LoadSymbols()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("LoadSymbols")
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
Dim i%, ToolbarImageInd%, ind1%, ind2%
Dim h$

If (iBigSymbols%) Then
    ToolbarImageInd% = 1
Else
    ToolbarImageInd% = 0
End If

With ParentForm
    For i% = 0 To 19
        .cmdToolbar(i%).Picture = .imgToolbar(ToolbarImageInd%).ListImages(i% + 1).ExtractIcon
    Next i%
    
    .cmdToolbar(16).Picture = LoadPicture("")
    
    h$ = iErsatzBilder$
    While (h$ <> "")
        ind1% = Val(Left$(h$, 2))
        ind2% = Val(Mid$(h$, 3, 2)) - 1
        .cmdToolbar(ind2%).Picture = .imgToolbar(ToolbarImageInd%).ListImages(ind1%).ExtractIcon
        h$ = Mid$(h$, 5)
    Wend
    
'    .cmdToolbar(11).Picture = LoadPicture("")
'    For i% = 14 To 18
'        .cmdToolbar(i%).Picture = LoadPicture("")
'    Next i%
    
'    Call .OpToolbarResized
End With

Call clsError.DefErrPop
End Sub

Sub Move(dest As Object, pBack As Object, Source As Object, x!, y!)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("Move")
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
Dim Suc&

With Source
    .Top = pBack.Top + dest.Top + y! - iDragY!
    .Left = pBack.Left + dest.Left + x! - iDragX!
    If (.Top <= 0) Then
        Position = 0
    ElseIf (.Top + .Height >= ParentForm.ScaleHeight) Then
        Position = 2
    ElseIf (.Top > 0) Then
        Position = 4
'        Call OpToolbarPosition(4)
    End If
End With

Call clsError.DefErrPop
End Sub

Sub UpdateFlag(Status%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("UpdateFlag")
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
Dim ToolbarImageInd%

If (iNewLine) Then

    If (Status%) Then
        ImgInd(6) = 17
    Else
        ImgInd(6) = 6
    End If
    
    Call LoadToolbarButton(6)
    Call ShowToolbar((iAktIndex))
Else

    If (iBigSymbols%) Then
        ToolbarImageInd% = 1
    Else
        ToolbarImageInd% = 0
    End If
    
    With ParentForm
        If (Status%) Then
            .cmdToolbar(6).Picture = .imgToolbar(ToolbarImageInd%).ListImages(22).ExtractIcon
        Else
            .cmdToolbar(6).Picture = .imgToolbar(ToolbarImageInd%).ListImages(7).ExtractIcon
        End If
    End With
    
End If

Call clsError.DefErrPop
End Sub


Sub RundungFlag(Status%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("RundungFlag")
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
Dim ToolbarImageInd%

If (iBigSymbols%) Then
    ToolbarImageInd% = 1
Else
    ToolbarImageInd% = 0
End If

With ParentForm
    If (Status%) Then
        .cmdToolbar(MENU_SF7).Picture = .imgToolbar(ToolbarImageInd%).ListImages(24).ExtractIcon
    Else
        .cmdToolbar(MENU_SF7).Picture = .imgToolbar(ToolbarImageInd%).ListImages(23).ExtractIcon
    End If
End With

Call clsError.DefErrPop
End Sub

Sub LoadToolbar()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("LoadToolbar")
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
Dim i%, j%, ind%
Dim sLw$, h$

'On Error Resume Next

If (iSize = 0) Then
    IconPixel = 32
ElseIf (iSize = 1) Then
    IconPixel = 48
Else
    IconPixel = 64
End If
FaktorY = IconPixel / 64

With ParentForm.picToolbarOrg
    .Height = (IconPixel + 2 * (5 + iSize)) * Screen.TwipsPerPixelY
    .Width = Screen.Width   ' ParentForm.ScaleWidth
    .Top = 0
    .Left = 0
'    .Cls
    
    .BackColor = wPara1.FarbeArbeit
    
    ToolbarModus = 0
    sLw = UCase(Left(CurDir, 1))
    If (sLw <> "Z") Then
        ToolbarModus = 2
        
        Dim l&
        Dim sKey$
        For i = 0 To 9
            sKey = "Suffix" + CStr(i + 1)
            h$ = Space(51)
            l& = GetPrivateProfileString("SqlSuffix", sKey, h$, h$, 50, wPara1.LwZ + "\user\OpDrives.ini")
            h = Trim$(Left$(h$, l&))
            If (sLw = UCase(h)) Then
                If (i >= 1) Then
                    ToolbarModus = 6
                End If
                Exit For
            End If
        Next i
'        h = sqlop1.SqlGetSuffix
'        If (h <> "") Then
'            If (Val(Mid(h, 2)) > 1) Then
'                ToolbarModus = 6
'            End If
'        End If
    End If
    If (ToolbarModus Mod 8 >= 4) Then
        ToolbarBackR = 135
        ToolbarBackG = 61
        ToolbarBackB = 52
    Else
        ToolbarBackR = 201
        ToolbarBackG = 123
        ToolbarBackB = 58
    End If
    
    If (ToolbarModus >= 8) Then
        .BackColor = RGB(ToolbarBackR + 50, ToolbarBackG + 110, ToolbarBackB + 110)
    End If
    
    .ForeColor = RGB(210, 210, 210)
    
    .Font.Name = "Verdana"
    If (iSize = 0) Then
        .Font.Size = 10
    ElseIf (iSize = 1) Then
        .Font.Size = 10
    Else
        .Font.Size = 12
    End If
'    For j% = 14 To 5 Step -1
'        .Font.Size = 18  'nötig wegen Canon-BJ; sonst ab 2.Ausdruck falsch
'        .Font.Size = j%
'
'        If (.TextHeight("sF2") < (.Height / 3)) Then
'            Exit For
'        End If
'    Next j%
    ButtonWidth = (IconPixel + 5) * Screen.TwipsPerPixelX + .TextWidth("sF9")
    SpaceWidth = 270 ' ButtonWidth / 3
    
    For i% = 1 To 16
        ImgInd(i) = i
        ind = InStr(iErsatzBilder, "," + Format(i, "00"))
        If (ind > 0) Then
            ImgInd(i) = Val(Mid$(iErsatzBilder, ind + 3, 2))
        End If
        
        Call LoadToolbarButton(i)
        
        ParentForm.cmdToolbar(i).Visible = False
    Next i%
    ParentForm.cmdToolbar(0).Visible = False
            
'    .DrawWidth = 2
    ParentForm.picToolbarOrg.Line (0, 0)-(.ScaleWidth, 0), RGB(150, 150, 150)
    ParentForm.picToolbarOrg.Line (0, .ScaleHeight - 15)-(.ScaleWidth, .ScaleHeight - 15), RGB(150, 150, 150)
'    .DrawWidth = 1
    
    .Visible = False
End With

'ParentForm.picSave.Visible = False

Call clsError.DefErrPop
End Sub

Sub LoadToolbarButton(Index%, Optional ToolbarInd% = 0, Optional ButtonState% = 0)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("LoadToolbarButton")
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
Dim i%, j%, k%, m%, x%, IconWidth%, ff%, ind%, iAdd%
Dim lColor&, lColor2&, lColor3&, lColor4&, lColor5&
Dim h$, h2$
Dim iToolbar As PictureBox

If (ToolbarInd = 0) Then
    Set iToolbar = ParentForm.picToolbarOrg
Else
    Set iToolbar = ParentForm.picToolbar
End If

iAdd = 0
If (ButtonState = 1) Then
    iAdd = 30
End If


'On Error Resume Next

With iToolbar
    IconWidth = IconPixel * Screen.TwipsPerPixelX
    For i% = Index To Index
        x% = (i - 1) * (ButtonWidth + 15) + 210
        If (i% >= 9) Then x% = x% + SpaceWidth%
        
'        If (ToolbarModus = 0) Then
            iToolbar.Line (x, 0)-(x% + ButtonWidth, 315 * FaktorY), RGB(75 + iAdd, 75 + iAdd, 75 + iAdd), BF
            iToolbar.Line (x, 315)-(x% + ButtonWidth, 330 * FaktorY), RGB(80 + iAdd, 80 + iAdd, 80 + iAdd), BF
            iToolbar.Line (x, 345)-(x% + ButtonWidth, 360 * FaktorY), RGB(85 + iAdd, 85 + iAdd, 85 + iAdd), BF
            iToolbar.Line (x, 375)-(x% + ButtonWidth, 390 * FaktorY), RGB(90 + iAdd, 90 + iAdd, 90 + iAdd), BF
            Call wPara1.FillGradient(iToolbar, x / Screen.TwipsPerPixelX, (390 * FaktorY) / Screen.TwipsPerPixelY, (x% + ButtonWidth) / Screen.TwipsPerPixelX, .ScaleHeight / Screen.TwipsPerPixelY, RGB(40 + iAdd, 40 + iAdd, 40 + iAdd), RGB(160 + iAdd, 160 + iAdd, 160 + iAdd))
'        End If
        
        If (ToolbarModus Mod 4 = 1) Or (ToolbarModus Mod 4 = 3) Then
            iToolbar.Line (x, 0)-(x% + ButtonWidth, 390 * FaktorY), RGB(ToolbarBackR, ToolbarBackG, ToolbarBackB), BF
        End If
        If (ToolbarModus Mod 4 > 1) Then
            Call wPara1.FillGradient(iToolbar, x / Screen.TwipsPerPixelX, (390 * FaktorY) / Screen.TwipsPerPixelY, (x% + ButtonWidth) / Screen.TwipsPerPixelX, .ScaleHeight / Screen.TwipsPerPixelY, RGB(ToolbarBackR + 20, ToolbarBackG + 20, ToolbarBackB + 20), RGB(ToolbarBackR + 50, ToolbarBackG + 100, ToolbarBackB + 100))
        End If
            
        If (ButtonState = 2) Then
            iToolbar.Line (x, 0)-(x% + ButtonWidth, .ScaleHeight), RGB(200, 200, 200), BF
        End If
        
        If (ParentForm.cmdToolbar(i).Enabled) Then
            .ForeColor = RGB(210, 210, 210)
        Else
            .ForeColor = RGB(10, 10, 10)
        End If

        With ParentForm.picTemp
            .BorderStyle = 0
            .Left = ParentForm.ScaleWidth - 3000
            .Top = 1600
            
            .BackColor = RGB(68, 68, 68)
            lColor2 = RGB(67, 69, 66)
            lColor3 = RGB(67, 67, 67)
            lColor4 = RGB(66, 68, 65)
            lColor5 = RGB(65, 66, 64)

            .Width = 1000
            .Height = 1000  'IconHeight%
            .AutoRedraw = True
            .Cls
'MsgBox (CStr(iSize) + " " + CStr(i) + " " + CStr(ImgInd(i)))
            .PaintPicture ParentForm.imgToolbar(2 + iSize).ListImages(ImgInd(i)).Picture, 0, 0, IconWidth, IconWidth
 '           MsgBox ("done")

'            If (i = 6) Then
'                ff% = clsDat.FileOpen("\user\pixel.txt", "O")
'            End If
            
            m = IconPixel - 2
            For j = 2 To m
                For k = 2 To m
                    lColor = GetPixel(.hdc, j, k)
                    If (lColor <> .BackColor) And (lColor <> lColor2) And (lColor <> lColor3) And (lColor <> lColor4) And (lColor <> lColor5) Then
'                        If (i = 6) Then
'                            h = Right("000000" + Hex(lColor), 6)
'                            h2 = Str(Val("&H" + Left(h, 2))) + Str(Val("&H" + Mid(h, 3, 2))) + Str(Val("&H" + Mid(h, 5, 2)))
'                            Print #ff%, Str(j) + Str(k) + h2
'                        End If
                        SetPixel iToolbar.hdc, (x + 15) / Screen.TwipsPerPixelY + j, 9 + k, lColor
                    End If
                Next k
            Next j
'            If (i = 6) Then
'                Close #ff%
'            End If
        End With

        iToolbar.Line (x% + ButtonWidth, 90)-(x% + ButtonWidth + 15, .ScaleHeight - 90), RGB(170, 170, 170), BF
    Next i%
End With
Set iToolbar = Nothing

Call clsError.DefErrPop
End Sub

Sub ShowToolbar(Optional AktIndex% = 0, Optional ButtonState% = 1)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("ShowToolbar")
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
Dim i%, j%, diff%, x%
Dim h$

'On Error Resume Next

If (iNewLine = 0) Then
    Call clsError.DefErrPop: Exit Sub
End If

h$ = Trim(ParentForm.cmdToolbar(AktIndex).ToolTipText)
If (AktIndex > 0) And (Not (ParentForm.cmdToolbar(AktIndex).Enabled) Or (h = "")) Then
    AktIndex = 0
End If
'iAktIndex = AktIndex

With ParentForm.picToolbar
    .Width = ParentForm.ScaleWidth
    .Height = ParentForm.picToolbarOrg.Height
    .Top = 0
    .Left = 0
    
'    If (.BorderStyle = 1) Then
'        diff = .Height - .ScaleHeight
'        .Width = .Width + diff
'        .Height = .Height + diff
'        .Top = .Top - (diff / 2)
'        .Left = .Left - (diff / 2)
'        Call wPara1.ControlBorderless(ParentForm.picToolbar)
'        MsgBox (Str(.Height - .ScaleHeight))
'    End If
'    .Cls
    
    .BackColor = wPara1.FarbeArbeit
    
    BitBlt .hdc, 0, 0, .ScaleWidth, .ScaleHeight, ParentForm.picToolbarOrg.hdc, 0, 0, SRCCOPY
    
    Set .Font = ParentForm.picToolbarOrg.Font
    
    If (AktIndex > 0) Then
        Call LoadToolbarButton(AktIndex, 1, ButtonState)
'        Call ShowToolbarButton(Index)
    End If
    
    On Error Resume Next
    For i% = 1 To 16
        j% = i%
        If (i% < MENU_SF2) Then
            j% = i% - 1
        End If
        With ParentForm.mnuBearbeitenInd(j%)
            If (.Enabled) Then
                .Enabled = (.Caption <> "")
                ParentForm.cmdToolbar(i%).Enabled = .Enabled
            End If
        End With
    Next i%
    On Error GoTo DefErr
    
    For i% = 1 To 16
        Call ShowToolbarButton(i)
    Next i%
            
    If (AktIndex = 0) Then
        With ParentForm.picToolTip
            .Visible = False
        End With
    ElseIf (AktIndex <> iAktIndex) Then
        h$ = ParentForm.cmdToolbar(AktIndex).ToolTipText
        With ParentForm.picToolTip
            .Width = .TextWidth(h$ + "x")
            .Height = .TextHeight(h$) + 45
    
            x% = (AktIndex - 1) * (ButtonWidth + 15) + 210
            If (AktIndex% >= 9) Then x% = x% + SpaceWidth
            If (x + .Width > ParentForm.picToolbar.Width) Then
                x = ParentForm.picToolbar.Width - .Width - 90
            End If
            .Left = x
            .Top = ParentForm.picToolbar.Top + ParentForm.picToolbar.Height + 15
            .Visible = True
            .Cls
            .CurrentX = 2 * Screen.TwipsPerPixelX
            .CurrentY = 0
            ParentForm.picToolTip.Print h$
        End With
    End If
    
    .Visible = True
End With
iAktIndex = AktIndex

ParentForm.picSave.Visible = False

Call clsError.DefErrPop
End Sub

Sub ShowToolbarButton(Index%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("ShowToolbarButton")
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
Dim i%, x%, iAdd%
Dim h$

With ParentForm.picToolbar
    For i% = Index To Index
        x% = (i - 1) * (ButtonWidth + 15) + 210
        If (i% >= 9) Then x% = x% + SpaceWidth
        
        h$ = Trim(ParentForm.cmdToolbar(i).ToolTipText)
        If (ParentForm.cmdToolbar(i).Enabled) And (h <> "") Then
            .ForeColor = RGB(230, 230, 230)
            .ForeColor = RGB(230, 230, 230)
        Else
            .ForeColor = RGB(120, 120, 120)
            
'            If (ToolbarModus = 0) Then
                iAdd = 0
                ParentForm.picToolbar.Line (x, 0)-(x% + ButtonWidth, 315 * FaktorY), RGB(75 + iAdd, 75 + iAdd, 75 + iAdd), BF
                ParentForm.picToolbar.Line (x, 315)-(x% + ButtonWidth, 330 * FaktorY), RGB(80 + iAdd, 80 + iAdd, 80 + iAdd), BF
                ParentForm.picToolbar.Line (x, 345)-(x% + ButtonWidth, 360 * FaktorY), RGB(85 + iAdd, 85 + iAdd, 85 + iAdd), BF
                ParentForm.picToolbar.Line (x, 375)-(x% + ButtonWidth, 390 * FaktorY), RGB(90 + iAdd, 90 + iAdd, 90 + iAdd), BF
                Call wPara1.FillGradient(ParentForm.picToolbar, x / Screen.TwipsPerPixelX, (390 * FaktorY) / Screen.TwipsPerPixelY, (x% + ButtonWidth) / Screen.TwipsPerPixelX, .ScaleHeight / Screen.TwipsPerPixelY, RGB(40 + iAdd, 40 + iAdd, 40 + iAdd), RGB(160 + iAdd, 160 + iAdd, 160 + iAdd))
'            End If
        
            If (ToolbarModus Mod 4 = 1) Or (ToolbarModus Mod 4 = 3) Then
                ParentForm.picToolbar.Line (x, 0)-(x% + ButtonWidth, 390 * FaktorY), RGB(ToolbarBackR, ToolbarBackG, ToolbarBackB), BF
            End If
            If (ToolbarModus Mod 4 > 1) Then
                Call wPara1.FillGradient(ParentForm.picToolbar, x / Screen.TwipsPerPixelX, (390 * FaktorY) / Screen.TwipsPerPixelY, (x% + ButtonWidth) / Screen.TwipsPerPixelX, .ScaleHeight / Screen.TwipsPerPixelY, RGB(ToolbarBackR + 20, ToolbarBackG + 20, ToolbarBackB + 20), RGB(ToolbarBackR + 50, ToolbarBackG + 100, ToolbarBackB + 100))
            End If
                
            ParentForm.picToolbar.Line (x% + ButtonWidth, 90)-(x% + ButtonWidth + 15, .ScaleHeight - 90), RGB(170, 170, 170), BF
        End If

        ParentForm.picToolbar.Line (0, 0)-(.ScaleWidth, 0), RGB(150, 150, 150)
        ParentForm.picToolbar.Line (0, .ScaleHeight - 15)-(.ScaleWidth, .ScaleHeight - 15), RGB(150, 150, 150)
        
        If (i% <= 8) Then
            h$ = "F" + Format(i% + 1, "0")
        ElseIf (i% <= 16) Then
            h$ = "sF" + Format(i% - 7, "0")
        Else
            h$ = "a" + Format(i% - 16, "0")
        End If
        
'        .FontBold = True
        .CurrentX = x% + ButtonWidth - .TextWidth(h$) - 60 * wPara1.BildFaktor ' 150
        .CurrentY = (.Height / 3 - .TextHeight(h)) / 2 + 15 ' 60 * wPara1.BildFaktor 'cmdToolbarSize% + 45
        ParentForm.picToolbar.Print h$
        .Font.Bold = False
    Next i%
End With

Call clsError.DefErrPop
End Sub

Sub Click(x!)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("Click")
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
Dim start
Dim i%, xx%, Index%
Dim h$
    
With ParentForm.picToolbar
    If (x = 9999) Then
'        ToolbarModus = ToolbarModus + 1
'        If (ToolbarModus > 15) Then
'            ToolbarModus = 0
'        End If
'        Call LoadToolbar
'        RaiseEvent Resized
        Call clsError.DefErrPop: Exit Sub
    End If
    
    If (x < 0) Then
        Index = Abs(x)
    Else
        Index = 0
        xx = 210
        For i% = 1 To 16
            If (i% = 9) Then xx% = xx% + SpaceWidth
            
            If (x >= xx) And (x <= xx% + ButtonWidth) Then
                Index = i
                Exit For
            End If
            
            xx% = xx% + ButtonWidth + 15
        Next i%
    End If
    
    If (Index = 0) Then
        iSize = (iSize + 1) Mod 3
        Call LoadToolbar
        RaiseEvent Resized
    ElseIf (ParentForm.cmdToolbar(Index).Enabled) Then
        Call ShowToolbar(Index, 2)
        
        DoEvents
        start = Timer
        Do
            If (Timer - start) > 0.2 Then
                Exit Do
            End If
        Loop
        Call ShowToolbar(Index)
        
        If (x > 0) Then
            ParentForm.cmdToolbar(Index).Value = True
        End If
    End If
End With

Call clsError.DefErrPop
End Sub
       
Sub MouseMove(x!)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("MouseMove")
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
Dim i%, xx%, Index%
Dim h$
    
With ParentForm.picToolbar
    Index = 0
    xx = 210
    For i% = 1 To 16
        If (i% = 9) Then xx% = xx% + SpaceWidth
        
        If (x >= xx) And (x <= xx% + ButtonWidth) Then
            Index = i
            Exit For
        End If
        
        xx% = xx% + ButtonWidth + 15
    Next i%
    
    Call ShowToolbar(Index)
End With

Call clsError.DefErrPop
End Sub

Sub LoadIcons()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("LoadIcons")
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
Dim i%, j%, k%, gef%, ind%
Dim h$, ProgName$

With ParentForm.imgToolbar(2).ListImages
    If (.Count > 0) Then
        Call clsError.DefErrPop: Exit Sub
    End If
End With

ProgName = iIniDatei
ind = InStr(ProgName, ".")
If (ind > 0) Then
    ProgName = Left(ProgName, ind - 1)
End If
ProgName = UCase(ProgName)

Call SearchIcons(ProgName)
For i = 0 To 2
    For j = 1 To 50
        h = Format(j, "00") + "_" + CStr(i + 1) + ".GIF"
        gef = 0
        'MsgBox (h)
        With ParentForm.lstSortierung
            For k = 0 To (.ListCount - 1)
                If (.List(k) = h) Then
                    gef = True
                    'MsgBox ("gef")
                    Exit For
                End If
            Next k
        End With
        With ParentForm.imgToolbar(i + 2).ListImages
            If (gef) Then
                .Add , , LoadPicture(CurDir + "\NewLine\" + ProgName + "\" + Mid$(ParentForm.Name, 4) + "\" + h)
            'ElseIf (j <= 16) Then
            Else
'                .Add , , LoadPicture(CurDir + "\NewLine\Leer.jpg")
                .Add , , LoadPicture(CurDir + "\NewLine\Leer.gif")
            End If
            'MsgBox (".. " + CStr(i) + " " + CStr(ParentForm.imgToolbar(i + 2).ListImages.Count))
        End With
    Next j
    'MsgBox ("LI " + CStr(i) + " " + CStr(ParentForm.imgToolbar(i + 2).ListImages.Count))
Next i

Call clsError.DefErrPop
End Sub

Sub SearchIcons(ProgName$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("SearchIcons")
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
Dim h$, DirMask$, EntryName$, MinTag$
Dim SearchHandle&
Dim FindDataRec As WIN32_FIND_DATA
Dim erg%, ind%, i%, k%, gef%

ParentForm.lstSortierung.Clear
DirMask$ = CurDir + "\NewLine\" + ProgName + "\" + Mid$(ParentForm.Name, 4) + "\*.*"

SearchHandle& = FindFirstFile(DirMask$, FindDataRec)
If (SearchHandle& = INVALID_HANDLE_VALUE) Then Call clsError.DefErrPop: Exit Sub
Do
    h$ = FindDataRec.cFileName
    ind% = InStr(h$, Chr$(0))
    If (ind% > 0) Then h$ = Left$(h$, ind% - 1)
    EntryName$ = h$
    
    If ((EntryName$ = ".") Or (EntryName$ = "..")) Then
    ElseIf (FindDataRec.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) Then
    Else
        ParentForm.lstSortierung.AddItem UCase(EntryName)
    End If
    
    erg% = FindNextFile(SearchHandle, FindDataRec)
    If (erg% = 0) Then Exit Do
Loop
erg% = FindClose(SearchHandle&)

Call clsError.DefErrPop
End Sub


