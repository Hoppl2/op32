VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsOpBereiche"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim iArbeitBackHeight%
Dim iArbeitAnzzeilen%
Dim iArbeitTitel%
Dim iArbeitLeerzeileOben%
Dim iArbeitWasDarunter%
Dim iInfoBackHeight%
Dim iInfoAnzZeilen%
Dim iInfoTitel%
Dim iInfoZusatz%
Dim iAnzahlButtons%

Dim iBreiteX%
Dim iGesamtY%
Dim iLinksOkX%
Dim iLinksEscX%
Dim iLinksOkEscX%
Dim iZeilenHoeheY%

Dim iAutoRedraw%

Dim ParentForm As Form
Dim ParentToolbar As clsToolbar


Private Const DefErrModul = "OPBEREICHE.CLS"

Sub InitBereich(iForm As Object, iToolbar As Object)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("InitBereich")
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
Set ParentForm = iForm
Set ParentToolbar = iToolbar

iAutoRedraw = True
If (iNewLine) Then
    iInfoTitel% = True
End If

Call clsError.DefErrPop
End Sub

Public Property Get ArbeitBackHeight%()
ArbeitBackHeight% = iArbeitBackHeight%
End Property
Public Property Let ArbeitBackHeight(ByVal vNewValue%)
iArbeitBackHeight% = vNewValue%
End Property

Public Property Get ArbeitAnzZeilen%()
ArbeitAnzZeilen% = iArbeitAnzzeilen%
End Property
Public Property Let ArbeitAnzZeilen(ByVal vNewValue%)
iArbeitAnzzeilen% = vNewValue%
End Property

Public Property Get ArbeitTitel%()
ArbeitTitel% = iArbeitTitel%
End Property
Public Property Let ArbeitTitel(ByVal vNewValue%)
iArbeitTitel% = vNewValue%
End Property

Public Property Get ArbeitLeerzeileOben%()
ArbeitLeerzeileOben% = iArbeitLeerzeileOben%
End Property
Public Property Let ArbeitLeerzeileOben(ByVal vNewValue%)
iArbeitLeerzeileOben% = vNewValue%
End Property

Public Property Get ArbeitWasDarunter%()
ArbeitWasDarunter% = iArbeitWasDarunter%
End Property
Public Property Let ArbeitWasDarunter(ByVal vNewValue%)
iArbeitWasDarunter% = vNewValue%
End Property

Public Property Get InfoBackHeight%()
InfoBackHeight% = iInfoBackHeight%
End Property
Public Property Let InfoBackHeight(ByVal vNewValue%)
iInfoBackHeight% = vNewValue%
End Property

Public Property Get InfoAnzZeilen%()
InfoAnzZeilen% = iInfoAnzZeilen%
End Property
Public Property Let InfoAnzZeilen(ByVal vNewValue%)
iInfoAnzZeilen% = vNewValue%
End Property

Public Property Get InfoTitel%()
InfoTitel% = iInfoTitel%
End Property
Public Property Let InfoTitel(ByVal vNewValue%)
If (iNewLine = 0) Then
    iInfoTitel% = vNewValue%
End If
End Property

Public Property Get InfoZusatz%()
InfoZusatz% = iInfoZusatz%
End Property
Public Property Let InfoZusatz(ByVal vNewValue%)
iInfoZusatz% = vNewValue%
End Property

Public Property Get AnzahlButtons%()
AnzahlButtons% = iAnzahlButtons%
End Property
Public Property Let AnzahlButtons(ByVal vNewValue%)
iAnzahlButtons% = vNewValue%
End Property

Public Property Get ZeilenHoeheY%()
ZeilenHoeheY% = iZeilenHoeheY%
End Property
Public Property Let ZeilenHoeheY(ByVal vNewValue%)
iZeilenHoeheY% = vNewValue%
End Property

Public Property Get AutoRedraw%()
AutoRedraw% = iAutoRedraw%
End Property
Public Property Let AutoRedraw(ByVal vNewValue%)
iAutoRedraw% = vNewValue%
End Property

Sub ResizeWindow()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("ResizeWindow")
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
Dim i%, BackX%, BackY%, ToolX%, ToolY%
Dim Suc&
    
'With iForm
'    If (.Name = frmAction.Name) Then
'        .MousePointer = vbHourglass
'        .picStartBack.Visible = True
'        Suc& = BitBlt(.picStartBack.hdc, 0, 0, .ScaleWidth, .ScaleHeight, .hdc, 0, 0, &HCC0020)
'        iForm.picStartBack.ZOrder 0
'        Call InitBereicheBack(iForm, iToolbar)
'        Call InitBereich(iForm, iBereich, iBereichInd%)
'        .MousePointer = vbDefault
'        .picStartBack.Visible = False
'        Call ZeigeBestellWerte(True)
'    Else
'        Call InitBereicheBack(iForm, iToolbar)
'        Call InitBereich(iForm, iBereich, iBereichInd%)
'    End If
'End With

With ParentForm
'    .picSave.Width = .Width
'    .picSave.Height = .Height
'    .picSave.Visible = True
'    Suc& = BitBlt(.picSave.hdc, 0, 0, .ScaleWidth, .ScaleHeight, .hdc, 0, 0, &HCC0020)
'    .picSave.ZOrder 0
    Call RefreshBereicheBack
    Call RefreshBereich
    .picBack(0).Visible = True
'    .picSave.Visible = False
End With

Call clsError.DefErrPop
End Sub

Sub RefreshBereicheBack()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("RefreshBereicheBack")
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
Dim i%, ToolbarPos%

On Error Resume Next
ToolbarPos% = ParentToolbar.Position

Call ParentToolbar.ResizeToolbar

With ParentForm
    If (iNewLine) Then
        .picBack(0).Top = .picToolbar.Height
        .picBack(0).Height = .ScaleHeight - .picToolbar.Height - (3 * Screen.TwipsPerPixelX + wPara1.FrmCaptionHeight) + 60
        .picBack(0).Left = 0
        .picBack(0).Width = .ScaleWidth
    ElseIf (ParentToolbar.Visible) Then
        If (ToolbarPos% = 0) Then
            .picBack(0).Top = .picToolbar.Height
            .picBack(0).Height = .ScaleHeight - .picToolbar.Height
            .picBack(0).Left = 0
            .picBack(0).Width = .ScaleWidth
        ElseIf (ToolbarPos% = 1) Then
            .picBack(0).Top = 0
            .picBack(0).Height = .ScaleHeight
            .picBack(0).Left = 0
            .picBack(0).Width = .ScaleWidth - .picToolbar.Width
        ElseIf (ToolbarPos% = 2) Then
            .picBack(0).Top = 0
            .picBack(0).Height = .picToolbar.Top
            .picBack(0).Left = 0
            .picBack(0).Width = .ScaleWidth
        ElseIf (ToolbarPos% = 3) Then
            .picBack(0).Top = 0
            .picBack(0).Height = .ScaleHeight
            .picBack(0).Left = .picToolbar.Width
            .picBack(0).Width = .ScaleWidth - .picToolbar.Width
        Else
            .picBack(0).Top = 0
            .picBack(0).Height = .ScaleHeight
            .picBack(0).Left = 0
            .picBack(0).Width = .ScaleWidth
        End If
    Else
        .picBack(0).Top = 0
        .picBack(0).Height = .ScaleHeight
        .picBack(0).Left = 0
        .picBack(0).Width = .ScaleWidth
    End If

'    For i% = 1 To UBound(iForm.Bereiche)
'        .picBack(i%).Top = .picBack(0).Top
'        .picBack(i%).Height = .picBack(0).Height
'        .picBack(i%).Left = .picBack(0).Left
'        .picBack(i%).Width = .picBack(0).Width
'    Next i%
End With

'For i% = 0 To UBound(iForm.Bereiche)
'    iForm.Bereiche(i%).BereichOk = False
'Next i%

Call RefreshBereichsKonstanten

Call clsError.DefErrPop
End Sub

Sub RefreshBereichsKonstanten()

'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("RefreshBereichsKonstanten")
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
With ParentForm
    iBreiteX% = .picBack(0).ScaleWidth - 2 * wPara1.LinksX
    iGesamtY% = .picBack(0).ScaleHeight
    
    iLinksOkX% = (.picBack(0).Width - (wPara1.ButtonX * 2 + 300)) / 2
    iLinksEscX% = iLinksOkX% + wPara1.ButtonX + 300
    iLinksOkEscX% = (.picBack(0).Width - wPara1.ButtonX) / 2
    
    With .flxarbeit(0)
        If (.Rows < 1) Then .Rows = 1
        iZeilenHoeheY% = .RowHeight(0)
    End With
End With
    
Call clsError.DefErrPop
End Sub

Sub RefreshBereich()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("RefreshBereich")
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
ParentForm.picBack(0).Visible = False
Call RefreshBereichsControls
Call RefreshBereichsFarben
With ParentForm
    Call .RefreshBereichsFlexSpalten
    Call .RefreshBereichsControlsAdd
    Call .RefreshBereichsFarbenAdd
End With
ParentForm.picBack(0).Visible = True
On Error Resume Next
ParentForm.ActiveControl.SetFocus

Call clsError.DefErrPop
End Sub

Sub RefreshBereichsControls()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("RefreshBereichsControls")
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
Dim i%, iAnzButtons%, x1%, x2%, y1%, Y2%, iAdd%, iAdd2%, iZeilenHoeheY2%
Dim sp&

'On Error Resume Next

With ParentForm
    .lblArbeit(0).Left = wPara1.LinksX
    .flxarbeit(0).Left = wPara1.LinksX
    .lblInfo(0).Left = wPara1.LinksX
    .flxInfo(0).Left = wPara1.LinksX
    .flxInfoZusatz(0).Left = wPara1.LinksX
    
    .lblArbeit(0).Width = iBreiteX%
    .flxarbeit(0).Width = iBreiteX%
    .lblInfo(0).Width = iBreiteX%
    .flxInfo(0).Width = iBreiteX%
    .flxInfoZusatz(0).Width = iBreiteX%
    
    
    .lblArbeit(0).Top = wPara1.TitelY
    If (iArbeitTitel%) Then
        .lblArbeit(0).Visible = True
        .flxarbeit(0).Top = wPara1.FlexY
    Else
        .lblArbeit(0).Visible = False
        .flxarbeit(0).Top = wPara1.TitelY
    End If
    If (iArbeitLeerzeileOben%) Then
        .flxarbeit(0).Top = .flxarbeit(0).Top + wPara1.FlexY
    End If
    
    .lblInfo(0).Top = wPara1.TitelY
    If (iInfoTitel%) Then
        .lblInfo(0).Visible = True
        .flxInfo(0).Top = wPara1.FlexY
    Else
        .lblInfo(0).Visible = False
        .flxInfo(0).Top = wPara1.TitelY
    End If

    .flxInfoZusatz(0).Height = iZeilenHoeheY% * 2 + 90
    .flxInfoZusatz(0).Top = wPara1.TitelY
    If (iInfoZusatz%) Then
        .flxInfoZusatz(0).Visible = True
        .flxInfo(0).Top = .flxInfoZusatz(0).Top + .flxInfoZusatz(0).Height + 15
    Else
        .flxInfoZusatz(0).Visible = False
'        .flxInfo(0).Top = wPara1.TitelY
    End If

End With


With ParentForm.flxInfo(0)
    .Height = iZeilenHoeheY% * iInfoAnzZeilen% + 90
    If (iInfoAnzZeilen% > 0) Then
        iInfoBackHeight% = .Top + .Height + wPara1.UntenFreiY
    Else
        iInfoBackHeight% = wPara1.OhneInfoY
        ParentForm.lblInfo(0).Caption = "Keine zusätzlichen Informationen vorhanden"
        .Visible = False
    End If
End With
    
iArbeitBackHeight% = iGesamtY% - iInfoBackHeight%
With ParentForm.flxarbeit(0)
    .Height = iArbeitBackHeight% - wPara1.UntenFreiY - .Top
    If (iArbeitWasDarunter%) Then
        .Height = .Height - iZeilenHoeheY% - 90
    End If
    iArbeitAnzzeilen% = (.Height - 90) \ iZeilenHoeheY%
    .Height = iZeilenHoeheY% * iArbeitAnzzeilen% + 90
End With

With ParentForm
    .lblInfo(0).Top = .lblInfo(0).Top + iArbeitBackHeight%
    .flxInfo(0).Top = .flxInfo(0).Top + iArbeitBackHeight%
    .flxInfoZusatz(0).Top = .flxInfoZusatz(0).Top + iArbeitBackHeight%
    
    
    .cmdOk(0).Width = wPara1.ButtonX
    .cmdOk(0).Height = wPara1.ButtonY
    .cmdEsc(0).Width = wPara1.ButtonX
    .cmdEsc(0).Height = wPara1.ButtonY
    
    iAnzButtons = Abs(iAnzahlButtons%)
    If (iAnzButtons% = 1) Then
        .cmdOk(0).Left = iLinksOkEscX%
        .cmdEsc(0).Left = iLinksOkEscX%
    Else
        .cmdOk(0).Left = iLinksOkX%
        .cmdEsc(0).Left = iLinksEscX%
    End If
    
    If (iAnzahlButtons% > 0) Then
        .cmdOk(0).Top = iArbeitBackHeight% - .cmdOk(0).Height - 60
    Else
        .cmdOk(0).Top = .picBack(0).Height + 90
    End If
    .cmdEsc(0).Top = .cmdOk(0).Top
    
End With

If (iNewLine) Then
    ParentForm.picBack(0).BackColor = wPara1.FarbeArbeit
    
    x1 = ParentForm.flxarbeit(0).Left
    x2 = ParentForm.flxarbeit(0).Left + ParentForm.flxarbeit(0).Width
    y1 = ParentForm.lblInfo(0).Top
    Y2 = ParentForm.flxInfo(0).Top + ParentForm.flxInfo(0).Height
    
    iAdd = wPara1.NlFlexBackY
    iAdd2 = 120 * iFaktorY
    
    ParentForm.lblInfo(0).Visible = False
    With ParentForm.flxInfo(0)
        iZeilenHoeheY2 = iZeilenHoeheY + 60
        For i = 0 To (.Rows - 1)
            .RowHeight(i) = iZeilenHoeheY2
        Next i
        .Height = iZeilenHoeheY2 * iInfoAnzZeilen% '+ 90
        .Left = x1 + iAdd
'        .Top = y1 + ((y2 - y1) - .Height) / 2
        .Top = ParentForm.picBack(0).ScaleHeight - .Height - iAdd - iAdd2 ' 90
        .Width = (x2 - x1) - 2 * iAdd
        .ScrollBars = flexScrollBarNone
        .BorderStyle = 0
        y1 = .Top - iAdd
    End With

    With ParentForm.flxInfoZusatz(0)
        .Rows = 2
        .Cols = 15
        .FixedRows = 0
        .FixedCols = 0
        
        .RowHeight(0) = iZeilenHoeheY
        .RowHeight(1) = iZeilenHoeheY2
        
        .Height = iZeilenHoeheY% + iZeilenHoeheY2
        .Width = (x2 - x1) - 2 * iAdd
        .Left = x1 + iAdd
        .Top = ParentForm.flxInfo(0).Top - .Height - 2 * iAdd - iAdd2
        
        .ScrollBars = flexScrollBarNone
        .BorderStyle = 0
        
        If (iInfoZusatz%) Then
            .Visible = True
            y1 = .Top - iAdd
        Else
            .Visible = False
    '        .flxInfo(0).Top = wPara1.TitelY
        End If
    End With
    
    With ParentForm.flxarbeit(0)
        If (.Tag <> "0") Then
            .Left = x1 + iAdd
            .Top = .Top + iAdd
            .Width = (x2 - x1) - 2 * iAdd
            .Height = .Height - iAdd
            .BorderStyle = 0
            .GridLines = flexGridFlat
            .GridLinesFixed = .GridLines
            .GridColorFixed = .GridColor
            
            .BackColor = wPara1.nlFlexBackColor 'vbWhite
            .BackColorBkg = wPara1.nlFlexBackColor  'vbWhite
            .BackColorFixed = wPara1.nlFlexBackColorFixed   ' RGB(199, 176, 123)
            If (.SelectionMode = flexSelectionFree) Then
                .BackColorSel = RGB(135, 61, 52)
                .ForeColorSel = vbWhite '.ForeColor
            Else
                .BackColorSel = wPara1.nlFlexBackColorSel  ' RGB(232, 217, 172)
                .ForeColorSel = .ForeColor
            End If
            .Appearance = 0
        
'            With ParentForm.picBack(0)
'        '        ParentForm.picBack(0).Line (0, y2 - 990)-(.Width, .Height), RGB(180, 180, 180), BF
'                .ForeColor = RGB(180, 180, 180) ' vbWhite
'                .FillStyle = vbSolid
'                .FillColor = vbWhite
'                RoundRect .hdc, x1 / Screen.TwipsPerPixelX, (ParentForm.flxarbeit(0).Top - 300) / Screen.TwipsPerPixelY, x2 / Screen.TwipsPerPixelX, (ParentForm.flxarbeit(0).Top + ParentForm.flxarbeit(0).Height + iAdd) / Screen.TwipsPerPixelY, 20, 20
'            End With
        End If
    End With
    With ParentForm.flxarbeit(0)
        .Height = y1 - wPara1.UntenFreiY - .Top
        If (iArbeitWasDarunter%) Then
            .Height = .Height - iZeilenHoeheY% - 90
        End If
        iArbeitAnzzeilen% = .Height \ iZeilenHoeheY%
        .Height = iZeilenHoeheY% * iArbeitAnzzeilen% '+ 90
    End With
    
    With ParentForm.picBack(0)
'        ParentForm.picBack(0).Line (0, y2 - 960)-(.Width, .Height), RGB(180, 180, 180), BF
        Y2 = .Height - (70 * iFaktorY - 10) * Screen.TwipsPerPixelY
        ParentForm.picBack(0).Line (0, Y2)-(.Width, .Height), RGB(180, 180, 180), BF
        .ForeColor = RGB(180, 180, 180) ' vbWhite
        .FillStyle = vbSolid
        .FillColor = vbWhite
        RoundRect .hdc, x1 / Screen.TwipsPerPixelX, (ParentForm.flxInfo(0).Top - iAdd) / Screen.TwipsPerPixelY, x2 / Screen.TwipsPerPixelX, (ParentForm.flxInfo(0).Top + ParentForm.flxInfo(0).Height + iAdd) / Screen.TwipsPerPixelY, 20, 20
        If (iInfoZusatz%) Then
            RoundRect .hdc, x1 / Screen.TwipsPerPixelX, (ParentForm.flxInfoZusatz(0).Top - iAdd) / Screen.TwipsPerPixelY, x2 / Screen.TwipsPerPixelX, (ParentForm.flxInfoZusatz(0).Top + ParentForm.flxInfoZusatz(0).Height + iAdd) / Screen.TwipsPerPixelY, 20, 20
        End If
'        RoundRect .hdc, x1 / Screen.TwipsPerPixelX, (ParentForm.flxarbeit(0).Top - iAdd) / Screen.TwipsPerPixelY, x2 / Screen.TwipsPerPixelX, (ParentForm.flxarbeit(0).Top + ParentForm.flxarbeit(0).Height + iAdd) / Screen.TwipsPerPixelY, 20, 20
        If (ParentForm.flxarbeit(0).Tag <> "0") Then
            RoundRect .hdc, x1 / Screen.TwipsPerPixelX, (ParentForm.flxarbeit(0).Top - iAdd) / Screen.TwipsPerPixelY, x2 / Screen.TwipsPerPixelX, (y1 - wPara1.UntenFreiY) / Screen.TwipsPerPixelY, 20, 20
        End If
    End With
End If


Call clsError.DefErrPop
End Sub

Sub RefreshBereichsFarben()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("RefreshBereichsFarben")
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
On Error Resume Next

With ParentForm
    .picBack(0).AutoRedraw = iAutoRedraw    'True
    
    If (iNewLine) Then
    Else
        .picBack(0).AutoRedraw = True
    
        .picBack(0).Line (0, 0)-(.ScaleWidth, iArbeitBackHeight%), wPara1.FarbeArbeit, BF
        .picBack(0).Line (0, iArbeitBackHeight%)-(.ScaleWidth, .ScaleHeight), wPara1.FarbeInfo, BF
        
        If (wPara1.FarbeRahmen) Then
            .picBack(0).DrawWidth = 3
            .picBack(0).Line (0, 0)-(.ScaleWidth, iArbeitBackHeight%), vbBlack, B
            .picBack(0).Line (0, iArbeitBackHeight%)-(.ScaleWidth, .ScaleHeight), vbBlack, B
        End If
        
        .lblArbeit(0).BackColor = wPara1.FarbeArbeit
        .flxarbeit(0).BackColorBkg = wPara1.FarbeArbeit
        
        .lblInfo(0).BackColor = wPara1.FarbeInfo
        .flxInfo(0).BackColorBkg = wPara1.FarbeInfo
    End If
End With
    
Call clsError.DefErrPop
End Sub

Public Sub WechselModus(NeuerModus%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("WechselModus")
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
Dim i%, ind%, erg%
Dim lColor&
Dim h$
Static EnabledArray%(20), SelectionModus&

On Error Resume Next
With ParentForm
    Select Case NeuerModus%
        Case 0
            .mnuDatei.Enabled = True
            .mnuBearbeiten.Enabled = True
            .mnuAnsicht.Enabled = True
            
'            On Error Resume Next
            For i% = MENU_F2 To MENU_SF9
                .mnuBearbeitenInd(i%).Enabled = EnabledArray%(i%)
            Next i%
'            On Error GoTo DefErr
        
            .mnuBearbeitenLayout.Checked = False
            
            .cmdOk(0).default = True
            .cmdEsc(0).Cancel = True
            
            With .flxarbeit(0)
                If (.Tag <> "0") Then
                    .SelectionMode = SelectionModus
                
                    If (iNewLine) Then
                        If (.SelectionMode = flexSelectionByRow) Then
                            lColor = RGB(232, 217, 172)
                        Else
                            lColor = RGB(135, 61, 52)
                        End If
                    Else
                        lColor = vbHighlight
                    End If
                    .BackColorSel = lColor
                        
                    If (.SelectionMode = flexSelectionByRow) Then
                        .col = 0
                        .ColSel = .Cols - 1
                    End If
                End If
            End With
            .flxInfo(0).BackColorSel = lColor
            
            .flxInfo(0).Enabled = False
            
            h$ = .Caption
            ind% = InStr(h$, " (EDITIER-MODUS)")
            If (ind% > 0) Then h$ = Left$(h$, ind% - 1)
            .Caption = h$
            
        Case 1
            .mnuDatei.Enabled = False
            .mnuBearbeiten.Enabled = True
            .mnuAnsicht.Enabled = False
            
            For i% = MENU_F2 To MENU_SF9
                EnabledArray%(i%) = .mnuBearbeitenInd(i%).Enabled
            Next i%
        
            .mnuBearbeitenInd(MENU_F2).Enabled = True
            .mnuBearbeitenInd(MENU_F3).Enabled = True
            .mnuBearbeitenInd(MENU_F4).Enabled = False
            .mnuBearbeitenInd(MENU_F5).Enabled = True
            .mnuBearbeitenInd(MENU_F6).Enabled = False
            .mnuBearbeitenInd(MENU_F7).Enabled = False
            .mnuBearbeitenInd(MENU_F8).Enabled = True
            .mnuBearbeitenInd(MENU_F9).Enabled = False
            .mnuBearbeitenInd(MENU_SF2).Enabled = False
            .mnuBearbeitenInd(MENU_SF3).Enabled = False
            .mnuBearbeitenInd(MENU_SF4).Enabled = False
            .mnuBearbeitenInd(MENU_SF5).Enabled = False
            .mnuBearbeitenInd(MENU_SF6).Enabled = False
            .mnuBearbeitenInd(MENU_SF7).Enabled = False
            .mnuBearbeitenInd(MENU_SF8).Enabled = False
            
            .mnuBearbeitenLayout.Checked = True
            
            .cmdOk(0).default = True
            .cmdEsc(0).Cancel = True
            
            With .flxarbeit(0)
                If (.Tag <> "0") Then
                    SelectionModus = .SelectionMode
                    .SelectionMode = flexSelectionFree
                    .col = .Cols - 1
                    .ColSel = .col
                    .BackColorSel = vbMagenta
                End If
            End With
            
            With .flxInfo(0)
                .BackColorSel = vbMagenta
                .row = 0
                .col = 1
                .Enabled = True
                On Error Resume Next
                .SetFocus
                On Error GoTo DefErr
            End With
            
            h$ = .Caption
            .Caption = h$ + " (EDITIER-MODUS)"
    End Select
    
    For i% = 0 To 7
        .cmdToolbar(i% + 1).Enabled = .mnuBearbeitenInd(i%).Enabled
    Next i%
    For i% = 8 To 15
        .cmdToolbar(i% + 1).Enabled = .mnuBearbeitenInd(i% + 1).Enabled
    Next i%
End With

Call clsError.DefErrPop
End Sub

Sub HighlightZeile(Optional NurNormalMachen% = False)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("HighlightZeile")
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
Dim arow%, aCol%, rInd%
Dim RkLaufNr&, lColor&
Dim h$, KalkText$, DirektWerte$
Static aRkLaufNr&, aFlexRow%

With ParentForm.flxarbeit(0)
    If (.SelectionMode = flexSelectionFree) Then
        If (NurNormalMachen%) Then
            .HighLight = flexHighlightNever
    
            arow% = .row
            aCol% = .col
    
            If (aFlexRow% < .Rows) Then
                .FillStyle = flexFillRepeat
                .row = aFlexRow%
                If (iNewLine) Then
                    .col = .Cols - 1
                    lColor = .CellBackColor
                    .col = 0
                    If (.FixedCols > 0) Then
                        .ColSel = .FixedCols - 1
                        .CellBackColor = .BackColorFixed
                        .col = .FixedCols
                    End If
                    .ColSel = .Cols - 1
                    .CellBackColor = lColor 'vbWhite
                Else
                    .col = 0
                    .ColSel = .Cols - 1
                    .CellForeColor = .ForeColor
                End If
                .FillStyle = flexFillSingle
                .col = aCol%
            End If
    
            aRkLaufNr& = -1&
            aFlexRow% = .row
            .HighLight = flexHighlightWithFocus
        Else
            RkLaufNr& = Val(.TextMatrix(.row, 3))
            If (RkLaufNr& <> aRkLaufNr&) Or (aFlexRow% <> .row) Then
    
                .HighLight = flexHighlightNever
    
                arow% = .row
                aCol% = .col
    
                .FillStyle = flexFillRepeat
    
                If (aFlexRow% < .Rows) Then
                    .row = aFlexRow%
                    If (iNewLine) Then
                        .col = .Cols - 1
                        lColor = .CellBackColor
                        .col = 0
                        If (.FixedCols > 0) Then
                            .ColSel = .FixedCols - 1
                            .CellBackColor = .BackColorFixed
                            .col = .FixedCols
                        End If
                        .ColSel = .Cols - 1
                        .CellBackColor = lColor 'vbWhite
                    Else
                        .col = 0
                        .ColSel = .Cols - 1
                        .CellForeColor = .ForeColor
                    End If
                    .row = arow%
                End If
    
                If (iNewLine) Then
                    .col = 0    '.FixedCols   ' 2
                    .ColSel = .Cols - 2
                    .CellBackColor = RGB(232, 217, 172)
                Else
                    .col = 0
                    .ColSel = .Cols - 1
                    .CellForeColor = wPara1.FarbeAktZeile
                End If
                
                .FillStyle = flexFillSingle
                .col = aCol%
    
                Call ParentForm.ZeigeInfo
                
                aRkLaufNr& = RkLaufNr&
                aFlexRow% = .row
                .HighLight = flexHighlightWithFocus
            End If
        End If
                
    '    .col = 0
    '    .ColSel = .Cols - 1
    Else
        Call ParentForm.ZeigeInfo
    End If
End With

Call clsError.DefErrPop
End Sub

Sub ZeigeInfoZusatz(pzn$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("ZeigeInfoZusatz")
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
Dim i%, j%, Monat%, erg%
Dim Jahr&, AltJahr&, Termin&, sp&
Dim iWert!
Dim ArtStat1 As clsArtStatistik

If (iInfoZusatz = 0) Then
    Call clsError.DefErrPop: Exit Sub
End If

With ParentForm.flxInfoZusatz(0)
    .Redraw = False
    
    If (iNewLine) Then
        .GridLines = flexGridFlat
        .BorderStyle = flexBorderNone
        .SelectionMode = flexSelectionFree
        
        .Rows = 2
        .Cols = 15
        .FixedRows = 0
        .FixedCols = 0
        
        .RowHeight(0) = iZeilenHoeheY
        .RowHeight(1) = iZeilenHoeheY + 60
        .Height = .RowHeight(0) + .RowHeight(1)
        
        sp& = .Width / 15 + 15
        For i% = 0 To 14
            .ColWidth(i%) = sp&
        Next i%
        
        For i% = 0 To 1
            .FillStyle = flexFillRepeat
            .row = i
            .col = 0
            .RowSel = .row
            .ColSel = .Cols - 1
            If (i% = 0) Then
                .CellBackColor = wPara1.nlFlexBackColorFixed    ' RGB(199, 176, 123)
            Else
                .CellBackColor = wPara1.nlFlexBackColorInfo     ' RGB(232, 217, 172)
                .CellFontSize = .Font.Size + 2
                .CellFontBold = False
            End If
            .CellAlignment = flexAlignCenterCenter
            .FillStyle = flexFillSingle
        Next i%
    Else
        .GridLines = flexGridInset
    
        .SelectionMode = flexSelectionFree
        .Rows = 2
        .FixedRows = 1
        .Cols = 15
        
        sp& = .Width / 15
        For i% = 0 To 14
            .ColWidth(i%) = sp&
        Next i%
        
        .FillStyle = flexFillRepeat
        .col = 0
        .row = 1
        .ColSel = .Cols - 1
        .RowSel = .Rows - 1
        .CellBackColor = vbWhite
        
        .col = 0
        .row = 0
        .ColSel = .Cols - 1
        .RowSel = .Rows - 1
        .CellAlignment = flexAlignCenterCenter
        .FillStyle = flexFillSingle
    End If
    
    Set ArtStat1 = New clsArtStatistik
    erg% = ArtStat1.StatistikRechnen(pzn$)
    If (erg%) Then
        AltJahr& = -1
        j% = 0
        For i% = 0 To 12
            Termin& = ArtStat1.Anfang - i% - 1
            Jahr& = (Termin& - 1) \ 12
            Monat% = Termin& - Jahr& * 12
            If (AltJahr& <> Jahr&) Then
                .TextMatrix(0, j%) = Str$(Jahr&)
                iWert! = ArtStat1.JahresWert(Jahr&)
                If (iWert! <> 0!) Then
                    .TextMatrix(1, j%) = Str$(iWert!)
                Else
                    .TextMatrix(1, j%) = ""
                End If
                .row = 0
                .col = j%
                .CellFontBold = True
                .CellFontSize = .Font.Size
                .row = 1
                If (iNewLine = 0) Then
                    .CellBackColor = .BackColorFixed
                    .CellFontSize = .Font.Size
'                Else
'                    .CellFontSize = .Font.Size + 2
                End If
                .CellFontBold = True
                j% = j% + 1
                AltJahr& = Jahr&
            End If

            .TextMatrix(0, j%) = Para1.MonatKurz(Monat%)

            iWert! = ArtStat1.MonatsWert(i% + 1)
            If (iWert! = 0) Then
                .TextMatrix(1, j%) = ""
            Else
                .TextMatrix(1, j%) = Str$(iWert!)
            End If
            j% = j% + 1
        Next i%
    Else
        For i% = 0 To .Cols - 1
            .TextMatrix(0, i%) = ""
            .TextMatrix(1, i%) = ""
        Next i%
        
        AltJahr& = -1
        j% = 0
        For i% = 0 To 12
            Termin& = ArtStat1.Anfang - i% - 1
            Jahr& = (Termin& - 1) \ 12
            Monat% = Termin& - Jahr& * 12
            If (AltJahr& <> Jahr&) Then
                .TextMatrix(0, j%) = Str$(Jahr&)
                .row = 0
                .col = j%
                .CellFontBold = True
                .CellFontSize = .Font.Size
                j% = j% + 1
                AltJahr& = Jahr&
            End If

            .TextMatrix(0, j%) = Para1.MonatKurz(Monat%)
            j% = j% + 1
        Next i%
    End If
    Set ArtStat1 = Nothing
    
    .Redraw = True
End With

Call clsError.DefErrPop
End Sub


