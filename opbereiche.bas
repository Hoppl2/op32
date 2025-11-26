Attribute VB_Name = "modOpBereiche"
Option Explicit

Private Const DefErrModul = "opbereiche.bas"

'Sub InitBereicheBack(iForm As Form, iToolbar As clsToolbar)
'Dim i%, ToolbarPos%
'
'On Error Resume Next
'ToolbarPos% = iToolbar.Position
'
'Call iToolbar.ResizeToolbar
'
'With iForm
'    If (iToolbar.Visible) Then
'        If (ToolbarPos% = 0) Then
'            .picBack(0).Top = .picToolbar.Height
'            .picBack(0).Height = .ScaleHeight - .picToolbar.Height
'            .picBack(0).Left = 0
'            .picBack(0).Width = .ScaleWidth
'        ElseIf (ToolbarPos% = 1) Then
'            .picBack(0).Top = 0
'            .picBack(0).Height = .ScaleHeight
'            .picBack(0).Left = 0
'            .picBack(0).Width = .ScaleWidth - .picToolbar.Width
'        ElseIf (ToolbarPos% = 2) Then
'            .picBack(0).Top = 0
'            .picBack(0).Height = .picToolbar.Top
'            .picBack(0).Left = 0
'            .picBack(0).Width = .ScaleWidth
'        ElseIf (ToolbarPos% = 3) Then
'            .picBack(0).Top = 0
'            .picBack(0).Height = .ScaleHeight
'            .picBack(0).Left = .picToolbar.Width
'            .picBack(0).Width = .ScaleWidth - .picToolbar.Width
'        Else
'            .picBack(0).Top = 0
'            .picBack(0).Height = .ScaleHeight
'            .picBack(0).Left = 0
'            .picBack(0).Width = .ScaleWidth
'        End If
'    Else
'        .picBack(0).Top = 0
'        .picBack(0).Height = .ScaleHeight
'        .picBack(0).Left = 0
'        .picBack(0).Width = .ScaleWidth
'    End If
'    For i% = 1 To UBound(iForm.Bereiche)
'        .picBack(i%).Top = .picBack(0).Top
'        .picBack(i%).Height = .picBack(0).Height
'        .picBack(i%).Left = .picBack(0).Left
'        .picBack(i%).Width = .picBack(0).Width
'    Next i%
'End With
'
'For i% = 0 To UBound(iForm.Bereiche)
'    iForm.Bereiche(i%).BereichOk = False
'Next i%
'
'Call iForm.InitBereichsKonstanten
'
'End Sub

'Sub InitBereichsKonstanten()
'
'With frmAction
'    BreiteX% = .picBack(0).ScaleWidth - (300 * BildFaktor!)
'    GesamtY% = .picBack(0).ScaleHeight
'
'    LinksOkX% = (.picBack(1).Width - (ButtonX% * 2 + 300)) / 2
'    LinksEscX% = LinksOkX% + ButtonX% + 300
'    LinksOkEscX% = (.picBack(1).Width - ButtonX%) / 2
'End With
'
'With frmAction.flxarbeit(0)
'    If (.Rows < 1) Then .Rows = 1
'    ZeilenHoeheY% = .RowHeight(0)
'End With
'
'End Sub

'Sub InitBereich(iForm As Form, iBereich As BereicheStruct, iBereichInd%)
'
'Call InitBereichsControls(iForm, iBereich, iBereichInd%)
'Call iForm.InitBereichsFlexSpalten(iBereichInd%)
'Call iForm.InitBereichsControlsAdd(iBereichInd%)
'Call InitBereichsFarben(iForm, iBereich, iBereichInd%)
'
'iBereich.BereichOk = True
'
'End Sub
'
'Sub InitBereichsControls(iForm As Form, iBereich As BereicheStruct, i%)
'Dim iAnzInfoZeilen%, iAnzButtons%
'
''On Error Resume Next
'
'With iForm
'    .lblArbeit(i%).Left = clsWinPara1.LinksX
'    .flxarbeit(i%).Left = clsWinPara1.LinksX
'    .lblInfo(i%).Left = clsWinPara1.LinksX
'    .flxInfo(i%).Left = clsWinPara1.LinksX
'
'    .lblArbeit(i%).Width = iForm.BreiteX%
'    .flxarbeit(i%).Width = iForm.BreiteX%
'    .lblInfo(i%).Width = iForm.BreiteX%
'    .flxInfo(i%).Width = iForm.BreiteX%
'
'
'    .lblArbeit(i%).Top = clsWinPara1.TitelY
'    If (iBereich.ArbeitTitelAnzeigen) Then
'        .lblArbeit(i%).Visible = True
'        .flxarbeit(i%).Top = clsWinPara1.FlexY
'    Else
'        .lblArbeit(i%).Visible = False
'        .flxarbeit(i%).Top = clsWinPara1.TitelY
'    End If
'    If (iBereich.ArbeitLeerzeileObenAnzeigen) Then
'        .flxarbeit(i%).Top = .flxarbeit(i%).Top + clsWinPara1.FlexY
'    End If
'
'    .lblInfo(i%).Top = clsWinPara1.TitelY
'    If (iBereich.InfoTitelAnzeigen) Then
'        .lblInfo(i%).Visible = True
'        .flxInfo(i%).Top = clsWinPara1.FlexY
'    Else
'        .lblInfo(i%).Visible = False
'        .flxInfo(i%).Top = clsWinPara1.TitelY
'    End If
'End With
'
'
'iAnzInfoZeilen% = iBereich.InfoAnzZeilen
'With iForm.flxInfo(i%)
'    .Height = iForm.ZeilenHoeheY% * iAnzInfoZeilen% + 90
'    If (iAnzInfoZeilen% > 0) Then
'        iBereich.InfoBackHeight = .Top + .Height + clsWinPara1.UntenFreiY
'    Else
'        iBereich.InfoBackHeight = clsWinPara1.OhneInfoY
'        iForm.lblInfo(i%).Caption = "Keine zusätzlichen Informationen vorhanden"
'        .Visible = False
'    End If
'End With
'
'iBereich.ArbeitBackHeight = iForm.GesamtY% - iBereich.InfoBackHeight
'With iForm.flxarbeit(i%)
'    If (i% = 3) Then
'        iBereich.ArbeitAnzZeilen = iBereich.InfoAnzZeilen
'    Else
'        .Height = iBereich.ArbeitBackHeight - clsWinPara1.UntenFreiY - .Top
'        If (iBereich.ArbeitWasDarunterAnzeigen) Then
'            .Height = .Height - iForm.ZeilenHoeheY% - 90
'        End If
'        iBereich.ArbeitAnzZeilen = (.Height - 90) \ iForm.ZeilenHoeheY%
'    End If
'    .Height = iForm.ZeilenHoeheY% * iBereich.ArbeitAnzZeilen + 90
'End With
'
'With iForm
'    .lblInfo(i%).Top = .lblInfo(i%).Top + iBereich.ArbeitBackHeight
'    .flxInfo(i%).Top = .flxInfo(i%).Top + iBereich.ArbeitBackHeight
'
'
'    .cmdOk(i%).Width = clsWinPara1.ButtonX
'    .cmdOk(i%).Height = clsWinPara1.ButtonY
'    .cmdEsc(i%).Width = clsWinPara1.ButtonX
'    .cmdEsc(i%).Height = clsWinPara1.ButtonY
'
'    iAnzButtons = Abs(iBereich.AnzahlButtons)
'    If (iAnzButtons% = 1) Then
'        .cmdOk(i%).Left = iForm.LinksOkEscX%
'        .cmdEsc(i%).Left = iForm.LinksOkEscX%
'    Else
'        .cmdOk(i%).Left = iForm.LinksOkX%
'        .cmdEsc(i%).Left = iForm.LinksEscX%
'    End If
'
'    If (iBereich.AnzahlButtons > 0) Then
'        .cmdOk(i%).Top = iBereich.ArbeitBackHeight - .cmdOk(i%).Height - 60
'    Else
'        .cmdOk(i%).Top = .picBack(i%).Height + 90
'    End If
'    .cmdEsc(i%).Top = .cmdOk(i%).Top
'End With
'
'End Sub
'
'Sub InitBereichsFarben(iForm As Form, iBereich As BereicheStruct, iBereichInd%)
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Call DefErrFnc("InitBereichsFarben")
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
'Dim i%
'
'On Error Resume Next
'
'With iForm
'    .picBack(iBereichInd%).AutoRedraw = True
'
'    .picBack(iBereichInd%).Line (0, 0)-(.ScaleWidth, iBereich.ArbeitBackHeight), clsWinPara1.FarbeArbeit, BF
'    .picBack(iBereichInd%).Line (0, iBereich.ArbeitBackHeight)-(.ScaleWidth, .ScaleHeight), clsWinPara1.FarbeInfo, BF
'
'    If (clsWinPara1.FarbeRahmen) Then
'        .picBack(iBereichInd%).DrawWidth = 3
'        .picBack(iBereichInd%).Line (0, 0)-(.ScaleWidth, iBereich.ArbeitBackHeight), vbBlack, B
'        .picBack(iBereichInd%).Line (0, iBereich.ArbeitBackHeight)-(.ScaleWidth, .ScaleHeight), vbBlack, B
'    End If
'
'    .lblArbeit(iBereichInd%).BackColor = clsWinPara1.FarbeArbeit
'    .flxarbeit(iBereichInd%).BackColorBkg = clsWinPara1.FarbeArbeit
'
'    .lblInfo(iBereichInd%).BackColor = clsWinPara1.FarbeInfo
'    .flxInfo(iBereichInd%).BackColorBkg = clsWinPara1.FarbeInfo
'End With
'
'Call iForm.InitBereichsFarbenAdd(iBereichInd%)
'
'Call DefErrPop
'End Sub
'
