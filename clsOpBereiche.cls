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
Dim iArbeitAnzZeilen%
Dim iArbeitTitel%
Dim iArbeitLeerzeileOben%
Dim iArbeitWasDarunter%
Dim iInfoBackHeight%
Dim iInfoAnzZeilen%
Dim iInfoTitel%
Dim iAnzahlButtons%

Dim iBreiteX%
Dim iGesamtY%
Dim iLinksOkX%
Dim iLinksEscX%
Dim iLinksOkEscX%
Dim iZeilenHoeheY%


Dim ParentForm As Form
Dim ParentToolbar As clsToolbar


Private Const DefErrModul = "clsBereiche.cls"

Sub InitBereich(iForm As Object, iToolbar As Object)
Set ParentForm = iForm
Set ParentToolbar = iToolbar
End Sub

Public Property Get ArbeitBackHeight%()
ArbeitBackHeight% = iArbeitBackHeight%
End Property
Public Property Let ArbeitBackHeight(ByVal vNewValue%)
iArbeitBackHeight% = vNewValue%
End Property

Public Property Get ArbeitAnzZeilen%()
ArbeitAnzZeilen% = iArbeitAnzZeilen%
End Property
Public Property Let ArbeitAnzZeilen(ByVal vNewValue%)
iArbeitAnzZeilen% = vNewValue%
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
iInfoTitel% = vNewValue%
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

Sub ResizeWindow()
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
'    .picsave.Width = .Width
'    .picsave.Height = .Height
'    .picsave.Visible = True
'    Suc& = BitBlt(.picsave.hdc, 0, 0, .ScaleWidth, .ScaleHeight, .hdc, 0, 0, &HCC0020)
'    .picsave.ZOrder 0
    Call RefreshBereicheBack
    Call RefreshBereich
    .picBack(0).Visible = True
'    .picsave.Visible = False
End With

End Sub

Sub RefreshBereicheBack()
Dim i%, ToolbarPos%

On Error Resume Next
ToolbarPos% = ParentToolbar.Position

Call ParentToolbar.ResizeToolbar

With ParentForm
    If (ParentToolbar.Visible) Then
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

End Sub

Sub RefreshBereichsKonstanten()

With ParentForm
    iBreiteX% = .picBack(0).ScaleWidth - (300 * wPara1.BildFaktor)
    iGesamtY% = .picBack(0).ScaleHeight
    
    iLinksOkX% = (.picBack(0).Width - (wPara1.ButtonX * 2 + 300)) / 2
    iLinksEscX% = iLinksOkX% + wPara1.ButtonX + 300
    iLinksOkEscX% = (.picBack(0).Width - wPara1.ButtonX) / 2
    
    With .flxarbeit(0)
        If (.Rows < 1) Then .Rows = 1
        iZeilenHoeheY% = .RowHeight(0)
    End With
End With
    
End Sub

Sub RefreshBereich()

Call RefreshBereichsControls
Call RefreshBereichsFarben
With ParentForm
    Call .RefreshBereichsFlexSpalten
    Call .RefreshBereichsControlsAdd
    Call .RefreshBereichsFarbenAdd
End With

End Sub

Sub RefreshBereichsControls()
Dim iAnzButtons%

'On Error Resume Next

With ParentForm
    .lblArbeit(0).Left = wPara1.LinksX
    .flxarbeit(0).Left = wPara1.LinksX
    .lblInfo(0).Left = wPara1.LinksX
    .flxInfo(0).Left = wPara1.LinksX
    
    .lblArbeit(0).Width = iBreiteX%
    .flxarbeit(0).Width = iBreiteX%
    .lblInfo(0).Width = iBreiteX%
    .flxInfo(0).Width = iBreiteX%
    
    
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
    iArbeitAnzZeilen% = (.Height - 90) \ iZeilenHoeheY%
    .Height = iZeilenHoeheY% * iArbeitAnzZeilen% + 90
End With

With ParentForm
    .lblInfo(0).Top = .lblInfo(0).Top + iArbeitBackHeight%
    .flxInfo(0).Top = .flxInfo(0).Top + iArbeitBackHeight%
    
    
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

End Sub

Sub RefreshBereichsFarben()

On Error Resume Next

With ParentForm
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
End With
    
End Sub


