VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsBesorger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Const INI_SECTION = "Besorger"
Const INFO_SECTION = "Infobereich Besorger"

Dim ParentForm As Form
Dim ParentToolbar As clsToolbar
Dim ParentInfo As clsInfoBereich
Dim ParentBereich As clsOpBereiche

Dim WamaBesorger$

Private Const DefErrModul = "BESORGER.CLS"

Sub Init(iForm As Object, iToolbar As Object, iInfo As Object, iBereich As Object)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("Init")
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
Dim h$

Set ParentForm = iForm
Set ParentToolbar = iToolbar
Set ParentInfo = iInfo
Set ParentBereich = iBereich

With ParentForm
    Call ParentToolbar.InitToolbar(ParentForm, INI_DATEI, INI_SECTION, "180218041808180918111812181418151816")
    
    .cmdToolbar(0).ToolTipText = "ESC Zurück: Zurückschalten auf vorige Bildschirmmaske"
    .cmdToolbar(1).ToolTipText = "F2 "
    .cmdToolbar(2).ToolTipText = "F3 Interner Programmwechsel"
    .cmdToolbar(3).ToolTipText = "F4 "
    .cmdToolbar(4).ToolTipText = "F5 Besorger auf abgeholt setzen/rücksetzen"
    .cmdToolbar(5).ToolTipText = "F6 Drucken der Besorger"
    .cmdToolbar(6).ToolTipText = "F7 Aktualisieren"
    .cmdToolbar(7).ToolTipText = "F8 "
    .cmdToolbar(8).ToolTipText = "F9 "
    .cmdToolbar(9).ToolTipText = "shift+F2 Artikel-Status"
    .cmdToolbar(10).ToolTipText = "shift+F3 "
    .cmdToolbar(11).ToolTipText = "shift+F4 "
    .cmdToolbar(12).ToolTipText = "shift+F5 Durchgriff auf Statistik-Anzeige"
    .cmdToolbar(13).ToolTipText = "shift+F6 "
    .cmdToolbar(14).ToolTipText = "shift+F7 "
    .cmdToolbar(15).ToolTipText = "shift+F8 "
    .cmdToolbar(16).ToolTipText = "shift+F9 Abholer komprimieren"
    'cmdToolbar(19).ToolTipText = "Programm beenden"
    
    Call ParentInfo.InitInfoBereich(.flxInfo(0), INI_DATEI, INFO_SECTION)
    Call ParentInfo.ZeigeInfoBereich("", False)
    .flxInfo(0).row = 0
    .flxInfo(0).col = 0
    
    Call ParentBereich.InitBereich(iForm, iToolbar)
    ParentBereich.ArbeitTitel = False
    ParentBereich.ArbeitLeerzeileOben = False
    ParentBereich.ArbeitWasDarunter = False
    ParentBereich.InfoTitel = False
    ParentBereich.InfoZusatz = ArtikelStatistik%
    ParentBereich.InfoAnzZeilen = iInfo.AnzInfoZeilen
    ParentBereich.AnzahlButtons = 0
    
    With .flxarbeit(0)
        .Cols = 16
        .Rows = 0
        .Rows = 2
        .FixedRows = 1
        .FormatString = "<PZN|^ |<Name|>Menge|^Meh|>BM|>Preis|>Abhol#|>Datum|>Uhrzeit|WAMA|^Rez|^LS|||"
    End With
        
    .mnuBearbeitenInd(0).Caption = ""
    .mnuBearbeitenInd(2).Caption = ""
    .mnuBearbeitenInd(3).Caption = "Abgeholt/Rücksetzen"
    .mnuBearbeitenInd(6).Caption = ""
    .mnuBearbeitenInd(7).Caption = ""
    .mnuBearbeitenInd(10).Caption = ""
    .mnuBearbeitenInd(11).Caption = ""
    .mnuBearbeitenInd(13).Caption = ""
    .mnuBearbeitenInd(14).Caption = ""
    .mnuBearbeitenInd(15).Caption = ""
    .mnuBearbeitenInd(16).Caption = "Abholer komprimieren"
    
    .mnuRahmenAnzeigen.Checked = wpara.FarbeRahmen
'    .mnuAktZeile.Checked = wpara.FarbeAktZeile
    .mnuZusatzInfo.Checked = ArtikelStatistik%
    
    .mnuEinzelBewertung.Enabled = False
    
    .mnuRueckrufe.Enabled = False
    .mnuLiefStammdaten.Enabled = False
    .mnuPosPartner.Enabled = False
    .mnuAepKalkulation.Enabled = False
    .mnuTeilAbschluss.Enabled = False

    .mnuEditAufteilung.Enabled = False
    
    .picBestellWerte.Visible = False
End With

BestellAnzeige% = 0

Call DefErrPop
End Sub

Sub EchtKurzInfo()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("EchtKurzInfo")
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
Dim i%, row%, iRow%, iCol%, KommWeg%, iDummy%
Dim dDummy#
Dim pzn$, ch$, ActKontrollen$, actzuordnung$, h$

With ParentForm
    row% = .flxarbeit(0).row
    pzn$ = RTrim$(.flxarbeit(0).TextMatrix(row%, 0))
    If (pzn$ = "XXXXXXX") Then Call DefErrPop: Exit Sub
    pzn$ = KorrPzn$(pzn$)
    
    iRow% = .flxInfo(0).row
    iCol% = .flxInfo(0).col
    Call ParentInfo.ZeigeInfoBereich(pzn$, True)
    If (.ActiveControl.Name <> .flxInfo(0).Name) Then
        Call ZeigeInfoBereichAdd
    End If
    .flxInfo(0).row = iRow%
    .flxInfo(0).col = iCol%
    
    If (ParentBereich.InfoZusatz) Then
        Call .ZeigeInfoZusatz(pzn$)
    End If
    
    FabsErrf% = ass.IndexSearch(0, pzn$, FabsRecno&)
    If (FabsErrf% = 0) Then
        .mnuBearbeitenInd(MENU_SF5).Enabled = True
    Else
        .mnuBearbeitenInd(MENU_SF5).Enabled = False
    End If
    .cmdToolbar(MENU_SF5).Enabled = .mnuBearbeitenInd(MENU_SF5).Enabled
    
            
    .picQuittieren.Visible = False
    
    .flxEinzelBewertung.Visible = False
    .flxDirektAufteilung.Visible = False
End With

Call DefErrPop
End Sub

Sub ZeigeInfoBereichAdd()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("ZeigeInfoBereichAdd")
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
Dim j%, Send%
Dim h$, tx$

With ParentForm.flxInfo(0)
    .redraw = False

    j% = 1
    
    Do While (j% <= ParentInfo.AnzInfoZeilen%)
        .TextMatrix(j% - 1, 0) = ""
        j% = j% + 1
    Loop
    
    .redraw = True
End With

Call DefErrPop
End Sub

Sub RefreshBereichsFlexSpalten()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("RefreshBereichsFlexSpalten")
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
Dim i%, j%, spBreite%
Dim sp&
            
With ParentForm.flxarbeit(0)
    ParentForm.Font.Bold = True
    .ColWidth(0) = 0
    .ColWidth(1) = ParentForm.TextWidth("X")
    .ColWidth(2) = 0
    .ColWidth(3) = ParentForm.TextWidth("XXXXXX")
    .ColWidth(4) = ParentForm.TextWidth("XXX")
    .ColWidth(5) = ParentForm.TextWidth("9999")
    .ColWidth(6) = ParentForm.TextWidth("99999.99")
    .ColWidth(7) = ParentForm.TextWidth("999999")
    .ColWidth(8) = ParentForm.TextWidth("99.99.9999")
    .ColWidth(9) = ParentForm.TextWidth("99:9999")
    .ColWidth(10) = ParentForm.TextWidth("XXXXXX")
    .ColWidth(11) = ParentForm.TextWidth("XXX")
    .ColWidth(12) = ParentForm.TextWidth("XX")
    .ColWidth(13) = 0
    .ColWidth(14) = 0
    .ColWidth(15) = wpara.FrmScrollHeight
    ParentForm.Font.Bold = False
    
    
    spBreite% = 0
    For i% = 1 To .Cols - 1
        If (.ColWidth(i%) > 0) Then
            .ColWidth(i%) = .ColWidth(i%) + ParentForm.TextWidth("X")
        End If
        spBreite% = spBreite% + .ColWidth(i%)
    Next i%
    If (spBreite% > .Width) Then
        spBreite% = .Width
    End If
    .ColWidth(2) = .Width - spBreite%
End With

With ParentForm.flxInfo(0)
    sp& = .Width / 8
    .ColWidth(0) = 2 * sp&
    For i% = 1 To 6
        .ColWidth(i%) = sp&
    Next i%
End With
        
With ParentForm.flxInfoZusatz(0)
    .Cols = 15
    sp& = .Width / 15
    For i% = 0 To 14
        .ColWidth(i%) = sp&
    Next i%
End With
        
Call DefErrPop
End Sub

Sub EditSatz()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("EditSatz")
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
Dim i%, ind%, lInd%, rInd%, EditCol%, aRow%, m%, aMeng%, iLief%, IstBesorger%
Dim aPreis#, nPreis#
Dim h$, h2$
Dim iFlex As MSFlexGrid
            
            
Set iFlex = ParentForm.flxarbeit(0)

EditCol% = iFlex.col
If (EditCol% = 6) And (iFlex.TextMatrix(iFlex.row, 11) <> "") Then
    Call iMsgBox("Achtung: Für Rezept-Besorger Preiseingabe unzulässig !", vbInformation, "Besorger-Verwaltung")
    Call DefErrPop: Exit Sub
End If

If (EditCol% = 5) Or (EditCol% = 6) Then

    EditModus% = 0
    If (EditCol% = 6) Then EditModus% = 4
                
    With iFlex
        KeinRowColChange% = True
        aRow% = .row
        .row = 0
        .CellFontBold = True
        .row = aRow%
        KeinRowColChange% = False
    End With
                
    Load frmEdit
    With frmEdit
        .Left = ParentForm.picBack(0).Left + iFlex.Left + iFlex.ColPos(EditCol%) + 45
        .Left = .Left + ParentForm.Left + wpara.FrmBorderHeight
        .Top = ParentForm.picBack(0).Top + iFlex.Top + (iFlex.row - iFlex.TopRow + 1) * iFlex.RowHeight(0)
        .Top = .Top + ParentForm.Top + wpara.FrmBorderHeight + wpara.FrmCaptionHeight + wpara.FrmMenuHeight
        .Width = iFlex.ColWidth(EditCol%)
        .Height = frmEdit.txtEdit.Height 'flxarbeit(0).RowHeight(1)
    End With
    With frmEdit.txtEdit
        .Width = frmEdit.ScaleWidth
        .Left = 0
        .Top = 0
        h2$ = iFlex.TextMatrix(iFlex.row, EditCol%)
        .text = h2$
        .BackColor = vbWhite
        .Visible = True
    End With
                
    frmEdit.Show 1
                
    With iFlex
        KeinRowColChange% = True
        aRow% = .row
        .row = 0
        .CellFontBold = False
        .row = aRow%
        KeinRowColChange% = False
    End With
    
    If (EditErg%) Then
        If (EditCol% = 5) Then
            aMeng% = Val(iFlex.TextMatrix(iFlex.row, EditCol%))
            m% = Val(EditTxt$)
            If (aMeng% <> m%) Then
                Call BesorgerFertig(1, m%)
            End If
        Else
            aPreis# = Val(iFlex.TextMatrix(iFlex.row, EditCol%))
            nPreis# = Val(EditTxt$)
            If (aPreis# <> nPreis#) Then
                Call BesorgerFertig(2, 0, nPreis#)
            End If
        End If
    End If
Else
    Call ZeigeBesorgerDaten
End If
'Call NaechsteBesorgerZeile

Call DefErrPop
End Sub
      
Sub BesorgerFertig(bStatus%, Optional bMenge% = 0, Optional bPreis# = 0#, Optional bPzn$ = "")
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("BesorgerFertig")
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
'Dim i%, erg%, Belegt%, KisteFertig%, KisteStatus%, BesorgerStatus%, s%, NurFertigeBesorger%, kInd%
'Dim h$, Arbeit$

Dim i%, j%, erg%, row%, KistenNr%, KistenInd%, KistenStatus%, BesorgerStatus%, BesorgerMenge%
Dim h$, BesorgerPzn$, Arbeit$

With ParentForm.flxarbeit(0)
    row% = .row
    h$ = RTrim$(.TextMatrix(row%, 7))
    KistenNr% = Val(h$)
    h$ = RTrim$(.TextMatrix(row%, 14))
    KistenInd% = Val(h$)
End With

KeinRowColChange% = True

kiste.OpenDatei
If (kiste.Belegt(KistenNr%)) Then
    kiste.GetKiste (KistenNr%)
    
    erg% = kiste.GetInhalt(KistenInd%)
    If (kiste.pzn = KorrPzn$) Then
        If (bStatus% And 1) Then
            kiste.menge = bMenge%
        End If
        If (bStatus% And 2) Then
            kiste.Preis = bPreis#
        End If
        If (bStatus% And 4) And (kiste.status <> 4) Then
            If (kiste.status = 3) Then
                kiste.status = 1
            Else
                kiste.status = 3
            End If
        End If
'        If (bStatus% And 8) Then
'            Kiste1.pzn = bPzn$
'        End If
        If (bStatus% And 16) Then
            If (kiste.status = 4) Then
                kiste.status = 1
            Else
                kiste.status = 4
            End If
        End If
        kiste.PutInhalt (KistenInd%)
    End If
    
    If (bStatus% And (4 Or 16)) Then
        KistenStatus% = 4
        BesorgerStatus% = 4

        For i% = 0 To 9
            erg% = kiste.GetInhalt(i%)
            If (erg%) Then
                If (kiste.WasTun = "B") Then
                    If (BesorgerStatus% > 3) And (kiste.status < 4) Then BesorgerStatus% = 3
                    If (kiste.status < 3) And (BesorgerStatus% > 1) Then BesorgerStatus% = 1
                End If
                If (InStr("RTBHP", kiste.WasTun) > 0) Then
                    If (KistenStatus% > 3) And (kiste.status < 4) Then KistenStatus% = 3
                    If (kiste.status < 3) And (KistenStatus% > 1) Then KistenStatus% = 1
                End If
            End If
        Next i%

        Arbeit$ = kiste.Arbeit
        Mid$(Arbeit$, 4, 1) = Chr$(BesorgerStatus%)
        kiste.Arbeit = Arbeit$
        kiste.KistenStatus = KistenStatus%
        kiste.PutKiste (KistenNr%)
    End If
End If
    

BesorgerMenge% = 1
For j% = 1 To 10
    h$ = RTrim$(kiste.InfoText(j% - 1))
    If Mid$(h$, 18, 3) = " x " Then
        BesorgerMenge% = -Val(Mid$(h$, 21, 4))
    End If
    If Mid$(h$, 14, 4) = "PZN=" Then
        BesorgerPzn$ = Mid$(h$, 18, 7)
    End If
Next j%

Call ZeigeBesorgerZeile(KistenNr%, KistenInd%, BesorgerPzn$, BesorgerMenge%)

kiste.CloseDatei

KeinRowColChange% = False

Call DefErrPop
End Sub

Sub flxInfoGotFocus(InfoRow%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("flxInfoGotFocus")
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

Call DefErrPop
End Sub

Public Sub ZeigeWerte(Optional AnzeigeFlag% = True)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("ZeigeWerte")
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

Call DefErrPop
End Sub
   
Sub mnuFontClick()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("mnuFontClick")
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

Call AuslesenBesorger

Call DefErrPop
End Sub

Sub MenuBearbeiten(Index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("MenuBearbeiten")
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

Select Case Index
    Case MENU_F5
        Call BesorgerFertig(16)
        Call NaechsteBesorgerZeile
    
    Case MENU_F7
        Call AuslesenBesorger
    
    Case MENU_SF2
        Call AnzeigeFenster(TEXT_BESTELLSTATUS, KorrPzn$, KorrTxt$)

    Case MENU_SF5
        ZeigeStatistik
        
    Case MENU_SF9
        Call AbholerKomprimieren
End Select

Call DefErrPop
End Sub

Sub ToggleBesorgerZeile()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("ToggleBesorgerZeile")
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

Call BesorgerFertig(4)

Call DefErrPop
End Sub

Sub NaechsteBesorgerZeile()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("NaechsteBesorgerZeile")
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

With ParentForm.flxarbeit(0)
    If (ParentForm.picQuittieren.Visible = False) Then
        If (.row < .Rows - 1) Then
            .row = .row + 1
            If (.TopRow + ParentBereich.ArbeitAnzZeilen - 1 <= .row) Then
                .TopRow = .row
            End If
        End If
    End If
End With

Call DefErrPop
End Sub

Sub flxArbeitKeyPress(KeyAscii As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("flxArbeitKeyPress")
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
Dim i%, ind%, ind2%, ind3%

If (KeyAscii = vbKeySpace) Then
    Call ToggleBesorgerZeile
    Call NaechsteBesorgerZeile
ElseIf (InStr("ABCDEFGHIJKLMNOPQRSTUVWXYZ", UCase(Chr$(KeyAscii))) > 0) Then
    Call frmAction.SelectZeile(UCase(Chr$(KeyAscii)))
End If

Call DefErrPop
End Sub

Public Sub cmdOkClick(h$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("cmdOkClick")
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

Call DefErrPop
End Sub

Sub ZeigeBesorgerDaten()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("ZeigeBesorgerDaten")
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
Dim j%, erg%, row%, KistenNr%, KistenInd%, BesorgerStatus%, BesorgerMenge%
Dim h$, BesorgerPzn$

With ParentForm.flxarbeit(0)
    row% = .row
    h$ = RTrim$(.TextMatrix(row%, 7))
    KistenNr% = Val(h$)
    h$ = RTrim$(.TextMatrix(row%, 14))
    KistenInd% = Val(h$)
End With

Call BesorgerInfo(KistenNr%, KorrPzn$, KorrTxt$)

KeinRowColChange% = True

kiste.OpenDatei
If (kiste.Belegt(KistenNr%)) Then
    kiste.GetKiste (KistenNr%)
    
    erg% = kiste.GetInhalt(KistenInd%)
    If (erg%) Then
        If (kiste.WasTun = "B") Then
            BesorgerStatus% = kiste.status
            BesorgerMenge% = 1
            For j% = 1 To 10
                h$ = RTrim$(kiste.InfoText(j% - 1))
                If Mid$(h$, 18, 3) = " x " Then
                    BesorgerMenge% = -Val(Mid$(h$, 21, 4))
                End If
                If Mid$(h$, 14, 4) = "PZN=" Then
                    BesorgerPzn$ = Mid$(h$, 18, 7)
                End If
            Next j%
            
            Call ZeigeBesorgerZeile(KistenNr%, KistenInd%, BesorgerPzn$, BesorgerMenge%)
        End If
    End If
End If
kiste.CloseDatei

With ParentForm.flxarbeit(0)
    .col = 2
    .ColSel = .col
End With

KeinRowColChange% = False

Call DefErrPop
End Sub

Sub mnuOptionenClick()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("mnuOptionenClick")
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

Call DefErrPop
End Sub

Sub mnuBearbeitenZusatzClick(Index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("mnuBearbeitenZusatzClick")
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

Call DefErrPop
End Sub

Sub ZeigeBesorgerZeile(KistenNr%, KistenInd%, pzn$, bm%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("ZeigeBesorgerZeile")
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
Dim i%, j%, Lief%, l%, row%, col%, iBold%, iItalic%, FirstRed%, xAktiv%, ind%
Dim BackDunkel%
Dim lBack&
Dim EK#, Rabatt!, bmo!
Dim h$, h2$, s$, LiefName$, ArtName$, ArtMenge$, ArtMeh$, VonWo$, nm$, zusatz$, KontrollChar$
Dim ActKontrollen$, actzuordnung$, SRT$, SQLStr$, ZusInfo$
Dim ArtikelKz As Byte
Dim BesorgerTxt$, WamaTxt$

WamaTxt$ = ""
h$ = pzn$ + Format(KistenNr%, "000") + "-"
ind% = InStr(WamaBesorger$, h$)
If (ind% > 0) Then
    h2$ = Mid$(WamaBesorger$, ind% + 11, 1)
    If (h2$ = "1") Then
        WamaTxt$ = "Best"
    Else
        WamaTxt$ = "WÜ"
    End If
End If

BesorgerTxt$ = ""

SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + pzn$
Set TaxeRec = TaxeDB.OpenRecordset(SQLStr$)
If (TaxeRec.EOF = False) Then
    Call Taxe2ast(pzn$)
    BesorgerTxt$ = ast.kurz + ast.meng + ast.meh
Else
    BesorgerTxt$ = Trim(kiste.InfoText(0))
End If


If (BestellAnzeige% = 0) Then
    h$ = Format(kiste.status, "0")
    h$ = h$ + CVDatum(Left$(kiste.VonWann, 2))
    h$ = h$ + Format(Asc(Mid$(kiste.VonWann, 3, 1)), "00") + Format(Asc(Mid$(kiste.VonWann, 4, 1)), "00")
    h$ = h$ + Format(KistenInd%, "0")
ElseIf (BestellAnzeige% = 1) Then
    h$ = Format(KistenNr%, "000") + Format(KistenInd%, "0")
Else
    h$ = BesorgerTxt$
    h$ = h$ + Format(KistenNr%, "000") + Format(KistenInd%, "0")
End If
SRT$ = h$

ArtName$ = BesorgerTxt$
ArtMenge$ = ""
ArtMeh$ = ""
    
If (TaxeRec.EOF = False) Then
    h$ = RTrim(Left$(BesorgerTxt$, 33))
    Call OemToChar(h$, h$)
    l% = Len(h$)
    xAktiv% = False
    For j% = l% To 2 Step -1
        h2$ = Mid$(h$, j%, 1)
        If (h2$ = " ") Then Exit For
        If (xAktiv%) And (InStr("0123456789", h2$) <= 0) Then Exit For
        If (InStr("xX", h2$) > 0) Then xAktiv% = True
    Next j%
    If (j% > 2) Then
        ArtName$ = Left$(h$, j%)
        ArtMenge$ = Mid$(h$, j% + 1)
    Else
        ArtName$ = h$
        ArtMenge$ = ""
    End If
    ArtMeh$ = Mid$(BesorgerTxt$, 34)
End If

KontrollChar$ = " "
If (kiste.status >= 3) Then
    KontrollChar$ = Chr$(214)
End If

'FirstRed% = False
'If (ww.fixiert = "2") Then
'    KontrollChar$ = Chr$(214)
'ElseIf (ActKontrollen$ <> "") Then
'    KontrollChar$ = "?"
'    If (ww.zukontrollieren = "1") Then
'        FirstRed% = True
'    End If
'End If

    
With frmAction.flxarbeit(0)

    col% = .col

    iItalic% = False
    
    lBack& = vbButtonFace
    
    BackDunkel% = 0
    If (kiste.status = 4) Then
        BackDunkel% = True
    End If

    If (BackDunkel%) Then lBack& = wpara.FarbeDunklerBereich ' FarbeGray&        'vbGrayText

    .FillStyle = flexFillRepeat
    .col = 0
    .ColSel = .Cols - 1
    .CellFontItalic = iItalic%
    .CellBackColor = lBack&
    .FillStyle = flexFillSingle

'    'damit Berechnung der Werte richtig
'    If (ww.zugeordnet = "N") Then
'        lBack& = wpara.FarbeDunklerBereich  ' FarbeGray& ' vbGrayText
'    End If

    .col = 1
    .CellFontName = "Symbol"
    If (FirstRed%) Then
        .CellBackColor = vbRed
    End If

    .col = 12
    .CellFontBold = True

    row% = .row
    .TextMatrix(row%, 0) = pzn$
    .TextMatrix(row%, 1) = KontrollChar$
    .TextMatrix(row%, 2) = ArtName$
    .TextMatrix(row%, 3) = ArtMenge$
    .TextMatrix(row%, 4) = ArtMeh$
    .TextMatrix(row%, 5) = Abs(bm%)
    .TextMatrix(row%, 6) = Format(kiste.Preis, "0.00")
    .TextMatrix(row%, 7) = Format(KistenNr%, "0")
    
    h$ = CVDatum(Left$(kiste.VonWann, 2))
    .TextMatrix(row%, 8) = Mid$(h$, 7, 2) + "." + Mid$(h$, 5, 2) + "." + Left$(h$, 4)
    
    h$ = kiste.VonWann
    .TextMatrix(row%, 9) = Format(Asc(Mid$(h$, 3, 1)), "00") + ":" + Format(Asc(Mid$(h$, 4, 1)), "00")
    
    .TextMatrix(row%, 10) = WamaTxt$
    
    h$ = ""
    If (kiste.RezeptNr > 0) Then
        h$ = "!"
    End If
    .TextMatrix(row%, 11) = h$
    
    h$ = ""
    If (kiste.HatLieferSchein) Then
        h$ = "!"
    End If
    .TextMatrix(row%, 12) = h$
    
    .TextMatrix(row%, 13) = SRT$
    .TextMatrix(row%, 14) = Format(KistenInd%, "0")
    
    .col = col%
End With

Call DefErrPop
End Sub

Sub AuslesenBesorger()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("AuslesenBesorger")
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
Dim i%, j%, erg%, KistenNr%, BesorgerStatus%, BesorgerMenge%, AnzBesorger%, MitBesorgerWeg%
Dim NurFertigeBesorger%, s%
Dim l&
Dim BesorgerPzn$, h$, Arbeit$
Dim tdatum As Date
Static IstAktiv%

If (IstAktiv%) Then Call DefErrPop: Exit Sub

IstAktiv% = True

h$ = ProgrammNamen$(ProgrammTyp%) + " - Sortierung nach "
If (BestellAnzeige% = 0) Then
    h$ = h$ + "Gruppen/Datum"
ElseIf (BestellAnzeige% = 1) Then
    h$ = h$ + "Abhol-Nummern"
Else
    h$ = h$ + "Artikel"
End If
frmAction.Caption = h$

Call StartAnimation(frmAction, "Besorger werden eingelesen ...")

KeinRowColChange% = True

With frmAction.flxarbeit(0)
    .redraw = False
    If (.Rows < 2) Then
        .Rows = 2
    End If
    .row = 0
End With

AnzBesorger% = 0

Call EinlesenWaMaBesorger

tdatum = DateValue(Left(LetztBesorgerWeg$, 2) + "." + Mid(LetztBesorgerWeg$, 3, 2) + "." + Mid(LetztBesorgerWeg$, 5, 2))
MitBesorgerWeg% = (DateDiff("D", tdatum, Now) > 0)
If (MitBesorgerWeg%) Then
    h$ = Format(Now, "DDMMYY")
    LetztBesorgerWeg$ = h$
    l& = WritePrivateProfileString("Bestellung", "LetztBesorgerWeg", h$, INI_DATEI)
End If

kiste.OpenDatei

For KistenNr% = 1 To 999
    If (kiste.Belegt(KistenNr%)) Then
        kiste.GetKiste (KistenNr%)
        
                
        NurFertigeBesorger% = 0
        If (MitBesorgerWeg%) And (kiste.KistenStatus = 4) Then
           Arbeit$ = kiste.Arbeit
           '* alles abgeholt
           '* wenn das nur Besorger waren, dann entfernen
           NurFertigeBesorger% = True
           For i% = 1 To 5
               s% = Asc(Mid$(Arbeit$, i%, 1))
               If (i% <> 4) And (s% <> 0) Then
                   NurFertigeBesorger% = 0
               End If
           Next i%
           If (NurFertigeBesorger%) Then
               '* AnfBlocks ausleeren
               For i% = 0 To 9
                   erg% = kiste.ClearInhalt(i%)
               Next i%
        
               kiste.ClearKiste (KistenNr%)
           End If
        End If
       
        If (NurFertigeBesorger%) Then
        Else
            For i% = 0 To 9
                erg% = kiste.GetInhalt(i%)
                If (erg%) Then
                    If (kiste.WasTun = "B") Then
                        BesorgerStatus% = kiste.status
                        BesorgerMenge% = 1
                        For j% = 1 To 10
                            h$ = RTrim$(kiste.InfoText(j% - 1))
                            If Mid$(h$, 18, 3) = " x " Then
                                BesorgerMenge% = -Val(Mid$(h$, 21, 4))
                            End If
                            If Mid$(h$, 14, 4) = "PZN=" Then
                                BesorgerPzn$ = Mid$(h$, 18, 7)
                            End If
                        Next j%
                        
                        With frmAction.flxarbeit(0)
                            AnzBesorger% = AnzBesorger% + 1
                    
                            If (.row >= (.Rows - 1)) Then .Rows = .Rows + 1
                            .row = .row + 1
                            Call ZeigeBesorgerZeile(KistenNr%, i%, BesorgerPzn$, BesorgerMenge%)
                        End With
                    
                    End If
                End If
            Next i%
        End If
    End If
Next KistenNr%

kiste.CloseDatei


With frmAction.flxarbeit(0)
    If (.row = 0) Then
        .Rows = 2
        .TextMatrix(1, 0) = "XXXXXXX"
        .TextMatrix(1, 1) = " "
        .TextMatrix(1, 2) = "Keine anzuzeigenden Besorger gespeichert !"
        For i% = 3 To (.Cols - 1)
            .TextMatrix(1, i%) = ""
        Next i%
    Else
        .Rows = .row + 1
    End If
    .row = 1
    .col = 13
    .RowSel = .Rows - 1 ' AnzBestellArtikel%
    .ColSel = .col
    .Sort = 5
    .redraw = True
    .TopRow = 1
    .row = 1
    .col = 2
    .ColSel = .col
    .SetFocus
End With

KeinRowColChange% = False
    
Call StopAnimation(frmAction)

Call ZeigeWerte

IstAktiv% = False

Call DefErrPop
End Sub

Sub RefreshDunklerBereich()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("RefreshDunklerBereich")
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

Call MenuBearbeiten(MENU_F7)

Call DefErrPop
End Sub
                  
Sub EinlesenWaMaBesorger()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("EinlesenWaMaBesorger")
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
Dim i%, j1%, j2%, j3%, k%, Max%, gAlm%, gefunden%
Dim pzn$, gPzn$, ret$

'ww.SatzLock (1)
Call ww.GetRecord(1)
Max% = ww.erstmax

ret$ = ""

For i% = 1 To Max%
    ww.GetRecord (i% + 1)
    If (ww.status > 0) And (ww.auto = "v") Then
        ret$ = ret$ + ww.pzn + Format(ww.AbholNr, "000") + "-" + Format(ww.status, "0") + ","
    End If
Next i%

'ww.SatzUnLock (1)
WamaBesorger$ = ret$

Call DefErrPop
End Sub

Sub AbholerKomprimieren()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("AbholerKomprimieren")
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

If (iMsgBox("Abholer komprimieren ?", vbYesNo Or vbDefaultButton2, "Besorger-Verwaltung") = vbYes) Then
    Call StartAnimation(frmAction, "Abholer werden komprimiert ...")
    Call kiste.KistenReorg
    Call StopAnimation(frmAction)
    Call AuslesenBesorger
End If

Call DefErrPop
End Sub


