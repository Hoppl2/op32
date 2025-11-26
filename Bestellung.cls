VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsBestellung"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Const INI_SECTION = "Bestellung"
Const INFO_SECTION = "Infobereich Bestellung"

Dim ParentForm As Form
Dim ParentToolbar As clsToolbar
Dim ParentInfo As clsInfoBereich
Dim ParentBereich As clsOpBereiche

Private Const DefErrModul = "BESTELLUNG.CLS"

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
    Call ParentToolbar.InitToolbar(ParentForm, INI_DATEI, INI_SECTION, "1717")
    
    .cmdToolbar(0).ToolTipText = "ESC Zurück: Zurückschalten auf vorige Bildschirmmaske"
    .cmdToolbar(1).ToolTipText = "F2 Erfassen zusätzlicher Artikel"
    .cmdToolbar(2).ToolTipText = "F3 Interner Programmwechsel"
    .cmdToolbar(3).ToolTipText = "F4 Quittieren der Nachricht"
    .cmdToolbar(4).ToolTipText = "F5 Entfernen eines Artikels oder einer Bedingung"
    .cmdToolbar(5).ToolTipText = "F6 Drucken der Bestellung"
    .cmdToolbar(6).ToolTipText = "F7 Aktualisieren"
    .cmdToolbar(7).ToolTipText = "F8 Zusatztext"
    .cmdToolbar(8).ToolTipText = "F9 Abmelden"
    .cmdToolbar(9).ToolTipText = "shift+F2 Artikel-Status"
    .cmdToolbar(10).ToolTipText = "shift+F3 akutellen Lieferanten wechseln"
    .cmdToolbar(11).ToolTipText = "shift+F4 Angebote"
    .cmdToolbar(12).ToolTipText = "shift+F5 Durchgriff auf Statistik-Anzeige"
    .cmdToolbar(13).ToolTipText = "shift+F6 Datenübertragung zum GH"
    .cmdToolbar(14).ToolTipText = "shift+F7 Warenwert"
    .cmdToolbar(15).ToolTipText = "shift+F8 Bestellvorschlag"
    .cmdToolbar(16).ToolTipText = "shift+F9 Rückkauf-Anfrage"
    'cmdToolbar(19).ToolTipText = "Programm beenden"
    
    Call ParentInfo.InitInfoBereich(.flxInfo(0), INI_DATEI, INFO_SECTION)
    Call ParentInfo.ZeigeInfoBereich("", False)
    .flxInfo(0).row = 0
    .flxInfo(0).col = 0
    
    Call ParentBereich.InitBereich(iForm, iToolbar)
    ParentBereich.ArbeitTitel = False
    ParentBereich.ArbeitLeerzeileOben = False
    ParentBereich.ArbeitWasDarunter = True
    ParentBereich.InfoTitel = False
    ParentBereich.InfoZusatz = ArtikelStatistik%
    ParentBereich.InfoAnzZeilen = iInfo.AnzInfoZeilen
    ParentBereich.AnzahlButtons = 0
    
    With .flxarbeit(0)
        .Cols = 35
        .Rows = 0
        .Rows = 2
        .FixedRows = 1
        .FormatString = "<PZN|^ |<Name|>Menge|^Meh|>BM|>NR|>%|^Lieferant|>POS|<Zusatz|<ActBedingung|||?|^!|v"
        .FormatString = .FormatString + "||||||||||||||||||<±"
    End With
        
    .mnuBearbeitenInd(10).Caption = "&Lieferant"
    .mnuBearbeitenInd(13).Caption = "&Sendevorgang"
    .mnuBearbeitenInd(14).Caption = "Wa&renwert"
    .mnuBearbeitenInd(15).Caption = "Bestell&vorschlag"
    
    .mnuBearbeitenZusatzInd(0).Caption = "Artikel zu W&Ü dazu"
    .mnuBearbeitenZusatzInd(0).Enabled = True
    
    Load .mnuBearbeitenZusatzInd(1)
    .mnuBearbeitenZusatzInd(1).Caption = "&Blind-Bestellung"
    .mnuBearbeitenZusatzInd(1).Enabled = True
        
    Load .mnuBearbeitenZusatzInd(2)
    .mnuBearbeitenZusatzInd(2).Caption = "D&irektbezug verwerfen"
    .mnuBearbeitenZusatzInd(2).Enabled = False
        
    Load .mnuBearbeitenZusatzInd(3)
    .mnuBearbeitenZusatzInd(3).Caption = "Letzte Artikelliste einlesen"
    .mnuBearbeitenZusatzInd(3).Enabled = False
        
    .mnuRahmenAnzeigen.Checked = wpara.FarbeRahmen
'    .mnuAktZeile.Checked = wpara.FarbeAktZeile
    .mnuZusatzInfo.Checked = ArtikelStatistik%
    
    .mnuEinzelBewertung.Enabled = True
    
    .mnuRueckrufe.Enabled = True
    .mnuLiefStammdaten.Enabled = True
    .mnuPosPartner.Enabled = True
    .mnuAepKalkulation.Enabled = False
    .mnuTeilAbschluss.Enabled = False
    
    .mnuEditAufteilung.Enabled = True
End With

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
    
    row% = .flxarbeit(0).row
    'lblInfo(0).Caption = " Information zu " + flxarbeit(0).TextMatrix(row%, 3)
    ActKontrollen$ = RTrim$(.flxarbeit(0).TextMatrix(row%, 11))
    If (ActKontrollen$ = "") Then
        .lblInfo(0).Caption = ""
    Else
        .lblInfo(0).Caption = "Kontrolle: " + ActKontrollen$
    End If
    actzuordnung$ = RTrim$(.flxarbeit(0).TextMatrix(row%, 18))
    
    iRow% = .flxInfo(0).row
    iCol% = .flxInfo(0).col
    Call ParentInfo.ZeigeInfoBereich(pzn$, True)
    If (.ActiveControl.Name <> .flxInfo(0).Name) Then
        Call ZeigeInfoBereichAdd(ActKontrollen$, actzuordnung$, Val(.flxarbeit(0).TextMatrix(row%, 27)), 0)
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
    
            
'    If (RTrim$(.flxarbeit(0).TextMatrix(row%, 14)) <> "") And _
'       ((.flxarbeit(0).TextMatrix(row%, 15) <> "v") Or (InStr(para.Benutz, "Y") <= 0)) Then
    If (BestellAnzeige% = 2) And (IstDirektLief%) Then
        .mnuBearbeitenInd(MENU_SF4).Enabled = True
    ElseIf (RTrim$(.flxarbeit(0).TextMatrix(row%, 15)) <> "") Then
        .mnuBearbeitenInd(MENU_SF4).Enabled = True
    Else
        .mnuBearbeitenInd(MENU_SF4).Enabled = False
    End If
    .cmdToolbar(MENU_SF4).Enabled = .mnuBearbeitenInd(MENU_SF4).Enabled
            
    
    KommWeg% = True
    If (.flxarbeit(0).TextMatrix(row%, 1) = "?") Then
        If (.flxarbeit(0).TextMatrix(row%, 16) = "v") Then
            h$ = .InternerKommentar$
            If (h$ <> "") Then
                Call .AnzeigeKommentar(h$, 3)
                KommWeg% = False
            End If
        End If
    End If
    If (KommWeg%) Then .picQuittieren.Visible = False
    
    .flxEinzelBewertung.Visible = False
    .flxDirektAufteilung.Visible = False
    If (BestellAnzeige% = 2) And (IstDirektLief%) Then
        If (.mnuEinzelBewertung.Checked) Then
            dDummy# = CalcDirektBewertung#(iDummy%, True, False)
        End If
        If (.mnuDirektAufteilung.Checked) Then
            Call ZeigeDirektAufteilung
        End If
    End If
End With

Call DefErrPop
End Sub

Sub ZeigeInfoBereichAdd(ActKontrollen$, actzuordnung$, schwellart%, Index%)
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

With ParentForm.flxInfo(Index%)
    .redraw = False

    h$ = LTrim(RTrim(ActKontrollen$))
    For j% = 1 To ParentInfo.AnzInfoZeilen
        If (h$ = "") Then Exit For
        
        tx$ = DecodKontrollen(h$, Send%)
                
        If (tx$ <> "") Then
            .row = j% - 1
            .col = 0
            .CellFontBold = True
            .CellForeColor = vbRed
            .TextMatrix(j% - 1, 0) = "? " + tx$
        End If
    Next j%
    
'    If (schwellart% > 0) Then
'        .row = j% - 1
'        .col = 0
'        .CellFontBold = True
'        .CellForeColor = vbBlue
'        If (schwellart% = 1) Then
'            tx$ = "Mindestumsatz"
'        Else
'            tx$ = "Günstigster Lieferant"
'        End If
'        .TextMatrix(j% - 1, 0) = "! " + tx$
'        j% = j% + 1
'    Else
        h$ = LTrim$(RTrim$(actzuordnung$))
        Do While (j% <= ParentInfo.AnzInfoZeilen%)
            If (h$ = "") Then Exit Do
            
            tx$ = DecodZuordnungen$(h$)
            
            If (tx$ <> "") Then
                .row = j% - 1
                .col = 0
                .CellFontBold = True
                .CellForeColor = vbBlue
                .TextMatrix(j% - 1, 0) = "! " + tx$
                
                j% = j% + 1
            End If
        Loop
'    End If
    
    Do While (j% <= ParentInfo.AnzInfoZeilen%)
        .TextMatrix(j% - 1, 0) = ""
        j% = j% + 1
    Loop
    
    .redraw = True
End With

Call DefErrPop
End Sub

Function DecodKontrollen$(ActKontr$, Send%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("DecodKontrollen$")
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
Dim ind%, k%
Dim h$, wert1$

h$ = ""
Send% = False
ActKontr$ = LTrim(RTrim(ActKontr$))

If (Len(ActKontr$) > 0) Then
    ind% = InStr(ActKontr$, ",")
    If (ind% > 0) Then
        k% = Val(Left$(ActKontr$, ind% - 1))
        ActKontr$ = Mid$(ActKontr$, ind% + 1)
    Else
        k% = Val(ActKontr$)
        ActKontr$ = ""
    End If
    
    If (k% = 100) Then
        h$ = "manuelle Sendesperre"
        Send% = True
    Else
        wert1$ = RTrim$(Kontrollen(k%).bedingung.wert1)
        If ((k% < AnzKontrollen%) And (wert1$ <> "")) Then
            h$ = wert1$ + " " + RTrim$(Kontrollen(k%).bedingung.op) + " " + RTrim$(Kontrollen(k%).bedingung.wert2)
            If (Kontrollen(k%).Send = "J") Then Send% = True
        End If
    End If
End If

DecodKontrollen$ = h$

Call DefErrPop
End Function

Function DecodZuordnungen$(ActZuord$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("DecodZuordnungen$")
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
Dim ind%, k%
Dim h$, wert1$

h$ = ""
ActZuord$ = LTrim(RTrim(ActZuord$))

If (Len(ActZuord$) > 0) Then
    ind% = InStr(ActZuord$, ",")
    If (ind% > 0) Then
        k% = Val(Left$(ActZuord$, ind% - 1))
        ActZuord$ = Mid$(ActZuord$, ind% + 1)
    Else
        k% = Val(ActZuord$)
        ActZuord$ = ""
    End If
    
    wert1$ = RTrim$(Zuordnungen(k%).bedingung.wert1)
    If ((k% < AnzZuordnungen%) And (wert1$ <> "")) Then
        h$ = wert1$ + " " + RTrim$(Zuordnungen(k%).bedingung.op) + " " + RTrim$(Zuordnungen(k%).bedingung.wert2)
    End If
End If

DecodZuordnungen$ = h$

Call DefErrPop
End Function

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
    .ColWidth(6) = ParentForm.TextWidth("999")
    If (IstDirektLief%) Then
        .ColWidth(7) = ParentForm.TextWidth("99.99")
    Else
        .ColWidth(7) = 0
    End If
    .ColWidth(8) = ParentForm.TextWidth("(XXXXXX)")
    .ColWidth(9) = ParentForm.TextWidth("9999")
    .ColWidth(10) = ParentForm.TextWidth("XXXXXXXX") '("XXXXXXXXXXXXX")
    .ColWidth(11) = 0
    .ColWidth(12) = 0
    .ColWidth(13) = 0
    .ColWidth(14) = ParentForm.TextWidth("WI")
    .ColWidth(15) = ParentForm.TextWidth("WI")
    .ColWidth(16) = ParentForm.TextWidth("WII") '+ wpara.FrmScrollHeight + 2 * wpara.FrmBorderHeight
    .ColWidth(17) = 0   'TextWidth("XXX")
    .ColWidth(18) = 0
    .ColWidth(19) = 0
    .ColWidth(20) = 0
    .ColWidth(21) = 0
    .ColWidth(22) = 0
    .ColWidth(23) = 0
    .ColWidth(24) = 0
    .ColWidth(25) = 0
    .ColWidth(26) = 0
    .ColWidth(27) = 0
    .ColWidth(28) = 0
    .ColWidth(29) = 0
    .ColWidth(30) = 0
    .ColWidth(31) = 0
    .ColWidth(32) = 0
    .ColWidth(33) = 0
    .ColWidth(34) = ParentForm.TextWidth("WI") + wpara.FrmScrollHeight + 2 * wpara.FrmBorderHeight
    ParentForm.Font.Bold = False
    
    
    spBreite% = 0
    For i% = 1 To .Cols - 1
        If (.ColWidth(i%) > 0) And (i% < 14) Then
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
Dim aZr!, zr!
Dim h$, h2$
Dim iFlex As MSFlexGrid
            
Set iFlex = ParentForm.flxarbeit(0)

IstBesorger% = (iFlex.TextMatrix(iFlex.row, 16) = "v") And (InStr(para.Benutz, "Y") > 0)

EditCol% = iFlex.col
If (EditCol% = 5) Or (EditCol% = 6) Then    'Or (EditCol% = 14) Then
'    If (iFlex.TextMatrix(iFlex.row, 15) = "v") And (InStr(para.Benutz, "Y") > 0) Then
    If (IstBesorger%) Then
        Beep
        Call DefErrPop: Exit Sub
    End If
End If

EditModus% = 0
If (EditCol% = 7) Then EditModus% = 4
            
frmAction.tmrAction.Enabled = False

If (EditCol% = 5) Or (EditCol% = 6) Or (EditCol% = 7) Or (EditCol% = 8) Or (EditCol% = 15) Then
            
    rInd% = SucheFlexZeile(False)
    If (rInd% <= 0) Then Call DefErrPop: Exit Sub

    If (iFlex.TextMatrix(iFlex.row, 22) = 0) Then
        
        If (EditCol% = 15) Then
            If (RTrim$(iFlex.TextMatrix(iFlex.row, 15)) <> "") Or (IstDirektLief%) Then
                Call ZeigeLocalAngebote
            End If
        ElseIf (EditCol% = 7) And (Abs(ww.zr) >= 200) Then
            EditTxt$ = Format(0, "0.00")
            Call SplitDirektRabatt(Val(EditTxt$))
        Else
            If (IstDirektLief% = False) Then
                If (iFlex.TextMatrix(iFlex.row, 15) = "!!") Or ((EditCol% = 6) And (IstDirektLief%)) Then
                    Call ZeigeLocalAngebote
                    frmAction.tmrAction.Enabled = True
                    Call DefErrPop: Exit Sub
                End If
            End If
            
            With iFlex
                KeinRowColChange% = True
                aRow% = .row
                .row = 0
                .CellFontBold = True
                .row = aRow%
                KeinRowColChange% = False
            End With
            
            Load frmEdit
            If (EditCol% = 5) Or (EditCol% = 6) Or (EditCol% = 7) Then
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
        '            .Height = frmEdit.ScaleHeight
                    .Left = 0
                    .Top = 0
                    h2$ = iFlex.TextMatrix(iFlex.row, EditCol%)
                    .text = h2$
                    .BackColor = vbWhite
                    .Visible = True
                End With
            ElseIf (EditCol% = 8) Then
                lInd% = -1
                With frmEdit
                    .Left = ParentForm.picBack(0).Left + iFlex.Left + iFlex.ColPos(EditCol%) + 45
                    .Left = .Left + ParentForm.Left + wpara.FrmBorderHeight
                    .Top = ParentForm.picBack(0).Top + iFlex.Top + iFlex.RowHeight(0)
                    .Top = .Top + ParentForm.Top + wpara.FrmBorderHeight + wpara.FrmCaptionHeight + wpara.FrmMenuHeight
                    .Width = ParentForm.lstLieferant.Width
                    .Height = iFlex.Height - iFlex.RowHeight(0)
                End With
                With frmEdit.lstEdit
                    .Height = frmEdit.ScaleHeight
                    frmEdit.Height = .Height
                    .Width = frmEdit.ScaleWidth
                    .Left = 0
                    .Top = 0
                    
                    .Clear
                    .AddItem "(keiner)"
                    .AddItem String$(50, "-")
                    For i% = 1 To AnzLiefNamen%
                        h$ = LiefNamen$(i% - 1)
                        .AddItem h$
                        iLief% = iFlex.TextMatrix(iFlex.row, 17)
                        If (iLief% > 0) And (lInd% < 0) Then
                            ind% = InStr(h$, "(")
                            h$ = Mid$(h$, ind% + 1)
                            If (iLief% = Val(Left$(h$, Len(h$) - 1))) Then
                                lInd% = i% + 1
                            End If
                        End If
                    Next i%
                    If (lInd% < 0) Then
                        .ListIndex = 0
                    Else
                        .ListIndex = lInd%
                    End If
                    .Visible = True
                End With
            End If
            
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
                If (EditCol% = 5) Or (EditCol% = 6) Then
                    ww.SatzLock (1)
                    rInd% = SucheFlexZeile(True)
                    If (rInd% > 0) Then
                        m% = Val(EditTxt$)
                        
                        If (EditCol% = 5) Then
                            aMeng% = ww.bm
                        Else
                            aMeng% = ww.nm
                        End If
        
                        If (aMeng% < 0) Then
                            m% = -m%
                        End If
        
                        If (aMeng% <> m%) Then
                            If (EditCol% = 5) Then
                                ww.bm = m%
                            Else
                                ww.nm = m%
                            End If
        
                            ww.zukontrollieren = Chr$(0)
                            ww.fixiert = Chr$(0)
                            ww.PutRecord (rInd% + 1)
                            Call AuslesenDateiSatz(rInd%, True)
                            Call ErhoeheCounter
                        End If
                    End If
                    ww.SatzUnLock (1)
                ElseIf (EditCol% = 7) Then
                    Call SplitDirektRabatt(Val(EditTxt$))
                Else
                    ind% = InStr(EditTxt$, "(")
                    EditTxt$ = Mid$(EditTxt$, ind% + 1)
            
                    ww.SatzLock (1)
                    rInd% = SucheFlexZeile%(True)
                    If (rInd% > 0) Then
'                        If (EditTxt$ = "(keiner)") Then
'                            bek.lief = 0
'                        Else
'                            bek.lief = Val(Left$(EditTxt$, Len(EditTxt$) - 1))
'                        End If
                        
                        If (EditTxt$ = "(keiner)") Then
                            iLief% = 0
                        Else
                            iLief% = Val(Left$(EditTxt$, Len(EditTxt$) - 1))
                        End If
                        If (iLief% <> ww.Lief) Then
                            ww.Lief = iLief%
                            ww.zukontrollieren = Chr$(0)
                            ww.fixiert = Chr$(0)
                            ww.IstSchwellArtikel = 0
                            ww.PutRecord (rInd% + 1)
                            Call AuslesenDateiSatz(rInd%, True)
                            Call ErhoeheCounter
                        End If
                    End If
                    ww.SatzUnLock (1)
                End If
            End If
            
        End If
    End If
End If

frmAction.tmrAction.Enabled = True

Call DefErrPop
End Sub
      
Sub SplitDirektRabatt(EditRabatt!)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("SplitDirektRabatt")
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
Dim i%, ind%, rInd%, MitAuswahl%
Dim iRabatt!, iZr!, iFr!
Dim txt$
Dim iFlex As MSFlexGrid

Set iFlex = ParentForm.flxarbeit(0)

MitAuswahl% = True
With iFlex
    If (Val(.TextMatrix(.row, 6)) > 0) And (DirektBezugFaktRabattTyp% = 1) Then
        MitAuswahl% = False
        EditErg% = True
        iZr! = EditRabatt!
    End If
End With

If (MitAuswahl%) Then
    With frmEdit
        .Left = ParentForm.picBack(0).Left + iFlex.Left + iFlex.ColPos(7) + 45
        .Left = .Left + ParentForm.Left + wpara.FrmBorderHeight
        .Top = ParentForm.picBack(0).Top + iFlex.Top + (iFlex.row - iFlex.TopRow + 1) * iFlex.RowHeight(0)
        .Top = .Top + ParentForm.Top + wpara.FrmBorderHeight + wpara.FrmCaptionHeight + wpara.FrmMenuHeight
    End With
    With frmEdit.flxEdit
        .Left = 0
        .Top = 0
        
        .Rows = 2
        .FixedRows = 1
        .FormatString = ">Ges.Rabatt|>ZeilenRab.|>FakturenRab"
        .Rows = 1
        .Cols = 3
        
        For i% = 0 To 2
            .ColWidth(i%) = frmEdit.TextWidth("999999999.99")
        Next i%
        
        txt$ = Format(EditRabatt! + DirektBezugFaktRabatt#, "0.00")
        txt$ = txt$ + vbTab + Format(EditRabatt!, "0.00")
        txt$ = txt$ + vbTab + Format(DirektBezugFaktRabatt#, "0.00")
        .AddItem txt$
    
        txt$ = Format(EditRabatt!, "0.00")
        txt$ = txt$ + vbTab + Format(EditRabatt!, "0.00")
        txt$ = txt$ + vbTab
        .AddItem txt$
    
        If (EditRabatt! > DirektBezugFaktRabatt#) Then
            txt$ = Format(EditRabatt!, "0.00")
            txt$ = txt$ + vbTab + Format(EditRabatt! - DirektBezugFaktRabatt#, "0.00")
            txt$ = txt$ + vbTab + Format(DirektBezugFaktRabatt#, "0.00")
            .AddItem txt$
        End If
        
        .Height = .RowHeight(0) * .Rows + 90
        .Width = .ColWidth(0) + .ColWidth(1) + .ColWidth(2) + 90
        
        frmEdit.Height = .Height
        frmEdit.Width = .Width
        
        .row = 1
        .col = 0
        .ColSel = .Cols - 1
        .Visible = True
    End With
    
    frmEdit.Show 1
    
    If (EditErg%) Then
        ind% = InStr(EditTxt$, vbTab)
        iRabatt! = xVal(Left$(EditTxt$, ind% - 1))
        EditTxt$ = Mid$(EditTxt$, ind% + 1)
        ind% = InStr(EditTxt$, vbTab)
        iZr! = xVal(Left$(EditTxt$, ind% - 1))
        iFr! = xVal(Mid$(EditTxt$, ind% + 1))
        If (iFr! > 0) Then
            iZr! = iZr! * (-1)
            If (iZr! = 0!) Then iZr! = -123.45
        End If
    End If
End If

If (EditErg%) Then
    ww.SatzLock (1)
    rInd% = SucheFlexZeile(True)
    If (rInd% > 0) Then
        If (Abs(ww.zr) >= 200) Then
            ww.zr = 200
            If (Abs(iZr!) <> 123.45) Then
                ww.zr = ww.zr + Abs(iZr!)
                If (iZr! < 0) Then
                    ww.zr = ww.zr * (-1)
                End If
            End If
        Else
            ww.zr = iZr!
        End If
        ww.PutRecord (rInd% + 1)
        Call AuslesenDateiSatz(rInd%, True)
    End If
    ww.SatzUnLock (1)

'    If (iFlex.Rows <= 2) Then
'        Call CheckDirektAngebot(1)
'    End If
End If

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
Dim ArbeitRow%
Dim h$

With ParentForm
    ArbeitRow% = .flxarbeit(0).row
    
    h$ = RTrim$(.flxarbeit(0).TextMatrix(ArbeitRow%, 0))
    If (h$ <> "XXXXXXX") Then
        .flxInfo(0).TextMatrix(InfoRow%, 0) = TEXT_BESTELLSTATUS
        InfoRow% = InfoRow% + 1
    End If
    h$ = RTrim$(.flxarbeit(0).TextMatrix(ArbeitRow%, 15))
    If (h$ <> "") Then
        .flxInfo(0).TextMatrix(InfoRow%, 0) = TEXT_SONDERANGEBOTE
        InfoRow% = InfoRow% + 1
    End If
    h$ = RTrim$(.flxarbeit(0).TextMatrix(ArbeitRow%, 9))
    If (h$ <> "") Then
        .flxInfo(0).TextMatrix(InfoRow%, 0) = TEXT_STATISTIK
        InfoRow% = InfoRow% + 1
    End If
'        h$ = RTrim$(.flxarbeit(0).TextMatrix(ArbeitRow%, 15))
'        If (h$ = "v") Then
'            .flxInfo(0).TextMatrix(InfoRow%, 0) = TEXT_ABHOLER
'            InfoRow% = InfoRow% + 1
'        End If
End With

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
Dim i%, j%, AnzRows%, ind%, r%, redraw%, OldRow%, oldCol%, oldStatus%, erg%, iiLief%, iaLief%
Dim MarkZeilen%, MarkPackungen%, AnzBestellWerteCols%, IstBewertungOk%
Dim Preis#, ZeilenWert#
Dim sp&, sp2&, x1&, x2&, y1&, y2&, TextH&, TextW&, NachLinks&, lColor&
Dim tx$, BestellWerte$(3), BestellWerteText$(3)

Dim GesStaffelAbw#, GesNrWert#, GesZeilenRabatt#, GesKo#, GesLagerKo#, GesPersKo#, vt#, Summe#
Dim StaffelAbw#, NrWert#, ZeilenRabatt#, LagerKo#, PersKo#, Bewertung#

DirektBewertung# = 0#

MarkWert# = 0#
GesamtWert# = 0#
Bewertung# = 0#

MarkZeilen% = 0
MarkPackungen% = 0

GesStaffelAbw# = 0#
GesNrWert# = 0#
GesZeilenRabatt# = 0#
GesKo# = 0#
GesLagerKo# = 0#
GesPersKo# = 0#

vt# = 0#

If (IstDirektLief%) Then
    AnzBestellWerteCols% = 4
Else
    AnzBestellWerteCols% = 3
End If

If (BestellAnzeige% = 2) Then
    With ParentForm.flxarbeit(0)
        OldRow% = .row
        oldCol% = .col
        oldStatus% = .redraw
        .redraw = False
        AnzRows% = .Rows - 1
        For i% = 1 To AnzRows%
            tx$ = .TextMatrix(i%, 13)   'EK
            If (tx$ <> "") Then
                iiLief% = Val(.TextMatrix(i%, 17))  'Lief
                iaLief% = Val(.TextMatrix(i%, 22)) 'AktivLief
                If (iaLief% = 0) And ((iiLief% = 0) Or (iiLief% = Lieferant%)) Then
                    .row = i%
                    .col = 5
                    sp2& = .CellBackColor
                    .col = 1
                    sp& = .CellBackColor
'                    If ((IstDirektLief% Or (sp& <> vbRed)) And ((IstDirektLief% = 0) Or (Val(.TextMatrix(i%, 5)) > 0))) Then   'BM
                    If ((IstDirektLief% Or (sp& <> vbRed)) And ((IstDirektLief% = 0) Or (sp2& = vbButtonFace))) Then 'BM
                        MarkZeilen% = MarkZeilen% + 1
                        MarkPackungen% = MarkPackungen% + Val(.TextMatrix(i%, 5))
                
                        Preis# = CDbl(.TextMatrix(i%, 13))
                        ZeilenWert# = Preis# * Val(.TextMatrix(i%, 5))
                        MarkWert# = MarkWert# + ZeilenWert#
                        
'                        If (IstDirektLief%) Then
'                            erg% = ZeilenWerte%(i%, ZeilenWert#, StaffelAbw#, NrWert#, ZeilenRabatt#, LagerKo#, PersKo#)
'                            If (erg%) Then
'                                GesStaffelAbw# = GesStaffelAbw# + StaffelAbw#
'                                GesNrWert# = GesNrWert# + NrWert#
'                                GesZeilenRabatt# = GesZeilenRabatt# + ZeilenRabatt#
'                                GesKo# = GesKo# + LagerKo# + PersKo#
'                                GesLagerKo# = GesLagerKo# + LagerKo#
'                                GesPersKo# = GesPersKo# + PersKo#
'                            End If
'                        End If
                        
                        GesamtWert# = GesamtWert# + ZeilenWert#
                        
'                        Debug.Print i%; ZeilenWert#; GesamtWert#
                    End If
                End If
            End If
        Next i%
        .row = OldRow%
        .col = oldCol%
        .redraw = oldStatus%
    End With


'    If ((IstDirektLief%) And (GesamtWert# <> 0#)) Then
''        Bewertung# = para.Tfaktor + 100# * (GesNrWert# / (GesamtWert# + GesNrWert#) + GesZeilenRabatt# / GesamtWert# + (vt# / 360#) * para.Lagerkosten / 100# - GesKo# / GesamtWert# - GesStaffelAbw# / GesamtWert#) - para.TransDataRabatt * para.FakBar
'        Bewertung# = GesNrWert# / (GesamtWert# + GesNrWert#)
'        Bewertung# = Bewertung# + (GesZeilenRabatt# - GesKo# - GesStaffelAbw#) / GesamtWert#
'        Bewertung# = Bewertung# + (vt# / 360#) * para.Lagerkosten / 100#
'        Bewertung# = Bewertung# * 100#
'        Bewertung# = Bewertung# + 100#
'        Bewertung# = Bewertung# - para.TransDataRabatt * para.FakBar
'    End If
    If (IstDirektLief%) Then DirektBewertung# = CalcDirektBewertung#(IstBewertungOk%, False)
   
    ParentForm.picBestellWerte.Visible = True
Else
    ParentForm.picBestellWerte.Visible = False
End If


If (AnzeigeFlag%) Then
    With ParentForm.picBestellWerte
        .Font = ParentForm.flxarbeit(0).Font
        .Cls

        sp& = .Width / AnzBestellWerteCols%
        x1& = sp&
        For i% = 1 To AnzBestellWerteCols%
            ParentForm.picBestellWerte.Line (x1&, 0)-(x1&, .ScaleHeight), vbBlack
            x1& = x1& + sp&
        Next i%
    
        TextH& = .TextHeight("Äg")
        TextW& = .TextWidth("999999.99") + 2 * Screen.TwipsPerPixelX
       
        
        BestellWerteText$(0) = "Zeilen"
        BestellWerteText$(1) = "Packungen"
        BestellWerteText$(2) = "Wert"
        BestellWerteText$(3) = "Bewertung"
        
        If (BestellAnzeige% = 2) Then
            BestellWerte$(0) = Format(MarkZeilen%, "0")
            BestellWerte$(1) = Format(MarkPackungen%, "0")
            BestellWerte$(2) = Format(MarkWert#, "0.00")
'            BestellWerte$(3) = Format(Bewertung#, "0.0 ") + Format(DirektBewertung#, "0.00")
            BestellWerte$(3) = Format(DirektBewertung#, "0.00")
        Else
            For i% = 0 To 3
                BestellWerte$(i%) = ""
            Next i%
        End If
        
        sp& = .Width / AnzBestellWerteCols%
        x1& = sp& - TextW& - 4 * Screen.TwipsPerPixelX
        x2& = sp& - 2 * Screen.TwipsPerPixelX
    
        For i% = 1 To AnzBestellWerteCols%
            y2& = ParentForm.picBestellWerte.ScaleHeight - 15
            
            tx$ = BestellWerteText$(i% - 1)
            NachLinks& = .TextWidth(tx$)
            .CurrentX = x1& - 2 * Screen.TwipsPerPixelX - NachLinks&
            .CurrentY = 0
            ParentForm.picBestellWerte.Print tx$
            
            lColor& = vbWhite
            If (IstDirektLief%) Then
                If (i% = 3) And (IstBewertungOk% And &H1) Then lColor& = vbRed
                If (i% = 4) And (IstBewertungOk% And &H2) Then lColor& = vbRed
            End If
            ParentForm.picBestellWerte.Line (x1&, 0)-(x2&, y2&), lColor&, BF
            ParentForm.picBestellWerte.Line (x1&, 0)-(x2&, y2&), vbBlack, B
            
            .CurrentX = x1& + 4 * Screen.TwipsPerPixelX
            .CurrentY = 0
            lColor& = .ForeColor
            If (IstDirektLief%) Then
                If (i% = 3) And (IstBewertungOk% And &H1) Then .ForeColor = vbWhite
                If (i% = 4) And (IstBewertungOk% And &H2) Then .ForeColor = vbWhite
            End If
            ParentForm.picBestellWerte.Print BestellWerte$(i% - 1)
            .ForeColor = lColor&
        
            x1& = x1& + sp&
            x2& = x2& + sp&
        Next i%
    End With
    
    If (IstDirektLief%) And (ManuellDirektReduce%) Then
        If (IstBewertungOk% And &H3) Then
            Call ReduceManuelleDirektBestellung
        Else
            ManuellDirektReduce% = 0
        End If
    End If

End If
        
Call DefErrPop
End Sub
   
'Function ZeilenWerte%(pos%, ZeilenWert#, StaffelAbw#, NrWert#, ZeilenRabatt#, LagerKo#, PersKo#)
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Call DefErrFnc("ZeilenWerte%")
'Call DefErrMod(DefErrModul)
'On Error GoTo DefErr
'GoTo DefErrEnd
'DefErr:
'Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
'Case vbRetry
'  Resume
'Case vbIgnore
'  Resume Next
'End Select
'End
'DefErrEnd:
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Dim i%, bm%, nm%, GesMenge%, RabSatz%, bmo%, DiffMenge%, wg%, zr%, ret%
'Dim proz!, RABMIT!
'Dim WieLange#, preis#, fr#
'
'fr# = 0#
'
'ret% = False
'
'ZeilenWert# = 0#
'StaffelAbw# = 0#
'NrWert# = 0#
'ZeilenRabatt# = 0#
'LagerKo# = 0#
'PersKo# = 0#
'
'With ParentForm.flxarbeit(0)
'    'ist etwas eingetragen ?
'    bm% = Val(.TextMatrix(pos%, 5))
'    nm% = Val(.TextMatrix(pos%, 6))
'    If (bm% <> 0) Or (nm% <> 0) Then
'        ret% = True
'
'        GesMenge% = bm% + nm%
'        'Bestellwert errechnen -------------------------------------------
'        preis# = CDbl(.TextMatrix(pos%, 13))
'        ZeilenWert# = preis# * bm%
'
'        'Großhändler Staffelrabatt ---------------------------------------
'        'nur bei 1-er Preisliste Artikel
'        wg% = Val(.TextMatrix(pos%, 24))
'        If (wg% = 1) Then
'
'            RabSatz% = 0
'            For i% = 1 To para.StaffelMax - 2
'                If (GesMenge% >= para.StaffelBM(i%)) And (GesMenge% <= para.StaffelBM(i% + 1)) Then
'                    RabSatz% = i%
'                    Exit For
'                End If
'            Next i%
'            If (RabSatz% > 0) Then
'                'rabatt zwischen RabSatz% und RabSatz%+1
'                proz! = (GesMenge% - para.StaffelBM(RabSatz%)) / (para.StaffelBM(RabSatz% + 1) - para.StaffelBM(RabSatz%))
'                RABMIT! = para.StaffelRabatt(RabSatz%) + (para.StaffelRabatt(RabSatz% + 1) - para.StaffelRabatt(RabSatz%)) * proz! 'linear
'            Else
'                RABMIT! = para.StaffelRabatt(para.StaffelMax)
'            End If
'
'            StaffelAbw# = GesMenge% * preis# * RABMIT! / 100# * para.FakBar
'        End If
'
'        'Naturalrabattwert (mit Faktor) ----------------------------------
'        NrWert# = nm% * preis# * para.FakNatu
'
'        'Zeilen-Barrabattwert (mit Faktor) -------------------------------
'        zr% = Val(.TextMatrix(pos%, 26))
'        ZeilenRabatt# = ((CDbl(zr%) + fr#) / 100#) * CDbl(bm%) * preis# * para.FakBar
''        ZeilenRabatt# = ((CDbl(zr%) / 10# + fr#) / 100#) * CDbl(bm%) * preis# * para.FAKBAR
'
'        'ist optimierung vorhanden ?  (-1)=nein
'        bmo% = Val(.TextMatrix(pos%, 25))
'        If (bmo% > 0) Then
'            'Übervorrat bestellt -----------------------------------------
'            If (GesMenge% > bmo%) Then
'                DiffMenge% = GesMenge% - bmo%
'                WieLange# = (CSng(DiffMenge%) / CSng(bmo%)) * para.BestellPeriode
'                LagerKo# = WieLange# / 360# * (para.Lagerkosten) / 100# * preis# * CDbl(DiffMenge%)
'                PersKo# = WieLange# / 360# * (para.PersonalKosten) / 100# * preis# * CDbl(DiffMenge%)
'            End If
'
'            'zuwenig bestellt --------------------------------------------
'            If (GesMenge% < bmo%) Then
'                DiffMenge% = bmo% - GesMenge%
'                WieLange# = (CSng(DiffMenge%) / CSng(bmo%)) * para.BestellPeriode
'                PersKo# = WieLange# / 360# * (para.PersonalKosten) / 100# * preis# * CDbl(DiffMenge%)
'            End If
'        End If
'    End If
'End With
'
'ZeilenWerte% = ret%
'
'Call DefErrPop
'End Function



Sub AuslesenDateiSatz(pos%, Optional BereitsGelockt% = False)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("AuslesenDateiSatz")
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
Dim i%, j%, Lief%, Max%, l%, AltLief%, ind%, row%, WasGeändert%, AnzeigeCol%
Dim h$, h2$, SRT$, autox$
Dim Lac$, wwOrg$, wwNeu$
Dim menge%
Dim lfa%

If (BereitsGelockt% = False) Then Call ww.SatzLock(1)

AktUhrzeit% = Val(Format(Now, "HHMM"))
NaechstLiefernderLieferant% = HoleNaechstLiefernden%

With ParentForm.flxarbeit(0)
    .redraw = False
    
    ww.GetRecord (pos% + 1)
    
    wwOrg$ = ww.RawData
    
    menge% = Abs(ww.bm)
    lfa% = ww.Lief
    
    If (ww.ssatz > 0) Then
        ass.GetRecord (ww.ssatz + 1)
        ww.PosLag = ass.PosLag
        ww.tplatz = ass.tplatz
        ww.tlager = ass.tlager
    End If
    
    Lac$ = " "
    If ((ww.asatz > 0) And (LaCo$ <> "")) Then
        ast.GetRecord (ww.asatz + 1)
        Lac$ = UCase(ast.Lac)
    End If
    
    Call CheckZuordnen(lfa%, Lac$, menge%)
    
    If (ww.best = "B") Then
        ww.zugeordnet = "J"
    Else
        ww.zugeordnet = "N"
    End If
    
    Call CheckKontrollen(menge%)
    
    Call ZeigeBestellZeile(pos%)
'    .TextMatrix(.row, 18) = SRT$
    
    wwNeu$ = ww.RawData
    If (wwOrg$ <> wwNeu$) And (ww.aktivlief = 0) Then
        ww.PutRecord (pos% + 1)
    End If
    
    
    If (BereitsGelockt% = False) Then Call ww.SatzUnLock(1)
    
    Call ZeigeWerte
    Call EchtKurzInfo
    
    .redraw = True
End With

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
Dim erg%

BekartCounter% = -1
Call AuslesenBestellung(True, False, True)

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
Dim row%, col%, ret%, iLief%
Dim l&
Dim h$, pzn$, txt$, erg$, OrgOpDirektPartner$
Dim IstSek, StartSek

Dim i%

Dim rInd%

frmAction.tmrAction.Enabled = False
Select Case Index

    Case MENU_F5
        Call LoescheBestellZeile
        Call NaechsteBestellZeile
    
    Case MENU_F7
        Call EntferneGeloeschteZeilen(1)
        Call AuslesenBestellung(True, False)
    
    Case MENU_F8
        With ParentForm.flxarbeit(0)
            If (.TextMatrix(.row, 16) = "v") And (InStr(para.Benutz, "Y") > 0) Then
                Call ZeigeBesorgerDaten
            Else
                Call ZusatzFenster(ZUSATZ_ARTIKEL, KorrPzn$, KorrTxt$)
            End If
        End With
    
    Case MENU_SF2
        With ParentForm.flxarbeit(0)
            If (.TextMatrix(.row, 16) = "v") And (InStr(para.Benutz, "Y") > 0) Then
                Call ZeigeBesorgerDaten
            Else
                Call AnzeigeFenster(TEXT_BESTELLSTATUS, KorrPzn$, KorrTxt$)
            End If
        End With
        
    Case MENU_SF3
        Call lif.GetRecord(Lieferant% + 1)
        txt$ = UCase(Trim$(lif.kurz))
        erg$ = MatchCode(1, pzn$, txt$, False, False)
        If (erg$ <> "") Then
            iLief% = Val(pzn$)
            lifzus.GetRecord (iLief% + 1)
            If (lifzus.IstDirektLieferant) Then
                LiefWechselOk% = True
            Else
                frmLieferantenWahl.Show 1
            End If
            If (LiefWechselOk%) Then
                Lieferant% = Val(pzn$)
                Call AuslesenBestellung(True, False, True)
            End If
        End If
    
    Case MENU_SF4
        With ParentForm.flxarbeit(0)
            If (RTrim$(.TextMatrix(.row, 15)) <> "") Or (IstDirektLief%) Then
                Call ZeigeLocalAngebote
            End If
        End With
            
    Case MENU_SF5
        ZeigeStatistik
        
    Case MENU_SF6
        Call AuslesenBestellung2
        If (AnzBestellArtikel% > 0) Then
            If (IstDirektLief%) Then
                Call DirektBezugSenden
            ElseIf (Wbestk2ManuellSenden%) Then
                Wbestk2ManuellVorbereitung% = True
                AutomaticSend% = False
                ManuellSendung% = True
                LeerAuftrag% = False
                frmSenden.Show 1
                Wbestk2ManuellVorbereitung% = False
'                Call HoleIniRufzeiten
'                i% = AnzRufzeiten%
'                Rufzeiten(i%).Lieferant = Lieferant%
'                Rufzeiten(i%).Aktiv = "J"
'                Rufzeiten(i%).AuftragsErg = "ZH"
'                Rufzeiten(i%).AuftragsArt = "  "
'                Rufzeiten(i%).Gewarnt = "J"
'                Rufzeiten(i%).RufZeit = 9999
'                Rufzeiten(i%).LieferZeit = 9999
'                Rufzeiten(i%).LetztSend = 0
'                AnzRufzeiten% = AnzRufzeiten% + 1
'                Call SpeicherIniRufzeiten
'                Call HoleFruehesteManuelleSendezeit%
                Call AuslesenBestellung(True, False)
            Else
                ret% = ModemAktivieren%
                If (ret%) Then
                    AutomaticSend% = False
                    ManuellSendung% = True
                    LeerAuftrag% = False
                    frmSenden.Show 1
                Else
                    Call iMsgBox("Modem momentan belegt !")
                End If
            End If
        End If
    
    Case MENU_SF7
        Call ZeigeWumsatz(DarfWumsatz%)
'        Call ZeigeAuswertung(False)
    
    Case MENU_SF8
        If (IstDirektLief%) Then
            OrgOpDirektPartner$ = OpDirektPartner$
            If (Trim(OpDirektPartner$) <> "") And (IstAutoDirektLief% = 0) Then
                frmDirektPartner.Show 1
            End If
        End If
        
        frmBestVors.Show 1
        Call EntferneGeloeschteZeilen(0)
        ManuellDirektReduce% = True
        Call AuslesenBestellung(True, False, True)
        If (ManuellDirektReduce%) Then
            ManuellDirektReduce% = False
            Call AuslesenBestellung(True, False, True)
        End If
        
        If (IstDirektLief%) Then
            OpDirektPartner$ = OrgOpDirektPartner$
        End If
    Case MENU_SF9
'        If (IstDirektLief%) Then Call CalcDirektBewertung#(ret%, True)
        frmRueckKauf.Show 1
        
End Select

frmAction.tmrAction.Enabled = True

Call DefErrPop
End Sub

Sub LoescheBestellZeile()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("LoescheBestellZeile")
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
Dim rInd%, Geloescht%
Dim pzn$, ch$

frmAction.tmrAction.Enabled = False

ww.SatzLock (1)

rInd% = SucheFlexZeile%(True)
If (rInd% > 0) Then

'    pzn$ = ww.pzn
'    ch$ = Left$(pzn$, 1)
'    Geloescht% = (Asc(ch$) > 127)
'
'    If (Geloescht%) Then
'        ch$ = Chr$(Asc(ch$) - 128)
'    Else
'        ch$ = Chr$(Asc(ch$) + 128)
'        If (ww.best = "B") Then
'            ww.best = " "
'        End If
'    End If
'
'    Mid$(pzn$, 1, 1) = ch$
'    ww.pzn = pzn$
    
    Geloescht% = ww.loesch
    If (Geloescht%) Then
        ww.loesch = 0
    Else
        ww.loesch = 1
    End If
    
    ww.PutRecord (rInd% + 1)
    Call AuslesenDateiSatz(rInd%, True)
    Call ErhoeheCounter
End If

ww.SatzUnLock (1)

frmAction.tmrAction.Enabled = True

Call DefErrPop
End Sub

Sub ToggleBestellZeile()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("ToggleBestellZeile")
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
Dim i%, rInd%

frmAction.tmrAction.Enabled = False

Call ww.SatzLock(1)

rInd% = SucheFlexZeile%(True)
If (rInd% > 0) Then
'    If (ww.zukontrollieren = "1") Or (ww.zukontrollieren = "2") Then
'        ww.zukontrollieren = "4"
'        ww.fixiert = "2"
'    ElseIf (ww.zugeordnet = "J") Then
'        ww.zukontrollieren = "1"
'        ww.fixiert = "1"
'    Else
'        If (ww.lief = 0) Then
'            ww.lief = Lieferant%
'        ElseIf (ww.fixiert = Chr$(0)) Or (ww.fixiert = "2") Then
'            ww.zukontrollieren = "1"
'            ww.fixiert = "1"
'        End If
'    End If
    If (ww.zukontrollieren = "1") Or (ww.zukontrollieren = "2") Then
        ww.zukontrollieren = "4"
        ww.fixiert = "2"
    Else
        ww.zukontrollieren = "1"
        ww.fixiert = "1"
    End If
    
    ww.PutRecord (rInd% + 1)
    Call AuslesenDateiSatz(rInd%, True)
    Call ErhoeheCounter
End If

ww.SatzUnLock (1)

frmAction.tmrAction.Enabled = True

Call DefErrPop
End Sub

Sub ErhoeheCounter()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("ErhoeheCounter")
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
    
ww.GetRecord (1)
ww.erstcounter = (ww.erstcounter + 1) Mod 100
ww.PutRecord (1)
Call DefErrPop
End Sub

Sub NaechsteBestellZeile()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("NaechsteBestellZeile")
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
'    .ColSel = .Cols - 1
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
    Call ToggleBestellZeile
    Call NaechsteBestellZeile
ElseIf (InStr("ABCDEFGHIJKLMNOPQRSTUVWXYZ", UCase(Chr$(KeyAscii))) > 0) Then
    Call frmAction.SelectZeile(UCase(Chr$(KeyAscii)))
End If

Call DefErrPop
End Sub

Sub CheckDirektAngebot(OldRow%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("CheckDirektAngebot")
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
Dim aRow%
Dim rInd%, angInd%, angBm%, angNm%, angLief%, angAep#, angIndOrg%, iAngebot%, angAngebotYOrg%
Dim TmrActionEnabled%
Dim angZr!
Dim l&
Dim h$, ch$, iHerst$
Dim IstTemporaer As Byte, IstManuell As Byte

KeinRowColChange% = True

TmrActionEnabled% = frmAction.tmrAction.Enabled
frmAction.tmrAction.Enabled = False


aRow% = ParentForm.flxarbeit(0).row
ParentForm.flxarbeit(0).row = OldRow%

rInd% = SucheFlexZeile%(True)
If (rInd% > 0) Then
    
    angBm% = Abs(ww.bm)
    If (ww.loesch = 0) And (angBm% > 0) Then
    
        IstTemporaer = 0
        IstManuell = 0
            
        angBm% = ww.bm
        angNm% = ww.nm
        angZr! = ww.zr
        
        angLief% = Lieferant%
        
        iHerst$ = ww.herst
        If (ww.Lief > 0) And (ww.bm <> 0) Then
            lifzus.GetRecord (ww.Lief + 1)
            If (lifzus.IstDirektLieferant) Then
                iHerst$ = iHerst$ + vbCr + Format(ww.Lief, "000")
            End If
        End If
        
        angInd% = -999
        
        angAngebotYOrg% = AngebotY%
        Call ZeigeAngebote(KorrPzn$, KorrTxt$, 1, angInd%, IstTemporaer, IstManuell, AngebotY%, angBm%, angNm%, angLief%, angAep#, angZr!, iHerst$, GhAllgAngebote$)
        If (angAngebotYOrg% <> AngebotY%) Then
            l& = WritePrivateProfileString(UserSection$, "AngebotY", Str$(AngebotY%), INI_DATEI)
        End If
    '    If (angInd% > 100) Then
    '    If (IstTemporaer) Then
    '        Call EigenesAngebotErfassen(KorrPzn$, KorrTxt$, angInd%, IstManuell, angBm%, angNm%, angLief%, angAep#, angZr!)
    '    ElseIf (angInd% >= 0) Then
        If (angInd% >= 0) Then
    '        If (angIndOrg% >= 100) Then angInd% = angInd% - 1
            WirklichKeinRowColChange% = True
            Call AngebotZuordnen(angInd%, IstTemporaer, IstManuell, angBm%, angNm%, angLief%, angAep#, angZr!)
            WirklichKeinRowColChange% = False
        End If
    End If
    
End If

ParentForm.flxarbeit(0).row = aRow%

KeinRowColChange% = False
frmAction.tmrAction.Enabled = TmrActionEnabled%

Call DefErrPop
End Sub

Sub ZeigeLocalAngebote()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("ZeigeLocalAngebote")
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
Dim rInd%, angInd%, angBm%, angNm%, angLief%, angAep#, angIndOrg%, iAngebot%, angAngebotYOrg%
Dim angZr!
Dim l&
Dim h$, ch$, iHerst$
Dim IstTemporaer As Byte, IstManuell As Byte

'With ParentForm.flxarbeit(0)
'    If (.TextMatrix(.row, 15) = "v") And (InStr(para.Benutz, "Y") > 0) Then
'        Beep
'        Call DefErrPop: Exit Sub
'    End If
'End With

rInd% = SucheFlexZeile%(True)
If (rInd% > 0) Then
    
    iAngebot% = ww.angebot
    
    IstTemporaer = 0
    IstManuell = 0
    If (iAngebot% And &H10) Then IstTemporaer = 1
    If (iAngebot% And &H20) Then IstManuell = 1
    
    angBm% = Abs(ww.bm)
    If (angBm% = 0) Then
        Call iMsgBox("Problem: Keine Bestellmenge eingetragen!", vbInformation)
        Call DefErrPop: Exit Sub
    End If

    
    If (iAngebot% And &H4) Then
        angInd% = ww.AngebotInd
        If (angInd% = 250) Then angInd% = 1000 + ww.Lief
    Else
        angInd% = -1
    End If
    
'    If (angInd% >= 100) Then
    If (IstTemporaer) Then
        angBm% = ww.bm
        angNm% = ww.nm
    End If
    angIndOrg% = angInd%
    
    If (IstDirektLief%) Then
        angLief% = Lieferant%
    Else
        angLief% = -1
    End If
    
    iHerst$ = ww.herst
    If (ww.Lief > 0) And (ww.bm <> 0) Then
        lifzus.GetRecord (ww.Lief + 1)
        If (lifzus.IstDirektLieferant) Then
            iHerst$ = iHerst$ + vbCr + Format(ww.Lief, "000")
        End If
    End If
    
    angAngebotYOrg% = AngebotY%
    Call ZeigeAngebote(KorrPzn$, KorrTxt$, 1, angInd%, IstTemporaer, IstManuell, AngebotY%, angBm%, angNm%, angLief%, angAep#, angZr!, iHerst$, GhAllgAngebote$)
    If (angAngebotYOrg% <> AngebotY%) Then
        l& = WritePrivateProfileString(UserSection$, "AngebotY", Str$(AngebotY%), INI_DATEI)
    End If
'    If (angInd% > 100) Then
'    If (IstTemporaer) Then
'        Call EigenesAngebotErfassen(KorrPzn$, KorrTxt$, angInd%, IstManuell, angBm%, angNm%, angLief%, angAep#, angZr!)
'    ElseIf (angInd% >= 0) Then
    If (angInd% >= 0) Then
'        If (angIndOrg% >= 100) Then angInd% = angInd% - 1
        Call AngebotZuordnen(angInd%, IstTemporaer, IstManuell, angBm%, angNm%, angLief%, angAep#, angZr!)
    End If
    
End If

Call DefErrPop
End Sub

Function CalcDirektBewertung#(IstBewertungOk%, AnzeigeFlag%, Optional GesamtBewertung% = True, Optional LiefEinlesen% = True, Optional iBevorratungsZeit% = -1)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("CalcDirektBewertung#")
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
Dim rInd%, angInd%, angBm%, angNm%, angLief%, angAep#, angIndOrg%, iAngebot%, angAngebotYOrg%
Dim iiLief%, iaLief%
Dim angZr!
Dim l&
Dim h$, ch$

Dim angBMopt!, angBMast%, angWgAst%, angTaxeAep#, angAstAep#, angPznLagernd%

Dim i%, OldRow%, oldCol%, oldStatus%, AnzRows%, spBreite%, Y%
Dim sp&
Dim ret#
Dim tx$, FormStr$

'With ParentForm.flxarbeit(0)
'    If (.TextMatrix(.row, 15) = "v") And (InStr(para.Benutz, "Y") > 0) Then
'        Beep
'        Call DefErrPop: Exit Sub
'    End If
'End With

If (LiefEinlesen%) Then lifzus.GetRecord (Lieferant% + 1)

ret# = 0#
IstBewertungOk% = 3 ' False
'If (GesamtBewertung%) Then
    With ParentForm.flxarbeit(0)
        Call InitDirektBewertung(Lieferant%)
        OldRow% = .row
        oldCol% = .col
        oldStatus% = .redraw
        .redraw = False
        AnzRows% = .Rows - 1
        For i% = 1 To AnzRows%
            If (GesamtBewertung% = False) Then i% = OldRow%
            tx$ = .TextMatrix(i%, 13)   'EK
            If (tx$ <> "") Then
                iiLief% = Val(.TextMatrix(i%, 17))  'Lief
                iaLief% = Val(.TextMatrix(i%, 22)) 'AktivLief
                If (iaLief% = 0) And ((iiLief% = 0) Or (iiLief% = Lieferant%)) Then
                    .row = i%
                    .col = 5
                    sp& = .CellBackColor
'                    If ((sp& <> vbRed) And ((IstDirektLief% = 0) Or (Val(.TextMatrix(i%, 5)) > 0))) Then    'BM
'                    If ((IstDirektLief% = 0) Or (Val(.TextMatrix(i%, 5)) > 0)) Then
                    If (sp& = vbButtonFace) Then
                        rInd% = SucheFlexZeile%(True)
                        If (rInd% > 0) Then
                            
                            If (iBevorratungsZeit% > 0) Then
                                angBm% = CalcDirektBM(iBevorratungsZeit%, xVal(.TextMatrix(i%, 25)), Val(.TextMatrix(i%, 9)))
                            Else
                                angBm% = Abs(ww.bm)
                            End If
                            
                            If (angBm% > 0) Then
                                angNm% = ww.nm
                                angZr! = ww.zr
                                
                                angBMopt! = CDbl(.TextMatrix(i%, 28))
                                angBMast% = Val(.TextMatrix(i%, 29))
                                angWgAst% = Val(.TextMatrix(i%, 31))
                                angTaxeAep# = CDbl(.TextMatrix(i%, 30))
                                angAstAep# = CDbl(.TextMatrix(i%, 32))
                                angPznLagernd% = Val(.TextMatrix(i%, 33))
                                
                                Call RechneDirektBewertung(KorrPzn$, angBm%, angNm%, angZr!, angBMopt!, angBMast%, angWgAst%, angTaxeAep#, angAstAep#, angPznLagernd%)
                            End If
                        End If
                    End If
                End If
            End If
            If (GesamtBewertung% = False) Then Exit For
        Next i%
        .row = OldRow%
        .col = oldCol%
        .redraw = oldStatus%
    End With
        
        If (GesamtBewertung%) Then
            angAngebotYOrg% = AngebotY%
            ret# = ZeigeDirektBewertung(IstBewertungOk%, AngebotY%, AnzeigeFlag%)
            If (angAngebotYOrg% <> AngebotY%) Then
                l& = WritePrivateProfileString(UserSection$, "AngebotY", Str$(AngebotY%), INI_DATEI)
            End If
        Else
            With ParentForm.flxEinzelBewertung
            
                .Cols = 2
                .Rows = 13
                .FixedRows = 1
                .FixedCols = 1
                
                .Height = .RowHeight(0) * 14 + 90
                
                FormStr$ = ""
                For i% = 1 To 30
                    FormStr$ = FormStr$ + "|^" + Mid$(Str$(i%), 2)
                Next i%
                FormStr$ = FormStr$ + ";|Lieferant|Orig.Angebot|Angebot|AEP|Rabatt|Preis|NR|aliq.Lager|aliq.Best|Staffel|kalkAEP|Gespart|%|||"
                .FormatString = FormStr$
                .SelectionMode = flexSelectionByColumn
                
                .FillStyle = flexFillRepeat
                .row = 11
                .col = 0
                .RowSel = .Rows - 1
                .ColSel = .Cols - 1
                .CellBackColor = vbWhite
                
                .row = 2
                .col = 0
                .RowSel = 3
                .ColSel = .Cols - 1
                .CellFontBold = True
                
                .row = 6
                .col = 0
                .RowSel = 7
                .ColSel = .Cols - 1
                .CellFontBold = True
                
                .FillStyle = flexFillSingle
                
                .Cols = 2
                For i% = 0 To .Cols - 1
                    .ColWidth(i%) = ParentForm.TextWidth("WWWWWWWW ")
                    .ColAlignment(i%) = flexAlignRightCenter
                Next i%
            
                For i% = 0 To .Cols - 1
                    spBreite% = spBreite% + .ColWidth(i%)
                Next i%
                .Width = spBreite% + 90
                
                Y% = ParentForm.flxarbeit(0).Top + ParentForm.flxarbeit(0).RowPos(OldRow%) + ParentForm.flxarbeit(0).RowHeight(0)
                If ((Y% + .Height) > ParentForm.picBack(0).Height) Then
                    Y% = ParentForm.flxarbeit(0).Top + ParentForm.flxarbeit(0).RowPos(OldRow%) - .Height
                End If
                .Top = Y%
                .Left = (ParentForm.ScaleWidth - .Width) / 2
                
                angAngebotYOrg% = AngebotY%
                AnzeigeFlag% = 2
                ret# = ZeigeDirektBewertung(IstBewertungOk%, AngebotY%, AnzeigeFlag%)
                
                .Visible = True
            End With
        End If
        
'Else
'    rInd% = SucheFlexZeile%(True)
'    If (rInd% > 0) Then
'
'        angInd% = -1
'        angBm% = ww.bm
'        angNm% = ww.nm
'        angZr! = ww.zr
'
'        If (IstDirektLief%) Then
'            angLief% = Lieferant%
'        Else
'            angLief% = -1
'        End If
'
'        angAngebotYOrg% = AngebotY%
'        Call ZeigeAngebote(KorrPzn$, KorrTxt$, 2, angInd%, 0, 0, AngebotY%, angBm%, angNm%, angLief%, angAep#, angZr!)
'        If (angAngebotYOrg% <> AngebotY%) Then
'            l& = WritePrivateProfileString(UserSection$, "AngebotY", Str$(AngebotY%), INI_DATEI)
'        End If
'    End If
'End If

CalcDirektBewertung# = ret#

Call DefErrPop
End Function

Sub CalcAlleDirektBM(iBevorratungsZeit%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("CalcAlleDirektBM")
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
Dim rInd%
Dim iiLief%, iaLief%
Dim i%, OldRow%, oldCol%, oldStatus%, AnzRows%
Dim tx$

With ParentForm.flxarbeit(0)
    OldRow% = .row
    oldCol% = .col
    oldStatus% = .redraw
    .redraw = False
    AnzRows% = .Rows - 1
    For i% = 1 To AnzRows%
        tx$ = .TextMatrix(i%, 13)   'EK
        If (tx$ <> "") Then
            iaLief% = Val(.TextMatrix(i%, 22)) 'AktivLief
            If (iaLief% = 0) Then
                .row = i%
                rInd% = SucheFlexZeile%(True)
                If (rInd% > 0) Then
                    If (ww.Lief = 0) Or (ww.Lief = Lieferant%) Or (ww.lief1 = Lieferant%) Then
                        ww.bm = CalcDirektBM(iBevorratungsZeit%, xVal(.TextMatrix(i%, 25)), Val(.TextMatrix(i%, 9)))
                        ww.PutRecord (rInd% + 1)
                    End If
                End If
            End If
        End If
    Next i%
    .row = OldRow%
    .col = oldCol%
    .redraw = oldStatus%
End With

Call DefErrPop
End Sub

Sub CalcAlleDirektAngebote()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("CalcAlleDirektAngebote")
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
Dim rInd%
Dim iiLief%, iaLief%
Dim i%, OldRow%, oldCol%, oldStatus%, AnzRows%
Dim tx$, h$

DirektZuordnungenHalten% = True

With ParentForm.flxarbeit(0)
    OldRow% = .row
    oldCol% = .col
    oldStatus% = .redraw
    .redraw = False
    AnzRows% = .Rows - 1
    For i% = 1 To AnzRows%
        tx$ = .TextMatrix(i%, 13)   'EK
        If (tx$ <> "") Then
            iaLief% = Val(.TextMatrix(i%, 22)) 'AktivLief
            If (iaLief% = 0) Then
                .row = i%
                rInd% = SucheFlexZeile%(True)
                If (rInd% > 0) Then
                    If (ww.Lief = 0) Or (ww.Lief = Lieferant%) Or (ww.lief1 = Lieferant%) Then
                        ww.angebot = 0
                        ww.Lief = ww.lief1
                        Call EinzelSatz(0, h$)
                        ww.PutRecord (rInd% + 1)
                    End If
                End If
            End If
        End If
    Next i%
    .row = OldRow%
    .col = oldCol%
    .redraw = oldStatus%
End With

DirektZuordnungenHalten% = False

Call DefErrPop
End Sub

Sub EigenesAngebotErfassen(pzn$, txt$, angInd%, angManuell As Byte, angBm%, angNm%, angLief%, angAep#, angZr!)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("EigenesAngebotErfassen")
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
Dim OrgManuellBm%

ManuellPzn$ = pzn$
ManuellTxt$ = txt$
ManuellBm% = angBm%
ManuellNm% = angNm%
OrgManuellBm% = ManuellBm%
Load frmErfassung
With frmErfassung
    .txtManuell(0).text = ManuellBm%
    .txtManuell(1).text = ManuellNm%
    .Show 1
End With
If (ManuellErg%) Then
    angBm% = ManuellBm%
    angNm% = ManuellNm%
    If (angBm% < OrgManuellBm%) Then angInd% = 999
    Call AngebotZuordnen(angInd%, 1, angManuell, angBm%, angNm%, angLief%, angAep#, angZr!)
End If

Call DefErrPop
End Sub

Sub AngebotZuordnen(angInd%, angTemporaer As Byte, angManuell As Byte, angBm%, angNm%, angLief%, angAep#, angZr!)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("AngebotZuordnen")
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
Dim rInd%, iAngebot%
Dim iFlex As MSFlexGrid
            
Set iFlex = ParentForm.flxarbeit(0)
If ((iFlex.TextMatrix(iFlex.row, 16) = "v") And (InStr(para.Benutz, "Y") > 0) And _
    (Val(iFlex.TextMatrix(iFlex.row, 5)) <> angBm%) And (angInd% > 0)) Then
    Beep
    Call DefErrPop: Exit Sub
End If


ww.SatzLock (1)

rInd% = SucheFlexZeile%(True)
If (rInd% > 0) Then
    
    iAngebot% = ww.angebot
    If (angInd% = 200) Then '999
        ww.bm = angBm%
        ww.nm = angNm%
        ww.Lief = angLief%
        ww.nnart = 0
        ww.NNAEP = 0#
        iAngebot% = iAngebot% And (&H34 Xor &HFFFF)
    ElseIf (angInd% > 0) Then
        ww.bm = angBm%
        ww.nm = angNm%
        ww.Lief = angLief%
        If (iAngebot% And &H1) Then
            iAngebot% = iAngebot% And (&H1 Xor &HFFFF)
            iAngebot% = iAngebot% Or &H2
        End If
        iAngebot% = iAngebot% Or &H4
        ww.nnart = 2
        ww.NNAEP = FNX(angAep#)
        ww.zr = angZr!
        iAngebot% = iAngebot% And (&H30 Xor &HFFFF)
        If (angTemporaer) Then iAngebot% = iAngebot% Or &H10
        If (angManuell) Then iAngebot% = iAngebot% Or &H20
        ww.AngebotInd = angInd%     '- 1
    Else
        ww.bm = ww.bm1
        ww.nm = ww.nm1
        ww.Lief = ww.lief1
        ww.nnart = 0
        ww.NNAEP = 0#
        ww.zr = -123.45 ' 0!
        iAngebot% = iAngebot% And (&H34 Xor &HFFFF)
    End If
    ww.angebot = iAngebot%
    
    ww.zukontrollieren = Chr$(0)
    ww.fixiert = Chr$(0)
            
    ww.PutRecord (rInd% + 1)
    Call AuslesenDateiSatz(rInd%, True)
    Call ErhoeheCounter
End If
ww.SatzUnLock (1)

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
Dim Lm%, BuchM%, rInd%

If (h$ = TEXT_SONDERANGEBOTE) Then
    Call MenuBearbeiten(MENU_SF4)
'ElseIf (h$ = TEXT_ABHOLER) Then
'    Call ZeigeBesorgerDaten
ElseIf (h$ = TEXT_BESTELLSTATUS) Then
    Call MenuBearbeiten(MENU_SF2)
ElseIf (h$ = TEXT_STATISTIK) Then
    Call MenuBearbeiten(MENU_SF5)
End If

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
Dim row%, AbholNr%
Dim h$

With ParentForm.flxarbeit(0)
    row% = .row
    h$ = RTrim$(.TextMatrix(row%, 12))
    AbholNr% = Val(h$)
End With

Call BesorgerInfo(AbholNr%, KorrPzn$, KorrTxt$)

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
OptionenNeu% = False
frmOptionen.Show 1
If (OptionenNeu%) Then
    Call AuslesenBestellung(True, False, True)
End If
Call DefErrPop
End Sub

Sub ManuellSpeichern(Optional MitAuslesen% = False)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("ManuellSpeichern")
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
Dim erg%

'If (ManuellSsatz% = 0) Then
'    If (ManuellPzn$ = "9999999") Or (ManuellAsatz% = 0) Then
'        ManuellBm% = -ManuellBm%: ManuellNm% = -ManuellNm%
'    Else
'        erg% = MsgBox("Ist dieser Artikel bereits auf Lager ?", vbYesNo Or vbInformation)
'        If (erg% = vbNo) Then
'            ManuellBm% = -ManuellBm%: ManuellNm% = -ManuellNm%
'        End If
'    End If
'End If

Call NeuerDS

If (MitAuslesen%) Then
    Call AuslesenBestellung(True, False, True)
End If

Call DefErrPop
End Sub

Sub ManuellFertig()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("ManuellFertig")
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
Dim erg%

Call AuslesenBestellung(True, False, True)

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
Dim i%, erg%, Max%, LetztArtikelListe%
Dim pzn$, h$

frmAction.tmrAction.Enabled = False
    
Select Case Index
    
    Case 0
'        erg% = frmWuFrage.EinlesenWuHeader%
'        If (erg% = 0) Then
'            Call iMsgBox("Problem: Für diesen Lieferanten ist noch keine WÜ gespeichert!", vbInformation)
'        Else
'            WuFrageTxt$ = " Welcher Lieferung soll " + KorrTxt$ + " angefügt werden :"
'            WuFrageErg% = 1
'            frmWuFrage.Show 1
'            If (WuFrageErg%) Then
'            End If
'        End If
        
        WuAuswahlModus% = 2
        frmWuAuswahl.Show 1
        If (WuLifDat$ <> "") Then
            Call UpdateEinzelZeile(WuLifDat$)
        End If
        
    
    Case 1
        If (iMsgBox("Blindbestellung durchführen ?", vbYesNo Or vbDefaultButton2) = vbYes) Then
            Call AuslesenBestellung2
            If (AnzBestellArtikel% > 0) Then
                BlindBestellung% = True
                Call HoleLieferantenDaten(Lieferant%)
                Call SucheSendeArtikel
                Call SaetzeVorbereiten
                Call SaetzeSenden
                Call UpdateBekartDat(Lieferant%, True)
                BlindBestellung% = False
            End If
        End If
    
    Case 2
        Call DirektBezugVerwerfen
        Call MenuBearbeiten(MENU_F7)
    
    Case Else
        lifzus.GetRecord (Lieferant% + 1)
        DirektBevorratungsZeitraum% = lifzus.TempBevorratungsZeitraum
        If (DirektBevorratungsZeitraum% = 0) Then DirektBevorratungsZeitraum% = lifzus.BevorratungsZeitraum
        If (DirektBevorratungsZeitraum% = 0) Then DirektBevorratungsZeitraum% = para.BestellPeriode
        
        PalettenModus% = 1
        
        LetztArtikelListe% = FileOpen("menglist.dat", "I")
        Do While Not (EOF(LetztArtikelListe%))
            Line Input #LetztArtikelListe%, h$
            pzn$ = Left$(h$, 7)
            If (Val(pzn$) > 0) Then Call ArtikelInPalette(pzn$)
        Loop
        Close #LetztArtikelListe%
        
        DirektZuordnungenHalten% = True
        Call MenuBearbeiten(MENU_F7)
        DirektZuordnungenHalten% = False
    
End Select

frmAction.tmrAction.Enabled = True

Call DefErrPop
End Sub

Sub DirektBezugVerwerfen(Optional AlleWeg% = True)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("DirektBezugVerwerfen")
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
Dim i%, Max%
        
ww.SatzLock (1)
ww.GetRecord (1)
Max% = ww.erstmax
For i% = 1 To Max%
    ww.GetRecord (i% + 1)
    If (ww.status = 1) And (ww.loesch = 0) And (ww.aktivlief = 0) And (ww.DirektTyp = 2) Then
        If (ww.Lief = Lieferant%) Or (AlleWeg% And (ww.lief1 = Lieferant%)) Then
            ww.loesch = 1
            ww.status = 0
            ww.PutRecord (i% + 1)
        End If
    End If
Next i%
ww.SatzUnLock (1)

Call DefErrPop
End Sub
    
Sub AuslesenBestellung2()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("AuslesenBestellung2")
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
Dim i%, j%, Lief%, Max%, l%, AltLief%, ind%, row%
Dim zr!
Dim Preis#, ZeilenWert#, EK#, fr#
Dim h$, h2$, SRT$, autox$, AktPzn$, tx$
Dim Lac$
Static IstAktiv%

If (IstAktiv% = True) Then Call DefErrPop: Exit Sub

IstAktiv% = True

ww.SatzLock (1)
ww.GetRecord (1)

Max% = ww.erstmax
BekartCounter% = ww.erstcounter
GlobBekMax% = Max%

frmAction!lstSortierung.Clear
frmAction!lstDirektSortierung.Clear

AnzBestellArtikel% = 0

MarkWert# = 0#
GesamtWert# = 0#

If (IstDirektLief%) Then
    For i% = 1 To Max%
        ww.GetRecord (i% + 1)
    
    '    If (Asc(ww.pzn) < 128) And (ww.aktivlief = 0) Then
        If (ww.status = 1) And (ww.loesch = 0) And (ww.aktivlief = 0) And ((ww.Lief = 0) Or (ww.Lief = Lieferant%)) Then
'            If ((ww.lief = Lieferant%) And (ww.zukontrollieren = "1") And (Abs(ww.bm) > 0)) Then
            If ((ww.Lief = Lieferant%) And (ww.zukontrollieren = "1") And (Abs(ww.bm) > 0) And (Abs(ww.bm) >= DirektBezugAbBm%)) Then
                Call iMsgBox("Direktbestellung nicht durchführbar, da noch Artikel zur Kontrolle !", vbInformation)
                IstAktiv% = False
                ww.SatzUnLock (1)
                Call DefErrPop: Exit Sub
            End If
        End If
    Next i%
End If
        
        
For i% = 1 To Max%
    ww.GetRecord (i% + 1)

    If (ww.status = 1) And (ww.loesch = 0) And (ww.aktivlief = 0) Then
        If (ww.Lief = Lieferant%) Or ((ww.Lief = 0) And (IstDirektLief% = 0)) Then
        
        
            Preis# = ww.aep
            ZeilenWert# = Preis# * Abs(ww.bm)
    
            GesamtWert# = GesamtWert# + ZeilenWert#
    
'            If (ww.zugeordnet = "J") And (ww.zukontrollieren <> "1") And ((IstDirektLief% = 0) Or (Abs(ww.bm) > 0)) Then
            If (ww.zugeordnet = "J") And (ww.zukontrollieren <> "1") And ((IstDirektLief% = 0) Or ((Abs(ww.bm) > 0) And (Abs(ww.bm) >= DirektBezugAbBm%))) Then
                AnzBestellArtikel% = AnzBestellArtikel% + 1
                h$ = Left$(ww.txt, 18) + Mid$(ww.txt, 29) + Format(i%, "0000") + ww.pzn + Format(Abs(ww.bm), "0000") + Format(ww.BekLaufNr, String(9, 48))
                If (para.Land = "A") Then
                    h$ = h$ + ww.wg
                End If
                frmAction!lstSortierung.AddItem h$
                ww.aktivlief = Lieferant%
'                If (Wbestk2ManuellSenden%) And (IstDirektLief% = 0) Then
                If (Wbestk2ManuellSenden%) Then
                    ww.aktivind = -AnzBestellArtikel%
                Else
                    ww.aktivind = AnzBestellArtikel%
                End If
                ww.PutRecord (i% + 1)
                MarkWert# = MarkWert# + ZeilenWert#
               
                If (IstDirektLief%) Then
                    Call InsertLstDirektSortierung
                End If
            End If
        End If
    End If
Next i%

If (AnzBestellArtikel% > 0) Then
    BekartCounter% = (BekartCounter% + 1) Mod 100
    ww.erstcounter = BekartCounter%
    ww.PutRecord (1)
End If

IstAktiv% = False

ww.SatzUnLock (1)

Call DefErrPop
End Sub

Sub ZeigeBestellZeile(pos%, Optional HatDirektLief% = False)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("ZeigeBestellZeile")
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
Dim i%, j%, Lief%, l%, row%, col%, iBold%, iItalic%, FirstRed%, iAngebot%, red14%, xAktiv%
Dim BackDunkel%
Dim lBack&
Dim EK#, Rabatt!, bmo!
Dim h$, h2$, s$, LiefName$, ArtName$, ArtMenge$, ArtMeh$, VonWo$, nm$, zusatz$, KontrollChar$
Dim ActKontrollen$, actzuordnung$, SRT$, SQLStr$, ZusInfo$
Dim ArtikelKz As Byte

KeinRowColChange% = True

bmo! = 0

BackDunkel% = 0

If (BestellAnzeige% = 0) Then
    h$ = ww.txt
    If (ww.aktivlief > 0) Then h$ = "ZZZ" + h$
ElseIf (BestellAnzeige% = 2) And (IstDirektLief%) Then
    If (ww.ssatz > 0) Then
        If (ass.opt <= 0) Then
            bmo! = 0
        Else
            bmo! = ass.opt
'            bmo% = Int(ass.opt + 0.501)
    '        If (bmo% = 0) Then bmo% = 1
        End If
    End If
    
    h$ = " "
    If (ww.bm <= 0) Then
        h$ = "*"
        BackDunkel% = True
    ElseIf (DirektBezugAbBm% > 0) And (ww.bm < DirektBezugAbBm%) Then
        h$ = "!"
        BackDunkel% = True
    End If
    h$ = h$ + ww.txt
Else
    h$ = ww.zukontrollieren + ww.txt
End If
SRT$ = h$ + Format(pos%, "00000")

LiefName$ = ""
Lief% = ww.Lief
'h$ = Format(lief%, "###    ")
If (Lief% > 0) And (Lief% <= lif.AnzRec) Then
    lif.GetRecord (Lief% + 1)
    h2$ = lif.kurz
'    If (Trim(h2$) = "") Then h2$ = "???"
    If (Trim(h2$) = "") Or (Asc(Left$(h2$, 1)) < 32) Then h2$ = Format(Lief%, "0")
    LiefName$ = h2$
End If
If (ww.fixiert = "1") Then
    LiefName$ = "(" + LiefName$ + ")"
End If

If (ww.pzn = "9999999") Then
    h$ = RTrim$(ww.txt)
    Call OemToChar(h$, h$)
    ArtName$ = h$
    ArtMenge$ = ""
    ArtMeh$ = ""
Else
    h$ = RTrim(Left$(ww.txt, 33))
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
    ArtMeh$ = Mid$(ww.txt, 34)
End If

EK# = ww.aep

nm$ = " "
If (Abs(ww.nm) > 0) Then
    nm$ = Format(Abs(ww.nm), " 0")
End If

zusatz$ = ""
If (ww.absage > 0) Then zusatz$ = zusatz$ + "Absage-" + Mid$(Str$(ww.absage), 2)
If (ww.bm = 0) And (ww.nm = 0) And (IstDirektLief% = 0) And (HatDirektLief% = 0) Then zusatz$ = zusatz$ + "Anfrage "
If (ww.ArtikelKz2 And KZ_ISTINTERIM) Then zusatz$ = "Interim"


VonWo$ = "m"
If (ww.auto = "v") Then
    VonWo$ = "v"
ElseIf (ww.auto = "+") Then
    VonWo$ = ""    '"a"
End If
'    zusatz$ = zusatz$ + VonWo$ + " "

'If (Asc(Left$(ww.pzn, 1)) > 127) Then
'    zusatz$ = "GELÖSCHT"
'End If
If (ww.loesch) Then
    zusatz$ = "GELÖSCHT"
End If

ActKontrollen$ = ""
If (ww.fixiert = "1") Then
    ActKontrollen$ = "100"
End If
For j% = 0 To 5
    If (ww.actkontrolle(j%) < 111) Then
        If (ActKontrollen$ <> "") Then ActKontrollen$ = ActKontrollen$ + ","
        ActKontrollen$ = ActKontrollen$ + Mid$(Str$(ww.actkontrolle(j%)), 2)
    Else
        Exit For
    End If
Next j%

actzuordnung$ = ""
If (ww.actzuordnung < 111) Then
    actzuordnung$ = Mid$(Str$(ww.actzuordnung), 2)
End If


KontrollChar$ = " "
FirstRed% = False
If (ww.fixiert = "2") Then
    KontrollChar$ = Chr$(214)
ElseIf (ActKontrollen$ <> "") Then
    KontrollChar$ = "?"
    If (ww.zukontrollieren = "1") Then
        FirstRed% = True
    End If
End If

ZusInfo$ = ""
ArtikelKz = ww.ArtikelKz
If (ArtikelKz And KZ_ZUSTEXT) Then
    If (ArtikelKz And KZ_MERKZETTEL) Then
        ZusInfo$ = "±"
    Else
        ZusInfo$ = "+"
    End If
ElseIf (ArtikelKz And KZ_MERKZETTEL) Then
    ZusInfo$ = "-"
End If
        
Rabatt! = 0!
If (IstDirektLief%) Then
    Rabatt! = ww.zr
    If (Rabatt! <= -200) Then
        Rabatt! = Rabatt! + 200
    ElseIf (Rabatt! >= 200) Then
        Rabatt! = Rabatt! - 200
    End If
    If (ww.zr < 0) Then
        If (Rabatt! = -123.45) Then Rabatt! = 0#
        If (ww.nm = 0) Or (DirektBezugFaktRabattTyp% <> 1) Then
            Rabatt! = Rabatt! * (-1) + DirektBezugFaktRabatt#
        End If
    End If
End If


With frmAction.flxarbeit(0)

    col% = .col

    iItalic% = False
    lBack& = vbButtonFace

'    If (ww.loesch) Or (ww.zukontrollieren = "1") Then
    If (ww.loesch) Then
        lBack& = wpara.FarbeDunklerBereich  ' FarbeGray& 'vbGrayText
    End If
'    If (NurZuKontrollierende% = False) And (ww.zugeordnet = "N") Then
'        lBack& = vbGrayText
'    End If
    If (ww.aktivlief > 0) Then
        lBack& = vbGreen
        lif.GetRecord (ww.aktivlief + 1)
        zusatz$ = lif.kurz
        zusatz$ = zusatz$ + " ..."
    End If
    If (ww.loesch) Then
        iItalic% = True
    End If

'    If ((BestellAnzeige% = 2) And (IstDirektLief%) And (ww.bm = 0)) Then lBack& = wpara.FarbeDunklerBereich ' FarbeGray&        'vbGrayText
    If (BackDunkel%) Then lBack& = wpara.FarbeDunklerBereich ' FarbeGray&        'vbGrayText

    .FillStyle = flexFillRepeat
    .col = 0
    .ColSel = .Cols - 1
    .CellFontItalic = iItalic%
    .CellBackColor = lBack&
    .FillStyle = flexFillSingle

    iAngebot% = ww.angebot

    'damit Berechnung der Werte richtig
    If (ww.zugeordnet = "N") Then
        lBack& = wpara.FarbeDunklerBereich  ' FarbeGray& ' vbGrayText
    End If

'    .col = 7
'    If (ww.IstSchwellArtikel) Then
'        .CellForeColor = vbBlue
'    Else
'        .CellForeColor = .ForeColor
'    End If
    
    .col = 1
    .CellFontName = "Symbol"
    If (FirstRed%) Then
        .CellBackColor = vbRed
    End If


    row% = .row
    .TextMatrix(row%, 0) = ww.pzn
    .TextMatrix(row%, 1) = KontrollChar$
    .TextMatrix(row%, 2) = ArtName$
    .TextMatrix(row%, 3) = ArtMenge$
    .TextMatrix(row%, 4) = ArtMeh$
    .TextMatrix(row%, 5) = Abs(ww.bm)
    .TextMatrix(row%, 6) = nm$
    .TextMatrix(row%, 7) = Format(Rabatt!, "0.00")
    .TextMatrix(row%, 8) = LiefName$
    If (ww.ssatz = 0) Then
        h$ = " "
    Else
        h$ = Format(ww.PosLag, "0")
    End If
    If (FremdPznOk%) Then
        FremdPznRec.Seek "=", ww.pzn
        If (FremdPznRec.NoMatch = False) Then
            h$ = "*" + h$
        End If
    End If
    .TextMatrix(row%, 9) = h$
    
    
    .TextMatrix(row%, 10) = zusatz$
    .TextMatrix(row%, 11) = ActKontrollen$  'bek.ActKontrolle
    .TextMatrix(row%, 12) = Val(ww.AbholNr)
    .TextMatrix(row%, 13) = Format(EK#, "0.00")
    .TextMatrix(row%, 14) = ww.alt

    If (iAngebot% And &H1) Then
        .TextMatrix(row%, 15) = " "
    ElseIf (iAngebot% And &H4) Then
        .TextMatrix(row%, 15) = "!!"
    Else
        .TextMatrix(row%, 15) = "!"
    End If

    .TextMatrix(row%, 16) = VonWo$
    .TextMatrix(row%, 17) = ww.Lief
    .TextMatrix(row%, 18) = actzuordnung$
    .TextMatrix(row%, 19) = SRT$
    If (ProgrammChar$ = "2") Or (lBack& = vbButtonFace) Then
        .TextMatrix(row%, 20) = "*"
    Else
        .TextMatrix(row%, 20) = " "
    End If
    .TextMatrix(row%, 21) = ww.BekLaufNr
    .TextMatrix(row%, 22) = ww.aktivlief
    .TextMatrix(row%, 23) = ww.loesch
    .TextMatrix(row%, 24) = ww.wg
    .TextMatrix(row%, 25) = bmo!
    .TextMatrix(row%, 26) = ww.zr
    .TextMatrix(row%, 27) = ww.IstSchwellArtikel
    
    .TextMatrix(row%, 34) = ZusInfo$

    If ((BestellAnzeige% = 2) And (IstDirektLief%)) Then
        If (ww.ssatz > 0) Then
            .TextMatrix(row%, 28) = ass.opt
            .TextMatrix(row%, 29) = ass.bm
        Else
            .TextMatrix(row%, 28) = 0
            .TextMatrix(row%, 29) = 0
        End If
       
        If (ww.asatz > 0) Then
            ast.GetRecord (ww.asatz + 1)
            .TextMatrix(row%, 32) = ast.aep
        Else
            .TextMatrix(row%, 32) = 0
        End If
        
        SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + ww.pzn
        Set TaxeRec = TaxeDB.OpenRecordset(SQLStr$)
        If (TaxeRec.EOF = False) Then
            .TextMatrix(row%, 30) = TaxeRec!EK / 100
            If (ww.asatz <= 0) Then
                Call Taxe2ast(ww.pzn)
            End If
            .TextMatrix(row%, 31) = ast.wg
        Else
            .TextMatrix(row%, 30) = 0
            .TextMatrix(row%, 31) = 0
        End If

        If (ww.asatz > 0) Or (TaxeRec.EOF = False) Then
            .TextMatrix(row%, 33) = 1
        Else
            .TextMatrix(row%, 33) = 0
        End If
    
    End If

    .col = col%
End With

If (WirklichKeinRowColChange% = False) Then KeinRowColChange% = False

Call DefErrPop
End Sub

Sub AuslesenBestellung(AnzeigeFlag%, BereitsGelockt%, Optional AnimationFlag% = False)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("AuslesenBestellung")
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
Dim i%, j%, Max%, l%, AltLief%, ind%, row%, WasGeõndert%, AnzeigeCol%, erg%, iLief%
Dim nRufzeit%, DoppeltKontrolle%, gef%, Bm0Kontrolle%
Dim h$, h2$, wwOrg$, wwNeu$
Dim AutoDirLiefs$, ManuellDirLiefs$, IniHinweis$, h3$, EinzelHinweis$
Static IstAktiv%

If (IstAktiv%) Then Call DefErrPop: Exit Sub

frmAction.tmrAction.Enabled = False

If (BereitsGelockt% = False) Then Call ww.SatzLock(1)

ww.GetRecord (1)
'If (GlobBekMax% = bek.erstmax) And (BekartCounter% = bek.erstcounter) Then
'    If (BereitsGelockt% = False) Then Call bek.SatzUnLock(1)
'    Call DefErrPop:    Exit Sub
'End If

IstAktiv% = True

WasGeõndert% = False
Max% = ww.erstmax
AltLief% = ww.erstlief
BekartCounter% = ww.erstcounter
GlobBekMax% = Max%

AktUhrzeit% = Val(Format(Now, "HHMM"))
NaechstLiefernderLieferant% = HoleNaechstLiefernden%

Call frmAction.InitAuslesenBestellung(AnzeigeCol%, AltLief%, Max%)

Call frmAction.ErzeugeRueckrufMenu

frmAction.lstSortierung.Clear

iLief% = 0
If (BestellAnzeige% = 2) Then iLief% = Lieferant%

nRufzeit% = HoleNaechsteRufzeit%(iLief%)

h$ = ProgrammNamen$(ProgrammTyp%) + " - "
If (BestellAnzeige% = 0) Then
    h$ = h$ + "Übersicht komplett, nächster Ruf: "
ElseIf (BestellAnzeige% = 1) Then
    h$ = h$ + "Artikel zur Kontrolle, nächster Ruf: "
Else
    h$ = h$ + "Vorschau für "
End If

If (iLief% > 0) And (iLief% <= lif.AnzRec) Then
    lif.GetRecord (iLief% + 1)
    h2$ = RTrim$(lif.kurz)
Else
    h2$ = "??????"
End If
h$ = h$ + h2$ + " "

lifzus.GetRecord (iLief% + 1)
IstDirektLief% = lifzus.IstDirektLieferant

IstAutoDirektLief% = 0
If (IstDirektLief%) Then
    IstAutoDirektLief% = lifzus.IstAutoDirektLieferant
End If

DarfBestVors% = True
DirektBezugFaktRabatt# = 0#
DirektBezugValutaStellung% = 0#
If (IstDirektLief%) Then
    DarfBestVors% = lifzus.ZuordnungenAktiv
    DirektBezugFaktRabatt# = lifzus.FakturenRabatt
    If (lifzus.TempFakturenRabatt > 0) Then DirektBezugFaktRabatt# = lifzus.TempFakturenRabatt
    DirektBezugFaktRabattTyp% = lifzus.FakturenRabattTyp
    DirektBezugValutaStellung% = lifzus.ValutaStellung
    If (lifzus.TempValutaStellung > 0) Then DirektBezugValutaStellung% = lifzus.TempValutaStellung
    DirektBezugAbBm% = lifzus.DirektMindestBM
End If


If (nRufzeit% >= 0) Then
    h2$ = " (" + Format(nRufzeit% \ 100, "00") + ":" + Format(nRufzeit% Mod 100, "00") + ") "
    h$ = h$ + h2$
End If

DirektBezugsKz% = 0
If (IstDirektLief%) Then
    DirektBezugsKz% = 2
    h$ = h$ + "- Direktlieferant"
End If

If (PruefeRueckruf(iLief%)) Then h$ = h$ + " (RR)"

frmAction.lblArbeit(0) = h2$
    
frmAction.Caption = h$

If (frmAction.flxarbeit(0).Rows < 2) Then
    frmAction.flxarbeit(0).Rows = 2
End If
frmAction.flxarbeit(0).row = 0




DoppeltKontrolle% = False
For i% = 0 To AnzKontrollen% - 1
    If (UCase(RTrim(Kontrollen(i%).bedingung.wert1)) = "DOPPELTKONTROLLE") Then
        DoppeltKontrolle% = True
        Exit For
    End If
Next i%
If (DoppeltKontrolle%) Then
    With frmAction.lstSortierung
        .Clear
        For i% = 1 To Max%
            ww.GetRecord (i% + 1)
            If (ww.status = 1) Then .AddItem ww.pzn
        Next i%
        i% = 0
        Do
            If (i% >= .ListCount) Then Exit Do
            
            j% = i% + 1
            gef% = False
            Do
                If (j% >= .ListCount) Then Exit Do
                If (.List(j%) = .List(i%)) Then
                    gef% = True
                    .RemoveItem (j%)
                Else
                    Exit Do
'                    j% = j% + 1
                End If
            Loop
            If (gef%) Then
                i% = i% + 1
            Else
                .RemoveItem (i%)
            End If
        Loop
'        For i% = 0 To .ListCount - 1
'            Debug.Print .List(i%)
'        Next i%
    End With
End If


'frmAction.lstSortierung.Clear


AutoDirLiefs$ = ""
ManuellDirLiefs$ = ""

AnzBestellArtikel% = 0

For i% = 1 To Max%
    ww.GetRecord (i% + 1)
    
    If (ww.status = 1) Then
    
        wwOrg$ = ww.RawData
        
        If (Lieferant% <> AltLief%) Then
            ww.zukontrollieren = Chr$(0)
        End If

        Call EinzelSatz(0, h$)
    
        Bm0Kontrolle% = False
        If (ww.Lief > 0) Then
            lifzus.GetRecord (ww.Lief + 1)
            If (lifzus.IstDirektLieferant) Then
                h$ = Format(ww.Lief, "000") + "-"
                If (ww.DirektTyp = 2) Then
                    ind% = InStr(ManuellDirLiefs$, h$)
                    If (ind% <= 0) Then ManuellDirLiefs$ = ManuellDirLiefs$ + h$ + "0,"
                Else
                    ind% = InStr(AutoDirLiefs$, h$)
                    If (ind% <= 0) Then
                        AutoDirLiefs$ = AutoDirLiefs$ + h$ + "0,"
                        ind% = InStr(AutoDirLiefs$, h$)
                    End If
                    h2$ = Mid$(AutoDirLiefs$, ind% + 4, 1)
                    If (ww.zukontrollieren = "1") Then
                        h2$ = "1"
                    ElseIf (ww.zukontrollieren = "2") And (h2$ <> "1") Then
                        h2$ = "2"
                    End If
                    Mid$(AutoDirLiefs$, ind% + 4, 1) = h2$
                End If
                
                If (Abs(ww.bm) = 0) Then Bm0Kontrolle% = True
            End If
        End If
    
        With frmAction.flxarbeit(0)
            If ((BestellAnzeige% = 0) And ((Bm0Anzeigen%) Or (Abs(ww.bm) > 0))) Or _
               ((BestellAnzeige% = 1) And ((ww.zukontrollieren = "1") Or (ww.zukontrollieren = "2")) And Not (Bm0Kontrolle%)) Or _
               ((BestellAnzeige% = 2) And (ww.zugeordnet = "J")) Then
                
                AnzBestellArtikel% = AnzBestellArtikel% + 1
    
                If (.row >= (.Rows - 1)) Then .Rows = .Rows + 1
                .row = .row + 1
                Call ZeigeBestellZeile(i%, Bm0Kontrolle%)
            End If
        End With
    
    
        wwNeu$ = ww.RawData
        If (wwOrg$ <> wwNeu$) And (ww.aktivlief = 0) Then
            ww.PutRecord (i% + 1)
            WasGeõndert% = True
        End If
    End If
Next i%

AltLief% = Lieferant%
ww.erstlief = AltLief%
ww.PutRecord (1)

If (BereitsGelockt% = False) Then Call ww.SatzUnLock(1)


'''''''''''
Do
    ind% = InStr(ManuellDirLiefs$, "-")
    If (ind% > 0) Then
        h$ = Left$(ManuellDirLiefs$, ind%)
        ManuellDirLiefs$ = Mid$(ManuellDirLiefs$, ind% + 3)

        ind% = InStr(AutoDirLiefs$, h$)
        If (ind% > 0) Then
            h2$ = AutoDirLiefs$
            AutoDirLiefs$ = Left$(h2$, ind% - 1) + Mid$(h2$, ind% + 6)
        End If
    Else
        Exit Do
    End If
Loop
'
'
IniHinweis$ = HoleIniString$("DirektBezugHinweis")
If (IniHinweis$ <> "") Then
    h2$ = ""
    Do
        If (IniHinweis$ = "") Then Exit Do

        ind% = InStr(IniHinweis$, "-")
        If (ind% > 0) Then
            h$ = Left$(IniHinweis$, ind%)
            ind% = InStr(AutoDirLiefs$, h$)
            If (ind% > 0) Then
                EinzelHinweis$ = Left$(IniHinweis$, 14)
                h3$ = Mid$(AutoDirLiefs$, ind% + 4, 1)
                Mid$(EinzelHinweis$, 9, 1) = h3$
'                    If (h3$ <> "0") Then
                    h2$ = h2$ + EinzelHinweis$
'                    End If
            End If
            IniHinweis$ = Mid$(IniHinweis$, 15)
        Else
            Exit Do
        End If
    Loop
    IniHinweis$ = h2$

    Call SpeicherIniString%("DirektBezugHinweis", IniHinweis$)
End If
''''''''''''

With frmAction.flxarbeit(0)
    If (.row = 0) Then
        .Rows = 2
        .TextMatrix(1, 0) = "XXXXXXX"
        .TextMatrix(1, 1) = " "
        .TextMatrix(1, 2) = "Keine anzuzeigenden Bestelldaten gespeichert !"
        For i% = 3 To (.Cols - 1)
            .TextMatrix(1, i%) = ""
        Next i%
    Else
        .Rows = .row + 1
    End If
    .row = 1
    .col = 19
    .RowSel = .Rows - 1 ' AnzBestellArtikel%
    .ColSel = 19
    .Sort = 5
End With

Call ActProgram.ZeigeWerte

Call frmAction.ExitAuslesenBestellung(AnzeigeCol%, Max%)

zTabelleAktiv% = True
zManuellAktiv% = False

IstAktiv% = False

Call frmAction.ErzeugeGesendeteMenu
Call frmAction.ErzeugeSchwellwerteMenu
Call frmAction.ErzeugeDirektbezugMenu

Call ShowSchwellwertWarnung
Call ShowDirektBezugHinweis

frmAction.tmrAction.Enabled = True

Call DefErrPop
End Sub

Sub AltRClick()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("AltRClick")
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
                  
'Sub AuslesenPalette()
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Call DefErrFnc("AuslesenPalette")
'Call DefErrMod(DefErrModul)
'On Error GoTo DefErr
'GoTo DefErrEnd
'DefErr:
'Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
'Case vbRetry
'  Resume
'Case vbIgnore
'  Resume Next
'End Select
'End
'DefErrEnd:
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Dim i%, j%, Lief%, Max%, l%, AltLief%, ind%, row%
'Dim Preis#, ZeilenWert#, EK#
'Dim h$, h2$, SRT$, autox$, AktPzn$, tx$
'Dim Lac$
'
'lifzus.GetRecord (Lieferant% + 1)
'DirektBevorratungsZeitraum% = lifzus.TempBevorratungsZeitraum
'If (DirektBevorratungsZeitraum% = 0) Then DirektBevorratungsZeitraum% = lifzus.BevorratungsZeitraum
'If (DirektBevorratungsZeitraum% = 0) Then DirektBevorratungsZeitraum% = para.BestellPeriode
'
'ww.GetRecord (1)
'Max% = ww.erstmax
'
'For i% = 1 To Max%
'    ww.GetRecord (i% + 1)
'
'    If (ww.status = 1) And (ww.Lief = Lieferant%) And (ww.loesch = 0) And (ww.aktivlief = 0) Then
'
'        FabsErrf% = ass.IndexSearch(0, ww.pzn, FabsRecno&)
'        If (FabsErrf% = 0) Then
'            ww.ssatz = CInt(FabsRecno&)
'            ass.GetRecord (ww.ssatz + 1)
'            ww.loesch = 1
'            ww.PutRecord (i% + 1)
'            Call ArtikelInPalette(ww.pzn)
'        End If
'
'    End If
'Next i%
'
'Call DefErrPop
'End Sub

Sub DirektBezugSenden()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("DirektBezugSenden")
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
Dim ind%, ret%, MitAusdruck%
Dim h$, h2$

frmDirektBezug.Show 1
If (DirektBezugErg$ = "") Then
    Call UpdateBekartDat(Lieferant%, False)
Else
    h$ = DirektBezugErg$
    If (Left$(h$, 1) = "*") Then
        MitAusdruck% = True
        h$ = Mid$(h$, 2)
    End If
    If (Left$(h$, 5) = "Modem") Then
        If (Wbestk2ManuellSenden%) Then
            Wbestk2ManuellVorbereitung% = True
            AutomaticSend% = False
            ManuellSendung% = True
            LeerAuftrag% = False
            frmSenden.Show 1
            Wbestk2ManuellVorbereitung% = False
            Call AuslesenBestellung(True, False)
            ret% = False
        Else
            ret% = ModemAktivieren%
            If (ret%) Then
                AutomaticSend% = False
                ManuellSendung% = True
                LeerAuftrag% = False
                UebertragungOk% = False
                frmSenden.Show 1
                ret% = UebertragungOk%
            Else
                Call iMsgBox("Modem momentan belegt !")
            End If
        End If
'    ElseIf (Left$(h$, 5) = "eMail") Then
'        Call DirektBezugMail
'    ElseIf (Left$(h$, 5) = "Compu") Then
    Else
        BlindBestellung% = True
        Call SucheSendeArtikel
        Call SaetzeVorbereiten
        Call SaetzeSenden
        Call UpdateBekartDat(Lieferant%, True)
        BlindBestellung% = False
        
        ret% = True
        MitAusdruck% = True
    End If
    
    If (ret%) Then
        lifzus.GetRecord (Lieferant% + 1)
        lifzus.TempBevorratungsZeitraum = 0
        lifzus.TempValutaStellung = 0
        lifzus.TempFakturenRabatt = 0
        lifzus.PutRecord (Lieferant% + 1)
        
        If (MitAusdruck%) Then Call DirektBezugAusdruck(True)
        
        Call DirektBezugVerwerfen(False)
    End If
    
End If

Call DefErrPop
End Sub

Function PruefeRueckruf%(iLief%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("PruefeRueckruf%")
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
Dim i%, ind%
Dim l&
Dim SollRR$, h$

h$ = Space$(100)
l& = GetPrivateProfileString(INI_SECTION, "Rueckrufe", " ", h$, 101, INI_DATEI)
SollRR$ = Trim(Left$(h$, l&))
    
h$ = Format(iLief%, "000")
ind% = InStr(SollRR$, h$)
PruefeRueckruf% = (ind% > 0)

Call DefErrPop
End Function


Sub ReduceManuelleDirektBestellung()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("ReduceManuelleDirektBestellung")
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
Dim i%, WasGeändert%, Max%, nUeber%, bm1%, mm1%
Dim SQLStr$
            
WasGeändert% = False
Call ww.SatzLock(1)
ww.GetRecord (1)

Max% = ww.erstmax
BekartCounter% = ww.erstcounter

For i% = 1 To Max%
    nUeber% = 0
    
    ww.GetRecord (i% + 1)
    If (ww.status = 1) And (ww.Lief = Lieferant%) Then
        'nur wenn über GH beziehbar !
        SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + ww.pzn
        Set TaxeRec = TaxeDB.OpenRecordset(SQLStr$)
        If (TaxeRec.EOF = False) Then
            If (TaxeRec!UeberGH = 1) Then
                nUeber% = True
            End If
        End If
        If (nUeber%) Then
        Else
            bm1% = 0
            FabsErrf% = ass.IndexSearch(0, ww.pzn, FabsRecno&)
            If (FabsErrf% = 0) Then
                ass.GetRecord (FabsRecno& + 1)
                mm1% = ass.vmm
                If (mm1% <= 0) Then mm1% = ass.MM
                If ((mm1% >= 1) And (ass.PosLag <= Int(mm1% / 2))) Then
                    bm1% = mm1% - ass.PosLag
                End If
                If (mm1% = 0) And (ass.PosLag = 0) Then
                    bm1% = 1
                End If
            End If
            
            If (bm1% > 0) Then
                ww.bm = bm1%
                ww.nm = 0
                ww.Lief = 0
                ww.nnart = 0
                ww.NNAEP = 0#
                ww.zr = 0!
                ww.angebot = ww.angebot And (&H34 Xor &HFFFF)
                
                ww.PutRecord (i% + 1)
                WasGeändert% = True
            End If
        End If
    End If
Next i%

If (WasGeändert%) Then
    BekartCounter% = (BekartCounter% + 1) Mod 100
    ww.erstcounter = BekartCounter%
    ww.PutRecord (1)
End If

Call ww.SatzUnLock(1)

Call DefErrPop
End Sub

Sub MakePartnerLief(iLief%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("MakePartnerLief")
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
Dim i%, rInd%

frmAction.tmrAction.Enabled = False

Call ww.SatzLock(1)

rInd% = SucheFlexZeile%(True)
If (rInd% > 0) Then
    ww.Lief = iLief%
    ww.zukontrollieren = Chr$(0)
    ww.fixiert = Chr$(0)
    ww.IstSchwellArtikel = 0
    ww.PutRecord (rInd% + 1)
    Call AuslesenDateiSatz(rInd%, True)
    Call ErhoeheCounter
    Call ErhoehePartnerCounter
End If

ww.SatzUnLock (1)

frmAction.tmrAction.Enabled = True

Call DefErrPop
End Sub
