VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsWarenÜber"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Const INI_SECTION = "WÜ"
Const INFO_SECTION = "Infobereich WÜ"

Private Type EwTextErstSatz
    Max As Single
    rest As String * 124
End Type
Private Type EwTextStruct
    pzn As String * 7
    Nr As String * 1
    NextSatz As Single
    Zeile As String * 60
    rest As String * 56
End Type
Dim EwTextErst As EwTextErstSatz
Dim EwTextRec As EwTextStruct


Dim WumsatzRec() As WumsatzStruct
Dim ASumsatzRec() As WumsatzStruct  'AbgabeSchlüssel...
Dim LagerListePzn$()
Dim DokuPflichtPzn$()
Dim IntStreichungRow%()

Dim ParentForm As Form
Dim ParentToolbar As clsToolbar
Dim ParentInfo As clsInfoBereich
Dim ParentBereich As clsOpBereiche

Dim ActLifDat$

Dim PreisKalkModus%

Dim SuperEdit%

Dim RetourFussTxt$, RetourKopfTxt$
Dim RetourFussHoehe%

Dim RowaDateiLen&

Dim NLbeleg$
Dim NLbelegDatum%
Dim OrgBeleg$
Dim OrgBelegDatum%

Private Const DefErrModul = "WARENÜBERNAHME.CLS"

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
Dim i%, erg%
Dim h$

Set ParentForm = iForm
Set ParentToolbar = iToolbar
Set ParentInfo = iInfo
Set ParentBereich = iBereich

With ParentForm
    Call ParentToolbar.InitToolbar(ParentForm, INI_DATEI, INI_SECTION, "2314241625151717")
    
    .cmdToolbar(0).ToolTipText = "ESC Zurück: Zurückschalten auf vorige Bildschirmmaske"
    .cmdToolbar(1).ToolTipText = "F2 Erfassen zusätzlicher Artikel"
    .cmdToolbar(2).ToolTipText = "F3 Interner Programmwechsel"
    .cmdToolbar(3).ToolTipText = "F4 Quittieren der Nachricht"
    .cmdToolbar(4).ToolTipText = "F5 Irrtum"
    .cmdToolbar(5).ToolTipText = "F6 Drucken der Lieferung"
    .cmdToolbar(6).ToolTipText = "F7 Aktualisieren"
    .cmdToolbar(7).ToolTipText = "F8 Zusatztext"
    .cmdToolbar(8).ToolTipText = "F9 Abmelden"
    .cmdToolbar(9).ToolTipText = "shift+F2 Bestell-Status"
    .cmdToolbar(10).ToolTipText = "shift+F3 aktuelle Lieferung wechseln"
    .cmdToolbar(11).ToolTipText = "shift+F4 Angebote"
    .cmdToolbar(12).ToolTipText = "shift+F5 Durchgriff auf Statistik-Anzeige"
    .cmdToolbar(13).ToolTipText = "shift+F6 Fertigmachen der Lieferung"
    .cmdToolbar(14).ToolTipText = "shift+F7 Stift-Erfassung"
    .cmdToolbar(15).ToolTipText = "shift+F8 Retouren erfassen"
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
        .Cols = 27
        .Rows = 0
        .Rows = 2
        .FixedRows = 1
'        .FormatString = "<PZN|^ |<Name|>Menge|^Meh|>BM|>NR|^Lieferant|>POS|<Zusatz|<ActBedingung|||?|^!|v"
        .FormatString = "<PZN|^ |<Name|>Menge|^Meh|>BM|>NR|>RM|>LM|>±RM|>±LM|^Verfall|>AEP|^!|^v|A"
        .FormatString = .FormatString + "|||||||||||<±"
    End With
        
    .mnuBearbeitenInd(10).Caption = "&Lieferung"
    .mnuBearbeitenInd(13).Caption = "L&ieferung fertig"
    .mnuBearbeitenInd(14).Caption = "St&ift-Erfassung"
    .mnuBearbeitenInd(15).Caption = "&Retouren erfassen"
    
    .mnuBearbeitenZusatzInd(0).Caption = "Artikel zurü&ck in Bestellung"
    .mnuBearbeitenZusatzInd(0).Enabled = True
    
    Load .mnuBearbeitenZusatzInd(1)
    .mnuBearbeitenZusatzInd(1).Caption = "Lieferung zurüc&k in Bestellung"
    .mnuBearbeitenZusatzInd(1).Enabled = True
        
    Load .mnuBearbeitenZusatzInd(2)
    .mnuBearbeitenZusatzInd(2).Caption = "Zusätzlichen &Umsatz eingeben"
    .mnuBearbeitenZusatzInd(2).Enabled = True
        
    Load .mnuBearbeitenZusatzInd(3)
    .mnuBearbeitenZusatzInd(3).Caption = "&Beleg-Daten"
        
    Load .mnuBearbeitenZusatzInd(4)
    .mnuBearbeitenZusatzInd(4).Caption = "Alle LM be&stätigen"
        
    If (para.Land = "A") Then
        Load .mnuBearbeitenZusatzInd(5)
        .mnuBearbeitenZusatzInd(5).Caption = "Aus&druck Kontroll-Blätter"
    End If
        
    .mnuRahmenAnzeigen.Checked = wpara.FarbeRahmen
'    .mnuAktZeile.Checked = wpara.FarbeAktZeile
    .mnuZusatzInfo.Checked = ArtikelStatistik%
    
    .mnuBearbeitenInd(MENU_SF3).Enabled = True
    .mnuBearbeitenInd(MENU_SF6).Enabled = True
    .cmdToolbar(10).Enabled = .mnuBearbeitenInd(MENU_SF3).Enabled
    .cmdToolbar(13).Enabled = .mnuBearbeitenInd(MENU_SF6).Enabled
    
    If (Dp04Ok%) Then
        .mnuBearbeitenInd(MENU_SF7).Enabled = True
    Else
        .mnuBearbeitenInd(MENU_SF7).Enabled = False
    End If
    .cmdToolbar(14).Enabled = .mnuBearbeitenInd(MENU_SF7).Enabled
    
    .mnuBearbeitenInd(MENU_SF8).Enabled = True
    .cmdToolbar(15).Enabled = .mnuBearbeitenInd(MENU_SF8).Enabled
    
    .mnuEinzelBewertung.Enabled = False
    
    .mnuRueckrufe.Enabled = False
    .mnuLiefStammdaten.Enabled = True
    .mnuPosPartner.Enabled = True
    .mnuAepKalkulation.Enabled = True
    .mnuTeilAbschluss.Enabled = True
    
    .mnuEditAufteilung.Enabled = False
    
    If (RowaAktiv%) Then
        .tmrRowa.Interval = 2000
    Else
        .tmrRowa.Interval = 5000
    End If
    
    .tmrRowa.Enabled = True
    
    If (SeriellScannerOk%) Then
        erg% = OpenCom(ParentForm, SeriellScannerParam$)
        If (erg% = False) Then Call MsgBox("Seriell-Scanner nicht verfügbar!", vbExclamation)
    End If
    
    .flxEinzelBewertung.Visible = False
    
    IstRowaWu% = 0

End With

Call DefErrPop
End Sub

Sub AuslesenWu(Optional RabattEinlesen% = False)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("AuslesenWu")
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
Dim i%, j%, IstOk%, Erst%, Max%, Lief%, row%, AnzCol%, AltLastChronologisch%
Dim LifDat$, h$, h2$, h3$, SRT$
Dim FrmActionTmrRowaEnabled%
Static IstAktiv%

If (IstAktiv%) Then Call DefErrPop: Exit Sub

'If (BereitsGelockt% = False) Then Call wu.SatzLock(1)

With frmAction.tmrRowa
    FrmActionTmrRowaEnabled% = .Enabled
    .Enabled = False
End With

ww.GetRecord (1)

IstAktiv% = True
RowaDateiLen& = 0&

Max% = ww.erstmax

If (Max% > 20) Then
    Call StartAnimation(frmAction, "WÜ wird eingelesen ...")
End If

EinzelneWu% = True
IstAltLast% = False
AltLastChronologisch% = False
WuInDm% = -1
IstDirektLief% = False
DirektBezugFaktRabatt# = 0#

If (RabattEinlesen%) Then
    ActBeleg$ = Space$(10)
    ActBelegDatum% = 0
End If

Call ParentForm.ToolbarUpdateFlag(False)

With ParentForm.flxarbeit(0)
    .redraw = False

    h$ = WuLifDat$
    If (Len(h$) < 3) Then
        If (Len(h$) < 2) Then EinzelneWu% = False
        h3$ = Left$(h$, 1)
        h$ = Mid$(h$, 2)
        If (h3$ = Chr$(0)) Then
            h2$ = "Alle Lieferanten"
        Else
            Lief% = Asc(h3$)
            lif.GetRecord (Lief% + 1)
            h2$ = RTrim$(lif.Name(0))
            
            lifzus.GetRecord (Lief% + 1)
            IstDirektLief% = lifzus.IstDirektLieferant
            If (IstDirektLief%) Then
                DirektBezugFaktRabatt# = lifzus.FakturenRabatt
                If (lifzus.TempFakturenRabatt > 0) Then DirektBezugFaktRabatt# = lifzus.TempFakturenRabatt
            End If
        End If
        If (h$ = "") Then
            h2$ = h2$ + " (alle Lieferungen)"
            IstAltLast% = True
        ElseIf (h$ = " ") Then
            h2$ = h2$ + " (alle aktuellen Lieferungen)"
        ElseIf (h$ = "*") Then
            h2$ = h2$ + " (alle Altlasten)"
            IstAltLast% = True
        Else
            IstAltLast% = True
            If (UCase(h$) = "G") Then
                h2$ = h2$ + " (alle Gutschriften)"
            ElseIf (UCase(h$) = "L") Then
                h2$ = h2$ + " (alle Nachlieferungen)"
            ElseIf (UCase(h$) = "R") Then
                h2$ = h2$ + " (alle Retourwaren)"
            Else
                h2$ = h2$ + " (alle Nachverrechnungen)"
            End If
            If (h$ = LCase(h$)) Then
                AltLastChronologisch% = True
                h2$ = Left$(h2$, Len(h2$) - 1) + ", chronologisch)"
            End If
        End If
    Else
        Lief% = Asc(Mid$(h$, 2, 1))
        lif.GetRecord (Lief% + 1)
        h2$ = RTrim$(lif.Name(0))
        h2$ = h2$ + " ("
        
            
        lifzus.GetRecord (Lief% + 1)
        IstDirektLief% = lifzus.IstDirektLieferant
        If (IstDirektLief%) Then
            DirektBezugFaktRabatt# = lifzus.FakturenRabatt
            If (lifzus.TempFakturenRabatt > 0) Then DirektBezugFaktRabatt# = lifzus.TempFakturenRabatt
        End If
        
        Erst% = True
        Do
            If (Left$(h$, 1) = "@") Then
                LifDat$ = Mid$(h$, 2, 11)
                If (Mid$(h$, 13, 1) <> " ") Then IstAltLast% = True
                h$ = Mid$(h$, 14)
            
                If (Erst% = False) Then
                    h2$ = h2$ + ", "
                End If
                h3$ = Mid$(LifDat$, 2, 6)
                h2$ = h2$ + Left$(h3$, 2) + "." + Mid$(h3$, 3, 2) + "." + Right$(h3$, 2)
                h3$ = Mid$(LifDat$, 8, 4)   'Format(CVI(Right$(LifDat$, 2)), "0000")
                h2$ = h2$ + " " + Left$(h3$, 2) + ":" + Mid$(h3$, 3)
                Erst% = False
            Else
                Exit Do
            End If
        Loop
        h2$ = h2$ + ")"
        
        If (RabattEinlesen%) Then
            Call rabtab.TabelleEinlesen(Lief%)
        End If
    End If
    
    ParentForm.lblArbeit(0) = h2$
    If (IstAltLast%) Then
        ParentForm.Caption = "Altlasten" + " - " + h2$
    Else
        ParentForm.Caption = ProgrammNamen$(ProgrammTyp%) + " - " + h2$
    End If
    
    
    AnzCol% = .col
    If (.Rows > 1) And (AnzeigeLaufNr& = 0) Then
        row% = .row
        AnzeigeLaufNr& = Val(RTrim$(.TextMatrix(row%, 20)))
    End If
    Call ParentForm.HighlightZeile(True)
    
'    If (.Rows < 2) Then
        .Rows = 2
'    End If
    .row = 0


    AnzBestellArtikel% = 0

    ActLifDat$ = ""
    
    For i% = 1 To Max%
        ww.GetRecord (i% + 1)
    
        If (ww.status = 2) Then
    
            If (ww.Lief = 0) Then
                ww.Lief = 1
                ww.PutRecord (i% + 1)
            End If
            
            LifDat$ = Chr$(ww.Lief) + ww.WuBestDatum + Format(CVI(ww.WuBestZeit), "0000")
            
            IstOk% = True
            If (Len(WuLifDat$) <= 2) Then
                h3$ = Left$(WuLifDat$, 1)
                h$ = Mid$(WuLifDat$, 2)
                If (h3$ <> Chr$(0)) And (h3$ <> Left$(LifDat$, 1)) Then
                    IstOk% = False
                ElseIf (h$ = " ") And (ww.IstAltLast) Then
                    IstOk% = False
                ElseIf (h$ = "*") And (ww.IstAltLast = 0) Then
                    IstOk% = False
                ElseIf (UCase(h$) = "G") And ((ww.IstAltLast = 0) Or (ww.WuNeuRm >= 0)) Then
                    IstOk% = False
                ElseIf (UCase(h$) = "L") And ((ww.IstAltLast = 0) Or (ww.WuNeuLm <= 0)) Then
                    IstOk% = False
                ElseIf (UCase(h$) = "R") And ((ww.IstAltLast = 0) Or (ww.WuNeuLm >= 0)) Then
                    IstOk% = False
                ElseIf (UCase(h$) = "V") And ((ww.IstAltLast = 0) Or (ww.WuNeuRm <= 0)) Then
                    IstOk% = False
                End If
'            ElseIf (Len(WuLifDat$) = 1) And (WuLifDat$ <> Left$(LifDat$, 1)) Then
'                IstOk% = False
            ElseIf (Len(WuLifDat$) > 2) And (InStr(WuLifDat$, "@" + LifDat$) = 0) Then
                IstOk% = False
            ElseIf (Asc(ww.WuBestDatum) = 0) Then
                IstOk% = False
            End If
            
            If (IstOk%) Then
                If (ActLifDat$ = "") Then ActLifDat$ = LifDat$
                
                If (EinzelneWu%) And (WuInDm% < 0) Then
                    WuInDm% = 0
                    If (Right$(ww.WuBestDatum, 2) = "01") And (Year(Now) > 2001) Then WuInDm% = 1
                End If
                
'                If (ww.WuFertig = 3) Then IstAltLast% = True
                
'                If (EinzelneWu%) And (ww.WuLm < 0) And (ActBeleg$ = Space$(10)) Then
                If (EinzelneWu%) And (RabattEinlesen%) And (ww.LmStatus) And (ActBeleg$ = Space$(10)) Then
                    ActBeleg$ = ww.WuBeleg
                    ActBelegDatum% = ww.WuBelegDatum
                End If
                
                AnzBestellArtikel% = AnzBestellArtikel% + 1
                If (.row >= (.Rows - 1)) Then .Rows = .Rows + 1
                .row = .row + 1
                Call ZeigeWuZeile
                .TextMatrix(.row, 16) = LifDat$
                
                h$ = " "
'                If (EinzelneWu%) And (IstAltLast% = False) And (WuLifDat$ <> "") And (ww.WuLm < 0) Then
                If (EinzelneWu%) And (IstAltLast% = False) And (WuLifDat$ <> "") Then
                    If (ww.LmStatus) Then
                        h$ = Format(ww.WuStatus, "0")
                    Else
                        h$ = h$ + CheckWuSortierung$
                    End If
                ElseIf (AltLastChronologisch%) Then
                    h2$ = ww.WuBestDatum
                    h$ = Right$(h2$, 2) + Mid$(h2$, 3, 2) + Left$(h2$, 2)
                    h2$ = Format(CVI(ww.WuBestZeit), "0000")
                    h$ = h$ + h2$
                End If
                
'                h$ = h$ + Left$(ww.txt, 6)
                h$ = h$ + ww.txt
                
                SRT$ = h$ + Format(i%, "00000")
                .TextMatrix(.row, 18) = SRT$
            End If
        End If
    Next i%
    
    If (ActLifDat$ = "") Then ActLifDat$ = Mid$(WuLifDat$, 2, 11)
End With


'If (BereitsGelockt% = False) Then Call wu.SatzUnLock(1)

With ParentForm.flxarbeit(0)
'    If (IstAltLast%) Then
'        ParentForm.Font.Bold = True
'        .ColWidth(9) = ParentForm.TextWidth("999")
'        .ColWidth(10) = ParentForm.TextWidth("999")
'        ParentForm.Font.Bold = False
'        .ColWidth(9) = .ColWidth(9) + ParentForm.TextWidth("X")
'        .ColWidth(10) = .ColWidth(10) + ParentForm.TextWidth("X")
'    Else
'        .ColWidth(9) = 0
'        .ColWidth(10) = 0
'    End If

    If (.row = 0) Then
        .Rows = 2
        .TextMatrix(1, 0) = "XXXXXXX"
        .TextMatrix(1, 1) = " "
        .TextMatrix(1, 2) = "Keine passende Warenübernahme gespeichert !"
        For i% = 3 To (.Cols - 1)
            .TextMatrix(1, i%) = ""
        Next i%
    Else
        .Rows = .row + 1
    End If
    .row = 1
    .col = 18
    .RowSel = .Rows - 1 ' AnzBestellArtikel%
    .ColSel = 18
    .Sort = 5
    
'    Debug.Print
'    For i% = 1 To .Rows - 1
'        .row = i%
'        .col = 1
'        Debug.Print Str$(i%); .CellFontName
'    Next i%
    
    .row = 1
    If (AnzeigeLaufNr& > 0) Then
        For i% = 1 To .Rows - 1
            If (Val(.TextMatrix(i%, 20)) = AnzeigeLaufNr&) Then
                .row = i%
                Exit For
            End If
        Next i%
        AnzeigeLaufNr& = 0
        
        If (NLbeleg$ <> "") Then
            OrgBeleg$ = ActBeleg$
            OrgBelegDatum% = ActBelegDatum%
            ActBeleg$ = NLbeleg$
            ActBelegDatum% = NLbelegDatum%
            NLbeleg$ = ""
        End If
    End If

    If (.row < .TopRow) Then
        .TopRow = .row
    Else
        If (.row >= (.TopRow + ParentBereich.ArbeitAnzZeilen - 2)) Then
            .TopRow = .row - ParentBereich.ArbeitAnzZeilen + 2
        End If
'        While ((.row - .TopRow) >= (ParentBereich.ArbeitAnzZeilen - 1))
'            .TopRow = .TopRow + 1
'        Wend
    End If

    .col = AnzCol%
    If (.col = 0) Or (ErstAuslesen%) Or (RabattEinlesen%) Then
        If (IstAltLast%) Then
            .col = 10
        Else
            .col = 8
        End If
        ErstAuslesen% = False
    End If
    .ColSel = .col

    Call ParentForm.HighlightZeile
    
    Call EchtKurzInfo
End With

Call ZeigeWerte

ParentForm.flxarbeit(0).redraw = True

If (Max% > 20) Then
    Call StopAnimation(frmAction)
End If

IstAktiv% = False

If (EinzelneWu% And RabattEinlesen% And Not (IstAltLast%)) Then
    BelegModus% = 0
    frmWuBeleg.Show 1
End If

If (EinzelneWu%) Then
    ParentForm.mnuBearbeitenZusatzInd(3).Enabled = True
    ParentForm.mnuBearbeitenZusatzInd(4).Enabled = True
Else
    ParentForm.mnuBearbeitenZusatzInd(3).Enabled = False
    ParentForm.mnuBearbeitenZusatzInd(4).Enabled = False
End If

If (IstAltLast%) Then
    ParentForm.cmdAltR.Enabled = False
    ParentForm.mnuBearbeitenInd(MENU_F2).Enabled = False
Else
    ParentForm.cmdAltR.Enabled = True
    ParentForm.mnuBearbeitenInd(MENU_F2).Enabled = True
End If
ParentForm.cmdToolbar(1).Enabled = ParentForm.mnuBearbeitenInd(MENU_F2).Enabled


Call ShowSchwellwertWarnung
Call ShowDirektBezugHinweis

With frmAction.tmrRowa
     .Enabled = FrmActionTmrRowaEnabled%
End With

Call DefErrPop
End Sub

Sub ZeigeWuZeile()
Dim row%, col%, j%, l%, Lief%
Dim EK#
Dim lBack&
Dim h$, h2$, zeit$
Dim ArtName$, ArtMenge$, ArtMeh$, KontrollChar$, LmBest$, Lac$, ZusInfo$
Dim ArtikelKz As Byte

Lac$ = ""
If (ShuttleAktiv%) Then
    FabsErrf% = ast.IndexSearch(0, ww.pzn, FabsRecno&)
    If (FabsErrf% = 0) Then
        ast.GetRecord (FabsRecno& + 1)
        If (ast.Lac = "!") Then Lac$ = "!"
    End If
End If

With ParentForm.flxarbeit(0)
    row% = .row
    col% = .col
    
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
        For j% = l% To 2 Step -1
            h2$ = Mid$(h$, j%, 1)
            If (h2$ = " ") Then Exit For
            If (InStr(".xX", h2$) > 0) Then
                h2$ = Mid$(h$, j% - 1, 1)
                If (InStr("0123456789", h2$) <= 0) Then
                    Exit For
                End If
            End If
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
    
    EK# = ww.WuAEP
    

'    If (anzeigen%) Then .redraw = False

    LmBest$ = " "
    lBack& = wpara.FarbeDunklerBereich  ' FarbeGray&  'vbGrayText
    .col = 1
    .CellFontName = wpara.FontName(0)
    KontrollChar$ = " "
'    If (IstAltLast And ww.NbBereitsEditiert) Or ((IstAltLast% = False) And (ww.WuLm <= 0)) Then
    If (ww.LmStatus) Or (IstAltLast% And ((ww.WuStatus) Or (ww.RmStatus <> 0))) Then
        lBack& = vbButtonFace
        LmBest$ = "*"
        If (ww.WuStatus = 1) Then
            KontrollChar$ = "$"
            If (ww.auto = "v") Then .CellBackColor = vbRed
        ElseIf (ww.WuStatus = 2) Then
            .CellFontName = "Symbol"
            KontrollChar$ = Chr$(214)
        ElseIf (EinzelneWu%) Then
            Call ParentForm.ToolbarUpdateFlag(True)
        End If
    End If

    .FillStyle = flexFillRepeat
    .col = 0
    .ColSel = .Cols - 1
    .CellBackColor = lBack&
    .FillStyle = flexFillSingle
    
    If (Lac$ = "!") Then
        .col = 15
        .CellBackColor = vbRed
    End If
    
    
    If (IstAltLast% = False) Then
        If (ww.WuRm <> ww.WuBm) Then
            .col = 7
            .CellBackColor = vbRed
        End If
    
        If (ww.WuLm <> (ww.WuBm + ww.WuNm)) Then
            .col = 8
            .CellBackColor = vbRed
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
        
    
    
    .TextMatrix(row%, 0) = ww.pzn
    .TextMatrix(row%, 1) = KontrollChar$
    .TextMatrix(row%, 2) = ArtName$
    .TextMatrix(row%, 3) = ArtMenge$
    .TextMatrix(row%, 4) = ArtMeh$
    
    .TextMatrix(row%, 5) = ww.WuBm
    .TextMatrix(row%, 6) = ww.WuNm
    .TextMatrix(row%, 7) = ww.WuRm
    .TextMatrix(row%, 8) = ww.WuLm
        
    h$ = ""
    If (ww.WuNeuRm <> 0) And (ww.WuNeuRm <> -999) Then h$ = Str$(ww.WuNeuRm)
    .TextMatrix(row%, 9) = h$
    
    h$ = ""
    If (ww.WuNeuLm <> 0) And (ww.WuNeuLm <> -999) Then h$ = Str$(ww.WuNeuLm)
    .TextMatrix(row%, 10) = h$
    
    .TextMatrix(row%, 11) = Mid$(ww.WuAblDatum, 3)
    .TextMatrix(row%, 12) = Format(EK#, "0.00")
    
    h$ = ""
    If (ww.nnart = 2) Then h$ = "!!"
    .TextMatrix(row%, 13) = h$
    
    h$ = "m"
    If (ww.auto = "v") Then
        h$ = "v"
    ElseIf (ww.auto = "+") Then
        h$ = ""    '"a"
    End If
    .TextMatrix(row%, 14) = h$
    
    h$ = ""
    If (ww.abl = "A") Then h$ = "A"
    .TextMatrix(row%, 15) = h$ + Lac$
    
    .TextMatrix(row%, 17) = ww.WuStatus
    
    .TextMatrix(row%, 19) = LmBest$
    .TextMatrix(row%, 20) = ww.BekLaufNr
    .TextMatrix(row%, 21) = Val(ww.AbholNr)
    .TextMatrix(row%, 22) = ww.LmAnzGebucht 'ww.WuAnzBereitsGebucht

    h$ = ""
    If (ww.WuNNaepOk) Then h$ = "ok"
    .TextMatrix(row%, 23) = h$
    
    .TextMatrix(row%, 24) = ww.IstAltLast   ' ww.WuFertig
    
    h$ = Trim(ww.WuBeleg)
    If (h$ <> "") Then h$ = h$ + " (" + sdate(ww.WuBelegDatum) + ")"
    .TextMatrix(row%, 25) = h$
    
    .TextMatrix(row%, 26) = ZusInfo$
    
    .col = col%
End With

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
Dim i%, row%, iRow%, iCol%, KommWeg%
Dim pzn$, ch$, ActKontrollen$, actzuordnung$, h$

With ParentForm
    If (.flxarbeit(0).Rows < 2) Then Call DefErrPop: Exit Sub
    row% = .flxarbeit(0).row
    pzn$ = RTrim$(.flxarbeit(0).TextMatrix(row%, 0))
    If (pzn$ = "XXXXXXX") Then Call DefErrPop: Exit Sub
    If (Len(pzn$) <> 7) Then Call DefErrPop: Exit Sub
    pzn$ = KorrPzn$(pzn$)
    
    row% = .flxarbeit(0).row
    iRow% = .flxInfo(0).row
    iCol% = .flxInfo(0).col
    Call ParentInfo.ZeigeInfoBereich(pzn$, True)
    If (.ActiveControl.Name <> .flxInfo(0).Name) Then
        Call ZeigeInfoBereichAdd(.flxarbeit(0).TextMatrix(row%, 16), 0)
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
    .cmdToolbar(12).Enabled = .mnuBearbeitenInd(MENU_SF5).Enabled
    
    
    If (RTrim$(.flxarbeit(0).TextMatrix(.flxarbeit(0).row, 13)) <> "") Then
        .mnuBearbeitenInd(MENU_SF4).Enabled = True
    Else
        .mnuBearbeitenInd(MENU_SF4).Enabled = False
    End If
    .cmdToolbar(11).Enabled = .mnuBearbeitenInd(MENU_SF4).Enabled

'    KommWeg% = True
'    If (.flxarbeit(0).TextMatrix(row%, 15) = "v") Then
'        h$ = .InternerKommentar$
'        If (h$ <> "") Then
'            Call .AnzeigeKommentar(h$)
'            KommWeg% = False
'        End If
'    End If
'    If (KommWeg%) Then .picQuittieren.Visible = False

    .flxEinzelBewertung.Visible = False
    .flxDirektAufteilung.Visible = False
    If (IstAltLast% = 0) And (IstDirektLief%) Then
        If (.mnuDirektAufteilung.Checked) Then
            Call ZeigeDirektAufteilung
        End If
    End If

End With

Call DefErrPop
End Sub

Sub ZeigeInfoBereichAdd(sLifDat$, Index%)
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
Dim j%, Lief%
Dim h$, h2$, h3$

h3$ = ""
With ParentForm.flxarbeit(Index%)
    If (Val(.TextMatrix(.row, 24)) > 0) Then
        If (Val(.TextMatrix(.row, 10)) > 0) Then
            h3$ = "Nachlieferung"
        ElseIf (Val(.TextMatrix(.row, 10)) < 0) Then
            h3$ = "Retourware"
        ElseIf (Val(.TextMatrix(.row, 9)) > 0) Then
            h3$ = "Nachverrechnung"
        ElseIf (Val(.TextMatrix(.row, 9)) < 0) Then
            h3$ = "Gutschrift"
        End If
    End If
End With

With ParentForm.flxInfo(Index%)
    .redraw = False
    
    .row = 0
    .col = 0
    .CellFontBold = True
    .CellForeColor = .ForeColor
    
    Lief% = Asc(sLifDat$)
    lif.GetRecord (Lief% + 1)
    h$ = RTrim$(lif.kurz)
    h$ = h$ + " ("
    h2$ = Mid$(sLifDat$, 2, 6)
    h$ = h$ + Left$(h2$, 2) + "." + Mid$(h2$, 3, 2) + "." + Right$(h2$, 2)
    h2$ = Mid$(sLifDat$, 8, 4)  ' Format(CVI(Mid$(sLifDat$, 8, 2)), "0000")
    h$ = h$ + " " + Left$(h2$, 2) + ":" + Mid$(h2$, 3)
    h$ = h$ + ")"
    .TextMatrix(0, 0) = h$
    
    .row = 1
    .col = 0
    .CellFontBold = True
    .CellForeColor = .ForeColor
    .CellAlignment = flexAlignLeftCenter
    .TextMatrix(1, 0) = ParentForm.flxarbeit(Index%).TextMatrix(ParentForm.flxarbeit(Index%).row, 25)
    
    .row = 2
    .col = 0
    .CellFontBold = True
    .CellForeColor = vbBlue
    .TextMatrix(2, 0) = h3$
    
    j% = 4
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
'    .ColWidth(5) = ParentForm.TextWidth("(XXXXXX)")
'    .ColWidth(6) = ParentForm.TextWidth("99:99:9999")
'    .ColWidth(7) = ParentForm.TextWidth("999:99")
    .ColWidth(5) = ParentForm.TextWidth("999")
    .ColWidth(6) = ParentForm.TextWidth("999")
    .ColWidth(7) = ParentForm.TextWidth("999")
    .ColWidth(8) = ParentForm.TextWidth("999")
    
    .ColWidth(9) = ParentForm.TextWidth("999 ")
    .ColWidth(10) = ParentForm.TextWidth("999 ")
    
    .ColWidth(11) = ParentForm.TextWidth("99:9999")
    
    .ColWidth(12) = ParentForm.TextWidth("99999.99")
    
    .ColWidth(13) = ParentForm.TextWidth("WI")
    .ColWidth(14) = ParentForm.TextWidth("WII")
    
    .ColWidth(15) = ParentForm.TextWidth("WI") '+ wpara.FrmScrollHeight + 2 * wpara.FrmBorderHeight
    .ColWidth(16) = 0
    .ColWidth(17) = 0
    .ColWidth(18) = 0
    .ColWidth(19) = 0
    .ColWidth(20) = 0
    .ColWidth(21) = 0
    .ColWidth(22) = 0
    .ColWidth(23) = 0
    .ColWidth(24) = 0
    .ColWidth(25) = 0
    .ColWidth(26) = ParentForm.TextWidth("WI") + wpara.FrmScrollHeight + 2 * wpara.FrmBorderHeight
    ParentForm.Font.Bold = False
    
    spBreite% = 0
    For i% = 1 To .Cols - 1
        If (.ColWidth(i%) > 0) And (i% <> 13) And (i% <> 14) Then
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

Sub EditSatz(Optional StartWert$ = "")
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
Dim i%, rInd%, EditCol%, nix%, bm%, Lm%, nm%, rm%, cLm%, cRm%, cZiel%
Dim Soll%(2), SollMax%, ind%, WuEditErg%
Dim aAep#, aep#
Dim txt$, SollStr$(2)
Dim iFlex As MSFlexGrid
            
Set iFlex = ParentForm.flxarbeit(0)

EditCol% = iFlex.col

With iFlex
'    If (WuLifDat$ = Chr$(0)) Then
    If (EinzelneWu% = False) Then
        If (Val(.TextMatrix(.row, 24)) > 0) Then
            txt$ = "*"
        Else
            txt$ = " "
        End If
        WuLifDat$ = "@" + .TextMatrix(.row, 16) + txt$
        Call AuslesenWu(True)
        iFlex.col = EditCol%
    End If

    If (EditCol% = 8) And (RTrim$(.TextMatrix(.row, 14)) <> "") Then
        If (InStr(para.Benutz, "Y") > 0) Then Call ZeigeBesorgerDaten(1)
'        Call DefErrPop: Exit Sub
    End If
    
    If (EditCol% = 13) Then
            If (RTrim$(.TextMatrix(.row, 13)) <> "") Then
'                pzn$ = .TextMatrix(.row, 0)
'                txt$ = RTrim$(.TextMatrix(.row, 2)) + "  " + RTrim$(.TextMatrix(.row, 3))
'                txt$ = txt$ + RTrim$(.TextMatrix(.row, 4))
                Call ZeigeLocalAngebote
            End If
        Call DefErrPop: Exit Sub
    End If

    nix% = True
    If (SuperEdit%) Then
        SuperEdit% = False
        nix% = False
    ElseIf ((EditCol% = 2) Or (EditCol% = 11) Or (EditCol% = 12)) Then
        nix% = False
    ElseIf ((IstAltLast%) And ((EditCol% = 9) Or (EditCol% = 10))) Then
        nix% = False
    ElseIf ((IstAltLast% = False) And (EditCol% = 8)) Then
        nix% = False
    ElseIf ((EditCol% = 7) And (.TextMatrix(.row, 19) = "*") And (DarfRmKontrolle%) And (IstAltLast% = False)) Then
        nix% = False
    End If
    If (nix%) Then Beep: Call DefErrPop: Exit Sub
End With

WuEditErg% = WuEditSatz(StartWert$)
If (WuEditErg%) And (EditCol% = 8) Then
ElseIf (EditCol% = 7) Then
    rInd% = SucheFlexZeile(False)
    If (rInd% <= 0) Then Call DefErrPop: Exit Sub
    
    bm% = ww.WuBm
    Lm% = ww.WuLm
    nm% = ww.WuNm
    rm% = ww.WuRm
        
    
    If ((rm% <> bm%) Or (Lm% <> (bm% + nm%))) Then
        Soll%(0) = bm% + nm%
        If (nm% > 0) Then
            SollStr$(0) = "Bm+Nm"
        Else
            SollStr$(0) = "Bm"
        End If
        If (Lm% = Soll%(0)) Then
            SollStr$(0) = SollStr$(0) + ",Lm"
            SollMax% = 0
        Else
            Soll%(1) = Lm%
            SollStr$(1) = "Lm"
            SollMax% = 1
        End If
        nix% = True
        For i% = 0 To SollMax%
            If (rm% = Soll%(i%)) Then
                SollStr$(i%) = SollStr$(i%) + ",Rm"
                nix% = False
                Exit For
            End If
        Next i%
        If (nix%) Then
            SollMax% = SollMax% + 1
            Soll%(SollMax%) = rm%
            SollStr$(SollMax%) = "Rm"
        End If
    
        
        With frmEdit
            .Left = ParentForm.picBack(0).Left + iFlex.Left + iFlex.ColPos(9) + 45
            .Left = .Left + ParentForm.Left + wpara.FrmBorderHeight
            .Top = ParentForm.picBack(0).Top + iFlex.Top + (iFlex.row - iFlex.TopRow + 1) * iFlex.RowHeight(0)
            .Top = .Top + ParentForm.Top + wpara.FrmBorderHeight + wpara.FrmCaptionHeight + wpara.FrmMenuHeight
        End With
        With frmEdit.flxEdit
            .Left = 0
            .Top = 0
            
            .Rows = 2
            .FixedRows = 1
            .FormatString = "Sollwert|^±Rm|^±Lm|<Typ"
            .Rows = 1
            .Cols = 4
            
            .ColWidth(0) = frmEdit.TextWidth("99 (=BM+NM,LM,RM)")
            .ColWidth(1) = frmEdit.TextWidth("99999")
            .ColWidth(2) = frmEdit.TextWidth("99999")
            .ColWidth(3) = frmEdit.TextWidth("RET,GS  ")
            
            For i% = 0 To SollMax%
                cRm% = Soll%(i%) - nm% - rm%
                cLm% = Soll%(i%) - Lm%
                txt$ = ""
                If (cRm% > 0) Then
                    txt$ = "NV"
                ElseIf (cRm% < 0) Then
                    txt$ = "GS"
                End If
                If (cLm% > 0) Then
                    If (txt <> "") Then txt$ = txt$ + ","
                    txt$ = txt$ + "NL"
                ElseIf (cLm% < 0) Then
                    If (txt <> "") Then txt$ = txt$ + ","
                    txt$ = txt$ + "RET"
                End If
                .AddItem Str$(Soll%(i%)) + "  (=" + SollStr$(i%) + ")" + vbTab + Str$(cRm%) + vbTab + Str$(cLm%) + vbTab + txt$
            Next i%
            .AddItem "verwerfen" + vbTab + Str$(0) + vbTab + Str$(0)
            
            .Height = .RowHeight(0) * .Rows + 90
            .Width = .ColWidth(0) + .ColWidth(1) + .ColWidth(2) + .ColWidth(3) + 90
            
            frmEdit.Height = .Height
            frmEdit.Width = .Width
            
            .row = 1
            .col = 0
            .ColSel = .Cols - 1
            .Visible = True
        End With
    
        frmEdit.Show 1

        cLm% = 0
        cRm% = 0
        If (EditErg%) Then
            ind% = InStr(EditTxt$, vbTab)
            cZiel% = Val(Left$(EditTxt$, ind% - 1))
            EditTxt$ = Mid$(EditTxt$, ind% + 1)
            ind% = InStr(EditTxt$, vbTab)
            cRm% = Val(Left$(EditTxt$, ind% - 1))
            cLm% = Val(Mid$(EditTxt$, ind% + 1))
            
'            If (EditTxt$ = "Bestellmenge") Then
'                cLm% = (bm% + nm%) - Lm%
'                cRm% = bm% - Rm%
'                cZiel% = bm%
'            ElseIf (EditTxt$ = "Liefermenge") Then
'                cRm% = (Lm% - nm%) - Rm%
'                cZiel% = Lm%
'            ElseIf (EditTxt$ = "Rechnungsmenge") Then
'                cLm% = Rm% - (Lm% - nm%)
'                cZiel% = Rm%
'            End If

'            If (EditTxt$ = "Nachlieferung") Or (EditTxt$ = "Retourware") Then
'                If ((EditTxt$ = "Nachlieferung") And (Rm% <= bm%)) Or _
'                   ((EditTxt$ = "Retourware") And (Rm% >= bm%)) Then
'                    cLm% = (bm% + nm%) - Lm%
'                    cRm% = bm% - Rm%
'                Else
'                    cLm% = Rm% - (Lm% - nm%)
'                    cRm% = 0
'                End If
'            ElseIf (EditTxt$ = "Gutschrift") Or (EditTxt$ = "Nachverrechnung") Then
'                cLm% = 0
'                cRm% = (Lm% - nm%) - Rm%
'            ElseIf (EditTxt$ = "keine Altlast") Then
'                cLm% = 0
'                cRm% = 0
'            End If
        End If
        
        ww.SatzLock (1)
        rInd% = SucheFlexZeile(True)
        If (rInd% > 0) Then
            ww.WuNeuLm = cLm%
            ww.WuNeuRm = cRm%
            ww.WuNeuZiel = cZiel%
            ww.WuRetMenge = 0
            ww.PutRecord (rInd% + 1)
            Call AuslesenDateiSatz(rInd%, True)
        End If
        ww.SatzUnLock (1)
    End If
End If

Call DefErrPop
End Sub
      
Function WuEditSatz%(StartWert$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("WuEditSatz%")
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
Dim i%, ind%, lInd%, rInd%, EditCol%, aRow%, m%, aMeng%, iLief%, neu%, ret%
Dim abl%, nix%, heute%, VerfaelltInMonaten%, TaxeVerfallMonate%, WarnungMonate%
Dim aAep#, aep#
Dim h$, h2$, pzn$, txt$, SQLStr$
Dim iFlex As MSFlexGrid
            
ret% = False

Set iFlex = ParentForm.flxarbeit(0)

EditModus% = 0
            
EditCol% = iFlex.col

rInd% = SucheFlexZeile(False)
If (rInd% <= 0) Then Call DefErrPop: Exit Function

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
    If (EditCol% = 2) Then .Top = .Top + iFlex.RowHeight(0)
End With
With frmEdit.txtEdit
    .Width = frmEdit.ScaleWidth
'            .Height = frmEdit.ScaleHeight
    .Left = 0
    .Top = 0
    h2$ = iFlex.TextMatrix(iFlex.row, EditCol%)
    
    If (EditCol% = 2) Then
        h2$ = Trim(ww.WuText)
        EditModus% = 1
    ElseIf (EditCol% = 11) Then
        If (Trim(h2$) = "") Then
            FabsErrf% = ass.IndexSearch(0, ww.pzn, FabsRecno&)
            If (FabsErrf% = 0) Then
                ass.GetRecord (FabsRecno& + 1)
                If (ass.PosLag > 0) Then
                    abl% = ass.abl2
                    If (ass.abl1 > abl%) Then abl% = ass.abl1
                    If (abl% > 0) Then
                        h2$ = Right$(sdate(abl%), 4)
                    End If
                End If
            End If
        End If
        EditModus% = 5
    ElseIf (EditCol% = 12) Then
        EditModus% = 4
    End If
    
    If (EditCol% >= 5) And (EditCol% <= 10) Then
        .MaxLength = 4
    Else
        .MaxLength = 0
    End If
    If (StartWert$ = "") Then
        .text = h2$
    Else
        .text = StartWert$
        frmEdit.tmrEdit.Enabled = True
    End If
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
    ww.SatzLock (1)
    rInd% = SucheFlexZeile(True)
    If (rInd% > 0) Then
        ret% = True
    
        neu% = False
        
'        If (ww.WuBelegDatum = 0) Then
        If (IstAltLast% = False) Then
            ww.WuBeleg = ActBeleg$
            ww.WuBelegDatum = ActBelegDatum%
            neu% = True
        End If
    
        If (EditCol% = 2) Then
            ww.WuText = Trim(EditTxt$)
            neu% = True
        ElseIf (EditCol% = 5) Then
            m% = Val(EditTxt$)
            ww.WuBm = m%
            neu% = True
        ElseIf (EditCol% = 6) Then
            m% = Val(EditTxt$)
            ww.WuNm = m%
            ww.WuRm = ww.WuBm
            ww.WuLm = ww.WuBm + ww.WuNm
            neu% = True
        ElseIf (EditCol% = 7) Then
            m% = Val(EditTxt$)
            ww.WuRm = m%
            ww.RmStatus = 1
            neu% = True
        ElseIf (EditCol% = 8) Then
            m% = Val(EditTxt$)
            ww.WuLm = m%
            ww.LmStatus = 1
            ww.RmStatus = 0
            ww.WuStatus = 0
            neu% = True
        ElseIf (EditCol% = 9) Or (EditCol% = 10) Then
            m% = Val(EditTxt$)
            If (EditCol% = 9) Then
                ww.WuNeuRm = m%
                ww.RmStatus = 1
            ElseIf (IstAltLast%) Then
                ww.WuNeuLm = m%
                ww.LmStatus = 1
            Else
                m% = -Abs(m%)
                ww.WuNeuLm = m%
                ww.WuNeuRm = m%
                ww.WuRetMenge = m%
            End If
            neu% = True
        ElseIf (EditCol% = 11) Then
            h2$ = ww.WuAblDatum
            h$ = "01" + EditTxt$
            If (h$ <> h2$) Then
                ww.WuAblDatum = h$
                neu% = True
            End If
        Else
            aAep# = ww.WuAEP
            aep# = Val(EditTxt$)
            If (aAep# <> aep#) Then
                ww.WuAEP = aep#
                neu% = True
            End If
        End If
        
        If (neu%) Then
            ww.PutRecord (rInd% + 1)
            Call AuslesenDateiSatz(rInd%, True)
        End If
    End If
    ww.SatzUnLock (1)
    
    If (EditCol% = 11) And (AnzVerfallWarnungen% > 0) Then
        SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + ww.pzn
        Set TaxeRec = TaxeDB.OpenRecordset(SQLStr$)
        If (TaxeRec.EOF = False) Then
            TaxeVerfallMonate% = TaxeRec!VerfallMonate
            If (TaxeVerfallMonate% > 0) Then
                abl% = Val(Mid$(EditTxt$, 3)) * 12 + Val(Left$(EditTxt$, 2))
                h$ = Format(Now, "MMYY")
                heute% = Val(Mid$(h$, 3)) * 12 + Val(Left$(h$, 2))
                VerfaelltInMonaten% = abl% - heute%
                
                WarnungMonate% = 0
                For i% = 0 To (AnzVerfallWarnungen% - 1)
                    If (VerfallWarnungen(i%).Laufzeit >= TaxeVerfallMonate%) Then
                        WarnungMonate% = VerfallWarnungen(i%).Warnung
                        Exit For
                    End If
                Next i%
                If (WarnungMonate% > 0) And (WarnungMonate% >= VerfaelltInMonaten%) Then
                    h$ = "Achtung: Artikel hat eine Laufzeit von" + Str$(TaxeVerfallMonate%) + " Monaten,"
                    h$ = h$ + vbCr + "verfällt aber in" + Str$(VerfaelltInMonaten%) + " Monaten !"
                    Call iMsgBox(h$, vbInformation, ww.txt)
                End If
            End If
        End If
    End If

End If

WuEditSatz% = ret%

Call DefErrPop
End Function
      
Sub WuBuchenAlle()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("WuBuchenAlle")
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
Dim i%, Lm%, BuchM%, rInd%, AltRow%, besMenge%, besStatus%, TuEs%, BenutzerNr%
Dim FrmActionTmrRowaEnabled%

BenutzerNr% = 1
If (IstRowaWu%) Then
    BenutzerNr% = 99
ElseIf (SignaturBeiBuchen%) Then
    BenutzerNr% = HoleBenutzerSignatur
End If
If (BenutzerNr% <= 0) Then
    Call MsgBox("Ohne gültiges Passwort keine Ware verbuchbar!" + vbCrLf + "Vorgang abgebrochen", vbExclamation)
    Call DefErrPop: Exit Sub
End If

With ParentForm.flxarbeit(0)
    .redraw = False
    AltRow% = .row
    
    If (.Rows > 10) Then
        Call StartAnimation(frmAction, "WÜ wird verbucht ...")
    End If

    ReDim LagerListePzn$(0)
    ReDim DokuPflichtPzn$(0)
    
    If (IstAltLast%) Then Call RetourenAusdruck
    
    With frmAction.tmrRowa
        FrmActionTmrRowaEnabled% = .Enabled
        .Enabled = False
    End With

    If (FremdPznOk%) And (IstDirektLief%) Then
        WARE_HANDLE% = FileOpen("opware.dat", "RW", "R", 50)
        If (WARE_HANDLE% > 0) Then
            If (LOF(WARE_HANDLE%) = 0) Then
                WuSatz$ = String(50, Chr$(0))
                Put WARE_HANDLE%, 1, WuSatz$
            End If
            Call iLock(WARE_HANDLE%, 1)
        End If
    End If

    ww.SatzLock (1)
    For i% = 1 To (.Rows - 1)
        If (.TextMatrix(i%, 19) = "*") Then
            .row = i%
            rInd% = SucheFlexZeile(True)
            If (rInd% > 0) Then
                TuEs% = False
                If (IstAltLast%) Then
                    If (ww.LmStatus) Or (ww.RmStatus) Then
                        TuEs% = True
                    End If
                ElseIf (ww.WuStatus = 0) Then
                    Lm% = ww.WuLm
                    BuchM% = ww.LmAnzGebucht
                    TuEs% = True
                End If
                If (TuEs%) Then
                    If (IstAltLast%) Then
                        ww.WuStatus = 0
                        If (ww.WuNeuLm = 0) And (ww.WuNeuRm = 0) Then ww.WuStatus = 2
                        If (ww.RmStatus) Then
                            ww.WuRm = ww.WuRm + ww.WuNeuRm
                            ww.RmAnzGebucht = ww.RmAnzGebucht + ww.WuNeuRm
                            ww.WuNeuRm = 0
                            ww.RmStatus = 0
                        End If
                        If (ww.LmStatus) Then
                            Call WuBuchen(RTrim$(.TextMatrix(.row, 0)), ww.WuNeuLm, BenutzerNr%)
                            If (HatPreisKontrolle%(ww.WuNeuLm)) Then ww.WuStatus = 1
                            ww.WuLm = ww.WuLm + ww.WuNeuLm
                            ww.LmAnzGebucht = ww.LmAnzGebucht + ww.WuNeuLm
                            ww.WuNeuLm = 0
                            ww.LmStatus = 0
                        End If
                        If (ww.WuStatus = 0) Then
                            If ((ww.WuRm = ww.WuBm) And (ww.WuLm = (ww.WuBm + ww.WuNm))) Or _
                               ((ww.WuRm = ww.WuNeuZiel) And (ww.WuLm = (ww.WuNeuZiel + ww.WuNm))) Then
                                ww.WuStatus = 2
                            End If
                        End If
                    Else
                        Call WuBuchen(RTrim$(.TextMatrix(i%, 0)), Lm% - BuchM%, BenutzerNr%)
                        ww.LmAnzGebucht = Lm%
                        ww.LmStatus = 2
                        If (HatPreisKontrolle%(ww.WuLm)) Then
                            ww.WuStatus = 1
                        Else
                            ww.WuStatus = 2
                        End If
                        If (RTrim$(.TextMatrix(i%, 14)) <> "") Then
                            If (InStr(para.Benutz, "Y") > 0) Then
                                besMenge% = Lm%     '- BuchM%
                                If (HatPreisKontrolle%(ww.WuLm)) Then
                                    besStatus% = 1
                                Else
                                    besStatus% = 5  '1 or 4
                                End If
                                Call BesorgerFertig(Val(RTrim$(.TextMatrix(i%, 21))), KorrPzn$, besStatus%, besMenge%)
                            End If
                        End If
                    End If
                    ww.PutRecord (rInd% + 1)
'                    Call AuslesenDateiSatz(rInd%, True)
                End If
            End If
        End If
    Next i%
    ww.SatzUnLock (1)
    
    If (FremdPznOk%) And (IstDirektLief%) Then
        Call iUnLock(WARE_HANDLE%, 1)
        Close #WARE_HANDLE%
    End If

    With frmAction.tmrRowa
         .Enabled = FrmActionTmrRowaEnabled%
    End With

    If (UBound(LagerListePzn$) > 0) And (DruckeLagerKontrollListe%) Then Call KontrollListe(0)
    
    If (UBound(DokuPflichtPzn$) > 0) Then
        Call DokuPflichtAnzeige
    End If
    
    If (.Rows > 10) Then
        Call StopAnimation(frmAction)
    End If

    .row = AltRow%
    .redraw = True
    
    Call ParentForm.ToolbarUpdateFlag(False)
End With

Call DefErrPop
End Sub

'Statistiksatz für Artikel:PZN$ aufbereiten MENGE% = Liefermenge -------
Sub WuBuchen(pzn$, menge%, Optional BenutzerNr% = 0)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("WuBuchen")
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

'Dim i%, ret%, row%, asatz%, ssatz%, wg%, found%, gilt%, lld%, AltUeblLief%, LagAlt%, LagNeu%
'Dim Verbr%, pp%, vBm%, vMm%, Verkauf%, xVerk%, mm%, llm%, dazu%, heute%
'Dim Tage%, ALS%, NLS%, XABL%, AB2%, AblMonat%, AblJahr%
'Dim aep#, AVP#, lager#, TagesVerbrauch#, SollVerbrauch#
'Dim h$, mw$, AbCode$, AltEan$, AufLager$, heuteC$, DirLif$, AblK$, abl$

Dim i%, asatz%, ssatz%, LagAlt%, LagNeu%, AblMonat%, AblJahr%, XABL%, AB2%
Dim lager#
Dim AbCode$, abl$

asatz% = 0
If (Val(pzn$) <> 0) And (pzn$ <> "9999999") Then
    If (FremdPznOk%) And (IstDirektLief%) And (IstAltLast% = 0) Then
        menge% = ZeigeDirektAufteilung%(menge%, 1)
    End If

    FabsErrf% = ast.IndexSearch(0, pzn$, FabsRecno&)
    If (FabsErrf% = 0) Then
        asatz% = CInt(FabsRecno&)
        ast.GetRecord (asatz% + 1)
'        aep# = ast.aep
'        AVP# = ast.AVP
'        mw$ = ast.mw
'        wg% = Val(ast.wg1)
    End If
End If


If (FabsErrf% <> 0) Or (asatz% = 0) Or (pzn$ = "9999999") Or (menge% = 0) Then
    Call DefErrPop
    Exit Sub
End If

AbCode$ = ast.abl

FabsErrf% = ass.IndexSearch(0, pzn$, FabsRecno&)
If (FabsErrf% = 0) Then
'    If (FremdPznOk%) And (IstDirektLief%) And (IstAltLast% = 0) Then
'        menge% = ZeigeDirektAufteilung%(menge%, 1)
'    End If

    ssatz% = CInt(FabsRecno&)
    ass.GetRecord (ssatz% + 1)
    LagAlt% = ass.PosLag
    If (LagAlt% <= 0) Then
        ass.abl1 = 0
        ass.abl2 = 0
    End If

    lager# = CSng(ass.PosLag) + CSng(menge%)
    If (lager# > 32000) Then lager# = 32000
    If (lager# < -32000) Then lager# = -32000
    ass.PosLag = CInt(lager#)
    
    'Ablauf eintragen ------------------------------------------------------------
    abl$ = RTrim(ww.WuAblDatum)
'    If (abl$ <> "") Then
    If (Len(abl$) > 2) Then
        AblMonat% = Mid$(abl$, 3, 2)
        AblJahr% = Mid$(abl$, 5, 2)
        AblMonat% = AblMonat% + 1
        If (AblMonat% > 12) Then
            AblMonat% = 1
            AblJahr% = AblJahr% + 1
        End If
        abl$ = "01" + Format(AblMonat%, "00") + Format(AblJahr%, "00")
        XABL% = iDate(abl$) - 1
        
        If (AbCode$ <> "A") And (para.AblK <> "A") Then
            ass.abl1 = 0
            ass.abl2 = 0
        Else
            'Beide Abl.daten vorhanden ?
            If (ass.abl1 = 0) Or (ass.abl2 = 0) Then
                'Abl.daten=Neues
                ass.abl1 = XABL%
                ass.abl2 = XABL%
            Else
                AB2% = ass.abl2
    '            If (AB2% > XABL%) And (Val(Right$(abl$, 2)) > 80 Or Val(Right$(X$, 2)) < 80) Then  'V2.83
                If (AB2% > XABL%) Then  'V2.83
                    'Neu bleibt,Alt = Neues
                    ass.abl1 = XABL%
                Else
                    'Alt = Neu   Neu = Neues
                    ass.abl1 = ass.abl2
                    ass.abl2 = XABL%
                End If
            End If
        End If
    End If

    ass.lagbew = "*"
    ass.pzn = pzn$
    
    ass.PutRecord (ssatz% + 1)
    
    If (PosArtikel%()) Then
        i% = UBound(LagerListePzn$) + 1
        ReDim Preserve LagerListePzn$(i%)
        LagerListePzn$(i%) = ass.pzn
    End If
    
    LagNeu% = ass.PosLag
    Call LagerBew(pzn$, LagAlt%, LagNeu%, BenutzerNr%)
End If

If (ww.ArtikelKz2 And KZ_DOKUPFLICHT) Then
    i% = UBound(DokuPflichtPzn$) + 1
    ReDim Preserve DokuPflichtPzn$(i%)
    DokuPflichtPzn$(i%) = pzn$ + "  " + ww.txt
End If

Call DefErrPop
End Sub

Function WuFertig%()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("WuFertig%")
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

'Dim i%, ret%, row%, asatz%, ssatz%, wg%, found%, gilt%, lld%, AltUeblLief%, LagAlt%, LagNeu%
'Dim Verbr%, pp%, vBm%, vMm%, Verkauf%, xVerk%, mm%, llm%, dazu%, heute%
'Dim Tage%, ALS%, NLS%, XABL%, AB2%, AblMonat%, AblJahr%
'Dim aep#, AVP#, lager#, TagesVerbrauch#, SollVerbrauch#
'Dim h$, mw$, AbCode$, AltEan$, AufLager$, heuteC$, DirLif$, AblK$, abl$

'Dim asatz%, ssatz%, LagAlt%, LagNeu%, AblMonat%, AblJahr%, XABL%, AB2%
'Dim lager#
'Dim AbCode$, abl$

Dim i%, ret%, asatz%, ssatz%, menge%, lld%, llm%, Rmenge%, Natu%, heute%
Dim pzn$


ret% = False

If (ww.auto = "v") And (Val(ww.AbholNr) > 0) Then
    Call BesorgerLoeschen(Val(ww.AbholNr))
End If


pzn$ = ww.pzn
If (IstAltLast%) Then
    menge% = ww.LmAnzGebucht
    Rmenge% = ww.RmAnzGebucht
    heute% = iDate(Format(Now, "DDMMYY"))
Else
    menge% = ww.WuLm
    Rmenge% = ww.WuRm
    heute% = ww.WuBelegDatum
End If
Natu% = menge% - Rmenge%

asatz% = 0
If (Val(pzn$) <> 0) And (pzn$ <> "9999999") Then
    FabsErrf% = ast.IndexSearch(0, pzn$, FabsRecno&)
    If (FabsErrf% = 0) Then
        asatz% = CInt(FabsRecno&)
'        ast.GetRecord (asatz% + 1)
    End If
End If

If (FabsErrf% <> 0) Or (asatz% = 0) Or (pzn$ = "9999999") Or (menge% = 0) Then
    WuFertig% = ret%
    Call DefErrPop
    Exit Function
End If

FabsErrf% = ass.IndexSearch(0, pzn$, FabsRecno&)
If (FabsErrf% = 0) Then
    ssatz% = CInt(FabsRecno&)
    ass.GetRecord (ssatz% + 1)
    
'    If (IstAltLast% = False) Then
    If (menge% > 0) Then
        lld% = ass.lld
        llm% = ass.llm
        
        If (lld% = heute%) Then
            menge% = llm% + menge%
        End If
        
        ass.lld = heute%
        ass.llm = menge%
    End If

    
    For i% = 0 To 3
        ass.llief(i%) = ass.llief(i% + 1)
        ass.lldat(i%) = ass.lldat(i% + 1)
        ass.llrm(i%) = ass.llrm(i% + 1)
        ass.llnm(i%) = ass.llnm(i% + 1)
    Next i%
    ass.llief(4) = ww.Lief
    ass.lldat(4) = heute%
    ass.llrm(4) = Rmenge%
    ass.llnm(4) = Natu%
    
    ass.lagbew = "*"
    ass.pzn = pzn$
    
    ass.PutRecord (ssatz% + 1)
    
'    If (PosArtikel%()) Then
'        i% = UBound(LagerListePzn$) + 1
'        ReDim Preserve LagerListePzn$(i%)
'        LagerListePzn$(i%) = ass.pzn
'    End If
    
    If (ass.halt = "S") Then
        i% = UBound(IntStreichungRow%) + 1
        ReDim Preserve IntStreichungRow%(i%)
        IntStreichungRow%(i%) = ParentForm.flxarbeit(0).row
    End If
    
    ret% = True
End If

WuFertig% = ret%

Call DefErrPop
End Function

''Statistiksatz für Artikel:PZN$ aufbereiten MENGE% = Liefermenge -------
'Sub aWuBuchen(pzn$, menge%)
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Call DefErrFnc("WuBuchen")
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
'Dim i%, ret%, row%, asatz%, ssatz%, wg%, found%, gilt%, lld%, AltUeblLief%, LagAlt%, LagNeu%
'Dim Verbr%, pp%, vBm%, vMm%, Verkauf%, xVerk%, MM%, llm%, dazu%, heute%
'Dim Tage%, ALS%, NLS%, XABL%, AB2%, AblMonat%, AblJahr%
'Dim aep#, AVP#, lager#, TagesVerbrauch#, SollVerbrauch#
'Dim h$, mw$, AbCode$, AltEan$, AufLager$, heuteC$, DirLif$, AblK$, abl$
'
'Dim LiefNr%, Rmenge%, Natu%, KontrDruck%
'
'LiefNr% = ww.Lief
'Rmenge% = ww.WuRm
'Natu% = ww.WuNm
'
'KontrDruck% = False
'
'h$ = Format(Now, "DDMMYY")
'heute% = iDate(h$)
'heuteC$ = MKDate(heute%)
'
'found% = 0
'asatz% = 0
'
'AufLager$ = ww.WuStat
'
'
'If (Val(pzn$) <> 0) And (pzn$ <> "9999999") Then
'    FabsErrf% = ast.IndexSearch(0, pzn$, FabsRecno&)
'    If (FabsErrf% = 0) Then
'        asatz% = CInt(FabsRecno&)
'    End If
'End If
'
'If (asatz% > 0) Then
'    ast.GetRecord (asatz% + 1)
'    aep# = ast.aep
'    AVP# = ast.AVP
'    mw$ = ast.mw
'    wg% = Val(ast.wg1)
'Else
''    AEP# = w.AEP#
''    AVP# = w.AVP#
''    mw$ = "?"
''    wg% = 9
'End If
'
'
'If (FabsErrf% <> 0) Or (asatz% = 0) Or (pzn$ = "9999999") Or (menge% = 0) Then
'    Call DefErrPop
'    Exit Sub
'End If
'
'AbCode$ = ast.abl
'
'gilt% = 0
'If (InStr(ast.rez, "P") <> 0) Then gilt% = True
''If (gilt%) And (w.AEP# <> AEP# Or w.AVP# <> AVP#) Then
''    If (w.AEP# <> 0#) Then ast.AEP = w.AEP#
''    ast.AVP = w.AVP#
''    ast.putRecord (asatz% + 1)
''End If
'
'
'LagAlt% = 0
'ssatz% = 0
'lld% = 0
'AltUeblLief% = 0
'
'FabsErrf% = ass.IndexSearch(0, pzn$, FabsRecno&)
'If (FabsErrf% = 0) Then
'    ssatz% = CInt(FabsRecno&)
'    ass.GetRecord (ssatz% + 1)
'    LagAlt% = ass.PosLag
'    If (LagAlt% <= 0) Then
'        ass.abl1 = 0
'        ass.abl2 = 0
'        ass.PutRecord (ssatz% + 1)
'    End If
'    AltUeblLief% = ass.Lief
'
'    lld% = ass.lld
''    lld$ = sDate$(lld%)
'    AltEan$ = ass.ean
''    If lld% = 0 Then LSet ass.lld$ = heuteC$
''    If ass.HALT$ = "S" Then GoSub InterneStreichung
'End If
'
'Verbr% = 0
'Verkauf = 0
'If (FabsErrf% <> 0) Then
'    If (AufLager$ <> "J") Then Call DefErrPop:    Exit Sub
''    If (EanIdx% <> 0) Then
''        Cmnd$ = "I\4\" + Space$(13): GoSub fabs: If errf% <> 0 Then Error 255
''        AltEan$ = Space$(13)
''    End If
''    Cmnd$ = "I\3\" + pzn$: GoSub fabs
''
''    If errf% = 0 Then ssatz% = rec%: ssatz = xrec Else Return
''    GET #F.ASS%, 1: SMAX = FNSATZ(CVI(LEFT$(F.ASS$, 2)))
''    IF SMAX < ssatz THEN LSET F.ASS$ = MKI$(ssatz%): PUT #F.ASS%, 1
''    LSet f.ass$ = String$(Len(f.ass$), 0): LSet ass.pzn$ = pzn$: ass.ean$ = AltEan$
''    GoSub 30000: LSet ass.bm$ = MKI$(km): LSet ass.mm$ = MKI$(min%)
''    LSet ass.lief$ = Chr$(AltUeblLief%): LSet ass.lld$ = heuteC$
''    GoSub PosArtikel
''    If pp% <> 0 Then LSet ass.PosLag$ = MKI$(menge%)
''    LSet ass.als$ = MKI$(0)
'Else
''    Call PosArtikel(pp%)
'    vBm% = ass.vBm
'    vMm% = ass.vMm
'    If (vBm% <> 0) Or (vMm% <> 0) Then
'        'vorgeschriebene BM noch aktuell ?
''        Call NochAktuell(1, text$, 0, "", vBm%, vMm%, f.ass%, ssatz%, f.ast%, f.ast$, f.lif%, POSAktivWG$) '2.96
'    Else
'         'willkürliche BM
'        If (pp% = 0) And (ass.bm < menge% And ww.Lief > 0) Then
'            lif.GetRecord (ww.Lief + 1)
'            DirLif$ = lif.DirektLieferant
'            If (DirLif$ <> "J") And (DirLif$ <> "V") Then
''                Call NochAktuell(2, text$, 0, "", CVI(ass.bm$), menge%, f.ass%, ssatz%, f.ast%, f.ast$, f.lif%, POSAktivWG$) '2.96
''                Call FussZeile(xm$, xm%)
'            End If
'        End If
'    End If
'
'    MM% = ass.MM
'    llm% = ass.llm
'    ALS% = ass.lag 'AlterLagerStand
'    For xVerk% = 1 To 10
'        Verkauf% = Verkauf% + ass.vkmng(xVerk% - 1)
'    Next xVerk%
'    If (pp%) Then
'      If (Verkauf% <> 0) Then NLS% = ass.PosLag
'      lager# = CSng(ass.PosLag) + CSng(menge%)
'      If (lager# > 32000) Then lager# = 32000
'      If (lager# < -32000) Then lager# = -32000
'      ass.PosLag = CInt(lager#)
'    End If
'
'    If (lld% = heute%) Then
'        menge% = llm% + menge%
''    ElseIf Typ$ = "N" And menge% < (llm% / 4) Then
''        menge% = menge% + llm%
''    ElseIf Typ$ = "N" And lm% <> 0 Then
''        menge% = menge% + llm%
'    Else
'        dazu% = 0
'        If (pp%) Then          'schon verbraucht ?
'            Tage% = heute% - lld%
'            TagesVerbrauch# = CDbl(llm%) / CDbl(Tage%)
'            SollVerbrauch# = ass.opt / para.bp1
'            '3.04 AND statt OR
'            If (TagesVerbrauch# >= SollVerbrauch# * 4) And (ass.opt <> 0) And (llm% >= 4) Then
''                XMALT$ = xm$
''                Call NochAktuell(4, text$, dazu%, lld$, Menge%, llm%, f.ass%, ssatz%, f.ast%, f.ast$, f.lif%, POSAktivWG$) '2.96
''                xm$ = XMALT$: Call FussZeile(xm$, xm%)
'            End If
'        End If
'        If (dazu% <> 0) Then
'            menge% = llm% + menge%
'        Else
'            If (pp% = 0) Or (Verkauf% = 0) Then
'                'bei POR oder POS ohne Verkäufe
'                'alter Lagerstand nicht aus Lagerstandfeld lesen
'                If (ALS% < 0) Then ALS% = MM%
'                If (ALS% > 32000) Then ALS% = 32000
'                If (NLS% < 0) Then NLS% = MM%
'                If (NLS% > 32000) Then NLS% = 32000
'            End If
'            'Verbrauch=LetzteLiefermenge+LagerAlt-LagerNeu für POR+POS
'            Verbr% = llm% + ALS% - NLS%
'            If (Verbr% <= 0) Then Verbr% = 1
'        End If
'    End If
'End If
'
'If (Verbr% <> 0) Then
'  ass.lag = NLS%
'  If (pp% = 0) Or (Verkauf% = 0) Then
'    Call Quartale
'    For i% = 0 To 16
'        ass.vkmon(i%) = ass.vkmon(i% + 1)
'        ass.vkmmng(i%) = ass.vkmmng(i% + 1)
'    Next i%
'    ass.vkmon(17) = ass.lld
'    ass.vkmmng(17) = Verbr%
'  End If
'  ass.lld = heute%
'End If
'
'ass.llm = menge%
'
''Ablauf eintragen ------------------------------------------------------------
'abl$ = RTrim(ww.WuAblDatum)
'If (abl$ <> "") Then
'    AblMonat% = Mid$(abl$, 3, 2)
'    AblJahr% = Mid$(abl$, 5, 2)
'    AblMonat% = AblMonat% + 1
'    If (AblMonat% > 12) Then
'        AblMonat% = 1
'        AblJahr% = AblJahr% + 1
'    End If
'    abl$ = "01" + Format(AblMonat%, "00") + Format(AblJahr%, "00")
'    XABL% = iDate(abl$) - 1
'
'    If (AbCode$ <> "A") And (para.AblK <> "A") Then
'        ass.abl1 = 0
'        ass.abl2 = 0
'    Else
'        'Beide Abl.daten vorhanden ?
'        If (ass.abl1 = 0) Or (ass.abl2 = 0) Then
'            'Abl.daten=Neues
'            ass.abl1 = XABL%
'            ass.abl2 = XABL%
'        Else
'            AB2% = ass.abl2
''            If (AB2% > XABL%) And (Val(Right$(abl$, 2)) > 80 Or Val(Right$(X$, 2)) < 80) Then  'V2.83
'            If (AB2% > XABL%) Then  'V2.83
'                'Neu bleibt,Alt = Neues
'                ass.abl1 = XABL%
'            Else
'                'Alt = Neu   Neu = Neues
'                ass.abl1 = ass.abl2
'                ass.abl2 = XABL%
'            End If
'        End If
'    End If
'End If
'
'For i% = 0 To 3
'    ass.llief(i%) = ass.llief(i% + 1)
'    ass.lldat(i%) = ass.lldat(i% + 1)
'    ass.llrm(i%) = ass.llrm(i% + 1)
'    ass.llnm(i%) = ass.llnm(i% + 1)
'Next i%
'ass.llief(4) = LiefNr%
'ass.lldat(4) = heute%
'ass.llrm(4) = Rmenge%
'ass.llnm(4) = Natu%
'
'ass.lagbew = "*"
'ass.pzn = pzn$
'ass.PutRecord (ssatz% + 1)
'found% = True
'
'LagNeu% = ass.PosLag
'Call LagerBew(pzn$, LagAlt%, LagNeu%)
'If pp% And KontrDruck% Then
'    ret% = MsgBox("Lagerkontrolliste drucken ?")
'    If (ret% = vbYes) Then
''        Call KontrollDruck(f.ast%, asatz, LagNeu%, Lieferant, Bedat$, Lieferung$, pr, d.DL%, fz1$, fz2$, pc$(2), pc$(3), pc$(3), pc$(4))       '2.96
'    End If
'End If
'
'Call DefErrPop
'
'End Sub

'Sub PosArtikel(pp%)
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Call DefErrFnc("PosArtikel")
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
'Dim errf%
'
'errf% = 0: pp% = 0
'If InStr("AF", ass.pp) <> 0 Then pp% = True
'If (InStr(para.PosAktivWG, ast.wg1) > 0) Then pp% = True
'
'Call DefErrPop
'End Sub


'SUB NochAktuell (Art%, text$, dazu%, s$, Menge1%, Menge2%, f.ass%, ssatz%, f.ast%, f.ast$, f.lif%, POSAktivWG$) STATIC
'
''vMengenAuto% Static!
'
'If Art% = 1 And vMengenAuto% Then Exit Sub
'
'TRUE = -1: BELL$ = CHR$(7): DEL$ = SPACE$(79)
'
'FIELD #f.ass%, 256 AS f.ass$
'FIELD #f.ass%, 7 AS ass.pzn$, 1 AS ASS.LIEF$, 1 AS ASS.HALT$, 2 AS ASS.LLD$, 2 AS ASS.LLM$, 2 AS ASS.ABL1$, 2 AS ass.ABL2$, 2 AS ASS.VBM$, 2 AS ASS.BM$, 2 AS ASS.MM$, 195 AS ass.Dummy1$,2 AS ass.vMM$, 12 AS ass.tara$, 24 AS ass.rest$
'
'LOCATE 15, 14: Print text$; " ";
'
'If Art% = 1 Then    'vBM, vMM
'  GoSub Vorgeschrieben
'ElseIf Art% = 2 Then 'willkürliche BM
'  GoSub WillkuerBM
'ElseIf Art% = 3 Then 'interne Streichung
'  GoSub IntS
'ElseIf Art% = 4 Then 'Schon verbraucht?
'  GoSub SchonVerbraucht
'End If
'
'Call farbe("N")
'LOCATE 15, 2: Print DEL$;: LOCATE 16, 2: Print DEL$;
'LOCATE 17, 2: Print DEL$;
'Exit Sub
'
''---------------------------------------------------------------------------
'
'Vorgeschrieben:
'
'  Call farbe("Nh"): Print "BM ="; Menge1%: Call farbe("N")
'  xm$ = "[000] vorgeschriebene BM entfernen  [F6] Fragen ausschalten  [F7] Statistik " + BELL$
'  Call FussZeile(xm$, xm%)
'
'  LOCATE 17, 14: Print "ist die vorgeschriebene BM von =>     noch aktuell ?";
'  X$ = Str$(Menge1%): Call MyTrim(X$)
'  ok% = 0
'  While Not (ok%)
'    xp$ = "4717042": Call ein500: X = Val(X$)
'    If lt% = 6 Then 'Fragen ausschalten
'      vMengenAuto% = True
'      ok% = True
'    ElseIf lt% = 7 Then
'      GoSub Statistikbild
'      ok% = 0: lt% = -1
'    ElseIf lt% <> 0 Or X < 0 Or X > 9999 Then
'      Print BELL$;
'    Else
'      ok% = True
'    End If
'  Wend
'  If (Not vMengenAuto%) Then
'    GET #f.ass%, FNSATZ(ssatz%) + 1
'    LSet ass.vBm$ = MKI$(X)
'    PUT #f.ass%, FNSATZ(ssatz%) + 1
'    If Menge2% <> 0 Then
'      xm$ = "[000] vorgeschriebene MM entfernen  [F6] Fragen ausschalten  [F7] Statistik " + BELL$
'      Call FussZeile(xm$, xm%)
'      LOCATE 15, 59: Call farbe("Nh")
'      Print "MM ="; Menge2%;: Call farbe("N")
'      LOCATE 17, 14: Print "ist die vorgeschriebene MM von =>     noch aktuell ?";
'      X$ = Str$(Menge2%): Call MyTrim(X$)
'      ok% = 0
'      While Not (ok%)
'        xp$ = "4717042": Call ein500: X = Val(X$)
'        If lt% = 6 Then 'Fragen ausschalten
'          vMengenAuto% = True
'          ok% = True
'        ElseIf lt% = 7 Then
'          GoSub Statistikbild
'          ok% = 0: lt% = -1
'        ElseIf lt% <> 0 Or X < 0 Or X > 9999 Then
'          Print BELL$;
'        Else
'          ok% = True
'        End If
'      Wend
'      If (Not vMengenAuto%) Then
'        GET #f.ass%, FNSATZ(ssatz%) + 1
'        LSet ass.vMm$ = MKI$(X)
'        PUT #f.ass%, FNSATZ(ssatz%) + 1
'      End If
'    End If
'  End If
'  ok% = 0
'
'  Return
'
'WillkuerBM:
'  Call farbe("Nh"): Print "BM ="; Menge1%
'  xm$ = "[<+] alles in die Statistik" + BELL$
'  Call FussZeile(xm$, xm%)
'
'  LOCATE 16, 22: Print "willkürliche Bestellmenge ="; Menge2%; "!": Call farbe("N")
'  LOCATE 17, 22: Print "in die Statistik eintragen     =>     "
'  lt% = -1
'  While lt% <> 0 Or X < 0 Or X > 9999
'    X$ = Str$(Menge2%): Call MyTrim(X$)
'    xp$ = "5517042": Call ein500: X = Val(X$)
'    If lt% <> 0 Or X < 0 Or X > 9999 Then Print BELL$;
'  Wend
'  Menge2% = X
'  Return
'
'IntS:
'  LOCATE 17, 14: Print "ist die interne Streichung noch aktuell ? =>  ";
'  xm$ = "[Leertaste] Streichung entfernen   [S] interne Streichung   [F7] Statistik " + BELL$
'  Call FussZeile(xm$, xm%)
'
'  X$ = s$
'  ok% = 0
'  While Not (ok%)
'    xp$ = "5917011"
'    Call ein500
'    X = InStr(" Ss", X$)
'    If lt% = 7 Then
'      GoSub Statistikbild
'      ok% = 0: lt% = -1
'    ElseIf lt% <> 0 Or X <= 0 Then
'      Print BELL$;
'    Else
'      ok% = True
'    End If
'  Wend
'  If X > 1 Then s$ = "S" Else s$ = " "
'  Return
'
'SchonVerbraucht:
'  Call farbe("Nh"): Print "LM ="; Menge1%
'
'  xm$ = "[N] noch nicht verbraucht  [J] bereits verbraucht": Call FussZeile(xm$, xm%)
'  LOCATE 17, 11
'  Print "sind die"; Menge2%; "Stück vom "; s$; " schon verbraucht ? (J/N) =>"
'  Call farbe("N")
'  X$ = "N": xp$ = "6917011"
'  X% = 0
'  While X% = 0
'    Call ein500
'    X% = InStr("JjNn", X$)
'    If X% = 0 Or lt% <> 0 Then Print BELL$;: X% = 0
'  Wend
'  If X% > 2 Then dazu% = True
'  Return
''----------------------------------------------------------------------------
'Statistikbild:
'  Call wsave
'  GET #f.ass%, FNSATZ(ssatz%) + 1
'  f.stb% = 18
'  Open "STATB" + Mid$(Str$(Val(User$)), 2) + ".$$$" For Output As #f.stb%
'  Print #f.stb%, f.ast$ + f.ass$
'  Close #f.stb%
'  Shell "STATBILD"
'  Call wload
'      ''CALL StatBild(f.ast%, f.ass%, f.lif%, f.ast$, f.ass$, POSAktivWG$)
'      'geht nicht, weil Platz f. Zeichenfolge zu gering
'  Return
''----------------------------------------------------------------------------
'End Sub

Sub Quartale()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("Quartale")
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
Dim i%, NextTag%, NextMenge%, quart%(3), xquart%, swap%, QuartQ%, QuartJ%, aq%, aj%
Dim NextTagX$, QuartTagX$, X$

NextMenge% = ass.vkmmng(0)
If (NextMenge% <> 0) Then
    For i% = 0 To 3
        quart%(i%) = ass.quart(i%)
    Next i%
    
    NextTag% = ass.vkmon(0)
    NextTagX$ = sdate(NextTag%)
    
    xquart% = ass.quarttag
    QuartTagX$ = sdate(xquart%)
    
    If (xquart% <> 0) And (NextTag% > xquart%) Then
        X$ = QuartTagX$: GoSub quart
        aq% = QuartQ%: aj% = QuartJ%
        X$ = NextTagX$: GoSub quart
        While (aq% <> QuartQ%) Or (aj% <> QuartJ%)
            For i% = 0 To 2
                quart%(i%) = quart%(i% + 1)
            Next i%
            quart%(3) = 0
            aq% = aq% + 1
            If (aq% > 4) Then
                aq% = 1
                aj% = aj% + 1
            End If
        Wend
        quart%(3) = quart%(3) + NextMenge%
    Else
        For i% = 0 To 2
            quart%(i%) = 0
        Next i%
        quart%(3) = NextMenge%
    End If
    
    ass.quarttag = NextTag%
    For i% = 0 To 3
        ass.quart(i%) = quart%(i%)
    Next i%
End If

Call DefErrPop
Exit Sub

'*-----------------------------------------------------------------
quart:

QuartQ% = Int((Val(Mid$(X$, 3, 2)) + 2) / 3)
QuartJ% = Val(Mid$(X$, 5, 2))
If QuartJ% >= 50 Then
  QuartJ% = QuartJ% + 1900
Else
  QuartJ% = QuartJ% + 2000
End If
Return
'*----------------------------------------------------------

End Sub

Sub WuFertigAlle(Optional KomplettAbschluss% = True, Optional NeuLIef% = 0)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("WuFertigAlle")
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
Dim i%, j%, Lm%, BuchM%, rInd%, AltRow%, AltCol%, found%, TuEs%, gef%, LoescheRowaLs%, iWg%
Dim AlleBearbeitet%, AlleEinwiegerAblauf%, SollPreisKalk%, WirdAltLast%, WumsatzMenge%, iRabatt%
Dim OrgGesamtWert#, Preis#, ZeilenWert#
Dim WumsatzDatum$, SQLStr$, sPzn$
Dim FrmActionTmrRowaEnabled%

With ParentForm.flxarbeit(0)
    
    If (EinzelneWu% = False) Then
        Call iMsgBox("Aktion nicht durchführbar, da noch keine Lieferung ausgewählt wurde !", vbInformation)
        Call DefErrPop: Exit Sub
    End If
    
    If (IstAltLast% = False) And (KomplettAbschluss%) Then
        AlleBearbeitet% = True
        KeinRowColChange% = True
        For i% = 1 To (.Rows - 1)
            If (.TextMatrix(i%, 19) <> "*") Or (Trim(.TextMatrix(i%, 1)) = "") Then
                AlleBearbeitet% = False
                Exit For
            End If
        Next i%
        
        If (AlleBearbeitet% = False) Then
            Call iMsgBox("Aktion nicht durchführbar, da noch nicht alle Positionen gebucht sind !", vbInformation)
            KeinRowColChange% = False
            Call DefErrPop: Exit Sub
        End If
    End If
    
    .redraw = False
    
    SollPreisKalk% = False
    For i% = 1 To (.Rows - 1)
        If (.TextMatrix(i%, 1) = "$") Then
            SollPreisKalk% = True
            Exit For
        End If
    Next i%
    
    If (IstAltLast% = False) Then
        If (SollPreisKalk%) And (DarfPreisKalk% = False) Then
            Call iMsgBox("Aktion nicht durchführbar, da die Tätigkeit 'Preiskalkulation' für Sie derzeit gesperrt ist !", vbInformation)
            .redraw = True
            KeinRowColChange% = False
            Call DefErrPop: Exit Sub
        End If
        
        If (SollPreisKalk%) And (para.Land = "A") Then
            AlleEinwiegerAblauf% = True
            For i% = 1 To (.Rows - 1)
                If (.TextMatrix(i%, 1) = "$") Then
                    sPzn$ = .TextMatrix(i%, 0)
                    If (sPzn$ <> "9999999") Then
                        FabsErrf% = ast.IndexSearch(0, sPzn$, FabsRecno&)
                        If (FabsErrf% = 0) Then
                            ast.GetRecord (FabsRecno& + 1)
                            iWg% = Val(ast.wg)
                            If (iWg% = 3) Or ((iWg% >= 30) And (iWg% <= 39)) Then
                                If (Trim(.TextMatrix(i%, 11)) = "") Then
                                    AlleEinwiegerAblauf% = False
                                    Exit For
                                End If
                            End If
                        End If
                    End If
                End If
            Next i%
            If (AlleEinwiegerAblauf% = False) Then
                Call iMsgBox("Aktion nicht durchführbar, da noch Verfalldaten für Einwieger fehlen!", vbInformation)
                KeinRowColChange% = False
                .redraw = True
                Call DefErrPop: Exit Sub
            End If
        End If
        
        If (DarfRmKontrolle% = False) Then
            AltRow% = .row
            found% = False
            For i% = 1 To (.Rows - 1)
                .row = i%
                If (KomplettAbschluss%) Or (Trim(.TextMatrix(i%, 1)) <> "") Then
                    rInd% = SucheFlexZeile(False)
                    If (rInd% > 0) Then
                        If (ww.RmStatus = 0) And (ww.WuLm <> (ww.WuBm + ww.WuNm)) Then
                            found% = True
                            Exit For
                        End If
                    End If
                End If
            Next i%
            .row = AltRow%
            If (found%) Then
                Call MsgBox("Aktion nicht durchführbar, da die Tätigkeit 'RM-Kontrolle' für Sie derzeit gesperrt ist !", vbInformation)
                .redraw = True
                KeinRowColChange% = False
                Call DefErrPop: Exit Sub
            End If
        End If
        
        Call ParentForm.HighlightZeile(True)
        For i% = 1 To (.Rows - 1)
            .row = i%
            If (KomplettAbschluss%) Or (Trim(.TextMatrix(i%, 1)) <> "") Then
                rInd% = SucheFlexZeile(False)
                If (rInd% > 0) Then
                    If (ww.RmStatus = 0) And (ww.WuLm <> (ww.WuBm + ww.WuNm)) Then
                        .redraw = True
                        .col = 7
                        If (.row < .TopRow) Then
                            .TopRow = .row
                        Else
                            If (.row >= (.TopRow + ParentBereich.ArbeitAnzZeilen - 2)) Then
                                .TopRow = .row - ParentBereich.ArbeitAnzZeilen + 2
                            End If
                    '        While ((.row - .TopRow) >= (ParentBereich.ArbeitAnzZeilen - 1))
                    '            .TopRow = .TopRow + 1
                    '        Wend
                        End If
                        Call ParentForm.HighlightZeile
                        Call EchtKurzInfo
                        Call EditSatz
                        .redraw = False
                    End If
                End If
            End If
        Next i%
    
        .redraw = False
        
        If NettoEk% Then
            For i% = 1 To (.Rows - 1)
                If (KomplettAbschluss%) Or (Trim(.TextMatrix(i%, 1)) <> "") Then
                    If (Trim$(.TextMatrix(i%, 23)) = "") Then
                        Call NNAEPRechnen(i%)
                    End If
                End If
            Next i%
        End If
    End If
    
    .redraw = False
    
    If (SollPreisKalk%) And (DarfPreisKalk%) Then
        If (para.Land = "A") Then
            frmPreisKalkA.Show 1
            If (PreisKalkAepChange%) Then
                Call AuslesenWu
            End If
        Else
            PreisKalkModus% = 1
            frmWuPreisKalk.Show 1
            If (PreisKalkErg% = False) Then
                .redraw = True
                KeinRowColChange% = False
                Call AuslesenWu
                Call DefErrPop: Exit Sub
            ElseIf (PreisKalkAepChange%) Then
                Call AuslesenWu
            End If
        End If
    End If
    
    
    TuEs% = True
    If (IstAltLast% = False) Then
        OrgGesamtWert# = GesamtWert#
        
        If (KomplettAbschluss% = 0) Then
            GesamtWert# = 0#
            For i% = 1 To (.Rows - 1)
                If (Trim(.TextMatrix(i%, 1)) <> "") Then
                    Preis# = CDbl(.TextMatrix(i%, 12))
                    ZeilenWert# = Preis# * Val(.TextMatrix(i%, 7))
                    GesamtWert# = GesamtWert# + ZeilenWert#
                End If
            Next i%
        End If
        
        
        frmWarenWert.Show 1
        If (Warenwert# > 0#) Then
'            Call WumsatzSpeichern
        ElseIf (Warenwert# < 0#) Then
            TuEs% = False
        End If
        
        GesamtWert# = OrgGesamtWert#
    End If
    
    If (TuEs%) Then
        ParentForm.MousePointer = vbHourglass
    
        ReDim WumsatzRec(0)
        ReDim ASumsatzRec(0)
'        ReDim LagerListePzn$(0)
        ReDim IntStreichungRow%(0)
    
        AltRow% = .row
        AltCol% = .col
        
        If (.Rows > 10) Then
            Call StartAnimation(frmAction, "WÜ wird abgeschlossen ...")
        End If

        If (IstAltLast% = False) Then Call WEProt(0)

        LoescheRowaLs% = True
       
        With frmAction.tmrRowa
            FrmActionTmrRowaEnabled% = .Enabled
            .Enabled = False
        End With

        If (FremdPznOk%) And (IstDirektLief%) And (IstAltLast% = 0) Then
            NNEK_HANDLE% = FileOpen("opnnek.dat", "RW", "R", 65)
            If (NNEK_HANDLE% > 0) Then
                If (LOF(NNEK_HANDLE%) = 0) Then
                    NnEkSatz$ = String(65, Chr$(0))
                    Put NNEK_HANDLE%, 1, NnEkSatz$
                End If
                Call iLock(NNEK_HANDLE%, 1)
            End If
        End If

'        nnek.OpenDatei
        ww.SatzLock (1)
        For i% = 1 To (.Rows - 1)
            .row = i%
            
            If (KomplettAbschluss%) Or (Trim(.TextMatrix(i%, 1)) <> "") Then
                rInd% = SucheFlexZeile(True)
                If (rInd% > 0) Then
                    If (KomplettAbschluss% = 0) Then ww.Lief = NeuLIef%
                
                    WumsatzMenge% = 0
                    If (IstAltLast%) Then
                        If (ww.WuStatus = 2) Then
                            If (ww.RmAnzGebucht <> 0) Then
                                WumsatzMenge% = ww.RmAnzGebucht
                                WumsatzDatum$ = ww.WuBestDatum
                            End If
                            found% = WuFertig%
                            If (found% = 0) Then
                                Call MerkzettelSpeichern(ww.pzn, Left$(ww.txt, 33) + " " + Mid$(ww.txt, 34, 2), iDate(ww.WuBestDatum), ww.WuBelegDatum, ww.Lief, ww.LmAnzGebucht)
                            End If
                        
                            ww.status = 0
                            ww.WuBestDatum = Space$(6)
                            ww.PutRecord (rInd% + 1)
                        Else
                            LoescheRowaLs% = False
                        End If
                    Else
                        WumsatzMenge% = ww.WuRm
                        WumsatzDatum$ = ww.WuBestDatum
                        
                        found% = WuFertig%
                        If (found% = 0) And (ww.WuLm <> 0 Or ww.bm = 0) Then
                            Call MerkzettelSpeichern(ww.pzn, Left$(ww.txt, 33) + " " + Mid$(ww.txt, 34, 2), iDate(ww.WuBestDatum), ww.WuBelegDatum, ww.Lief, ww.WuLm)
                        End If
                    
                        If (NettoEk% = 0) Then
                            ww.WuNNAep = ww.WuAEP
                            ww.WuNNart = 9
                        End If
                        If (ww.pzn <> "9999999") And (ww.WuRm > 0) Then
                            Call NNAEPSpeichern
                        End If
                    
                        If (ww.pzn <> "9999999") Then Call WEProt(1)
                        WirdAltLast% = False
                        If (ww.WuRetMenge < 0) Then
                            ww.WuBm = 0
                            ww.WuNm = 0
                            ww.WuRm = Abs(ww.WuRetMenge)
                            ww.WuLm = Abs(ww.WuRetMenge)
                            ww.WuNeuRm = ww.WuRetMenge
                            ww.WuNeuLm = ww.WuRetMenge
                            ww.WuNeuZiel = 0
                            WirdAltLast% = True
                        ElseIf ((ww.WuRm <> ww.WuBm) Or (ww.WuLm <> (ww.WuBm + ww.WuNm))) And ((ww.WuNeuLm <> 0) Or (ww.WuNeuRm <> 0)) Then
                            WirdAltLast% = True
                        End If
                        
                        If (WirdAltLast%) Then
                            ww.IstAltLast = 1
                            ww.WuStatus = 0
                            ww.LmStatus = 0
                            ww.RmStatus = 0
                            ww.LmAnzGebucht = 0
                            ww.RmAnzGebucht = 0
                            ww.auto = " "
                            
                            LoescheRowaLs% = False
                        Else
                            ww.status = 0
                            ww.WuBestDatum = Space$(6)
                        End If
                        ww.PutRecord (rInd% + 1)
                    End If
                    
                    If (WumsatzMenge% <> 0) Then
                        iRabatt% = rabtab.HatRabatt(ww.Lief, ww.pzn, ww.WuNNart, Abs(ww.WuLm))
                        
                        gef% = False
                        For j% = 1 To UBound(WumsatzRec)
                            If (WumsatzRec(j%).Lief = ww.Lief) And (WumsatzRec(j%).bdatum = WumsatzDatum$) And (WumsatzRec(j%).Rabatt = iRabatt%) Then
                                WumsatzRec(j%).Wert = WumsatzRec(j%).Wert + (ww.WuAEP * WumsatzMenge%)
                                gef% = True
                                Exit For
                            End If
                        Next j%
                        If (gef% = False) Then
                            j% = UBound(WumsatzRec) + 1
                            ReDim Preserve WumsatzRec(j%)
                            WumsatzRec(j%).Lief = ww.Lief
                            WumsatzRec(j%).bdatum = WumsatzDatum$
                            WumsatzRec(j%).Wert = (ww.WuAEP * WumsatzMenge%)
                            WumsatzRec(j%).Rabatt = iRabatt%
                        End If
                    
                        If (para.Land = "D") Then
                            SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + ww.pzn
                            Set TaxeRec = TaxeDB.OpenRecordset(SQLStr$)
                            If (TaxeRec.EOF = False) Then
                                'Rabatt wird her verwendet für den Abgabeschlüssel
                                iRabatt% = TaxeRec!AbgabeBest
                            
                                gef% = False
                                For j% = 1 To UBound(ASumsatzRec)
                                    If (ASumsatzRec(j%).Lief = ww.Lief) And (ASumsatzRec(j%).bdatum = WumsatzDatum$) And (ASumsatzRec(j%).Rabatt = iRabatt%) Then
                                        ASumsatzRec(j%).Wert = ASumsatzRec(j%).Wert + (ww.WuAEP * WumsatzMenge%)
                                        gef% = True
                                        Exit For
                                    End If
                                Next j%
                                If (gef% = False) Then
                                    j% = UBound(ASumsatzRec) + 1
                                    ReDim Preserve ASumsatzRec(j%)
                                    ASumsatzRec(j%).Lief = ww.Lief
                                    ASumsatzRec(j%).bdatum = WumsatzDatum$
                                    ASumsatzRec(j%).Wert = (ww.WuAEP * WumsatzMenge%)
                                    ASumsatzRec(j%).Rabatt = iRabatt%
                                End If
'                            Else
'                                Beep
                            End If
                        End If
                    End If
                    
                End If
            End If
        Next i%
        ww.SatzUnLock (1)
'        nnek.CloseDatei
        
        If (FremdPznOk%) And (IstDirektLief%) And (IstAltLast% = 0) Then
            Call iUnLock(NNEK_HANDLE%, 1)
            Close #NNEK_HANDLE%
        End If

        With frmAction.tmrRowa
             .Enabled = FrmActionTmrRowaEnabled%
        End With

        If (RowaAktiv% And LoescheRowaLs%) Then Call RowaLsDatei(True)

        If (IstAltLast% = False) Then Call WEProt(2)
        
        If (UBound(WumsatzRec) > 0) Then Call WumsatzEinzeln(WumsatzRec)
        If (UBound(ASumsatzRec) > 0) Then Call ASumsatzEinzeln(ASumsatzRec)
        
        If (UBound(IntStreichungRow%) > 0) Then Call CheckInterneStreichung
        
        If (.Rows > 10) Then
            Call StopAnimation(frmAction)
        End If

        .row = AltRow%
        .col = AltCol%
        
        Call EntferneGeloeschteZeilen(0)
        
        WuLifDat$ = Chr$(0)
        
        ParentForm.MousePointer = vbDefault
    End If
    
    .redraw = True
    KeinRowColChange% = False
    
    Call AuslesenWu
End With

Call DefErrPop
End Sub

Function HatPreisKontrolle%(Lmenge%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("HatPreisKontrolle%")
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
Dim Aufschlag%, wg%, ret%
Dim pzn$, SQLStr$, sAbgabe$

ret% = True

pzn$ = ww.pzn


'If (ww.WuLm <> 0) And (pzn$ <> "9999999") Then
If (Lmenge% <= 0) Then
    ret% = False
ElseIf (para.Land = "A") Then
    If (pzn$ <> "9999999") Then
        FabsErrf% = ast.IndexSearch(0, pzn$, FabsRecno&)
        If (FabsErrf% = 0) Then
            ast.GetRecord (FabsRecno& + 1)
        Else
            SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + pzn$
            Set TaxeRec = TaxeDB.OpenRecordset(SQLStr$)
            If (TaxeRec.EOF = False) Then
                FabsErrf = 0
                Call Taxe2ast(pzn$)
            End If
        End If

        If (FabsErrf% = 0) Then
            wg% = Val(ast.wg1)
            Aufschlag% = Val(ast.ka)
            
            ret% = False
            If (wg% <> 1) And (wg% <> 2) And (wg% <> 5) Then ret% = True
            If (InStr(ast.kasz, "P") <> 0) Then ret% = True
            If (Aufschlag% = 9) Then ret% = True
            If (ast.aep = 0) Or (ast.AVP = 0) Then ret% = True
        End If
    End If
ElseIf (pzn$ <> "9999999") Then
    SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + pzn$
    Set TaxeRec = TaxeDB.OpenRecordset(SQLStr$)
    If (TaxeRec.EOF = False) Then
'        If (InStr("23", Chr$(TaxeRec!AbgabeBest)) = 0) And (TaxeRec!EK > 0) And (TaxeRec!VK > 0) Then ret% = False
        sAbgabe$ = "23"
        If (KalkNichtRezPflichtigeAM%) Then
            sAbgabe$ = sAbgabe$ + "0478"
        End If
        If (InStr(sAbgabe$, Format(TaxeRec!AbgabeBest, "0")) = 0) And (TaxeRec!EK > 0) And (TaxeRec!VK > 0) Then ret% = False
    End If
End If


'If (ww.WuLm <> 0) And (pzn$ <> "9999999") Then
'    FabsErrf% = ast.IndexSearch(0, pzn$, FabsRecno&)
'    If (FabsErrf% = 0) Then
'        ast.GetRecord (FabsRecno& + 1)
'
'        wg% = Val(ast.wg1)
'        Aufschlag% = Val(ast.ka)
'        If (InStr(ast.Rez, "P") <> 0) Then ret% = True
'        If (Aufschlag% = 9) Then ret% = True
'        If (ast.aep = 0) Or (ast.AVP = 0) Then ret% = True
'
'        FabsErrf% = ass.IndexSearch(0, pzn$, FabsRecno&)
'        If (ww.auto = "v") And (FabsErrf% <> 0) Then ret% = True
'    End If
'End If
  
HatPreisKontrolle% = ret%

Call DefErrPop

End Function

'Function HatPreisKontrolle%()
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Call DefErrFnc("HatPreisKontrolle%")
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
'Dim gilt%, liefm%, Aufschlag%, wg%, ret%
'Dim xaep#, xavp#
'Dim pzn$
'
'ret% = False
'
'gilt% = 0
'pzn$ = Left$(wu.at, 7)
''aText$ = wu.at
'liefm% = wu.Lm
'xaep# = wu.AEP
'xavp# = wu.AVP
'If (xaep# <= 0#) Then xaep# = 0.01
'If (xavp# <= 0#) Then xavp# = 0.01
'If (pzn$ <> "9999999") Then
'    FabsErrf% = ast.IndexSearch(0, pzn$, FabsRecno&)
'    If (FabsErrf% = 0) Then
'        ast.GetRecord (FabsRecno& + 1)
'
'        wg% = Val(ast.wg1)
'        Aufschlag% = Val(ast.ka)
'        If (InStr(ast.Rez, "P") <> 0) Then gilt% = True
'        If (Aufschlag% = 9) Then gilt% = True
'        If (ast.AEP = 0) Or (ast.AVP = 0) Then gilt% = True
'        '2.96 AVP in Stammdaten meist aktueller als in WÜ-Satz
'        If (ast.AVP <> 0) Then xavp# = ast.AVP
'
'        FabsErrf% = ass.IndexSearch(0, pzn$, FabsRecno&)
'        If (wu.besorger = "B") And (FabsErrf% <> 0) Then gilt% = True
'    Else
'        If (wu.ssatz <> 0) Or (wu.asatz <> 0) Then '3.09
'            wu.asatz = 0
'            wu.ssatz = 0
''            PUT #f.b%, bes%(i) + 1
'        End If
'    End If
'End If
'
'If (gilt%) And (liefm% <> 0) Then
'    ret% = True
''    pk% = pk% + 1: pk%(pk%) = bes%(i)
''    pkArrStr$ = pkArrStr$ + MKI$(bes%(i)) '23.1.95 für Übergabe an AlleKontrNr
''    If wg% = 3 And Val(ast.Meng2$) <> 0 And Val(ast.MENG$) <> 0 And land$ <> "D" Then
''      PreisFaktor = Val(ast.Meng2$) / Val(ast.MENG$)
''    Else
''      PreisFaktor = 1
''    End If
''    AEP# = xaep# * PreisFaktor: AVP# = xavp# * PreisFaktor
''    TAX$ = ast.TAX$: kzauf$ = ast.ka$         '25.1.95 ergänzt
''    LOCATE pk% + 3, 1: GoSub PkZeiAnz
'End If
'
'HatPreisKontrolle% = ret%
'
'Call DefErrPop
'
'End Function

'Function SucheFlexZeile%(Optional BereitsGelockt% = False)
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Call DefErrFnc("SucheFlexZeile%")
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
'Dim i%, row%, pos%, ret%, Max%
'Dim LaufNr$
'Dim pzn$, ch$
'
'ret% = False
'
'With ParentForm.flxarbeit(0)
'    row% = .row
'    LaufNr$ = .TextMatrix(row%, 14)
'    pos% = Val(Right$(.TextMatrix(row%, 15), 5))
'
'    If (BereitsGelockt% = False) Then Call ww.SatzLock(1)
'    ww.GetRecord (1)
'    Max% = ww.erstmax
'
'    ret% = SucheDateiZeile%(pos%, Max%, LaufNr$)
'
'    If (ret%) Then
''        If (bek.aktivlief > 0) Then
''            Call MsgBox("Bestellsatz gesperrt!")
''            ret% = False
''        End If
'    Else
'        Call MsgBox("WÜ-Satz nicht mehr vorhanden!")
'    End If
'
'    If (BereitsGelockt% = False) Then Call ww.SatzunLock(1)
'End With
'
'SucheFlexZeile% = ret%
'
'Call DefErrPop
'End Function

'Function SucheDateiZeile%(pos%, Max%, LaufNr$)
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Call DefErrFnc("SucheDateiZeile%")
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
'Dim i%, ret%
'Dim iLaufNr$
'
'ret% = False
'
'If (pos% <= Max%) Then
'    ww.GetRecord (pos% + 1)
'    iLaufNr$ = wu.li + wu.bd + wu.zeit + Format(wu.WuLaufNr, "00000")
'    If (iLaufNr$ = LaufNr$) Then
'        ret% = pos%
'    End If
'End If
'
'If (ret% = False) Then
'    For i% = 1 To Max%
'        ww.GetRecord (i% + 1)
'        iLaufNr$ = wu.li + wu.bd + wu.zeit + Format(wu.WuLaufNr, "00000")
'        If (iLaufNr$ = LaufNr$) Then
'            ret% = i%
'            Exit For
'        End If
'    Next i%
'End If
'
'SucheDateiZeile% = ret%
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
Dim LifDat$, h$, SRT$

If (BereitsGelockt% = False) Then Call ww.SatzLock(1)

With ParentForm.flxarbeit(0)
    .redraw = False
    
    ww.GetRecord (pos% + 1)
    
    LifDat$ = Chr$(ww.Lief) + ww.WuBestDatum + Format(CVI(ww.WuBestZeit), "0000")
    
    Call ZeigeWuZeile
    .TextMatrix(.row, 16) = LifDat$
    
    
    h$ = " "
'    h$ = h$ + Left$(ww.txt, 6)
    h$ = h$ + ww.txt
    SRT$ = h$ + Format(pos%, "00000")
    .TextMatrix(.row, 18) = SRT$
    
    .redraw = True
End With


If (BereitsGelockt% = False) Then Call ww.SatzUnLock(1)

Call ZeigeWerte

Call frmAction.HighlightZeile

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
Dim i%, j%, AnzRows%, ind%, r%, redraw%
Dim Preis#, ZeilenWert#
Dim sp&, x1&, x2&, y1&, y2&, TextH&, TextW&, NachLinks&
Dim tx$, BestellWerte$(3, 2), BestellWerteText$(3, 2)

MarkWert# = 0#
GesamtWert# = 0#

AnzBestellWerteRows% = 1

With ParentForm.flxarbeit(0)
    AnzRows% = .Rows - 1
    For i% = 1 To AnzRows%
        tx$ = .TextMatrix(i%, 18)
        If (tx$ <> "") Then
            Preis# = CDbl(.TextMatrix(i%, 12))
            ZeilenWert# = Preis# * Val(.TextMatrix(i%, 7))
            GesamtWert# = GesamtWert# + ZeilenWert#
            
            If (.TextMatrix(i%, 19) = "*") Then
                MarkWert# = MarkWert# + 1
            End If
        End If
    Next i%
    MarkWert# = AnzRows% - MarkWert#
    If (.TextMatrix(1, 0) = "XXXXXXX") Then
        MarkWert# = 0#
        AnzRows% = 0
    End If
End With

If (AnzeigeFlag%) Then
    With ParentForm.picBestellWerte
        .Font = ParentForm.flxarbeit(0).Font
        .Cls

        sp& = .Width / 3
        x1& = sp&
        For i% = 1 To 3
            ParentForm.picBestellWerte.Line (x1&, 0)-(x1&, .ScaleHeight), vbBlack
            x1& = x1& + sp&
        Next i%
    
        TextH& = .TextHeight("Äg")
        TextW& = .TextWidth("999999.99") + 2 * Screen.TwipsPerPixelX
       
        
        BestellWerteText$(0, 0) = "Anzahl Artikel"
        BestellWerte$(0, 0) = Format(AnzRows%, "0")
        
        BestellWerteText$(0, 1) = "Unbestätigte LM"
        BestellWerte$(0, 1) = Format(MarkWert#, "0")
        
        BestellWerteText$(0, 2) = "Wert"
        BestellWerte$(0, 2) = Format(GesamtWert#, "0.00")
        
        
        sp& = .Width / 3
        x1& = sp& - TextW& - 4 * Screen.TwipsPerPixelX
        x2& = sp& - 2 * Screen.TwipsPerPixelX
    
        For i% = 1 To 3
            For j% = 1 To AnzBestellWerteRows%
                y1& = (j% - 1) * TextH&
                If (AnzBestellWerteRows% = 1) Then
                    y2& = ParentForm.picBestellWerte.ScaleHeight - 15
                Else
                    y2& = j% * TextH&
                End If
                
                tx$ = BestellWerteText$(j% - 1, i% - 1)
                NachLinks& = .TextWidth(tx$)
                .CurrentX = x1& - 2 * Screen.TwipsPerPixelX - NachLinks&
                .CurrentY = y1&
                ParentForm.picBestellWerte.Print tx$
                
                ParentForm.picBestellWerte.Line (x1&, y1&)-(x2&, y2&), vbWhite, BF
                ParentForm.picBestellWerte.Line (x1&, y1&)-(x2&, y2&), vbBlack, B
                
                .CurrentX = x1& + 4 * Screen.TwipsPerPixelX
                .CurrentY = y1&
                ParentForm.picBestellWerte.Print BestellWerte$(j% - 1, i% - 1)
            Next j%
        
            x1& = x1& + sp&
            x2& = x2& + sp&
        Next i%
    End With
End If
        
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
Dim erg%, row%, col%, ret%, ind%, rInd%
Dim l&
Dim h$, pzn$, txt$, mErg$, Jetzt$, iLifDat$

Select Case Index

    Case MENU_F5
        If (IstAltLast% = False) Then
            col% = ParentForm.flxarbeit(0).col
            If (col% = 7) Or (col% = 8) Then
                ww.SatzLock (1)
                rInd% = SucheFlexZeile(True)
                If (rInd% > 0) Then
                    If (col% = 7) Then
                        ww.WuRm = ww.WuBm
                        ww.RmStatus = 0
                    ElseIf (ww.WuStatus = 0) Then
                        ww.WuLm = ww.WuBm + ww.WuNm
                        ww.LmStatus = 0
                    End If
                    ww.PutRecord (rInd% + 1)
                    Call AuslesenDateiSatz(rInd%, True)
                End If
                ww.SatzUnLock (1)
            End If
        End If
    
    Case MENU_F7
        Call WuBuchenAlle
        IstRowaWu% = 0
        Call AuslesenWu
    
    Case MENU_F8
        With ParentForm.flxarbeit(0)
            If (.TextMatrix(.row, 14) = "v") And (InStr(para.Benutz, "Y") > 0) Then
                Call ZeigeBesorgerDaten
            Else
                Call ZusatzFenster(ZUSATZ_ARTIKEL, KorrPzn$, KorrTxt$)
            End If
        End With
            
    Case MENU_SF2
        With ParentForm.flxarbeit(0)
            If (.TextMatrix(.row, 14) = "v") And (InStr(para.Benutz, "Y") > 0) Then
                Call ZeigeBesorgerDaten
            Else
                Call AnzeigeFenster(TEXT_BESTELLSTATUS, KorrPzn$, KorrTxt$)
            End If
        End With
        
    Case MENU_SF3
        h$ = WuLifDat$
        WuAuswahlModus% = 0
        frmWuAuswahl.Show 1
        If (WuLifDat$ <> h$) Then
            Call ActProgram.AuslesenWu(True)
        End If
    
    Case MENU_SF4
        With ParentForm.flxarbeit(0)
            If (RTrim$(.TextMatrix(.row, 13)) <> "") Then
'                pzn$ = .TextMatrix(.row, 0)
'                txt$ = RTrim$(.TextMatrix(.row, 2)) + "  " + RTrim$(.TextMatrix(.row, 3))
'                txt$ = txt$ + RTrim$(.TextMatrix(.row, 4))
                Call ZeigeLocalAngebote
            End If
        End With
        
    Case MENU_SF5
        ZeigeStatistik
    
    Case MENU_SF6
        Call WuFertigAlle
    
    Case MENU_SF7
        Call Dp04Einlesen
    
    Case MENU_SF8
        mErg$ = MatchCode(2, pzn$, txt$, False, False)
        If (mErg$ <> "") Then
            Jetzt$ = Format(Now, "DDMMYY") + Format(Val(Left$(Time$, 2)) * 100 + Val(Mid$(Time$, 4, 2)), "0000")
            Do
                If (mErg$ = "") Then Exit Do
                
                ind% = InStr(mErg$, vbTab)
                h$ = Left$(mErg$, ind% - 1)
                mErg$ = Mid$(mErg$, ind% + 1)
                
                ind% = InStr(h$, "@")
                pzn$ = Left$(h$, ind% - 1)
                h$ = Mid$(h$, ind% + 1)
                
                ind% = InStr(h$, "@")
                txt$ = Left$(h$, ind% - 1)
                h$ = Mid$(h$, ind% + 1)
                
                ManuellPzn$ = pzn$
                ManuellTxt$ = txt$
                Call ManuellBefuellen
                
                ind% = InStr(h$, "@")
                txt$ = Left$(h$, ind% - 1)
                h$ = Mid$(h$, ind% + 1)
                
                ind% = InStr(txt$, "(")
                If (ind% > 0) Then
                    txt$ = Mid$(txt$, ind% + 1)
                    iLifDat$ = Chr$(Val(Left$(txt$, Len(txt$) - 1))) + Jetzt$
                    Call NeuerDS(iLifDat$, h$)
                End If
            Loop
            Call ActProgram.ManuellFertig
        End If
        
    Case MENU_SF9
        frmAction.flxarbeit(0).Rows = 2
        frmRueckKauf.Show 1
        Call AuslesenWu
        
End Select

Call DefErrPop
End Sub

Sub Dp04Einlesen()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("Dp04Einlesen")
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
Dim i%, Dp04Deb%, Dp04Menge%, AnzAblDaten%
Dim h$, AblDaten$(1)
                
AnzAblDaten% = 0
AblDaten$(0) = ""
AblDaten$(1) = ""
        
Call LeseStiftEinlesen
Dp04Deb% = FileOpen("winwawi.dp0", "I")
Do While Not (EOF(Dp04Deb%))
    Line Input #Dp04Deb%, h$
    If (Left$(h$, 1) = "<") Then
        EingabeStr$ = PruefeDp04Zeile(h$, Dp04Menge%)
        If (Dp04Menge% = 0) Then Dp04Menge% = 1
        If (Len(EingabeStr$) = 13) And _
           ((Left$(EingabeStr$, 8) = "99999978") Or (Left$(EingabeStr$, 8) = "99999979")) Then
            If (AnzAblDaten% < 2) Then
                AblDaten$(1) = AblDaten$(0)
                AblDaten$(0) = EingabeStr$
                AnzAblDaten% = AnzAblDaten% + 1
            End If
        Else
            Call SucheZeile(True, Dp04Menge%)
            If (AnzAblDaten% = 2) Then
                For i% = 0 To 1
                    EingabeStr$ = AblDaten$(i%)
                    Dp04Menge% = 1
                    Call SucheZeile(True, Dp04Menge%)
                Next i%
            End If
            AnzAblDaten% = 0
            AblDaten$(0) = ""
            AblDaten$(1) = ""
        End If
    End If
Loop
Close #Dp04Deb%

Call DefErrPop
End Sub

Sub BackToBestellungAlle(Optional GanzeLieferung% = True)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("BackToBestellungAlle")
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
Dim i%, rInd%, ret%
Dim h$

If (GanzeLieferung%) Then
    ret% = iMsgBox("Wollen Sie wirklich alle Artikel von '" + ParentForm.lblArbeit(0) + "' zurück in die Bestellung stellen ?", vbYesNo Or vbDefaultButton2)
Else
    With ParentForm.flxarbeit(0)
        h$ = Trim(.TextMatrix(.row, 2)) + " " + Trim(.TextMatrix(.row, 3)) + Trim(.TextMatrix(.row, 4))
    End With
    ret% = iMsgBox("Wollen Sie wirklich '" + h$ + "' zurück in die Bestellung stellen ?", vbYesNo Or vbDefaultButton2)
End If
If (ret% <> vbYes) Then Call DefErrPop: Exit Sub

ww.SatzLock (1)

With ParentForm.flxarbeit(0)
    .redraw = False
    If (GanzeLieferung%) Then
        For i% = 1 To (.Rows - 1)
            .row = i%
            rInd% = SucheFlexZeile(True)
            If (rInd% > 0) Then
                ww.status = 1
                ww.PutRecord (rInd% + 1)
            End If
        Next i%
        WuLifDat$ = Chr$(0)
    Else
        rInd% = SucheFlexZeile(True)
        If (rInd% > 0) Then
            ww.status = 1
            ww.Lief = 0
            ww.aktivlief = 0
            ww.best = " "
            ww.zugeordnet = "N"
            ww.zukontrollieren = Chr$(0)
            ww.fixiert = Chr$(0)
            ww.absage = ww.absage + 1 'Lieferant%
            ww.PutRecord (rInd% + 1)
        End If
    End If
    .redraw = True
End With

ww.GetRecord (1)
ww.erstcounter = (ww.erstcounter + 1) Mod 100
ww.PutRecord (1)

ww.SatzUnLock (1)

Call AuslesenWu

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
Dim ArbeitRow%, Lm%, BuchM%, rInd%
Dim h$

rInd% = SucheFlexZeile(False)
If (rInd% <= 0) Then Call DefErrPop: Exit Sub

With ParentForm
    ArbeitRow% = .flxarbeit(0).row
    
    h$ = RTrim$(.flxarbeit(0).TextMatrix(ArbeitRow%, 19))
    If (h$ = "*") Then
        If (IstAltLast%) Then
            .flxInfo(0).TextMatrix(InfoRow%, 0) = TEXT_WU_BUCHEN
            InfoRow% = InfoRow% + 1
        ElseIf (Val(.flxarbeit(0).TextMatrix(ArbeitRow%, 17)) = 0) Then
            .flxInfo(0).TextMatrix(InfoRow%, 0) = TEXT_WU_BUCHEN
            InfoRow% = InfoRow% + 1
        ElseIf (.flxarbeit(0).TextMatrix(ArbeitRow%, 1) = "$") Then
            .flxInfo(0).TextMatrix(InfoRow%, 0) = TEXT_WU_PREIS
            InfoRow% = InfoRow% + 1
        End If
    End If
    
End With

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
Dim Lm%, BuchM%, rInd%, row%, besMenge%, besStatus%, TuEs%

If (h$ = TEXT_WU_BUCHEN) Then
    With ParentForm.flxarbeit(0)
        TuEs% = True
        ww.SatzLock (1)
        If (IstAltLast% = False) Then
            Lm% = Val(.TextMatrix(.row, 8))
            BuchM% = Val(.TextMatrix(.row, 22))
        End If
        If (TuEs%) Then
            rInd% = SucheFlexZeile(True)
            If (rInd% > 0) Then
                If (IstAltLast%) Then
                    ww.WuStatus = 0
                    If (ww.RmStatus) Then
                        ww.WuRm = ww.WuRm + ww.WuNeuRm
                        ww.RmAnzGebucht = ww.RmAnzGebucht + ww.WuNeuRm
                        ww.WuNeuRm = 0
                        ww.RmStatus = 0
                    End If
                    If (ww.LmStatus) Then
                        Call WuBuchen(RTrim$(.TextMatrix(.row, 0)), ww.WuNeuLm)
                        If (HatPreisKontrolle%(ww.WuNeuLm)) Then ww.WuStatus = 1
                        ww.WuLm = ww.WuLm + ww.WuNeuLm
                        ww.LmAnzGebucht = ww.LmAnzGebucht + ww.WuNeuLm
                        ww.WuNeuLm = 0
                        ww.LmStatus = 0
                    End If
                    If (ww.WuStatus = 0) Then
                        If ((ww.WuRm = ww.WuBm) And (ww.WuLm = (ww.WuBm + ww.WuNm))) Or _
                           ((ww.WuRm = ww.WuNeuZiel) And (ww.WuLm = (ww.WuNeuZiel + ww.WuNm))) Then
                            ww.WuStatus = 2
                        End If
                    End If
                Else
                    Call WuBuchen(RTrim$(.TextMatrix(.row, 0)), Lm% - BuchM%)
                    ww.LmAnzGebucht = Lm%
                    ww.LmStatus = 2
                    If (HatPreisKontrolle%(ww.WuLm)) Then
                        ww.WuStatus = 1
                    Else
                        ww.WuStatus = 2
                    End If
                    If (RTrim$(.TextMatrix(.row, 14)) <> "") Then
                        If (InStr(para.Benutz, "Y") > 0) Then
                            besMenge% = Lm%     '- BuchM%
                            If (HatPreisKontrolle%(ww.WuLm)) Then
                                besStatus% = 1
                            Else
                                besStatus% = 5  '1 or 4
                            End If
                            Call BesorgerFertig(Val(RTrim$(.TextMatrix(.row, 21))), KorrPzn$, besStatus%, besMenge%)
                        End If
                    End If
                End If
                ww.PutRecord (rInd% + 1)
                Call AuslesenDateiSatz(rInd%, True)
            End If
        End If
        ww.SatzUnLock (1)
    End With
ElseIf (h$ = TEXT_WU_PREIS) Then
    If (DarfPreisKalk% = False) Then
        Call iMsgBox("Aktion nicht durchführbar, da die Tätigkeit 'Preiskalkulation' für Sie derzeit gesperrt ist !", vbInformation)
    Else
        With ParentForm.flxarbeit(0)
            If (Trim$(.TextMatrix(.row, 23)) = "") Then
                Call NNAEPRechnen(.row)
            End If
        End With
        
        PreisKalkModus% = 0
        frmWuPreisKalk.Show 1
        Call AuslesenWu
    End If
End If

Call DefErrPop
End Sub

Sub PreisKalkBefuellen()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("PreisKalkBefuellen")
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
Dim j%, l%, rInd%, PosLag%
Dim tAvp#, xAep#, xavp#
Dim pzn$, SQLStr$, h$, h2$, text$
Dim ArtName$, ArtMenge$, ArtMeh$, LifDat$

Dim i%, AltRow%

If (PreisKalkModus% = 0) Then       'aktuelle Zeile
    rInd% = SucheFlexZeile(False)
    If (rInd% > 0) Then
        Call PreisKalkZeile(rInd%)
    End If
Else
    With ParentForm.flxarbeit(0)
        .redraw = False
        AltRow% = .row
        For i% = 1 To (.Rows - 1)
            If (.TextMatrix(i%, 1) = "$") Then
                .row = i%
                rInd% = SucheFlexZeile(False)
                If (rInd% > 0) Then
                    Call PreisKalkZeile(rInd%)
                End If
            End If
        Next i%
        .row = AltRow%
        .redraw = True
    End With
End If

Call DefErrPop
End Sub

Sub PreisKalkZeile(rInd%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("PreisKalkZeile")
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
Dim j%, l%, ind%, PosLag%, PkInd%, iPkInd%
Dim tAep#, tAvp#, xAep#, xavp#, KalkPreis#, KalkAvp#
Dim pzn$, SQLStr$, h$, h2$, text$
Dim ArtName$, ArtMenge$, ArtMeh$, KalkText$, LifDat$, mw$, RundungsKz$, EuroKz$

pzn$ = ww.pzn

tAep# = 0
tAvp# = 0
SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + pzn$
Set TaxeRec = TaxeDB.OpenRecordset(SQLStr$)
If (TaxeRec.EOF = False) Then
    tAep# = TaxeRec!EK / 100
    tAvp# = TaxeRec!VK / 100
    mw$ = "2"
    If (TaxeRec!MWSTKz) Then mw$ = "1"
End If


xAep# = 0#
xavp# = ww.WuAVP
If (pzn$ <> "9999999") Then
    FabsErrf% = ast.IndexSearch(0, pzn$, FabsRecno&)
    If (FabsErrf% = 0) Then
        ast.GetRecord (FabsRecno& + 1)
        xAep# = ww.WuAEP    ' ast.aep
        '2.96 AVP in Stammdaten meist aktueller als in WÜ-Satz
        If (ast.AVP <> 0) Then xavp# = ast.AVP
        mw$ = ast.mw
    End If
End If

PosLag% = 0
PkInd% = 0
FabsErrf% = ass.IndexSearch(0, pzn$, FabsRecno&)
If (FabsErrf% = 0) Then
    Call ass.GetRecord(FabsRecno& + 1)
    PosLag% = ass.PosLag
    PkInd% = ass.pk
    If (PkInd% = 8224) Then PkInd% = 0
End If
If (ww.auto = "v") And (AufschlagsTabelle(MAX_AUFSCHLAEGE - 1).PreisBasis <> 0) Then PkInd% = MAX_AUFSCHLAEGE

text$ = ww.txt
If (pzn$ = "9999999") Then
    h$ = RTrim$(text$)
    Call OemToChar(h$, h$)
    ArtName$ = h$
    ArtMenge$ = ""
    ArtMeh$ = ""
Else
    h$ = RTrim(Left$(text$, 33))
    Call OemToChar(h$, h$)
    l% = Len(h$)
    For j% = l% To 2 Step -1
        h2$ = Mid$(h$, j%, 1)
        If (h2$ = " ") Then Exit For
        If (InStr(".xX", h2$) > 0) Then
            h2$ = Mid$(h$, j% - 1, 1)
            If (InStr("0123456789", h2$) <= 0) Then
                Exit For
            End If
        End If
    Next j%
    If (j% > 2) Then
        ArtName$ = Left$(h$, j%)
        ArtMenge$ = Mid$(h$, j% + 1)
    Else
        ArtName$ = h$
        ArtMenge$ = ""
    End If
    ArtMeh$ = Mid$(text$, 34)
End If

'PkInd% = Asc(Right$(pzn$, 1)) - 48

'If (PkInd% = 0) Then
'    KalkPreis# = 0#
'    KalkText$ = ""
'Else
'    ind% = AufschlagsTabelle(PkInd% - 1).Preisbasis
'    If (ind% > 0) Then
'        If (ind% = 1) Then
'            KalkText$ = "NN-AEP"
'            KalkPreis# = ww.WuNNAep
'        ElseIf (ind% = 2) Then
'            KalkText$ = "Stamm-AEP"
'            KalkPreis# = xaep#
'        Else
'            KalkText$ = "Taxe-AEP"
'            KalkPreis# = tAep#
'        End If
'        If (AufschlagsTabelle(PkInd% - 1).Aufschlag) Then
'            KalkPreis# = KalkPreis# * (100# + AufschlagsTabelle(PkInd% - 1).Aufschlag) / 100#
'            KalkText$ = KalkText$ + " +" + Str$(AufschlagsTabelle(PkInd% - 1).Aufschlag) + "%  "
'        End If
'    End If
'End If

If (WuInDm% > 0) Then
    tAep# = FNX(tAep# * EURO_FAKTOR)
    tAvp# = FNX(tAvp# * EURO_FAKTOR)
    xavp# = FNX(xavp# * EURO_FAKTOR)
End If


h$ = pzn$ + vbTab + vbTab + ArtName$ + vbTab + ArtMenge$ + vbTab + ArtMeh$ + vbTab
h$ = h$ + Format(xAep#, "0.00") + vbTab + Format(xavp#, "0.00") + vbTab
h$ = h$ + vbTab
If (KalkPreis# <> 0#) Then h$ = h$ + Format(KalkPreis#, "0.00")
h$ = h$ + vbTab
If (KalkText$ <> "") Then h$ = h$ + KalkText$
h$ = h$ + vbTab
h$ = h$ + Format(tAvp#, "0.00") + vbTab + Format(PosLag%, "0") + vbTab
If (ww.auto = "v") Then h$ = h$ + "v"
h$ = h$ + vbTab + ast.wg + vbTab + ast.ka + vbTab

'    LifDat$ = wu.li + wu.bd + wu.zeit
'    LifDat$ = LifDat$ + Format(wu.WuLaufNr, "00000")
'    h$ = h$ + LifDat$ + vbTab

h$ = h$ + Str$(ww.BekLaufNr) + vbTab

h2$ = Format(rInd%, "00000")
h$ = h$ + h2$ + vbTab

iPkInd% = PkInd%
RundungsKz$ = ""
If (iPkInd% > 1000) Then
    RundungsKz$ = "*"
    iPkInd% = iPkInd% - 1000
End If
EuroKz$ = ""
If (iPkInd% >= 500) Then
    EuroKz$ = "*"
    iPkInd% = iPkInd% - 500
End If
h$ = h$ + Format(iPkInd%, "0") + vbTab

h$ = h$ + Format(ww.WuNNAep, "0.00") + vbTab
h$ = h$ + Format(tAep#, "0.00") + vbTab
h$ = h$ + mw$ + vbTab
h$ = h$ + Format(xavp#, "0.00") + vbTab

h$ = h$ + RundungsKz$ + vbTab
h$ = h$ + EuroKz$ + vbTab

With frmWuPreisKalk.flxarbeit(0)
'    .AddItem h$
    'so umständlich wegen fehler 'typen unverträglich' !
    .AddItem "a"
    .row = .Rows - 1
    For j% = 0 To 30
        If (h$ = "") Then Exit For
        ind% = InStr(h$, vbTab)
        If (ind% > 0) Then
            h2$ = Left$(h$, ind% - 1)
            h$ = Mid$(h$, ind% + 1)
        Else
            h2$ = h$
            h$ = ""
        End If
        .TextMatrix(.row, j%) = h2$
    Next j%

    Call frmWuPreisKalk.AvpKalkulation(PkInd%, KalkAvp#, KalkText$)
    .TextMatrix(.row, 9) = KalkText$
    If (Trim(KalkText$) = "") Then
        .TextMatrix(.row, 8) = ""
        .TextMatrix(.row, 7) = ""
    Else
        .TextMatrix(.row, 8) = Format(KalkAvp#, "0.00")
        .col = 7
        If (PkInd% > 1000) Then
            .TextMatrix(.row, 7) = .TextMatrix(.row, 8)
            .CellFontItalic = True
        Else
            .TextMatrix(.row, 7) = frmWuPreisKalk.PruefeRundung
            .CellFontItalic = False
        End If
    End If
End With


Call DefErrPop
End Sub

Function AvpKalkulation#(PkInd%, NnAep1#, StammAep1#, TaxeAep1#, Mw1$, KalkText$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("AvpKalkulation#")
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
Dim ind%, aufschl%, mw%, KalkAvp#

KalkAvp# = 0#
KalkText$ = ""

If (PkInd% > 0) Then
    ind% = AufschlagsTabelle(PkInd% - 1).PreisBasis
    If (ind% > 0) Then
        If (ind% = 1) Then
            KalkText$ = "NN-AEP"
            KalkAvp# = NnAep1#
        ElseIf (ind% = 2) Then
            KalkText$ = "Stamm-AEP"
            KalkAvp# = StammAep1#
        Else
            KalkText$ = "Taxe-AEP"
            KalkAvp# = TaxeAep1#
        End If
        
        aufschl% = AufschlagsTabelle(PkInd% - 1).Aufschlag
        mw% = para.Mwst(Val(Mw1$))
        If (aufschl% = 211) Then
            KalkAvp# = CalcAMPV(KalkAvp#, mw%)
            KalkText$ = KalkText$ + " + AMPV"
        Else
            KalkAvp# = KalkAvp# * (100# + aufschl%) / 100#
            KalkAvp# = FNX(KalkAvp# * (100# + mw%) / 100#)
            If (aufschl%) Then KalkText$ = KalkText$ + " +" + Str$(aufschl%) + "%  "
        End If
    End If
End If

AvpKalkulation# = KalkAvp#

Call DefErrPop
End Function

'Sub AvpKalkulation(PkInd%, KalkAvp#, KalkText$)
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Call DefErrFnc("AvpKalkulation")
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
'Dim ind%, aufschl%, mw%
'Dim KalkBasis$
'
'KalkAvp# = ActProgram.AvpKalkulation(PkInd%, iNnAep#, iStammAep#, iTaxeAep#, iMw$, KalkText$)
'
'KalkAvp# = 0#
'KalkText$ = ""
'
'If (PkInd% > 0) Then
'    With frmWuPreisKalk.flxWuPreisKalk
'        ind% = AufschlagsTabelle(PkInd% - 1).Preisbasis
'        If (ind% > 0) Then
'            If (ind% = 1) Then
'                KalkText$ = "NN-AEP"
'                KalkBasis$ = .TextMatrix(.row, 18)
'            ElseIf (ind% = 2) Then
'                KalkText$ = "Stamm-AEP"
'                KalkBasis$ = .TextMatrix(.row, 5)
'            Else
'                KalkText$ = "Taxe-AEP"
'                KalkBasis$ = .TextMatrix(.row, 19)
'            End If
'
'            KalkAvp# = CDbl(KalkBasis$)
'            aufschl% = AufschlagsTabelle(PkInd% - 1).Aufschlag
'            mw% = para.mwst(Val(.TextMatrix(.row, 20)))
'            If (aufschl% = 211) Then
'                KalkAvp# = CalcAMPV(KalkAvp#, mw%)
'                KalkText$ = KalkText$ + " + AMPV"
'            Else
'                KalkAvp# = KalkAvp# * (100# + aufschl%) / 100#
'                KalkAvp# = FNX(KalkAvp# * (100# + mw%) / 100#)
'                If (aufschl%) Then KalkText$ = KalkText$ + " +" + Str$(aufschl%) + "%  "
'            End If
'        End If
'    End With
'End If
'
'Call DefErrPop
'End Sub

Function PruefeRundung#(OrgAvp1#, KalkAvp1#)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("PruefeRundung#")
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
Dim i%, kGueltig%, ind%, bedTyp%, diff%, Bed1Ok%, Bed2Ok%, BereichOk%, AktiverBereich%
Dim val1%, val2%, rWert%
Dim OrgAvp&, KalkAvp&
Dim dVal2#
Dim wert1$, wert2$, op$, RundWert$, h$

OrgAvp& = CLng(OrgAvp1# * 100#)
KalkAvp& = CLng(KalkAvp1# * 100#)


Bed1Ok% = True
Bed2Ok% = True

BereichOk% = True
AktiverBereich% = False

For i% = 0 To (AnzRundungen% - 1)
    kGueltig% = False
    
    wert1$ = UCase(RTrim(Rundungen(i%).bedingung.wert1))
    op$ = UCase(RTrim(Rundungen(i%).bedingung.op))
    wert2$ = UCase(RTrim(Rundungen(i%).bedingung.wert2))
    RundWert$ = UCase(RTrim(Rundungen(i%).Gerundet))
    
    bedTyp% = ZeilenTyp%(wert1$)
    
'    If (BereichOk%) And (bedTyp% >= 6) And (i% > 0) Then
    If (BereichOk%) And (AktiverBereich%) And (bedTyp% >= 6) Then
        Exit For
    End If
    
    If (BereichOk% = 0) And (bedTyp% < 6) Then
    Else
    
        If (bedTyp% = 2) Then
            ' "ABBRUCH BEI VERTEUERUNG (%)"
            If (KalkAvp& > OrgAvp&) And (OrgAvp& > 0) Then
                kGueltig% = True
                val1% = Int(CDbl(KalkAvp&) / CDbl(OrgAvp&) * 100# - 100# + 0.501)
                val2% = Val(wert2$)
            End If
        ElseIf (bedTyp% = 3) Then
            ' "ABBRUCH BEI VERTEUERUNG (DM)"
            If (KalkAvp& > OrgAvp&) Then
                kGueltig% = True
                val1% = KalkAvp& - OrgAvp&
                val2% = Int(Val(wert2$) * 100# + 0.501) 'früher: cdbl
            End If
        ElseIf (bedTyp% = 4) Then
            ' "ABBRUCH BEI VERBILLIGUNG (%)"
            If (KalkAvp& < OrgAvp&) And (OrgAvp& > 0) Then
                kGueltig% = True
                val1% = Int(Abs(CDbl(KalkAvp&) / CDbl(OrgAvp&) * 100# - 100#) + 0.501)
                val2% = Val(wert2$)
            End If
        ElseIf (bedTyp% = 5) Then
            ' "ABBRUCH BEI VERBILLIGUNG(DM)"
            If (KalkAvp& < OrgAvp&) Then
                kGueltig% = True
                val1% = OrgAvp& - KalkAvp&
                val2% = Int(Val(wert2$) * 100# + 0.501) 'früher: cdbl
            End If
        ElseIf (bedTyp% = 6) Then
            ' "WENN"
            dVal2# = Val(wert2$)
            BereichOk% = False
            Select Case op$
                Case "<"
                    If (KalkAvp1# < dVal2#) Then BereichOk% = True
                Case "<="
                    If (KalkAvp1# <= dVal2#) Then BereichOk% = True
                Case "="
                    If (KalkAvp1# = dVal2#) Then BereichOk% = True
                Case "<>"
                    If (KalkAvp1# <> dVal2#) Then BereichOk% = True
                Case ">="
                    If (KalkAvp1# >= dVal2#) Then BereichOk% = True
                Case ">"
                    If (KalkAvp1# > dVal2#) Then BereichOk% = True
            End Select
            AktiverBereich% = BereichOk%
        ElseIf (bedTyp% = 7) Then
            ' "SONST"
            BereichOk% = True
            AktiverBereich% = True
        Else
            AktiverBereich% = True
            If ((bedTyp% = 0) And (Bed1Ok%)) Or ((bedTyp% = 1) And (Bed2Ok%)) Then
                kGueltig% = True
                val2% = Val(wert2$)
                
                If (bedTyp% = 0) Then
                    val1% = KalkAvp& Mod 10
                ElseIf (bedTyp% = 1) Then
                    val1% = KalkAvp& Mod 100
                End If
            End If
        End If
    
    End If
            
        
    If (kGueltig%) Then
        kGueltig% = False
        Select Case op$
            Case "<"
                If (val1% < val2%) Then kGueltig% = True
            Case "<="
                If (val1% <= val2%) Then kGueltig% = True
            Case "="
                If (val1% = val2%) Then kGueltig% = True
            Case "<>"
                If (val1% <> val2%) Then kGueltig% = True
            Case ">="
                If (val1% >= val2%) Then kGueltig% = True
            Case ">"
                If (val1% > val2%) Then kGueltig% = True
        End Select
    End If
        
    If (kGueltig%) Then
        rWert% = Int(Val(RundWert$) * 100# + 0.501)
        If (bedTyp% = 0) Then
            Bed1Ok% = False
            diff% = rWert% - 110
            If (diff% >= 10) Then
                KalkAvp& = KalkAvp& + 10
            ElseIf (diff% < 0) Then
                KalkAvp& = KalkAvp& - 10
            End If
            KalkAvp& = (KalkAvp& \ 10) * 10 + (rWert% Mod 10)
        ElseIf (bedTyp% = 1) Then
            Bed2Ok% = False
            diff% = rWert% - 100
            If (diff% >= 100) Then
                KalkAvp& = KalkAvp& + 100
            ElseIf (diff% < 0) Then
                KalkAvp& = KalkAvp& - 100
            End If
            KalkAvp& = (KalkAvp& \ 100) * 100 + (rWert% Mod 100)
        Else
            KalkAvp& = 0
            Exit For
        End If
    End If
    
Next i%

PruefeRundung# = KalkAvp& / 100#

Call DefErrPop
End Function

'Function PruefeRundung$()
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Call DefErrFnc("PruefeRundung$")
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
'Dim i%, kGueltig%, ind%, bedTyp%, diff%, Bed1Ok%, Bed2Ok%
'Dim val1%, val2%, rWert%
'Dim OrgAvp&, KalkAvp&
'Dim wert1$, wert2$, op$, RundWert$, h$
'
'With frmWuPreisKalk.flxWuPreisKalk
'    OrgAvp& = CLng(CDbl(.TextMatrix(.row, 6)) * 100#)
'    KalkAvp& = CLng(CDbl(.TextMatrix(.row, 8)) * 100#)
'End With
'
'
'Bed1Ok% = True
'Bed2Ok% = True
'
'For i% = 0 To (AnzRundungen% - 1)
'    kGueltig% = False
'
'    wert1$ = UCase(RTrim(Rundungen(i%).bedingung.wert1))
'    op$ = UCase(RTrim(Rundungen(i%).bedingung.op))
'    wert2$ = UCase(RTrim(Rundungen(i%).bedingung.wert2))
'    RundWert$ = UCase(RTrim(Rundungen(i%).Gerundet))
'
'    bedTyp% = ZeilenTyp%(wert1$)
'
'    If (bedTyp% = 2) Then
'        ' "ABBRUCH BEI VERTEUERUNG (%)"
'        If (KalkAvp& > OrgAvp&) And (OrgAvp& > 0) Then
'            kGueltig% = True
'            val1% = Int(CDbl(KalkAvp&) / CDbl(OrgAvp&) * 100# - 100# + 0.501)
'            val2% = Val(wert2$)
'        End If
'    ElseIf (bedTyp% = 3) Then
'        ' "ABBRUCH BEI VERTEUERUNG (DM)"
'        If (KalkAvp& > OrgAvp&) Then
'            kGueltig% = True
'            val1% = KalkAvp& - OrgAvp&
'            val2% = Int(CDbl(wert2$) * 100# + 0.501)
'        End If
'    ElseIf (bedTyp% = 4) Then
'        ' "ABBRUCH BEI VERBILLIGUNG (%)"
'        If (KalkAvp& < OrgAvp&) And (OrgAvp& > 0) Then
'            kGueltig% = True
'            val1% = Int(Abs(CDbl(KalkAvp&) / CDbl(OrgAvp&) * 100# - 100#) + 0.501)
'            val2% = Val(wert2$)
'        End If
'    ElseIf (bedTyp% = 5) Then
'        ' "ABBRUCH BEI VERBILLIGUNG(DM)"
'        If (KalkAvp& < OrgAvp&) Then
'            kGueltig% = True
'            val1% = OrgAvp& - KalkAvp&
'            val2% = Int(CDbl(wert2$) * 100# + 0.501)
'        End If
'
'    ElseIf ((bedTyp% = 0) And (Bed1Ok%)) Or ((bedTyp% = 1) And (Bed2Ok%)) Then
'
'        kGueltig% = True
'        val2% = Val(wert2$)
'
'        If (bedTyp% = 0) Then
'            val1% = KalkAvp& Mod 10
'        ElseIf (bedTyp% = 1) Then
'            val1% = KalkAvp& Mod 100
'        End If
'    End If
'
'
'    If (kGueltig%) Then
'        kGueltig% = False
'        Select Case op$
'            Case "<"
'                If (val1% < val2%) Then kGueltig% = True
'            Case "<="
'                If (val1% <= val2%) Then kGueltig% = True
'            Case "="
'                If (val1% = val2%) Then kGueltig% = True
'            Case "<>"
'                If (val1% <> val2%) Then kGueltig% = True
'            Case ">="
'                If (val1% >= val2%) Then kGueltig% = True
'            Case ">"
'                If (val1% > val2%) Then kGueltig% = True
'        End Select
'    End If
'
'    If (kGueltig%) Then
'        rWert% = Int(Val(RundWert$) * 100# + 0.501)
'        If (bedTyp% = 0) Then
'            Bed1Ok% = False
'            diff% = rWert% - 110
'            If (diff% >= 10) Then
'                KalkAvp& = KalkAvp& + 10
'            ElseIf (diff% < 0) Then
'                KalkAvp& = KalkAvp& - 10
'            End If
'            KalkAvp& = (KalkAvp& \ 10) * 10 + (rWert% Mod 10)
'        ElseIf (bedTyp% = 1) Then
'            Bed2Ok% = False
'            diff% = rWert% - 100
'            If (diff% >= 100) Then
'                KalkAvp& = KalkAvp& + 100
'            ElseIf (diff% < 0) Then
'                KalkAvp& = KalkAvp& - 100
'            End If
'            KalkAvp& = (KalkAvp& \ 100) * 100 + (rWert% Mod 100)
'        Else
'            KalkAvp& = 0
'            Exit For
'        End If
'    End If
'
'Next i%
'
'If (KalkAvp& > 0) Then
'    PruefeRundung$ = Format(KalkAvp& / 100#, "0.00")
'Else
'    PruefeRundung$ = ""
'End If
'
'Call DefErrPop
'End Function

Function PruefeAutomaticPreis$(RundAvp#)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("PruefeAutomaticPreis$")
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
Dim PkInd%, MitRundung%
Dim tAep#, xAep#, KalkAvp#
Dim ret$, pzn$, SQLStr$, mw$, KalkText$

If (para.Land = "A") Then
    PruefeAutomaticPreis$ = ""
    Call DefErrPop: Exit Function
End If


ret$ = ""
RundAvp# = 0#

pzn$ = ww.pzn

PkInd% = 0
FabsErrf% = ass.IndexSearch(0, pzn$, FabsRecno&)
If (FabsErrf% = 0) Then
    Call ass.GetRecord(FabsRecno& + 1)
    PkInd% = ass.pk
    If (PkInd% = 8224) Then PkInd% = 0
End If
If (ww.auto = "v") And (AufschlagsTabelle(MAX_AUFSCHLAEGE - 1).PreisBasis <> 0) Then PkInd% = MAX_AUFSCHLAEGE

'If (PkInd% > 0) Then PkInd% = PkInd% + 1000

MitRundung% = True
If (PkInd% > 1000) Then
    PkInd% = PkInd% - 1000
    MitRundung% = False
End If
If (PkInd% >= 500) Then
    PkInd% = PkInd% - 500
End If

If (PkInd% > 0) And (PkInd% < 100) Then
    tAep# = 0
    SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + pzn$
    Set TaxeRec = TaxeDB.OpenRecordset(SQLStr$)
    If (TaxeRec.EOF = False) Then
        tAep# = TaxeRec!EK / 100
        mw$ = "2"
        If (TaxeRec!MWSTKz) Then mw$ = "1"
    End If
    
    xAep# = 0#
    If (pzn$ <> "9999999") Then
        FabsErrf% = ast.IndexSearch(0, pzn$, FabsRecno&)
        If (FabsErrf% = 0) Then
            ast.GetRecord (FabsRecno& + 1)
            xAep# = ww.WuAEP    ' ast.aep
            mw$ = ast.mw
        End If
    End If

    KalkAvp# = AvpKalkulation(PkInd%, ww.WuNNAep, xAep#, tAep#, mw$, KalkText$)
    If (MitRundung%) Then
        RundAvp# = PruefeRundung(ww.WuAVP, KalkAvp#)
    Else
        RundAvp# = KalkAvp#
    End If
    If (RundAvp# <> 0#) Then
        ret$ = "Preisvorschlag:" + vbCr + Format(KalkAvp#, "0.00") + "  (" + KalkText$ + ")"
        If (MitRundung%) Then
            ret$ = ret$ + vbCr + Format(RundAvp#, "0.00") + "  (gerundet)"
        End If
    End If
End If

PruefeAutomaticPreis$ = ret$

Call DefErrPop
End Function

Sub SpeicherKalkPreise(rInd%, Optional PkInd% = 0)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("SpeicherKalkPreise")
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
Dim besStatus%, TuEs%
Dim dWuAep#, dWuAvp#
Dim pzn$
                
ww.WuStatus = 2
ww.PutRecord (rInd% + 1)

dWuAep# = ww.WuAEP
dWuAvp# = ww.WuAVP
If (WuInDm% > 0) Then
    dWuAep# = FNX(dWuAep# / EURO_FAKTOR)
    dWuAvp# = FNX(dWuAvp# / EURO_FAKTOR)
End If
    
If (ww.auto = "v") And (InStr(para.Benutz, "Y") > 0) Then
    besStatus% = 6  '2 or 4
    Call BesorgerFertig(Val(ww.AbholNr), ww.pzn, besStatus%, 0, dWuAvp#)
'    Call BesorgerFertig(Val(ww.AbholNr), ww.pzn, besStatus%, 0, ww.WuAVP)
End If

pzn$ = ww.pzn
If (Val(pzn$) <> 0) And (pzn$ <> "9999999") Then
    FabsErrf% = ast.IndexSearch(0, pzn$, FabsRecno&)
    If (FabsErrf% = 0) Then
        ast.GetRecord (FabsRecno& + 1)
        TuEs% = False
'        If (ww.WuAep <> ast.aep Or ww.WuAVP <> ast.AVP) Then
'            If (ww.WuAep <> 0#) Then ast.aep = ww.WuAep
'            ast.AVP = ww.WuAVP
'            TuEs% = True
'        End If
        If (dWuAep# <> ast.aep Or dWuAvp# <> ast.AVP) Then
            If (dWuAep# <> 0#) Then ast.aep = dWuAep#
            ast.AVP = dWuAvp#
            TuEs% = True
        End If
        If (PkInd% > 0) Then
            ast.ka = "1"
            TuEs% = True
        End If
        If (TuEs%) Then ast.PutRecord (FabsRecno& + 1)
    End If
    If (PkInd% > 0) Then
        FabsErrf% = ass.IndexSearch(0, pzn$, FabsRecno&)
        If (FabsErrf% = 0) Then
            ass.GetRecord (FabsRecno& + 1)
            If (ass.pk <> PkInd%) Then
                ass.pk = PkInd%
                ass.PutRecord (FabsRecno& + 1)
            End If
        End If
    End If
End If

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

Call AuslesenWu

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
Dim rInd%
Dim RundAvp#
Dim h$

'Debug.Print Format(KeyAscii, "000")

If (KeyAscii = vbKeySpace) Then
    With frmAction.flxarbeit(0)
        If (.TextMatrix(.row, 1) = "$") And (frmAction.picQuittieren.Visible) Then
            ww.SatzLock (1)
            rInd% = SucheFlexZeile(True)
            If (rInd% > 0) Then
                h$ = PruefeAutomaticPreis(RundAvp#)
                If (h$ <> "") Then
                    ww.WuAVP = RundAvp#
                    Call SpeicherKalkPreise(rInd%)
                    Call AuslesenDateiSatz(rInd%, True)
                End If
            End If
            ww.SatzUnLock (1)
        End If
    End With
ElseIf ((KeyAscii >= 48) And (KeyAscii <= 57)) Then
    EingabeStr$ = EingabeStr$ + Chr$(KeyAscii)
ElseIf (InStr("ABCDEFGHIJKLMNOPQRSTUVWXYZ", UCase(Chr$(KeyAscii))) > 0) Then
    Call frmAction.SelectZeile(UCase(Chr$(KeyAscii)))
End If
Call DefErrPop
End Sub

Sub SucheZeile(Optional Dp04Aktiv% = False, Optional Dp04Menge%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("SucheZeile")
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
Dim i%, rInd%, gef%, AutomaticIncr%, IstPzn%, hErg%, ret%, sMax%
Dim AltRecno&, AltNnSatz&
Dim AltAep#, AltAvp#, AltKp#, AbholerErsatzPreis#
Dim pzn$, txt$, SQLStr$, OrgEingabeStr$, OrgEan$, mErg$, lDat$, AltPzn$, AltEan$, mc$, vnum$
Dim AbholerErsatzName$
Dim MatchArt As clsMatchArtikel

If (OrgBeleg$ <> "") Then
    ActBeleg$ = OrgBeleg$
    ActBelegDatum% = OrgBelegDatum%
    OrgBeleg$ = ""
End If

With ParentForm.flxarbeit(0)
    If (Len(EingabeStr$) < 7) Then Call DefErrPop: Exit Sub

    If (Len(EingabeStr$) = 13) And ((Left$(EingabeStr$, 8) = "99999978") Or (Left$(EingabeStr$, 8) = "99999979")) Then
        If (Dp04Aktiv%) Then
            rInd% = SucheFlexZeile(False)
            If (rInd% > 0) Then
                txt$ = RTrim(ww.WuAblDatum)
                If (txt$ = "") Then txt$ = "011210"
            
                If (Left$(EingabeStr$, 8) = "99999978") Then
                    txt$ = Left$(txt$, 4) + Mid$(EingabeStr$, 11, 2)
                ElseIf (Left$(EingabeStr$, 8) = "99999979") Then
                    txt$ = Left$(txt$, 2) + Mid$(EingabeStr$, 11, 2) + Right$(txt$, 2)
                End If
                ww.WuAblDatum = txt$
                ww.PutRecord (rInd% + 1)
                Call AuslesenDateiSatz(rInd%, True)
            End If
        Else
            txt$ = EingabeStr$
            .col = 11
            Call EditSatz(txt$)
        End If
        Call DefErrPop: Exit Sub
    End If
    
    OrgEingabeStr$ = EingabeStr$
    EingabeStr$ = Right$(Space$(13) + EingabeStr$, 13)
    FabsErrf% = ass.IndexSearch(1, EingabeStr$, FabsRecno&)
    If (FabsErrf% = 0) Then
        ass.GetRecord (FabsRecno& + 1)
        EingabeStr$ = ass.pzn
    Else
        If (Len(OrgEingabeStr$) = 7) Then
            EingabeStr$ = OrgEingabeStr$
        ElseIf (para.Land = "D") Then
            EingabeStr$ = Mid$(EingabeStr$, 6, 7)
        ElseIf (Len(OrgEingabeStr$) = 13) Then
            EingabeStr$ = Mid$(EingabeStr$, 7, 6)
            Call PZNPrüfZiffer(EingabeStr$)
        Else
            EingabeStr$ = Mid$(EingabeStr$, 6, 7)
        End If
    End If
    
    
    If (EinzelneWu% = False) Then
        lDat$ = ""
        For i% = 1 To (.Rows - 1)
            If (.TextMatrix(i%, 0) = EingabeStr$) Then
                If (lDat$ = "") Then lDat$ = .TextMatrix(i%, 16)
                If (lDat$ <> .TextMatrix(i%, 16)) Then
                    Call MenuBearbeiten(MENU_SF3)
                    Call DefErrPop: Exit Sub
                End If
            End If
        Next i%
    End If
    
    gef% = 0
    For i% = 1 To (.Rows - 1)
        If (.TextMatrix(i%, 0) = EingabeStr$) Then
            gef% = i%
            If (Trim$(.TextMatrix(i%, 19)) = "") Or (Val(.TextMatrix(i%, 8)) < Val(.TextMatrix(i%, 5)) + Val(.TextMatrix(i%, 6))) Then
                Exit For
            End If
        End If
    Next i%
        
    If (gef% = 0) Then
        IstPzn% = False
        
        FabsErrf% = ast.IndexSearch(0, EingabeStr$, FabsRecno&)
        If (FabsErrf% = 0) Then
            ast.GetRecord (FabsRecno& + 1)
            IstPzn% = True
        End If
        
        If (IstPzn% = False) Then
            SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + EingabeStr$
            Set TaxeRec = TaxeDB.OpenRecordset(SQLStr$)
            If (TaxeRec.EOF = False) Then
                IstPzn% = True
                Call Taxe2ast(EingabeStr$)
            End If
        End If
        
        pzn$ = ast.pzn
        
        If (IstPzn%) Then
            WuFrageTxt$ = Trim(ast.kurz) + " " + ast.meng + ast.meh + " gehört nicht zu dieser Lieferung!"
            AbholerErsatzName$ = ast.kurz + ast.meng + " " + ast.meh
            AbholerErsatzPreis# = ast.AVP
        Else
            WuFrageTxt$ = "Strichcode " + OrgEingabeStr$ + " ist der EDV unbekannt!"
        End If
        
        WuFrageErg% = 0
        frmWuFrage.Show 1
        
        If (WuFrageErg% = 0) Then
            If (IstPzn%) Then
                Call ManuellErfassen(pzn$, "", False, 1, 0)
                Call AuslesenWu
                .col = 5
                SuperEdit% = True
                Call EditSatz
                .col = 6
                SuperEdit% = True
                Call EditSatz
                gef% = .row
            Else
                mErg$ = ArtikelAuswahl$
                If (mErg$ <> "") Then
                    Call AuslesenWu
                    gef% = .row
                Else
                    Call DefErrPop: Exit Sub
                End If
            End If
        ElseIf (WuFrageErg% = 1) Or (WuFrageErg% = 2) Then
            StrichCodeErg$ = "Zu ersetzenden Artikel wählen"
            frmStrichCode.Show 1
            If (StrichCodeErg$ <> "") Then
                If (IstPzn% = False) Then
                    mErg$ = ArtikelAuswahl$(False)
                    If (mErg$ <> "") Then
                        pzn$ = mErg$
                    Else
                        Call DefErrPop: Exit Sub
                    End If
                End If
            
                gef% = Val(StrichCodeErg$)
                
                If (WuFrageErg% = 2) Then
                
                    AltPzn$ = .TextMatrix(gef%, 0)
                    FabsErrf% = ass.IndexSearch(0, AltPzn$, FabsRecno&)
                    If (FabsErrf% = 0) Then
                        ass.GetRecord (FabsRecno& + 1)
                        AltEan$ = ass.ean
                    Else
                        Call iMsgBox("Aktion abgebrochen: zu ersetzender Artikel ist kein Lagerartikel !", vbInformation)
                        Call DefErrPop: Exit Sub
                    End If
                    
                    Set MatchArt = New clsMatchArtikel
                    Call MatchArt.InitMatch
                    
                    WuFrageTxt$ = "PZN " + pzn$
                    hErg% = True
                    SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + pzn$
                    Set TaxeRec = TaxeDB.OpenRecordset(SQLStr$)
                    If (TaxeRec.EOF = False) Then
                '        If (InStr("23", Chr$(TaxeRec!AbgabeBest)) = 0) And (TaxeRec!EK > 0) And (TaxeRec!VK > 0) Then ret% = False
                        If (InStr("023478", Format(TaxeRec!AbgabeBest, "0")) = 0) And (TaxeRec!EK > 0) And (TaxeRec!VK > 0) Then hErg% = False
                        Call Taxe2ast(pzn$)
                        WuFrageTxt$ = Trim(ast.kurz) + " " + ast.meng + ast.meh
                    End If

                    FabsErrf% = ass.IndexSearch(0, pzn$, FabsRecno&)
                    If (FabsErrf% <> 0) Then
                        ret% = iMsgBox(WuFrageTxt$ + " auf Lager nehmen ?", vbYesNo Or vbDefaultButton2)
                        If (ret% <> vbYes) Then Call DefErrPop: Exit Sub
                        
                        ret% = InStammAufnehmen%(pzn$)
                        If (ret% = False) Then Call DefErrPop: Exit Sub
                    End If
                                        
                    Load frmWuAlternativ
    
                    frmWuAlternativ.chkWuAlternativ(0).Enabled = hErg%
                    If (hErg%) Then frmWuAlternativ.chkWuAlternativ(0).Value = 1
                    
                    hErg% = False
                    If (Len(Trim(AltEan$)) > 7) Then hErg% = True
                    frmWuAlternativ.chkWuAlternativ(1).Enabled = hErg%
                    
                    frmWuAlternativ.chkWuAlternativ(2).Value = 1
                    
                    frmWuAlternativ.Show 1
                    
                    If (WuFrageErg% < 0) Then Call DefErrPop: Exit Sub
                    
                    If (WuFrageErg% And 1) Then
                        FabsErrf% = ast.IndexSearch(0, AltPzn$, FabsRecno&)
                        If (FabsErrf% = 0) Then
                            ast.GetRecord (FabsRecno& + 1)
                            AltAep# = ast.aep
                            AltAvp# = ast.AVP
                            AltKp# = ast.kp
                            FabsErrf% = ast.IndexSearch(0, pzn$, FabsRecno&)
                            If (FabsErrf% = 0) Then
                                ast.GetRecord (FabsRecno& + 1)
                                If (AltAep# <> 0#) Then ast.aep = AltAep#
                                If (AltAvp# <> 0#) Then ast.AVP = AltAvp#
                                If (AltKp# <> 0#) Then ast.kp = AltKp#
                                ast.PutRecord (FabsRecno& + 1)
                            End If
                        End If
                    End If
                    
                    AltRecno& = 0&
                    FabsErrf% = ass.IndexSearch(0, pzn$, FabsRecno&)
                    If (FabsErrf% = 0) Then
                        AltRecno& = FabsRecno&
                        ass.GetRecord (FabsRecno& + 1)
                        AltEan$ = ass.ean
                    End If
                    
                    FabsErrf% = ass.IndexSearch(0, AltPzn$, FabsRecno&)
                    If (FabsErrf% = 0) Then
                        ass.GetRecord (FabsRecno& + 1)
                        ass.pzn = pzn$
                        If (AltRecno& > 0) Then
                            If (WuFrageErg% And 2) Then
                                FabsErrf% = ass.IndexReplace(1, AltRecno&, AltEan$, ass.ean, FabsRecno&)
                            Else
                                ass.ean = AltEan$
                            End If
                        Else
                            FabsErrf% = ass.IndexInsert(0, pzn$, FabsRecno&)
                            If (FabsErrf% <> 0) Then Call DefErrPop: Exit Sub
                            AltRecno& = FabsRecno&
                            If Not (WuFrageErg% And 2) Then ass.ean = Space$(13)
                            FabsErrf% = ass.IndexInsert(1, ass.ean, FabsRecno&)
                            If (FabsErrf% <> 0) Then Call DefErrPop: Exit Sub
                            If (FabsRecno& <> AltRecno&) Then Call DefErrPop: Exit Sub
                            
                            Call ass.GetRecord(1)
                            If (FabsRecno& > ass.erstmax) Then
                                ass.erstmax = FabsRecno&
                                Call ass.PutRecord(1)
                            End If
                        End If
                        ass.PutRecord (AltRecno& + 1)
                    
                        'NNAEP
                        If (para.Land <> "A") Then
                            AltNnSatz& = 0&
                            FabsErrf% = nnek.IndexSearch(0, AltPzn$, FabsRecno&)
                            If (FabsErrf% = 0) Then
                                AltNnSatz& = FabsRecno&
                                nnek.GetRecord (FabsRecno& + 1)
                                FabsErrf% = nnek.IndexSearch(0, pzn$, FabsRecno&)
                                If (FabsErrf% <> 0) Then
                                    FabsErrf% = nnek.IndexInsert(0, pzn$, FabsRecno&)
                                    If (FabsErrf% <> 0) Then Call DefErrPop: Exit Sub
                                End If
                                nnek.pzn = pzn$
                                nnek.PutRecord (FabsRecno& + 1)
                                
                                nnek.GetRecord (1)
                                If (FabsRecno& > nnek.erstmax) Then
                                    nnek.erstmax = FabsRecno&
                                    nnek.PutRecord (1)
                                End If
                            End If
                        End If
                    End If
                    
                    If (WuFrageErg% And 4) Then
                        FabsErrf% = ast.IndexSearch(0, AltPzn$, FabsRecno&)
                        If (FabsErrf% = 0) Then
                            ast.GetRecord (FabsRecno& + 1)
                            FabsErrf% = ast.IndexDelete(0, FabsRecno&, AltPzn$, FabsRecno&)
    
'                            mc$ = ast.kurz + ast.meng + ast.meh
'                            mc$ = Left$(MatchArt.MakeMatchCode(mc$, 2) + Space$(35), 35)
                            vnum$ = ast.Vnr:    vnum$ = MatchArt.MakeVnum(vnum$)
                            mc$ = ast.kurz + ast.meng + ast.meh
                            mc$ = MatchArt.MakeMatchCode(mc$, 2)
                            mc$ = Left$(mc$ + Space$(10), 10) + vnum$ + Mid$(mc$, 11)
                            FabsErrf% = ast.IndexDelete(1, FabsRecno&, mc$, FabsRecno&)
                            
                            ast.pzn = String(Len(ast.pzn), 48)
                            ast.kurz = "A" + Space$(Len(ast.kurz) - 1)
                            ast.meng = Space$(Len(ast.meng))
                            ast.meh = Space$(Len(ast.meh))
                            ast.Vnr = String(Len(ast.Vnr), 0)
                            ast.PutRecord (FabsRecno& + 1)
                        End If

                        FabsErrf% = ass.IndexSearch(0, AltPzn$, FabsRecno&)
                        If (FabsErrf% = 0) Then
                            ass.GetRecord (FabsRecno& + 1)
                            AltEan$ = ass.ean
                            ass.pzn = String(Len(ass.pzn), 48)
                            ass.ean = String(Len(ass.ean), 48)
                            ass.PutRecord (FabsRecno& + 1)
                            FabsErrf% = ass.IndexDelete(0, FabsRecno&, AltPzn$, FabsRecno&)
                            FabsErrf% = ass.IndexDelete(1, FabsRecno&, AltEan$, FabsRecno&)
                        End If
                        
                        'NNEK löschen
                        If (para.Land <> "A") Then
                            If (AltNnSatz& > 0) Then
                                nnek.GetRecord (AltNnSatz& + 1)
                                nnek.pzn = String$(Len(nnek.pzn), 0)
                                nnek.PutRecord (AltNnSatz& + 1)
                                FabsErrf% = nnek.IndexDelete(0, AltNnSatz&, AltPzn$, FabsRecno&)
                            End If
                        End If
                    End If
                End If
                
                .redraw = False
                
                .row = gef%
                rInd% = SucheFlexZeile(False)
                If (rInd% > 0) Then
                    ww.status = 0
                    ww.PutRecord (rInd% + 1)

                    If (ww.auto = "v") And (InStr(para.Benutz, "Y") > 0) Then
                        Call BesorgerFertig(Val(ww.AbholNr), ww.pzn, 10, 0, AbholerErsatzPreis#, pzn$ + AbholerErsatzName$)
                    End If
                    
                    Call ManuellErfassen(pzn$, "", False, ww.bm, 0, ww.AbholNr)
                    Call AuslesenWu
                    gef% = .row
                End If
            Else
                Call DefErrPop: Exit Sub
            End If
        ElseIf (WuFrageErg% = 3) Then
            StrichCodeErg$ = "Strichcode <" + OrgEingabeStr$ + "> zuordnen"
            frmStrichCode.Show 1
            If (StrichCodeErg$ <> "") Then
                gef% = Val(StrichCodeErg$)
                
                FabsErrf% = ass.IndexSearch(0, .TextMatrix(gef%, 0), FabsRecno&)
                If (FabsErrf% = 0) Then
                    ass.GetRecord (FabsRecno& + 1)
                    OrgEan$ = ass.ean
                    ass.ean = Right$(Space$(13) + OrgEingabeStr$, 13)
                    ass.PutRecord (FabsRecno& + 1)
                    AltRecno& = FabsRecno&
                    FabsErrf% = ass.IndexReplace(1, AltRecno&, OrgEan$, ass.ean, FabsRecno&)
                End If
            Else
                Call DefErrPop: Exit Sub
            End If
        ElseIf (WuFrageErg% = 4) Then
            ret% = IstInNachlieferung%(True, OrgEingabeStr$)
            NLbeleg$ = ActBeleg$
            NLbelegDatum% = ActBelegDatum%
            Call AuslesenWu
        Else
            Call DefErrPop: Exit Sub
        End If
    ElseIf (IstDirektLief%) Then
        pzn$ = EingabeStr$
        FabsErrf% = ass.IndexSearch(0, pzn$, FabsRecno&)
        If (FabsErrf% <> 0) Then
            WuFrageTxt$ = "PZN " + pzn$
            SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + pzn$
            Set TaxeRec = TaxeDB.OpenRecordset(SQLStr$)
            If (TaxeRec.EOF = False) Then
                Call Taxe2ast(pzn$)
                WuFrageTxt$ = Trim(ast.kurz) + " " + ast.meng + ast.meh
            
                ret% = iMsgBox(WuFrageTxt$ + " auf Lager nehmen ?", vbYesNo Or vbDefaultButton2)
                If (ret% = vbYes) Then
                    ret% = InStammAufnehmen%(pzn$)
                    If (ret% = False) Then Call DefErrPop: Exit Sub
                
                    ret% = InStatistikAufnehmen%(pzn$)
                    If (ret% = False) Then Call DefErrPop: Exit Sub
                End If
            End If
        End If
    End If
        
    If (gef% > 0) Then
        Call ParentForm.HighlightZeile(True)
        .row = gef%
        .col = 8
        
'        If (WuLifDat$ = Chr$(0)) Then
        If (EinzelneWu% = False) Then
            If (Val(.TextMatrix(.row, 24)) > 0) Then
                txt$ = "*"
            Else
                txt$ = " "
            End If
            WuLifDat$ = "@" + .TextMatrix(.row, 16) + txt$
            Call AuslesenWu(True)
        Else
            If (.row < .TopRow) Then
                .TopRow = .row
            Else
                If (.row >= (.TopRow + ParentBereich.ArbeitAnzZeilen - 2)) Then
                    .TopRow = .row - ParentBereich.ArbeitAnzZeilen + 2
                End If
        '        While ((.row - .TopRow) >= (ParentBereich.ArbeitAnzZeilen - 1))
        '            .TopRow = .TopRow + 1
        '        Wend
            End If
            Call ParentForm.HighlightZeile
            Call EchtKurzInfo
        End If
        
        
        AutomaticIncr% = False
        If (.TextMatrix(.row, 14) = "v") Then
'            Call ZeigeBesorgerDaten(1)
        ElseIf (Dp04Aktiv% = False) And (Val(.TextMatrix(.row, 5)) <= para.LangsamLM) Then
            AutomaticIncr% = True
        End If
            
        If (AutomaticIncr%) Or (Dp04Aktiv%) Then
            rInd% = SucheFlexZeile(False)
            If (rInd% > 0) Then
'                If (ww.WuBelegDatum = 0) Then
                If (IstAltLast% = False) Then
                    ww.WuBeleg = ActBeleg$
                    ww.WuBelegDatum = ActBelegDatum%
                End If
                
                If (AutomaticIncr%) Then
                    If (ww.LmStatus = 0) Then ww.WuLm = 0
                    ww.WuLm = ww.WuLm + 1
                Else
                    ww.WuLm = Dp04Menge%
                End If
                ww.LmStatus = 1
                ww.RmStatus = 0
                ww.WuStatus = 0
                ww.PutRecord (rInd% + 1)
                Call AuslesenDateiSatz(rInd%, True)
            End If
        Else
            Call EditSatz
        End If
            
        .col = 11
'        While (para.AblK = " ") And (.TextMatrix(.row, 15) = "A") And (RTrim(.TextMatrix(.row, 11)) = "")
        While (para.AblK = " ") And (Left$(.TextMatrix(.row, 15), 1) = "A") And (RTrim(.TextMatrix(.row, 11)) = "")
            Call EditSatz
        Wend
'    ElseIf (Len(EingabeStr$) = 7) Then
'        SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + EingabeStr$
'        Set TaxeRec = TaxeDB.OpenRecordset(SQLStr$)
'        If (TaxeRec.EOF = False) Then
'            pzn$ = TaxeRec!pzn
'            txt$ = Trim$(TaxeRec!Name) + "  " + Trim$(TaxeRec!menge) + " " + Trim$(TaxeRec!einheit)
'            Call ManuellErfassen(pzn$, txt$, True)
'        End If
    End If
End With

Call DefErrPop
End Sub

Function InStammAufnehmen%(pzn$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("InStammAufnehmen%")
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
Dim AltRecno&
Dim vnum$, mc$
Dim MatchArt As clsMatchArtikel
                        
InStammAufnehmen% = False

FabsErrf% = ast.IndexSearch(0, pzn$, FabsRecno&)
If (FabsErrf% <> 0) Then
    Set MatchArt = New clsMatchArtikel
    Call MatchArt.InitMatch
                    
    FabsErrf% = ast.IndexInsert(0, pzn$, FabsRecno&)
    If (FabsErrf% <> 0) Then Call DefErrPop: Exit Function
    AltRecno& = FabsRecno&
    
    vnum$ = ast.Vnr:    vnum$ = MatchArt.MakeVnum(vnum$)
    mc$ = ast.kurz + ast.meng + ast.meh
    mc$ = MatchArt.MakeMatchCode(mc$, 2)
    mc$ = Left$(mc$ + Space$(10), 10) + vnum$ + Mid$(mc$, 11)
    FabsErrf% = ast.IndexInsert(1, mc$, FabsRecno&)
    If (FabsErrf% <> 0) Then Call DefErrPop: Exit Function
    
    If (FabsRecno& <> AltRecno&) Then
        Call iMsgBox("Vorgang abgebrochen:" + vbCr + "Unterschiedliche Records im Stammdaten-Index" + vbCr + "Bitte reorganisieren ! ", vbOKOnly Or vbCritical, "FABSP")
        Call DefErrPop: Exit Function
    End If

    
    ast.PutRecord (FabsRecno& + 1)
    
    Call ast.GetRecord(1)
    If (FabsRecno& > ast.erstmax) Then
        ast.erstmax = FabsRecno&
        Call ast.PutRecord(1)
    End If
End If

InStammAufnehmen% = True

Call DefErrPop
End Function

Function InStatistikAufnehmen%(pzn$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("InStatistikAufnehmen%")
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
Dim AltRecno&
                        
InStatistikAufnehmen% = False

FabsErrf% = ass.IndexSearch(0, pzn$, FabsRecno&)
If (FabsErrf% <> 0) Then
    ass.InitStruct
    ass.pzn = pzn$
    FabsErrf% = ass.IndexInsert(0, pzn$, FabsRecno&)
    If (FabsErrf% <> 0) Then Call DefErrPop: Exit Function
    AltRecno& = FabsRecno&
    ass.ean = Space$(13)
    FabsErrf% = ass.IndexInsert(1, ass.ean, FabsRecno&)
    If (FabsErrf% <> 0) Then Call DefErrPop: Exit Function
    If (FabsRecno& <> AltRecno&) Then Call DefErrPop: Exit Function
    
    If (FabsRecno& <> AltRecno&) Then
        Call iMsgBox("Vorgang abgebrochen:" + vbCr + "Unterschiedliche Records im Statistik-Index" + vbCr + "Bitte reorganisieren ! ", vbOKOnly Or vbCritical, "FABSP")
        Call DefErrPop: Exit Function
    End If

    ass.PutRecord (FabsRecno& + 1)

    Call ass.GetRecord(1)
    If (FabsRecno& > ass.erstmax) Then
        ass.erstmax = FabsRecno&
        Call ass.PutRecord(1)
    End If
End If
                  
InStatistikAufnehmen% = True

Call DefErrPop
End Function

Sub LagerBew(pzn$, LagAlt%, LagNeu%, BenutzerNr%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("LagerBew")
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
Dim Zeiger!, Groesse!
Dim Lb1 As clsLagerBewegung

If (LagAlt% <> LagNeu%) And (InStr(para.Benutz, "V") > 0) Then
    Set Lb1 = New clsLagerBewegung
    Lb1.OpenDatei
    
    Lb1.datum = iDate(Format(Now, "DDMMYY"))
    Lb1.zeit = Int(((Val(Left$(Time$, 2)) * 60! + Val(Mid$(Time$, 4, 2))) * 60! + Val(Right$(Time$, 2))) / 3!)
    Lb1.pzn = pzn$
    Lb1.LagAlt = LagAlt%
    Lb1.LagNeu = LagNeu%
    Lb1.User = Val(para.User)
    Lb1.prog = 7
    Lb1.pers = BenutzerNr%

'  f.tmp% = 6
'  If ProgrammNr$ = "" Then
'    Open "MPOS.DAT" For Random Access Read Write Shared As #f.tmp% Len = 100
'    FIELD #f.tmp%, 2 AS F.PWCODE$, 2 AS F.Z$, 2 AS F.M$, 20 AS F.N$,2 AS F.FCOMP$, 2 AS F.PRG$
'    GET #f.tmp%, VAL(user$) + 1
'    werwardas$ = Chr$(CVI(f.PWCODE$))
'    Close #f.tmp%
'    ProgrammNr$ = Chr$(0)
'    '1 VERKAUF
'    '2 INFOART
'    '3 STAMA
'    '4 VERKERF
'    '5 POSDRUCK
'    '6 FAKT
'    '7 WARENEIN
'    ProgrammNr$ = Chr$(7)
'  End If

    Lb1.SatzLock (1)
    Lb1.GetRecord (1)
    Zeiger! = Lb1.erstzeiger
    Groesse! = Lb1.erstgroesse
    If (Groesse! <= 0.1) Then
        Groesse! = 100000!
        Zeiger! = 0!
    End If
    If (Zeiger! <= 0.1) Or (Zeiger! > 100000.1) Then Zeiger! = 0
    Zeiger! = Zeiger! + 1
    If (Zeiger! > Groesse!) Then Zeiger! = 1
    Lb1.PutRecord (Zeiger! + 1)
    Lb1.erstzeiger = Zeiger!
    Lb1.erstgroesse = Groesse!
    Lb1.PutRecord (1)
    Lb1.SatzUnLock (1)
    Lb1.CloseDatei
End If

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
OptionenModus% = 0
frmWuOptionen.Show 1
If (OptionenNeu%) Then
    Call AuslesenWu
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
                        
Call NeuerDS(ActLifDat$)

If (MitAuslesen%) Then
    Call AuslesenWu
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
                        
Call AuslesenWu

Call DefErrPop
End Sub

Sub ZeigeBesorgerDaten(Optional bModus% = 0)
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
Dim row%, AbholNr%, rInd%
Dim h$

bModus% = 0

With ParentForm.flxarbeit(0)
    row% = .row
    h$ = RTrim$(.TextMatrix(row%, 21))
    AbholNr% = Val(h$)
End With

Call BesorgerInfo(AbholNr%, KorrPzn$, KorrTxt$, bModus%)
If (bModus% = 2) Then
    rInd% = SucheFlexZeile(False)
    If (rInd% > 0) Then
        If (ww.WuLm > 0) Then ww.WuLm = 0
        ww.WuLm = -(ww.WuBm + ww.WuNm)
        ww.PutRecord (rInd% + 1)
        Call AuslesenDateiSatz(rInd%, True)
    End If
End If

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

If (Index = 0) Then
    Call BackToBestellungAlle(False)
ElseIf (Index = 1) Then
    Call BackToBestellungAlle
ElseIf (Index = 2) Then
    Call UmsatzDazu
ElseIf (Index = 3) Then
    BelegModus% = 1
    frmWuBeleg.Show 1
ElseIf (Index = 4) Then
    Call WuAlleBestaetigen
Else
    Call DruckeKontrollBlaetter
End If

Call DefErrPop
End Sub

Sub WuAepKalk()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("WuAepKalk")
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
Dim txt$

If (EinzelneWu% = False) Then
    With ParentForm.flxarbeit(0)
        If (EinzelneWu% = False) Then
            If (Val(.TextMatrix(.row, 24)) > 0) Then
                txt$ = "*"
            Else
                txt$ = " "
            End If
            WuLifDat$ = "@" + .TextMatrix(.row, 16) + txt$
            Call AuslesenWu(True)
            .col = 12
        End If
    End With
End If

rInd% = SucheFlexZeile(False)
If (rInd% <= 0) Then Call DefErrPop: Exit Sub

frmDirektAepKalk.Show 1
    
Call DefErrPop
End Sub
                  
Sub TeilAbschluss()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("TeilAbschluss")
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
Dim rInd%, iLief%
Dim erg$, pzn$, txt$

txt$ = ""
erg$ = MatchCode(1, pzn$, txt$, False, False)
If (erg$ <> "") Then
    iLief% = Val(pzn$)
    lifzus.GetRecord (iLief% + 1)
    Call rabtab.TabelleEinlesen(iLief%)
    Call WuFertigAlle(False, iLief%)
End If

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
    Call ZeigeAngebote(KorrPzn$, KorrTxt$, 0, angInd%, IstTemporaer, IstManuell, AngebotY%, angBm%, angNm%, angLief%, angAep#, angZr!, iHerst$)
    If (angAngebotYOrg% <> AngebotY%) Then
        l& = WritePrivateProfileString(UserSection$, "AngebotY", Str$(AngebotY%), INI_DATEI)
    End If
    
'    Call ZeigeAngebote(KorrPzn$, KorrTxt$, 0, angInd%, 0, 0, AngebotY%, angBm%, angNm%, angLief%, angAep#)
'    If (angInd% > 100) Then
'        Call EigenesAngebotErfassen(KorrPzn$, KorrTxt$, angInd%, angBm%, angNm%, angLief%, angAep#)
'    ElseIf (angInd% >= 0) Then
'        If (angIndOrg% >= 100) Then angInd% = angInd% - 1
'        Call AngebotZuordnen(angInd%, angBm%, angNm%, angLief%, angAep#)
'    End If
    
End If

Call DefErrPop
End Sub

Sub NNAEPRechnen(row%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("NNaepRechnen")
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
Dim rInd%, art%, orgart%, NNEingabe%, EingabeArt%
Dim aep#, NNAEP#

'If (para.Land = "A") Then Call DefErrPop: Exit Sub

ParentForm.flxarbeit(0).row = row%
rInd% = SucheFlexZeile(False)
If (rInd% > 0) Then

    If (IstDirektLief%) Then
        NNAEP# = ww.WuAEP
        If (ww.WuLm <> 0) Then NNAEP# = (NNAEP# * (Abs(ww.WuLm) - Abs(ww.WuNm))) / Abs(ww.WuLm)
        NNAEP# = FNX(NNAEP#)
        ww.WuNNAep = NNAEP#
        ww.WuNNart = 4
        ww.WuNNaepOk = 1
        ww.PutRecord (rInd% + 1)
    Else
        NNEingabe% = 0
        If (ww.pzn <> "9999999") Then
            FabsErrf% = ass.IndexSearch(0, ww.pzn, FabsRecno&)
            If (FabsErrf% = 0) Or (ww.WuStat = "J") Then
              
        '        If InStr(Benutz$, "&") > 0 And liefvor% > 9 Then
        '            Call Ueberweisungsauftrag(liefvor%, gh%, f.lif%)
        '        End If
              
                art% = ww.WuNNart
                NNAEP# = ww.WuNNAep
                If (art% <> 1) And (art% <> 2) Then art% = Int(art% / 10)
                If (art% <> 1) And (art% <> 2) Then art% = 0
                If (art% <> 1) Then
                    orgart% = art%
                Else
                    orgart% = 0
                End If
                
                If (art% <> 1) Then
                  aep# = ww.WuAEP
                  'wenn bei Angeboten nicht geliefert wird, was bestellt wurde oder
                  'bei Bestellung mit NM und Teillieferung -> NNEK unklar --> Eingabe
                  If (orgart% = 2 Or Abs(ww.WuNm) > 0) And Abs(ww.WuLm) <> Abs(ww.WuBm) + Abs(ww.WuNm) Then NNEingabe% = True
                  If (rabtab.TabOk) And (Abs(ww.WuRm) > 0) And Not (NNEingabe%) Then
                    '3.11 Rm immer LM-Naturalrabatt, bei Teillieferung wird dieser Teil
                    'ohnehin umgangen
                    Call rabtab.CalcNNEK(Abs(ww.WuLm), Abs(ww.WuLm) - Abs(ww.WuNm), NNAEP#, art%, ww.txt, aep#, ww.pzn, WuInDm%)
                  End If
                  If Not (rabtab.TabOk) Or NNEingabe% Then
                    If Abs(ww.WuRm) > 0 Then             '3.11
                      EkEingabeArt% = -1
                      If (NNEingabe%) Then
                        If (orgart% = 2) Then
                          EkEingabeArt% = 2
                        Else
                          EkEingabeArt% = -2
                        End If
                      End If
                      If (NNAEP# = 0#) Then NNAEP# = aep#
                      Call rabtab.EkEingabe(ww.txt, NNAEP#, 1, EkEingabeArt%)
                    End If
                    art% = 5
                  End If
                End If
                If (ww.WuAEP = 0) Then
                    ww.WuAEP = aep#
                '      Wert# = Wert# + (AEP# * Abs(CVI(b.rm$)))
                End If
                If (NNAEP# < 0.01) Then NNAEP# = ww.WuAEP
        '        LSet b.ghlief$ = Chr$(gh%)
                ww.WuNNAep = NNAEP#
                ww.WuNNart = orgart% * 10 + art%
                ww.WuNNaepOk = 1
                ww.PutRecord (rInd% + 1)
            End If
        End If
    End If
End If

'GoSub Preiskorrektur

Call DefErrPop
End Sub
      
Sub NNAEPSpeichern()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("NNaepSpeichern")
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
Dim i%, menge%, nnmax&, dWert#

'If (para.Land = "A") Then Call DefErrPop: Exit Sub

FabsErrf% = nnek.IndexSearch(0, ww.pzn, FabsRecno&)
If (FabsErrf% = 0) Then
    nnek.GetRecord (FabsRecno& + 1)
Else
    FabsErrf% = nnek.IndexSearch(0, Space$(7), FabsRecno&)
    If (FabsErrf% = 0) Then
        FabsErrf% = nnek.IndexReplace(0, FabsRecno&, Space$(7), ww.pzn, FabsRecno&)
    Else
        FabsErrf% = nnek.IndexInsert(0, ww.pzn, FabsRecno&)
    End If
    
    If (FabsErrf% <> 0) Then Call DefErrPop: Exit Sub
    
    nnek.GetRecord (1)
    nnmax& = nnek.erstmax
    If (FabsRecno& > nnmax&) Then
        nnmax& = FabsRecno&
        nnek.erstmax = nnmax&
        nnek.PutRecord (1)
    End If
    
    For i% = 0 To 4
        nnek.BelegDat(i%) = 0
        nnek.NNAEP(i%) = 0#
        nnek.art(i%) = 0
        nnek.Beleg(i%) = Space$(Len(nnek.Beleg(0)))
        nnek.Lm(i%) = 0
        nnek.rm(i%) = 0
        nnek.LiefNr(i%) = 0
        nnek.GhNr(i%) = 0 '?
    Next i%
End If

For i% = 4 To 1 Step -1
    nnek.BelegDat(i%) = nnek.BelegDat(i% - 1)
    nnek.NNAEP(i%) = nnek.NNAEP(i% - 1)
    nnek.art(i%) = nnek.art(i% - 1)
    nnek.Beleg(i%) = nnek.Beleg(i% - 1)
    nnek.Lm(i%) = nnek.Lm(i% - 1)
    nnek.rm(i%) = nnek.rm(i% - 1)
    nnek.LiefNr(i%) = nnek.LiefNr(i% - 1)
    nnek.GhNr(i%) = nnek.GhNr(i% - 1)
Next i%
    
dWert# = ww.WuNNAep
If (WuInDm% > 0) Then dWert# = FNX(dWert# / EURO_FAKTOR)
    

nnek.pzn = ww.pzn
nnek.BelegDat(0) = ww.WuBelegDatum
nnek.NNAEP(0) = dWert#  ' ww.WuNNAep
nnek.art(0) = ww.WuNNart
nnek.Beleg(0) = ww.WuBeleg
nnek.Lm(0) = Abs(ww.WuLm)
nnek.rm(0) = Abs(ww.WuRm)
nnek.LiefNr(0) = ww.Lief
nnek.GhNr(0) = 0 '?
    
nnek.PutRecord (FabsRecno& + 1)

If (FremdPznOk%) And (IstDirektLief%) And (IstAltLast% = 0) Then
    menge% = CInt(dWert# * 100)
    menge% = ZeigeDirektAufteilung%(-menge%, 2)
End If

Call DefErrPop
End Sub
                  
Sub WumsatzSpeichern()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("WumsatzSpeichern")
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
Dim Handle%
Dim h$

Handle% = FreeFile
Open "wumsatz.dat" For Append As #Handle%
h$ = Mid$(ActLifDat$, 2, 6) + " " + Right$(Space$(3) + Str$(Asc(Left$(ActLifDat$, 1))), 3)
h$ = h$ + " " + Right$(Space$(9) + Str$(Int(Warenwert# * 100# + 0.5)), 9)
h$ = Left$(h$ + Space$(32), 30)
Print #Handle%, h$
Close #Handle%

Call DefErrPop
End Sub
                  
Sub KontrollListe(typ%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("KontrollListe")
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
Dim i%, j%, ret%, sp%(8), Y%, anz%, ind%, bis%
Dim h$, KopfZeile$, header$, tx$, PosLag$, pzn$, Lac$, AvpAlt$, VkLager$

frmAction.lstSortierung.Clear

If (typ% = 0) Then      'Lagerkontroll-Liste
    KopfZeile$ = "Lagerkontroll-Liste"
    header$ = frmAction.lblArbeit(0).Caption
    
    ret% = iMsgBox("Lagerkontroll-Liste drucken ?", vbYesNo Or vbDefaultButton1)
    If (ret% <> vbYes) Then Call DefErrPop: Exit Sub
    
    Call StartAnimation(frmAction, "Ausdruck wird erstellt ...")
    
    
    For i% = 1 To UBound(LagerListePzn$)
        FabsErrf% = ast.IndexSearch(0, LagerListePzn$(i%), FabsRecno&)
        If (FabsErrf% = 0) Then
            ast.GetRecord (FabsRecno& + 1)
            
            PosLag$ = " "
            VkLager$ = " "
            FabsErrf% = ass.IndexSearch(0, LagerListePzn$(i%), FabsRecno&)
            If (FabsErrf% = 0) Then
                ass.GetRecord (FabsRecno& + 1)
                PosLag$ = Format(ass.PosLag, "0")
                If (ass.tplatz > 0) Then
                    VkLager$ = "*"
                End If
            End If
            
            h$ = ast.kurz + vbTab + ast.meng + vbTab + ast.meh
            h$ = h$ + vbTab + Format(ast.aep, "0.00") + vbTab + Format(ast.AVP, "0.00")
            h$ = h$ + vbTab + PosLag$
            h$ = h$ + vbTab + ast.Lac
            h$ = h$ + vbTab + VkLager$
            h$ = h$ + vbTab + ast.pzn + vbTab
            frmAction.lstSortierung.AddItem h$
        End If
    Next i%
Else
    KopfZeile$ = "Preis-Aktualisierungen"
    header$ = frmAction.lblArbeit(0).Caption
    
    Call StartAnimation(frmAction, "Ausdruck wird erstellt ...")
    
    
    With frmWuPreisKalk.flxarbeit(0)
        For i% = 1 To (.Rows - 1)
            If (.TextMatrix(i%, 1) = Chr$(214)) Then
                pzn$ = .TextMatrix(i%, 0)
                AvpAlt$ = " "
                PosLag$ = " "
                Lac$ = " "
                
                FabsErrf% = ast.IndexSearch(0, pzn$, FabsRecno&)
                If (FabsErrf% = 0) Then
                    ast.GetRecord (FabsRecno& + 1)
                    Lac$ = ast.Lac
                    AvpAlt$ = Format(ast.AVP, "0.00")
                    
                    FabsErrf% = ass.IndexSearch(0, pzn$, FabsRecno&)
                    If (FabsErrf% = 0) Then
                        ass.GetRecord (FabsRecno& + 1)
                        PosLag$ = Format(ass.PosLag, "0")
                    End If
                End If
                    
                h$ = .TextMatrix(i%, 2) + vbTab + .TextMatrix(i%, 3) + vbTab + .TextMatrix(i%, 4)
                h$ = h$ + vbTab + AvpAlt$ + vbTab + Format(CDbl(.TextMatrix(i%, 6)), "0.00")
                h$ = h$ + vbTab + PosLag$
                h$ = h$ + vbTab + Lac$
                h$ = h$ + vbTab + pzn$ + vbTab
                frmAction.lstSortierung.AddItem h$
            End If
        Next i%
    End With
End If


Printer.ScaleMode = vbTwips
Printer.Font.Name = "Arial"

DruckSeite% = 0
    
Call DruckKopf(header$, "", KopfZeile$)
Printer.Font.Size = 11
        
sp%(0) = Printer.TextWidth(String$(9, "9"))
sp%(1) = sp%(0) + Printer.TextWidth(String$(27, "X"))
sp%(2) = sp%(1) + Printer.TextWidth(String$(5, "X"))
sp%(3) = sp%(2) + Printer.TextWidth(String$(4, "X"))
sp%(4) = sp%(3) + Printer.TextWidth(String$(9, "9"))
sp%(5) = sp%(4) + Printer.TextWidth(String$(9, "9"))
sp%(6) = sp%(5) + Printer.TextWidth(String$(6, "9"))
sp%(7) = sp%(6) + Printer.TextWidth(String$(5, "X"))
If (typ% = 0) Then
    sp%(8) = sp%(7) + Printer.TextWidth(String$(4, "X"))
Else
    sp%(8) = sp%(7)
End If
        
Printer.CurrentX = 0
Printer.Print "P Z N";
Printer.CurrentX = sp%(0)
Printer.Print "A R T I K E L";
If (typ% = 0) Then
    tx$ = "A E P"
Else
    tx$ = "AVP alt"
End If
Printer.CurrentX = sp%(4) - Printer.TextWidth(tx$)
Printer.Print tx$;
tx$ = "A V P"
Printer.CurrentX = sp%(5) - Printer.TextWidth(tx$)
Printer.Print tx$;
tx$ = "POS"
Printer.CurrentX = sp%(6) - Printer.TextWidth(tx$)
Printer.Print tx$;
tx$ = "ORT"
Printer.CurrentX = sp%(7) - Printer.TextWidth(tx$)
Printer.Print tx$;
If (typ% = 0) Then
    tx$ = "VK"
    Printer.CurrentX = sp%(8) - Printer.TextWidth(tx$)
    Printer.Print tx$;
End If

Printer.Print " "
Y% = Printer.CurrentY
Printer.Line (0, Y%)-(sp%(8), Y%)

With frmAction.lstSortierung
    anz% = .ListCount
    For i% = 1 To anz%
        .ListIndex = i% - 1
        h$ = .text
        
        If (typ% = 0) Then
            bis% = 8
        Else
            bis% = 7
        End If
        For j% = 0 To bis%
            ind% = InStr(h$, vbTab)
            tx$ = Left$(h$, ind% - 1)
            h$ = Mid$(h$, ind% + 1)
            
            If (j% = bis%) Then
                Printer.CurrentX = 0
            ElseIf (j% = 0) Or (j% = 2) Then
                Printer.CurrentX = sp%(j%)
            Else
                Printer.CurrentX = sp%(j% + 1) - Printer.TextWidth(tx$ + "x")
            End If
            
            Printer.Print tx$;
        Next j%
        
        Printer.Print " "
        If (i% Mod 3 = 0) Then
            Printer.CurrentY = Printer.CurrentY + 60
        End If
        
        If (Printer.CurrentY > Printer.ScaleHeight - 1000) Then
            Call DruckFuss
            Call DruckKopf(header$, "", KopfZeile$)
            Printer.Font.Size = 11
        End If
    Next i%
End With
    
Call DruckFuss(False)

Printer.EndDoc

Call StopAnimation(frmAction)

Call DefErrPop
End Sub
                  
Sub RetourenAusdruck()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("RetourenAusdruck")
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
Dim i%, j%, k%, ret%, rInd%, Y%, sp%(5), anz%, ind%, AktLief%, Erst%, AnzRetourArtikel%, Handle%
Dim ZentrierX%
Dim RetourWert#
Dim tx$, h$, KopfZeile$, header$

With frmAction.lstSortierung
    .Clear
'    ww.SatzLock (1)
    For i% = 1 To (ParentForm.flxarbeit(0).Rows - 1)
        If (ParentForm.flxarbeit(0).TextMatrix(i%, 19) = "*") Then
            ParentForm.flxarbeit(0).row = i%
            rInd% = SucheFlexZeile(True)
            If (rInd% > 0) Then
                If (ww.LmStatus) And (ww.WuNeuLm < 0) Then
                    .AddItem Format(ww.Lief, "000") + Format(i%, "0000")
                End If
            End If
        End If
    Next i%
'    ww.SatzUnLock (1)
    
    ret% = vbNo
    If (.ListCount > 0) Then
        ret% = iMsgBox("Retouren-Liste drucken ?", vbYesNo Or vbDefaultButton1)
    End If
    If (ret% <> vbYes) Then Call DefErrPop: Exit Sub

    
    Call StartAnimation(frmAction, "Ausdruck wird erstellt ...")
    KopfZeile$ = "Retouren-Liste"
    header$ = frmAction.lblArbeit(0).Caption
    
    
    Printer.ScaleMode = vbTwips
    Printer.Font.Name = "Arial"
    Printer.Font.Size = 18  'nötig wegen Canon-BJ; sonst ab 2.Ausdruck falsch
    Printer.Font.Size = 11
    
    
    On Error Resume Next
    
    RetourFussTxt$ = ""
    RetourFussHoehe% = 0
    Handle% = FreeFile
    Open "\user\text\retour0.txt" For Input As #Handle%
    If (Err = 0) Then
        Do While Not (EOF(Handle%))
            Line Input #Handle%, h$
            If (RetourFussTxt$ <> "") Then RetourFussTxt$ = RetourFussTxt$ + vbCrLf
            RetourFussTxt$ = RetourFussTxt$ + h$
        Loop
        Close #Handle%
        RetourFussTxt$ = RetourFussTxt$ + vbCrLf
        Call OemToChar(RetourFussTxt$, RetourFussTxt$)
        RetourFussHoehe% = Printer.TextHeight(RetourFussTxt$)
    End If
    RetourFussHoehe% = RetourFussHoehe% + Printer.TextHeight("Seite") + 567
        
    RetourKopfTxt$ = ""
    Handle% = FreeFile
    Open "\user\text\retour1.txt" For Input As #Handle%
    If (Err = 0) Then
        Do While Not (EOF(Handle%))
            Line Input #Handle%, h$
            If (RetourKopfTxt$ <> "") Then RetourKopfTxt$ = RetourKopfTxt$ + vbCrLf
            RetourKopfTxt$ = RetourKopfTxt$ + h$
        Loop
        Close #Handle%
    Else
        RetourKopfTxt$ = para.Fistam(0) + vbCrLf + para.Fistam(1)
    End If
    Call OemToChar(RetourKopfTxt$, RetourKopfTxt$)
    
    On Error GoTo DefErr
    
    
   
    sp%(0) = Printer.TextWidth(String$(5, "9"))
    sp%(1) = sp%(0) + Printer.TextWidth(String$(8, "9"))
    sp%(2) = sp%(1) + Printer.TextWidth(String$(12, "X"))
    sp%(3) = sp%(2) + Printer.TextWidth(String$(9, "9"))
    sp%(4) = sp%(3) + Printer.TextWidth(String$(28, "X"))
    sp%(5) = sp%(4) + Printer.TextWidth(String$(10, "9"))
            
    ZentrierX% = (Printer.ScaleWidth - (sp%(5) - (sp%(0) - Printer.TextWidth("Anz." + "x")))) / 2
    For i% = 0 To 5
        sp%(i%) = sp%(i%) + ZentrierX%
    Next i%

'    ww.SatzLock (1)
    
    For k% = 1 To AnzRetourenDruck%
        AktLief% = -1
        Erst% = True
    
        Printer.ScaleMode = vbTwips
        Printer.Font.Name = "Arial"
        Printer.Font.Size = 18  'nötig wegen Canon-BJ; sonst ab 2.Ausdruck falsch
        Printer.Font.Size = 11
    
        anz% = .ListCount
        For i% = 1 To anz%
            .ListIndex = i% - 1
            h$ = .text
            
            ind% = Val(Right$(h$, 4))
            ParentForm.flxarbeit(0).row = ind%
            rInd% = SucheFlexZeile(True)
            If (rInd% > 0) Then
                If (ww.Lief <> AktLief%) Then
                    AktLief% = ww.Lief
                    If (Erst% = False) Then
                        Y% = Printer.CurrentY
                        Printer.Line (0, Y%)-(sp%(5), Y%)
                        Printer.CurrentX = sp%(3) + Printer.TextWidth("x")
                        Printer.Print Format(AnzRetourArtikel%, "0") + " Position(en)";
                        tx$ = Format(RetourWert#, "0.00")     'MarkWert#
                        Printer.CurrentX = sp%(5) - Printer.TextWidth(tx$ + "x")
                        Printer.Print tx$;
                        Printer.Print " "
                        Call RetourenFuss
                    End If
                    Erst% = False
                    DruckSeite% = 0
                    AnzRetourArtikel% = 0
                    RetourWert# = 0#
                    Call RetourenKopf(AktLief%)
                
                    tx$ = "Anz."
                    Printer.CurrentX = sp%(0) - Printer.TextWidth(tx$ + "x")
                    Printer.Print tx$;
                    tx$ = "Datum"
                    Printer.CurrentX = sp%(1) - Printer.TextWidth(tx$ + "x")
                    Printer.Print tx$;
                    tx$ = "Beleg"
                    Printer.CurrentX = sp%(1) + Printer.TextWidth("x")
                    Printer.Print tx$;
                    tx$ = "P Z N"
                    Printer.CurrentX = sp%(3) - Printer.TextWidth(tx$ + "x")
                    Printer.Print tx$;
                    tx$ = "A R T I K E L"
                    Printer.CurrentX = sp%(3) + Printer.TextWidth("x")
                    Printer.Print tx$;
                    tx$ = "A E P"
                    Printer.CurrentX = sp%(5) - Printer.TextWidth(tx$ + "x")
                    Printer.Print tx$;
                    Printer.Print " "
                    
                    tx$ = "Verfall"
                    Printer.CurrentX = sp%(1) - Printer.TextWidth(tx$ + "x")
                    Printer.Print tx$;
                    tx$ = "Begründung"
                    Printer.CurrentX = sp%(1) + Printer.TextWidth("x")
                    Printer.Print tx$;
                    Printer.Print " "
                    
                    Y% = Printer.CurrentY
                    Printer.Line (ZentrierX%, Y%)-(sp%(5), Y%)
                End If
                
                AnzRetourArtikel% = AnzRetourArtikel% + 1
                RetourWert# = RetourWert# + (Abs(ww.WuNeuLm) * ww.WuAEP)
                
                h$ = Str$(Abs(ww.WuNeuLm)) + vbTab + sdate(ww.WuBelegDatum) + vbTab + ww.WuBeleg
                h$ = h$ + vbTab + ww.pzn + vbTab + ww.txt + vbTab + Format(ww.WuAEP, "0.00") + vbTab
            
                For j% = 0 To 5
                    ind% = InStr(h$, vbTab)
                    tx$ = Left$(h$, ind% - 1)
                    h$ = Mid$(h$, ind% + 1)
                    
                    If (j% = 2) Or (j% = 4) Then
                        Printer.CurrentX = sp%(j% - 1) + Printer.TextWidth("x")
                    Else
                        Printer.CurrentX = sp%(j%) - Printer.TextWidth(tx$ + "x")
                    End If
                    
                    Printer.Print tx$;
                Next j%
                Printer.Print " "
                
                h$ = vbTab + Mid$(ww.WuAblDatum, 3) + vbTab + ww.WuText + vbTab
                For j% = 0 To 2
                    ind% = InStr(h$, vbTab)
                    tx$ = Left$(h$, ind% - 1)
                    h$ = Mid$(h$, ind% + 1)
                    
                    If (j% = 2) Or (j% = 4) Then
                        Printer.CurrentX = sp%(j% - 1) + Printer.TextWidth("x")
                    Else
                        Printer.CurrentX = sp%(j%) - Printer.TextWidth(tx$ + "x")
                    End If
                    
                    Printer.Print tx$;
                Next j%
                Printer.Print " "
            End If
            
            If (Printer.CurrentY > Printer.ScaleHeight - RetourFussHoehe%) Then
                Call RetourenFuss
                Call RetourenKopf(AktLief%)
            End If
        Next i%
        
    '    ww.SatzUnLock (1)
    
        Y% = Printer.CurrentY
        Printer.Line (ZentrierX%, Y%)-(sp%(5), Y%)
        Printer.CurrentX = sp%(3) + Printer.TextWidth("x")
        Printer.Print Format(AnzRetourArtikel%, "0") + " Position(en)";
        tx$ = Format(RetourWert#, "0.00")     'MarkWert#
        Printer.CurrentX = sp%(5) - Printer.TextWidth(tx$ + "x")
        Printer.Print tx$;
        Printer.Print " "
        
        Call RetourenFuss(False)
        Printer.EndDoc
    Next k%
    
    Call StopAnimation(frmAction)
End With

Call DefErrPop
End Sub

Sub RetourenKopf(Lief%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("RetourenKopf")
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
Dim l&, he&, i%, pos%, X%, Y%
Dim h$, heute$, SeitenNr$, header$

header$ = "Retourwaren"

heute$ = Format(Day(Date), "00") + "-"
heute$ = heute$ + Format(Month(Date), "00") + "-"
heute$ = heute$ + Format(Year(Date), "0000")
'heute$ = heute$ + " " + Left$(Time$, 5)

With Printer
    
    .CurrentX = 0: .CurrentY = 0
    .Font.Size = 11
    Printer.Print RetourKopfTxt$
    Y% = .CurrentY
    Printer.Line (0, Y%)-(.ScaleWidth, Y%)
    .CurrentX = 0
    .CurrentY = .CurrentY + 300
    .Font.Size = 14
    
    
    
'    .CurrentX = 0: .CurrentY = 0
'    .Font.Size = 14
    
    DruckSeite% = DruckSeite% + 1
'    SeitenNr$ = "-" + Str$(DruckSeite%) + " -"
'    l& = .TextWidth(SeitenNr$)
'    .CurrentX = (.ScaleWidth - l&) / 2: .CurrentY = 0
'    Printer.Print SeitenNr$
    
'    l& = .TextWidth(heute$)
'    .CurrentX = .ScaleWidth - l& - 10: .CurrentY = 0
'    Printer.Print heute$
    
'    .CurrentX = 0
'    l& = .TextHeight("A") + 300
'    .CurrentY = l&
    
    
    lif.GetRecord (Lief% + 1)
    For i% = 0 To 3
        h$ = RTrim$(lif.Name(i%))
        Printer.Print h$
        If (i% = 1) Or (i% = 3) Then Printer.Print
    Next i%
    h$ = RTrim$(lif.kurz)
    Printer.Print h$ + "  (" + Mid$(Str$(Lief%), 2) + ")"
    .Font.Size = 11
    Printer.Print
    
    .Font.Size = 18
    l& = .TextWidth(header$)
    he& = .TextHeight("A")
        
    X% = (.ScaleWidth - (l& + 800)) / 2
    Y% = .CurrentY
    .DrawWidth = 2
'    RoundRect .hdc, X% / .TwipsPerPixelX, Y% / .TwipsPerPixelY, (X% + l& + 800) / .TwipsPerPixelX, (Y% + he& + 800) / .TwipsPerPixelY, 200, 200
    .CurrentX = (.ScaleWidth - l&) / 2
    .CurrentY = Y% + 400
    Printer.Print header$;
    
    l& = .TextWidth(heute$)
    .CurrentX = .ScaleWidth - l& - 10
    Printer.Print heute$
    
    Printer.Print
    .Font.Size = 11
    Printer.Print
    Printer.Print
End With

Call DefErrPop
End Sub

Sub RetourenFuss(Optional NewPage% = True)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("RetourenFuss")
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
Dim Y%
Dim h$

With Printer
    .Font.Bold = False
    .Font.Size = 11
    
    h$ = "Seite" + Str$(DruckSeite%)
            
    .CurrentX = (.ScaleWidth - Printer.TextWidth(h$)) / 2
    .CurrentY = .ScaleHeight - Printer.TextHeight(h$)
    Printer.Print h$
    
    If (RetourFussTxt$ <> "") Then
        .CurrentX = 0
        .CurrentY = .CurrentY - (RetourFussHoehe% - 567)
        Printer.Print RetourFussTxt$;
    End If
    
    If (NewPage% = True) Then .NewPage
End With

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

frmAction.flxarbeit(0).col = 10
SuperEdit% = True
Call EditSatz

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
Call AuslesenWu
Call DefErrPop
End Sub
                  
Sub EinlesenAlteWue()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("EinlesenAlteWue")
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
Dim i%, j%, bm%, nm%, AltWueMax%, ret%, wwMax%
Dim text$, iPzn$

wu.GetRecord (1)
AltWueMax% = wu.erstmax

ret% = vbNo
If (AltWueMax% > 0) Then
    ret% = iMsgBox("Sollen die Daten der DOS-Warenübernahme übernommen (und dort gelöscht) werden ?", vbYesNo Or vbDefaultButton2)
End If

If (ret% = vbYes) Then
    wu.SatzLock (1)
    wu.GetRecord (1)
    AltWueMax% = wu.erstmax

    ww.SatzLock (1)
    ww.GetRecord (1)
    wwMax% = ww.erstmax
    
    For i% = 1 To AltWueMax%
        wu.GetRecord (i% + 1)

        text$ = wu.at
        iPzn$ = Left$(text$, 7)
        
        If (Asc(Left$(iPzn$, 1)) <= 127) Then
            ww.status = 2
        
            ww.pzn = Left$(text$, 7)
        
            text$ = Mid$(text$, 9)
            ww.txt = Left$(text$, 33) + Mid$(text$, 35)
            
            ww.Lief = Asc(wu.li)
                            
            ww.WuBestDatum = wu.bd
            ww.WuBestZeit = wu.zeit
            
            ww.WuStat = wu.stat
            bm% = wu.bm
            nm% = wu.nm
            If (wu.stat = "N") Then
                bm% = -bm%
                nm% = -nm%
            End If
            
            ww.bm = bm%
            ww.nm = nm%
            ww.WuBm = Abs(bm%)
            ww.WuNm = Abs(nm%)
            ww.WuRm = Abs(bm%)
            ww.WuLm = Abs(bm%) + Abs(nm%)
                                        
            ww.aep = wu.aep
            ww.WuAEP = ww.aep
                            
            ww.AVP = wu.AVP
            ww.WuAVP = ww.AVP
            
            ww.WuAm = wu.am
            ww.abl = wu.ac
            ww.asatz = wu.asatz
            ww.ssatz = wu.ssatz
            ww.km = wu.km
            
            ww.WuLa = wu.la
            ww.wg = wu.wg
                            
            If (wu.besorger = "B") Then
                ww.auto = "v"
            Else
                ww.auto = Chr$(0)
            End If
                
            ww.AbholNr = wu.AbholNr
            ww.WuAblDatum = wu.abl
                            
            ww.nnart = wu.nnart
            ww.WuNNart = ww.nnart
                            
            ww.NNAEP = wu.NNAEP
            ww.WuNNAep = ww.NNAEP
                            
                           
                            
            ww.WuBelegDatum = 0
            ww.WuBeleg = Space$(10)
            ww.WuRetMenge = 0
            ww.WuNNaepOk = 0
            ww.WuText = Space$(Len(ww.WuText))
            ww.WuNeuLm = 0
            ww.WuNeuRm = 0
            ww.WuNeuZiel = 0
            
            ww.IstAltLast = 0
            ww.WuStatus = 0
            ww.LmStatus = 0
            ww.RmStatus = 0
            ww.LmAnzGebucht = 0
            ww.RmAnzGebucht = 0
            
            ww.aktivlief = 0
            ww.aktivind = 0
                            
            ''''''''
            ww.best = " "
            ww.absage = 0
            ww.angebot = 0
            
            ww.alt = Chr$(0)
            
            ww.besorger = " "
            
    '        ww.beklaufnr = Val(Format(Day(Date), "00") + Right$(Format(Now, "HHMMSS"), 4) + Right$(ww.pzn, 3))
            ww.BekLaufNr = CalcLaufNr&(ww.pzn)
            AnzeigeLaufNr& = ww.BekLaufNr
                
            
            ww.zugeordnet = Chr$(0)
            ww.zukontrollieren = Chr$(0)
            ww.fixiert = Chr$(0)
            ww.DirektTyp = 0
            
            For j% = 0 To 5
                ww.actkontrolle(j%) = 111
            Next j%
            ww.actzuordnung = 111
            ww.PosLag = 0
            
            ww.herst = Space$(Len(ww.herst))
            
            ww.loesch = 0
            '''''''
        
            wwMax% = wwMax% + 1
            ww.erstmax = wwMax%
            ww.PutRecord (1)
            ww.PutRecord (wwMax% + 1)
        End If
    Next i%
    
    wu.erstmax = 0
    wu.PutRecord (1)
    wu.SatzUnLock (1)

    ww.SatzUnLock (1)
End If
                
Call DefErrPop
End Sub

Sub WuAlleBestaetigen()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("WuAlleBestaetigen")
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
Dim i%, rInd%, AltRow%, besMenge%, besStatus%, TuEs%, ret%, FrageRet%
Dim pzn$, SQLStr$

With ParentForm.flxarbeit(0)
    .redraw = False
    AltRow% = .row
    
    ww.SatzLock (1)
    For i% = 1 To (.Rows - 1)
        If (.TextMatrix(i%, 19) <> "*") Then
            .row = i%
            rInd% = SucheFlexZeile(True)
            If (rInd% > 0) Then
                
                If (IstAltLast% = False) Then
                    ww.WuBeleg = ActBeleg$
                    ww.WuBelegDatum = ActBelegDatum%
                End If
                
                ww.LmStatus = 1
                ww.RmStatus = 0
                ww.WuStatus = 0
                
                ww.PutRecord (rInd% + 1)
            End If
        End If
    Next i%
    ww.SatzUnLock (1)
    .row = AltRow%
    .redraw = True
    
    If (IstDirektLief%) Then
        FrageRet% = -1
        For i% = 1 To (.Rows - 1)
            pzn$ = .TextMatrix(i%, 0)
            FabsErrf% = ass.IndexSearch(0, pzn$, FabsRecno&)
            If (FabsErrf% <> 0) Then
                SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + pzn$
                Set TaxeRec = TaxeDB.OpenRecordset(SQLStr$)
                If (TaxeRec.EOF = False) Then
                    Call Taxe2ast(pzn$)
                
                    If (FrageRet% = -1) Then FrageRet% = iMsgBox("Alle Artikel dieser Lieferung auf Lager nehmen ?", vbYesNo Or vbDefaultButton2)
                    If (FrageRet% = vbYes) Then
                        ret% = InStammAufnehmen%(pzn$)
                        If (ret% = False) Then Call DefErrPop: Exit Sub
                    
                        ret% = InStatistikAufnehmen%(pzn$)
                        If (ret% = False) Then Call DefErrPop: Exit Sub
                    End If
                End If
            End If
        Next i%
    End If
    
    Call AuslesenWu
End With

Call DefErrPop
End Sub

Sub WEProt(pModus%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("WEProt")
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
Dim l&
Dim h$, Pfad$, text$, LastDatum$, heute$, SQLStr$, abl$
Dim HerstellerRec As Recordset
Static ProtId%
Static HerstellerDB As Database


If (pModus% = 0) Then
    ProtId% = 0
    
    h$ = "0"
    l& = GetPrivateProfileString("WEPROT", "Aktiv", "0", h$, 2, "\user\dp.ini")
    h$ = Left$(h$, l&)
    If (Val(h$) <> 0) Then
        h$ = Space$(100)
        l& = GetPrivateProfileString("WEPROT", "Pfad", " ", h$, 101, "\user\dp.ini")
        Pfad$ = Trim(Left$(h$, l&))
        If (Pfad$ = "") Then Pfad$ = CurDir$
        If (Right$(Pfad$, 1) <> "\") Then Pfad$ = Pfad$ + "\"
        
        h$ = Space$(100)
        l& = GetPrivateProfileString("WEPROT", "LastDat", " ", h$, 101, "\user\dp.ini")
        LastDatum$ = Trim(Left$(h$, l&))
        
        h$ = Format(Now, "DDMMYY")
        heute$ = DateLong(h$)
        
        If (LastDatum$ <> heute$) Then
            l& = WritePrivateProfileString("WEPROT", "LastDat", heute$, "\user\dp.ini")
        End If
    
        erg% = 0
        h$ = para.TaxeLw + ":\taxe\hersteller.mdb"
        On Error GoTo ErrorHandler
        Set HerstellerDB = OpenDatabase(h$, False, True)
        If (erg% > 0) Then
            erg% = 0
            ProtId% = 0
            Call DefErrPop: Exit Sub
        End If
        On Error GoTo DefErr

        h$ = Pfad$ + "WEPROT.ASC"
        ProtId% = FreeFile
        On Error GoTo ErrorHandler
        If (LastDatum$ <> heute$) Then
            Open h$ For Output As #ProtId%
        Else
            Open h$ For Append As #ProtId%
        End If
        If (erg% > 0) Then
            erg% = 0
            ProtId% = 0
        End If
        On Error GoTo DefErr
    End If

ElseIf (pModus% = 1) Then
    
    If (ProtId% > 0) Then
        SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + ww.pzn
        Set TaxeRec = TaxeDB.OpenRecordset(SQLStr$)
        If (TaxeRec.EOF = False) Then
            text$ = UCase$(Left$(TaxeRec!Name + Space$(28), 28) + Space$(6))
        
            SQLStr$ = "SELECT * FROM HERSTELLER WHERE NUMMER = " + Format(TaxeRec!HerstellerNr, "00000")
            Set HerstellerRec = HerstellerDB.OpenRecordset(SQLStr$)
            If (HerstellerRec.EOF = False) Then
                text$ = text$ + Left$(HerstellerRec!Name + Space$(34), 34)
            Else
                text$ = text$ + Space$(34)
            End If
            
            text$ = text$ + Left$(TaxeRec!DarForm + Space$(3), 3) + Space$(21) 'Darreichungsform
            text$ = text$ + Right$(Space$(5) + LTrim$(TaxeRec!menge), 5)  'Menge
            text$ = text$ + " " + Left$(LTrim$(TaxeRec!einheit) + Space$(2), 2) 'Meh
            
            h$ = ""
            abl$ = Trim$(ww.WuAblDatum)
            If (Len(abl$) > 2) Then h$ = DateLong(abl$)
            text$ = text$ + Left$(Left$(h$, 6) + Space$(8), 8) 'Ablauf
            
            lif.GetRecord (ww.Lief + 1)
            h$ = RTrim$(lif.kurz)
            text$ = text$ + Left$(h$ + Space$(24), 24)
            
            Print #ProtId%, text$
        End If
    End If

Else
    
    If (ProtId% > 0) Then
        HerstellerDB.Close
        Close #ProtId%
    End If

End If
    
Call DefErrPop: Exit Sub
    
ErrorHandler:
    erg% = Err
    Err = 0
    Resume Next
    Return

Call DefErrPop
End Sub

Sub CheckInterneStreichung()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("CheckInterneStreichung")
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
Dim i%, ret%, iRow%
Dim h$, pzn$

For i% = 1 To UBound(IntStreichungRow%)
    With ParentForm.flxarbeit(0)
        .redraw = True
        Call frmAction.HighlightZeile(True)
        iRow% = IntStreichungRow%(i%)
        .row = iRow%
        pzn$ = .TextMatrix(.row, 0)
        
        If (.row < .TopRow) Then
            .TopRow = .row
        Else
            If (.row >= (.TopRow + ParentBereich.ArbeitAnzZeilen - 2)) Then
                .TopRow = .row - ParentBereich.ArbeitAnzZeilen + 2
            End If
        End If

    End With

    FabsErrf% = ast.IndexSearch(0, pzn$, FabsRecno&)
    If (FabsErrf% = 0) Then
        ast.GetRecord (FabsRecno& + 1)
        h$ = Trim(ast.kurz) + "  " + Trim(ast.meng) + ast.meh
        
        ret% = iMsgBox("Interne Streichung noch aktuell ?", vbYesNo Or vbDefaultButton1, h$)
        If (ret% <> vbYes) Then
            FabsErrf% = ass.IndexSearch(0, pzn$, FabsRecno&)
            If (FabsErrf% = 0) Then
                ass.GetRecord (FabsRecno& + 1)
                ass.halt = " "
                ass.PutRecord (FabsRecno& + 1)
            End If
        End If
    End If
Next i%

Call DefErrPop
End Sub
                  
Sub RowaLsDatei(Optional Loeschen% = False)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("RowaLsDatei")
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
Dim ROWALS%
Dim i%, LsInd%, anz%, sLen%, gef%, BelegLen%
Dim SollLifDat$, h$, DatName$
Dim RowaLsSatz As RowaLsStruct

gef% = False
sLen% = Len(RowaLsSatz)
BelegLen% = Len(RowaLsSatz.Beleg)

h$ = Mid$(ActLifDat$, 8, 4)
SollLifDat$ = Left$(ActLifDat$, 7) + MKI(Val(h$))

ROWALS% = FileOpen%("\user\rowals.dat", "RW", "R", sLen%)
If (ROWALS% > 0) Then
    anz% = LOF(ROWALS%) / sLen%
    LsInd% = anz% + 1
    For i% = 1 To anz%
        Get #ROWALS%, i%, RowaLsSatz
        If (RowaLsSatz.Beleg = String$(BelegLen%, 0)) Then LsInd% = i%
        If (RowaLsSatz.LifDat = SollLifDat$) Then
            gef% = True
            LsInd% = i%
            Exit For
        End If
    Next i%
        
    If (Loeschen%) Then
        If (gef%) Then
            RowaLsSatz.Beleg = String$(BelegLen%, 0)
            RowaLsSatz.LifDat = String$(Len(RowaLsSatz.LifDat), 0)
            RowaLsSatz.Ln = " "
            RowaLsSatz.pos = 0
        End If

        DatName$ = Format(Asc(Mid$(WuLifDat$, 2, 1)), "000")
        DatName$ = DatName$ + Mid$(WuLifDat$, 3, 10) + ".dat"
        On Error Resume Next
        Kill DatName$
        On Error GoTo DefErr
    Else
        RowaLsSatz.Beleg = Left$(Trim$(ActBeleg$) + Space$(BelegLen%), BelegLen%)
        RowaLsSatz.LifDat = SollLifDat$
        RowaLsSatz.Ln = " "
        RowaLsSatz.pos = 0
    End If
    Put #ROWALS%, LsInd%, RowaLsSatz
        
    Close #ROWALS%
End If

Call DefErrPop
End Sub
                  

Sub RowaWu()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("RowaWu")
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
Dim i%, j%, WAWI_ARTIKEL%, gef%, rInd%, DatAnz%, AktTransNr%, Max%, NeuMax%, TransNr%, ok%
Dim DatLen&
Dim DatName$, h$, pzn$, Verfall$, AHbd$, AHze$, AHli$
Dim RowaSatz As String * 13
Dim RowaArtikel$()
        
ReDim RowaArtikel$(0)
        
If (RowaAktiv%) Then
    DatName$ = Format(Asc(Mid$(WuLifDat$, 2, 1)), "000")
    DatName$ = DatName$ + Mid$(WuLifDat$, 3, 10) + ".dat"
    
    h$ = ""
    WAWI_ARTIKEL% = FileOpen(DatName$, "RW", "R", Len(RowaSatz))
    DatLen& = LOF(WAWI_ARTIKEL%)
    If (DatLen& > 0) And (RowaDateiLen& <> DatLen&) Then
        DatAnz% = DatLen& / Len(RowaSatz)
        For j% = 1 To DatAnz%
            Get WAWI_ARTIKEL%, j%, RowaSatz
            If (RowaSatz <> String(13, 0)) Then
                i% = UBound(RowaArtikel$) + 1
                ReDim Preserve RowaArtikel$(i%)
                RowaArtikel$(i%) = RowaSatz
                RowaSatz = String(Len(RowaSatz), 0)
                Put WAWI_ARTIKEL%, j%, RowaSatz
            End If
        Next j%
        RowaDateiLen& = DatLen&
    End If
    Close (WAWI_ARTIKEL%)
    
    On Error Resume Next
    If (DatLen& = 0) Then Kill DatName$
    On Error GoTo DefErr
Else
    SHAREDDAT% = FileOpen%("\USER\SHARED.DAT", "RW", "R", Len(SharedDatSatz))
    If (LOF(SHAREDDAT%) = 0) Then
        Close #SHAREDDAT%
        Kill "\USER\SHARED.DAT"
    Else
        AUTOHEAD% = FileOpen%("\USER\AUTOHEAD.DAT", "RW", "R", Len(AutoHeadSatz))
    
        Call iLock(SHAREDDAT%, 1)
        Get #SHAREDDAT%, 1, SharedDatSatz
        Max% = CVI(Left$(SharedDatSatz, 2))
        NeuMax% = 0
        
        AktTransNr% = -1
    
        For i% = 1 To Max%
            Get #SHAREDDAT%, i% + 1, SharedDatSatz
            
            gef% = False
            
            'nur Datensätze von Komm-PC lesen "0" +
            'nur "01"-er Sätze für WÜ relevant
            If (Left$(SharedDatSatz, 3) = "001") Then
            
                pzn$ = Mid$(SharedDatSatz, 7, 7)
                Verfall$ = Mid$(SharedDatSatz, 15, 4) + Mid$(SharedDatSatz, 21, 2)
                TransNr% = Val(Mid$(SharedDatSatz, 4, 3)) - 800
                
                'nur Transaktionsnummern von 800 - 899 sind WÜ-Transaktionsnummern
                If (TransNr% >= 0) And (TransNr% <= 99) Then
                    'hole Bestelldaten zu Transaktionsnummer
                                        
                    If (AktTransNr% < 0) Then
                        Get #AUTOHEAD%, TransNr% + 1, AutoHeadSatz
                        
                        AHbd$ = Mid$(AutoHeadSatz, 11, 2) + Mid$(AutoHeadSatz, 8, 2) + Mid$(AutoHeadSatz, 5, 2)
                        AHze$ = Format(Val(Mid$(AutoHeadSatz, 14, 2)) * 100 + Val(Mid$(AutoHeadSatz, 17, 2)), "0000")
                        AHli$ = Chr$(Val(Mid$(AutoHeadSatz, 20, 3)))
                        h$ = AHli$ + AHbd$ + AHze$
                        If (InStr(WuLifDat$, h$) > 0) Then
                            AktTransNr% = TransNr%
                        End If
                    End If
                    
                    If (TransNr% = AktTransNr%) Then
                    
                        gef% = True
                        ok% = True
                        If (Val(pzn$) = 0) Then
                            'Endedatensatz erkannt: Transaktionsnummer freigeben
                            AutoHeadSatz = Space$(Len(AutoHeadSatz) - 2) + Chr$(13) + Chr$(10)
                            Put #AUTOHEAD%, TransNr% + 1, AutoHeadSatz
                            pzn$ = "XXXXXXX"
                        ElseIf (Mid$(SharedDatSatz, 14, 1) <> "2") Then
                            ok% = False
                        End If
                        
                        If (ok%) Then
                            j% = UBound(RowaArtikel$) + 1
                            ReDim Preserve RowaArtikel$(j%)
                            RowaArtikel$(j%) = pzn$ + Verfall$
                        End If
                          
                    End If
                Else
                    '01-er Satz aber falsche Transaktionsnummer???
                    gef% = True
                End If
            End If
            
            If (gef% = False) Then
                NeuMax% = NeuMax% + 1
                If (NeuMax% <> i%) Then Put #SHAREDDAT%, NeuMax% + 1, SharedDatSatz
            End If
        Next i%
        
        'übrige Datensätze löschen
        SharedDatSatz = String$(Len(SharedDatSatz), 0)
        For i% = NeuMax% + 1 To Max%
            Put #SHAREDDAT%, i% + 1, SharedDatSatz
        Next i%
        
        Get #SHAREDDAT%, 1, SharedDatSatz
        SharedDatSatz = MKI(NeuMax%) + Mid$(SharedDatSatz, 3)
        Put #SHAREDDAT%, 1, SharedDatSatz

        Call iUnLock(SHAREDDAT%, 1)
        Close #SHAREDDAT%
        Close #AUTOHEAD%
    End If
End If

If (UBound(RowaArtikel$) > 0) Then
    For i% = 1 To UBound(RowaArtikel$)
        pzn$ = Left$(RowaArtikel$(i%), 7)
        Verfall$ = Mid$(RowaArtikel$(i%), 8, 6)
        
        If (pzn$ = "XXXXXXX") Then
            IstRowaWu% = True
            Call MenuBearbeiten(MENU_F7)
            Exit For
        Else
            With ParentForm.flxarbeit(0)
                gef% = 0
                For j% = 1 To (.Rows - 1)
                    If (.TextMatrix(j%, 0) = pzn$) Then
                        gef% = j%
                        If (Trim$(.TextMatrix(j%, 19)) = "") Or (Val(.TextMatrix(j%, 8)) < Val(.TextMatrix(j%, 5)) + Val(.TextMatrix(j%, 6))) Then
                            Exit For
                        End If
                    End If
                Next j%
            
                If (gef% > 0) Then
                    Call ParentForm.HighlightZeile(True)
                    .row = gef%
                    .col = 8
                    
                    If (.row < .TopRow) Then
                        .TopRow = .row
                    Else
                        If (.row >= (.TopRow + ParentBereich.ArbeitAnzZeilen - 2)) Then
                            .TopRow = .row - ParentBereich.ArbeitAnzZeilen + 2
                        End If
                    End If
                    Call ParentForm.HighlightZeile
                    Call EchtKurzInfo
                    
                    rInd% = SucheFlexZeile(False)
                    If (rInd% > 0) Then
                        If (IstAltLast% = False) Then
                            ww.WuBeleg = ActBeleg$
                            ww.WuBelegDatum = ActBelegDatum%
                        End If
                        
                        If (ww.LmStatus = 0) Then ww.WuLm = 0
                        ww.WuLm = ww.WuLm + 1
                        ww.LmStatus = 1
                        ww.RmStatus = 0
                        ww.WuStatus = 0
                        
                        If (Verfall$ <> Space$(6)) Then ww.WuAblDatum = Verfall$
                        ww.PutRecord (rInd% + 1)
                        Call AuslesenDateiSatz(rInd%, True)
                        Call EchtKurzInfo
                    End If
                End If
            End With
        End If
    Next i%
End If

Call DefErrPop
End Sub

Sub UmsatzDazu()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("UmsatzDazu")
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
Dim erg$, pzn$, txt$

erg$ = MatchCode(1, pzn$, txt$, False, False)
If (erg$ <> "") Then
    WuFrageErg% = Val(pzn$)
    frmWumsatzAdd.Show 1
End If

Call DefErrPop
End Sub
                  

'SUB GemAufschlag (f.auf%, RMenge%, aep#, avp#, mw$, wg%, MwSt1, Menge%, text$)
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Call DefErrFnc("UmsatzDazu")
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
'
'pzn$ = Left$(text$, 7)
'
'If heute% < AufAbDatum% Or heute% > AufBisDatum% Then Exit Sub
'
'If RMenge% <= 0 Then Exit Sub
'If InStr("234", mw$) <> 0 Then Exit Sub
'
'Call wsave: X$ = "11410538CW": Call wmake(X$)
'XX1$ = xm$
'Call farbe("W")
'LOCATE 11, 41: Print "   Ermittlung der Verkaufspreise    ";
'LOCATE 12, 42
'If mw$ <> "1" Then
'    Call farbe("Nh")
'    Print " Mehrwertsteuer unbekannt !" + Str$(MwSt1); "% ? ";
'Else
'    Print "         für" + Str$(MwSt1) + "% Artikel           ";
'End If
'
'Call farbe("W")
'LOCATE 13, 42: Print Mid$(text$, 9);
'LOCATE 14, 42: Print USING; "AEP => #####.## "; aep#;
'LOCATE 15, 42: Print USING; "AVP => #####.## "; AVP#;
'
'x1# = aep#: X2# = AVP#
'xm$ = "[F5] Standardkalkulation      AVP=[00] -> Artikel wird nicht berücksichtigt" + BELL$
'Call FussZeile(xm$, xm%)
'LOCATE 15, 59: Call farbe("W")
'If aep# <> 0# Then Print "laut Preisliste";: GoTo 53300
'Print "Preise unbekannt";
'
'53190 x1# = aep#: X2# = AVP#
'    X$ = Str$(FNX#(x1#)): Call MyTrim(X$)
'    xp$ = "4914083": Call ein500: X# = FNX#(Val(X$))
'    x1# = X#: If X# = 0 Then X2# = 0#
'    If lt% <> 0 Then Print BELL$;: GoTo 53190
'    LOCATE 14, 49: Print USING; "#####.##"; x1#;
'53300 X$ = Str$(FNX#(X2#)): Call MyTrim(X$)
'    xp$ = "4915083": Call ein500: X# = FNX#(Val(X$))
'    If lt% = 1 Or lt% = 99 Then GoTo 53190
'    If lt% = 5 Then
'        X2# = Int(FNX#(x1# * 1.44 * (1# + MwSt1 / 100#)) * 10#) / 10#
'        LOCATE 15, 59: Print "laut Kalkulation";: GoTo 53300
'    End If
'    X2# = X#: If X2# = 0# Then x1# = 0#
'53321 If x1# <= X2# Then GoTo 53323
'    XMX$ = xm$:
'    xm$ = "Verkaufspreis kann nicht kleiner als Einkaufspreis sein !" + BELL$
'    xm% = 5: Call FussZeile(xm$, xm%): xm$ = XMX$: Call FussZeile(xm$, xm%): GoTo 53300
'53323 If X2# <= x1# * 3 Then GoTo 53325
'    XMX$ = xm$: xm$ = "Verkaufspreis zu hoch !" + BELL$: xm% = 5: Call FussZeile(xm$, xm%)
'    xm$ = XMX$: Call FussZeile(xm$, xm%): GoTo 53300
'53325 LOCATE 15, 49: Print USING; "#####.##"; X2#;
'    If lt% <> 0 Then Print BELL$;: GoTo 53300
'Call wload
'
'xm$ = XX1$: Call FussZeile(xm$, xm%)
'If x1# = 0# Or X2# = 0# Then Exit Sub
'aep# = x1#: AVP# = X2#
'
'
'
'ga.SatzLock (1)
'ga.GetFirstRecord
'aufmax% = ga.erstmax
'
'ga.pzn = pzn$
'ga.dummy = " "
'LSet auf.text$ = Mid$(text$, 9)
'ga.aep = aep#
'ga.AVP = AVP#
'ga.menge = menge%
'LSet auf.datum$ = heuteC$
'ga.kz = Chr$(0)
'aufmax% = aufmax% + 1
'ga.PutRecord (aufmax% + 1)
'
'ga.GetFirstRecord
'ga.erstmax = aufmax%
'ga.PutRecord (1)
'ga.SatzUnLock (1)
'
'Call DefErrPop
'End Sub

Function IstInNachlieferung%(Optional InWuHinein% = False, Optional sEingabe$ = "")
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("IstInNachlieferung%")
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
Dim i%, Max%, ret%

ret% = 0

If (sEingabe$ = "") Then
    sEingabe$ = EingabeStr$
End If

ww.GetRecord (1)
Max% = ww.erstmax

For i% = 1 To Max%
    ww.GetRecord (i% + 1)

    If (ww.status = 2) Then
        If (ww.pzn = sEingabe$) And (ww.IstAltLast) And (ww.WuNeuLm > 0) Then
            If (Chr$(ww.Lief) = Left$(ActLifDat$, 1)) Then
                If (InWuHinein%) Then
                    BelegModus% = 1
                    frmWuBeleg.Show 1
                    
                    ww.WuBeleg = ActBeleg$
                    ww.WuBelegDatum = ActBelegDatum%
                    
                    ww.WuBestDatum = Mid$(ActLifDat$, 2, 6)
                    ww.WuBestZeit = MKI(Mid$(ActLifDat$, 8, 4))
                
                    ww.WuAEP = ww.aep
                    ww.WuLm = ww.WuNeuLm
                    ww.WuRm = ww.WuNeuRm
                    ww.WuNeuLm = 0
                    ww.WuNeuRm = 0
                    ww.PutRecord (i% + 1)
                    
                    AnzeigeLaufNr& = ww.BekLaufNr
                    
                End If
                ret% = True
                Exit For
            End If
        End If
    End If
Next i%

IstInNachlieferung% = ret%

Call DefErrPop
End Function


Sub DruckeKontrollBlaetter()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("DruckeKontrollBlaetter")
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
Dim i%, j%, fWz%, WzMax%, menge%, iHeute%, Erst%, Y%, EW_TEXT%, AnzTexte%, EwTextX%
Dim TextMax!, TextSatz!
Dim aep#
Dim KontrNr$, h$, heuteC$, pzn$, Meng2$, AltLief$, NeuLIef$, X$, h2$
Dim WkZwi As clsWkZwi
Dim hTaxe As clsHilfsTaxe

If (iMsgBox("Kontrollblätter zur ID-Prüfung ausdrucken ?", vbYesNo Or vbDefaultButton2) <> vbYes) Then
    Call DefErrPop: Exit Sub
End If

Call StartAnimation(frmAction, "Ausdruck wird erstellt ...")
    

Set WkZwi = New clsWkZwi
WkZwi.OpenDatei
'WkZwi.SatzLock (1)
WkZwi.GetFirstRecord
WzMax% = WkZwi.erstmax

Set hTaxe = New clsHilfsTaxe
hTaxe.OpenDatei

Erst% = True
With Printer
    For i% = 1 To WzMax%
        WkZwi.GetRecord (i% + 1)
        
        If (WkZwi.DruckKz <> "D") Then
            FabsErrf% = ast.IndexSearch(0, WkZwi.pzn, FabsRecno&)
            If (FabsErrf% = 0) Then
                ast.GetRecord (FabsRecno& + 1)
            End If
            
            X$ = sdate(WkZwi.datum)
    
            NeuLIef$ = WkZwi.Lief + X$
            
            If (Erst%) Then
                Erst% = 0
                Call DruckKontrollBlaetterKopf
            End If
            
            If (AltLief$ <> NeuLIef$) Then
                If (Printer.CurrentY > Printer.ScaleHeight - 1000) Then
                    Call DruckFuss
                    Call DruckKontrollBlaetterKopf
                    Printer.Font.Size = 11
                End If
                
                AltLief$ = NeuLIef$
                
                .Font.Bold = True
                .CurrentY = .CurrentY + 450
                X$ = Left$(X$, 2) + "." + Mid$(X$, 3, 2) + "." + Mid$(X$, 5)
                h$ = "Lieferdatum:  " + X$ + "      Lieferant: " + WkZwi.Lief
                Printer.Print h$
                .Font.Bold = False

                Y% = .CurrentY + 30
                Printer.Line (0, Y%)-(.ScaleWidth, Y%)
                .CurrentX = 0
                .CurrentY = .CurrentY + 60
            End If
            
            If (Printer.CurrentY > Printer.ScaleHeight - 1000) Then
                Call DruckFuss
                Call DruckKontrollBlaetterKopf
                Printer.Font.Size = 11
            End If
            
            h$ = ast.kurz
            Printer.Print h$;
            
            Printer.Font.Size = 10
            h$ = Format(WkZwi.menge, "0") + " Stück zu "
            h$ = h$ + WkZwi.Meng2 + " " + ast.meh + "    "
            .CurrentX = .ScaleWidth * 3 / 10
            Printer.Print h$;
            
            If (WkZwi.AblaufDatum > 0) Then
                .CurrentX = .ScaleWidth * 0.49
'                Debug.Print .CurrentX
                h2$ = sdate(WkZwi.AblaufDatum)
                Printer.Print "Ablauf: " + Left$(h2$, 2); "." + Mid$(h2$, 3, 2) + "." + Mid$(h2$, 5, 2);
            End If
            Printer.Font.Size = 11
            
            .CurrentX = .ScaleWidth * 2 / 3
            h$ = "ID-Datum:                 Paraphe:"
            Printer.Print h$
            
            .CurrentY = .CurrentY + 45
            Printer.Print "Prüfung:   ";
            EwTextX% = .CurrentX
            Printer.Print
            
            FabsErrf% = hTaxe.IndexSearch(0, WkZwi.pzn, FabsRecno&)
            If (FabsErrf% = 0) Then
                Call hTaxe.GetRecord(FabsRecno& + 1)
                If (hTaxe.TextSatz > 0) Then
                    TextSatz! = hTaxe.TextSatz
                    
                    EW_TEXT% = FileOpen("EW-TEXT.DAT", "R", "R", Len(EwTextRec))
                    Get #EW_TEXT%, 1, EwTextErst
                    TextMax! = EwTextErst.Max
                    Call DxToIEEEs(TextMax!)
                    
                    For j% = 1 To 100
                        If (TextSatz! <= 0) Or (TextSatz! > TextMax!) Then
                            Exit For
                        End If
                        
                        Get #EW_TEXT%, TextSatz! + 1, EwTextRec
                        If (EwTextRec.pzn = hTaxe.pzn) Then
                            .CurrentX = EwTextX%
                            h2$ = Trim(EwTextRec.Zeile)
                            Call OemToChar(h2$, h2$)
                            Printer.Print h2$
                        Else
                            Exit For
                        End If
                        
                        TextSatz! = EwTextRec.NextSatz
                        Call DxToIEEEs(TextSatz!)
                    Next j%
                    Close #EW_TEXT%
                End If
            End If
    
            Printer.Print
            Y% = .CurrentY
            Printer.Line (0, Y%)-(.ScaleWidth, Y%)
            .CurrentX = 0
            .CurrentY = .CurrentY + 60
            
            WkZwi.DruckKz = "D"
            WkZwi.PutRecord (i% + 1)
        End If
    Next i%
End With

If (Erst% = 0) Then
    Call DruckFuss(False)
    Printer.EndDoc
End If

Call StopAnimation(frmAction)

'WkZwi.SatzUnLock (1)
WkZwi.CloseDatei

hTaxe.CloseDatei

Call DefErrPop
End Sub

Sub DruckKontrollBlaetterKopf()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("DruckKontrollBlaetterKopf")
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
'Dim l&, i%, pos%, h&, x%, y%, heute$, SeitenNr$
Dim l&
Dim h$

With Printer
    .CurrentX = 0: .CurrentY = 0
    .Font.Size = 18
    h$ = "Kontrollblatt zu ID-Prüfung"
    l& = .TextWidth(h$)
    .CurrentX = (.ScaleWidth - l&) / 2:
    .CurrentY = 90
    Printer.Print h$
    
    .Font.Size = 11
    .CurrentY = .CurrentY + 210
    
'    DruckSeite% = DruckSeite% + 1
'    SeitenNr$ = "-" + Str$(DruckSeite%) + " -"
'    l& = .TextWidth(SeitenNr$)
'    .CurrentX = (.ScaleWidth - l&) / 2: .CurrentY = 0
'    Printer.Print SeitenNr$
End With

Call DefErrPop
End Sub

Sub PZNPrüfZiffer(s As String)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("PZNPrüfZiffer")
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

Dim X As Long
Dim Pruefziffer As Integer
Dim t As String
Dim IZ As Integer
Dim IW As Long

X = Val(Left(s, 6)): Pruefziffer% = 10
t = Right$(String$(6, "0") + Mid(Str(X), 2) + "0", 7)
IW = 0
For IZ = 1 To 6
  IW = IW + Val(Mid(t, IZ, 1)) * (IZ + 1)
Next IZ
Pruefziffer% = (IW Mod 11)

If Pruefziffer% = 10 Then
  Mid(t, 7, 1) = " "            'ungültige PZN
Else
  Mid(t, 7, 1) = Right(Str(Pruefziffer%), 1)
  If X = 999999 Then Pruefziffer% = 10
End If
s = t
Call DefErrPop
End Sub

Sub DokuPflichtAnzeige()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("DokuPflichtAnzeige")
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
Dim i%
Dim h$

h$ = "Dokumentationspflicht gem. Transfusionsgesetz für folgende Artikel:" + vbCrLf + vbCrLf
For i% = 1 To UBound(DokuPflichtPzn$)
    h$ = h$ + DokuPflichtPzn$(i%) + vbCrLf
Next i%

Call MsgBox(h$, vbOKOnly, "Artikel mit Dokumentationspflicht")

Call DefErrPop
End Sub
                  

