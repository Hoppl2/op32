Attribute VB_Name = "modSenden"
Option Explicit

Type SerAuftrag
  i01 As Integer
  i02 As Integer
  i03 As Integer
  sta As Integer
  mds As Integer
  dat As Integer
  tim As Integer
  key As Integer
  i04 As Integer
  def As Integer
  i05 As Integer
  tfn As String * 25
  rfn As String * 25
  te1 As String * 20
  te2 As String * 20
  ke1 As String * 7
  ke2 As String * 7
  i06 As Integer
End Type

Const LINTECHPFAD$ = "C:\LINTECH\"

Const ModemSetup$ = "[00000R=0000]"
Const ModemAnswer$ = "[00000A=000360]": 'Wartezeit 360 Minuten = 60 Stunden
Const ModemDialP$ = "[00000W=000CRI"
Const ModemDialS$ = ";]"
Const empfange$ = "[00000E=000180]"
Const trennStr$ = "[00000T=00]"

Public BestSendenAbbruch%
Public BlindBestellung%

Public Lintech%, seriell%, ModemInDOS%

Public TelGh$, TelApo$
Public GhIDF$, ApoIDF$

Public SendeStatusBereichInd%

Public AnzMinutenWarnung%
Public AnzMinutenWarten%
Public AnzMinutenVerspaetung%

Public ModemOk%

Public SendeLog%

Public Bestaetigung$

Public AutomatikFertig%
Public AutomatikFehler$
Public AutomatikAktivSenden%

Public FaxDruckKopfTxt$
Public FaxDruckFussTxt$
Public FaxDruckFussHoehe%
Public AnzFaxDruckArtikel%, AnzFaxDruckPackungen%
Public FaxDruckWert#, FaxDruckWert2#

Public AnzDruckSpalten%
Public DruckSpalte() As DruckSpalteStruct
Public DruckFontSize%





Dim sap As SerAuftrag
Dim sap1 As SerAuftrag

Dim ISDNRet$(10)
Dim ISDNStatus$(10)

'Dim xParams$(6)
Dim SendPara$

Dim SRT$()
Dim mnr$()
Dim mbm%()
Dim mlauf&()

Dim Wg3Str$

Dim SendFile$, RecvFile$
Dim ZJOB%
Dim AktivSenden%
Dim RecWarte%, OrgRecWarte%

Dim folgenr%

Public LiefName1$, LiefName2$, LiefName3$, LiefName4$
Dim AuftragsErg$, AuftragsArt$
Dim HeuteDatStr$, xcBestDatum%

Dim LOGBUCH%, LEITUNGBUCH%, DRUCKBUCH%
Dim SendDruckDateiName$

Dim AbsagenZeilen$()
Dim AnzAbsagenZeilen%

Dim RueckmeldungenFlag%, AbsagenFlag%

Dim AbsagenWav%

Public SendeForm As Form

Public LeerAuftrag%

Public MaxSendSatz%
Public SendSatz$() '100

Public MacheEtiketten%
Public TeilDefekte%

Public CommEvent%
Public IsdnEndeDelay%

Public SHAREDDAT%
Public SharedDatSatz As String * 128
Public AUTOHEAD%
Public AutoHeadSatz As String * 24

Public DirektBezugErg$

Public PharmaBoxHandShake%

Public UebertragungOk%

Public AutomatikDrucker$

Public RueckKaufSendung%

Public ManuellSendung%

Public Wbestk2ManuellVorbereitung%

Public AbsagenMitNL$

Dim TimeOutZeit%, FehlerTimeOut%, FehlerBcc%, FehlerNak%, TimeOut%
Dim xParams$(6)

Dim AustriaSetup$, AustriaDialP$, AustriaDialS$, AustriaAnswer$, AustriaExit$, AustriaWait$
Dim ImSenden%
Dim ACK%, ackGH%

Dim MaxRecSatz%
Dim RecSatz$() '100

Dim RS$, STX$, ETX$, EOT$, ENQ$, NAK$, TDD$, WACK$, ACK0$, ACK1$, DLEEOT$

Dim RecStr$


Private Const DefErrModul = "SENDEN.BAS"

Sub InitIsdnAnzeige()

'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("InitIsdnAnzeige")
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
ISDNRet$(0) = "Treiber nicht geladen"
ISDNRet$(1) = "Tabelle voll"
ISDNRet$(2) = "falsche Auftragsnummer"
ISDNRet$(3) = "Auftrags-Tabelle gelockt"
ISDNRet$(4) = "falsche Funktion"
ISDNRet$(5) = "DFÜ nicht möglich bei diesem Auftrag"
ISDNRet$(6) = "DFÜ erfolglos nach n Wahlversuchen"
ISDNRet$(7) = "Interface-Karte fehlt/defekt"
ISDNRet$(8) = "Übertragungsprotokoll unzulässig"
ISDNRet$(9) = "Fehler beim Initialisieren"
ISDNRet$(10) = "Auftrag mit Key nicht vorhanden"

ISDNStatus$(1) = "bereit"
ISDNStatus$(2) = "erledigt"
ISDNStatus$(5) = "Fehler beim Verbindungsaufbau"
ISDNStatus$(6) = "Fehler beim Senden der Daten"
ISDNStatus$(7) = "Fehler beim Empfang der Daten"
ISDNStatus$(8) = "Fehler beim Sichern der Daten"
ISDNStatus$(9) = "gesperrt"
ISDNStatus$(10) = "Verbindungsabbruch ohne Rückmeldung"
Call DefErrPop
End Sub

Sub InitBestellung()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("InitBestellung")
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
Dim i%, ok%, d%
Dim h$

If (Lieferant% < 0) Then Call DefErrPop: Exit Sub

'ok% = TestModemPar%
'If (ok% = False) Then Call DefErrPop: Exit Sub

Call HoleLieferantenDaten(Lieferant%)

Call SucheSendeArtikel

AktivSenden% = False


With SendeForm
    If (RueckKaufSendung%) Then
        .Caption = "Rückkauf-Anfrage für " + LiefName1$
    Else
        .Caption = "Sendeauftrag für " + LiefName1$
    End If
    .txtAuftrag(0).text = "ZH"
    .txtAuftrag(1).text = "  "
    .optAuftrag(0).Value = True
    .chkAuftrag(0).Value = 1
    .chkAuftrag(1).Value = 1
    .lblLieferant(3).Caption = LiefName1$
    .lblLieferant(4).Caption = LiefName2$
    .lblLieferant(5).Caption = LiefName3$
    .lblLieferant(6).Caption = LiefName4$
    .lblLieferantWert(0).Caption = ApoIDF$
    .lblLieferantWert(1).Caption = GhIDF$
    .lblLieferantWert(2).Caption = TelGh$
    .fmeAuftrag.Caption = "Auftrag (" + Str$(AnzBestellArtikel%) + " Artikel)"
    
    .lblModemWert.Caption = ZeigeModemTyp$
End With

Call DefErrPop
End Sub
    
Function TestModemPar%()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("TestModemPar%")
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
Dim ok%, suche2%
Dim x$, h$

'def SEG = &H40
'  com1% = 256! * PEEK(1) + PEEK(0): com2% = 256! * PEEK(3) + PEEK(2)
'  LPT1% = 256! * PEEK(9) + PEEK(8): LPT2% = 256! * PEEK(11) + PEEK(10)
'def SEG
'Port% = LPT1%

If (para.isdn) Then
    Lintech% = False
    ok% = ModemParameter("LO-ISDN", SendPara$, False)
    If (ok% = False) Then
        ok% = ModemParameter("LINTECH", SendPara$, False)
        If (ok%) Then Lintech% = True
    End If
    If (ok% = False) Then
        para.isdn = False
    Else
        TestModemPar% = True
        Call DefErrPop: Exit Function
    End If
End If

seriell% = True
ok% = ModemParameter("MODEM-S", SendPara$, True)
If (ok% = False) Then
    suche2% = True
    seriell% = False
    ok% = ModemParameter("MODEM-P", SendPara$, False)
End If
If (ok% = False) Then
    If (suche2% = True) Then seriell% = True
    If (seriell%) Then x$ = "Seriell" Else x$ = "Parallel"
    x$ = x$ + "modem-Geräteparameter nicht vorhanden!"
    If (Wbestk2ManuellSenden%) Then
    Else
        Call iMsgBox(x$, vbCritical)
    End If
    TestModemPar% = False
    Call DefErrPop: Exit Function
End If
'If (seriell%) Then
'  If (Left$(xFilePara$, 4) = "COM1" And com1% = 0) Or (Left$(xFilePara$, 4) = "COM2" And com2% = 0) Or (Left$(xFilePara$, 3) <> "COM") Or (Left$(xFilePara$, 3) <> "COM") Or (InStr("12", Mid$(xFilePara$, 4, 1)) = 0) Then
'    Call meldung("Seriell-Modem an " + xFilePara$ + " nicht möglich!")
'    abbruch% = True: Return
'  End If
'Else
'  If (Left$(xFilePara$, 4) = "LPT1" And LPT1% = 0) Or (Left$(xFilePara$, 4) = "LPT2" And LPT2% = 0) Or (Left$(xFilePara$, 3) <> "LPT") Or (InStr("12", Mid$(xFilePara$, 4, 1)) = 0) Then
'    Call meldung("Parallel-Modem an " + xFilePara$ + " nicht möglich!")
'    abbruch% = True: Return
'  End If
'  If Left$(xFilePara$, 4) = "LPT1" Then Port% = LPT1% Else Port% = LPT2%
'End If

If (para.Land = "A") Then
    Call GetModemBefehle(xParams$())
    AustriaSetup$ = xParams$(1)
    AustriaDialP$ = xParams$(2)
    AustriaDialS$ = xParams$(3)
    AustriaAnswer$ = xParams$(4)
    AustriaExit$ = xParams$(5)
    If xParams$(6) = "" Then xParams$(6) = "ATS0=1" + vbCr
    AustriaWait$ = xParams$(6)
End If

TestModemPar% = True

Call DefErrPop
End Function

'Function ModemParameter%(TestName$, xFilePara$, Optional Parameter% = False)
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Call DefErrFnc("ModemParameter%")
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
'Dim i%, ok%, xq%, xParam%, DPA%, aa%
'Dim s$, X$, xGeraet$, xFileName$
'Dim XparamFix As String * 128
'Dim DparamFix As String * 40
'
'X$ = para.User
'If (Val(X$) = 0) Then X$ = ""
's$ = "\user\xparam" + X$ + ".dat"
'xParam% = FileOpen(s$, "RW", "R", Len(XparamFix))
'
''1 * belegt sonst nicht belegt, 8 Name,30 Schnittstelle incl.params
'i% = 1: ok% = False
'xFilePara$ = ""
'While (i% <= 9) And (ok% = 0)
'    Get xParam%, i% + 1, XparamFix
'    If (Left$(XparamFix$, 1) = "*") Then
'        X$ = Mid$(XparamFix$, 2, 8): X$ = LTrim$(X$): X$ = RTrim$(X$): xGeraet$ = X$
'        xFileName$ = "\user\" + X$ + ".dpa"
'        X$ = Mid$(XparamFix$, 10, 30): X$ = LTrim$(X$): X$ = RTrim$(X$): xFilePara$ = X$
'        If (Left$(xGeraet$, Len(TestName$)) = TestName$) Then ok% = True
'    End If
'    i% = i% + 1
'Wend
'Close #xParam%
'
'If (ok% And Parameter%) Then
'    DPA% = FileOpen(xFileName$, "RW", "R", Len(DparamFix))
'    For xq% = 1 To 5
'        Get #DPA%, xq%, DparamFix: X$ = DparamFix
'        X$ = RTrim$(X$): xParams$(xq%) = X$
'    Next xq%
'    Close #DPA%
'End If
'
'X$ = Mid$(xParams$(5), 21): X$ = LTrim$(X$): X$ = RTrim$(X$): xParams$(6) = X$
'X$ = Left$(xParams$(5), 20): X$ = LTrim$(X$): X$ = RTrim$(X$): xParams$(5) = X$
'
'ModemParameter% = ok%
'
'Call DefErrPop
'End Function

Sub SucheSendeArtikel()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("SucheSendeArtikel")
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
Dim i%, ind%, Max%
Dim h$, wg$

'Get #BEKART%, 1, Max%
'For i% = 1 To Max%
'    Get #BEKART%, i% + 1, bek
'    If ((Asc(bek.pzn) < 128) And (bek.Zugeordnet = "J")) Then
'        If (bek.ZuKontrollieren = "N") Or (bek.MussKontrollieren = "N") Or (bek.NochZuKontrollieren = "N") Then
'            h$ = Left$(bek.txt, 18) + Mid$(bek.txt, 29) + Format(i%, "0000") + bek.pzn + Format(Abs(bek.bm), "0000")
'            frmAction!lstSortierung.AddItem h$
'        End If
'    End If
'Next i%

    
'MaxZuSenden% = frmAction!lstSortierung.ListCount - 1
ReDim SRT$(AnzBestellArtikel%)
ReDim mnr$(AnzBestellArtikel%)
ReDim mbm%(AnzBestellArtikel%)
ReDim mlauf&(AnzBestellArtikel%)
ReDim SendSatz$(AnzBestellArtikel% + 10)

Wg3Str$ = ""

For i% = 0 To (AnzBestellArtikel% - 1)
    frmAction!lstSortierung.ListIndex = i%
    h$ = RTrim$(frmAction!lstSortierung.text)
    
    If (para.Land = "A") Then
        Wg3Str$ = Wg3Str$ + Right$(h$, 1)
        h$ = Left$(h$, Len(h$) - 1)
    End If
    
    SRT$(i%) = Left$(h$, 29)
    mnr$(i%) = Mid$(h$, 30, 7)
    mbm%(i%) = Val(Mid$(h$, 37, 4))
    mlauf&(i%) = Val(Mid$(h$, 41))
Next i%

Call DefErrPop
End Sub

Sub SendAutomatic()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("SendAutomatic")
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
Dim ret%, ind%, Ueberleiten%
Dim h$

AktivSenden% = AutomatikAktivSenden%
'RueckmeldungenFlag% = False
'AbsagenFlag% = True

Call GetSendungParameter

With SendeForm
'    AuftragsErg$ = Left$(RTrim$(.txtAuftrag(0).text) + Space$(2), 2)
'    AuftragsArt$ = Left$(RTrim$(.txtAuftrag(1).text) + Space$(2), 2)
    
    If (AktivSenden%) Then
        .optAuftrag(0).Value = True
    Else
        .optAuftrag(1).Value = True
    End If

End With

BestSendenAbbruch% = False

Call SucheSendeArtikel

Call SaetzeVorbereiten

ret% = SaetzeSenden%

If (ret%) Then
    If (RueckmeldungenFlag%) Then
        Call ZeigeRueckmeldungen
    End If
    If (AbsagenFlag%) Then
        Call AbsagenEntfernen
    End If
End If
        
AutomatikFertig% = ret%
'Unload frmSenden

Call DefErrPop
End Sub

Sub SendGermany()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("SendGermany")
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
Dim ret%, ind%, Ueberleiten%
Dim h$

Call GetSendungParameter

With SendeForm
    If (.optAuftrag(0).Value = True) Then
        AktivSenden% = True
    Else
        AktivSenden% = False
    End If
End With

BestSendenAbbruch% = False

Call SaetzeVorbereiten

ret% = SaetzeSenden%
    
If (ret%) Then
    If (RueckmeldungenFlag%) Then
        Call ZeigeRueckmeldungen
    End If
    If (AbsagenFlag%) Then
        Call AbsagenEntfernen
    End If
End If
        
If (AutomaticSend%) Then
'    Call UpdateBekartDat(Lieferant%, ret%)
    AutomatikFertig% = ret%
    Unload SendeForm
Else
    Ueberleiten% = False
    If (ret% = True) Then
        ret% = iMsgBox("Bestellung korrekt empfangen ?", vbYesNo Or vbDefaultButton2)
        If (ret% = vbYes) Then
            Ueberleiten% = True
        End If
    Else
        ret% = iMsgBox("Datenfernübertragung wiederholen ?", vbYesNo Or vbDefaultButton1)
        If (ret% <> vbYes) Then
            Ueberleiten% = True
        End If
    End If
    
    If (Ueberleiten%) Then
        If (RueckKaufSendung%) Then
            Call UpdateRueckKaufDat(Lieferant%, Ueberleiten%)
        Else
            Call UpdateBekartDat(Lieferant%, Ueberleiten%)
        End If
        
        UebertragungOk% = True
        
        Unload SendeForm
    Else
        Call SendeForm.SetFormModus(0)
    End If
End If

Call DefErrPop
End Sub

Sub GetSendungParameter()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("GetSendungParameter")
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
Dim ind%
Dim h$

With SendeForm
'    AuftragsErg$ = Left$(RTrim$(.txtAuftrag(0).text) + Space$(2), 2)
'    AuftragsArt$ = Left$(RTrim$(.txtAuftrag(1).text) + Space$(2), 2)
    
    h$ = UCase(Trim$(.cboAuftrag(0).text))
    ind% = InStr(h$, "(")
    If (ind% > 0) Then
        h$ = Mid$(h$, ind% + 1, 2)
    Else
        h$ = Left$(h$ + Space$(2), 2)
    End If
    AuftragsErg$ = h$
    
    h$ = UCase(Trim$(.cboAuftrag(1).text))
    ind% = InStr(h$, "(")
    If (ind% > 0) Then
        h$ = Mid$(h$, ind% + 1, 2)
    Else
        h$ = Left$(h$ + Space$(2), 2)
    End If
    AuftragsArt$ = h$
    
    RueckmeldungenFlag% = .chkAuftrag(0).Value
    AbsagenFlag% = .chkAuftrag(1).Value
'    If (para.Land = "A") Then
'        AbsagenFlag% = 0
'    End If
End With

Call DefErrPop
End Sub

Sub SaetzeVorbereiten()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("SaetzeVorbereiten")
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
Dim i%, j%, k%, gesendet%, menge%, IstFrei%
Dim satz$, char$, fistsatz$, BitFeld$, TestPZN$, sMenge$, txt$

If (para.Land = "A") Then
    Call SaetzeVorbereitenA
    Call DefErrPop: Exit Sub
End If

AnzAbsagenZeilen% = 0

If (BlindBestellung%) Then
    AktivSenden% = False
    AuftragsErg$ = "  "
    AuftragsArt$ = "  "
End If
    


'2 bytes satztyp
'3 bytes gesamtlänge

'startsatz
SendSatz$(0) = "01027" + GhIDF$ + ApoIDF$ + "3000    "
'                                       ^
'                                       S    Ruf durch GH
If Not (AktivSenden%) Then Mid$(SendSatz$(0), 24, 1) = "S"
'                    DS eventuell disponieren
'                    AA Artikelanzeige
'                    NL Nachliefern
'                    vv
'satz$(2) = "020231  DS00020108996ZH" 'ZH Zustellung Heute/ ZM Morgen
'            .....Aaa  MMMM#######xx --^^
'                  ^^
'                  ??
j% = 0

satz$ = "06256            " + para.Fistam(0) + "   " + para.Fistam(1)
For i% = 1 To Len(satz$)
  char$ = UCase(Mid$(satz$, i%, 1))
  Mid$(satz$, i%, 1) = char$
Next i%
satz$ = Left$(satz$ + Space$(256), 256)
fistsatz$ = satz$


'nur schicken wenn 3er Satz vorhanden
'J = J + 1: satz$(J) = satz$


BitFeld$ = ""

'1.58 Sätze nicht mehr lesen - sind schon in Arrays gespeichert
For i% = 0 To (AnzBestellArtikel% - 1)
  IstFrei% = TestBitFrei%(BitFeld$, i%)
  If (IstFrei%) Then
    TestPZN$ = mnr$(i%)
    menge% = Abs(mbm%(i%))
    k% = i% + 1
    While (k% < AnzBestellArtikel%)
      IstFrei% = TestBitFrei%(BitFeld$, k%)
      If (IstFrei%) Then
        If (mnr$(k%) = TestPZN$) Then
          menge% = menge% + Abs(mbm%(k%))
          Call SetBit(BitFeld$, k%)
        End If
      End If
      k% = k% + 1
    Wend
    sMenge$ = Right$("0000" + Mid$(Str$(menge%), 2), 4)
    If (TestPZN$ = "9999999") And (menge% <> 0) Then
      If (LeerAuftrag%) Then
        txt$ = Left$(SRT$(i), 25) + Space$(10)
      Else
        If (RueckKaufSendung%) Then
            rk.GetRecord (Val(Right$(SRT$(i), 4)) + 1)
            txt$ = rk.txt
        Else
            ww.GetRecord (Val(Right$(SRT$(i), 4)) + 1)
            txt$ = ww.txt
        End If
      End If
      satz$ = "030481aa  ####ppppdd" + Left$(txt$, 26) + "ää"
      '        12345678901234567890123
      '             ^     Auftragsnummer
      '                ^^ Artikelbezogener Hinweis
      Mid$(satz$, 7, 2) = AuftragsArt$
      Mid$(satz$, 11, 4) = sMenge$
      Mid$(satz$, 15, 4) = Mid$(txt$, 30, 4) 'Packungsgröße
      Mid$(satz$, 19, 2) = Mid$(txt$, 34, 2) 'Darreichungsform
      Mid$(satz$, 47, 2) = AuftragsErg$
    ElseIf (menge% <> 0) Then
      satz$ = "020231aa  ####0000000ää"
      '        12345678901234567890123
      '                      PZN
      '             ^     Auftragsnummer
      '                ^^ Artikelbezogener Hinweis
      'Laufnummer 00-99 später einbauen
      'mid$(satz$,9,2)=right$("00"+mid$(str$(j),2),2)
      Mid$(satz$, 7, 2) = AuftragsArt$
      Mid$(satz$, 11, 4) = sMenge$
      Mid$(satz$, 15, 7) = TestPZN$
      Mid$(satz$, 22, 2) = AuftragsErg$
    End If
    If (menge% <> 0) Then                   '1.59
      j% = j% + 1: SendSatz$(j%) = satz$
    End If
    Call SetBit(BitFeld$, i%)
  End If
Next i%

'lt. DrS Auftragsergänzung mit SA 06 schicken
If (AuftragsErg$ <> Space$(Len(AuftragsErg$)) And AuftragsErg$ <> "ZH") Then
  Mid$(fistsatz$, 15, 1) = "A"
  Mid$(fistsatz$, 16, 2) = AuftragsErg$
  j% = j% + 1: SendSatz$(j%) = fistsatz$
End If

MaxSendSatz% = j% + 1
'endsatz
SendSatz$(MaxSendSatz%) = "99019" + GhIDF$ + ApoIDF$

Call DefErrPop
End Sub

Function TestBitFrei%(BitFeld$, bit%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("TestBitFrei%")
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
Dim ok%, ind%, wert%, iBit%

iBit% = bit% + 1

ok% = True
ind% = (iBit% - 1) \ 8 + 1
wert% = 2 ^ (7 - ((iBit% - 1) Mod 8))
If (Len(BitFeld$) < ind%) Then
  BitFeld$ = BitFeld$ + String$(ind% - Len(BitFeld$), 0)
End If
If (Asc(Mid$(BitFeld$, ind%, 1)) And wert%) > 0 Then ok% = False

TestBitFrei% = ok%
Call DefErrPop
End Function

Sub SetBit(BitFeld$, bit%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("SetBit")
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
Dim ind%, wert%, iBit%

iBit% = bit% + 1

ind% = (iBit% - 1) \ 8 + 1
wert% = 2 ^ (7 - ((iBit% - 1) Mod 8))
If (Len(BitFeld$) < ind%) Then
  BitFeld$ = BitFeld$ + String$(ind% - Len(BitFeld$), 0)
End If
Mid$(BitFeld$, ind%, 1) = Chr$(Asc(Mid$(BitFeld$, ind%, 1)) Or wert%)

Call DefErrPop
End Sub

Function SaetzeSenden%()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("SaetzeSenden%")
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
Dim j%, sFile%, ret%, erg%
Dim l2&
Dim l$, h$, h2$, edat$

SendeLog% = True

AbsagenWav% = False

edat$ = Format(Day(Date), "00") + Format(Month(Date), "00") + Format(Year(Date), "0000")
h$ = Left$(edat$, 4) + Right$(edat$, 2)
HeuteDatStr$ = h$
xcBestDatum% = iDate(h$)

If (SendeLog%) Then
    erg% = wpara.CreateDirectory("winw")
    h2$ = "winw\" + Format(Year(Date) Mod 100, "00") + Format(Month(Date), "00") + Format(Day(Date), "00")
    erg% = wpara.CreateDirectory(h2$)
    SendDruckDateiName$ = Format(Lieferant%, "000") + Format(Now, "HHMM")
'    If (AutomaticSend% = False) Then
    If (ManuellSendung%) Then
        SendDruckDateiName$ = SendDruckDateiName$ + "m"
    End If
    
    If (Dir("winw\*.sp9") <> "") Then
        On Error Resume Next
        Kill "winw\*.sp1"
        For j% = 2 To 9
            h$ = Dir("winw\*.sp" + Format(j%, "0"))
            If (h$ <> "") Then
                h$ = "winw\" + h$
                Name h$ As Left$(h$, Len(h$) - 1) + Format(j% - 1, "0")
            End If
        Next j%
        On Error GoTo DefErr
    End If
    
    h2$ = h2$ + "\" + SendDruckDateiName$
    
    LOGBUCH% = FreeFile
    Open h2$ + ".log" For Output As LOGBUCH%
    For j% = 0 To MaxSendSatz%
      Print #LOGBUCH%, ">" + SendSatz$(j%)
    Next j%
    
    If (ModemInDOS% = False) Then
        LEITUNGBUCH% = FreeFile
        Open h2$ + ".deb" For Output As LEITUNGBUCH%
    End If
End If

If (BlindBestellung%) Then
    MaxSendSatz% = 0
    ret% = True
ElseIf (para.isdn) Then
    Call LoescheLintechDateien
    Call StarteLintechServer
    
'    Call RueckMeldungTestDatei

    SendFile$ = LINTECHPFAD$ + "S_" + Right$(String$(5, 48) + Mid$(Str$(Lieferant%), 2), 5) + "1" + ".I01"
    RecvFile$ = LINTECHPFAD$ + "S_" + Right$(String$(5, 48) + Mid$(Str$(Lieferant%), 2), 5) + "1" + ".RUK"

    sFile% = FreeFile
    Open SendFile$ For Output As #sFile%
    For j% = 0 To MaxSendSatz%
      Print #sFile%, SendSatz$(j%) + Chr$(10);
    Next j%
    Close #sFile%
    
    ret% = ISDNBestellung%
    
    Call LoescheLintechDateien(ret%)
ElseIf (ModemInDOS%) Then
    Call LoeschePharmaboxDateien
    Call StartePharmaboxServer(CurDir + "\" + h2$ + ".deb")
    
    ret% = PharmaboxBestellung%
    
    Call LoeschePharmaboxDateien
Else
'    ret% = ModemAktivieren%
    ret% = True
    If (ret% = True) Then
        ret% = SeriellBestellung%
        ' Serielle Schnittstelle schließen.
'        frmAction.comSenden.PortOpen = False
    End If
End If


If (SendeLog%) Then
    Close (LOGBUCH%)
    If (ModemInDOS% = False) Then Close (LEITUNGBUCH%)
End If

SaetzeSenden% = ret%
Call DefErrPop
End Function

Sub LoescheLintechDateien(Optional MitWarten% = False)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("LoescheLintechDateien")
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
Dim l&
Dim SuchPattern$, h$

If (MitWarten%) Then
    For i% = 1 To IsdnEndeDelay%
        Call IsdnPause
    Next i%
End If

On Error Resume Next
l& = EntferneTask("Isdn Server")

Kill LINTECHPFAD$ + "Z.JOB"

For i% = 1 To 2
    If (i% = 1) Then
        SuchPattern$ = LINTECHPFAD$ + "S_*" + ".I0*"
    Else
        SuchPattern$ = LINTECHPFAD$ + "S_*" + ".RUK"
    End If
    
    h$ = Dir$(SuchPattern$)
    Do
        If (Trim$(h$) = "") Then Exit Do
        Kill LINTECHPFAD$ + h$
        h$ = Dir
    Loop
Next i%

Call DefErrPop
End Sub

Sub StarteLintechServer()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("StarteLintechServer")
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
Dim l&, TaskId&

On Error Resume Next
l& = GetTaskId("Isdn Server")
If (l& = 0) Then
    TaskId = Shell("c:\lintech\server.exe", vbNormalNoFocus)
    SendeForm.SetFocus
    DoEvents
End If

Call DefErrPop
End Sub

Sub LoeschePharmaboxDateien()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("LoeschePharmaboxDateien")
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
Dim l&

On Error Resume Next
l& = EntferneTask("PharmBox")
Kill "\user\PharmBox.Out"
Kill "\user\PharmBox.In"
Kill "\user\PharmBox.log"
On Error GoTo DefErr

Call DefErrPop
End Sub

Sub StartePharmaboxServer(LeitungsProtDatei$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("StartePharmaboxServer")
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
Dim j%, sFile%
Dim l&, TaskId&
Dim h$

sFile% = FreeFile
Open "\user\PharmBox.Out" For Output As #sFile%

Print #sFile%, SendPara$
Print #sFile%, ModemSetup$
Print #sFile%, ModemAnswer$
Print #sFile%, ModemDialP$
Print #sFile%, ModemDialS$
Print #sFile%, empfange$
Print #sFile%, trennStr$
Print #sFile%, LeitungsProtDatei$

If (AktivSenden%) Then
    h$ = "A" + TelGh$
ElseIf (AutomaticSend%) Then
    h$ = "P" + Format$(AnzMinutenWarten%, "0")
Else
    h$ = "P" + Format(360, "0")
End If
Print #sFile%, h$

For j% = 0 To MaxSendSatz%
  Print #sFile%, SendSatz$(j%)
Next j%

Close #sFile%
    

ZJOB% = FreeFile
Open "\user\PharmBox.log" For Binary Access Read Write Shared As #ZJOB%

If (LOF(ZJOB%) = 0) Then
  Lock ZJOB%
  h$ = "0"
  Put ZJOB%, , h$
  Unlock ZJOB%
End If



On Error Resume Next
'l& = GetTaskId("Isdn Server")
'If (l& = 0) Then
'    TaskId = Shell("PharmBox.bat", vbNormalNoFocus)
    TaskId = Shell("\user\PharmBox.bat", vbMinimizedNoFocus)
    Call IsdnPause
    SendeForm.SetFocus
    DoEvents
'End If

Call DefErrPop
End Sub

Function ISDNBestellung%()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("ISDNBestellung%")
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
Dim rc%, sta%, fehler%
Dim MeldungStr$

rc% = ISDNBest%(sta%)

If (rc% = 1) Then
    Call ISDNRueckMeldungen
      
    fehler% = 0
    On Error GoTo ErrorISDNBestellung
    Kill RecvFile$
    Kill SendFile$
    On Error GoTo 0
    ISDNBestellung% = True
Else
    If (rc% = 99) Then rc% = 0
    If (rc% = -1) Then
        MeldungStr$ = "Datenübertragung im Hintergrund  - Defekte werden nicht berücksichtigt!"
    ElseIf rc% = -2 Then
        MeldungStr$ = "Datenübertragung abgebrochen!"
    ElseIf (rc% = 0) Then
        MeldungStr$ = "Fehler bei ISDN-Übertragung: wartet auf Anruf"
    ElseIf (rc% = 1) Then
        MeldungStr$ = "Fehler bei ISDN-Übertragung: Rückmeldungen vorhanden"
    ElseIf (rc% = 2) Then
        MeldungStr$ = "Fehler bei ISDN-Übertragung: undokumentiert"
    ElseIf (rc% = 3) Then
        MeldungStr$ = "Fehler bei ISDN-Übertragung: Übertragungsfehler"
    ElseIf (rc% = 4) Then
        MeldungStr$ = "Fehler bei ISDN-Übertragung: S_XXXXXX.i0X fehlt"
    ElseIf (rc% = 5) Then
        MeldungStr$ = "Fehler bei ISDN-Übertragung: > maximale Satzlänge"
    ElseIf (rc% = 6) Then
        MeldungStr$ = "Fehler bei ISDN-Übertragung: Wählversuche gescheitert"
    ElseIf (rc% = 7) Then
        MeldungStr$ = "Fehler bei ISDN-Übertragung: Rückmeldungen nicht übertragbar"
    ElseIf (rc% = 8) Then
        MeldungStr$ = "Fehler bei ISDN-Übertragung: Wählstring-Fehler"
    ElseIf (rc% = 9) Then
        MeldungStr$ = "Fehler bei ISDN-Übertragung: GH nimmt Satz 06 nicht an"
    ElseIf (rc% = 10) Then
        MeldungStr$ = "Fehler bei ISDN-Übertragung: ankommende GH-BGA nicht in Z.JOB"
    Else
        MeldungStr$ = "Fehler bei ISDN-Übertragung unbekannt:" + Str$(rc%)
    End If
    
    If (AutomaticSend% = False) Then
        Call iMsgBox(MeldungStr$, vbCritical)
    Else
        Call StatusZeile(MeldungStr$)
        AutomatikFehler$ = MeldungStr$
    End If
  
    If (rc% <> -1) Then Kill SendFile$
    ISDNBestellung% = False
End If

Call DefErrPop: Exit Function

ErrorISDNBestellung:
    fehler% = Err
    Err = 0
    Resume Next
    Return


Call DefErrPop
End Function

Function SeriellBestellung%()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("SeriellBestellung%")
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
Dim ret%, ret2%
Dim SendStr$, RecStr$
Dim MeldungStr$

If (para.Land = "A") Then
    SeriellBestellung% = SeriellBestellungA%
    Call DefErrPop: Exit Function
End If

ret% = SeriellBest%(MeldungStr$)

If (ret% = True) Then
    RecWarte% = 120
    ret% = SeriellRueckMeldungen%
    SendStr$ = trennStr$
    ret2% = SeriellSend%(SendStr$, RecStr$)
    MeldungStr$ = "Abbruch bei Empfang der Rückmeldungen"
End If
If (ret% = False) Then
    Call StatusZeile(MeldungStr$)
    If (AutomaticSend% = False) Then
        Call iMsgBox(MeldungStr$, vbCritical)
    Else
        AutomatikFehler$ = MeldungStr$
    End If
End If

SeriellBestellung% = ret%

Call DefErrPop
End Function

Function SeriellBest%(MeldungStr$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("SeriellBest%")
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
Dim j%, ret%, Warte%, OrgWarte%
Dim status$, SendStr$, RecStr$, satz01$, deb$

SeriellBest% = False

RecWarte% = 2

status$ = "Warten auf Modembereitschaft ...": Call StatusZeile(status$)
ret% = SeriellReceive%(RecStr$)
deb$ = "< " + RecStr$: Call StatusZeile(deb$)
'If (ret% = False) Then Call DefErrPop: Exit Function
'

RecWarte% = 120: OrgRecWarte% = RecWarte%
RecWarte% = 30

status$ = "Modem-Initialisierung wird durchgeführt.": Call StatusZeile(status$)
SendStr$ = ModemSetup$
ret% = SeriellSend%(SendStr$, RecStr$)
MeldungStr$ = "Fehler bei Modem-Initialisierung"
If (ret% = False) Then Call DefErrPop: Exit Function

RecWarte% = OrgRecWarte%

If (AktivSenden%) Then
    status$ = "Telefonnummer " + TelGh$ + " wird gewählt.": Call StatusZeile(status$)
    folgenr% = 0
    SendStr$ = ModemDialP$ + TelGh$ + ModemDialS$
    ret% = SeriellSend%(SendStr$, RecStr$)
    MeldungStr$ = "Fehler beim Wählen"
    If (ret% = False) Then Call DefErrPop: Exit Function
    SendStr$ = empfange$
    ret% = SeriellSend%(SendStr$, RecStr$)
    satz01$ = RecStr$
    MeldungStr$ = "Fehler beim Empfangen"
    If (ret% = False) Then Call DefErrPop: Exit Function
Else
    If (AutomaticSend%) Then
        status$ = "Warten auf Anruf (max." + Str$(AnzMinutenWarten%) + " Minuten) ..."
    Else
        status$ = "Warten auf Anruf (max. 6 Stunden) ..."
    End If
    Call StatusZeile(status$)
    folgenr% = 0
    SendStr$ = ModemAnswer$
    If (AutomaticSend%) Then
        RecWarte% = AnzMinutenWarten% * 60
    Else
        RecWarte% = 6 * 60 * 60
    End If
    ret% = SeriellSend%(SendStr$, RecStr$)
    RecWarte% = OrgRecWarte%
    MeldungStr$ = "Fehler beim Warten auf Anruf"
    If (ret% = False) Then Call DefErrPop: Exit Function
    SendStr$ = empfange$
    ret% = SeriellSend%(SendStr$, RecStr$)
    satz01$ = RecStr$
    MeldungStr$ = "Fehler beim Empfangen"
    If (ret% = False) Then Call DefErrPop: Exit Function
End If

If (Mid$(RecStr$, 9, 1) <> "3") Then
    status$ = "Teilnehmer antwortet nicht !": Call StatusZeile(status$)
    MeldungStr$ = status$
    Call DefErrPop: Exit Function
End If

If (Mid$(satz01$, 20, 7) <> GhIDF$) Then
    status$ = "falsche Großhandlung " + Mid$(RecStr$, 20, 7) + "!": Call StatusZeile(status$)
    MeldungStr$ = status$
    Call DefErrPop: Exit Function
End If

status$ = "Bestellung wird gesendet ...": Call StatusZeile(status$)
Mid$(SendSatz$(0), Len(SendSatz$(0)) - 3, 4) = Mid$(satz01$, Len(satz01$) - 4, 4)
For j% = 0 To MaxSendSatz%
    SendStr$ = "[00000D=00" + SendSatz$(j%) + "]"
    ret% = SeriellSend%(SendStr$, RecStr$)
    MeldungStr$ = "Fehler bei Übertragung"
    If (ret% = False) Then Call DefErrPop: Exit Function
Next j%

MeldungStr$ = ""

SeriellBest% = True

Call DefErrPop
End Function

Function SeriellSend%(SendStr$, RecStr$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("SeriellSend%")
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
Dim ret%
Dim l$, deb$

Mid$(SendStr$, 6, 1) = Right$(Str$(folgenr%), 1)
l$ = Mid$(Str$(Len(SendStr$) - 2), 2)
l$ = Right$("0000" + l$, 4)
Mid$(SendStr$, 2, 4) = l$
Call MakeBcc(SendStr$)

SendeForm.comSenden.Output = SendStr$
deb$ = "> " + SendStr$: Call StatusZeile(deb$)

folgenr% = folgenr% + 1

ret% = SeriellReceive%(RecStr$)
deb$ = "< " + RecStr$: Call StatusZeile(deb$)

If (ret%) Then RecStr$ = Left$(RecStr$, Len(RecStr$) - 1)

SeriellSend% = ret%

Call DefErrPop
End Function

Function SeriellReceive%(RecStr$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("SeriellReceive%")
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
Dim l%, ret%
Dim char$, deb$
Dim TimerEnd
Dim ch As Variant
Dim chByte() As Byte

ret% = False

RecStr$ = ""
TimerEnd = Timer + RecWarte%
Do
    If (Timer > TimerEnd) Then Exit Do
    
    
    l% = SendeForm.comSenden.InBufferCount
    
    If (CommEvent% <> 0) Then
        deb$ = "- " + Str$(CommEvent%) + Str$(l%): Call StatusZeile(deb$)
        CommEvent% = 0
    End If
    
    If (l% > 0) Then
'        char$ = SendeForm.comSenden.Input
        ch = SendeForm.comSenden.Input
        chByte = ch
        char$ = Chr$(chByte(0))
        RecStr$ = RecStr$ + char$
        If (ret%) Then Exit Do
        If (char$ = "]") Then
            ret% = True
            TimerEnd = Timer + 2
        End If
    Else
        DoEvents
        If (BestSendenAbbruch% = True) Then Exit Do
    End If
Loop

SeriellReceive% = ret%

Call DefErrPop
End Function

'alt
'Function SeriellReceive%(RecStr$)
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Call DefErrFnc("SeriellReceive%")
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
'Dim ok%, l%, ret%
'Dim char$
'Dim timeranf
'
'ret% = False
'
'RecStr$ = ""
'ok% = 0
'timeranf = Timer
'Do
'    If (Timer - timeranf > RecWarte%) Then Exit Do
'
'    l% = SendeForm.comSenden.InBufferCount
'    If (l% > 0) Then
'        char$ = SendeForm.comSenden.Input
'        RecStr$ = RecStr$ + char$
'        If (ok%) Then
'            ret% = True
'            Exit Do
'        End If
'        If (char$ = "]") Then ok% = 1
'    Else
'        DoEvents
'        If (BestSendenAbbruch% = True) Then Exit Do
'    End If
'Loop
'
'SeriellReceive% = ret%
'
'Call DefErrPop
'End Function

'ganz alt
'Function SeriellReceive%(RecStr$)
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Call DefErrFnc("SeriellReceive%")
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
'Dim ok%, l%
'Dim char$, bcc$, deb$
'Dim timeranf
'Dim vchar As Variant
'Dim bchar() As Byte
'
'SeriellReceive% = False
'
'RecStr$ = ""
'ok% = 0
'timeranf = Timer
'While (ok% = 0)
'
'    If (Timer - timeranf > RecWarte%) Then ok% = 2
'
'    l% = SendeForm.comSenden.InBufferCount
'    If (l% > 0) Then
'        char$ = SendeForm.comSenden.Input
'        RecStr$ = RecStr$ + char$
'        If (char$ = "]") Then ok% = 1
'    Else
'        DoEvents
'        If (BestSendenAbbruch% = True) Then
'            ok% = 2
'        End If
'    End If
'Wend
'If (ok% = 1) Then
'    char$ = SendeForm.comSenden.Input
'    bcc$ = char$
'    SeriellReceive% = True
'End If
'
'Call DefErrPop
'End Function

Sub MakeBcc(SendStr$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("MakeBcc")
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
Dim i%, bcc%, l%
Dim str1$

bcc% = 0
str1$ = SendStr$
l% = Len(SendStr$)
For i% = 2 To l%
  str1$ = Right$(str1$, l% - i% + 1)
  bcc% = (bcc% Xor Asc(Left$(str1$, 1)))
Next i%
SendStr$ = SendStr$ + Chr$(bcc%)

Call DefErrPop
End Sub

Function ISDNBest%(sta%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("ISDNBest%")
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
Dim hSendFile$, hRecvFile$
Dim hTelGH$, hTelApo$
Dim hGhIDF$, hApoIDF$, ae$, satz$, modus$, status$, sFile$, rFile$, TGH$, TApo$, IGH$, IApo$, h2$
Dim hLieferant%, ind%, ret%, FileLen%, Funkt%, RAlt%
Dim StartSek, IstSek

Dim Interval!

'Parameter vor Überschreiben schützen
hSendFile$ = SendFile$
hRecvFile$ = RecvFile$
hTelGH$ = TelGh$
hTelApo$ = TelApo$
hGhIDF$ = GhIDF$
hApoIDF$ = ApoIDF$
hLieferant% = Lieferant%

ind% = hLieferant% * 10 + 1

'f.tmp% = 1
FileLen% = 134
satz$ = String$(FileLen%, 0)

'CALL KopfZeile(xh$)
Interval! = 1

'Datensatz zusammenstellen
ret% = 0
sap.sta = 0
sap1.sta = -1
If Not (AktivSenden%) Then sap.mds = 0 Else sap.mds = 2
'datum = ((Val(Right$(Date$, 4)) - 1980) * 16 + Val(Left$(Date$, 2))) * 32 + Val(Mid$(Date$, 4, 2))
'zeit = (Val(Left$(Time$, 2)) * 64) + Val(Mid$(Time$, 4, 2))
'Auftrag% = hLieferant% * 10 + 1

ZJOB% = FreeFile
Open LINTECHPFAD$ + "Z.JOB" For Binary Access Read Write Shared As #ZJOB%

If (LOF(ZJOB%) = 0) Then
  Lock ZJOB%
  ae$ = Space$(62) + Chr$(10)
  Put ZJOB%, , ae$
  Unlock ZJOB%
End If

If (sap.mds = 0) Then
  modus$ = "Warten auf Anruf "
ElseIf (sap.mds = 1) Then
  modus$ = "Anruf durchführen "
ElseIf (sap.mds = 2) Then
  modus$ = "Anruf sofort durchführen "
Else
  modus$ = "Unbekannt "
End If
modus$ = modus$ + "(" + Str$(sap.mds) + ")"
status$ = Time$ + " Modus = " + modus$
Call StatusZeile(status$)

'Einfügen
Call WriteZjob

'    h2$ = Format$(Now, "HH:MM:SS")
'    StartSek& = Val(Left$(h2$, 2)) * 3600 + Val(Mid$(h2$, 4, 2)) * 60 + Val(Right$(h2$, 2))
StartSek = Timer

While (ret% = 0)  'solange nicht übertragen

    DoEvents
    If (BestSendenAbbruch% = True) Then
        ret% = -1
'        ret% = 1
    End If

    If (ret% = 0) Then
        ret% = CheckZjob%
        Call IsdnPause
    End If

    If (ret% = 0) And (AktivSenden% = False) And (sap.sta = 0) Then
'            h2$ = Format$(Now, "HH:MM:SS")
'            IstSek& = Val(Left$(h2$, 2)) * 3600 + Val(Mid$(h2$, 4, 2)) * 60 + Val(Right$(h2$, 2))
        IstSek = Timer
        If ((IstSek - StartSek) > (AnzMinutenWarten% * 60)) Then
            ret% = -1
        End If
    End If
    
Wend

If (ret% = -1) Then ret% = -2
Call LoescheZjob

Close ZJOB%

sta% = sap.sta

ISDNBest% = ret%
Call DefErrPop
End Function

Sub IsdnPause(Optional PauseSekunden% = 1)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("IsdnPause")
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
Dim StartSek

StartSek = Timer
Do
    If ((Timer - StartSek) > PauseSekunden%) Then Exit Do   'or (timer<2)
    DoEvents
Loop

Call DefErrPop
End Sub


Function CheckFertig%()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("CheckFertig%")
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
Dim ind%, Funkt%, ret%

ind% = Lieferant% * 10 + 1
Funkt% = 12
ret% = IsdnSer%(Funkt%, ind%, sap)

CheckFertig01:
Call DisplayStatus(Funkt%, ret%)
If (sap.mds = 1) And (sap.sta = 6 Or sap.sta = 7) Then
  Funkt% = 11    'Verbindung unterbrochen, Auftrag sofort nochmal schicken
  sap.sta = 1
  ret% = IsdnSer%(Funkt%, ind%, sap)
  Call DisplayStatus(Funkt%, ret%)

  GoTo CheckFertig01
End If

CheckFertig% = ret%

Call DefErrPop
End Function

Sub DisplayStatus(Funkt%, ret%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("DisplayStatus")
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
Static AltFunkt%
Dim anzeigen%
Dim status$

status$ = ""

anzeigen% = False
If (sap1.sta <> sap.sta) Then anzeigen% = True
If (sap1.mds <> sap.mds) Then anzeigen% = True
If (AltFunkt% <> Funkt%) Then anzeigen% = True
If (anzeigen%) Then
    AltFunkt% = Funkt%
    sap1 = sap
    status$ = status$ + Time$ + " "
    If (Funkt% = 1) Then
        status$ = status$ + "Auftrag vorbereiten "
    ElseIf (Funkt% = 2) Then
        status$ = status$ + "Auftrag löschen     "
    ElseIf (Funkt% = 5) Then
        status$ = status$ + "Auftrag freigeben   "
    ElseIf (Funkt% = 11) Then
        status$ = status$ + "Auftrag wiederholen "
    ElseIf (Funkt% = 12) Then
        status$ = status$ + "Statusabfrage       "
    Else
        status$ = status$ + "Funkt:" + Str$(Funkt%) + " "
    End If
    If (sap.sta <= 9) Then
        status$ = status$ + LTrim$(ISDNStatus$(sap.sta)) + " "
    ElseIf (sap.sta = 10) Then
        status$ = status$ + "Senden aktiv "
    Else
        status$ = status$ + "Status unbekannt (" + Mid$(Str$(sap.sta), 2) + ") "
    End If
    If (Funkt% = 1) Then
        status$ = status$ + "Modus="
        If (sap.mds = 1) Then
            status$ = status$ + "Anruf durchführen "
        ElseIf (sap.mds = 2) Then
            status$ = status$ + "Warten auf Anruf  "
        ElseIf (sap.mds = 9) Then
            status$ = status$ + "Gesperrt          "
        Else
            status$ = status$ + "Unbekannt :" + Str$(sap.mds)
        End If
    End If
    status$ = status$ + "(" + Mid$(Str$(ret%), 2) + ") "
    Call StatusZeile(status$)
End If

Call DefErrPop
End Sub

Sub WriteZjob()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("WriteZjob")
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
Dim wPos&
Dim satz$

wPos& = -1
Call iLock(ZJOB%, 1)
Seek ZJOB%, 64
satz$ = Space$(142)
Do While Not (EOF(ZJOB%))
  Get ZJOB%, , satz$
  If (Left$(satz$, 7) = GhIDF$) Then
    wPos& = Seek(ZJOB%) - 142
    Exit Do
  End If
Loop
If (wPos& < 0) Then wPos& = LOF(ZJOB%) + 1
satz$ = Left$(GhIDF$ + Space$(7), 7) + Right$(String$(5, 48) + Mid$(Str$(Lieferant%), 2), 5)
satz$ = satz$ + "11         72110" + Left$(ApoIDF$ + Space$(7), 7) + "01200T105    "
satz$ = satz$ + Left$(TelGh$ + Space$(31), 31) + "1" + Mid$(Str$(sap.mds), 2) + "0"
satz$ = Left$(satz$ + Space$(141), 141) + Chr$(10)
Put ZJOB%, wPos&, satz$
Call iUnLock(ZJOB%, 1)

Call DefErrPop
End Sub

Function ReadZjob%()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("ReadZjob%")
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
Dim ret%
Dim satz$

Call iLock(ZJOB%, 1)
Seek #ZJOB%, 64
satz$ = Space$(142)
Do While Not (EOF(ZJOB%))
  Get ZJOB%, , satz$
  If (Left$(satz$, 7) = GhIDF$) Then
    Exit Do
  End If
Loop
Call iUnLock(ZJOB%, 1)

ret% = Val(Mid$(satz$, 28, 1))
sap.sta = Asc(Mid$(satz$, 36, 1)) - 48
sap.mds = Val(Mid$(satz$, 81, 1))

ReadZjob% = ret%
Call DefErrPop
End Function

Function CheckZjob%()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("CheckZjob%")
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
Dim ret%, ind%
Dim stat$, status$

'LOCATE 15, 70: Print Time$;

ind% = Lieferant% * 10 + 1

CheckZjob01:
ret% = ReadZjob
If (ret% = 0) Then
  If (sap1.sta <> sap.sta) Then
    sap1 = sap

    If (sap.sta = 0) Then
      stat$ = "bereit "
    Else
      stat$ = "Anruf eingeleitet "
    End If

    stat$ = stat$ + "(" + Str$(sap.sta) + ")"
    status$ = Time$ + " Status = " + stat$
    Call StatusZeile(status$)
  End If
End If
If (ret% = 3) Then
  Call WriteZjob    'Verbindung unterbrochen, Auftrag sofort nochmal schicken
'  GoTo CheckFertig01
End If

CheckZjob% = ret%

Call DefErrPop
End Function

Sub LoescheZjob()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("LoescheZjob")
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
Dim i%, aeJOB%
Dim ae$, satz$

aeJOB% = FreeFile
Open LINTECHPFAD$ + "AE.JOB" For Output As #aeJOB%
ae$ = Space$(62) + Chr$(10)
Print #aeJOB%, ae$;

Call iLock(ZJOB%, 1)
Seek ZJOB%, 64
satz$ = Space$(142)
i% = 0
Do
  Get ZJOB%, 63 + (i% * 142) + 1, satz$
  If (EOF(ZJOB%)) Then Exit Do
  If (Left$(satz$, 7) <> GhIDF$) Then
    Print #aeJOB%, satz$;
  End If
  i% = i% + 1
Loop
Call iUnLock(ZJOB%, 1)

Close aeJOB%
Close ZJOB%
Kill LINTECHPFAD$ + "Z.JOB"
Name LINTECHPFAD$ + "AE.JOB" As LINTECHPFAD$ + "Z.JOB"

Open LINTECHPFAD$ + "Z.JOB" For Binary Access Read Write Shared As #ZJOB%

Call DefErrPop
End Sub

Function PharmaboxBestellung%()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("PharmaboxBestellung%")
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
Dim ret%
Dim MeldungStr$, satz$

Do
    DoEvents
    If (BestSendenAbbruch% = True) Then
        Call iLock(ZJOB%, 1)
        satz$ = "4"
        Put ZJOB%, 1, satz$
        Call iUnLock(ZJOB%, 1)
        BestSendenAbbruch% = False

'        MeldungStr$ = "Abbruch durch Benutzer"
'        ret% = -1
'        Exit Do
    End If

    ret% = CheckPharmaboxZjob%(MeldungStr$)
    If (ret% <> 0) Then Exit Do
    
    Call IsdnPause
Loop

ret% = ret% + 1

If (ret% = True) Then
    Call PharmaboxRueckMeldungen
Else
    If (AutomaticSend% = False) Then
        Call iMsgBox(MeldungStr$, vbCritical)
    Else
        AutomatikFehler$ = MeldungStr$
    End If
End If

Close ZJOB%

PharmaboxBestellung% = ret%

Call DefErrPop
End Function

Function CheckPharmaboxZjob%(MeldungStr$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("CheckPharmaboxZjob%")
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
Dim ret%, status%, ind%
Dim satz$, h$

ret% = 0

Call iLock(ZJOB%, 1)
satz$ = Space$(LOF(ZJOB%))
Get ZJOB%, 1, satz$
Call iUnLock(ZJOB%, 1)

status% = Val(Left$(satz$, 1))
satz$ = Mid$(satz$, 2)
Call OemToChar(satz$, satz$)
If (status% <> 0) And (status% <> 4) Then
    With SendeForm.flxAuftrag
        .redraw = False
        .Rows = 0
        Do
            ind% = InStr(satz$, vbCr)
            If (ind% > 0) Then
                h$ = Left$(satz$, ind% - 1)
                satz$ = Mid$(satz$, ind% + 1)
                .AddItem h$
            Else
                Exit Do
            End If
        Loop
        
        MeldungStr$ = h$
    
        .redraw = True
        If (.Rows > 4) Then
            .TopRow = .Rows - 4
            If (.FixedRows > 0) Then
                .TopRow = .TopRow + .FixedRows
            End If
        End If
        
        If (.Rows = 0) Then .AddItem ""
        
        .row = .Rows - 1
    End With
    
    If (status% = 3) Then
        status% = 0
        Call iLock(ZJOB%, 1)
        satz$ = "0"
        Put ZJOB%, 1, satz$
        Call iUnLock(ZJOB%, 1)
    End If
    
    ret% = status% * (-1)
End If

CheckPharmaboxZjob% = ret%

Call DefErrPop
End Function

Sub PharmaboxRueckMeldungen()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("PharmaboxRueckMeldungen")
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

RecvFile$ = "\user\PharmBox.In"
Call ISDNRueckMeldungen

Call DefErrPop
End Sub

Sub ISDNRueckMeldungen()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("ISDNRueckMeldungen")
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
Dim i%, j%, rFile%, SATZLEN%, ende%, satz%, fehler%
Dim recv$, Oldrecv$, satz1$

MaxSendSatz% = 0
ReDim SendSatz$(MaxSendSatz%)
SendSatz$(0) = Space$(14) + "DATENÜBERTRAGUNG RÜCKMELDUNGEN: "

ende% = 0

fehler% = 0
On Error GoTo ErrorISDNRueckMeldungen
rFile% = FreeFile
Open RecvFile$ For Input As #rFile%
On Error GoTo 0

If (fehler% = 0) Then
  Line Input #rFile%, recv$
  While Not (ende%)
    SATZLEN% = Val(Mid$(recv$, 3, 3))
    satz1$ = Left$(recv$, SATZLEN%)

    MaxSendSatz% = MaxSendSatz% + 1
    ReDim Preserve SendSatz$(MaxSendSatz%)
    SendSatz$(MaxSendSatz%) = satz1$
    
'    If (MaxSendSatz% < 700) Then
'        MaxSendSatz% = MaxSendSatz% + 1
'        SendSatz$(MaxSendSatz%) = satz1$
'    End If

    Oldrecv$ = recv$      '1.84 bei Lintech-ISDN-Server werden Rückmelde-
                          'sätze nur durch CHR(10) getrennt --> möglicherweise
                          'hängen Rückmeldungen aneinander
    recv$ = ""
    If (Len(Oldrecv$) > (SATZLEN% + 5)) Then
        Oldrecv$ = Mid$(Oldrecv$, SATZLEN% + 1)
        i% = InStr(Oldrecv$, Chr$(10))    'prüfen, ob weiterer Satz dranhängt
        If (i% > 0) Then Oldrecv$ = Mid$(Oldrecv$, i% + 1)
        If (Len(Oldrecv$) >= Val(Mid$(Oldrecv$, 3, 3))) Then recv$ = Oldrecv$
    End If
    If (recv$ = "") Then
        If (EOF(rFile%)) Then
            ende% = True
            recv$ = ""
        Else
            Line Input #rFile%, recv$
        End If
    End If
  Wend
End If
Close #rFile%


If (SendeLog%) Then
    For j% = 1 To MaxSendSatz%
      Print #LOGBUCH%, "<" + SendSatz$(j%)
    Next j%
End If


Call DefErrPop: Exit Sub

ErrorISDNRueckMeldungen:
    fehler% = Err
    Err = 0
    Resume Next
    Return

Call DefErrPop
End Sub

Function SeriellRueckMeldungen%()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("SeriellRueckMeldungen%")
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
Dim j%, ret%, ende%, telok%, SATZLEN%, KeinSatz%
Dim satz$, satz1$, status$, SendStr$, RecStr$

SeriellRueckMeldungen% = False

status$ = "Empfang der Rückmeldungen.": Call StatusZeile(status$)

MaxSendSatz% = 0
ReDim SendSatz$(MaxSendSatz%)
SendSatz$(0) = Space$(14) + "DATENÜBERTRAGUNG RÜCKMELDUNGEN: "

ende% = 0
KeinSatz% = True

While Not (ende%)
    SendStr$ = empfange$
    ret% = SeriellSend%(SendStr$, RecStr$)
    If (ret% = False) Then
        If (KeinSatz%) Then
            Call DefErrPop: Exit Function
        Else
            ende% = True
        End If
    End If
    
    If (ret%) Then
        KeinSatz% = False
        
        telok% = (Mid$(RecStr$, 9, 2) = "30")
        satz$ = Mid$(RecStr$, 15)
        satz$ = Left$(satz$, Len(satz$) - 1)
        
        If (telok% = 0) Then Call DefErrPop: Exit Function
        
        DoEvents
        If (BestSendenAbbruch% = True) Then Call DefErrPop: Exit Function
        
        While (satz$ <> "")
            SATZLEN% = Val(Mid$(satz$, 3, 3))
            satz1$ = Left$(satz$, SATZLEN%)
            satz$ = Mid$(satz$, SATZLEN% + 1)
            If (Left$(satz$, 1) = Chr$(10)) Then satz$ = Mid$(satz$, 2)
            If (Mid$(satz1$, 1, 2) = "99") Or (Mid$(RecStr$, 9, 1) <> "3") Then
                ende% = True
            Else
                MaxSendSatz% = MaxSendSatz% + 1
                ReDim Preserve SendSatz$(MaxSendSatz%)
                SendSatz$(MaxSendSatz%) = satz1$
            End If
        Wend
    End If
Wend

If (SendeLog%) Then
    For j% = 1 To MaxSendSatz%
      Print #LOGBUCH%, "<" + SendSatz$(j%)
    Next j%
End If
    

SeriellRueckMeldungen% = True

Call DefErrPop
End Function

Sub AbsagenEntfernen()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("AbsagenEntfernen")
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
Dim BABSAGE%, i%, d%, satz1%, SatzTyp%, SN%, babMax%, Max%, rInd%, upd%, ind%, TuEs%, st%, mn%
Dim OrgMenge%, FehlMenge%, KriegtMenge%, SollMenge%
Dim h$, x$, pzn$, xc$, DefektGrund$, bZeit$

st% = Val(Left$(Time$, 2))
mn% = Val(Mid$(Time$, 4, 2)) + 1
If (mn% > 59) Then
    mn% = 0
    st% = st% + 1
    If (st% > 23) Then
        st% = 0
    End If
End If
bZeit$ = MKI(st% * 100 + mn%)

absagen.OpenDatei
absagen.GetRecord (1)
babMax% = absagen.erstmax

ww.SatzLock (1)
ww.GetRecord (1)
Max% = ww.erstmax

upd% = False
For satz1% = 1 To MaxSendSatz%
    x$ = SendSatz$(satz1%)
    If (para.Land = "D") Then
        x$ = Left$(x$, Len(x$) - 1)
    End If
    SatzTyp% = Val(Mid$(x$, 1, 2))
    
    TuEs% = False
    If (para.Land = "A") Then
        If (SatzTyp% = 3) Then
            TuEs% = True
            OrgMenge% = Val(Mid$(x$, 5, 4))
            pzn$ = Mid$(x$, 50, 7)
            FehlMenge% = Val(Mid$(x$, 9, 4))
            DefektGrund$ = Mid$(x$, 13, 2)
    
            If (Right$(AbsagenMitNL$, 1) <> ",") Then
                AbsagenMitNL$ = AbsagenMitNL$ + ","
            End If
            
            If (InStr(AbsagenMitNL$, DefektGrund$ + ",") > 0) Then
                
                For i% = 0 To (AnzBestellArtikel% - 1)
                    If (pzn$ = mnr$(i%)) Then
                        SN% = Val(Right$(SRT$(i%), 4))
                        If (SN% > 0) Then                 '1.88
                            rInd% = SucheDateiZeile(SN%, Max%, mlauf&(i%))
                            If (rInd%) And (ww.pzn = pzn$) And (ww.zugeordnet = "J") Then
                                ww.status = 2
                                ww.Lief = Lieferant%
                                
                                If (ww.nnart = 2) Then
                                    ww.WuAEP = ww.NNAEP
                                Else
                                    ww.WuAEP = ww.aep
                                End If
                                
                                ww.WuBestDatum = HeuteDatStr$
                                ww.WuBm = Abs(ww.bm)
                                ww.WuNm = Abs(ww.nm)
                                
                                ww.WuStat = "J"
                                If (ww.bm < 0) Or (ww.nm < 0) Then ww.WuStat = "N"
                                
                                ww.WuRm = 0
                                ww.WuLm = 0
                                ww.WuAm = 0
                                ww.WuAblDatum = Space$(6)
                                ww.WuLa = 0 'MM%
                                ww.WuBestZeit = bZeit$
                                ww.WuAVP = ww.avp
                                ww.WuBelegDatum = 0
                                ww.WuBeleg = Space$(10)
                                ww.WuRetMenge = 0
                                ww.WuNNaepOk = 0
                                ww.WuNNart = ww.nnart
                                ww.WuNNAep = ww.NNAEP
                                ww.WuText = Space$(Len(ww.WuText))
                                ww.WuNeuLm = Abs(ww.bm) + Abs(ww.nm)
                                ww.WuNeuRm = Abs(ww.bm)
                                ww.WuNeuZiel = 0
                                
                                ww.IstAltLast = 1
                                ww.WuStatus = 0
                                ww.LmStatus = 0
                                ww.RmStatus = 0
                                ww.LmAnzGebucht = 0
                                ww.RmAnzGebucht = 0
                                
                                ww.aktivlief = 0
                                ww.aktivind = 0
                
                                ww.PutRecord (rInd% + 1)
                                SRT$(i%) = String$(Len(SRT$(i%)), 48)
                                
                            End If
                        End If
                        
                        Exit For
                    End If
                Next i%
                
                TuEs% = 0
            End If
        End If
    Else
        If (SatzTyp% = 4) Then
            TuEs% = True
            OrgMenge% = Val(Mid$(x$, 18, 4))
            pzn$ = Mid$(x$, 22, 7)
            FehlMenge% = Val(Mid$(x$, 31, 4))
        End If
    End If
    
    If (TuEs%) Then
        KriegtMenge% = OrgMenge% - FehlMenge%
        
        With frmAction!lstSortierung
            .Clear
            For i% = 0 To (AnzBestellArtikel% - 1)
                If (pzn$ = mnr$(i%)) Then
                    SN% = Val(Right$(SRT$(i%), 4))
                    If (SN% > 0) Then                 '1.88
                        rInd% = SucheDateiZeile(SN%, Max%, mlauf&(i%))
                        If (rInd%) And (ww.pzn = pzn$) And (ww.zugeordnet = "J") Then
                            h$ = " "
                            If (ww.auto = "+") Then
                                h$ = "Z"
                            End If
                            h$ = h$ + Format(i%, "0000")
                            .AddItem h$
                        End If
                    End If
                End If
            Next i%
            
            For i% = 1 To (.ListCount)
                upd% = True
                
                .ListIndex = i% - 1
                h$ = RTrim$(.text)
                ind% = Val(Right$(h$, 4))
                SN% = Val(Right$(SRT$(ind%), 4))
                If (SN% > 0) Then                 '1.88
                    rInd% = SucheDateiZeile(SN%, Max%, mlauf&(ind%))
                    If (rInd%) And (ww.pzn = pzn$) And (ww.zugeordnet = "J") Then
                        SollMenge% = Abs(ww.bm)
                        If (KriegtMenge% = 0) Then
                            If (Left$(h$, 1) = " ") Then AbsagenWav% = True
                            
                            ww.aktivlief = 0
                            ww.best = " "
                            ww.zugeordnet = "N"
                            ww.zukontrollieren = Chr$(0)
                            ww.fixiert = Chr$(0)
                            ww.absage = ww.absage + 1 'Lieferant%
                            If (ww.absage = 1) Then
                                ww.Lief = 0
                            End If
                            ww.IstSchwellArtikel = 0
                            ww.PutRecord (rInd% + 1)
                            SRT$(ind%) = String$(Len(SRT$(ind%)), 48)
                            
                            ReDim Preserve AbsagenZeilen$(AnzAbsagenZeilen%)
                            AbsagenZeilen$(AnzAbsagenZeilen%) = DruckZeile$(True)
                            AnzAbsagenZeilen% = AnzAbsagenZeilen% + 1
                
                        ElseIf (SollMenge% > KriegtMenge%) Then
                            If (Left$(h$, 1) = " ") Then AbsagenWav% = True
                            
                            SollMenge% = SollMenge% - KriegtMenge%
                            
                            If (ww.bm < 0) Then KriegtMenge% = -KriegtMenge%
                            ww.bm = KriegtMenge%
                            ww.PutRecord (rInd% + 1)
                            
                            If (Left$(h$, 1) = " ") Or (TeilDefekte%) Then
                                If (ww.bm < 0) Then SollMenge% = -SollMenge%
                                ww.bm = SollMenge%
                                ww.aktivlief = 0
                                ww.best = " "
                                ww.zugeordnet = "N"
                                ww.zukontrollieren = Chr$(0)
                                ww.fixiert = Chr$(0)
                                ww.absage = ww.absage + 1 'Lieferant%
                                If (ww.absage = 1) Then
                                    ww.Lief = 0
                                End If
                                ww.IstSchwellArtikel = 0
                                
                                Max% = Max% + 1
                                ww.PutRecord (Max% + 1)
                            End If
                            
                            ReDim Preserve AbsagenZeilen$(AnzAbsagenZeilen%)
                            AbsagenZeilen$(AnzAbsagenZeilen%) = DruckZeile$(True)
                            AnzAbsagenZeilen% = AnzAbsagenZeilen% + 1
                
                            KriegtMenge% = 0
                        Else
                            KriegtMenge% = KriegtMenge% - SollMenge%
                        End If
                    End If
                End If
            Next i%
            
            If (.ListCount > 0) Then
                'Absage-Datei ------------------------------
                absagen.datum = xcBestDatum%
                absagen.Lief = Lieferant%
                absagen.pzn = ww.pzn
                absagen.filler = " "
                absagen.text = Left$(ww.txt, 33) + " " + Mid$(ww.txt, 34, 2)
                absagen.menge = FehlMenge%  'ww.bm
                absagen.rest = Trim$(Mid$(x$, 70, 15)) + String$(absagen.DateiLen, 0)
                
                babMax% = babMax% + 1
                If (babMax% > 32000) Then babMax% = 1   'früher:2000
                absagen.PutRecord (babMax% + 1)
                
                absagen.GetRecord (1)
                absagen.erstmax = babMax%
                absagen.PutRecord (1)
            End If
        End With
    End If
Next satz1%

absagen.CloseDatei

If (upd%) Then
    ww.erstmax = Max%
    ww.erstcounter = (ww.erstcounter + 1) Mod 100
    ww.PutRecord (1)
End If
ww.SatzUnLock (1)

If (AbsagenWav%) And (Dir("wwbesorg.wav") <> "") Then
    Call PlaySound("\user\wwbesorg.wav", 0, SND_FILENAME Or SND_ASYNC)
End If
    
Call DefErrPop
End Sub

'Sub AbsagenEntfernen()
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Call DefErrFnc("AbsagenEntfernen")
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
'Dim BABSAGE%, i%, d%, satz1%, SatzTyp%, menge%, SN%, babMax%, Max%, rInd%, upd%
'Dim h$, X$, pzn$, xc$
'
'absagen.OpenDatei
'absagen.GetRecord (1)
'babMax% = absagen.erstmax
'
'ww.SatzLock (1)
'ww.GetRecord (1)
'Max% = ww.erstmax
'
'upd% = False
'For satz1% = 1 To MaxSendSatz%
'    X$ = SendSatz$(satz1%)
'    X$ = Left$(X$, Len(X$) - 1)
'    SatzTyp% = Val(Mid$(X$, 1, 2))
'    pzn$ = Mid$(X$, 22, 7)
'    menge% = Val(Mid$(X$, 31, 4))
'    If (SatzTyp% = 4) Then
'        For i% = 0 To (AnzBestellArtikel% - 1)
'            SN% = Val(Right$(SRT$(i%), 4))
'            If (SN% > 0) Then                 '1.88
''                bek.GetRecord (SN% + 1)
'                rInd% = SucheDateiZeile(SN%, Max%, mlauf&(i%))
'                If (rInd%) And (ww.pzn = pzn$) And (menge% = Abs(ww.bm)) And (ww.zugeordnet = "J") Then
'
'                    AbsagenZeilen$(AnzAbsagenZeilen%) = DruckZeile$(True)
'                    AnzAbsagenZeilen% = AnzAbsagenZeilen% + 1
'
'                    upd% = True
'
'                    'Absage-Datei ------------------------------
'                    absagen.datum = xcBestDatum%
'                    absagen.lief = Lieferant%
'                    absagen.pzn = ww.pzn
'                    absagen.filler = " "
'                    absagen.text = Left$(ww.txt, 33) + " " + Mid$(ww.txt, 34, 2)
'                    absagen.menge = ww.bm
'                    absagen.rest = Trim$(Mid$(X$, 70, 15)) + String$(absagen.DateiLen, 0)
'
'                    babMax% = babMax% + 1
'                    If (babMax% > 2000) Then babMax% = 1
'                    absagen.PutRecord (babMax% + 1)
'
'                    absagen.GetRecord (1)
'                    absagen.erstmax = babMax%
'                    absagen.PutRecord (1)
'
'                    'Bestell-Datei -------------------------------------------
'                    ww.lief = 0
'                    ww.aktivlief = 0
'                    ww.best = " "
'                    ww.zugeordnet = "N"
'                    ww.zukontrollieren = Chr$(0)
'                    ww.absage = Lieferant%
''                    bek.PutRecord (SN% + 1)
'                    ww.PutRecord (rInd% + 1)
'                    SRT$(i%) = String$(Len(SRT$(i%)), 48)
'                    Exit For
'                End If
'            End If
'        Next i%
'    End If
'Next satz1%
'
'absagen.CloseDatei
'
'If (upd%) Then
'    ww.erstcounter = (ww.erstcounter + 1) Mod 100
'    ww.PutRecord (1)
'End If
'ww.SatzUnLock (1)
'
'Call DefErrPop
'End Sub

Sub StatusZeile(h$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("StatusZeile")
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
Dim h2$, h3$

If (para.Land = "A") Then
    h2$ = ""
    For i% = 1 To Len(h$)
        h3$ = Mid$(h$, i%, 1)
        If (Asc(h3$) < 32) Then
            h2$ = h2$ + "<" + Format(Asc(h3$), "0") + ">"
        Else
            h2$ = h2$ + h3$
        End If
    Next i%
    h$ = h2$
Else
    h2$ = Right$(h$, 1)
    If (Asc(h2$) < 32) Then
        h$ = Left$(h$, Len(h$) - 1) + "<" + Mid$(Str$(Asc(h2$)), 2) + ">"
    End If
End If

With SendeForm.flxAuftrag
    .AddItem h$
    If (.Rows > 4) Then
        .TopRow = .Rows - 4
        If (.FixedRows > 0) Then
            .TopRow = .TopRow + .FixedRows
        End If
    End If
End With

If (SendeLog%) Then
    h2$ = h$
    Call CharToOem(h2$, h2$)
    Print #LEITUNGBUCH%, Format(Now, "HHMMSS ") + h2$
End If

DoEvents
Call DefErrPop
End Sub

Sub TestBestellung()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("TestBestellung")
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
Dim wPos&
Dim satz$
Dim h$

h$ = LINTECHPFAD$ + "SAVE\S_" + Right$(String$(5, 48) + Mid$(Str$(Lieferant%), 2), 5) + "1" + ".I01"
RecvFile$ = LINTECHPFAD$ + "S_" + Right$(String$(5, 48) + Mid$(Str$(Lieferant%), 2), 5) + "1" + ".RUK"
FileCopy h$, RecvFile$

wPos& = -1
Call iLock(ZJOB%, 1)
Seek ZJOB%, 64
satz$ = Space$(142)
Do While Not (EOF(ZJOB%))
  Get ZJOB%, , satz$
  If (Left$(satz$, 7) = GhIDF$) Then
    wPos& = Seek(ZJOB%) - 142
    Exit Do
  End If
Loop
If (wPos& < 0) Then Call DefErrPop: Exit Sub
Mid$(satz$, 28, 1) = "1"
satz$ = Left$(satz$ + Space$(141), 141) + Chr$(10)
Put ZJOB%, wPos&, satz$
Call iUnLock(ZJOB%, 1)

Call DefErrPop
End Sub

Sub UpdateBekartDat(Lief%, Ueberleiten%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("UpdateBekartDat")
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
Dim i%, j%, SN%, bm%, nm%, MM%, ppMax%, ppGes%, bMax%, pp%, ssatz%, Max%, ok%, rInd%, TransNr%
Dim EK#
Dim bZeit$, shdatum$, shzeit$, shlieferant$, pzn$, OrgPzn$, text$, tx$, ShuttleHeader$, s$, sTrans$, lKurz$, ean$
Dim k&

TransNr% = 0

bZeit$ = MKI(Val(Left$(Time$, 2)) * 100 + Val(Mid$(Time$, 4, 2)))

ShuttleHeader$ = Format(Now, "yy.mm.dd") + " " + Left$(Time$, 2) + ":" + Mid$(Time$, 4, 2)
ShuttleHeader$ = ShuttleHeader$ + " " + Right$("   " + Str$(Lief%), 3)
Call ShuttleProtokoll(ShuttleHeader$)
            
ww.SatzLock (1)
ww.GetRecord (1)
Max% = ww.erstmax

If (Ueberleiten%) Then
    If (MacheEtiketten%) Then
        Call etik.SatzLock(1)
        Call etik.GetRecord(1)
        ppMax% = etik.erstmax
        ppGes% = etik.erstges
    End If
    
    DRUCKBUCH% = FileOpen("winw\" + SendDruckDateiName$ + ".sp9", "O")
    For i% = 0 To (AnzAbsagenZeilen% - 1)
        Print #DRUCKBUCH%, AbsagenZeilen$(i%)
    Next i%
    AnzAbsagenZeilen% = 0
    
    For i% = 0 To (AnzBestellArtikel% - 1)
        SN% = Val(Right$(SRT$(i%), 4))
        If (SN% > 0) Then                 '1.88
'            bek.GetRecord (SN% + 1)
            rInd% = SucheDateiZeile(SN%, Max%, mlauf&(i%))
            If (rInd%) And (ww.pzn = mnr$(i%)) And (ww.aktivlief = Lief%) Then
                
                text$ = DruckZeile$
                Print #DRUCKBUCH%, text$
                
                bm% = ww.bm: nm% = ww.nm: pzn$ = ww.pzn
                If Left$(pzn$, 4) = "9999" And Mid$(pzn$, 5, 3) <> "999" Then
                    ww.status = 0
                Else
                    ssatz% = 0
                    OrgPzn$ = pzn$
                    FabsErrf% = ass.IndexSearch(0, pzn$, FabsRecno&)
                    If (FabsErrf% = 0) Then
                        ssatz% = FabsRecno&
                        ass.GetRecord (FabsRecno& + 1)
                        MM% = ass.MM
                    Else
                        MM% = 0
                    End If
                    
                    tx$ = pzn$ + Format(ssatz%, " ####0") + Format(bm%, " ###0")
                    If (ShuttleAktiv%) And (ssatz% > 0) And (bm% <> 0) Then
                        tx$ = tx$ + Format(ass.tplatz, " 0")
                        If (ass.tplatz = 0) Then
                            FabsErrf% = ast.IndexSearch(0, pzn$, FabsRecno&)
                            If (FabsErrf% = 0) Then
                                ast.GetRecord (FabsRecno& + 1)
                                tx$ = tx$ + " " + ast.Lac
                                If (ast.Lac = "!") Then
                                
                                    If (TransNr% = 0) Then

                                        SHAREDDAT% = FileOpen%("\USER\SHARED.DAT", "RW", "R", Len(SharedDatSatz))
                                        If (LOF(SHAREDDAT%) = 0) Then
                                            Call ShuttleProtokoll("SharedDat 0")
                                            Close #SHAREDDAT%
                                            Kill "\USER\SHARED.DAT"
                                            TransNr% = -1
                                        Else
                                            TransNr% = ShuttleHeaderNr%(ShuttleHeader$)
                                            Call ShuttleProtokoll("ShuttleHeaderNr:" + Format(TransNr%, " ###0"))
                                        End If

                                        If (TransNr% > 0) Then
                                            'Header für Com-Pc (Automaten-Artikel)
                                            sTrans$ = Right$(Str$(TransNr%), 3)
                        
                                            lif.GetRecord (Lief% + 1)
                                            lKurz$ = lif.kurz
                                            
                                            s$ = "01" + sTrans$ + "0000000  "
                                            s$ = s$ + "Lieferant " + lKurz$ + " (" + Mid$(Str$(Lief%), 2) + ")"
                                            Call ShuttleArtikel(s$)
                                        End If
                                    End If
                                
                                    If (TransNr% > 0) Then
                                        For j% = 1 To Abs(bm%)
                                            s$ = "01" + sTrans$ + pzn$ + "02"
                                            '1.62
                                            ean$ = ass.ean
                                            Call ConvertEan(ean$)
                                            s$ = s$ + ean$
                                            Call ShuttleArtikel(s$)
                                        Next j%
                                    End If
                                
                                End If
                            End If
                        End If
                    End If
                    Call ShuttleProtokoll(tx$)

                    text$ = pzn$ + " " + Left$(ww.txt, 33) + " " + Mid$(ww.txt, 34, 2)
                    If (bm% = 0) And (nm% = 0) Then     '1.60 Merkzettel
                        Call MerkZettel(pzn$, Mid$(text$, 9), xcBestDatum%, xcBestDatum%, Lieferant%, 0)
                        ww.status = 0
                    Else
                        ww.status = 2
                        ww.Lief = Lief%
                        
                        If (ww.nnart = 2) Then
                            ww.WuAEP = ww.NNAEP
                        Else
                            ww.WuAEP = ww.aep
                        End If
                        
                        
                        If (IstDirektLief%) And (ww.zr < 0) Then
                            If (ww.nm = 0) Or (DirektBezugFaktRabattTyp% <> 1) Then
                                ww.WuAEP = ww.WuAEP * (100# - DirektBezugFaktRabatt#) / 100#
                            End If
                        End If
                
                        
                        
                        ww.WuBestDatum = HeuteDatStr$
                        ww.WuBm = Abs(bm%)
                        ww.WuNm = Abs(nm%)
                        
                        ww.WuStat = "J"
                        If (bm% < 0) Or (nm% < 0) Then ww.WuStat = "N"
                        
                        ww.WuRm = Abs(bm%)
                        ww.WuLm = Abs(bm%) + Abs(nm%)
                        ww.WuAm = 0
                        ww.WuAblDatum = Space$(6)
                        ww.WuLa = MM%
                        ww.WuBestZeit = bZeit$
                        ww.WuAVP = ww.avp
                        ww.WuBelegDatum = 0
                        ww.WuBeleg = Space$(10)
                        ww.WuRetMenge = 0
'                            ww.WuAnzBereitsGebucht = 0
'                            ww.WuFertig = 0
                        ww.WuNNaepOk = 0
                        ww.WuNNart = ww.nnart
                        ww.WuNNAep = ww.NNAEP
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
                    End If
                  
                    If (MacheEtiketten%) Then
                        pp% = 0
                        If (ssatz% <> 0) Then If (InStr("AFP", ass.pp) <> 0) Then pp% = True
                        If (InStr(para.PosAktivWG, ww.wg) <> 0) Then pp% = True
                        If (pp%) And (OrgPzn$ <> "9999999") Then
                            If (ww.auto <> "v") Then
                                etik.pzn = OrgPzn$
                                etik.menge = Abs(bm%) + Abs(nm%)
                                etik.druck = " "
                                etik.zusatz = "  "
                                etik.asatz = ww.asatz
                                etik.Lief = Lieferant%
                                etik.datum = xcBestDatum%
                                etik.Knr = String(2, 0)
                                etik.rest = String(44, 0)
                                ppMax% = ppMax% + 1
                                etik.PutRecord (ppMax% + 1)
                                ppGes% = ppGes% + Abs(bm%) + Abs(nm%)
                            End If
                        End If
                    End If
                End If
                ww.PutRecord (rInd% + 1)
            End If
        End If
    Next i%
    
    If (TransNr% > 0) Then Close #SHAREDDAT%
    
    For i% = 1 To MaxSendSatz%
        text$ = SendSatz$(i%)
        If (para.Land = "D") Then
            text$ = Left$(text$, Len(text$) - 1)
        End If
        Print #DRUCKBUCH%, "RM: " + text$
    Next i%
    Close #DRUCKBUCH%
                        
    If (MacheEtiketten%) Then
        Call etik.GetRecord(1)
        etik.erstmax = ppMax%
        etik.erstges = ppGes%
        Call etik.PutRecord(1)
        Call etik.SatzUnLock(1)
    End If
    
'    j% = 0
'    For i% = 1 To Max%
'        ww.GetRecord (i% + 1)
'        If (ww.Status > 0) Then
'            j% = j% + 1
'            ww.PutRecord (j% + 1)
'        End If
'    Next i%
'    ww.erstmax = j%
'    ww.PutRecord (1)
    Call EntferneGeloeschteZeilen(0, True)
Else
    For i% = 1 To Max%
        ww.GetRecord (i% + 1)
        
        ok% = False
        If (Lief% = -1) Or (ww.aktivlief = Lief%) Then ok% = True
        If (ww.status <> 1) Then ok% = False
        If (Lief% = -1) And (ww.aktivind < 0) Then ok% = False
        
        If (ok%) Then
            ww.aktivlief = 0
            ww.aktivind = 0
            ww.PutRecord (i% + 1)
        End If
    Next i%
End If


ww.erstcounter = (ww.erstcounter + 1) Mod 100
ww.PutRecord (1)
ww.SatzUnLock (1)

Call DefErrPop
End Sub

Sub UpdateEinzelZeile(LifDat$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("UpdateEinzelZeile")
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
Dim i%, j%, SN%, bm%, nm%, MM%, ppMax%, ppGes%, bMax%, pp%, ssatz%, Max%, ok%, rInd%, Lief%
Dim xcBdatum%
Dim EK#
Dim bZeit$, shdatum$, shzeit$, shlieferant$, pzn$, OrgPzn$, text$, tx$, ShuttleHeader$, s$, sTrans$, lKurz$, ean$
Dim bdatum$
Dim k&

Lief% = Asc(Left$(LifDat$, 1))
bdatum$ = Mid$(LifDat$, 2, 6)
xcBdatum% = iDate(bdatum$)
bZeit$ = MKI(Val(Mid$(LifDat$, 8, 4)))

ww.SatzLock (1)
ww.GetRecord (1)
Max% = ww.erstmax

If (MacheEtiketten%) Then
    Call etik.SatzLock(1)
    Call etik.GetRecord(1)
    ppMax% = etik.erstmax
    ppGes% = etik.erstges
End If

rInd% = SucheFlexZeile(True)
If (rInd%) Then
    
    bm% = ww.bm: nm% = ww.nm: pzn$ = ww.pzn
    If Left$(pzn$, 4) = "9999" And Mid$(pzn$, 5, 3) <> "999" Then
        ww.status = 0
    Else
        ssatz% = 0
        OrgPzn$ = pzn$
        FabsErrf% = ass.IndexSearch(0, pzn$, FabsRecno&)
        If (FabsErrf% = 0) Then
            ssatz% = FabsRecno&
            ass.GetRecord (FabsRecno& + 1)
            MM% = ass.MM
        Else
            MM% = 0
        End If
        
        text$ = pzn$ + " " + Left$(ww.txt, 33) + " " + Mid$(ww.txt, 34, 2)
        If (bm% = 0) And (nm% = 0) Then     '1.60 Merkzettel
            Call MerkZettel(pzn$, Mid$(text$, 9), xcBdatum%, xcBdatum%, Lieferant%, 0)
            ww.status = 0
        Else
            ww.status = 2
            ww.Lief = Lief%
            
            If (ww.nnart = 2) Then
                ww.WuAEP = ww.NNAEP
            Else
                ww.WuAEP = ww.aep
            End If
            
            ww.WuBestDatum = bdatum$
            ww.WuBm = Abs(bm%)
            ww.WuNm = Abs(nm%)
            
            ww.WuStat = "J"
            If (bm% < 0) Or (nm% < 0) Then ww.WuStat = "N"
            
            ww.WuRm = Abs(bm%)
            ww.WuLm = Abs(bm%) + Abs(nm%)
            ww.WuAm = 0
            ww.WuAblDatum = Space$(6)
            ww.WuLa = MM%
            ww.WuBestZeit = bZeit$
            ww.WuAVP = ww.avp
            ww.WuBelegDatum = 0
            ww.WuBeleg = Space$(10)
            ww.WuRetMenge = 0
'                            ww.WuAnzBereitsGebucht = 0
'                            ww.WuFertig = 0
            ww.WuNNaepOk = 0
            ww.WuNNart = ww.nnart
            ww.WuNNAep = ww.NNAEP
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
        End If
      
        If (MacheEtiketten%) Then
            pp% = 0
            If (ssatz% <> 0) Then If (InStr("AFP", ass.pp) <> 0) Then pp% = True
            If (InStr(para.PosAktivWG, ww.wg) <> 0) Then pp% = True
            If (pp%) And (OrgPzn$ <> "9999999") Then
                If (ww.auto <> "v") Then
                    etik.pzn = OrgPzn$
                    etik.menge = Abs(bm%) + Abs(nm%)
                    etik.druck = " "
                    etik.zusatz = "  "
                    etik.asatz = ww.asatz
                    etik.Lief = Lieferant%
                    etik.datum = xcBdatum%
                    etik.Knr = String(2, 0)
                    etik.rest = String(44, 0)
                    ppMax% = ppMax% + 1
                    etik.PutRecord (ppMax% + 1)
                    ppGes% = ppGes% + Abs(bm%) + Abs(nm%)
                End If
            End If
        End If
    End If
    ww.PutRecord (rInd% + 1)
End If

If (MacheEtiketten%) Then
    Call etik.GetRecord(1)
    etik.erstmax = ppMax%
    etik.erstges = ppGes%
    Call etik.PutRecord(1)
    Call etik.SatzUnLock(1)
End If

'Call EntferneGeloeschteZeilen(0, True)


ww.erstcounter = (ww.erstcounter + 1) Mod 100
ww.PutRecord (1)
ww.SatzUnLock (1)

Call DefErrPop
End Sub

Function DruckZeile$(Optional IstAbsage% = False)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("DruckZeile$")
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
Dim EK#
Dim text$, tx$
Dim DatenFile As Object

If (RueckKaufSendung%) Then
    Set DatenFile = rk
Else
    Set DatenFile = ww
End If

EK# = DatenFile.aep
tx$ = Left$(DatenFile.txt, 28)
If (IstAbsage%) Then
    EK# = 0#
    tx$ = Left$("Absage: " + DatenFile.txt, 28)
End If

text$ = tx$ + vbTab + Mid$(DatenFile.txt, 29, 5)
text$ = text$ + vbTab + Mid$(DatenFile.txt, 34, 2)

tx$ = Format(Abs(DatenFile.bm), "0")
text$ = text$ + vbTab + tx$

tx$ = ""
If (Abs(DatenFile.nm > 0)) Then
    tx$ = Format(Abs(DatenFile.nm), "0")
End If
text$ = text$ + vbTab + tx$

tx$ = Format(EK# * Abs(DatenFile.bm), "0.00")
text$ = text$ + vbTab + tx$
text$ = text$ + vbTab + DatenFile.pzn + vbTab

'EK# = ww.aep
'tx$ = Left$(ww.txt, 28)
'If (IstAbsage%) Then
'    EK# = 0#
'    tx$ = Left$("Absage: " + ww.txt, 28)
'End If
'
'text$ = tx$ + vbTab + Mid$(ww.txt, 29, 5)
'text$ = text$ + vbTab + Mid$(ww.txt, 34, 2)
'
'tx$ = Format(Abs(ww.bm), "0")
'text$ = text$ + vbTab + tx$
'
'tx$ = ""
'If (Abs(ww.nm > 0)) Then
'    tx$ = Format(Abs(ww.nm), "0")
'End If
'text$ = text$ + vbTab + tx$
'
'tx$ = Format(EK# * Abs(ww.bm), "0.00")
'text$ = text$ + vbTab + tx$
'text$ = text$ + vbTab + ww.pzn + vbTab

DruckZeile$ = text$

Call DefErrPop
End Function

Sub MerkZettel(pzn$, text$, dat%, dt%, Lief%, Lm%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("Merkzettel")
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
Dim i%, STAMMLOS%, Version%
Dim MzMax&
Dim satz!, xrec!, xrec1!
Dim h$, feld$, match$, Cmnd$, ch$, SQLStr$


'besorgt.GetRecord (1)
'Version% = Asc(Mid$(StammlosFix$, 4, 1))

SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + pzn$
Set TaxeRec = TaxeDB.OpenRecordset(SQLStr$)
If (TaxeRec.EOF = False) Then text$ = UCase(text$)

BESORGT.pzn = pzn$
BESORGT.flag = " "
BESORGT.text = text$
BESORGT.f1 = " "
BESORGT.dat = dat%
BESORGT.Leer = String$(4, 0)
BESORGT.f2 = " "
BESORGT.dt = dt%
BESORGT.leer1 = String$(4, 0)
BESORGT.Lief = Lief%
BESORGT.Lm = Lm%
BESORGT.rest = Chr$(13) + Chr$(10)

If (para.Land = "A") Then
    BESORGT.SatzLock (1)
    BESORGT.GetRecord (1)
    MzMax& = FNSATZ&(BESORGT.erstmax)
    
    MzMax& = MzMax& + 1
    If (MzMax& > 65535) Then MzMax& = 65535
    BESORGT.PutRecord (MzMax& + 1)
    
    BESORGT.erstmax = fnint%(MzMax&)
    BESORGT.erstdruck = "*"
    BESORGT.PutRecord (1)
    BESORGT.SatzUnLock (1)
Else
    match$ = Left$(BESORGT.text, 6) + BESORGT.pzn + MKDate(BESORGT.dat)
      
    'CHR$(0) und \ ersetzen durch Space
    For i% = 1 To Len(match$)
        ch$ = Mid$(match$, i%, 1)
        If (ch$ = Chr$(0)) Or (ch$ = "\") Then Mid$(match$, i%, 1) = " "
    Next i%
    
    FabsErrf% = BESORGT.IndexInsert(0, BESORGT.pzn, FabsRecno&)
    xrec! = FabsRecno&
    
    FabsErrf% = BESORGT.IndexInsert(1, match$, FabsRecno&)
    xrec1! = FabsRecno&
    
    If (xrec! <> xrec1!) Then
        Call iMsgBox("Unterschiedliche Records im Merkzettel!", vbOKOnly Or vbCritical, "FABSP")
    End If
    
    satz! = xrec!
    
    BESORGT.PutRecord (satz! + 1)
    
    BESORGT.SatzLock (1)
    BESORGT.GetRecord (1)
    'LockError! = 0
    'MzMax& = FNSATZ%(CVI%(Left$(StammlosFix$, 2)))
    MzMax& = FNSATZ&(BESORGT.erstmax)
    If (xrec! > MzMax&) Then
        MzMax& = xrec!
        If (MzMax& > 65535) Then MzMax& = 65535
        BESORGT.erstmax = fnint%(MzMax&)
    '    StammlosFix$ = MKI$(fnint%(MzMax&)) + "*" + Mid$(StammlosFix$, 4)
        BESORGT.PutRecord (1)
    End If
    BESORGT.SatzUnLock (1)
    'Close #STAMMLOS%
End If

Call DefErrPop
End Sub

'Sub Merkzettel(text$, dat$, DT$, lief$, lm%)
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Call DefErrFnc("Merkzettel")
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
'Dim STAMMLOS%, Version%
'Dim MzMax&
'Dim satz!, xrec!, xrec1!
'Dim h$, feld$, match$, Cmnd$
'
'
'h$ = "stammlos.dat"
'STAMMLOS% = FileOpen(h$, "RW", "R", Len(StammlosFix$))
'
'
''FIELD #f.mz%, 65 AS f.mz$
'
'Get #STAMMLOS%, 1, StammlosFix$
'Version% = Asc(Mid$(StammlosFix$, 4, 1))
'
'feld$ = String$(65, 32)
'Mid$(feld$, 1, 44) = text$
'If (Version% = 1) Then     'altes Format
'    Mid$(feld$, 46, 6) = dat$
'    Mid$(feld$, 53, 6) = DT$
'Else                    'neues Format
'    Mid$(feld$, 46, 2) = MKDate$(iDate%(dat$))
'    Mid$(feld$, 48, 4) = String$(4, 0)
'    Mid$(feld$, 53, 6) = MKDate$(iDate%(DT$))
'    Mid$(feld$, 55, 4) = String$(4, 0)
'End If
'
'Mid$(feld$, 59, 1) = lief$
'Mid$(feld$, 60, 4) = Right$("    " + Str$(lm%), 4)
'Mid$(feld$, 64, 2) = Chr$(13) + Chr$(10)
'
'If (Version% > 1) Then
'
'    match$ = Mid$(feld$, 9, 6) + Mid$(feld$, 1, 7) + Mid$(feld$, 46, 2)
'
'    FabsErrf% = besorgt.IndexInsert(0, Mid$(feld$, 1, 7), FabsRecno&)
'    xrec! = FabsRecno&
'
'    FabsErrf% = besorgt.IndexInsert(1, match$, FabsRecno&)
'    xrec1! = FabsRecno&
'
'    If (xrec! <> xrec1!) Then
'        Call MsgBox("Unterschiedliche Records im Merkzettel!", vbOKOnly Or vbCritical, "FABSP")
'    End If
'
'    satz! = xrec!
'
'Else
'    satz! = MzMax& + 1
'End If
'
'StammlosFix$ = feld$
'Put #STAMMLOS%, satz! + 1, StammlosFix$
'
'Call iLock(STAMMLOS%, 1)
'Get #STAMMLOS%, 1, StammlosFix$
''LockError! = 0
'MzMax& = FNSATZ%(CVI%(Left$(StammlosFix$, 2)))
'If (xrec! > MzMax&) Then
'    MzMax& = xrec!
'    If (MzMax& > 65535) Then MzMax& = 65535
'    StammlosFix$ = MKI$(fnint%(MzMax&)) + "*" + Mid$(StammlosFix$, 4)
'    Put #STAMMLOS%, 1, StammlosFix$
'End If
'Call iUnLock(STAMMLOS%, 1)
'Close #STAMMLOS%
'
'Call DefErrPop
'End Sub

Function fnint%(ByVal x!)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("fnint%")
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
fnint% = Int(x! + (x! > 32767) * 65536)
Call DefErrPop
End Function

Function FNSATZ&(x%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("FNSATZ&")
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
FNSATZ& = x% - (x% < 0) * 65536!
Call DefErrPop
End Function

Function IsdnSer%(Funkt%, ind%, sap As SerAuftrag)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("IsdnSer%")
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
End Function

Function ModemAktivieren%()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("ModemAktivieren%")
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
Dim ret%

ret% = True
If (seriell% And Not (ModemInDOS%)) Then
    ret% = OpenModemCom(SendeForm, SendPara$)
End If

ModemAktivieren% = ret%

Call DefErrPop
End Function

Function OpenModemCom%(sForm As Object, xFilePara$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("OpenModemCom%")
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
Dim i%, ret%, fehler%, CommPort%, ind%, ind2%, stByte%
Dim h$, Settings$

ret% = True

fehler% = 0
' COM einsetzen.
ind% = InStr(xFilePara$, "COM")
If (ind% > 0) Then
    CommPort% = Val(Mid$(xFilePara$, ind% + 3, 1))
    Settings$ = Mid$(xFilePara$, ind% + 5)
    
    stByte% = 1
    For i% = 1 To 4
        ind2% = InStr(stByte%, Settings$, ",")
        If (ind2% > 0) Then
            If (i% = 4) Then
                Settings$ = Left$(Settings$, ind2% - 1)
            Else
                stByte% = ind2% + 1
            End If
        Else
            Exit For
        End If
    Next i%
End If

With sForm
    .comSenden.CommPort = CommPort%
    .comSenden.Settings = Settings$
'    .comSenden.InputMode = comInputModeText
    .comSenden.InputMode = comInputModeBinary
    .comSenden.Handshaking = PharmaBoxHandShake%    ' comRTSXOnXOff
    .comSenden.InputLen = 1
    
    ' _Anschluß öffnen.
    On Error GoTo ErrorHandler
    .comSenden.PortOpen = True
End With

If (fehler%) Then ret% = False

OpenModemCom% = ret%

Call DefErrPop: Exit Function

ErrorHandler:
    fehler% = Err
    Err = 0
    Resume Next
    Return

End Function

Sub HoleLieferantenDaten(iLieferant%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("HoleLieferantenDaten")
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
Dim i%, ok%
Dim h2$, x$, BetrNr$

lif.GetRecord (iLieferant% + 1)
'h2$ = lif.Name(0): Call OemToChar(h2$, h2$): LiefName1$ = RTrim$(h2$)
'h2$ = lif.Name(1): Call OemToChar(h2$, h2$): LiefName2$ = RTrim$(h2$)
'h2$ = lif.Name(2): Call OemToChar(h2$, h2$): LiefName3$ = RTrim$(h2$)
'h2$ = lif.Name(3): Call OemToChar(h2$, h2$): LiefName4$ = RTrim$(h2$)
h2$ = lif.Name(0): LiefName1$ = RTrim$(h2$)
h2$ = lif.Name(1): LiefName2$ = RTrim$(h2$)
h2$ = lif.Name(2): LiefName3$ = RTrim$(h2$)
h2$ = lif.Name(3): LiefName4$ = RTrim$(h2$)

GhIDF$ = lif.IdfLieferant
BetrNr$ = lif.IdfApo
'GhIDF$ = Mid$(lif.rest, 92, 7)
'BetrNr$ = Mid$(lif.rest, 79, 7)

'X$ = Mid$(lif.rest, 17, 14): X$ = LTrim$(X$): X$ = RTrim$(X$): TelGh$ = X$
x$ = lif.Telefon: x$ = LTrim$(x$): x$ = RTrim$(x$): TelGh$ = x$
i% = 1
While (i% <= Len(TelGh$))
  If InStr("0123456789:<=>;.TP& ", Mid$(TelGh$, i%, 1)) = 0 Then
    TelGh$ = Mid$(TelGh$, 1, i% - 1) + Mid$(TelGh$, i% + 1)
  Else
    i% = i% + 1
  End If
Wend

'1.79 Lieferantendurchwahl - nur zulassen, wenn rein numerisch
'X$ = Mid$(lif.rest, 75, 4): X$ = LTrim$(X$): X$ = RTrim$(X$)
x$ = lif.Durchwahl: x$ = LTrim$(x$): x$ = RTrim$(x$)
ok% = True
For i% = 1 To 4
  If InStr("0123456789 ", Mid$(x$, i%, 1)) = 0 Then ok% = 0
Next i
If (ok%) Then TelGh$ = TelGh$ + x$

ApoIDF$ = BetrNr$

lifzus.GetRecord (iLieferant% + 1)

Call DefErrPop
End Sub

Function ZeigeModemTyp$()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("ZeigeModemTyp$")
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

h$ = ""
If (para.isdn) Then
    h$ = "ISDN-Karte"
    If (Lintech%) Then h$ = h$ + " (Lintech)"
Else
    h$ = "Seriellmodem (" + SendPara$ + ")"
End If

ZeigeModemTyp$ = h$

Call DefErrPop
End Function

Sub EntferneAktivKz(Lief%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("EntferneAktivKz")
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
    If (Lief% = -1) Or (ww.aktivlief = Lief%) Then
        ww.aktivlief = 0
        ww.aktivind = 0
        ww.PutRecord (i% + 1)
    End If
Next i%
ww.SatzUnLock (1)

Call DefErrPop
End Sub

Sub RueckMeldungTestDatei()
Dim f%
Dim s$

s$ = "\lintech\s_00" + Format(Lieferant%, "000") + "1.ruk"
f% = FreeFile
'Open "\lintech\s_000021.ruk" For Output As #f%
Open s$ For Output As #f%

s$ = "0408440003561    00017554032TA0001  10 N1  LYPHAN FILM               ZUR ZEIT DEFEKT"
s$ = Left$(s$ + Space$(256), 84) + Chr$(10)
Print #f%, s$;
s$ = "0408440003561    00014377397  0000   1 ROL OMNIFIX ELASTIC 10CMX 6M HNEU: 0255585"
s$ = Left$(s$ + Space$(256), 84) + Chr$(10)
Print #f%, s$;
s$ = "0408440003561    00044325667KA0003  50 N2  SAROTEN RETARD 50MG       DEFEKT HERSTELL"
s$ = Left$(s$ + Space$(256), 84) + Chr$(10)
Print #f%, s$;
s$ = "0408440003561    00026963018TA0002  50 N2  TRIAMPUR FORTE FILM       NICHT GEFUEHRT"
s$ = Left$(s$ + Space$(256), 84) + Chr$(10)
Print #f%, s$;
s$ = "082564000356IHRE SANACORP AG<BEDANKT SICH FUER DEN AUFTRAG.<   17 ARTIKEL EMPFANGEN<    3 ARTIKEL DEFEKT<   200,13 AUFTRAGSWERT<    14,29 DURCHSCHN. ZEILENWERT<TOUR 2207<07.06.00 17:22<>"
s$ = Left$(s$ + Space$(256), 256) + Chr$(10)
Print #f%, s$;
s$ = "9901930029049310274"
s$ = Left$(s$ + Space$(256), 19) + Chr$(10)
Print #f%, s$;

'Close #f%
'Exit Sub

'"         1         2         3         4         5         6         7         8"
'"12345678901234567890123456789012345678901234567890123456789012345678901234567890"
s$ = "0408430029041    00010167415  0001     1 STAUSTAUSCHSET STE 13164 HEPFEHLT ZUR ZEIT"
s$ = Left$(s$ + Space$(256), 84) + Chr$(10)
Print #f%, s$;
s$ = "062563002904     <<<          NACHLIEFERN MOEGLICH          <0167415 AUSTAUSCHARTIKEL"
s$ = Left$(s$ + Space$(256), 256) + Chr$(10)
Print #f%, s$;
s$ = "0408430029041    00020950546  0002   650 GRREDUPECT PLUS             NICHT LIEFERBAR"
s$ = Left$(s$ + Space$(256), 84) + Chr$(10)
Print #f%, s$;
s$ = "0408430029041    00027366508  0002    10 STREISEKAUGUMMI STADA       AHD SEIT       "
s$ = Left$(s$ + Space$(256), 84) + Chr$(10)
Print #f%, s$;
s$ = "0408430029041    00020224722  0002     2 STBERUHIG SGR KFR LA 6-188/PNICHT GEFUNDEN "
s$ = Left$(s$ + Space$(256), 84) + Chr$(10)
Print #f%, s$;
s$ = "062563002904     <<<          DISPO MOEGLICH: FIRMA         <BUETTNER-FRANK         "
s$ = Left$(s$ + Space$(256), 256) + Chr$(10)
Print #f%, s$;
s$ = "0408430029041    00011012548SL0001   2,5 GRTERRAMYCIN AUGEN          NICHT LIEFERBAR"
s$ = Left$(s$ + Space$(256), 84) + Chr$(10)
Print #f%, s$;
s$ = "062563002904     AUFTRAGSBESTAETIGUNG<ODERBRUCH APOTHEKE            <LIEFERUNG      "
s$ = Left$(s$ + Space$(256), 256) + Chr$(10)
Print #f%, s$;
s$ = "070353002904EPHOENIX  DEFEKTE      "
s$ = Left$(s$ + Space$(256), 35) + Chr$(10)
Print #f%, s$;
's$ = "082563002904AUFTRAGSBESTAETIGUNG<ODERBRUCH APOTHEKE            <LIEFERUNG 20:30      "
s$ = "082564000250GEHE NL BERLIN           9:25 UHR<<  6 ARTIKEL UEBERTRAGEN.<<  WIR KOENNEN IHREN AUFTRAG<  LEIDER NICHT BEARBEITEN !!<<  BITTE SETZEN SIE SICH MIT UNSERER<  AUFTRAGS-ANNAHME IN VERBINDUNG<>"
s$ = Left$(s$ + Space$(256), 256) + Chr$(10)
Print #f%, s$;
s$ = "9901930029049310274"
s$ = Left$(s$ + Space$(256), 19) + Chr$(10)
Print #f%, s$;




Close #f%

End Sub

Function ShuttleHeaderNr%(s$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("ShuttleHeaderNr%")
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
Dim i%, ret%
Dim x$

ret% = 0

AUTOHEAD% = FileOpen%("\USER\AUTOHEAD.DAT", "RW", "R", Len(AutoHeadSatz))
If (AUTOHEAD% > 0) Then

    For i% = 0 To 99
        Get #AUTOHEAD%, i% + 1, AutoHeadSatz
        If (Val(Left$(AutoHeadSatz, 3)) = 0) Then
            ret% = 800 + i%
            x$ = Right$("   " + Str$(ret%), 3) + " " + s$
            x$ = Left$(x$ + Space$(22), 22) + Chr$(13) + Chr$(10)
            AutoHeadSatz = x$
            Put #AUTOHEAD%, i% + 1, AutoHeadSatz
            Exit For
        End If
    Next i%
    
    Close #AUTOHEAD%
    
End If

ShuttleHeaderNr% = ret%

Call DefErrPop
End Function

Sub ShuttleArtikel(Datensatz$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("ShuttleArtikel")
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
Dim Max%, rc%
Dim satz1$

Call iLock(SHAREDDAT%, 1)
Get #SHAREDDAT%, 1, SharedDatSatz
satz1$ = SharedDatSatz
Max% = CVI(Left$(SharedDatSatz, 2))
rc% = Asc(Mid$(SharedDatSatz, 3, 1))
SharedDatSatz = Left$("1" + Datensatz$ + Space$(128), 128)
Max% = Max% + 1
Put #SHAREDDAT%, Max% + 1, SharedDatSatz

rc% = rc% + 1
If (rc% > 255) Then rc% = 1
SharedDatSatz = MKI(Max%) + Chr$(rc%) + Mid$(satz1$, 4)
Put #SHAREDDAT%, 1, SharedDatSatz
Call iUnLock(SHAREDDAT%, 1)

Call ShuttleProtokoll("ShuttleArtikel:" + Format(Max%, " ###0 ") + Datensatz$)

Call DefErrPop
End Sub

Sub ConvertEan(ean$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("ConvertEan")
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

ean$ = Trim$(ean$)

i% = 1
While (i% <= Len(ean$))
    If (InStr("0123456789", Mid$(ean$, i%, 1)) = 0) Then
        ean$ = Left$(ean$, i% - 1) + Mid$(ean$, i% + 1)
    Else
        i% = i% + 1
    End If
Wend
ean$ = Right$(Space$(13) + ean$, 13)

Call DefErrPop
End Sub

Sub ShuttleProtokoll(s$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("ShuttleProtokoll")
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

If (ShuttleAktiv%) Then
    Handle% = FreeFile
    Open "shbest.txt" For Append As #Handle%
    Print #Handle%, s$
    Close #Handle%
End If

Call DefErrPop
End Sub
                 
Sub DirektBezugAusdruck(Optional ManuellDruck% = True)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("DirektBezugAusdruck")
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
Dim ZentrierX%, DruckerWechsel%
Dim RetourWert#
Dim tx$, h$, AktDruckerName$

DruckerWechsel% = False
If (ProgrammChar$ = "2") Then
    If (AutomatikDrucker$ <> "") And (AutomatikDrucker$ <> Printer.DeviceName) Then
        DruckerWechsel% = True
    End If
End If

If (DruckerWechsel%) Then
    AktDruckerName$ = Printer.DeviceName
    For i% = 0 To (Printers.Count - 1)
        h$ = Printers(i%).DeviceName
        If (h$ = AutomatikDrucker$) Then
            Set Printer = Printers(i%)
            Exit For
        End If
    Next i%
End If


AktLief% = Lieferant%

With frmAction.lstDirektSortierung
    Do
    
'        If (ManuellDruck%) Then Call StartAnimation("Ausdruck wird erstellt ...")
        
        On Error Resume Next
        FaxDruckFussHoehe% = 45
        FaxDruckKopfTxt$ = ""
        Handle% = FreeFile
        Open "\user\text\retour1.txt" For Input As #Handle%
        If (Err = 0) Then
            Do While Not (EOF(Handle%))
                Line Input #Handle%, h$
                If (FaxDruckKopfTxt$ <> "") Then FaxDruckKopfTxt$ = FaxDruckKopfTxt$ + vbCrLf
                FaxDruckKopfTxt$ = FaxDruckKopfTxt$ + h$
            Loop
            Close #Handle%
            Call OemToChar(FaxDruckKopfTxt$, FaxDruckKopfTxt$)
        Else
            FaxDruckKopfTxt$ = para.Fistam(0) + vbCrLf + para.Fistam(1)
        End If
        On Error GoTo DefErr
        
        
        AnzDruckSpalten% = 11
        ReDim DruckSpalte(AnzDruckSpalten% - 1)
        
        With DruckSpalte(0)
            .Titel = "P Z N"
            .TypStr = String$(7, "9")
            .Ausrichtung = "L"
        End With
        With DruckSpalte(1)
            .Titel = "A R T I K E L"
            .TypStr = String$(25, "X")  '28
            .Ausrichtung = "L"
        End With
        With DruckSpalte(2)
            .Titel = ""
            .TypStr = String$(6, "X")
            .Ausrichtung = "R"
        End With
        With DruckSpalte(3)
            .Titel = ""
            .TypStr = String$(5, "X")
            .Ausrichtung = "L"
        End With
        With DruckSpalte(4)
            .Titel = "A E P"
            .TypStr = "99999.99"
            .Ausrichtung = "R"
        End With
        With DruckSpalte(5)
            .Titel = "B M"
            .TypStr = String$(5, "9")
            .Ausrichtung = "R"
        End With
        With DruckSpalte(6)
            .Titel = "N R"
            .TypStr = String$(5, "9")
            .Ausrichtung = "R"
        End With
        With DruckSpalte(7)
            .Titel = "Z R"
            .TypStr = "99.99"
            .Ausrichtung = "R"
        End With
        With DruckSpalte(8)
            .Titel = "AngAEP"
            .TypStr = "99999.99"
            .Ausrichtung = "R"
        End With
        With DruckSpalte(9)
            .Titel = "FR"
            .TypStr = "999"
            .Ausrichtung = "R"
        End With
        With DruckSpalte(10)
            .Titel = "Wert"
            .TypStr = "99999.99"
            .Ausrichtung = "R"
        End With
        
        Call InitDruckZeile(True)
        
'        For k% = 1 To AnzRetourenDruck%
            DruckSeite% = 0
            AnzFaxDruckArtikel% = 0
            AnzFaxDruckPackungen% = 0
            FaxDruckWert# = 0#
            FaxDruckWert2# = 0#
            Call FaxDruckKopf(AktLief%)
            
            anz% = .ListCount
            For i% = 1 To anz%
                .ListIndex = i% - 1
                h$ = .text + vbTab
                
                ' nur wenn BM>0
                If (InStr(h$, vbTab + "0" + vbTab) = 0) Then
                    ind% = InStr(h$, Chr$(10))
                    If (ind% > 0) Then
                        h$ = Mid$(h$, ind% + 1) + Left$(h$, ind% - 1) + vbTab
                    End If
                    
                    Call FaxDruckZeile(h$)
                End If
                
                If (Printer.CurrentY > Printer.ScaleHeight - FaxDruckFussHoehe%) Then
                    Call FaxDruckFuss
                    Call FaxDruckKopf(AktLief%)
                End If
            Next i%
            
            Call FaxDruckSumme
            Call FaxDruckFuss(False)
            Printer.EndDoc
'        Next k%
        
        If (ManuellDruck%) Then
'            Call StopAnimation
            If (iMsgBox("Ausdruck in Ordnung ?", vbYesNo Or vbDefaultButton1) = vbYes) Then Exit Do
        Else
            Exit Do
        End If
    Loop
End With

If (DruckerWechsel%) Then
    For i% = 0 To (Printers.Count - 1)
        h$ = Printers(i%).DeviceName
        If (h$ = AktDruckerName$) Then
            Set Printer = Printers(i%)
            Exit For
        End If
    Next i%
End If

Call DefErrPop
End Sub

Sub InitDruckZeile(Optional ZentrierX% = False)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("InitDruckZeile")
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
Dim i%, j%, zentr%
Dim GesBreite&
Dim h$
        
For j% = 12 To 5 Step -1
    Printer.ScaleMode = vbTwips
    Printer.Font.Name = "Arial"
    Printer.Font.Size = 18  'nötig wegen Canon-BJ; sonst ab 2.Ausdruck falsch
    Printer.Font.Size = j%
        
    DruckSpalte(0).StartX = 0
    For i% = 0 To (AnzDruckSpalten% - 1)
        If (i% = 0) Then
            DruckSpalte(0).StartX = 0
        Else
            DruckSpalte(i%).StartX = DruckSpalte(i% - 1).StartX + DruckSpalte(i% - 1).BreiteX + Printer.TextWidth("  ")
        End If
        DruckSpalte(i%).BreiteX = Printer.TextWidth(RTrim(DruckSpalte(i%).TypStr))
    Next i%
    
    GesBreite& = DruckSpalte(AnzDruckSpalten% - 1).StartX + DruckSpalte(AnzDruckSpalten% - 1).BreiteX
    If (GesBreite& < Printer.ScaleWidth) Then Exit For
Next j%

DruckFontSize% = j%

If (ZentrierX%) Then
    zentr% = (Printer.ScaleWidth - GesBreite&) / 2
    For i% = 0 To (AnzDruckSpalten% - 1)
        DruckSpalte(i%).StartX = DruckSpalte(i%).StartX + zentr%
    Next i%
End If

Call DefErrPop
End Sub

Sub FaxDruckKopf(Lief%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("FaxDruckKopf")
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
Dim l&, he&, i%, pos%, x%, Y%, GesBreite&
Dim h$, heute$, SeitenNr$, header$

header$ = "Bestellung"

heute$ = Format(Day(Date), "00") + "-"
heute$ = heute$ + Format(Month(Date), "00") + "-"
heute$ = heute$ + Format(Year(Date), "0000")
'heute$ = heute$ + " " + Left$(Time$, 5)

With Printer
    
    .CurrentX = 0: .CurrentY = 0
    .Font.Size = 11
    Printer.Print FaxDruckKopfTxt$
    Y% = .CurrentY
    Printer.Line (0, Y%)-(.ScaleWidth, Y%)
    .CurrentX = 0
    .CurrentY = .CurrentY + 300
    .Font.Size = 14
    
    DruckSeite% = DruckSeite% + 1
    
    If (lifzus.DirektBestFaxKz) Then
        h$ = "Per Fax an " + Trim(lifzus.DirektBestFax)
        l& = .TextWidth(h$)
        .CurrentX = .ScaleWidth - l& - 10
        Printer.Print h$;
        .CurrentX = 0
    End If
    
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
        
    .CurrentX = (.ScaleWidth - l&) / 2
    .CurrentY = .CurrentY + 400
    Printer.Print header$;
    
    l& = .TextWidth(heute$)
    .CurrentX = .ScaleWidth - l& - 10
    Printer.Print heute$
    
    Printer.Print
    .Font.Size = DruckFontSize%
    Printer.Print
    Printer.Print
    
    For i% = 0 To (AnzDruckSpalten% - 1)
        h$ = RTrim(DruckSpalte(i%).Titel)
        If (DruckSpalte(i%).Ausrichtung = "L") Then
            x% = DruckSpalte(i%).StartX
        Else
            x% = DruckSpalte(i%).StartX + DruckSpalte(i%).BreiteX - Printer.TextWidth(h$)
        End If
        .CurrentX = x%
        Printer.Print h$;
    Next i%
    
    Printer.Print " "
    
    Y% = Printer.CurrentY
    GesBreite& = DruckSpalte(AnzDruckSpalten% - 1).StartX + DruckSpalte(AnzDruckSpalten% - 1).BreiteX
    Printer.Line (DruckSpalte(0).StartX, Y%)-(GesBreite&, Y%)

    Y% = Printer.CurrentY
    Printer.CurrentY = Y% + 30
    
End With

Call DefErrPop
End Sub

Sub FaxDruckFuss(Optional NewPage% = True)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("FaxDruckFuss")
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
            
    .CurrentX = (.ScaleWidth - .TextWidth(h$)) / 2
    .CurrentY = .ScaleHeight - .TextHeight(h$)
    Printer.Print h$
    
    If (FaxDruckFussTxt$ <> "") Then
        .CurrentX = 0
        .CurrentY = .CurrentY - (FaxDruckFussHoehe% - 567)
        Printer.Print FaxDruckFussTxt$;
    End If
    
    If (NewPage% = True) Then .NewPage
End With

Call DefErrPop
End Sub

Sub FaxDruckZeile(ZeilenText$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("FaxDruckZeile")
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
Dim i%, ind%, x%
Dim h$, tx$

h$ = ZeilenText$

For i% = 0 To (AnzDruckSpalten% - 1)
    ind% = InStr(h$, vbTab)
    tx$ = Left$(h$, ind% - 1)
    h$ = Mid$(h$, ind% + 1)
    
    If (tx$ = Chr$(214)) Then
        Printer.Font.Name = "Symbol"
    End If
    
    If (DruckSpalte(i%).Ausrichtung = "L") Then
        x% = DruckSpalte(i%).StartX
    Else
        x% = DruckSpalte(i%).StartX + DruckSpalte(i%).BreiteX - Printer.TextWidth(tx$)
    End If
    Printer.CurrentX = x%
    Printer.Print tx$;
    
    If (tx$ = Chr$(214)) Then
        Printer.Font.Name = "Arial"
    End If
    
    If (i% = (AnzDruckSpalten% - 1)) Then
        FaxDruckWert# = FaxDruckWert# + CDbl(tx$)
    End If
Next i%
Printer.Print " "
                
AnzFaxDruckArtikel% = AnzFaxDruckArtikel% + 1

ind% = InStr(h$, vbTab)
tx$ = Left$(h$, ind% - 1)
h$ = Mid$(h$, ind% + 1)
FaxDruckWert2# = FaxDruckWert2# + CDbl(tx$)
    
ind% = InStr(h$, vbTab)
tx$ = Left$(h$, ind% - 1)
h$ = Mid$(h$, ind% + 1)
AnzFaxDruckPackungen% = AnzFaxDruckPackungen% + Val(tx$)
    
Call DefErrPop
End Sub

Sub FaxDruckSumme()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("FaxDruckSumme")
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
Dim Y%, GesBreite&, tx$
            

Y% = Printer.CurrentY
GesBreite& = DruckSpalte(AnzDruckSpalten% - 1).StartX + DruckSpalte(AnzDruckSpalten% - 1).BreiteX
Printer.Line (DruckSpalte(0).StartX, Y%)-(GesBreite&, Y%)

Y% = Printer.CurrentY
Printer.CurrentY = Y% + 30

Printer.CurrentX = DruckSpalte(1).StartX
Printer.Print Format(AnzFaxDruckArtikel%, "0") + " Position(en) / " + Format(AnzFaxDruckPackungen%, "0") + " Packung(en)";
tx$ = Format(FaxDruckWert#, "0.00")     'MarkWert#
Printer.CurrentX = GesBreite& - Printer.TextWidth(tx$)
Printer.Print tx$;
Printer.Print " "

If (FaxDruckWert# <> FaxDruckWert2#) Then
    Printer.CurrentX = DruckSpalte(7).StartX
    Printer.Print "-" + Str$(DirektBezugFaktRabatt#) + "% Fakt.Rabatt";
    tx$ = Format(FaxDruckWert2# - FaxDruckWert#, "0.00")
    Printer.CurrentX = GesBreite& - Printer.TextWidth(tx$)
    Printer.Print tx$;
    Printer.Print " "
    
    Y% = Printer.CurrentY
    Printer.Line (DruckSpalte(10).StartX, Y%)-(GesBreite&, Y%)
    Y% = Printer.CurrentY
    Printer.CurrentY = Y% + 30
    tx$ = Format(FaxDruckWert2#, "0.00")
    Printer.CurrentX = GesBreite& - Printer.TextWidth(tx$)
    Printer.Print tx$;
    Printer.Print " "
End If

If (DirektBezugValutaStellung% > 0) Then
    Printer.Print " "
    Printer.Print "Valutastellung:" + Str$(DirektBezugValutaStellung%) + " Tage"
End If
    
Printer.Print " "
tx$ = "(AEP .. TaxeAep, ZR .. ZeilenRabatt, AngAEP .. AngebotsAep, FR .. mit FakturenRabatt, "
tx$ = tx$ + "Wert .. Zeilenwert)"
Printer.Print tx$
           
Call DefErrPop
End Sub

Sub DirektBezugMail()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("DirektBezugMail")
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
Dim i%, anz%, ind%
Dim text$, OrgDir$, h$

OrgDir$ = CurDir

'Call StartAnimation("Email wird erstellt ...")

With frmAction.lstDirektSortierung
    text$ = ""
    anz% = .ListCount
    For i% = 1 To anz%
        .ListIndex = i% - 1
        h$ = .text + vbTab
        
        ind% = InStr(h$, Chr$(10))
        If (ind% > 0) Then
            h$ = Mid$(h$, ind% + 1) + Left$(h$, ind% - 1)
        End If
        
        text$ = text$ + h$ + vbCrLf
        
                
        
    Next i%
End With

'With frmAction
'    .MAPISession1.SignOn
'    If Err <> 0 Then
'        MsgBox "Logon Failure: " + Error$
'    End If
'    .MAPIMessages1.SessionID = .MAPISession1.SessionID
'    .MAPIMessages1.MsgIndex = -1
'    .MAPIMessages1.Compose
'    .MAPIMessages1.RecipDisplayName = RTrim(lifzus.DirektBestMail)
'    .MAPIMessages1.RecipAddress = RTrim(lifzus.DirektBestMail)
'    'MAPIMessages1.AddressResolveUI = True
'    'MAPIMessages1.ResolveName
'    .MAPIMessages1.MsgSubject = "Bestellung"
'    .MAPIMessages1.MsgNoteText = text$
'    .MAPIMessages1.Send
'    .MAPISession1.SignOff
'End With

'Call StopAnimation

ChDrive (OrgDir$)
ChDir (OrgDir$)

'If (iMsgBox("Email in Ordnung ?", vbYesNo Or vbDefaultButton1) = vbYes) Then Exit Do

Call DefErrPop
End Sub

Sub UpdateRueckKaufDat(Lief%, Ueberleiten%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("UpdateRueckKaufDat")
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
Dim i%, j%, SN%, bm%, nm%, MM%, bMax%, ssatz%, Max%, ok%, rInd%
Dim bZeit$, pzn$, OrgPzn$, text$, tx$, s$, lKurz$
Dim k&

bZeit$ = MKI(Val(Left$(Time$, 2)) * 100 + Val(Mid$(Time$, 4, 2)))

rk.GetRecord (1)
Max% = rk.erstmax

If (Ueberleiten%) Then
    DRUCKBUCH% = FileOpen("winw\" + SendDruckDateiName$ + ".sp9", "O")
    For i% = 0 To (AnzAbsagenZeilen% - 1)
        Print #DRUCKBUCH%, AbsagenZeilen$(i%)
    Next i%
    AnzAbsagenZeilen% = 0
    
    For i% = 0 To (AnzBestellArtikel% - 1)
        SN% = Val(Right$(SRT$(i%), 4))
        If (SN% > 0) Then                 '1.88
'            bek.GetRecord (SN% + 1)
            rInd% = SucheRueckKaufZeile(SN%, Max%, mlauf&(i%))
            If (rInd%) And (rk.pzn = mnr$(i%)) And (rk.aktivlief = Lief%) Then
                
                text$ = DruckZeile$
                Print #DRUCKBUCH%, text$
                
                bm% = rk.bm: nm% = rk.nm: pzn$ = rk.pzn
                If Left$(pzn$, 4) = "9999" And Mid$(pzn$, 5, 3) <> "999" Then
                    rk.status = 0
                Else
                    ssatz% = 0
                    OrgPzn$ = pzn$
                    FabsErrf% = ass.IndexSearch(0, pzn$, FabsRecno&)
                    If (FabsErrf% = 0) Then
                        ssatz% = FabsRecno&
                        ass.GetRecord (FabsRecno& + 1)
                        MM% = ass.MM
                    Else
                        MM% = 0
                    End If
                    
                    tx$ = pzn$ + Format(ssatz%, " ####0") + Format(bm%, " ###0")

                    text$ = pzn$ + " " + Left$(rk.txt, 33) + " " + Mid$(rk.txt, 34, 2)
                    If (bm% = 0) And (nm% = 0) Then     '1.60 Merkzettel
'                        Call Merkzettel(pzn$, Mid$(text$, 9), xcBestDatum%, xcBestDatum%, Lieferant%, 0)
                        rk.status = 0
                    Else
                        rk.status = 2
                        rk.Lief = Lief%
                        rk.lief1 = Lief%
                        
                        rk.WuAEP = rk.aep
                        
                        rk.WuBestDatum = HeuteDatStr$
                        rk.WuBm = Abs(bm%)
                        rk.WuNm = Abs(nm%)
                        
                        rk.WuStat = "J"
                        If (bm% < 0) Or (nm% < 0) Then rk.WuStat = "N"
                        
                        rk.WuRm = Abs(bm%)
                        rk.WuLm = Abs(bm%) + Abs(nm%)
                        rk.WuAm = 0
'                        rk.WuAblDatum = Space$(6)
                        rk.WuLa = MM%
                        rk.WuBestZeit = bZeit$
                        rk.WuAVP = rk.avp
                        rk.WuBelegDatum = 0
                        rk.WuBeleg = Space$(10)
                        rk.WuRetMenge = 0
'                            ww.WuAnzBereitsGebucht = 0
'                            ww.WuFertig = 0
                        rk.WuNNaepOk = 0
                        rk.WuNNart = rk.nnart
                        rk.WuNNAep = rk.NNAEP
                        rk.WuText = Space$(Len(rk.WuText))
                        rk.WuNeuLm = 0
                        rk.WuNeuRm = 0
                        rk.WuNeuZiel = 0
                        
                        rk.IstAltLast = 0
                        rk.WuStatus = 0
                        rk.LmStatus = 0
                        rk.RmStatus = 0
                        rk.LmAnzGebucht = 0
                        rk.RmAnzGebucht = 0
                        
                        rk.aktivlief = 0
                        rk.aktivind = 0
                    End If
                  
                End If
                rk.PutRecord (rInd% + 1)
            End If
        End If
    Next i%
    
    For i% = 1 To MaxSendSatz%
        text$ = SendSatz$(i%)
        text$ = Left$(text$, Len(text$) - 1)
        Print #DRUCKBUCH%, "RM: " + text$
    Next i%
    Close #DRUCKBUCH%
                        
    Call EntferneGeloeschteRueckKauf
Else
    For i% = 1 To Max%
        rk.GetRecord (i% + 1)
        
        ok% = False
        If (Lief% = -1) Or (rk.aktivlief = Lief%) Then ok% = True
        If (rk.status <> 1) Then ok% = False
        
        If (ok%) Then
            rk.aktivlief = 0
            rk.aktivind = 0
            rk.PutRecord (i% + 1)
        End If
    Next i%
End If


rk.erstcounter = (rk.erstcounter + 1) Mod 100
rk.PutRecord (1)

Call DefErrPop
End Sub

Sub ResetWbestk2Senden()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("ResetWbestk2Senden")
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
Dim i%, j%, Max%
Dim h$, Leer$

ww.SatzLock (1)
ww.GetRecord (1)

Max% = ww.erstmax

For i% = 1 To Max%
    ww.GetRecord (i% + 1)
    
    If (ww.aktivind < 0) Then
        ww.aktivlief = 0
        ww.aktivind = 0
        ww.PutRecord (i% + 1)
    End If
Next i%

ww.erstcounter = (ww.erstcounter + 1) Mod 100
ww.PutRecord (1)
ww.SatzUnLock (1)


rk.OpenDatei
If (rk.DateiLen = 0) Then
    rk.erstmax = 0
    rk.erstlief = 0
    rk.erstcounter = 0
    rk.erstrest = String(rk.DateiLen, 0)
    rk.PutRecord (1)
End If

rk.GetRecord (1)
Max% = rk.erstmax

For i% = 1 To Max%
    rk.GetRecord (i% + 1)
    
    If (rk.aktivind < 0) Then
        rk.aktivlief = 0
        rk.aktivind = 0
        rk.PutRecord (i% + 1)
    End If
Next i%

rk.erstcounter = (rk.erstcounter + 1) Mod 100
rk.PutRecord (1)
rk.CloseDatei

Call DefErrPop
End Sub

Sub InsertLstDirektSortierung()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("InsertLstDirektSortierung")
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
Dim zr!
Dim EK#, AngEK#, fr#, ZeilenWert#
Dim SQLStr$, h$, tx$

EK# = ww.aep
SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + ww.pzn
Set TaxeRec = TaxeDB.OpenRecordset(SQLStr$)
If (TaxeRec.EOF = False) Then
    EK# = TaxeRec!EK / 100
End If

AngEK# = ww.aep
If (ww.NNAEP > 0) Then
    AngEK# = ww.NNAEP
End If

h$ = Left$(ww.txt, 28) + vbTab + Mid$(ww.txt, 29, 5)
h$ = h$ + vbTab + Mid$(ww.txt, 34, 2)
h$ = h$ + vbTab + Format(EK#, "0.00")
h$ = h$ + vbTab + Format(Abs(ww.bm), "0")

tx$ = ""
If (Abs(ww.nm > 0)) Then
    tx$ = Format(Abs(ww.nm), "0")
End If
h$ = h$ + vbTab + tx$

zr! = ww.zr
fr# = 0

If (zr! <= -200) Then
    zr! = zr! + 200
ElseIf (zr! >= 200) Then
    zr! = zr! - 200
End If

If (ww.zr < 0) Then
    If (zr! = -123.45) Then zr! = 0#
    zr! = Abs(zr!)
    If (ww.nm = 0) Or (DirektBezugFaktRabattTyp% <> 1) Then
        fr# = DirektBezugFaktRabatt#
    End If
End If

tx$ = ""
If (zr! > 0) Then
    tx$ = Format(zr!, "0")
End If
h$ = h$ + vbTab + tx$
h$ = h$ + vbTab + Format(AngEK#, "0.00")

tx$ = ""
If (fr# > 0) Then
    tx$ = Chr$(214) ' Format(fr#, "0")
End If
h$ = h$ + vbTab + tx$

ZeilenWert# = AngEK# * Abs(ww.bm)

'                h$ = h$ + vbTab + Format(EK# * Abs(ww.bm), "0.00")
h$ = h$ + vbTab + Format(ZeilenWert#, "0.00")

ZeilenWert# = ZeilenWert# * (100# - fr#) / 100#
h$ = h$ + vbTab + Format(ZeilenWert#, "0.00")

h$ = h$ + vbTab + Format(Abs(ww.bm) + Abs(ww.nm), "0")

h$ = h$ + Chr$(10) + ww.pzn
frmAction!lstDirektSortierung.AddItem h$

Call DefErrPop
End Sub

Sub SpeicherManuellSendung()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("SpeicherManuellSendung")
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
Dim AktivSend$
    
Call GetSendungParameter

With SendeForm
    If (.optAuftrag(0).Value = True) Then
        AktivSend$ = "J"
    Else
        AktivSend$ = "N"
    End If
End With

Call HoleIniRufzeiten
i% = AnzRufzeiten%
Rufzeiten(i%).Lieferant = Lieferant%
Rufzeiten(i%).Aktiv = AktivSend$
Rufzeiten(i%).AuftragsErg = AuftragsErg$ ' "ZH"
Rufzeiten(i%).AuftragsArt = AuftragsArt$ ' "  "
Rufzeiten(i%).Gewarnt = "J"
Rufzeiten(i%).RufZeit = 9999
Rufzeiten(i%).LieferZeit = 9999
Rufzeiten(i%).LetztSend = 0
AnzRufzeiten% = AnzRufzeiten% + 1

If (RueckKaufSendung%) Then
    Rufzeiten(i%).RufZeit = 9998
    Rufzeiten(i%).LieferZeit = 9998
End If

Call SpeicherIniRufzeiten
Call HoleFruehesteManuelleSendezeit%

Call DefErrPop
End Sub

Sub SaetzeVorbereitenA()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("SaetzeVorbereitenA")
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
Dim i%, j%, k%, gesendet%, menge%, IstFrei%, IW%, im%, ind%
Dim NUM&
Dim satz$, char$, fistsatz$, BitFeld$, TestPZN$, sMenge$, txt$, x$
Dim RufKennung$, KundenStandard$, DatumUhrzeit$, AuftragsNummer$, BetrNr$, pzn2$

AnzAbsagenZeilen% = 0

If (BlindBestellung%) Then
    AktivSenden% = False
    AuftragsErg$ = "  "
    AuftragsArt$ = "  "
End If
    


'2 bytes satztyp
'2 bytes gesamtlänge

'A Apotheke, G Großhandel , O Großhandel ohne Übernahme
RufKennung$ = "A"
KundenStandard$ = "00"

'Datum Uhrzeit
x$ = Right$(Date$, 2) + Mid$(Date$, 1, 2) + Mid$(Date$, 4, 2)
x$ = x$ + Left$(Time$, 2) + Mid$(Time$, 4, 2)
ind% = 1
While (ind% <> 0)
    ind% = InStr(x$, " "): If (ind% > 0) Then Mid$(x$, ind%, 1) = "0"
Wend
DatumUhrzeit$ = x$

AuftragsNummer$ = "01"

NUM& = Val(para.BetriebsNummer)
BetrNr$ = Right$("00000" + Mid$(Str$(NUM&), 2) + "00", 7)
Mid$(BetrNr$, 5, 1) = "1": 'Nur für Apotheken!

Do
    IW% = 0
    For i% = 1 To 6
        char$ = Mid$(BetrNr$, i%, 1)
        IW% = IW% + Val(char$) * (i% + 1)
    Next i%
    im% = (IW% Mod 11)                         '1.80
    If (im% = 10) Then
        char$ = Mid$(BetrNr$, 6, 1)
        Mid$(BetrNr$, 6, 1) = Right$(Str$(Val(char$) + 1), 1)
    Else
        Exit Do
    End If
Loop
Mid$(BetrNr$, 7, 1) = Right$(Str$(im%), 1)
'ApoIDF$ = BetrNr$


'startsatz
SendSatz$(0) = "0142" + GhIDF$ + ApoIDF$ + RufKennung$ + AuftragsErg$ + "999" + "D0"
SendSatz$(0) = SendSatz$(0) + KundenStandard$ + DatumUhrzeit$ + AuftragsArt$ + AuftragsNummer$


j% = 0

'nur schicken wenn 3er Satz vorhanden
'J = J + 1: satz$(J) = satz$


BitFeld$ = ""

'1.58 Sätze nicht mehr lesen - sind schon in Arrays gespeichert
For i% = 0 To (AnzBestellArtikel% - 1)
    TestPZN$ = mnr$(i%)
    char$ = Left$(TestPZN$, 1)
    
    IstFrei% = True
    If (char$ = "8") Then
        ' PZNs mit 8 am Anfang sind freie Nummern von MEDATA bzw. SANODAT
        IstFrei% = False
    ElseIf (char$ = "9") Then
        ' PZNs mit 9 am Anfang sind freie Nummern
        'Spezial PZNs für Order und ähnliches
        If Left$(TestPZN$, 4) <> "9999" Or Mid$(TestPZN$, 5, 3) = "999" Then IstFrei% = False
    End If
    
    If (IstFrei%) Then
        If Val(TestPZN$) <= 0 Then IstFrei% = False
    End If
    If (IstFrei%) Then
        If mbm%(i) = 0 Then IstFrei% = False
    End If

    If (IstFrei%) Then
        IstFrei% = TestBitFrei%(BitFeld$, i%)
    End If
    
    If (IstFrei%) Then
        satz$ = "0217  00010000023"
        
        menge% = Abs(mbm%(i%))
        k% = i% + 1
        While (k% < AnzBestellArtikel%)
          IstFrei% = TestBitFrei%(BitFeld$, k%)
          If (IstFrei%) Then
            If (mnr$(k%) = TestPZN$) Then
              menge% = menge% + Abs(mbm%(k%))
              Call SetBit(BitFeld$, k%)
            End If
          End If
          k% = k% + 1
        Wend
        
        sMenge$ = Right$("0000" + Mid$(Str$(menge%), 2), 4)
        Mid$(satz$, 7, 4) = sMenge$
        
        If (Mid$(Wg3Str$, i% + 1, 1) = "3") Then
            FabsErrf% = ast.IndexSearch(0, TestPZN$, FabsRecno&)
            If (FabsErrf% = 0) Then
                ast.GetRecord (FabsRecno& + 1)
                pzn2$ = ast.rez + ast.kas + ast.kasz
                If Val(pzn2$) <> 0 Then TestPZN$ = pzn2$
            End If
        End If
        Mid$(satz$, 11, 7) = TestPZN$
        
        j% = j% + 1: SendSatz$(j%) = satz$

        Call SetBit(BitFeld$, i%)
    End If
Next i%

MaxSendSatz% = j% + 1
'endsatz
SendSatz$(MaxSendSatz%) = "9904"

Call DefErrPop
End Sub

Function SeriellBestellungA%()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("SeriellBestellungA%")
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
Dim ret%, ret2%
Dim SendStr$, RecStr$, h$
Dim MeldungStr$

ret% = SeriellBestA%(MeldungStr$)

'ret% = True
'MaxSendSatz% = 2
'ReDim SendSatz$(MaxSendSatz%)
'h$ = Left$("Betaserc" + Space$(35), 35)
'SendSatz$(1) = "035600020002NL" + h$ + "2425057"
'h$ = Left$("Effortil" + Space$(35), 35)
'SendSatz$(2) = "0356000300030R" + h$ + "0792165"

If (ret%) Then
    If Not (AktivSenden%) Then
        If (AustriaExit$ <> "") Then
            SendStr$ = AustriaExit$
            ret2% = TransmitBlock%(SendStr$, RecStr$)
        End If
    End If
Else
    MeldungStr$ = "Die Datenfernübertragung wurde abgebrochen!"
    Call StatusZeile(MeldungStr$)
    If (AutomaticSend% = False) Then
        Call iMsgBox(MeldungStr$, vbCritical)
    Else
        AutomatikFehler$ = MeldungStr$
    End If
End If

SeriellBestellungA% = ret%

Call DefErrPop
End Function

Function SeriellBestA%(MeldungStr$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("SeriellBestA%")
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
Dim j%, ret%, Warte%, OrgWarte%, AnzBesetzt%
Dim status$, SendStr$, RecStr$, satz01$, deb$
Dim ch$

SeriellBestA% = False

RS$ = Chr$(&H1E)
STX$ = Chr$(2)
ETX$ = Chr$(3)
EOT$ = Chr$(4)
ENQ$ = Chr$(5)
NAK$ = Chr$(&H15)
TDD$ = Chr$(2) + Chr$(5)
WACK$ = Chr$(&H10) + "3"
ACK0$ = Chr$(&H10) + "0"
ACK1$ = Chr$(&H10) + "1"
DLEEOT$ = Chr$(&H10) + Chr$(4)

With SendeForm.comSenden
    While (.InBufferCount > 0)
        ch$ = .Input
    Wend
End With


TimeOutZeit% = 60
FehlerTimeOut% = 0
FehlerBcc% = 0
FehlerNak% = 0


status$ = "Modem-Initialisierung wird durchgeführt.": Call StatusZeile(status$)
SendStr$ = AustriaSetup$
ret% = TransmitBlock%(SendStr$, RecStr$)
MeldungStr$ = "Fehler bei Modem-Initialisierung"
If (ret% = False) Then Call DefErrPop: Exit Function

Call IsdnPause(4)

If (AktivSenden%) Then
    status$ = "Telefonnummer " + TelGh$ + " wird gewählt.": Call StatusZeile(status$)
    SendStr$ = AustriaDialP$ + TelGh$ + AustriaDialS$
    ret% = TransmitBlock%(SendStr$, RecStr$)
    
    MeldungStr$ = "Fehler beim Wählen"
    If (ret% = False) Then Call DefErrPop: Exit Function
Else
    status$ = "Wartemodus wird aktiviert ..."
    Call StatusZeile(status$)
'    If (warten%) Then
        SendStr$ = AustriaWait$
'    Else
'        SendStr$ = AustriaAnswer$
'    End If
    ret% = TransmitBlock%(SendStr$, RecStr$)
    
    MeldungStr$ = "Fehler beim Warten"
End If

If (ret% = False) Then Call DefErrPop: Exit Function

Call IsdnPause(4)
status$ = "Warten auf Telefonverbindung und Datenton ...": Call StatusZeile(status$)

While (Mid$(RecStr$, 1, 7) <> "CONNECT")
'    If warten% Then LOCATE 17, 73: Print Left$(Time$, 5);
    
    ret% = ReceiveBlock%(RecStr$)

    If (ret% = 0) And (Mid$(RecStr$, 1, 4) = "BUSY") Then
        AnzBesetzt% = AnzBesetzt% + 1
        If (AnzBesetzt% < 3) Then
            deb$ = "Neuerlicher Versuch in 60 Sekunden ...": Call StatusZeile(deb$)
            Call IsdnPause(60)
            status$ = "Telefonnummer " + TelGh$ + " wird gewählt.": Call StatusZeile(status$)
            SendStr$ = AustriaDialP$ + TelGh$ + AustriaDialS$
            ret% = TransmitBlock%(SendStr$, RecStr$)
            MeldungStr$ = "Fehler beim Wählen"
        End If
    End If
    
    If (Mid$(RecStr$, 1, 4) = "RING") Then
        status$ = "Telefon hat geläutet, Modem hebt ab ...": Call StatusZeile(status$)
    End If
    
    If (AktivSenden%) Then
    Else
        If (TimeOut%) Then ret% = True
    End If
    
    If (ret% = False) Then Call DefErrPop: Exit Function
Wend

status$ = "Die Verbindung ist korrekt hergestellt!": Call StatusZeile(status$)

status$ = "Empfang der Großhandel-Kennung.": Call StatusZeile(status$)

ACK% = 0: ackGH% = 0
Do
    ret% = RecvBlocks%
    
    If (RecStr$ <> "") Then Exit Do
    
    If (ret% = False) Then Call DefErrPop: Exit Function
Loop

status$ = "Senden der Bestellung ...": Call StatusZeile(status$)
ImSenden% = True: ret% = SendBlocks%: ImSenden% = 0
If (ret%) Then
    '* ab hier gilt es als bestellt

    '* lt.Hestag hier wieder mit ACK0 beginnen
    ACK% = 0

    status$ = "Empfangen der Rückmeldungen vom Großhandel ...": Call StatusZeile(status$)
    ret% = RecvBlocks%
    ret% = True
    
    If (MaxRecSatz% > 1) Then MaxRecSatz% = MaxRecSatz% - 1
            
    MaxSendSatz% = MaxRecSatz%
    ReDim SendSatz$(MaxSendSatz%)
    For j% = 0 To MaxSendSatz%
        SendSatz$(j%) = RecSatz$(j%)
    Next j%
End If


SeriellBestA% = ret%

Call DefErrPop
End Function

Function TransmitBlock%(SendStr$, RecStr$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("TransmitBlock%")
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
Dim ret%

Call OutputComStr(SendStr$)

ret% = ReceiveBlock%(RecStr$)

'If (ret%) Then RecStr$ = Left$(RecStr$, Len(RecStr$) - 1)

TransmitBlock% = ret%

Call DefErrPop
End Function

Function ReceiveBlock%(RecStr$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("ReceiveBlock%")
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
Dim l%, ret%, TimeOut2%, FehlerTimeOut%
Dim char$, deb$
Dim TimerEnd
Dim ch As Variant
Dim chByte() As Byte


Do
    ret% = False
    RecStr$ = ""
    TimeOut% = 0
    TimerEnd = Timer + TimeOutZeit
    
    Do
        char$ = ""
        
        If (Timer > TimerEnd) Then
            TimeOut% = True
            Exit Do
        End If
        
        l% = SendeForm.comSenden.InBufferCount
        If (l% > 0) Then
            ch = SendeForm.comSenden.Input
            chByte = ch
            char$ = Chr$(chByte(0))
'            Debug.Print "<" + Format(chByte(0), "0") + ">" + "  " + char$
            
            If (char$ = vbLf) Or (Asc(char$) > 127) Or (Asc(char$) = 0) Then
                char$ = ""
            End If
            If (char$ <> vbCr) And (char$ <> "") Then
                RecStr$ = RecStr$ + char$
                TimerEnd = Timer + TimeOutZeit
            ElseIf (RecStr$ <> "") Then
                ret% = True
                Exit Do
            End If
        Else
            DoEvents
            If (BestSendenAbbruch% = True) Then
                Exit Do
            End If
        End If
        If (Len(RecStr$) > 500) Then
            RecStr$ = ""
            TimeOut% = True
            Exit Do
        End If
    Loop
    
    deb$ = "< " + RecStr$ + char$: Call StatusZeile(deb$)
            
    If (TimeOut%) Then
        FehlerTimeOut% = FehlerTimeOut% + 1
        If (ImSenden%) And Not (TimeOut2%) Then
            TimeOut2% = True
            Call OutputComStr(ENQ$ + vbCr)
        Else
            Exit Do
        End If
    ElseIf (ret%) Then
        TimeOut2% = 0
        If (Left$(RecStr$, 1) = STX$) Then
            If (BCCcheck%(RecStr$)) Then
                Exit Do
            Else
                FehlerBcc% = FehlerBcc% + 1
                Call OutputComStr(NAK$ + vbCr)
            End If
        ElseIf (RecStr$ = WACK$) Then
            Call OutputComStr(ENQ$ + vbCr)
        ElseIf (RecStr$ = TDD$) Then
            Call OutputComStr(NAK$ + vbCr)
        Else
            Exit Do
        End If
    Else
        Exit Do
    End If
Loop

If (Mid$(RecStr$, 1, 5) = "ERROR") Then
    deb$ = "Keine Telefonverbindung vorhanden! (Fehler)": Call StatusZeile(deb$)
    ret% = False
End If
If (Mid$(RecStr$, 1, 11) = "NO DIALTONE") Then
    deb$ = "Keine Telefonverbindung vorhanden! (kein Wählton)": Call StatusZeile(deb$)
    ret% = False
End If
If (Mid$(RecStr$, 1, 4) = "BUSY") Then
    deb$ = "Keine Telefonverbindung vorhanden! (besetzt)": Call StatusZeile(deb$)
    ret% = False
End If
If (Mid$(RecStr$, 1, 10) = "NO CARRIER") Then
    deb$ = "Keine Telefonverbindung vorhanden! (kein Datenton)": Call StatusZeile(deb$)
    ret% = False
End If

ReceiveBlock% = ret%

Call DefErrPop
End Function

Sub OutputComStr(SendStr$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("OutputComStr")
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
Dim deb$

SendeForm.comSenden.Output = SendStr$
deb$ = "> " + SendStr$: Call StatusZeile(deb$)

Call DefErrPop
End Sub

Function BCCcheck%(s$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("BCCcheck%")
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
Dim i%, ret%
Dim OrgBcc$, bcc$, bcc2$, Pruef$

ret% = False

If (Len(s$) >= 4) Then
    OrgBcc$ = Right$(s$, 2)
    Pruef$ = Mid$(s$, 2, Len(s$) - 3)
    
    bcc$ = MakeBccA(Pruef$)
    
    If (OrgBcc$ = bcc$) Then ret% = True
End If

BCCcheck% = ret%

Call DefErrPop
End Function

Function MakeBccA$(Pruef$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("MakeBccA$")
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
Dim i%, bcc%
Dim bcc1$, bcc2$
    
bcc% = 0
For i% = 1 To Len(Pruef$)
    bcc% = (bcc% Xor Asc(Mid$(Pruef$, i%, 1)))
Next i%
bcc1$ = Chr$(&H20 + (bcc% And &HF))
bcc2$ = Chr$(&H20 + (bcc% And &HF0) / 16)

MakeBccA$ = bcc1$ + bcc2$

Call DefErrPop
End Function

Function RecvBlocks%()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("RecvBlocks%")
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
Dim i%, ret%, ind%, ENQtest%, ok%
Dim x$, satz$, sACK$

RecvBlocks% = False

MaxRecSatz% = 0
ReDim RecSatz$(MaxSendSatz%)
RecSatz$(0) = Space$(14) + "DATENÜBERTRAGUNG RÜCKMELDUNGEN: "



TimeOutZeit = 30
ENQtest% = 0
While (RecStr$ <> ENQ$)
    ret% = ReceiveBlock%(RecStr$)
    ENQtest% = ENQtest% + 1
    If (ret% = False) Or (ENQtest% > 10) Then Call DefErrPop: Exit Function
Wend

ok% = True
Do
    If (ACK% = 0) Then
        sACK$ = ACK0$
    Else
        sACK$ = ACK1$
    End If
    
    If ok% Then
        ret% = TransmitBlock%(sACK$ + vbCr, RecStr$)
    Else
        ret% = ReceiveBlock%(RecStr$)
    End If
    
    If (ret% = False) Then
        Call DefErrPop: Exit Function
    End If
    
    If (RecStr$ = EOT$) Or (RecStr$ = DLEEOT$) Then Exit Do
    
    ok% = 0
    If (Left$(RecStr$, 1) = STX$) Then
        'Blöcke Zerlegen
        x$ = Mid$(RecStr$, 2, Len(RecStr$) - 4)
        While (Len(x$) > 0)
            ind% = InStr(x$, RS$): If (ind% = 0) Then ind% = Len(x$) + 1
            satz$ = Left$(x$, ind% - 1)
            
            MaxRecSatz% = MaxRecSatz% + 1
            ReDim Preserve RecSatz$(MaxRecSatz%)
            RecSatz$(MaxRecSatz%) = satz$
            
            x$ = Mid$(x$, ind% + 1)
        Wend
        ACK% = Not (ACK%): ok% = True
    End If
    If (RecStr$ = ENQ$) Then ok% = True
    If (RecStr$ = EOT$) Then ACK% = Not (ACK%)
Loop

If (RecStr$ = DLEEOT$) Then
'    TimeOutZeit = 3
'    For i% = 1 To 3
'        ret% = TransmitBlock%(DLEEOT$ + vbCr, RecStr$)
'        If (Mid$(RecStr$, 1, 10) = "NO CARRIER") Then
'            Exit For
'        End If
'    Next i%
    Call DefErrPop: Exit Function
Else
    TimeOutZeit = 10
    ret% = TransmitBlock%(ENQ$ + vbCr, RecStr$)
    
    If (ackGH% = 0) Then
        sACK$ = ACK0$
    Else
        sACK$ = ACK1$
    End If
    
    If (ret%) Then
        Do
            If (RecStr$ = ACK0$) Then
                Exit Do
            End If
            
            ret% = ReceiveBlock%(RecStr$)
            
            If (ret% = False) Then Call DefErrPop: Exit Function
        Loop
    End If
    
    ackGH% = Not (ackGH%)
End If

RecvBlocks% = ret%

Call DefErrPop
End Function

Function SendBlocks%()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("SendBlocks%")
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
Dim j%, SatzCount%, NakCount%, Blockung%, iNAK%, SatzSend%, ret%
Dim deb$, bcc$, sACK$, altACK$, SendStr$

SendBlocks% = False

TimeOutZeit = 10
SatzCount% = 0: Blockung% = True: NakCount% = 0

deb$ = Right$("    " + Str$(MaxSendSatz% + 1), 4) + " Sätze noch zu senden !": Call StatusZeile(deb$)

While (SatzCount% <= MaxSendSatz%)
    deb$ = Right$("    " + Str$(MaxSendSatz% - SatzCount% + 1), 4) + " Sätze noch zu senden !"
    Call StatusZeile(deb$)
    
    iNAK% = 0
    If (Blockung%) Then
        SendStr$ = "": SatzSend% = 0
        Do
            j% = SatzCount% + SatzSend%
            If (j% > MaxSendSatz%) Then Exit Do
            If (Len(SendStr$ + SendSatz$(j%)) >= 256) Then Exit Do
        
            SendStr$ = SendStr$ + SendSatz$(j%) + RS$
            SatzSend% = SatzSend% + 1
        Loop
    Else
        SendStr$ = SendSatz$(SatzCount%) + RS$: SatzSend% = 1
    End If
    SendStr$ = SendStr$ + ETX$
    bcc$ = MakeBccA(SendStr$)
    SendStr$ = STX$ + SendStr$ + bcc$ + vbCr
    
    If (ackGH% = 0) Then
        sACK$ = ACK0$
    Else
        sACK$ = ACK1$
    End If
    
    ret% = TransmitBlock(SendStr$, RecStr$)
    While (RecStr$ <> sACK$) And (RecStr$ <> NAK$) And (ret%)
        ret% = ReceiveBlock%(RecStr$)
        
        'altes ACK ?
        If (ackGH% = 0) Then
            altACK$ = ACK1$
        Else
            altACK$ = ACK0$
        End If
        
        If (ret%) And (RecStr$ = altACK$) Then
            ret% = TransmitBlock(SendStr$, RecStr$)
        End If
    Wend
    If (RecStr$ = NAK$) Then iNAK% = True
    If (RecStr$ = sACK$) Then ackGH% = Not (ackGH%)
    
    If (iNAK%) Then
        FehlerNak% = FehlerNak% + 1
        If (NakCount% >= 2) And (Blockung%) Then
            Blockung% = False
        ElseIf (NakCount% < 3) Then
            NakCount% = NakCount% + 1
        Else
            ret% = False
        End If
    Else
        SatzCount% = SatzCount% + SatzSend%
        NakCount% = 0
    End If

    If (ret% = False) Then Call DefErrPop: Exit Function
Wend

deb$ = Right$("    " + Str$(MaxSendSatz% + 1), 4) + " Sätze gesendet                 "
Call StatusZeile(deb$)

ret% = TransmitBlock%(EOT$ + vbCr, RecStr$)
While (RecStr$ <> ENQ$) And (ret%)
    ret% = ReceiveBlock(RecStr$)
    'altes ACK ?
    If (ackGH% = 0) Then
        altACK$ = ACK1$
    Else
        altACK$ = ACK0$
    End If
    If (RecStr$ = altACK$) And (ret%) Then
        ret% = TransmitBlock%(EOT$ + vbCr, RecStr$)
    End If
Wend

SendBlocks% = ret%

Call DefErrPop
End Function
