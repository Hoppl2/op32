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

Public Lintech%, seriell%

Public TelGh$, TelApo$
Public GhIDF$, ApoIDF$

Public SendeStatusBereichInd%

Public AnzMinutenWarnung%
Public AnzMinutenWarten%

Public ModemOk%

Public SendeLog%

Dim sap As SerAuftrag
Dim sap1 As SerAuftrag

Dim ISDNRet$(10)
Dim ISDNStatus$(10)

Dim xParams$(6)
Dim xFilePara$

Dim SRT$()
Dim mnr$()
Dim mbm%()

Dim SendFile$, RecvFile$
Dim ZJOB%
Dim AktivSenden%
Dim MaxSendSatz%
Dim RecWarte%, OrgRecWarte%

Dim folgenr%

Public LiefName1$, LiefName2$, LiefName3$, LiefName4$
Dim AuftragsErg$, AuftragsArt$
Dim SendSatz$(100)
Dim HeuteDatStr$, xcBestDatum%

Dim LOGBUCH%, LEITUNGBUCH%

Private Const DefErrModul = "wsenden.bas"

Sub InitIsdnAnzeige()

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
End Sub

Sub InitBestellung()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("InitBestellung")
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
Dim i%, ok%, d%
Dim X$, h$, edat$

If (Lieferant% < 0) Then Call DefErrPop: Exit Sub

edat$ = Format(Day(Date), "00") + Format(Month(Date), "00") + Format(Year(Date), "0000")
X$ = Left$(edat$, 4) + Right$(edat$, 2)
'edat$ = Mid$(Date, 1, 3) + Mid$(Date, 4, 3) + Right$(Date, 4)
'x$ = Left$(edat$, 2) + Mid$(edat$, 4, 2) + Right$(edat$, 2)
HeuteDatStr$ = X$
xcBestDatum% = iDate(X$)

'ok% = TestModemPar%
'If (ok% = False) Then Exit Sub

Call HoleLieferantenDaten

Call SucheSendeArtikel

AktivSenden% = True


With frmSenden
    .Caption = "Sendeauftrag für " + LiefName1$
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
Dim ok%, suche2%
Dim X$, h$

'def SEG = &H40
'  com1% = 256! * PEEK(1) + PEEK(0): com2% = 256! * PEEK(3) + PEEK(2)
'  LPT1% = 256! * PEEK(9) + PEEK(8): LPT2% = 256! * PEEK(11) + PEEK(10)
'def SEG
'Port% = LPT1%

If (para.isdn) Then
    Lintech% = False
    ok% = ModemParameter%("LO-ISDN")
    If (ok% = False) Then
        ok% = ModemParameter%("LINTECH")
        If (ok%) Then Lintech% = True
    End If
    If (ok% = False) Then
        para.isdn = False
    Else
        TestModemPar% = True
        Exit Function
    End If
End If

seriell% = True
ok% = ModemParameter%("MODEM-S", True)
If (ok% = False) Then
    suche2% = True
    seriell% = False
    ok% = ModemParameter%("MODEM-P")
End If
If (ok% = False) Then
    If (suche2% = True) Then seriell% = True
    If (seriell%) Then X$ = "Seriell" Else X$ = "Parallel"
    X$ = X$ + "modem-Geräteparameter nicht vorhanden!"
    Call MsgBox(X$, vbCritical)
    TestModemPar% = False
    Exit Function
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

TestModemPar% = True

End Function

Function ModemParameter%(TestName$, Optional Parameter% = False)
Dim i%, ok%, xq%, xParam%, DPA%, aa%
Dim s$, X$, xGeraet$, xFileName$
Dim XparamFix As String * 128
Dim DparamFix As String * 40

s$ = "\user\xparam" + para.User + ".dat"
xParam% = FileOpen(s$, "RW", "R", Len(XparamFix))

'1 * belegt sonst nicht belegt, 8 Name,30 Schnittstelle incl.params
i% = 1: ok% = False
While (i% <= 9) And (ok% = 0)
    Get xParam%, i% + 1, XparamFix
    If (Left$(XparamFix$, 1) = "*") Then
        X$ = Mid$(XparamFix$, 2, 8): X$ = LTrim$(X$): X$ = RTrim$(X$): xGeraet$ = X$
        xFileName$ = "\user\" + X$ + ".dpa"
        X$ = Mid$(XparamFix$, 10, 30): X$ = LTrim$(X$): X$ = RTrim$(X$): xFilePara$ = X$
        If (Left$(xGeraet$, Len(TestName$)) = TestName$) Then ok% = True
    End If
    i% = i% + 1
Wend
Close #xParam%

If (ok% And Parameter%) Then
    DPA% = FileOpen(xFileName$, "RW", "R", Len(DparamFix))
    For xq% = 1 To 5
        Get #DPA%, xq%, DparamFix: X$ = DparamFix
        X$ = RTrim$(X$): xParams$(xq%) = X$
    Next xq%
    Close #DPA%
End If

X$ = Mid$(xParams$(5), 21): X$ = LTrim$(X$): X$ = RTrim$(X$): xParams$(6) = X$
X$ = Left$(xParams$(5), 20): X$ = LTrim$(X$): X$ = RTrim$(X$): xParams$(5) = X$

ModemParameter% = ok%

End Function

Sub SucheSendeArtikel()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("SucheSendeArtikel")
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
Dim i%, ind%, Max%
Dim h$

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

For i% = 0 To (AnzBestellArtikel% - 1)
    frmAction!lstSortierung.ListIndex = i%
    h$ = RTrim$(frmAction!lstSortierung.text)
    SRT$(i%) = Left$(h$, 29)
    mnr$(i%) = Mid$(h$, 30, 7)
    mbm%(i%) = Val(Mid$(h$, 37, 4))
Next i%

Call DefErrPop
End Sub

Sub SendAutomatic()
Dim i%, Max%, ret%
Dim d%
Dim X$, edat$

edat$ = Format(Day(Date), "00") + Format(Month(Date), "00") + Format(Year(Date), "0000")
X$ = Left$(edat$, 4) + Right$(edat$, 2)
HeuteDatStr$ = X$
xcBestDatum% = iDate(X$)

Call HoleLieferantenDaten

Call SucheSendeArtikel

AuftragsErg$ = Left$(RTrim$(Rufzeiten(AutomaticInd%).AuftragsErg) + Space$(2), 2)
AuftragsArt$ = Left$(RTrim$(Rufzeiten(AutomaticInd%).AuftragsArt) + Space$(2), 2)
If (Rufzeiten(AutomaticInd%).Aktiv = "J") Then
    AktivSenden% = True
Else
    AktivSenden% = False
End If

BestSendenAbbruch% = False

Call SaetzeVorbereiten

ret% = SaetzeSenden%
    
If (ret% = True) Then
    If (MaxSendSatz% > 0) Then
        Call AbsagenEntfernen
    End If
'    Call UeberleitenInWU
End If
    
'Call EntferneAktivKz(Lieferant%)
Call UpdateBekartDat(Lieferant%, ret%)

End Sub

Sub SendGermany()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("SendGermany")
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
Dim ret%, ind%, Ueberleiten%, RueckmeldungenFlag%, AbsagenFlag%
Dim h$


With frmSenden
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
    
    If (.optAuftrag(0).Value = True) Then
        AktivSenden% = True
    Else
        AktivSenden% = False
    End If
    RueckmeldungenFlag% = .chkAuftrag(0).Value
    AbsagenFlag% = .chkAuftrag(1).Value
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
        
Ueberleiten% = False
If (ret% = True) Then
    ret% = MsgBox("Bestellung korrekt empfangen ?", vbYesNo Or vbDefaultButton1)
    If (ret% = vbYes) Then
        Ueberleiten% = True
    End If
Else
    ret% = MsgBox("Datenfernübertragung wiederholen ?", vbYesNo Or vbDefaultButton1)
    If (ret% <> vbYes) Then
        Ueberleiten% = True
    End If
End If

If (Ueberleiten%) Then
'    Call UeberleitenInWU
'    Call EntferneAktivKz(Lieferant%)
    Call UpdateBekartDat(Lieferant%, Ueberleiten%)
    Unload frmSenden
Else
    Call frmSenden.SetFormModus(0)
End If

Call DefErrPop
End Sub

Sub SaetzeVorbereiten()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("SaetzeVorbereiten")
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
Dim i%, j%, k%, gesendet%, menge%, IstFrei%
Dim satz$, char$, fistsatz$, BitFeld$, TestPZN$, sMenge$

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

satz$ = "06256            " + para.FISTAM(0) + "   " + para.FISTAM(1)
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
      bek.GetRecord (Val(Right$(SRT$(i), 4)) + 1)
      satz$ = "030481aa  ####ppppdd" + Left$(bek.txt, 26) + "ää"
      '        12345678901234567890123
      '             ^     Auftragsnummer
      '                ^^ Artikelbezogener Hinweis
      Mid$(satz$, 7, 2) = AuftragsArt$
      Mid$(satz$, 11, 4) = sMenge$
      Mid$(satz$, 15, 4) = Mid$(bek.txt, 30, 4) 'Packungsgröße
      Mid$(satz$, 19, 2) = Mid$(bek.txt, 34, 2) 'Darreichungsform
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
Dim j%, sFile%, ret%, erg%
Dim l$, h2$

SendeLog% = True

If (SendeLog%) Then
    erg% = CreateDirectory%("winw")
    h2$ = "winw\" + Format(Year(Date) Mod 100, "00") + Format(Month(Date), "00") + Format(Day(Date), "00")
    erg% = CreateDirectory%(h2$)
    h2$ = h2$ + "\" + Format(Lieferant%, "000") + Format(Now, "HHMM")
    If (AutomaticSend% = False) Then
        h2$ = h2$ + "m"
    End If
    LOGBUCH% = FreeFile
    Open h2$ + ".log" For Output As LOGBUCH%
    For j% = 0 To MaxSendSatz%
      Print #LOGBUCH%, ">" + SendSatz$(j%)
    Next j%
    
    LEITUNGBUCH% = FreeFile
    Open h2$ + ".deb" For Output As LEITUNGBUCH%
End If

If (para.isdn) Then
    SendFile$ = LINTECHPFAD$ + "S_" + Right$(String$(5, 48) + Mid$(Str$(Lieferant%), 2), 5) + "1" + ".I01"
    RecvFile$ = LINTECHPFAD$ + "S_" + Right$(String$(5, 48) + Mid$(Str$(Lieferant%), 2), 5) + "1" + ".RUK"

    sFile% = FreeFile
    Open SendFile$ For Output As #sFile%
    For j% = 0 To MaxSendSatz%
      Print #sFile%, SendSatz$(j%) + Chr$(10);
    Next j%
    Close #sFile%
    
    '  GoSub OpenDeb
    ret% = ISDNBestellung%
    '  Close #f.deb%
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
    Close (LEITUNGBUCH%)
End If

SaetzeSenden% = ret%
Call DefErrPop
End Function

Function ISDNBestellung%()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("ISDNBestellung%")
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
        Call MsgBox(MeldungStr$, vbCritical)
    Else
        Call StatusZeile(MeldungStr$)
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
Dim ret%, ret2%
Dim SendStr$, RecStr$

ret% = SeriellBest%

If (ret% = True) Then
    ret% = SeriellRueckMeldungen%
    SendStr$ = trennStr$
    ret2% = SeriellSend%(SendStr$, RecStr$)
End If

SeriellBestellung% = ret%

End Function

Function SeriellBest%()
Dim j%, ret%, Warte%, OrgWarte%
Dim Status$, SendStr$, RecStr$, satz01$

SeriellBest% = False

RecWarte% = 120: OrgRecWarte% = RecWarte%

Status$ = "Warten auf Modembereitschaft ...": Call StatusZeile(Status$)
ret% = SeriellReceive%(RecStr$)
If (ret% = False) Then Exit Function
'
Status$ = "Modem-Initialisierung wird durchgeführt.": Call StatusZeile(Status$)
SendStr$ = ModemSetup$
ret% = SeriellSend%(SendStr$, RecStr$)
If (ret% = False) Then Exit Function

If (AktivSenden%) Then
    Status$ = "Telefonnummer " + TelGh$ + " wird gewählt.": Call StatusZeile(Status$)
    folgenr% = 0
    SendStr$ = ModemDialP$ + TelGh$ + ModemDialS$
    ret% = SeriellSend%(SendStr$, RecStr$)
    If (ret% = False) Then Exit Function
    SendStr$ = empfange$
    ret% = SeriellSend%(SendStr$, RecStr$)
    satz01$ = RecStr$
    If (ret% = False) Then Exit Function
Else
    If (AutomaticSend%) Then
        Status$ = "Warten auf Anruf (max." + Str$(AnzMinutenWarten%) + " Minuten) ..."
    Else
        Status$ = "Warten auf Anruf (max. 6 Stunden) ..."
    End If
    Call StatusZeile(Status$)
    folgenr% = 0
    SendStr$ = ModemAnswer$
    If (AutomaticSend%) Then
        RecWarte% = AnzMinutenWarten% * 60
    Else
        RecWarte% = 6 * 60 * 60
    End If
    ret% = SeriellSend%(SendStr$, RecStr$)
    RecWarte% = OrgRecWarte%
    If (ret% = False) Then Exit Function
    SendStr$ = empfange$
    ret% = SeriellSend%(SendStr$, RecStr$)
    satz01$ = RecStr$
    If (ret% = False) Then Exit Function
End If

If (Mid$(RecStr$, 9, 1) <> "3") Then
    Status$ = "Teilnehmer antwortet nicht !": Call StatusZeile(Status$)
    Exit Function
End If

If (Mid$(satz01$, 20, 7) <> GhIDF$) Then
    Status$ = "falsche Großhandlung " + Mid$(RecStr$, 20, 7) + "!": Call StatusZeile(Status$)
    Exit Function
End If

Status$ = "Bestellung wird gesendet ...": Call StatusZeile(Status$)
Mid$(SendSatz$(0), Len(SendSatz$(0)) - 3, 4) = Mid$(satz01$, Len(satz01$) - 4, 4)
For j% = 0 To MaxSendSatz%
    SendStr$ = "[00000D=00" + SendSatz$(j%) + "]"
    ret% = SeriellSend%(SendStr$, RecStr$)
    If (ret% = False) Then Exit Function
Next j%

SeriellBest% = True

End Function

Function SeriellSend%(SendStr$, RecStr$)
Dim ret%
Dim l$, deb$

Mid$(SendStr$, 6, 1) = Right$(Str$(folgenr%), 1)
l$ = Mid$(Str$(Len(SendStr$) - 2), 2)
l$ = Right$("0000" + l$, 4)
Mid$(SendStr$, 2, 4) = l$
Call MakeBcc(SendStr$)

frmSenden.comSenden.Output = SendStr$
deb$ = "> " + SendStr$: Call StatusZeile(deb$)

folgenr% = folgenr% + 1

ret% = SeriellReceive%(RecStr$)
deb$ = "< " + RecStr$: Call StatusZeile(deb$)

SeriellSend% = ret%

End Function

Function SeriellReceive%(RecStr$)
Dim ok%, l%
Dim char$, bcc$, deb$
Dim timeranf
Dim vchar As Variant
Dim bchar() As Byte

SeriellReceive% = False

RecStr$ = ""
ok% = 0
timeranf = Timer
While (ok% = 0)
    
    If (Timer - timeranf > RecWarte%) Then ok% = 2
    
    l% = frmSenden.comSenden.InBufferCount
    If (l% > 0) Then
        char$ = frmSenden.comSenden.Input
        RecStr$ = RecStr$ + char$
        If (char$ = "]") Then ok% = 1
    Else
        DoEvents
        If (BestSendenAbbruch% = True) Then
            ok% = 2
        End If
    End If
Wend
If (ok% = 1) Then
    char$ = frmSenden.comSenden.Input
    bcc$ = char$
    SeriellReceive% = True
End If

End Function

Sub MakeBcc(SendStr$)
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

End Sub

Function ISDNBest%(sta%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("ISDNBest%")
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
Dim hSendFile$, hRecvFile$
Dim hTelGH$, hTelApo$
Dim hGhIDF$, hApoIDF$, ae$, satz$, modus$, Status$, sFile$, rFile$, TGH$, TApo$, IGH$, IApo$, h2$
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
Status$ = Time$ + " Modus = " + modus$
Call StatusZeile(Status$)

'Einfügen
Call WriteZjob

'    h2$ = Format$(Now, "HH:MM:SS")
'    StartSek& = Val(Left$(h2$, 2)) * 3600 + Val(Mid$(h2$, 4, 2)) * 60 + Val(Right$(h2$, 2))
StartSek = Timer

While (ret% = 0)  'solange nicht übertragen

    DoEvents
    If (BestSendenAbbruch% = True) Then
        ret% = -1
    End If

    If (ret% = 0) Then
        ret% = CheckZjob%
    End If

    If (ret% = 0) And (AktivSenden% = False) Then
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

Function CheckFertig%()
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

End Function

Sub DisplayStatus(Funkt%, ret%)
Static AltFunkt%
Dim anzeigen%
Dim Status$

Status$ = ""

anzeigen% = False
If (sap1.sta <> sap.sta) Then anzeigen% = True
If (sap1.mds <> sap.mds) Then anzeigen% = True
If (AltFunkt% <> Funkt%) Then anzeigen% = True
If (anzeigen%) Then
    AltFunkt% = Funkt%
    sap1 = sap
    Status$ = Status$ + Time$ + " "
    If (Funkt% = 1) Then
        Status$ = Status$ + "Auftrag vorbereiten "
    ElseIf (Funkt% = 2) Then
        Status$ = Status$ + "Auftrag löschen     "
    ElseIf (Funkt% = 5) Then
        Status$ = Status$ + "Auftrag freigeben   "
    ElseIf (Funkt% = 11) Then
        Status$ = Status$ + "Auftrag wiederholen "
    ElseIf (Funkt% = 12) Then
        Status$ = Status$ + "Statusabfrage       "
    Else
        Status$ = Status$ + "Funkt:" + Str$(Funkt%) + " "
    End If
    If (sap.sta <= 9) Then
        Status$ = Status$ + LTrim$(ISDNStatus$(sap.sta)) + " "
    ElseIf (sap.sta = 10) Then
        Status$ = Status$ + "Senden aktiv "
    Else
        Status$ = Status$ + "Status unbekannt (" + Mid$(Str$(sap.sta), 2) + ") "
    End If
    If (Funkt% = 1) Then
        Status$ = Status$ + "Modus="
        If (sap.mds = 1) Then
            Status$ = Status$ + "Anruf durchführen "
        ElseIf (sap.mds = 2) Then
            Status$ = Status$ + "Warten auf Anruf  "
        ElseIf (sap.mds = 9) Then
            Status$ = Status$ + "Gesperrt          "
        Else
            Status$ = Status$ + "Unbekannt :" + Str$(sap.mds)
        End If
    End If
    Status$ = Status$ + "(" + Mid$(Str$(ret%), 2) + ") "
    Call StatusZeile(Status$)
End If

End Sub

Sub WriteZjob()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("WriteZjob")
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
Dim ret%, ind%
Dim stat$, Status$

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
    Status$ = Time$ + " Status = " + stat$
    Call StatusZeile(Status$)
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

Sub ISDNRueckMeldungen()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("ISDNRueckMeldungen")
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
Dim i%, j%, rFile%, SatzLen%, ende%, satz%, fehler%
Dim recv$, Oldrecv$, satz1$

MaxSendSatz% = 0
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
    SatzLen% = Val(Mid$(recv$, 3, 3))
    satz1$ = Left$(recv$, SatzLen%)
    If (MaxSendSatz% < 700) Then
        MaxSendSatz% = MaxSendSatz% + 1
        SendSatz$(MaxSendSatz%) = satz1$
    End If
    Oldrecv$ = recv$      '1.84 bei Lintech-ISDN-Server werden Rückmelde-
                          'sätze nur durch CHR(10) getrennt --> möglicherweise
                          'hängen Rückmeldungen aneinander
    recv$ = ""
    If (Len(Oldrecv$) > (SatzLen% + 5)) Then
        Oldrecv$ = Mid$(Oldrecv$, SatzLen% + 1)
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
Dim j%, ret%, ende%, telok%, SatzLen%
Dim satz$, satz1$, Status$, SendStr$, RecStr$

SeriellRueckMeldungen% = False

Status$ = "Empfang der Rückmeldungen.": Call StatusZeile(Status$)

MaxSendSatz% = 0
SendSatz$(0) = Space$(14) + "DATENÜBERTRAGUNG RÜCKMELDUNGEN: "

ende% = 0

While Not (ende%)
    SendStr$ = empfange$
    ret% = SeriellSend%(SendStr$, RecStr$)
    If (ret% = False) Then Exit Function
  
    
    telok% = (Mid$(RecStr$, 9, 2) = "30")
    satz$ = Mid$(RecStr$, 15)
    satz$ = Left$(satz$, Len(satz$) - 1)
    
    If (telok% = 0) Then Exit Function
    
    DoEvents
    If (BestSendenAbbruch% = True) Then Exit Function
    
    While (satz$ <> "")
        SatzLen% = Val(Mid$(satz$, 3, 3))
        satz1$ = Left$(satz$, SatzLen%)
        satz$ = Mid$(satz$, SatzLen% + 1)
        If (Left$(satz$, 1) = Chr$(10)) Then satz$ = Mid$(satz$, 2)
        If (Mid$(satz1$, 1, 2) = "99") Or (Mid$(RecStr$, 9, 1) <> "3") Then
            ende% = True
        Else
            If (MaxSendSatz% < 700) Then
                MaxSendSatz% = MaxSendSatz% + 1
                SendSatz$(MaxSendSatz%) = satz1$
            End If
        End If
    Wend
Wend

If (SendeLog%) Then
    For j% = 1 To MaxSendSatz%
      Print #LOGBUCH%, "<" + SendSatz$(j%)
    Next j%
End If
    

SeriellRueckMeldungen% = True

End Function

Sub RueckmeldungenBefuellen()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("RueckmeldungenBefuellen")
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
Dim i%
Dim h$
        
With frmRueckmeldungen.flxRueck
    If (MaxSendSatz% > 1) Then
        For i% = 1 To MaxSendSatz%
            h$ = Trim$(SendSatz$(i%))
            .AddItem Left$(h$, 2) + vbTab + Mid$(h$, 3, 3) + vbTab + Mid$(h$, 6)
        Next i%
    Else
        .AddItem "" + vbTab + "" + vbTab + "keine Rückmeldungen empfangen !"
    End If
End With

Call DefErrPop
End Sub

Sub ZeigeRueckmeldungen()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("ZeigeRueckmeldungen")
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
Dim i%
Dim h$
        
frmRueckmeldungen.Show 1

Call DefErrPop
End Sub

Sub AbsagenEntfernen()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("AbsagenEntfernen")
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
Dim BABSAGE%, i%, d%, satz1%, satztyp%, menge%, SN%, babMax%
Dim h$, X$, pzn$, xc$

absagen.OpenDatei
absagen.GetRecord (1)
babMax% = absagen.erstmax

For satz1% = 1 To MaxSendSatz%
    X$ = SendSatz$(satz1%)
    X$ = Left$(X$, Len(X$) - 1)
    satztyp% = Val(Mid$(X$, 1, 2))
    pzn$ = Mid$(X$, 22, 7)
    menge% = Val(Mid$(X$, 31, 4))
    If (satztyp% = 4) Then
        For i% = 0 To (AnzBestellArtikel% - 1)
            SN% = Val(Right$(SRT$(i%), 4))
            If (SN% > 0) Then                 '1.88
                bek.GetRecord (SN% + 1)
                If (bek.pzn = pzn$) And (menge% = Abs(bek.bm)) And (bek.zugeordnet = "J") Then
                    
                    'Absage-Datei ------------------------------
                    absagen.datum = xcBestDatum%
                    absagen.lief = Lieferant%
                    absagen.pzn = bek.pzn
                    absagen.filler = " "
                    absagen.text = Left$(bek.txt, 33) + " " + Mid$(bek.txt, 34, 2)
                    absagen.menge = bek.bm
                    absagen.rest = String$(absagen.DateiLen, 0)
                    
                    babMax% = babMax% + 1
                    If (babMax% > 2000) Then babMax% = 1
                    absagen.PutRecord (babMax% + 1)
                    
                    absagen.GetRecord (1)
                    absagen.erstmax = babMax%
                    absagen.PutRecord (1)
                    
                    'Bestell-Datei -------------------------------------------
                    bek.lief = 0
                    bek.best = " "
                    bek.zugeordnet = "N"
                    bek.absage = Lieferant%
                    bek.PutRecord (SN% + 1)
                    SRT$(i%) = String$(Len(SRT$(i%)), 48)
                    Exit For
                End If
            End If
        Next i%
    End If
Next satz1%

absagen.CloseDatei

Call DefErrPop
End Sub

Sub StatusZeile(h$)

'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("StatusZeile")
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

With frmSenden.flxAuftrag
    .AddItem h$
    If (.Rows > 4) Then
        .TopRow = .Rows - 4
        If (.FixedRows > 0) Then
            .TopRow = .TopRow + .FixedRows
        End If
    End If
End With

If (SendeLog%) Then
    Print #LEITUNGBUCH%, Format(Now, "HHMMSS ") + h$
End If

DoEvents
Call DefErrPop
End Sub

Sub TestBestellung()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("TestBestellung")
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

Sub UpdateBekartDat(lief%, Ueberleiten%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("UpdateBekartDat")
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
Dim i%, j%, bArtikel%, SN%, bm%, nm%, mm%, ppMax%, ppGes%, bMax%, pp%, ssatz%, Max%, ok%
Dim bzeit$, shdatum$, shzeit$, shlieferant$, pzn$, OrgPzn$, text$

bzeit$ = MKI(Val(Left$(Time$, 2)) * 100 + Val(Mid$(Time$, 4, 2)))

'shdatum$ = Mid$(dat$, 5, 2) + "." + Mid$(dat$, 3, 2) + "." + Mid$(dat$, 1, 2)
'shzeit$ = Right$(str$(Int(CVI(bzeit$) / 100)), 2) + ":" + Right$("00" + Mid$(str$(CVI(bzeit$) Mod 100), 2), 2)
'shlieferant$ = Right$("   " + str$(Lieferant%), 3)

'nm.DR% = 0: KEIN.nr% = 0: s.AEP# = 0#: s.AVP# = 0#

bek.SatzLock (1)

If (Ueberleiten%) Then
    Call etik.SatzLock(1)
    Call wu.SatzLock(1)
    
    wu.GetRecord (1)
    bMax% = wu.erstmax
    
    Call etik.GetRecord(1)
    ppMax% = etik.erstmax
    ppGes% = etik.erstges
    
    bArtikel% = 0
End If
    
bek.GetRecord (1)
Max% = bek.erstmax
j% = 0
For i% = 1 To Max%
    bek.GetRecord (i% + 1)
    
    ok% = False
    If (lief% = -1) Or (bek.aktivlief = lief%) Then ok% = True
    
    If (Ueberleiten%) Then
        If (ok%) Then
            bm% = bek.bm: nm% = bek.nm: pzn$ = bek.pzn
            If Left$(pzn$, 4) = "9999" And Mid$(pzn$, 5, 3) <> "999" Then
            Else
                ssatz% = 0
                OrgPzn$ = pzn$
                FabsErrf% = ass.IndexSearch(0, pzn$, FabsRecno&)
                If (FabsErrf% = 0) Then
                    ssatz% = FabsRecno&
                    ass.GetRecord (FabsRecno& + 1)
                    mm% = ass.mm
                Else
                    mm% = 0
                End If
                text$ = pzn$ + " " + Left$(bek.txt, 33) + " " + Mid$(bek.txt, 34, 2)
                If (bm% = 0) And (nm% = 0) Then     '1.60 Merkzettel
                    Call Merkzettel(pzn$, Mid$(text$, 9), xcBestDatum%, xcBestDatum%, Lieferant%, 0)
                Else
                    wu.stat = "J"
                    If (bm% < 0) Or (nm% < 0) Then wu.stat = "N"
                    wu.at = text$
                    wu.AEP = bek.AEP
                    wu.Kz = "B"
                    wu.bd = HeuteDatStr$
                    wu.AVP = bek.AVP
                    wu.li = Chr$(Lieferant%)
                    wu.bm = Abs(bm%)
                    wu.nm = Abs(nm%)
                    wu.ac = bek.abl
                    wu.rm = Abs(bm%)
                    wu.lm = Abs(bm%) + Abs(nm%)
                    wu.am = 0
                    wu.asatz = bek.asatz
                    wu.abl = Space$(6)
                    wu.zeit = bzeit$
                    wu.ssatz = bek.ssatz
                    wu.la = nm%
                    wu.km = bek.km
                    wu.wg = bek.wg
                    
                    If (bek.auto = "v") Then
                        wu.besorger = "B"
                    Else
                        wu.besorger = " "
                    End If
                    wu.AbholNr = bek.AbholNr
                    wu.nnart = bek.nnart
                    wu.NNAep = bek.NNAep
                    
                    bMax% = bMax% + 1
                    wu.PutRecord (bMax% + 1)
                    
                    bArtikel% = bArtikel% + 1
                End If
              
                pp% = 0
                If (ssatz% <> 0) Then If (InStr("AFP", ass.pp) <> 0) Then pp% = True
                If (InStr(para.PosAktivWG, bek.wg) <> 0) Then pp% = True
                If (pp%) And (OrgPzn$ <> "9999999") Then
                    If (bek.auto <> "v") Then
                        etik.pzn = OrgPzn$
                        etik.menge = Abs(bm%) + Abs(nm%)
                        etik.druck = " "
                        etik.zusatz = "  "
                        etik.asatz = bek.asatz
                        etik.lief = Lieferant%
                        etik.datum = xcBestDatum%
                        etik.knr = String(2, 0)
                        etik.rest = String(44, 0)
                        ppMax% = ppMax% + 1
                        etik.PutRecord (ppMax% + 1)
                        ppGes% = ppGes% + Abs(bm%) + Abs(nm%)
                    End If
                End If
            End If
          
        Else
            j% = j% + 1
            bek.PutRecord (j% + 1)
        End If
    ElseIf (ok%) Then
        bek.aktivlief = 0
        bek.aktivind = 0
        bek.PutRecord (i% + 1)
    End If
Next i%

If (Ueberleiten%) Then
    Call etik.GetRecord(1)
    etik.erstmax = ppMax%
    etik.erstges = ppGes%
    Call etik.PutRecord(1)
    
    Call wu.GetRecord(1)
    wu.erstmax = bMax%
    Call wu.PutRecord(1)
    
    Call etik.SatzUnLock(1)
    Call wu.SatzUnLock(1)
    
    bek.erstmax = j%
    bek.PutRecord (1)
End If

bek.erstcounter = (bek.erstcounter + 1) Mod 100
bek.PutRecord (1)
bek.SatzUnLock (1)

Call frmAction.AuslesenBestellung(True, False, True)

'Call WProtokoll(B.Max%, "BESTK2 nach Überleitung" + str$(bArtikel%) + " Artikel")
'Close #f.pp%
'GoTo ProgrammEnde

Call DefErrPop
End Sub

Sub Merkzettel(pzn$, text$, dat%, dt%, lief%, lm%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("Merkzettel")
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
Dim STAMMLOS%, Version%
Dim MzMax&
Dim satz!, xrec!, xrec1!
Dim h$, feld$, match$, Cmnd$


'besorgt.GetRecord (1)
'Version% = Asc(Mid$(StammlosFix$, 4, 1))

BESORGT.pzn = pzn$
BESORGT.Flag = " "
BESORGT.text = text$
BESORGT.f1 = " "
BESORGT.dat = dat%
BESORGT.Leer = String$(4, 0)
BESORGT.f2 = " "
BESORGT.dt = dt%
BESORGT.leer1 = String$(4, 0)
BESORGT.lief = lief%
BESORGT.lm = lm%
BESORGT.rest = Chr$(13) + Chr$(10)


match$ = Left$(BESORGT.text, 6) + BESORGT.pzn + MKDate(BESORGT.dat)

FabsErrf% = BESORGT.IndexInsert(0, BESORGT.pzn, FabsRecno&)
xrec! = FabsRecno&

FabsErrf% = BESORGT.IndexInsert(1, match$, FabsRecno&)
xrec1! = FabsRecno&

If (xrec! <> xrec1!) Then
    Call MsgBox("Unterschiedliche Records im Merkzettel!", vbOKOnly Or vbCritical, "FABSP")
End If

satz! = xrec!

BESORGT.PutRecord (satz! + 1)

BESORGT.SatzLock (1)
BESORGT.GetRecord (1)
'LockError! = 0
'MzMax& = FNSATZ%(CVI%(Left$(StammlosFix$, 2)))
MzMax& = FNSATZ%(BESORGT.erstmax)
If (xrec! > MzMax&) Then
    MzMax& = xrec!
    If (MzMax& > 65535) Then MzMax& = 65535
    BESORGT.erstmax = fnint%(MzMax&)
'    StammlosFix$ = MKI$(fnint%(MzMax&)) + "*" + Mid$(StammlosFix$, 4)
    BESORGT.PutRecord (1)
End If
BESORGT.SatzUnLock (1)
'Close #STAMMLOS%

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

Function fnint%(ByVal X!)

'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("fnint%")
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
fnint% = Int(X! + (X! > 32767) * 65536)

Call DefErrPop
End Function

Function FNSATZ%(X%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("FNSATZ%")
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
 
FNSATZ% = X% - (X% < 0) * 65536!
 
Call DefErrPop
End Function

Sub DruckeBestellung()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("DruckeBestellung")
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
Dim i%, j%, pos%, sp%(9), SN%, Y%, Max%, ind%
Dim EK#
Dim header$, tx$, h$

Call StartAnimation("Ausdruck wird erstellt ...")




bek.SatzLock (1)
bek.GetRecord (1)
Max% = bek.erstmax

frmAction.lstSortierung.Clear

For i% = 1 To Max%
    bek.GetRecord (i% + 1)
'For i% = 0 To (AnzBestellArtikel% - 1)
'    SN% = Val(Right$(SRT$(i%), 4))
'    If (SN% > 0) Then                 '1.88
'        bek.GetRecord (SN% + 1)
        If ((Asc(bek.pzn) < 128) And (bek.zugeordnet = "J")) Then
'            If (bek.zukontrollieren = "N") Or (bek.MussKontrollieren = "N") Or (bek.nochzukontrollieren = "N") Then
            If (bek.zukontrollieren <> "1") Then
                EK# = bek.AEP
    
                                
                h$ = Left$(bek.txt, 28) + vbTab + Mid$(bek.txt, 29, 5)
                h$ = h$ + vbTab + Mid$(bek.txt, 34, 2)

                tx$ = Format(Abs(bek.bm), "0")
                h$ = h$ + vbTab + tx$
                
                tx$ = ""
                If (Abs(bek.nm > 0)) Then
                    tx$ = Format(Abs(bek.nm), "0")
                End If
                h$ = h$ + vbTab + tx$
                
                tx$ = Format(EK# * Abs(bek.bm), "0.00")
                h$ = h$ + vbTab + tx$
                h$ = h$ + vbTab + bek.pzn + vbTab
                
                frmAction.lstSortierung.AddItem h$
            End If
        End If
'    End If
Next i%

bek.SatzUnLock (1)



Printer.ScaleMode = vbTwips
Printer.Font.Name = "Arial"

DruckSeite% = 0
    
header$ = frmAction.lblArbeit(0).Caption
    
Call DruckKopf(header$, "")
        
sp%(0) = Printer.TextWidth(String$(9, "9"))
sp%(1) = sp%(0) + Printer.TextWidth(String$(28, "X"))
sp%(2) = sp%(1) + Printer.TextWidth(String$(6, "X"))
sp%(3) = sp%(2) + Printer.TextWidth(String$(5, "X"))
sp%(4) = sp%(3) + Printer.TextWidth(String$(5, "9"))
sp%(5) = sp%(4) + Printer.TextWidth(String$(5, "9"))
sp%(6) = sp%(5) + Printer.TextWidth(String$(12, "9"))
        
Printer.CurrentX = 0
Printer.Print "P Z N";
Printer.CurrentX = sp%(0)
Printer.Print "A R T I K E L";
'tx$ = "Menge"
'Printer.CurrentX = sp%(2) - Printer.TextWidth(tx$)
'Printer.Print tx$;
'tx$ = "Me"
'Printer.CurrentX = sp%(3) - Printer.TextWidth(tx$)
'Printer.Print tx$;
tx$ = "B M"
Printer.CurrentX = sp%(4) - Printer.TextWidth(tx$)
Printer.Print tx$;
tx$ = "N M"
Printer.CurrentX = sp%(5) - Printer.TextWidth(tx$)
Printer.Print tx$;
tx$ = "Zeilenwert"
Printer.CurrentX = sp%(6) - Printer.TextWidth(tx$)
Printer.Print tx$;
Printer.Print " "
Y% = Printer.CurrentY
Printer.Line (0, Y%)-(sp%(6), Y%)
                
    
With frmAction.lstSortierung
    AnzBestellArtikel% = .ListCount
    For i% = 1 To AnzBestellArtikel%
        .ListIndex = i% - 1
        h$ = .text
        
        
        For j% = 0 To 6
            ind% = InStr(h$, vbTab)
            tx$ = Left$(h$, ind% - 1)
            h$ = Mid$(h$, ind% + 1)
            
            If (j% = 6) Then
                Printer.CurrentX = 0
            ElseIf (j% = 0) Or (j% = 2) Then
                Printer.CurrentX = sp%(j%)
            Else
                Printer.CurrentX = sp%(j% + 1) - Printer.TextWidth(tx$ + "x")
            End If
            
            Printer.Print tx$;
        Next j%
        
        Printer.Print " "
        
        If (Printer.CurrentY > Printer.ScaleHeight - 1000) Then
            Call DruckFuss
            Call DruckKopf(header$, "")
        End If
    Next i%
End With
    
Y% = Printer.CurrentY
Printer.Line (0, Y%)-(sp%(6), Y%)
Printer.CurrentX = sp%(0)
Printer.Print Format(AnzBestellArtikel%, "0") + " Positionen";
tx$ = Format(MarkWert#, "0.00")
Printer.CurrentX = sp%(6) - Printer.TextWidth(tx$ + "x")
Printer.Print tx$;

Printer.Print " "

Call DruckFuss(False)

Printer.EndDoc

Call StopAnimation

Call DefErrPop
End Sub

Function IsdnSer%(Funkt%, ind%, sap As SerAuftrag)

End Function

Function ModemAktivieren%()
Dim ret%, fehler%, CommPort%, ind%
Dim h$, Settings$

Call DefErrFnc("ModemAktivieren")
On Error GoTo DefErr

ret% = True

If (seriell%) Then
    fehler% = 0
    ' COM einsetzen.
    ind% = InStr(xFilePara$, "COM")
    If (ind% > 0) Then
        CommPort% = Val(Mid$(xFilePara$, ind% + 3, 1))
        Settings$ = Mid$(xFilePara$, ind% + 5)
    End If
    frmSenden.comSenden.CommPort = CommPort%
    ' 9600 Baud, keine Parität, 8 Datenbits und 1 Stopbit
    frmSenden.comSenden.Settings = Settings$
    frmSenden.comSenden.InputMode = comInputModeText
    '    frmAction.comSenden.InputMode = comInputModeBinary
    frmSenden.comSenden.Handshaking = comRTSXOnXOff
    ' Steuerelement anweisen, daß es den gesamten
    ' Pufferinhalt lesen soll, wenn die Input-Eigenschaft
    ' verwendet wird.
    frmSenden.comSenden.InputLen = 1
    
    ' _Anschluß öffnen.
    On Error GoTo ErrorHandler
    frmSenden.comSenden.PortOpen = True
    
    'If (fehler% = 1) Then ret% = False
    If (fehler%) Then ret% = False
End If

ModemAktivieren% = ret%
'Open xFilePara$ For Random As #cm%

Call DefErrPop
Exit Function

ErrorHandler:
    fehler% = Err
    Err = 0
    Resume Next
    Return

'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
DefErr:
Call DefErrAnswer2(Err.Source, Err.Number, Err.Description, DefErrModul)
fehler% = 1
Resume Next
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
End Function

Sub HoleLieferantenDaten()
Dim i%, ok%
Dim h2$, X$, BetrNr$

lif.GetRecord (Lieferant% + 1)
h2$ = lif.Name(0): Call OemToChar(h2$, h2$): LiefName1$ = RTrim$(h2$)
h2$ = lif.Name(1): Call OemToChar(h2$, h2$): LiefName2$ = RTrim$(h2$)
h2$ = lif.Name(2): Call OemToChar(h2$, h2$): LiefName3$ = RTrim$(h2$)
h2$ = lif.Name(3): Call OemToChar(h2$, h2$): LiefName4$ = RTrim$(h2$)

GhIDF$ = Mid$(lif.rest, 92, 7)
BetrNr$ = Mid$(lif.rest, 79, 7)

X$ = Mid$(lif.rest, 17, 14): X$ = LTrim$(X$): X$ = RTrim$(X$): TelGh$ = X$
i% = 1
While (i% <= Len(TelGh$))
  If InStr("0123456789:<=>;.TP& ", Mid$(TelGh$, i%, 1)) = 0 Then
    TelGh$ = Mid$(TelGh$, 1, i% - 1) + Mid$(TelGh$, i% + 1)
  Else
    i% = i% + 1
  End If
Wend

'1.79 Lieferantendurchwahl - nur zulassen, wenn rein numerisch
X$ = Mid$(lif.rest, 75, 4): X$ = LTrim$(X$): X$ = RTrim$(X$)
ok% = True
For i% = 1 To 4
  If InStr("0123456789 ", Mid$(X$, i%, 1)) = 0 Then ok% = 0
Next i
If (ok%) Then TelGh$ = TelGh$ + X$

ApoIDF$ = BetrNr$

End Sub

Function ZeigeModemTyp$()
Dim h$

h$ = ""
If (para.isdn) Then
    h$ = "ISDN-Karte"
    If (Lintech%) Then h$ = h$ + " (Lintech)"
Else
    h$ = "Seriellmodem (" + xFilePara$ + ")"
End If

ZeigeModemTyp$ = h$

End Function

Sub EntferneAktivKz(lief%)
Dim i%, Max%

bek.SatzLock (1)
bek.GetRecord (1)
Max% = bek.erstmax
For i% = 1 To Max%
    bek.GetRecord (i% + 1)
    If (lief% = -1) Or (bek.aktivlief = lief%) Then
        bek.aktivlief = 0
        bek.aktivind = 0
        bek.PutRecord (i% + 1)
    End If
Next i%
bek.SatzUnLock (1)

End Sub

