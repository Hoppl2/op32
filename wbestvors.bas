Attribute VB_Name = "modBestvors"
Option Explicit

Public BestvorsAbbruch%

Public ByteLeiste$
Dim AbholerPzns$

Dim bPZN$()
Dim nbPZN$()
Dim bkPZN$()
Dim besPZN$()

Dim bMax%
Dim nbMax%
Dim bkMax%
Dim besMax%

Dim asatz%, ssatz%

Dim TaraMM%
Dim Retour%
Dim AbholerMenge%
Dim MinMen%, BestMen%

Dim xheute%

Dim AbholNummer%
Dim AnzArtikelDazu%

Private Const DefErrModul = "wbestvors.bas"

Sub BestellVorschlag(Optional HintergrundAktiv% = True)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("BestellVorschlag")
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
Dim i%, sMax%, VonSatz%, BisSatz%
Dim Prozent!, StartZeit!, Dauer!, GesamtDauer!, RestDauer!
Dim lSatz&
Dim h$

AnzArtikelDazu% = 0

h$ = Format(Now, "DDMMYY")
xheute% = iDate(h$)

Call EinlesenWÜ

If (para.BVRetour = "J") Then
    Call EinlesenNB
End If

Call EinlesenBestellung

If (InStr(para.Benutz, "Y") > 0) Then
    Call AbholerTesten
    AbholerMenge% = 0 'damit normale Bestellung nicht falsch wird
End If

Call ass.GetRecord(1)
sMax% = ass.erstmax

If (sMax% <= 0) Then Call DefErrPop: Exit Sub

If (HintergrundAktiv%) Then

    VonSatz% = HintergrundSsatz% + 1
    BisSatz% = VonSatz% + HintergrundAnz% - 1
    If (VonSatz% = 1) Then VonSatz% = 2
    
    For ssatz% = VonSatz% To BisSatz%
        If (ssatz% > sMax%) Then Exit For
        
        lSatz& = ssatz%
        Call ass.GetRecord(lSatz&)
        
        If (InStr(AbholerPzns$, " " + ass.pzn) = 0) Then
            Call ArtikelInBestellung
        End If
    Next ssatz%
    
    HintergrundSsatz% = HintergrundSsatz% + HintergrundAnz%
    If (HintergrundSsatz% > sMax%) Then HintergrundSsatz% = 0
    
Else
    
    BestvorsAbbruch% = False

    frmBestVors!prgBestVors.Max = sMax%
    StartZeit! = Timer

    ass.GetRecord (1)

    For ssatz% = 1 To sMax%
        ass.GetRecord

        If (InStr(AbholerPzns$, " " + ass.pzn) = 0) Then
            Call ArtikelInBestellung
        End If

        If (ssatz% Mod 100 = 0) Then
            frmBestVors!lblBestVorsStatusWert(0).Caption = ssatz%
            frmBestVors!lblBestVorsStatusWert(1).Caption = AnzArtikelDazu%
            Dauer! = Timer - StartZeit!
            frmBestVors!lblBestVorsDauerWert(0).Caption = Format$(Dauer! \ 60, "##0") + ":" + Format$(Dauer! Mod 60, "00")
            Prozent! = (ssatz% / sMax%) * 100!
            If (Prozent! > 0) Then
                GesamtDauer! = (Dauer! / Prozent!) * 100!
            Else
                GesamtDauer! = Dauer!
            End If
            RestDauer! = GesamtDauer! - Dauer!
            frmBestVors!lblBestVorsDauerWert(1).Caption = Format$(RestDauer! \ 60, "##0") + ":" + Format$(RestDauer! Mod 60, "00")
            frmBestVors!prgBestVors.Value = ssatz%
            frmBestVors!lblBestVorsProzent.Caption = Format$(Prozent!, "##0") + " %"
            
            h$ = Format$(Prozent!, "##0") + " %"
            With frmBestVors!picBestVorsProgress
                .Cls
                .CurrentX = (.ScaleWidth - .TextWidth(h$)) \ 2
                .CurrentY = (.ScaleHeight - .TextHeight(h$)) \ 2
                frmBestVors!picBestVorsProgress.Print h$
                frmBestVors!picBestVorsProgress.Line (0, 0)-((Prozent! * .ScaleWidth) \ 100, .ScaleHeight), vbHighlight, BF
'                Call BitBlt(.hdc, 0, 0, .ScaleWidth, .ScaleHeight, .hdc, 0, 0, &HCC0020)
            End With
            
            DoEvents
            If (BestvorsAbbruch% = True) Then
                Exit For
            End If
        End If
    Next ssatz%

End If

Call DefErrPop
End Sub

Sub EinlesenWÜ()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("EinlesenWÜ")
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
Dim i%, j%, Max%

wu.GetRecord (1)
Max% = wu.erstmax
ReDim bPZN$(Max%)
j% = -1
For i% = 1 To Max%
    wu.GetRecord
    If (wu.besorger <> "B") Then
        j% = j% + 1
        bPZN$(j%) = Left$(wu.at, 7) + wu.li
    End If
Next i%
bMax% = j%

Call DefErrPop
End Sub

Sub EinlesenNB()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("EinlesenNB")
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
Dim i%, j%, k%, Max%, gefunden%, gAlm%
Dim pzn$, gPzn$

nachb.OpenDatei ("R")
nachb.GetRecord (1)
Max% = nachb.erstmax
ReDim nbPZN$(Max%)
j% = -1
For i% = 1 To Max%
    nachb.GetRecord
    If (nachb.alm <> 0) Then
        gefunden% = False
        pzn$ = Left$(nachb.at, 7)
        For k% = 0 To j%
            'Suchen, ob PZN schon da. Wenn ja - Menge addieren
            gPzn$ = Left$(nbPZN$(k%), 7)
            gAlm% = CVI(Mid$(nbPZN$(k%), 8, 2))
            If (gPzn$ = pzn$) Then
                nbPZN$(k%) = gPzn$ + MKI(gAlm% + nachb.alm)
                gefunden% = True
                Exit For
            End If
        Next k%
        If (gefunden% = False) Then
            j% = j% + 1
            nbPZN$(j%) = pzn$ + MKI(nachb.alm)
        End If
    End If
Next i%
nbMax% = j%
nachb.CloseDatei

Call DefErrPop
End Sub

Sub EinlesenBestellung()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("EinlesenBestellung")
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
Dim i%, j%, Max%

Call bek.GetRecord(1)
Max% = bek.erstmax
ReDim bkPZN$(Max%)
j% = -1
For i% = 1 To Max%
    bek.GetRecord
    If (bek.auto <> "v") Then
        j% = j% + 1
        bkPZN$(j%) = bek.pzn
    End If
Next i%
bkMax% = j%

Call DefErrPop
End Sub

Sub ArtikelInBestellung()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("ArtikelInBestellung")
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
Dim BestellMen%, Einfuegen%, vorhanden%

If ((Val(ass.pzn) = 0) Or (Chr$(ass.halt) = "S")) Then Call DefErrPop: Exit Sub

If (ass.vmm = 8224) Then
    ass.vmm = 0
End If

MinMen% = ass.mm
If ((ass.vmm > 0) Or ((ass.vmm = 0) And (ass.vbm > 0))) Then MinMen% = ass.vmm

Retour% = RetourenSuchen%

BestellMen% = ass.bm
If (ass.vbm > 0) Then BestellMen% = ass.vbm

If (para.LagerAuffuellen) Then
    BestMen% = BestellMen% + AbholerMenge% + MinMen% - ass.poslag
    If (para.BVRetour = "J") Then
        BestMen% = BestMen% - Retour%
    End If
Else
    BestMen% = BestellMen%
End If

If (para.TaraKontrolle) Then Call TaraKontrolleAction

Einfuegen% = TaraMMKontrollieren%

If (Einfuegen%) Then
    'Test ob reduzieren und ob dann MM reicht
    Call ReduTest
    Einfuegen% = TaraMMKontrollieren%
End If
If (Einfuegen%) Then
    FabsErrf% = ast.IndexSearch(0, ass.pzn, FabsRecno&)
    If (FabsErrf% = 0) Then
        asatz% = CInt(FabsRecno&)
        ast.GetRecord (FabsRecno& + 1)
        Einfuegen% = PosArtikel%()
    Else
        Einfuegen% = False
    End If
End If
If (Einfuegen%) Then
    vorhanden% = VorhTest%()
    If (vorhanden% <> 0) Then
        Einfuegen% = False
        'Anzeige warum nicht .....
    End If
End If
If (Einfuegen%) Then
    If (ast.lic <> "?") Then
        Call Bestellen
    End If
End If
    
Call DefErrPop
End Sub

Sub TaraKontrolleAction()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("TaraKontrolleAction")
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
Dim i%, KleinstTaraPlatz%, iTaraPlatz%

TaraMM% = 0
If (ass.wert("max", 1) = 0) Then Call DefErrPop: Exit Sub  'nur max. 1 Taralager

'Tara-Mindestmenge = Summe der Plätze minus dem kleinsten plus 1
'damit kann nicht passieren, daß ein Platz leer ist und nix zum nachräumen da ist,
'selbst wenn alle anderen Lagerplätze voll sind und nur einer leer, wird bestellt

KleinstTaraPlatz% = 9999
For i% = 0 To 3
    iTaraPlatz% = ass.Max(i%)
    If (iTaraPlatz% > 0) Then
        TaraMM% = TaraMM% + iTaraPlatz%
        If (iTaraPlatz% < KleinstTaraPlatz%) Then
            KleinstTaraPlatz% = iTaraPlatz%
        End If
    End If
Next i%
TaraMM% = TaraMM% - KleinstTaraPlatz% + 1

If (TaraMM% <= MinMen%) Then TaraMM% = 0

Call DefErrPop
End Sub

'Sub TaraKontrolleAction()
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Call DefErrFnc("TaraKontrolleAction")
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
'Dim i%, KleinstTaraPlatz%, iTaraPlatz%
'
'TaraMM% = 0
'If (ass.teinzel(1).teplatz = 0) Then Call DefErrPop: Exit Sub   'nur max. 1 Taralager
'
''Tara-Mindestmenge = Summe der Plätze minus dem kleinsten plus 1
''damit kann nicht passieren, daß ein Platz leer ist und nix zum nachräumen da ist,
''selbst wenn alle anderen Lagerplätze voll sind und nur einer leer, wird bestellt
'
'KleinstTaraPlatz% = 9999
'For i% = 0 To 3
'    iTaraPlatz% = ass.teinzel(i%).teplatz
'    If (iTaraPlatz% > 0) Then
'        TaraMM% = TaraMM% + iTaraPlatz%
'        If (iTaraPlatz% < KleinstTaraPlatz%) Then
'            KleinstTaraPlatz% = iTaraPlatz%
'        End If
'    End If
'Next i%
'TaraMM% = TaraMM% - KleinstTaraPlatz% + 1
'
'If (TaraMM% <= MinMen%) Then TaraMM% = 0
'
'Call DefErrPop
'End Sub

Function TaraMMKontrollieren%()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("TaraMMKontrollieren%")
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
Dim assMM%, ret%

ret% = True

assMM% = MinMen%
If ((ass.poslag - AbholerMenge% + Retour%) > MinMen%) Then
    If (TaraMM% > 0) Then
        MinMen% = TaraMM%
        BestMen% = ass.bm + Retour% + MinMen% - AbholerMenge% - ass.poslag
        If ((ass.poslag - AbholerMenge% + Retour%) > MinMen%) Then ret% = False
    Else
        ret% = False
    End If
    If ((InStr(para.Benutz, "*") > 0) And (ass.poslag <= ass.flager) And (para.LstGemeinsam = "J")) Then
        BestMen% = assMM%
    End If
End If

TaraMMKontrollieren% = ret%

Call DefErrPop
End Function

Function RetourenSuchen%()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("RetourenSuchen%")
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
Dim i%, ret%

ret% = 0
If (para.BVRetour = "J") Then
    For i% = 0 To nbMax%
        If (ass.pzn = Left$(nbPZN$(i%), 7)) Then
            ret% = CVI(Mid$(nbPZN$(i%), 8, 2))
            Exit For
        End If
    Next i%
End If

RetourenSuchen% = ret%

Call DefErrPop
End Function

Sub ReduTest()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("ReduTest")
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
Dim BmStamm%, MmStamm%, verbrauch%, zeit%, lld%, llm%
Dim BmoAkt!
Dim xc$

BmStamm% = ass.bm
MmStamm% = ass.mm
If ((BmStamm% = 1) And (MmStamm% = 0)) Then Call DefErrPop: Exit Sub

verbrauch% = POSverbrauch%(zeit%)

If (verbrauch% < 1) Then
    lld% = ass.lld
    llm% = ass.llm
    verbrauch% = ass.lag + llm% - ass.poslag
    zeit% = xheute% - lld%
End If
If (zeit% < 1) Then Call DefErrPop: Exit Sub

BmoAkt! = CSng(verbrauch%) / CSng(zeit%) * para.bp1
If (BmoAkt! < Abs(BmStamm%)) Then Call Reduzieren

Call DefErrPop
End Sub

Sub Reduzieren()

'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("Reduzieren")
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
Call DefErrPop
End Sub

'Function POSverbrauch%(zeit%)
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Call DefErrFnc("POSverbrauch%")
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
'Dim j%, X%, verbrauch%
'Dim xc$
'
'verbrauch% = 0
'For j% = 9 To 0 Step -1
'    xc$ = ass.verk(j%).Datum
'    X% = CVDat%(xc$)
'    If ((X% + para.bp1 * 2) >= xheute%) Then
'        verbrauch% = verbrauch% + ass.verk(j%).Menge
'        zeit% = xheute% - X%
'    End If
'Next j%
'
'POSverbrauch% = verbrauch%
'
'Call DefErrPop
'End Function

Function POSverbrauch%(zeit%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("POSverbrauch%")
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
Dim j%, X%, verbrauch%
Dim xc$

verbrauch% = 0
For j% = 9 To 0 Step -1
    X% = ass.vkdat(j%)
    If ((X% + para.bp1 * 2) >= xheute%) Then
        verbrauch% = verbrauch% + ass.vkmng(j%)
        zeit% = xheute% - X%
    End If
Next j%

POSverbrauch% = verbrauch%

Call DefErrPop
End Function

Function PosArtikel%()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("PosArtikel%")
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

ret% = False
If (InStr("AF", ass.pp) > 0) Then ret% = True
If (InStr(para.PosAktivWG, Left$(ast.wg, 1)) > 0) Then ret% = True

PosArtikel% = ret%

Call DefErrPop
End Function

Function VorhTest%()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("VorhTest%")
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
Dim i%, mm1%, gefunden%, AnzInWü%, DirLief%, ret%

'schon in der Bestellung?
ret% = 0
For i% = 0 To bkMax%
    If (bkPZN$(i%) = ass.pzn) Then
        ret% = 1
        Exit For
    End If
Next i%

'WÜ auch prüfen ???
If (ret% = 0) Then
    'man muß wissen, ob Artikel mehrmals in WÜ steht
    AnzInWü% = 0
    DirLief% = 0
    For i% = 0 To bMax%
        If (Left$(bPZN$(i%), 7) = ass.pzn) Then
            ret% = 2
            AnzInWü% = AnzInWü% + 1
            DirLief% = Asc(Mid$(bPZN$(i%), 8, 1))
            If (AnzInWü% > 1) Then Exit For
        End If
    Next i%
End If
If ((ret% = 2) And (AnzInWü% = 1) And (para.DirektLieferanten <> "")) Then
    'bei Artikeln mit MM>1 prüfen, ob Lagerstand<=MM/2. Wenn ja, auf MM auffüllen
    '(Defekt bei Direktlieferantenbestellung verhindern)
    Call DirektLieferant(DirLief%, para.DirektLieferanten, gefunden%)
    If (gefunden%) Then
        mm1% = ass.vmm
        If (mm1% <= 0) Then mm1% = ass.mm
        If ((mm1% >= 1) And (ass.poslag < Int(mm1% / 2))) Then
            ret% = 0
            BestMen% = mm1% - ass.poslag
        End If
    End If
End If

VorhTest% = ret%

Call DefErrPop
End Function

Sub DirektLieferant(AktLief%, DirLief$, gefunden%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("DirektLieferant")
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
Dim s%, von%, bis%
Dim lief$, Such$

'Suche, ob der WÜ-Lieferant unter den in PDATEI angeführten Direktlieferanten ist
gefunden% = False
lief$ = DirLief$
While ((lief$ <> "") And Not (gefunden%))
    s% = InStr(lief$, ",")
    If (s% > 0) Then
        Such$ = Left$(lief$, s% - 1)
        lief$ = Mid$(lief$, s% + 1)
    Else
        Such$ = lief$
        lief$ = ""
    End If
    
    s% = InStr(Such$, "-")
    If (s% > 0) Then
        von% = Val(Left$(Such$, s% - 1))
        bis% = Val(Mid$(Such$, s% + 1))
        If (bis% = 0) Then bis% = 200
        If ((von% <= AktLief%) And (AktLief% <= bis%)) Then gefunden% = True
    Else
        If (Val(Such$) = AktLief%) Then gefunden% = True
    End If
Wend

Call DefErrPop
End Sub

Sub Bestellen()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("Bestellen")
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
Dim i%, X%, Max%
Dim xc$, iSatz$

bek.pzn = ast.pzn
bek.txt = ast.kurz + ast.meng + ast.meh
bek.asatz = asatz%
bek.ssatz = ssatz% '??
bek.nm = 0
bek.AEP = ast.AEP
bek.AVP = ast.AVP
bek.abl = ast.abl: If (bek.abl <> "A") Then bek.abl = "X"
bek.wg = Left$(ast.wg, 1)
bek.best = " "
bek.absage = 0
bek.bm = 1
bek.km = 1
bek.angebot = 0
bek.lief = 0
bek.auto = "+"
bek.alt = Chr$(0)
bek.nnart = 0
bek.NNAep = 0#

bek.zugeordnet = Chr$(0)
bek.zukontrollieren = Chr$(0)
bek.fixiert = Chr$(0)
'bek.musskontrollieren = Chr$(0)
bek.nochzukontrollieren = Chr$(0)

For i% = 0 To 5
    bek.actkontrolle(i%) = 111
Next i%
bek.ActZuordnung = 111
bek.aktivlief = 0
bek.aktivind = 0
bek.poslag = 0

bek.herst = ast.herst

If (ssatz% <> 0) Then
    bek.lief = ass.lief
    If (BestMen% > 0) Then
        bek.bm = BestMen%
        bek.km = ass.bm
    End If
    
    X% = ass.lld
    If ((X% + para.MonNBest) < xheute%) Then bek.alt = "?"
End If

bek.BekLaufNr = 0

bek.SatzLock (1)
bek.GetRecord (1)
bkMax% = bek.erstmax
bkMax% = bkMax% + 1
bek.erstmax = bkMax%
bek.PutRecord (1)
bek.PutRecord (bkMax% + 1)
bek.SatzUnLock (1)

ReDim Preserve bkPZN$(bkMax%)
bkPZN$(bkMax%) = ast.pzn

AnzArtikelDazu% = AnzArtikelDazu% + 1

BekartCounter% = -1

Call DefErrPop
End Sub

Sub AbholerTesten()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("AbholerTesten")
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
Dim i%, j%, erg%, KistenNr%, AbholerStatus%, gefunden%, gMenge%
Dim h$, pzn$, gPzn$

kiste.OpenDatei

besMax% = -1
For KistenNr% = 1 To 999
    If (kiste.Belegt(KistenNr%)) Then
        kiste.GetKiste (KistenNr%)
        
        For i% = 0 To 9
            erg% = kiste.GetInhalt(i%)
            If (erg%) Then
                If (kiste.WasTun = "B") Then
                    AbholerStatus% = Asc(kiste.Status)
                    If (AbholerStatus% = 1) Then
                        'fertig, aber nicht abgeholt
                        AbholerMenge% = 1
                        For j% = 1 To 10
                            h$ = RTrim$(kiste.InfoText(j% - 1))
                            If Mid$(h$, 18, 3) = " x " Then
                                AbholerMenge% = -Val(Mid$(h$, 21, 4))
                            End If
                            If Mid$(h$, 14, 4) = "PZN=" Then
                                pzn$ = Mid$(h$, 18, 7)
                            End If
                        Next j%
                        If ((Val(pzn$) > 0) And (pzn$ <> "999999")) Then
                            gefunden% = False
                            For j% = 0 To besMax%
                                'Suchen, ob PZN schon da. Wenn ja - Menge addieren
                                gPzn$ = Left$(besPZN$(j%), 7)
                                gMenge% = CVI(Mid$(besPZN$(j%), 8, 2))
                                If (gPzn$ = pzn$) Then
                                    besPZN$(j%) = gPzn$ + MKI(gMenge% + AbholerMenge%)
                                    gefunden% = True
                                    Exit For
                                End If
                            Next j%
                            If (gefunden% = False) Then
                                besMax% = besMax% + 1
                                ReDim Preserve besPZN$(besMax%)
                                besPZN$(besMax%) = pzn$ + MKI(AbholerMenge%)
                            End If
                        End If
                    End If
                End If
            End If
        Next i%
    End If
Next KistenNr%

kiste.CloseDatei

AbholerPzns$ = ""
For i% = 0 To besMax%
    pzn$ = Left$(besPZN$(i%), 7)
    AbholerMenge% = CVI(Mid$(besPZN$(i%), 8, 2))
    If (AbholerMenge% <> 0) Then
        FabsErrf% = ass.IndexSearch(0, pzn$, FabsRecno&)
        If (FabsErrf% = 0) Then
            Call ass.GetRecord(FabsRecno& + 1)
            Call ArtikelInBestellung
        End If
        AbholerPzns$ = AbholerPzns$ + " " + pzn$
    End If
Next i%

Call DefErrPop
End Sub

