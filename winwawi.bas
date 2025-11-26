Attribute VB_Name = "modWinWawi"
'1.0.52 03.05.05 AE Für A: Einwieger brauche Ablaufdatum (Prüfung in WuFertigAlle, Speichern in WkZwi, Ausdruck Kontrollblätter)
'1.0.51 31.03.05 AE op32.dll von AndiNote, LiefNamen(250)
'1.0.50 22.03.05 AE in EinzelSatz auch Prüfung Merkzettel abhängig von Fistam (F91)
'1.0.49 16.03.05 AE Merkzettel abhängig von Fistam (wenn 't' - MDB, sonst stammlos.dat)
'                   wenn RowaWu, dann keine Frage nach Signatur (immer Benutzer 99)
'                   250 Lieferanten
'1.0.48 01.02.05 AE Merkzettel jetzt als MDB
'1.0.47 01.12.04 AE EasyLink: in WÜ war bei neuerlichem Programmaufruf immer Aufteilung weg. Gefunden: nur wenn direkt WÜ aufgerufen wird,
'                   wenn vorher Bestellung, dann ok. Lag an Variable Lieferant%, behoben!
'                   Lagerkontroll-Liste: zus. Spalte VKLager
'                   wenn easyLink, dann * jetzt linksbündig (uch Bereiche in dll)
'                   wenn int.Streichung und poslag<=0, dann in Auswahl NICHT fett
'                   Signatur bei Buchen, abhängig von INi-Eintrag
'                   Anzeige der Artikel mit DokuPflicht in WÜ eingebaut
'1.0.46 15.11.04 AE Direktbezug adaptiert: CalcAlleDirektAngebote wieder aktiviert;
'                   den Teil, wo für DirLiefs wenn ein günstiges Angebot für einen GH gefunden wird, die BM uU auf 0 gesetzt wird, unter Kommentar gesetzt!
'1.0.45 11.10.04 AE Kontrollgrund 'Übl.Lieferant' richtiggestellt
'                   Kontrollgrund 'Interne Streichung' eingeführt
'1.0.44 30.08.04 AE EasyLInk: auch bei mir nicht-lagernde Artikel in Direktbezug einbeziehen (FremdArtikelInPalette)
'1.0.43 12.08.04 AE Lieferantenumsätze nach Abgabeschlüssel speichern, Ansicht in wumsatz; LiefManag: uU mehrere Besorger zu einem Auftrag zusammenfassen
'                   PreisKalkA: ESC nur möglich, wenn für alle Besorger auch AVP eingegeben
'1.0.42 02.08.04 AE Direktbezug Partner; verzögerter Start Einlesen Bestellung nach F3 (um mehrmaliges Weiterschalten zu beschleunigen)
'1.0.41 13.07.04 GS Etiketten Ö: wahlweise doch alle (abhängig von INI-Param)
'1.0.40 11.06.04 AE Auswahl Direktliefs für Verbund, nur wenn mind. einer
'1.0.39 02.06.04 AE Direktaufteilung: neuer Menüpunkt 'Aufteilung Direktbezug bearbeiten'; dort Edit EinzelBM
'1.0.38 11.05.04 AE WÜ-Optionen: 'Lieferantenabfrage bei man. Erfassen'; wenn gesetzt (default nein), dann Lieferanten-Auswahl
'                       vor Aufblenden Artikelmatchcode, danach wird dieser Lief in alle ausgewählten (freien) Aritkel gespeichert
'1.0.37 22.04.04 AE für A: Textbesorger kommen jetzt auch in Kalkulation
'                       kalk. Preise werden beim Abholer gespeichert und Besorger auf fertig gesetzt
'                       auch Wg 30-39 wird für Kontrollliste berücksichtigt
'                       Bereitstellen Etiketten: nut wenn Artikel kalkulierbar
'                       wenn Besetzt, bi szu 3x wählen, mit 90 Sek Pause
'                       an neuem Kalendertag werden bei AUfruf Besorgerverwaltung alle abgeholten Besorger entfernt
'1.0.36 22.03.04 AE Neue Kontrollgründe: AM RX, AM NON RX, NICHTAM
'                   WÜ: nach man. Ablaufbestätigung in LM nächste Zeile springen
'                   Rabatttabellen: bei Ausnahmen zus. 'AM RX'
'                   BM/EK-tabelle um Gruppen erweitert
'1.0.35 17.03.04 AE Neuer Kontrollgrund: Interim
'                   Artikel zu WÜ dazu: jetzt in jeder Ansicht möglich; Auswahlfenster angepasst
'1.0.34 09.03.04 AE bei Optionen: neuer Karteireiter 8 - Automatenware: Eingaben eines Lagercodes,
'                   dann die Lieferanten eingeben, die bei der Bestellung anstelle durch Angebote ausgewählte Lieferanten verwendet werden sollen
'                   Einbau der Auswertung der Automatenliefs
'1.0.33 04.03.04 AE 2-stellige Lagercodes eingeführt: Bedingungen, Infobereich
'                   neuer Kontrollgrund: Rezeptpfl. AM
'1.0.32 26.02.04 AE Besorger-Verwaltung: jetzt werden auch wieder die Namen von Textbesorgern angezeigt
'1.0.31 17.02.04 GS PznPrüfziffer: Endlosschleife bei Prüfziffer 10 behoben
'1.0.30 19.12.03 AE Kontrollgrund Text-Besorger; Parameter 'NichtRezeptpflichtige für Kalk anbieten'
'1.0.29 11.12.03 AE Handling Wg3 für Preiskalk, Bestellung. Ausdruck Kontrollblätter
'                   wieder richtiger Spaltenindex für allg.Lieferanten
'                   IN A: in 13-Code PZN vermuten
'                   Interimsbestellung: keine Angebote prüfen, Anzeige 'Interim' in Zeile
'1.0.28 21.11.03 AE DAO351 durch DAO360 ersetzt. Damit in allen Progs einheitlich !
'1.0.27 10.11.03 AE WÜ-SucheZeile: neuer Eintrag Nachlieferung, wenn Artikel in NL vorhanden; bei Auswahl wird er in WÜ gestellt
'                   Ö: Absagen entfernen überarbeitet, Auswahl welche Absagen in NL gehen sollen
'                   neuer Sound: wwfertig (für wbestk2)
'                   cldBesorger: Handling fertig/ABgeholt auch für Kistenstatus
'1.0.25 30.10.03 AE Für Ö: Feiertage, auch nn-Datei speichern ..
'1.0.24 14.10.03 AE Für Ö: Einbau Merkzettel ohne Indexdateien ...
'1.0.23 18.09.03 AE Einbau Partnerapos für Direktbezug
'1.0.22 12.09.03 AE Interimsbestellung: auch wenn Artikel für DirektLief in BESTELLUNG (VorhTest)
'1.0.21 09.09.03 AE Teilbestellungen bei Ladenhüter-Bestellungen
'1.0.19 02.07.03 AE Lieferprofile - Besorger
'1.0.16 04.04.03 AE wenn PZN < 7 Stellen, mit führenden Nullen auffüllen (EinzelSatz)
'                   Prüfung Interimsbestellung auch nach sF8 (manueller Produktpalette)
'       08.04.03 AE Direktbezug - 'Direkt möglich ab BM ..' funktional eingebaut
'1.0.15 27.03.02 AE Adaption Handling Direktbezug laut Liste DrS
'1.0.10 27.10.02 AE Erweiterungen Direktbezug
'1.0.8  06.9.02 AE  Änderungen Direktbezug, Prio-Liste
'1.0.7  28.8.02 AE  Änderungen Direktbezug bzgl. Handling
'1.0.6  19.8.02 AE  Erweiterungen Direktbezug laut Liste OM
'1.0.5  12.8.02 AE  Erweiterungen Direktbezug
'1.0.4  17.7.02 AE  bei Anzeige Kontrollen: Artikel nur anzeige, wenn BM>0 oder nicht für Diretklief
'                   bei ALternativ-Artikel: Besorgerinfo übernehmen; in Kiste PZN ändern

Option Explicit


'Public clsfabs As New clsFabsKlasse

Public ProgrammChar$

Public ProgrammNamen$(3)
Public ProgrammTyp%

Global buf$

Public TaxeDB As Database
Public TaxeRec As Recordset

Public MerkzettelDB As Database
Public MerkzettelRec As Recordset
Public MerkzettelParaRec As Recordset

Public ProgrammModus%
Public TaxeOk%
Public UebergabeStr$

Public FabsErrf%
Public FabsRecno&

Public DruckSeite%

Public PROTOKOLL%, prot%

Public UserSection$

Public Lieferant%
Public MarkWert#, GesamtWert#, Warenwert#

Public KeinRowColChange%
Public WirklichKeinRowColChange%


Public GlobBekMax%

Public HintergrundSsatz%, HintergrundAnz%

Public WochenTag$(6)

Public AutomaticSend%
Public AutomaticInd%
Public VorAutomaticLief%


Public ast As clsStamm
Public ass As clsStatistik
Public taxe As clsTaxe
Public lif As clsLieferanten
Public lifzus As clsLiefZusatz
'Public WawiGh As clsWawiDat
Public WawiDirekt As clsWawiDat
Public ww As clsWawiDat
Public bek As clsBekart
Public nachb As clsNachbearbeitung
Public wu As clsWÜ
Public BESORGT As clsBesorgt
Public Merkzettel As clsMerkzettel
Public absagen As clsAbsagen
Public kiste As clsKiste
Public ArtText As clsArttext
Public lieftext As clsLieftext
Public etik As clsEtiketten
Public ang As clsAngebote
Public ManAng As clsManuellAngebote
Public para As clsOpPara
Public wpara As clsWinPara
Public artstat As clsArtStatistik
Public rabtab As clsRabattTabellen
Public nnek As clsNNEK
Public rk As clsRueckKauf
'Public ga As clsGemAufschlag
Public ZusWv3 As clsZusWv3

Public EditErg%
Public EditModus%
Public EditTxt$
Public EditAnzGefunden%
Public EditGef%(49)

Public ArtikelStatistik%

Public NaechstLiefernderLieferant%
Public AktUhrzeit%

Public LiefWechselOk%

Public KommentarOk&()

Public ActProgram As Object

Public WuLifDat$
Public RkLifDat$

Public ErstAuslesen%

Public OptionenNeu%

Public EkEingabeModus%, EkEingabeArt%
Public EkEingabePreis#

Public NettoEk%

Public OptionenModus%

Public EinzelneWu%

Public PreisKalkErg%
Public PreisKalkAepChange%

Public AltLastNamen$(7)
Public AltLastStr$

Public IstAltLast%

Public ActBenutzer%

Public DarfAlternativArtikel%
Public DarfAltlastenLoschen%
Public DarfFertigMachen%
Public DarfHinzufuegen%
Public DarfPreisKalk%
Public DarfPreisKalkBesorger%
Public DarfPreisKontrolle%
Public DarfRmKontrolle%
Public DarfWumsatz%

Public StrichCodeErg$
Public WuFrageTxt$
Public WuFrageErg%

Public FreiKalkPreise#(2)
Public FreiKalkMw$

Public BestellAnzeige%

Public Dp04Ok%
'Public Dp04Para$
Public SeriellScannerOk%
Public SeriellScannerParam$

Public DruckeLagerKontrollListe%

Public AngebotY%

'Public FarbeGray&
Public INI_DATEI As String

Public RowaAktiv%

Public TageSpeichern%

Public vAnzeigeSperren%
Public AnzRetourenDruck%

Public DirektBewertung#

Public WuInDm%

Public KalkOhnePreis%

Public BestVorsAktiv%

Public AepKalkX%, AepKalkY%

Public Bm0Anzeigen%
Public ManuellDirektReduce%

Public LetztBesorgerWeg$

Public EditAufteilungPzn$

Public IstRowaWu%

Private Const DefErrModul = "WINWAWI.BAS"

Function InitProgramm%(Optional Visible% = True)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("InitProgramm%")
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
Dim i%, auf%, sMax%
Dim h$

frmAction.tmrAction.Enabled = False
Call PruefeTaetigkeiten

Select Case (ProgrammChar$)
    Case "B"
        
        Call EntferneGeloeschteZeilen(1)
        Call StopAnimation(frmAction)
            
        BekartCounter% = -1
        zTabelleAktiv% = True
        zManuellAktiv% = False
        Call ActProgram.AuslesenBestellung(True, False, True)
        Call frmAction.WechselModus(0)
        
    Case "W"
        Call StopAnimation(frmAction)
        
'        Call PruefeTaetigkeiten
        
        Call ActProgram.EinlesenAlteWue
        
        If (AufschlagsTabelle(0).PreisBasis = 0) Then
            For i% = 1 To 9
                auf% = para.Aufschlag(i%)
                If (auf% > 100) Then
                    AufschlagsTabelle(i% - 1).PreisBasis = 0
                    AufschlagsTabelle(i% - 1).Aufschlag = 0
                Else
                    AufschlagsTabelle(i% - 1).PreisBasis = 2
                    If (auf% = 100) Then
                        AufschlagsTabelle(i% - 1).Aufschlag = 211
                    Else
                        AufschlagsTabelle(i% - 1).Aufschlag = auf%
                    End If
                End If
            Next i%
            
            Call ass.GetRecord(1)
            sMax% = ass.erstmax

            If (para.Land <> "A") Then
                For i% = 1 To sMax%
                    ass.GetRecord (i% + 1)
                    FabsErrf% = ast.IndexSearch(0, ass.pzn, FabsRecno&)
                    If (FabsErrf% = 0) Then
                        ast.GetRecord (FabsRecno& + 1)
                        ass.pk = Val(ast.ka)
                        ass.PutRecord (i% + 1)
                    End If
                Next i%
            End If
            
        End If

        WuLifDat$ = Chr$(0)
        Call ActProgram.AuslesenWu
        BestVorsAktiv% = True
        
    Case "V"
        Call StopAnimation(frmAction)
        Call ActProgram.AuslesenBesorger
'        Call frmAction.WechselModus(0)
        
End Select

InitProgramm% = True
Call DefErrPop
End Function

Sub InitMisc()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("InitMisc")
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
'ProgrammRec(0).Name = "Taxe"
'ProgrammRec(0).ProgrammChar = "T"
'ProgrammRec(0).Hotkey = "T"

ProgrammNamen$(0) = "Bestellung"
ProgrammNamen$(1) = "Direkt-Bezug"
ProgrammNamen$(2) = "Warenübernahme"
ProgrammNamen$(3) = "Verwaltung Besorger"

WochenTag$(0) = "Montag"
WochenTag$(1) = "Dienstag"
WochenTag$(2) = "Mittwoch"
WochenTag$(3) = "Donnerstag"
WochenTag$(4) = "Freitag"
WochenTag$(5) = "Samstag"
WochenTag$(6) = "Sonntag"
    
AltLastNamen$(0) = "nur Gutschriften (alphabetisch)"
AltLastNamen$(1) = "nur Gutschriften (chronologisch)"
AltLastNamen$(2) = "nur Nachlieferungen (alphabetisch)"
AltLastNamen$(3) = "nur Nachlieferungen (chronologisch)"
AltLastNamen$(4) = "nur Nachverrechnungen (alphabetisch)"
AltLastNamen$(5) = "nur Nachverrechnungen (chronologisch)"
AltLastNamen$(6) = "nur Retourwaren (alphabetisch)"
AltLastNamen$(7) = "nur Retourwaren (chronologisch)"
    
AltLastStr$ = "GgLlVvRr"

Call DefErrPop
End Sub

Sub Main()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("Main")
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
Dim i%, erg%, tmp%, ind%, t%, m%, j%, c%
Dim l&, hWnd&
Dim h$, SQLStr$, DirRet$, FuDate$, CdInfoVer$, s$
Dim WinText As String * 255
Dim a!

'If (UCase(App.Path) = "C:\VB5SRC\WINWAWI") Then
'    ChDrive "P"
'    ChDrive "C"
'End If

Call InitMisc
If (App.PrevInstance) Then
    hWnd& = GetForegroundWindow()
    hWnd& = GetWindow(hWnd&, GW_HWNDFIRST)
    Do Until (hWnd& = 0)
        l& = GetWindowText(hWnd&, WinText, 255)
        h$ = Left$(WinText, l&)
        If (Left$(h$, 9) = "Altlasten") Then
            AppActivate h$
            End
        End If
        For i% = 0 To 3
            j% = Len(h$)
            m% = Len(ProgrammNamen$(i%))
            If (j% > m%) Then
                s$ = Left$(h$, m%)
                If (s$ = ProgrammNamen$(i%)) Then
                    AppActivate h$
                    End
                End If
            End If
        Next i%
        hWnd& = GetWindow(hWnd&, GW_HWNDNEXT)
    Loop
    End
End If

ReDim KommentarOk&(0)

If (Dir$("fistam.dat") = "") Then ChDir "\user"
INI_DATEI = CurDir + "\winop.ini"

Set ast = New clsStamm
Set ass = New clsStatistik
Set taxe = New clsTaxe
Set lif = New clsLieferanten
Set lifzus = New clsLiefZusatz
Set ww = New clsWawiDat
Set bek = New clsBekart
Set nachb = New clsNachbearbeitung
Set wu = New clsWÜ
Set BESORGT = New clsBesorgt
Set Merkzettel = New clsMerkzettel
Set absagen = New clsAbsagen
Set kiste = New clsKiste
Set ArtText = New clsArttext
Set lieftext = New clsLieftext
Set etik = New clsEtiketten
Set ang = New clsAngebote
Set ManAng = New clsManuellAngebote
Set para = New clsOpPara
Set wpara = New clsWinPara
Set artstat = New clsArtStatistik
Set rabtab = New clsRabattTabellen
Set nnek = New clsNNEK
Set rk = New clsRueckKauf
'Set ga = New clsGemAufschlag
Set ZusWv3 = New clsZusWv3

UserSection$ = "Computer" + Format(Val(para.User))
Call wpara.HoleWindowsParameter
Dp04Ok% = True

'SeriellScannerOk% = True
'SeriellScannerOk% = ModemParameter("SCANNER", SeriellScannerParam$)
SeriellScannerOk% = False

Call para.HoleFirmenStamm

frmAction.Show

Call StartAnimation(frmAction, "Parameter werden eingelesen ...")

Call para.AuslesenPdatei
Call para.HoleZuzahlungen
Call para.EinlesenPersonal

ActBenutzer% = HoleActBenutzer%

Call LoescheAlteTage
    
Call GetTaskId("Text-API")

'a! = Timer
'While (Timer - a!) < 20
'Wend
    
ast.OpenDatei
ass.OpenDatei

TaxeOk% = False
        
erg% = 0
h$ = para.TaxeLw + ":\taxe\taxe.mdb"
'h$ = "m:\taxe\taxe.mdb"
Set TaxeDB = taxe.OpenDatenbank(h$, False, True)
'On Error GoTo ErrorHandler
'Set TaxeDB = OpenDatabase(h$, False, True)
'If (erg% > 0) Then
'    erg% = MsgBox("Die Windows-Taxe existiert auf Laufwerk " + TaxeLw$ + ": nicht !", vbInformation)
'End If
'On Error GoTo 0

TaxeOk% = True
Set TaxeRec = TaxeDB.OpenRecordset("Taxe", dbOpenTable)
    
OpPartnerLiefs$ = ""
OpDirektPartner$ = ""
FremdPznOk% = OpenOpConnect(FremdPznDB, OpPartnerDB)
If (FremdPznOk%) Then
    Set FremdPznRec = FremdPznDB.OpenRecordset("Artikel", dbOpenTable)
    FremdPznRec.Index = "Unique"
    Set DirektAufteilungRec = FremdPznDB.OpenRecordset("DirektAufteilung", dbOpenTable)
    DirektAufteilungRec.Index = "Unique"
    Set OpPartnerRec = OpPartnerDB.OpenRecordset("PartnerProfile", dbOpenTable)
    OpPartnerRec.Index = "Unique"
    Call EinlesenOpPartnerLiefs
End If

ww.OpenDatei
If (ww.DateiLen = 0) Then
    ww.erstmax = 0
    ww.erstlief = 0
    ww.erstcounter = 0
    ww.erstrest = String(ww.DateiLen, 0)
    ww.PutRecord (1)
End If

bek.OpenDatei
If (bek.DateiLen = 0) Then
    bek.erstmax = 0
    bek.erstlief = 0
    bek.erstcounter = 0
    bek.erstrest = String(bek.DateiLen, 0)
    bek.PutRecord (1)
End If

wu.OpenDatei
If (wu.DateiLen = 0) Then
    wu.erstmax = 0
    wu.erstrest = String(wu.DateiLen, 0)
    h$ = wu.erstrest
    Mid$(h$, 13, 1) = Chr$(1)   'version !
    wu.erstrest = h$
    wu.PutRecord (1)
End If

etik.OpenDatei
If (etik.DateiLen = 0) Then
    etik.erstmax = 0
    etik.erstges = 0
    etik.erstrest = String(etik.DateiLen, 0)
    etik.PutRecord (1)
End If

lif.OpenDatei
Call EinlesenLieferanten

ang.OpenDatei
ManAng.OpenDatei

ArtText.OpenDatei
lieftext.OpenDatei

'BESORGT.OpenDatei
erg% = OpenCreateMerkzettelDB%

nnek.OpenDatei

If (para.Land = "A") Then
'    ga.OpenDatei
    ZusWv3.OpenDatei
End If

AnzBestellWerteRows% = 1

Call HoleIniKontrollen
Call HoleIniZuordnungen
Call HoleIniRufzeiten
Call HoleRufzeitenLieferanten
Call HoleIniTaetigkeiten
Call HoleIniAufschlagsTabelle
Call HoleIniRundungen
Call HoleIniWuSortierungen
Call HoleIniFeiertage
Call HoleIniVerfallWarnungen
Call HoleIniSignaturen
            
lifzus.OpenDatei
If (lifzus.DateiLen = 0) Then
    Call lifzus.InitDatei
    If (para.Land = "D") Then
        Call rabtab.KonvertTabelle
    End If
    Call KonvSchwellwerte
End If
'lifzus.EinlesenLiefFuerHerst

Call InitIsdnAnzeige

ModemOk% = TestModemPar%
If (Wbestk2ManuellSenden%) Then ModemOk% = True

Dp04Ok% = ModemParameter("DP-0", h$, False)
SeriellScannerOk% = ModemParameter("SCANNER", SeriellScannerParam$, False)


NettoEk% = False
If (InStr(para.Benutz, "#") > 0) Or (InStr(para.Benutz, "&") > 0) Then
    NettoEk% = True
End If
NettoEk% = (para.Land = "D")

Call InitWumsatzDat

If (InStr(para.Benutz, "Y") <= 0) Then
    With frmAction
        .mnuDateiInd(3).Enabled = False
        .cmdDatei(3).Enabled = False
    End With
End If



erg% = InitProgramm%

Call DefErrPop: Exit Sub
    
ErrorHandler:
    erg% = Err
    If ((erg% > 0) And (erg% <> 3024) And (erg% <> 3044)) Then
        Call iMsgBox("Fehler" + Str$(Err) + " beim Öffnen der Taxe " + h$ + vbCr + Err.Description, vbCritical, "OpenDatabase")
        End
    End If
    Err = 0
    Resume Next
    Return

Call DefErrPop
End Sub

Sub Programmende()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("ProgrammEnde")
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
Dim h$

'FabsErrf% = clsfabs.Cmnd("K\20", FabsRecno&)
'If (FabsErrf%) Then
'    erg% = MsgBox("Fehler " + Str$(FabsErrf%) + " beim Schliessen der Statistik-Daten", vbInformation)
'End If
    
wpara.ExitEndSub

TaxeDB.Close

If (InStr(para.Benutz, "t") > 0) Then
    MerkzettelDB.Close
Else
    BESORGT.CloseDatei
End If

ast.CloseDatei
ass.CloseDatei
'taxe.CloseDatei
lif.CloseDatei
lifzus.CloseDatei
ww.CloseDatei
bek.CloseDatei
nachb.CloseDatei
wu.CloseDatei
'BESORGT.CloseDatei
ArtText.CloseDatei
lieftext.CloseDatei
etik.CloseDatei
ang.CloseDatei
ManAng.CloseDatei
'para.CloseDatei
'wpara.CloseDatei
nnek.CloseDatei

If (para.Land = "A") Then
'    ga.CloseDatei
    ZusWv3.CloseDatei
End If

ast.FreeClass
ass.FreeClass
'taxe.FreeClass
lif.FreeClass
lifzus.FreeClass
ww.FreeClass
bek.FreeClass
nachb.FreeClass
wu.FreeClass
BESORGT.FreeClass
absagen.FreeClass
'kiste.FreeClass
ArtText.FreeClass
lieftext.FreeClass
etik.FreeClass
ang.FreeClass
ManAng.FreeClass
'para.FreeClass
'wpara.FreeClass
rabtab.FreeClass
nnek.FreeClass
rk.FreeClass
'ga.FreeClass
ZusWv3.FreeClass

Set ast = Nothing
Set ass = Nothing
Set taxe = Nothing
Set lif = Nothing
Set lifzus = Nothing
Set ww = Nothing
Set bek = Nothing
Set nachb = Nothing
Set wu = Nothing
Set BESORGT = Nothing
Set Merkzettel = Nothing
Set absagen = Nothing
Set kiste = Nothing
Set ArtText = Nothing
Set lieftext = Nothing
Set etik = Nothing
Set ang = Nothing
Set ManAng = Nothing
Set para = Nothing
Set wpara = Nothing
Set artstat = Nothing
Set rabtab = Nothing
Set nnek = Nothing
Set rk = Nothing
'Set ga = Nothing
Set ZusWv3 = Nothing

Call frmAction.frmActionUnload
    
End
Call DefErrPop
End Sub

Sub EinlesenLieferanten()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("EinlesenLieferanten")
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
Dim i%, j%, h$

With frmAction.lstSortierung
    .Clear
    j% = 0
    For i% = 1 To lif.AnzRec
        Call lif.GetRecord(i% + 1)
        h$ = lif.kurz
        h$ = Trim$(h$)
        If (h$ <> "") Then
            If (Asc(Left$(h$, 1)) >= 32) Then
                h$ = h$ + " (" + Mid$(Str$(i%), 2) + ")"
                .AddItem h$
                If (i% <= 10) Then
                    LiefNamen$(j%) = h$
                    lif.LiefName(j%) = h$
                    j% = j% + 1
                End If
            End If
        End If
    Next i%
    
'    ReDim LiefAlleNamen(.ListCount - 1)
    LiefNamen$(j%) = String$(50, "-")
    lif.LiefName(j%) = String$(50, "-")
    j% = j% + 1
    For i% = 0 To (.ListCount - 1)
        .ListIndex = i%
        LiefNamen$(j%) = RTrim$(.text)
        lif.LiefName(j%) = RTrim$(.text)
        j% = j% + 1
    Next i%
    AnzLiefNamen% = j%
    lif.AnzLiefNamen = j%
End With

Call DefErrPop
End Sub

Sub DruckeBestellung(Optional DruckDatei$ = "")
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("DruckeBestellung")
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
Dim i%, j%, pos%, sp%(9), SN%, Y%, Max%, ind%, DRUCKHANDLE%, iLief%, iRufzeit%
Dim EK#, SendWert#
Dim header$, tx$, h$, KopfZeile$

Call StartAnimation(frmAction, "Ausdruck wird erstellt ...")



frmAction.lstSortierung.Clear


If (DruckDatei$ = "") Then
    KopfZeile$ = ""
    header$ = frmAction.lblArbeit(0).Caption
    
    ww.SatzLock (1)
    ww.GetRecord (1)
    Max% = ww.erstmax
    
    For i% = 1 To Max%
        ww.GetRecord (i% + 1)
        If (ww.status = 1) Then
'            If ((Asc(ww.pzn) < 128) And (ww.zugeordnet = "J")) Then
            If ((ww.loesch = 0) And (ww.zugeordnet = "J")) Then
                If (ww.zukontrollieren <> "1") And (ww.aktivlief = 0) Then
                    EK# = ww.aep
        
                                    
                    h$ = Left$(ww.txt, 28) + vbTab + Mid$(ww.txt, 29, 5)
                    h$ = h$ + vbTab + Mid$(ww.txt, 34, 2)
                    Call OemToChar(h$, h$)
    
                    tx$ = Format(Abs(ww.bm), "0")
                    h$ = h$ + vbTab + tx$
                    
                    tx$ = ""
                    If (Abs(ww.nm > 0)) Then
                        tx$ = Format(Abs(ww.nm), "0")
                    End If
                    h$ = h$ + vbTab + tx$
                    
                    tx$ = Format(EK# * Abs(ww.bm), "0.00")
                    h$ = h$ + vbTab + tx$
                    h$ = h$ + vbTab + ww.pzn + vbTab
                    
                    frmAction.lstSortierung.AddItem h$
                End If
            End If
    '    End If
        End If
    Next i%
    
    ww.SatzUnLock (1)
Else
    KopfZeile$ = "gesendete Bestellung"
    header$ = "(unbekannt)"
    iLief% = Val(Left$(DruckDatei$, 3))
    If (iLief% > 0) And (iLief% <= lif.AnzRec) Then
        lif.GetRecord (iLief% + 1)
        h$ = RTrim$(lif.Name(0))
        
        iRufzeit% = Val(Mid$(DruckDatei$, 4, 4))
        h$ = h$ + "  (" + Format(iRufzeit% \ 100, "00") + ":" + Format(iRufzeit% Mod 100, "00") + ")"
        
        If (InStr(DruckDatei$, "m.") > 0) Then h$ = h$ + "  manuell"
        header$ = h$
    End If
            
    
    DRUCKHANDLE% = FileOpen("winw\" + DruckDatei$, "I")
    
    Do While Not EOF(DRUCKHANDLE%)
        Line Input #DRUCKHANDLE%, h$
        If (Left$(h$, 4) <> "RM: ") Then frmAction.lstSortierung.AddItem h$
    Loop
    
    Close #DRUCKHANDLE%
End If



Printer.ScaleMode = vbTwips
Printer.Font.Name = "Arial"

DruckSeite% = 0
    
Call DruckKopf(header$, "", KopfZeile$)
        
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
                
    
SendWert# = 0#

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
            
            If (j% = 5) Then SendWert# = SendWert# + CDbl(tx$)
            
            Printer.Print tx$;
        Next j%
        
        Printer.Print " "
        
        If (Printer.CurrentY > Printer.ScaleHeight - 1000) Then
            Call DruckFuss
            Call DruckKopf(header$, "", KopfZeile$)
        End If
    Next i%
End With
    
Y% = Printer.CurrentY
Printer.Line (0, Y%)-(sp%(6), Y%)
Printer.CurrentX = sp%(0)
Printer.Print Format(AnzBestellArtikel%, "0") + " Positionen";
tx$ = Format(SendWert#, "0.00")     'MarkWert#
Printer.CurrentX = sp%(6) - Printer.TextWidth(tx$ + "x")
Printer.Print tx$;

Printer.Print " "

Call DruckFuss(False)

Printer.EndDoc

Call StopAnimation(frmAction)

Call DefErrPop
End Sub

Sub DruckeWu()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("DruckeWu")
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
Dim i%, j%, pos%, sp%(12), SN%, Y%, Max%, ind%, iLief%, AnzWuArtikel%
Dim header$, tx$, h$, KopfZeile$

Call StartAnimation(frmAction, "Ausdruck wird erstellt ...")


h$ = frmAction.Caption
ind% = InStr(h$, "-")
If (ind% > 0) Then
    KopfZeile$ = Left$(h$, ind% - 1)
Else
    KopfZeile$ = h$
End If
header$ = frmAction.lblArbeit(0).Caption

Printer.ScaleMode = vbTwips
Printer.Font.Name = "Arial"

DruckSeite% = 0
    
Call DruckKopf(header$, "", KopfZeile$)
Printer.Font.Size = 11
        
sp%(0) = Printer.TextWidth(String$(2, "9"))
sp%(1) = sp%(0) + Printer.TextWidth(String$(9, "9"))
sp%(2) = sp%(1) + Printer.TextWidth(String$(23, "X"))
sp%(3) = sp%(2) + Printer.TextWidth(String$(6, "X"))
sp%(4) = sp%(3) + Printer.TextWidth(String$(3, "X"))
sp%(5) = sp%(4) + Printer.TextWidth(String$(4, "9"))
sp%(6) = sp%(5) + Printer.TextWidth(String$(4, "9"))
sp%(7) = sp%(6) + Printer.TextWidth(String$(4, "9"))
sp%(8) = sp%(7) + Printer.TextWidth(String$(4, "9"))
sp%(9) = sp%(8) + Printer.TextWidth(String$(4, "9"))
sp%(10) = sp%(9) + Printer.TextWidth(String$(4, "9"))
sp%(11) = sp%(10) + Printer.TextWidth(String$(7, "9"))
sp%(12) = sp%(11) + Printer.TextWidth(String$(8, "9"))
        
Printer.CurrentX = sp%(0)
Printer.Print "P Z N";
Printer.CurrentX = sp%(1)
Printer.Print "A R T I K E L";

With frmAction.flxarbeit(0)
    For i% = 5 To 12
        tx$ = .TextMatrix(0, i%)
        If (i% = 9) Or (i% = 10) Then tx$ = LCase(tx$)
        Printer.CurrentX = sp%(i%) - Printer.TextWidth(tx$)
        Printer.Print tx$;
    Next i%
End With

Printer.Print " "
Y% = Printer.CurrentY
Printer.Line (0, Y%)-(sp%(12), Y%)
                
    
With frmAction.flxarbeit(0)
    AnzWuArtikel% = .Rows - 1
    For i% = 1 To AnzWuArtikel%
        
        Printer.CurrentX = 0
        tx$ = .TextMatrix(i%, 1)
        If (tx$ = Chr$(214)) Then
            Printer.Font.Name = "Symbol"
        End If
        Printer.Print .TextMatrix(i%, 1);
        
        Printer.Font.Name = "Arial"
        
        For j% = 0 To 11
            If (j% = 0) Then
                tx$ = .TextMatrix(i%, j%)
            Else
                tx$ = .TextMatrix(i%, j% + 1)
            End If
            
            If (j% = 0) Or (j% = 1) Or (j% = 3) Then
                Printer.CurrentX = sp%(j%)
            Else
                Printer.CurrentX = sp%(j% + 1) - Printer.TextWidth(tx$)
            End If
            
'            If (j% = 5) Then SendWert# = SendWert# + CDbl(tx$)
            
            Printer.Print tx$;
        Next j%
        
        Printer.Print " "
        
        If (Printer.CurrentY > Printer.ScaleHeight - 1000) Then
            Call DruckFuss
            Call DruckKopf(header$, "", KopfZeile$)
            Printer.Font.Size = 11
        End If
    Next i%
End With
    
Y% = Printer.CurrentY
Printer.Line (0, Y%)-(sp%(12), Y%)
Printer.CurrentX = sp%(1)
Printer.Print Format(AnzBestellArtikel%, "0") + " Positionen";
tx$ = Format(GesamtWert#, "0.00")
Printer.CurrentX = sp%(12) - Printer.TextWidth(tx$)
Printer.Print tx$;

Printer.Print " "

Call DruckFuss(False)

Printer.EndDoc

Call StopAnimation(frmAction)

Call DefErrPop
End Sub

Sub PruefeTaetigkeiten()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("PruefeTaetigkeiten")
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
Dim i%, k%, gef%

DarfAlternativArtikel% = True
DarfAltlastenLoschen% = True
DarfFertigMachen% = True
DarfHinzufuegen% = True
DarfPreisKalk% = True
DarfPreisKalkBesorger% = True
DarfPreisKontrolle% = True
DarfRmKontrolle% = True
DarfWumsatz% = True


For i% = 0 To (AnzTaetigkeiten% - 1)
    gef% = False
    For k% = 0 To 49
        If (Taetigkeiten(i%).pers(k%) = ActBenutzer%) Then
            gef% = True
            Exit For
        End If
    Next k%
    If (gef% = False) Then
        If (Trim(Taetigkeiten(i%).Taetigkeit) = "Alternativ-Artikel") Then
            DarfAlternativArtikel% = False
        ElseIf (Trim(Taetigkeiten(i%).Taetigkeit) = "Altlasten löschen") Then
            DarfAltlastenLoschen% = False
        ElseIf (Trim(Taetigkeiten(i%).Taetigkeit) = "Fertigmachen") Then
            DarfFertigMachen% = False
        ElseIf (Trim(Taetigkeiten(i%).Taetigkeit) = "Hinzufügen") Then
            DarfHinzufuegen% = False
        ElseIf (Trim(Taetigkeiten(i%).Taetigkeit) = "Preiskalkulation") Then
            DarfPreisKalk% = False
        ElseIf (Trim(Taetigkeiten(i%).Taetigkeit) = "Preiskalk-Besorger") Then
            DarfPreisKalkBesorger% = False
        ElseIf (Trim(Taetigkeiten(i%).Taetigkeit) = "Preiskontrolle") Then
            DarfPreisKontrolle% = False
        ElseIf (Trim(Taetigkeiten(i%).Taetigkeit) = "RM-Kontrolle") Then
            DarfRmKontrolle% = False
        ElseIf (Trim(Taetigkeiten(i%).Taetigkeit) = "Lieferanten-Umsatz") Then
            DarfWumsatz% = False
        End If
    End If
Next i%

With frmAction
    If (DarfHinzufuegen%) Then
        .mnuBearbeitenInd(MENU_F2).Enabled = True
    Else
        .mnuBearbeitenInd(MENU_F2).Enabled = False
    End If
    .cmdToolbar(1).Enabled = .mnuBearbeitenInd(MENU_F2).Enabled

    If (DarfFertigMachen%) Then
        .mnuBearbeitenInd(MENU_SF6).Enabled = True
    Else
        .mnuBearbeitenInd(MENU_SF6).Enabled = False
    End If
    .cmdToolbar(13).Enabled = .mnuBearbeitenInd(MENU_SF6).Enabled
End With

Call DefErrPop
End Sub

Sub LoescheAlteTage()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("LoescheAlteTage")
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
Dim h$, DirMask$, EntryName$, MinTag$
Dim SearchHandle&
Dim FindDataRec As WIN32_FIND_DATA
Dim erg%, ind%, i%, k%, gef%

If (TageSpeichern% > 0) Then
    MinTag$ = Format(Now - TageSpeichern%, "YYMMDD")
    
    DirMask$ = "winw\*.*"
    
    SearchHandle& = FindFirstFile(DirMask$, FindDataRec)
    If (SearchHandle& = INVALID_HANDLE_VALUE) Then Exit Sub
    Do
        h$ = FindDataRec.cFileName
        ind% = InStr(h$, Chr$(0))
        If (ind% > 0) Then h$ = Left$(h$, ind% - 1)
        EntryName$ = h$
        
        If ((EntryName$ = ".") Or (EntryName$ = "..")) Then
        ElseIf (FindDataRec.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) Then
            If (EntryName$ < MinTag$) Then
                h$ = "winw\" + EntryName
                Call DeleteDirectory(h$, 1)
'                Debug.Print h$
            End If
        End If
        
        erg% = FindNextFile(SearchHandle, FindDataRec)
        If (erg% = 0) Then Exit Do
    Loop
    erg% = FindClose(SearchHandle&)
End If

Call DefErrPop
End Sub

Sub WumsatzEinzeln(wRec() As WumsatzStruct)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("WumsatzEinzeln")
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
Dim i%, j%, Handle%, iLief%
Dim dWert#
Dim h$

Handle% = FreeFile
Open "wumsatz.dat" For Append As #Handle%

AnzSchwellLief% = 0
        
For j% = 1 To UBound(wRec)
    iLief% = wRec(j%).Lief
    iLief% = lifzus.GetWumsatzLief(iLief%)
    
    dWert# = wRec(j%).Wert
    If (WuInDm% > 0) Then dWert# = FNX(dWert# / EURO_FAKTOR)
    
    h$ = wRec(j%).bdatum + " " + Right$(Space$(3) + Str$(iLief%), 3)
'    h$ = h$ + " " + Right$(Space$(9) + Str$(Int(wRec(j%).Wert * 100# + 0.5)), 9)
    h$ = h$ + " " + Right$(Space$(9) + Str$(Int(dWert# * 100# + 0.5)), 9)
    If (wRec(j%).Rabatt = False) Then h$ = h$ + "*"
    h$ = Left$(h$ + Space$(32), 30)
    Print #Handle%, h$
    
    If (IstSchwellLieferant%(iLief%) < 0) Then
        ReDim Preserve SchwellLief(AnzSchwellLief%)
        Call InitSchwellLief(AnzSchwellLief%)
        SchwellLief(AnzSchwellLief%).Lief = iLief%
        AnzSchwellLief% = AnzSchwellLief% + 1
    End If
Next j%

Close #Handle%

Call SpeicherSchwellwertDaten

Call DefErrPop
End Sub
                  
Sub ASumsatzEinzeln(wRec() As WumsatzStruct)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("ASumsatzEinzeln")
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
Dim i%, j%, Handle%, iLief%
Dim dWert#
Dim h$, ASumsatzName$
Dim ASumsatzDB As Database
Dim ASumsatzRec As Recordset
Dim Td As TableDef
Dim Idx As Index
Dim Fld As Field
Dim IxFld As Field

ASumsatzName$ = "ASumsatz.mdb"

On Error Resume Next
Err.Clear
Set ASumsatzDB = OpenDatabase(ASumsatzName$, False, False)
Set ASumsatzRec = ASumsatzDB.OpenRecordset("ASumsatz", dbOpenTable)
If Err.Number <> 0 Then
  On Error GoTo DefErr
  If Dir(ASumsatzName$) <> "" Then Kill ASumsatzName$
  Set ASumsatzDB = CreateDatabase(ASumsatzName$, dbLangGeneral)

  Set Td = ASumsatzDB.CreateTableDef("ASumsatz")

  Set Fld = Td.CreateField("Lieferant", dbInteger)
  Td.Fields.Append Fld
  Set Fld = Td.CreateField("Datum", dbDate)
  Td.Fields.Append Fld
  Set Fld = Td.CreateField("Wert", dbDouble)
  Td.Fields.Append Fld
  Set Fld = Td.CreateField("AbgabeSchlüssel", dbByte)
  Td.Fields.Append Fld

  ' Indizes für ASumsatz
  Set Idx = Td.CreateIndex()
  Idx.Name = "Lieferant"
  Idx.Primary = False
  Idx.Unique = False
  Set IxFld = Idx.CreateField("Lieferant")
  Idx.Fields.Append IxFld
  Set IxFld = Idx.CreateField("Datum")
  Idx.Fields.Append IxFld
  Td.Indexes.Append Idx

  ASumsatzDB.TableDefs.Append Td

  ASumsatzDB.Close
  Set ASumsatzDB = OpenDatabase(ASumsatzName$, False, False)
  Set ASumsatzRec = ASumsatzDB.OpenRecordset("ASumsatz", dbOpenTable)
End If
ASumsatzRec.Index = "Lieferant"


For j% = 1 To UBound(wRec)
    iLief% = wRec(j%).Lief
    iLief% = lifzus.GetWumsatzLief(iLief%)
    
    h$ = wRec(j%).bdatum
    
    ASumsatzRec.AddNew
    ASumsatzRec!Lieferant = iLief%
    ASumsatzRec!datum = Left$(h$, 2) + "." + Mid$(h$, 3, 2) + ".20" + Mid$(h$, 5, 2)
    ASumsatzRec!Wert = wRec(j%).Wert
    ASumsatzRec!AbgabeSchlüssel = wRec(j%).Rabatt
    ASumsatzRec.Update
Next j%

ASumsatzDB.Close

Call DefErrPop
End Sub
                  
Sub ZeigeRueckmeldungen()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("ZeigeRueckmeldungen")
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
        
frmRueckmeldungen.Show 1

Call DefErrPop
End Sub

Sub ShowDirektBezugHinweis()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("ShowDirektBezugHinweis")
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
Dim ind%, DirektKontrollStatus%, iLief%, HinweisTyp%
Dim ErstZeit%, JetztZeit%
Dim h$, IniHinweis$, HinweisText$


IniHinweis$ = HoleIniString$("DirektBezugHinweis")

Do
    ind% = InStr(IniHinweis$, ",")
    If (ind% > 0) Then
        h$ = Left$(IniHinweis$, ind% - 1)
        IniHinweis$ = Mid$(IniHinweis$, ind% + 1)
        
        iLief% = Val(Left$(h$, 3))
        HinweisTyp% = Val(Mid$(h$, 9, 1))
        
        If (HinweisTyp% = 2) Then
            ErstZeit% = Val(Mid$(h$, 5, 4))
            ErstZeit% = (ErstZeit% \ 100) * 60 + (ErstZeit% Mod 100)
            JetztZeit% = Val(Format(Now, "HHMM"))
            JetztZeit% = (JetztZeit% \ 100) * 60 + (JetztZeit% Mod 100)
        
            If (JetztZeit% < ErstZeit%) Then JetztZeit% = JetztZeit% + 24 * 60
            If ((JetztZeit% - ErstZeit%) >= DirektBezugKontrollenMinunten%) Then       '720
                HinweisTyp% = 0
            End If
        End If
        
        If (HinweisTyp% > 0) Then
            ErstZeit% = Val(Mid$(h$, 10, 4))
            ErstZeit% = (ErstZeit% \ 100) * 60 + (ErstZeit% Mod 100)
            JetztZeit% = Val(Format(Now, "HHMM"))
            JetztZeit% = (JetztZeit% \ 100) * 60 + (JetztZeit% Mod 100)
        
            If (JetztZeit% < ErstZeit%) Then JetztZeit% = JetztZeit% + 24 * 60
            If ((JetztZeit% - ErstZeit%) < DirektBezugWarnungMinuten%) Then     '10
                HinweisTyp% = 0
            End If
        End If
            
        If (HinweisTyp% > 0) Then
            lif.GetRecord (iLief% + 1)
            h$ = Trim$(lif.kurz)
            If (h$ = String$(Len(h$), 0)) Then h$ = ""
            If (h$ = "") Then
                h$ = "(" + Str$(iLief%) + ")"
            End If
    
            If (HinweisTyp% = 3) Then
                HinweisText$ = "Direktlieferant " + h$ + "überprüfen !"
                Call iMsgBox(HinweisText$, vbInformation, "Prüfung Direktlieferant")
            Else
                HinweisText$ = "Direktbezug für Lieferant" + vbCrLf + vbCrLf + h$ + vbCrLf$ + vbCrLf$
                If (HinweisTyp% = 1) Then
                    h$ = "ist DRINGEND zu kontrollieren !"
                Else
                    h$ = "wäre zu  kontrollieren !"
                End If
                HinweisText$ = HinweisText$ + h$
                Call iMsgBox(HinweisText$, vbInformation, "Kontrollen Direktbezug")
            End If
            
            IniHinweis$ = HoleIniString$("DirektBezugHinweis")
            h$ = Format(iLief%, "000") + "-"
            ind% = InStr(IniHinweis$, h$)
            If (ind% > 0) Then
                Mid$(IniHinweis$, ind% + 9, 4) = Format(Now, "HHMM")
                Call SpeicherIniString%("DirektBezugHinweis", IniHinweis$)
            End If
            
        End If
    Else
        Exit Do
    End If
Loop

Call DefErrPop
End Sub



Function ZeigeDirektAufteilung%(Optional iMenge% = 0, Optional AnzeigeModus% = 0)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("ZeigeDirektAufteilung")
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
Dim i%, iProfilNr%, ActBm%, OrgBm%, iSortNr%, spBreite%, iBm%, iRest%, iSum%, MaxRest%, MaxRow%
Dim maxDs%, iLief%
Dim l&
Dim multi!, bm!
Dim h$, pzn$, sName$, sLief$

pzn$ = KorrPzn$
OrgBm% = 0
iLief% = Lieferant%
    
With frmAction.flxarbeit(0)
    If (iMenge% > 0) Then
        ActBm% = iMenge%
    ElseIf (ProgrammChar$ = "B") Then
        ActBm% = Val(.TextMatrix(.row, 5)) + Val(.TextMatrix(.row, 6)) 'BM+NM
    ElseIf (ProgrammChar$ = "W") Then
        ActBm% = Val(.TextMatrix(.row, 8)) 'LM
        iLief% = Asc(.TextMatrix(.row, 16))
    End If
End With

With frmAction.flxDirektAufteilung
    .Cols = 6
    .Rows = 1
    .FixedRows = 0
    .FixedCols = 0
    .Rows = 0
    .SelectionMode = flexSelectionByRow
    
    .ColWidth(0) = 0    'TextWidth("999999")
    .ColWidth(1) = 0    'TextWidth("999999")
    .ColWidth(2) = frmAction.TextWidth("999999")
    .ColWidth(3) = frmAction.TextWidth(String(30, "X"))
    .ColWidth(4) = frmAction.TextWidth("99999")
    .ColWidth(5) = 0    'frmAction.TextWidth("99999")

    spBreite% = 0
    For i% = 0 To .Cols - 1
        spBreite% = spBreite% + .ColWidth(i%)
    Next i%
    .Width = spBreite% + 90

    .Clear
End With
    
'Partner-Apos
If (FremdPznOk%) Then
    DirektAufteilungRec.Seek "=", pzn$
    If (DirektAufteilungRec.NoMatch = False) Then
        Do
            If (DirektAufteilungRec.EOF) Then
                Exit Do
            End If
            If (DirektAufteilungRec!pzn <> pzn$) Then
                Exit Do
            End If
            
            If (DirektAufteilungRec!Lief = iLief%) Then
                iSortNr% = 999
                sName$ = ""
                sLief$ = ""
                iProfilNr% = DirektAufteilungRec!ProfilNr
                If (iProfilNr% > 0) Then
                    h$ = Format(iProfilNr%, "000")
                    If (InStr(OpDirektPartner$, h$ + ",") > 0) Then
                        OpPartnerRec.Seek "=", iProfilNr%
                        If (OpPartnerRec.NoMatch = False) Then
                            iSortNr% = OpPartnerRec!IntSortNr
                            sName$ = OpPartnerRec!Name
                            sLief$ = Format(OpPartnerRec!IntLiefNr, "0")
                        Else
                            sName$ = "(" + Format(iProfilNr%, "0") + ")"
                        End If
                        
                        If (OrgBm% = 0) Then
                            OrgBm% = DirektAufteilungRec!BvGes
                            If (OrgBm% = 0) Then
                                multi! = 1
                            Else
                                multi! = ActBm% * 1# / OrgBm%
                            End If
                        End If
                        iSum% = iSum% + DirektAufteilungRec!bv
                        
                        If (OrgBm% = ActBm%) Then
                            bm! = DirektAufteilungRec!bv
                        Else
                            bm! = DirektAufteilungRec!bv * multi!
'                                bm! = Int(DirektAufteilungRec!bv * multi! + 0.501)
                        End If
                        iBm% = Int(bm! + 0.501)
                        iRest% = Int(bm! * 100# + 0.501) Mod 100
                        
                        h$ = Format(iSortNr%, "0")
                        h$ = h$ + vbTab + Format(iProfilNr, "0")
                        h$ = h$ + vbTab + sLief$
                        h$ = h$ + vbTab + sName$
                        h$ = h$ + vbTab + Format(iBm%, "0")
                        h$ = h$ + vbTab + Format(iRest%, "0")
                        frmAction.flxDirektAufteilung.AddItem h$
                    End If
                End If
            End If
            
            DirektAufteilungRec.MoveNext
        Loop
    End If
End If
    
With frmAction.flxDirektAufteilung
    iSortNr% = 0
    iProfilNr = 0
    sLief$ = ""
    sName$ = "Eigenbedarf"
    If (.Rows = 0) Or (OrgBm% = 0) Then
        bm! = ActBm%
    Else
        bm! = OrgBm% - iSum%
        bm! = bm! * multi!
'        bm% = Int(bm% * multi! + 0.501)d
    End If
    iBm% = Int(bm! + 0.501)
    iRest% = Int(bm! * 100# + 0.501) Mod 100
    
    h$ = Format(iSortNr%, "0")
    h$ = h$ + vbTab + Format(iProfilNr, "0")
    h$ = h$ + vbTab + sLief$
    h$ = h$ + vbTab + sName$
    h$ = h$ + vbTab + Format(iBm%, "0")
    h$ = h$ + vbTab + Format(iRest%, "0")
    .AddItem h$
    
    .row = 0
    .col = 0
    .RowSel = .Rows - 1
    .ColSel = .col
    .Sort = 5
    
    .Height = .RowHeight(0) * .Rows + 90
    .ZOrder 0
    
    Do
        iSum% = 0
        For i% = 0 To (.Rows - 1)
            iSum% = iSum% + Val(.TextMatrix(i%, 4))
        Next i%
        
        If (iSum% = ActBm%) Then
            Exit Do
        ElseIf (iSum% < ActBm%) Then
            MaxRow% = -1
            MaxRest% = 0
            For i% = 0 To (.Rows - 1)
                iRest% = Val(.TextMatrix(i%, 5))
                If (iRest% < 50) And (iRest% > MaxRest%) Then
                    MaxRow% = i%
                    MaxRest% = iRest%
                End If
            Next i%
            If (MaxRow% >= 0) Then
                .TextMatrix(MaxRow%, 4) = Format(Val(.TextMatrix(MaxRow%, 4)) + 1, "0")
                .TextMatrix(MaxRow%, 5) = ""
            Else
                Exit Do
            End If
        Else
            MaxRow% = -1
            MaxRest% = 99
            For i% = 0 To (.Rows - 1)
                iRest% = Val(.TextMatrix(i%, 5))
                If (iRest% >= 50) And (iRest% < MaxRest%) Then
                    MaxRow% = i%
                    MaxRest% = iRest%
                End If
            Next i%
            If (MaxRow% >= 0) Then
                .TextMatrix(MaxRow%, 4) = Format(Val(.TextMatrix(MaxRow%, 4)) - 1, "0")
                .TextMatrix(MaxRow%, 5) = ""
            Else
                Exit Do
            End If
        End If
    Loop
    
    ZeigeDirektAufteilung% = Val(.TextMatrix(0, 4))
    
    .Visible = (AnzeigeModus% = 0)
    
    If (AnzeigeModus% > 0) Then
        For i% = 1 To (.Rows - 1)
            iBm% = Val(.TextMatrix(i%, 4))
            If (iBm% > 0) Then
                h$ = Space$(10)
                iProfilNr% = Val(.TextMatrix(i%, 1))
                OpPartnerRec.Seek "=", iProfilNr%
                If (OpPartnerRec.NoMatch = False) Then
                    h$ = Left$(OpPartnerRec!kurz + Space$(10), 10)
                End If
                
                If (AnzeigeModus% = 1) Then
                    WuRec.KurzEmpfänger = h$
                    WuRec.KurzAbsender = Left$(IdBeiPartnern$ + Space$(10), 10)
                    WuRec.pzn = pzn$
                    WuRec.menge = Right$(Space$(4) + Format(iBm%, "0"), 4)
                    WuRec.Rmenge = Right$(Space$(4) + Format(iBm%, "0"), 4)
                    WuRec.datum = Format(Now, "DDMMYY")
                    WuRec.Verfall = Left$(Trim(ww.WuAblDatum) + Space$(6), 6)
                    WuRec.Lief = Right$(Space$(3) + Format(ww.Lief, "0"), 3)
                    
                    Get WARE_HANDLE%, 1, WuSatz$
                    maxDs% = CVI(WuSatz$) + 1
                    WuSatz$ = MKI(maxDs%) + String(48, Chr$(0))
                    Put WARE_HANDLE%, 1, WuSatz$
                    Put #WARE_HANDLE%, maxDs% + 1, WuRec
                Else
                    NnEkRec.KurzEmpfänger = h$
                    NnEkRec.KurzAbsender = Left$(IdBeiPartnern$ + Space$(10), 10)
                    NnEkRec.pzn = pzn$
                    NnEkRec.Lmenge = Right$(Space$(4) + Format(Abs(ww.WuLm), "0"), 4)
                    NnEkRec.Rmenge = Right$(Space$(4) + Format(Abs(ww.WuRm), "0"), 4)
                    NnEkRec.BelegDatum = sdate(ww.WuBelegDatum)
                    NnEkRec.BelegNr = ww.WuBeleg
                    NnEkRec.nnek = Right$(Space$(8) + Format(Abs(iMenge%), "0"), 8)
                    NnEkRec.nnart = Right$(Space$(3) + Format(ww.WuNNart, "0"), 3)
                    NnEkRec.Lief = Right$(Space$(3) + Format(ww.Lief, "0"), 3)
                    
                    Get NNEK_HANDLE%, 1, NnEkSatz$
                    maxDs% = CVI(NnEkSatz$) + 1
                    NnEkSatz$ = MKI(maxDs%) + String(57, Chr$(0))
                    Put NNEK_HANDLE%, 1, NnEkSatz$
                    Put #NNEK_HANDLE%, maxDs% + 1, NnEkRec
                End If
            End If
        Next i%
    End If
End With
   
    
With frmAction
    .flxDirektAufteilung.Top = .picBack(0).Top + .flxarbeit(0).Top + .flxarbeit(0).RowPos(.flxarbeit(0).row) + .flxarbeit(0).RowHeight(0)
    .flxDirektAufteilung.Left = .picBack(0).Left + .flxarbeit(0).Left + .flxarbeit(0).ColPos(0)
End With


Call DefErrPop
End Function




