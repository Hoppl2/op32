Attribute VB_Name = "modWbestk2"
'1.0.51 31.03.05 AE analog WinWawi2
'1.0.50 22.03.05 AE analog WinWawi2
'1.0.49 16.03.05 AE Merkzettel abhängig von Fistam (wenn 't' - MDB, sonst stammlos.dat)
'1.0.48 01.02.05 AE Merkzettel jetzt als MDB
'1.0.47 25.11.04 AE mitkompiliert wegen Neuer DLL
'1.0.46 15.11.04 AE Direktbezug adaptiert: CalcAlleDirektAngebote wieder aktiviert;
'                   den Teil, wo für DirLiefs wenn ein günstiges Angebot für einen GH gefunden wird, die BM uU auf 0 gesetzt wird, unter Kommentar gesetzt!
'1.0.45 AE  111004  Kontrollgrund 'Übl.Lieferant' richtiggestellt
'                   Kontrollgrund 'Interne Streichung' eingeführt
'1.0.44 AE  300804  Mitkompiliert wegen WinWawi
'1.0.43 AE  120804  Mitkompiliert wegen WinWawi
'1.0.42 AE  020804  Mitkompiliert wegen WinWawi
'1.0.41 GS  130704  Etiketten Ö: wahlweise doch alle (abhängig von INI-Param)
'1.0.40 AE  110604  wenn AutomatikBest für DirLief, dann von Verbundbestellung ausnehmen (IstAutoDirektLief%)
'1.0.39 AE  020604  wegen winwawi mitkompiliert (EasyLink - DirektAufteilung)
'1.0.38 AE  110504  wegen winwawi mitkompiliert (ManuellLief)
'1.0.37 AE  040422  Für Österreich: wenn besetzt, wird nach 90 Sek nochmals gewählt (insgesamt 3x)
'                   Fenster in Vordergrund bringen mittels SetWindowPos
'                   Bereitstellen Etiketten: nut wenn Artikel kalkulierbar
'1.0.36 AE  040322  Neue Kontrollgründe: AM RX, AM NON RX, NICHTAM
'                   WÜ: nach man. Ablaufbestätigung in LM nächste Zeile springen
'                   Rabatttabellen: bei Ausnahmen zus. 'AM RX'
'                   BM/EK-tabelle um Gruppen erweitert
'1.0.35 AE  040317  Neuer Kontrollgrund: Interim
'1.0.34 AE  040309  Einbau der Auswertung der AutomatenLiefs
'1.0.33 AE  040304  2-stellige Lagercodes eingeführt: Bedingungen, Infobereich
'                   neuer Kontrollgrund: Rezeptpfl. AM
'1.0.30 AE  191203  Kontrollgrund Text-Besorger
'1.0.29 AE  111203  Interimsbestellung: keine Angebote prüfen, Anzeige 'Interim' in Zeile
'1.0.28 AE  211103  DAO351 durch DAO360 ersetzt; wegen Einheitlichkeit Programme
'1.0.27 AE  101103  neuer Sound wwfertig
'                   Für Ö: AbsagenEntfernen
'1.0.25 AE  301003  Für Ö: feiertage, nn-aep speichern. Mitkompiliert.
'1.0.24 AE  141003  Für Ö: Einbau Merkzettel ohne Indexdateien ...
'1.0.23 AE  180903  Einbau Partnerapos für Direktbezug
'1.0.22 AE  120903  Interimsbestellung: auch wenn Artikel für DirektLief in BESTELLUNG (VorhTest)
'1.0.21 AE  090903  Teilbestellungen bei Ladenhüter-Bestellungen
'1.0.19 AE  020703  Neue op32.dll
'1.0.18 AE  230603  Neue op32.dll
'1.0.16 AE  080403  Direktbezug- 'Direktbezug möglich ab ' eingebaut
'1.0.10 AE  271002  Direktbezug
'1.0.8 AE   060902  Direktbezug
'1.0.6 AE   020819  Erweiterungen Direktbezug
'1.0.5 AE   020812  Erweiterungen Direktbezug
'1.0.4 AE   020717  kompiliert wegen neuer DLL
'1.0.2 AE   010508
' F380 in SucheSendeArtikel: anscheinend kommt Timer während TimerAbarbeitung; deshalb in
' tmrAction_timer timer disabled und enabled, genaue Überlegungen wo

Option Explicit


Public Const RUFZEITENANZEIGE = 0
Public Const SENDEHINWEIS = 1
Public Const SENDEANZEIGE = 2
Public Const DIREKTBEZUG = 3

Public Const ZEILEN = 0
Public Const LIEFERANTENWAHL = 1
Public Const WAWIOPTIONEN = 2
Public Const INFOLAYOUT = 3
Public Const INFOGENAU = 4
Public Const ARTIKELDAZU = 5
Public Const BESTVORS = 6
Public Const BESTSENDEN = 7
Public Const RUECKMELDUNGEN = 8
Public Const LIEFERANTENZUORDNEN = 9

Public Const WINWAWI_INI = "\USER\WINWAWI.INI"
Public Const GLOBAL_SECTION = "Global"
Public Const INFO_BESTELLUNG_SECTION = "Infobereich Bestellung"
Public Const INFO_ERFASSUNG_SECTION = "Infobereich Erfassung"

Public Const MAX_INFO_ZEILEN = 6
Public Const MIN_INFO_ZEILEN = 4



'Public clsfabs As New clsFabsKlasse

Public ProgrammRec(1) As ProgrammStruct


'Public ass As StatistikStruct
'Public ast As StammStruct
'Public bek As BekartStruct
'Public bekOrg As BekartStruct
'Public bekFirst As BekartErstSatz
'Public bekFix As String * 128
'Public best As BestelStruct
'Public bestFirst As BestelErstSatz
'Public nachb As NachbeStruct
'Public lif As LieferantenStruct
'Public etik As EtikettenStruct
'Public etikFirst As EtikettenErstSatz
'Public StammlosFix As String * 65
'Public BabsageFix As String * 100
'Public ArtTextFix As String * 192
'Public GhAnbotRec As GhAngeboteStruct
'Public LocalAnbotRec(50) As LocalAngeboteStruct
'Public LocalAnbotStr(50) As LocalAngeboteString
'Public Herst As HerstellerStruct
'Public PasswordRec As PasswordStruct

Public ProgrammChar$

Global ASTAMM%
Global ASTATIS%
Global BEKART%
Global BESTEL%
Global LDATEI%
Global POSDRUCK%
Global GHANBOT%
'Global Arttext%

Global buf$
Global AnzAnzeige%
Global grecno&

Public TaxeDB As Database
Public TaxeRec As Recordset
Public TaxeRec2 As Recordset

Public MerkzettelDB As Database
Public MerkzettelRec As Recordset
Public MerkzettelParaRec As Recordset

Public Personal$(50)

Public ProgrammModus%
Public TaxeOk%
Public UebergabeStr$
Public BildFaktor!
Public ZuzFeld!(6)

Public FabsErrf%
Public FabsRecno&

Public fName$
Public fSize%

Public MaxRowsArbeit%(11)
Public MaxRowsInfo%(11)
Public MaxTextRows%

Public DruckSeite%

Public OrgFormWidth%, OrgWindowState%
Public LetztEingabe$
Public LetztEingabeRow%

Public FarbeArbeit&, FarbeInfo&
Public FarbeRahmen%

Public PROTOKOLL%, prot%

Public AngeboteIndex&

Public HochfahrenAktiv%
Public DDEstr$

Public AnzGh%

Public UserSection$

Public Lieferant%
Public Exklusiv%
Public VonBM%, BisBM%
Public WaGr$, auto$
Public MarkWert#, GesamtWert#
Public AddMarkWert#, AddGesamtWert#
'Public iSRT%(1000)

Public KeinRowColChange%

Public BmMulti%

Public GhAnbotAuswahl%
Public NNAEP#
Public menge%, NrMenge%
Public angebot%
Public lfa%

Public GhAnbotMax&
Public BMopt!, BMast%

Public NrWert#, BarRab#, LaKo#, PeKo#, Manko#, StaffelAb#, aep#

Public AnzInfoZeilen%(1)

Public InRueckmeldungen%

Public ErfassungBestmen%, ErfassungNatu%
Public ErfassungAsatz%, ErfassungSsatz%
Public ErfassungPzn$

Public AnzLocalAngebote%
Public DarfInfoBereich%

Public GlobBekMax%
Public BekUpdate%

Public HintergrundSsatz%, HintergrundAnz%

Public WochenTag$(6)

Public AutomaticSend%
Public AutomaticInd%
Public VorAutomaticLief%

Dim fFile%


''''''''''''''''''''''''''''''''''''''''''''''''''
Public ast As clsStamm
Public ass As clsStatistik
Public taxe As clsTaxe
Public lif As clsLieferanten
Public lifzus As clsLiefZusatz
Public ww As clsWawiDat
Public bek As clsBekart
Public nachb As clsNachbearbeitung
Public wu As clsWÜ
Public BESORGT As clsBesorgt
Public Merkzettel As clsMerkzettel
Public absagen As clsAbsagen
Public kiste As clsKiste
Public arttext As clsArttext
Public etik As clsEtiketten
Public ang As clsAngebote
Public ManAng As clsManuellAngebote
Public para As clsOpPara
Public wpara As clsWinPara
Public rabtab As clsRabattTabellen
Public rk As clsRueckKauf


Public ProgrammNamen$(3)
Public ProgrammTyp%

Public NaechstLiefernderLieferant%
Public AktUhrzeit%

Public NurZuKontrollierende%
''''''''''''''''''''''''''''''''''''''''''''''''''''

Public EditModus%
Public EditErg%
Public EditTxt$

Public LiefWechselOk%


Public ActProgram As Object

Public INI_DATEI As String

'Public DirektBezugAktiv%




Private Const DefErrModul = "wbestk2.bas"

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
Dim h$

ProgrammChar$ = "B"

Select Case (ProgrammChar$)
    Case "B"
        
        ProgrammChar$ = "2"
        Call EntferneGeloeschteZeilen(1)
        Call UpdateBekartDat(-1, False)
        Call frmAction.WechselModus(RUFZEITENANZEIGE)
        
End Select

InitProgramm% = True
Call DefErrPop
End Function

Sub CloseProgramm()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("CloseProgramm")
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

'Call iClose(ASTAMM%)
'Call iClose(ASTATIS%)
'Call iClose(BEKART%)
'Call iClose(BESTEL%)
'Call iClose(POSDRUCK%)
'Call iClose(Arttext%)

Call DefErrPop
End Sub

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

ProgrammNamen$(0) = "Wbestk2"

WochenTag$(0) = "Montag"
WochenTag$(1) = "Dienstag"
WochenTag$(2) = "Mittwoch"
WochenTag$(3) = "Donnerstag"
WochenTag$(4) = "Freitag"
WochenTag$(5) = "Samstag"
WochenTag$(6) = "Sonntag"
    

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
Dim i%, erg%, tmp%, ind%, t%, m%, j%, c%, sMax%
Dim h$, SQLStr$, DirRet$, FuDate$, CdInfoVer$, s$

If (App.PrevInstance) Then End

'If (Dir$("fistam.dat") = "") Then ChDir "\user"
ChDir "\user"
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
Set arttext = New clsArttext
Set etik = New clsEtiketten
Set ang = New clsAngebote
Set ManAng = New clsManuellAngebote
Set para = New clsOpPara
Set wpara = New clsWinPara
Set rabtab = New clsRabattTabellen
Set rk = New clsRueckKauf

Set ActProgram = New clsWbestk2

UserSection$ = "Computer" + Format(Val(para.User))

Call wpara.HoleWindowsParameter
Call InitMisc


frmAction.Show
Call ShowSysTrayIcon(frmAction, 1, frmAction.Icon, RTrim$(frmAction.Caption))
    
'Call StartAnimation("Parameter werden eingelesen ...")

Call para.HoleFirmenStamm
Call para.AuslesenPdatei


erg% = TestModemPar%
If (erg% = False) Then End



ast.OpenDatei
ass.OpenDatei
ass.GetRecord (1)
sMax% = ass.erstmax

If (BestVorsPeriodisch%) Then
    HintergrundAnz% = sMax% \ (BestVorsPeriodischMinuten% * 3) 'timer alle 5 Sek, suchen nur alle 20 Sek!
End If

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
FremdPznOk% = OpenOpConnect(FremdPznDB, OpPartnerDB)
If (FremdPznOk%) Then
    Set FremdPznRec = FremdPznDB.OpenRecordset("Artikel", dbOpenTable)
    FremdPznRec.Index = "Unique"
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
'Call EinlesenLieferanten

ang.OpenDatei
ManAng.OpenDatei

arttext.OpenDatei

'BESORGT.OpenDatei
erg% = OpenCreateMerkzettelDB%

rk.OpenDatei
If (rk.DateiLen = 0) Then
    rk.erstmax = 0
    rk.erstlief = 0
    rk.erstcounter = 0
    rk.erstrest = String(rk.DateiLen, 0)
    rk.PutRecord (1)
End If

ManuellLief% = 0

IstAutoDirektLief% = True

AnzBestellWerteRows% = 1

Call HoleIniKontrollen
Call HoleIniZuordnungen
Call HoleIniRufzeiten
Call HoleRufzeitenLieferanten
Call HoleIniFeiertage
            
lifzus.OpenDatei
If (lifzus.DateiLen = 0) Then
    Call lifzus.InitDatei
    If (para.Land = "D") Then
        Call rabtab.KonvertTabelle
    End If
    Call KonvSchwellwerte
End If
'Call lifzus.EinlesenLiefFuerHerst

Call InitIsdnAnzeige

'ModemOk% = TestModemPar%

Call InitWumsatzDat

erg% = InitProgramm%

Call DefErrPop: Exit Sub
    
ErrorHandler:
    erg% = Err
    If ((erg% > 0) And (erg% <> 3024) And (erg% <> 3044)) Then
        Call MsgBox("Fehler" + Str$(Err) + " beim Öffnen der Taxe " + h$ + vbCr + Err.Description, vbCritical, "OpenDatabase")
        End
    End If
    Err = 0
    Resume Next
    Return

Call DefErrPop
End Sub

'Sub Main()
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Call DefErrFnc("Main")
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
'Dim erg%, tmp%, ind%, t%, m%, j%, c%
'Dim l&
'Dim h$, SQLStr$, DirRet$, FuDate$, CdInfoVer$, s$
'Dim NotifyIconDataRec As NOTIFYICONDATA
'
'
'HochfahrenAktiv% = True
'
'Call InitMisc
'
'OrgWindowState% = -1
'Call frmAction.InitGrafik
'
'
'Call ShowSysTrayIcon(frmAction, 1, frmAction.Icon, RTrim$(frmAction.Caption))
'
'
'
'FabsErrf% = clsfabs.Cmnd("Y\2", FabsRecno&)
'FabsErrf% = clsfabs.Cmnd("O\astatis.idx\20", FabsRecno&)
'If (FabsErrf%) Then
'    erg% = MsgBox("Fehler " + Str$(FabsErrf%) + " beim Öffnen der Statistik-Daten", vbInformation)
'End If
'FabsErrf% = clsfabs.Cmnd("O\artikel1.idx\19", FabsRecno&)
'If (FabsErrf%) Then
'    erg% = MsgBox("Fehler " + Str$(FabsErrf%) + " beim Öffnen der Stamm-Daten", vbInformation)
'End If
'FabsErrf% = clsfabs.Cmnd("O\artikel2.idx\18", FabsRecno&)
'
'FabsErrf% = clsfabs.Cmnd("O\besorgt.ix1\17", FabsRecno&)
'
'FabsErrf% = clsfabs.Cmnd("O\arttext.ix1\16", FabsRecno&)
'
'
'Call frmAction.LadeAnzeige(100)
'
'frmAction.picToolbar.Visible = True
'frmAction.picStartBack.Visible = False
'
'
'TaxeOk% = False
'
'erg% = 0
'h$ = TaxeLw$ + ":\taxe\taxe.mdb"
'On Error GoTo ErrorHandler
'Set TaxeDB = OpenDatabase(h$, False, True)
'If (erg% > 0) Then
'    erg% = MsgBox("Die Windows-Taxe existiert auf Laufwerk " + TaxeLw$ + ": nicht !", vbInformation)
'End If
'On Error GoTo 0
'
'TaxeOk% = True
'Set TaxeRec = TaxeDB.OpenRecordset("Taxe", dbOpenTable)
'
'
'HochfahrenAktiv% = False
'
'erg% = InitProgramm%
'
'If (DDEstr$ <> "") Then
'    Call frmAction.DecodDDE(DDEstr$, c%)
'End If
'
'Call DefErrPop: Exit Sub
'
'ErrorHandler:
'    erg% = Err
'    If ((erg% > 0) And (erg% <> 3024) And (erg% <> 3044)) Then
'        Call MsgBox("Fehler" + Str$(Err) + " beim Öffnen der Taxe " + h$ + vbCr + Err.Description, vbCritical, "OpenDatabase")
'        End
'    End If
'    Err = 0
'    Resume Next
'    Return
'
'ErrorHandler2:
'    erg% = Err
'    Call MsgBox("Fehler" + Str$(Err) + " beim Öffnen von " + h$ + vbCr + Err.Description, vbCritical, "OpenDatabase")
'    Err = 0
'    Resume Next
'    Return
'
'Call DefErrPop
'End Sub

Sub ProgrammEnde()
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

wpara.ExitEndSub

TaxeDB.Close

If (InStr(para.Benutz, "t") > 0) Then
    MerkzettelDB.Close
Else
    BESORGT.CloseDatei
End If

ast.CloseDatei
ass.CloseDatei
lif.CloseDatei
lifzus.CloseDatei
ww.CloseDatei
bek.CloseDatei
nachb.CloseDatei
wu.CloseDatei
'BESORGT.CloseDatei
arttext.CloseDatei
etik.CloseDatei
ang.CloseDatei
ManAng.CloseDatei
rk.CloseDatei

ast.FreeClass
ass.FreeClass
lif.FreeClass
lifzus.FreeClass
ww.FreeClass
bek.FreeClass
nachb.FreeClass
wu.FreeClass
BESORGT.FreeClass
absagen.FreeClass
arttext.FreeClass
etik.FreeClass
ang.FreeClass
ManAng.FreeClass
rabtab.FreeClass
rk.FreeClass

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
Set arttext = Nothing
Set etik = Nothing
Set ang = Nothing
Set ManAng = Nothing
Set para = Nothing
Set wpara = Nothing
Set rabtab = Nothing
Set rk = Nothing

Call frmAction.frmActionUnload
    
End
Call DefErrPop
End Sub


Function GetTaskId(ByVal TaskName As String) As Long
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("GetTaskId")
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
Dim CurrWnd As Long
Dim Length As Integer
Dim ListItem As String
Dim X As Long
Dim ThreadID As Long
Dim i&

GetTaskId = 0

TaskName = UCase(TaskName)
CurrWnd = GetWindow(frmAction.hWnd, GW_HWNDFIRST)
Do While (CurrWnd <> 0)
    Length = GetWindowTextLength(CurrWnd)
    ListItem = Space(Length + 1)
    Length = GetWindowText(CurrWnd, ListItem, Length + 1)
    If (Length > 0) Then
        If (UCase(Left(ListItem, Len(TaskName))) = TaskName) Then
            X = GetWindowThreadProcessId(CurrWnd, ThreadID)
'            Print #PROTOKOLL%, ListItem; ThreadID
            GetTaskId = ThreadID
            Exit Do
        End If
    End If
    CurrWnd = GetWindow(CurrWnd, GW_HWNDNEXT)
Loop

Call DefErrPop
End Function

Function EntferneTask(ByVal TaskName As String) As Long
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("EntferneTask")
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
Dim CurrWnd As Long
Dim Length As Integer
Dim ListItem As String
Dim X As Long
Dim ThreadID As Long
Dim ProcHandle&
Dim i&

EntferneTask = 0

TaskName = UCase(TaskName)
CurrWnd = GetWindow(frmAction.hWnd, GW_HWNDFIRST)
Do While (CurrWnd <> 0)
    Length = GetWindowTextLength(CurrWnd)
    ListItem = Space(Length + 1)
    Length = GetWindowText(CurrWnd, ListItem, Length + 1)
    If (Length > 0) Then
        If (UCase(Left(ListItem, Len(TaskName))) = TaskName) Then
            X = GetWindowThreadProcessId(CurrWnd, ThreadID)
            ProcHandle& = OpenProcess(PROCESS_SYNCHRONIZE Or PROCESS_TERMINATE, False, ThreadID)
            X = PostMessage(CurrWnd, WM_CLOSE, 0, 0)
            X = WaitForSingleObject(ProcHandle&, 5000&)
            
'            Print #PROTOKOLL%, ListItem; ThreadID; ProcHandle&; X
            EntferneTask = ThreadID
            
            If (X <> 0) Then
                X = TerminateProcess(ProcHandle&, 0)
            End If
            Exit Do
        End If
    End If
    CurrWnd = GetWindow(CurrWnd, GW_HWNDNEXT)
Loop

Call DefErrPop
End Function

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

Call DefErrPop
End Sub


