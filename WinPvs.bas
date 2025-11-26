Attribute VB_Name = "modWinPvs"
Option Explicit

'1.0.16 20.04.06 GS Bar-Minuskunden: nicht, wenn v.storno
'1.0.15 30.01.06 GS Kunden ohne Kassenrezept und mind. 1 negativen Artikel zählen negativ bei Kundenanzahl
'1.0.14 25.11.04 AE neue op32.dll
'1.0.13 12.11.04 GS Create ApoControlDB gelöscht -> control.exe
'       rausgestellte RezArtikel (Ö) berücksichtigt
'1.0.12 26.10.04 GS ApoControl
'1.0.11 09.07.04 GS Div/Null in AuswertungDrucken behoben
'1.0.10 16.01.04 AE NNEK eingebaut wegen F91 in Matchcode
'1.0.9  11.12.03 AE Feiertage an Österreich angepasst
'1.0.8  21.11.03 AE DAO351 durch DAO360 ersetzt; CreateDatabase mit Default-Version
'1.0.7  29.10.03 GS HoleVK: bei Vorschau letzten VK-Satz nicht suchen, sondern maxVK
'1.0.6  23.6.03 AE Neu Kompiliert wegen op32.dll
'1.0.5  13.5.03 AE Richtigstellkung SucheDatum in PBA (tat bei manchen Auswertungen nix!)
'1.0.4  7.4.03 AE Richtigstellung Anzahl Wochen für Auswertung PBA
'                 F6 bei AuswertungRechnen behoben
'1.0.3  1.4.03 GS Überleitung: Beginn bei VK-Satz 2 statt 1; cntRez% weg, da Überlauf

Private Const DefErrModul = "WINPVS.BAS"

Type pssType
  KZ As String * 1
  datum As String * 2
  zeit As Integer
  LaufNr As Integer
  RezNr As Integer
  User As Integer
  BS As Byte
  pzn As String * 7
  Pnr As Byte
  RezEnde As Integer
  KuEnde As Integer
  wg As Byte
  tLager As Byte
  Lac As String * 1
  Preis As String * 8
  mwst As String * 1
  pFlag As Byte
  RabProz As Integer
  RabBetr As String * 8
  Multi As Integer
  Basis As String * 8
  sPrm As String * 4
  rest As String * 3
End Type

Type pss2Type
  lAusDat As String * 2
  lAusTim As Integer
  lAusVon As String * 2
  lAusBis As String * 2
  UebOK As Integer
  UebErst As Integer
  lUebDat As String * 2
  lUebTim As Integer
  lAnz As String * 4
  gAnz As String * 4
  rest As String * 40
End Type

Type pss3Type
  tENDE As Integer
  GndPrm As Integer
  PrmBas As String * 1
  SndPrm As Integer
  ZVkPrm As Integer
  RabAbzg As String * 1
  Normtag As Integer
  PauseKl As Integer
  geraete As Integer
  summen As Integer
  params As Integer
  Detail As Integer
  zeiten As Integer
  Liefauch As String * 1
  TagAnf As Integer
  TagEnd As Integer
  PauseGr As Integer
  PrivRez As String * 1
  rest As String * 32
End Type


'DS 4 - 13
Type pss4Type
  ind As String * 1
  wGrp As String * 25
  tLager As String * 4
  LgCode As String * 10
  vonAVP As Integer
  bisAVP As Integer
  vonSP As Integer
  bisSP As Integer
  RpPfl As String * 1
  Geräte As String * 10
  rest As String * 5
End Type


Type FeiertageStruct
    Name As String * 25
    KalenderTag As Date
    Aktiv As String * 1
End Type

Type OffenStruct
    von(1) As Integer
    Bis(1) As Integer
    iAbOffen As Integer
    iBisOffen As Integer
End Type

Public p As pssType
Public p2 As pss2Type
Public p3 As pss3Type
Public p4 As pss4Type
Public fpss As Long

Dim mwst%(4)
Const MAXPERSONAL% = 50

Public MitArb$(MAXPERSONAL%), iMitArb(MAXPERSONAL%) As Boolean
Dim geraete As Boolean, summen As Boolean, params As Boolean, Detail As Boolean
Public NORMTAGZEIT%, PauseKl%, PauseGr%
'Public TagAnf%, TagEnd%
Public chkSPANNE As Boolean
Public chkTARA As Boolean
Public chkAVP As Boolean
Public chkREZPF As Boolean
Public chkLGCOD As Boolean
        
Global operator$(10), wg$(10), wg1$(10), tLager$(10), Geräte$(10), vonSP#(10), bisSP#(10), LgCode$(10), Rp$(10), vonAVP#(10), bisAVP#(10)
        
Dim gAnzNormTage!
Public Vorschau As Boolean
Public vonAuswD As Date, bisAuswD As Date
Public ErstesMal%, UebDatBis As Date
Public abbruch As Boolean

Dim anzahl%
        
Public Const HAT_WERTE_SPALTE% = 28

Const MAX_FEIERTAGE = 20
Dim Feiertage(MAX_FEIERTAGE) As FeiertageStruct
Public aTage%
Const aGeräte% = 10
    'Rechenfeld-Tabellen dynamisch !
    '(Tag/Seite/Mitarb)
'Dim Basis#(aTage%, 1, MAXPERSONAL%)            'PrämienBasis
Dim Basis#()            'PrämienBasis
Dim Zusatz#()            'Zusatzverk.Basis
Dim Sonder#()            'SonderPrämienBasis
Dim AnzKd!()            'Anzahl Kunden
Dim AnzPKd!()           'Anzahl PrämKunden
Dim AnzRp!()            'Anzahl Rezepte
Dim AnzBar!()           'Anzahl Barverkäufe
Dim erstKd%()           'erster Kunde         '1.09
Dim letztKd%()          'letzter Kunde        '1.09
Dim Pause1Anz%()
Dim Pause1Zeit&()
Dim Pause2Anz%()
Dim Pause2Zeit&()
Dim AnzRezKd!()
Dim AnzZusKd!()
Dim AnzSPrmKd!()
Dim PrivUms#()
Dim ZusatzUms#()

  '(Gerät/Seite/Mitarb)
Dim ganzKd!()           'Anzahl Kunden
Dim ganzRp!()           'Anzahl Rezepte
Dim gAnzPB!()           'Anzahl PrämienBasen
Dim gPB#()              'Summe PrämienBasis
Dim gAnzSP!()           'Anzahl SonderPrämien
Dim gSP#()              'SonderPrämie
Dim gAnzZV!()           'Anzahl ZusatzVerkäufe
Dim gZV#()              'ZusatzVerkauf
'(MitArb)
Dim AnzRab!()                  'Anzahl Rabatte
Dim SumRab#()                  'Summe Rabatte
Dim Umsatz#()                  'Brutto-Umsatz
        
Public aDetail$()
Public aInfo$()          'Infobereich
Public dInfo$()          'Werteauswahl für Diagramme
Public DiagPnr%()        'PersNr für Diagramme bei Vorschau

Dim Pers_Flex%()  'welche PersNr in welchem Flex
        
Dim AnzFeiertage%
        
Public DefErrFncStr(50) As String
Public DefErrStk As Integer

Public flx_Loaded As Integer

Global ActFileHandle As Integer
Global OpendFileName() As String
Global ActFileName As String
Global OpenForms As Integer

Global clsfabs As New clsFabsKlasse

Public ProgrammChar$

Public ProgrammNamen$(2)
Public ProgrammTyp%

Global buf$
Public LeereZeilen As Boolean

Public ProgrammModus%
Public TaxeOk%
Public UebergabeStr$

Public FabsErrf%
Public FabsRecno&

Public DruckSeite%

Public UserSection$

Public KeinRowColChange%
Public WirklichKeinRowColChange%


Public ast As clsStamm
Public ass As clsStatistik
Public v As clsVerkauf
Public vt As clsVerkRtab
Public nnek As clsNNEK

Public TaxeDB As Database
Public TaxeRec As Recordset
Public taxe As clsTaxe
Public arttext As clsArttext
Public besorgt As clsBesorgt
Public para As clsOpPara
Public wpara As clsWinPara

Public EditErg%
Public EditModus%
Public EditTxt$
Public EditAnzGefunden%
Public EditGef%(49)

Public ActProgram As Object

Public ActBenutzer%

Public INI_DATEI As String

Public PersonalFarben$(50)
Public PersonalInitialen$(50)
Public DiagrammTyp%(4)
Public DiagrammWas%(4)
Public pb#, zpb#, spb#, PrämBasis$
Public RabAbzug As Boolean, LSAuch As Boolean, PrivRez As Boolean
Public LegendenPosStr$
Public sF2Vorschau As Boolean

Const SPRM_DB = "PSTSPRM.mdb"
Public SprmDB As Database
Public sp As Recordset


Public Const APOCONTROL_DB = "CONTROL.MDB"
Public Const APOCONTROLW_DB = "CONTROLW.MDB"
Public ApoControlDB As Database
Public ApoControlWDB As Database
Public ApoControlRec As Recordset
Public ApoControlTRec As Recordset
Public ApoControlWRec As Recordset
Public ApoCDBda As Boolean

Public PrgAction As Byte
Public Const AUSWERTUNG As Byte = 0
Public Const ARTIKEL_EINLESEN As Byte = 1
Public Const VERKAUF_EINLESEN As Byte = 2

Public OffenRec(5) As OffenStruct
Public OffenRec2(5) As OffenStruct
Public ToleranzOffen%, ToleranzVergleich%

Function OpenCreateSprmArtikel%()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("OpenCreateSprmArtikel")
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

Dim SprmTd As TableDef
Dim SprmIdx As Index
Dim SprmFld As Field
Dim IxFld As Field

Dim kkey$
Dim ActRecNo&
Dim ret%

ret% = True

On Error Resume Next
Err.Clear
Set SprmDB = OpenDatabase(SPRM_DB)
Set sp = SprmDB.OpenRecordset("Sonderprämienartikel", dbOpenTable)
If Err.Number <> 0 Then
  On Error GoTo DefErr
  If Dir(SPRM_DB) <> "" Then Kill SPRM_DB
  Set SprmDB = CreateDatabase(SPRM_DB, dbLangGeneral)   ', dbVersion30)
  
  Set SprmTd = SprmDB.CreateTableDef("Sonderprämienartikel")
  
  Set SprmFld = SprmTd.CreateField("PZN", dbText)
  SprmFld.AllowZeroLength = False
  SprmFld.Size = 7
  SprmTd.Fields.Append SprmFld
  
  ' Indizes für Sonderprämienartikel
  Set SprmIdx = SprmTd.CreateIndex()
  SprmIdx.Name = "Unique"
  SprmIdx.Primary = True
  SprmIdx.Unique = True
  Set IxFld = SprmIdx.CreateField("PZN")
  SprmIdx.Fields.Append IxFld
  SprmTd.Indexes.Append SprmIdx
  
  SprmDB.TableDefs.Append SprmTd
  
  Set sp = SprmDB.OpenRecordset("Sonderprämienartikel", dbOpenTable)
  
End If
On Error GoTo DefErr

OpenCreateSprmArtikel% = ret%
Call DefErrPop
End Function

Function OpenCreateControlDB() As Boolean
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("OpenCreateControlDB")
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
Dim Td As TableDef
Dim Idx As Index
Dim fld As Field
Dim IxFld As Field
Dim i%, j%, k%, dummy%, DBneu%, max%
Dim OldClass As Object
Dim iRec As Recordset
Dim BezName$(100)
Dim BezTabelle%(100)
Dim BezFormatStr$(100)
Dim erg As Long

On Error Resume Next
Err.Clear
Set ApoControlDB = OpenDatabase(APOCONTROL_DB, False, False)
If (Err.Number <> 0) Then
    erg = iMsgBox("Zum Erstellen der ApoControl-Datenbank bitte ApoControl anwählen.", vbOKOnly Or vbInformation, "Verkaufsplatzperipherie")
    Call DefErrPop
    Exit Function
End If
Set ApoControlRec = ApoControlDB.OpenRecordset("Bezeichnungen", dbOpenTable)
Set ApoControlTRec = ApoControlDB.OpenRecordset("Tabellen", dbOpenTable)
On Error Resume Next
Err.Clear
Set ApoControlWDB = OpenDatabase(APOCONTROLW_DB, False, False)
If (Err.Number <> 0) Then
    erg = iMsgBox("Zum Erstellen der ApoControl-Datenbank bitte ApoControl anwählen.", vbOKOnly Or vbInformation, "Verkaufsplatzperipherie")
    Call DefErrPop
    Exit Function
End If
Set ApoControlWRec = ApoControlWDB.OpenRecordset("Werte", dbOpenTable)
On Error GoTo DefErr
OpenCreateControlDB = True
Call DefErrPop
End Function




Sub InitPersStat()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("InitPersStat")
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
Dim buf As String
Dim Tag As Date
Dim i%

v.GetRecord 2
Tag = dDatum(v.datum)
Tag = DateAdd("d", -1, Tag)

'Rec1: Header
buf = MKS(63) + MKS(64) + Space(56)
Put fpss, 1, buf

'Rec2: Variable
p2.lAusDat = MKDatum(Tag)
p2.lAusTim = 0
p2.lAusVon = MKDatum(Tag)
p2.lAusBis = MKDatum(Tag)
p2.UebOK = 0
p2.UebErst = -1
p2.lUebDat = MKDatum(Tag)
p2.lUebTim = 0
p2.lAnz = MKS(0)
p2.gAnz = MKS(0)
p2.rest = String(Len(p2.rest), 0)
Put fpss, 64 + 1, p2

'Rec3: Parameter
p3.tENDE = 1800
p3.GndPrm = 225
p3.PrmBas = "I"
p3.SndPrm = 0
p3.ZVkPrm = 0
p3.RabAbzg = "N"
p3.Liefauch = "N"
p3.Normtag = 480
p3.PauseKl = 30
p3.PauseGr = 60
p3.geraete = 0
p3.summen = 0
p3.params = 0
p3.Detail = -1
p3.zeiten = 0
p3.PrivRez = "N"
p3.rest = String(Len(p3.rest), 0)
Put fpss, 2 * 64 + 1, p3

'Rec4-13: Parameter-Untergruppen
p4.ind = " "
p4.wGrp = "1-9"   ' + SPACE$(23)
p4.tLager = Space(Len(p4.tLager))
p4.LgCode = Space(Len(p4.LgCode))
p4.vonAVP = 0
p4.bisAVP = 0
p4.vonSP = 0
p4.bisSP = 0
p4.RpPfl = "J"
p4.Geräte = Space(Len(p4.Geräte))
p4.rest = Space(Len(p4.rest))
Put fpss, 3 * 64 + 1, p4

p4.wGrp = Space(Len(p4.wGrp))
p4.RpPfl = " "
For i% = 4 To 12
  Put fpss, i% * 64 + 1, p4
Next i%

'Rec14-63: Mitarbeiter (reserviert)
buf = Space(64)
For i% = 13 To 62
  Put fpss, i% * 64 + 1, buf
Next i%

Call DefErrPop
End Sub


Sub ParseWarenGruppe(sugz%, iwg%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("ParseWarenGruppe")
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
  
Dim x1$, tok$
Dim i%, i1%, i2%, j%, zw%
  
x1$ = wg$(sugz%): wg1$(sugz%) = ""
x1$ = Trim(x1$)
iwg% = 0
i1% = 0: i2% = 0
Do While Len(x1$) > 0
  j% = 0
  i% = InStr(x1$, ",")
  If i% Then
    tok$ = Left$(x1$, i% - 1)
  Else
    tok$ = x1$
    x1$ = ""
  End If
  j% = InStr(tok$, "-")
  If j% Then
    If Val(tok$) > 0 And Val(tok$) <= 99 Then
      i1% = Val(tok$)
      If Val(Mid$(tok$, j% + 1)) > 0 And Val(Mid$(tok$, j% + 1)) <= 99 Then
        i2% = Val(Mid$(tok$, j% + 1))
        If i1% > i2% Then
          zw% = i1%
          i1% = i2%
          i2% = zw%
        End If
        If i1% < 10 Then i1% = i1% * 10
        If i2% < 10 Then i2% = i2% * 10 + 9
      Else
        Exit Do
      End If
    Else
      Exit Do
    End If
  Else
    If Val(tok$) > 0 And Val(tok$) <= 99 Or tok$ = String$(Len(tok$), "0") Then
      i1% = Val(tok$)
      If i1% > 0 And i1% < 10 Then
        i1% = i1% * 10
        i2% = i1% + 9
      Else
        i2% = i1%
      End If
    Else
      Exit Do
    End If
  End If
  If iwg% >= 0 Then
    If i1% > i2% Then
      zw% = i1%
      i1% = i2%
      i2% = zw%
    End If
    For j% = i1% To i2%
      iwg% = iwg% + 1
      If iwg% <= 100 Then
        wg1$(sugz%) = wg1$(sugz%) + "," + Right$(Str$(1000 + j%), 2)
      Else
        Exit Do
      End If
    Next j%
    x1$ = Mid$(x1$, i% + 1)
  End If
Loop
Call DefErrPop
End Sub


Sub AnfangsBedingungen()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("AnfangsBedingungen")
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

Dim i%, iwg%

Get #fpss, 65, p2
vonAuswD = CVDatum(p2.lAusVon)
bisAuswD = CVDatum(p2.lAusBis)
ErstesMal% = p2.UebErst
UebDatBis = CVDatum(p2.lUebDat)

Get #fpss, 2 * 64 + 1, p3
geraete = False: If p3.geraete <> 0 Then geraete = True
summen = False: If p3.summen <> 0 Then summen = True
params = False: If p3.params <> 0 Then params = True
Detail = False: If p3.Detail <> 0 Then Detail = True
pb# = p3.GndPrm / 100#
zpb# = p3.ZVkPrm / 100#
spb# = p3.SndPrm / 100#
PrämBasis$ = p3.PrmBas
If p3.RabAbzg = "N" Then
  RabAbzug = False
Else
  RabAbzug = True
End If
If p3.Liefauch = "N" Then
  LSAuch = False
Else
  LSAuch = True
End If
If p3.PrivRez = "N" Then
  PrivRez = False
Else
  PrivRez = True
End If

NORMTAGZEIT% = p3.Normtag
If NORMTAGZEIT% = 0 Then NORMTAGZEIT% = 480   '1.23
PauseKl% = p3.PauseKl
PauseGr% = p3.PauseGr: If PauseGr% = 0 Then PauseGr% = 60
'TagAnf% = p3.TagAnf: If TagAnf% = 0 Then TagAnf% = 745
'TagEnd% = p3.TagEnd: If TagEnd% = 0 Then TagEnd% = 1815

chkSPANNE = False
chkTARA = False
chkAVP = False
chkREZPF = False
chkLGCOD = False

For i% = 1 To 10
  Get #fpss, (2 + i%) * 64 + 1, p4
  operator$(i%) = p4.ind
  wg$(i%) = p4.wGrp
  Call ParseWarenGruppe(i%, iwg%)
  tLager$(i%) = p4.tLager
  vonSP#(i%) = CDbl(p4.vonSP)
  If vonSP#(i%) > 0 Then
    'vonSP$(i%) = Right(Space(5) + Str(vonSP#(i%)), 5)
    chkSPANNE = True
'  Else
'    vonSP$(i%) = Space(5)
  End If
  bisSP#(i%) = CDbl(p4.bisSP)
  If bisSP#(i%) > 0 Then
    'bisSP$(i%) = Right$(Space(5) + Str$(bisSP#(i%)), 5)
    chkSPANNE = True
 ' Else
 '   bisSP$(i%) = Space(5)
  End If
  If bisSP#(i%) = 0 Then bisSP#(i%) = 9999999#
  Geräte$(i%) = p4.Geräte
  Rp$(i%) = p4.RpPfl
  LgCode$(i%) = p4.LgCode
  LgCode$(i%) = Left(LgCode$(i%) + Space(10), 10)
  vonAVP#(i%) = CDbl(p4.vonAVP)
  If vonAVP#(i%) > 0 Then
'    vonAVP$(i%) = Right(Space(5) + Str(vonAVP#(i%)), 5)
    chkAVP = True
 ' Else
 '   vonAVP$(i%) = Space(5)
  End If
  bisAVP#(i%) = CDbl(p4.bisAVP)
  If bisAVP#(i%) > 0 Then
'    bisAVP$(i%) = Right(Space(5) + Str(bisAVP#(i%)), 5)
    chkAVP = True
'  Else
'    bisAVP$(i%) = Space(5)
  End If
  If bisAVP#(i%) = 0 Then bisAVP#(i%) = 9999999#
  tLager$(i%) = Trim(tLager$(i%))
  If Len(tLager$(i%)) > 0 Then chkTARA = True
  LgCode$(i%) = Trim(LgCode$(i%))
  If LgCode$(i%) <> Space$(Len(LgCode$(i%))) Then chkLGCOD = True
  Rp$(i%) = Trim(Rp$(i%))
  If Rp$(i%) = "J" Then chkREZPF = True
Next i%

Call DefErrPop
End Sub

Sub DatumEin()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("DatumEin")
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
Dim erg As Integer, satz As Long
Dim fhTmp As Long, buf As String

PrgAction = AUSWERTUNG
abbruch = True
Vorschau = False

Call AnfangsBedingungen
Call AuswahlAnzeigen

frmDatum.Show vbModal

If Not abbruch Then
  With frmAction.cboMitarbeiter
    .Clear
    .AddItem "Gesamt"
    For i% = 1 To anzahl%
        h$ = RTrim$(para.Personal(i%))
        If (h$ <> "") And iMitArb(i%) Then
            .AddItem h$
        End If
    Next i%
  End With

  Call StartAnimation(frmAction, "Ausgabe wird erstellt ...")
  Call MitArbSpeichern(anzahl%)
  If vonAuswD > UebDatBis Then
    Vorschau = True
    Close #fpss
    On Error Resume Next
    Kill "PSTPREV.DAT"
    On Error GoTo DefErr
    fpss = fopen("PSTPREV.DAT", "RW")
    Call InitPersStat
    fhTmp = fopen("PSTATIS.DAT", "r")
    buf = String(64, 0)
    For satz = 2 To 12
      Get #fhTmp, satz * 64 + 1, buf
      Put #fpss, satz * 64 + 1, buf
    Next satz
    Close #fhTmp
    Call Überleitung
  End If
  If LOF(fpss) > 0 Then Call AuswertungDurchfuehren
  Call StopAnimation(frmAction)
  If Not Vorschau Then
    Get #fpss, 65, p2
    LSet p2.lAusVon = MKDatum(vonAuswD)
    LSet p2.lAusBis = MKDatum(bisAuswD)
    Put #fpss, 65, p2
  Else
    Close #fpss
    On Error Resume Next
    Kill "PSTPREV.DAT"
    On Error GoTo DefErr
    fpss = fopen("PSTATIS.DAT", "RW")
    If LOF(fpss) = 0 Then
      Call InitPersStat
    End If
  End If
  Call AnfangsBedingungen
  Call frmAction.cboMitarbeiter_Click
End If
Call DefErrPop
End Sub

Sub FeiertageAnpassen(Optional Jahr$ = "")
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("FeiertageAnpassen")
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

Dim i%, j%, ind%
Dim l&
Dim wert1$, h$, key$, Aktiv$, mond$
Dim OsterSonntag As Date, advent1 As Date, BetTag As Date

If Jahr$ = "" Then Jahr$ = Format(Now, "YYYY")

Select Case Val(Jahr$) Mod 19
    Case 0
        mond$ = "15.04."
    Case 1
        mond$ = "03.04."
    Case 2
        mond$ = "23.03."
    Case 3
        mond$ = "11.04."
    Case 4
        mond$ = "31.03."
    Case 5
        mond$ = "18.04."
    Case 6
        mond$ = "08.04."
    Case 7
        mond$ = "28.03."
    Case 8
        mond$ = "16.04."
    Case 9
        mond$ = "05.04."
    Case 10
        mond$ = "25.03."
    Case 11
        mond$ = "13.04."
    Case 12
        mond$ = "02.04."
    Case 13
        mond$ = "22.03."
    Case 14
        mond$ = "10.04."
    Case 15
        mond$ = "30.03."
    Case 16
        mond$ = "17.04."
    Case 17
        mond$ = "07.04."
    Case 18
        mond$ = "27.03."
End Select
mond$ = mond$ + Jahr$
  
OsterSonntag = CDate(mond$) + 7 - (WeekDay(mond$) - 1)

advent1 = "04.12." + Jahr$
While (WeekDay(advent1) <> vbSunday)
    advent1 = advent1 - 1
Wend
BetTag = advent1 - 7 - 4

Feiertage(0).KalenderTag = CDate("01.01." + Jahr$)
Feiertage(1).KalenderTag = CDate("06.01." + Jahr$)
Feiertage(2).KalenderTag = CDate(Format(OsterSonntag - 2, "DD.MM.YYYY"))
Feiertage(3).KalenderTag = CDate(Format(OsterSonntag + 1, "DD.MM.YYYY"))
Feiertage(4).KalenderTag = CDate("01.05." + Jahr$)
Feiertage(5).KalenderTag = CDate(Format(OsterSonntag + 39, "DD.MM.YYYY"))
Feiertage(6).KalenderTag = CDate(Format(OsterSonntag + 50, "DD.MM.YYYY"))
Feiertage(7).KalenderTag = CDate(Format(OsterSonntag + 60, "DD.MM.YYYY"))
Feiertage(8).KalenderTag = CDate("15.08." + Jahr$)

If (para.Land = "A") Then
    Feiertage(9).KalenderTag = CDate("26.10." + Jahr$)
Else
    Feiertage(9).KalenderTag = CDate("03.10." + Jahr$)
End If

Feiertage(10).KalenderTag = CDate("31.10." + Jahr$)
Feiertage(11).KalenderTag = CDate("01.11." + Jahr$)

If (para.Land = "A") Then
    Feiertage(12).KalenderTag = CDate("08.12." + Jahr$)
Else
    Feiertage(12).KalenderTag = CDate(Format(BetTag, "DD.MM.YYYY"))
End If

Feiertage(13).KalenderTag = CDate("25.12." + Jahr$)
Feiertage(14).KalenderTag = CDate("26.12." + Jahr$)

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
Dim l&, hwnd&
Dim h$, SQLStr$, DirRet$, FuDate$, CdInfoVer$, s$
Dim WinText As String * 255
Dim a!
Dim chef%, Menu As Boolean
Dim buf As String

Call InitMisc

If (App.PrevInstance) Then
    hwnd& = GetForegroundWindow()
    hwnd& = GetWindow(hwnd&, GW_HWNDFIRST)
    Do Until (hwnd& = 0)
        l& = GetWindowText(hwnd&, WinText, 255)
        h$ = Left$(WinText, l&)
        For i% = 0 To 2
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
        hwnd& = GetWindow(hwnd&, GW_HWNDNEXT)
    Loop
    End
End If
Set para = New clsOpPara
Set wpara = New clsWinPara
Call para.HoleFirmenStamm
If (InStr(Oem2Ansi(para.Benutz), "ä") <= 0) Then
  Call iMsgBox("Dieses Programm hat Ihre Apotheke nicht gekauft !", vbCritical)
  wpara.ExitEndSub
'    Call frmAction.frmActionUnload
  End
  Call DefErrPop: Exit Sub
End If
ActBenutzer% = HoleActBenutzer%(Menu)

'If Not Menu Then
'  erg% = iMsgBox("Das Programm muss aus dem Optipharm Menü gestartet werden !")
'  End
'End If
If ActBenutzer% <> 1 Then
  frmPass.Show vbModal
  If ActBenutzer <> 1 Then
    erg% = iMsgBox("Sie haben keine Benutzungsberechtigung für dieses Programm !")
    wpara.ExitEndSub
    End
  End If
End If

If (Dir$("fistam.dat") = "") Then ChDir "\user"
INI_DATEI = CurDir + "\winop.ini"

Set ast = New clsStamm
Set ass = New clsStatistik
Set taxe = New clsTaxe
Set arttext = New clsArttext
Set besorgt = New clsBesorgt
Set v = New clsVerkauf
Set vt = New clsVerkRtab
Set nnek = New clsNNEK

UserSection$ = "Computer" + Format(Val(para.User))
Call wpara.HoleWindowsParameter
Call HoleIniFeiertage
Call HoleIniPersonalFarben
Call HoleIniDiagramme
Call HoleIniOffen


frmAction.Show

Call StartAnimation(frmAction, "Parameter werden eingelesen ...")

Call para.AuslesenPdatei
Call para.EinlesenPersonal
Call MitArbLaden(anzahl%)

ast.OpenDatei
ass.OpenDatei
v.OpenDatei
vt.OpenDatei

nnek.OpenDatei

TaxeOk% = False
        
erg% = 0
h$ = para.TaxeLw + ":\taxe\taxe.mdb"
Set TaxeDB = taxe.OpenDatenbank(h$, False, True)
'
TaxeOk% = True
Set TaxeRec = TaxeDB.OpenRecordset("Taxe", dbOpenTable)
    

arttext.OpenDatei
besorgt.OpenDatei
erg% = OpenCreateSprmArtikel


Call StopAnimation(frmAction)
Call frmAction.WechselModus(0)

fpss = fopen("PSTATIS.DAT", "RW")
If LOF(fpss) = 0 Then Call InitPersStat
Call AnfangsBedingungen

If Not abbruch Then
  If ErstesMal% <> 0 Then
    Call Überleitung
    If ErstesMal% = 0 Then Call DatumEin
  Else
    Call DatumEin
  End If
End If

Call DefErrPop
End Sub

Public Sub Überleitung()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("Überleitung")
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
PrgAction = VERKAUF_EINLESEN
abbruch = False
If Not Vorschau Then
  frmDatum.Show vbModal
End If
If Not abbruch Then
  frmFortschritt.Show vbModal
End If
Call DefErrPop
End Sub


Sub AuswahlAnzeigen()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("AuswahlAnzeigen")
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

With frmDatum!lstWer
  .Clear
  For i% = 1 To anzahl%
    If Left(MitArb$(i%), 5) <> "     " Then
      .AddItem Left$(MitArb$(i%), 20)
      If iMitArb(Val(Mid(MitArb$(i%), 22, 2))) Then
      'If Mid$(MitArb$(i%), 21, 1) = " " Then
        .Selected(.ListCount - 1) = True
      Else
        .Selected(.ListCount - 1) = False
      End If
    End If
  Next i%
  .ListIndex = 0
End With
Call DefErrPop
End Sub


Function iMsgBox%(prompt$, Optional buttons% = 0, Optional title$ = "")
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("iMsgBox")
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

Dim OrgKeinRowColChange%, ret%

OrgKeinRowColChange% = KeinRowColChange%
KeinRowColChange% = True
If (title$ = "") Then title$ = "Personal-Verkaufstatistik"
If (title$ <> "") Then
    ret% = MsgBox(prompt$, buttons%, title$)
Else
    ret% = MsgBox(prompt$, buttons%)
End If
KeinRowColChange% = OrgKeinRowColChange%

iMsgBox% = ret%
Call DefErrPop
End Function


Sub AuswertungDurchfuehren()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("AuswertungDurchfuehren")
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

Dim d As Date, von As Date, Bis As Date
Dim loopFlag As Boolean, gefunden As Boolean

von = vonAuswD
Bis = bisAuswD
If Vorschau Then
  aTage% = anzahl%
Else
  aTage% = DateValue(Bis) - DateValue(von) + 2
End If
Call TabelleLoeschen
Call FeiertageAnpassen(Format(von, "YYYY"))
Call AuswertungRechnen(gefunden, von, Bis)
Call AuswertungDrucken(von, Bis)
frmAction.Show

Call DefErrPop
End Sub


Sub AuswertungDrucken(von As Date, Bis As Date)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("AuswertungDrucken")
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

'ReDim kdProz!(10, 2)
Dim m%, g%, a%, d%, z%, i%, j%, k%, spalte%
Dim BasisSumme#, BasisPraemie#, SonderSumme#, SonderPraemie#, ZusatzSumme#, ZusatzPraemie#, PraemienSumme#
Dim Kd10tel#, Rp10tel#, AVP10tel#
Dim ArbeitsZeit&, AnzahlNormTage!, zw!
Dim NT!
Dim AnfTag%, EndTag%, zeile%
Dim p$, s$, h$
ReDim aDetail$(15, anzahl%, aTage%)
ReDim aInfo$(6, 5, anzahl%)
ReDim dInfo$(29, 2, anzahl%)
ReDim AnzNT(anzahl%, 1) As Single
ReDim DiagPnr%(0)

Get #fpss, (2 * 64) + 1, p3
pb# = p3.GndPrm / 100#
zpb# = p3.ZVkPrm / 100#
GoSub AuswertungDruckenKopfXT

If Vorschau Then
  Detail = True: geraete = False
  AnfTag% = 1
  EndTag% = AnfTag%
Else
  AnfTag% = 1
  EndTag% = AnfTag% + aTage% - 2
End If

'rückwärts laufen, damit GesamtSummen (Index 0) richtig berechnet werden
For m% = anzahl% To 0 Step -1
  ArbeitsZeit& = 0
  'ReDim kdProz!(10, 2)
  If m% = 0 Then
    a% = 0
  Else
    a% = Val(Right(MitArb$(m%), 2))
  End If
  If (m% = 0 Or Mid$(MitArb$(m%), 21, 1) <> " " Or Vorschau) And (Basis#(0, 0, a%) > 0 Or (AnzKd!(0, 1, a%)) > 0) Then
    If Detail Then
'">Tag|>PrämBasis|>ZusatzBasis|>SonderBasis|>AnzKd|>AnzRez|>%PrämKd|>%PrivKd|>%RezKd|>Erster|>Letzter|>kl.|>Pausen|>gr.|>Pausen"
      If Vorschau Then
        aDetail$(0, m%, 0) = "Pnr"
      Else
        aDetail$(0, m%, 0) = "Tag"
      End If
      aDetail$(1, m%, 0) = "PrämBasis"
      aDetail$(2, m%, 0) = "Zus.Basis"
      aDetail$(3, m%, 0) = "Snd.Basis"
      aDetail$(4, m%, 0) = "AnzKd"
      aDetail$(5, m%, 0) = "AnzRez"
      aDetail$(6, m%, 0) = "PräKd%"
      aDetail$(7, m%, 0) = "PrivKd%"
      aDetail$(8, m%, 0) = "RezKd%"
      aDetail$(9, m%, 0) = "Erster"
      aDetail$(10, m%, 0) = "Letzter"
      aDetail$(11, m%, 0) = "kleine"
      aDetail$(12, m%, 0) = "Pausen"
      aDetail$(13, m%, 0) = "große"
      aDetail$(14, m%, 0) = "Pausen"
      If Vorschau Then aDetail$(15, m%, 0) = "Kd/NT"
      
      
      zeile% = 0
      For d% = AnfTag% To EndTag%
        zeile% = zeile% + 1
        
        z% = 1
        If m% > 0 Then
          Pause1Anz%(0, z%, a%) = Pause1Anz%(0, z%, a%) + Pause1Anz%(d%, z%, a%)
          Pause1Zeit(0, z%, a%) = Pause1Zeit(0, z%, a%) + Pause1Zeit(d%, z%, a%)
          Pause2Anz%(0, z%, a%) = Pause2Anz%(0, z%, a%) + Pause2Anz%(d%, z%, a%)
          Pause2Zeit(0, z%, a%) = Pause2Zeit(0, z%, a%) + Pause2Zeit(d%, z%, a%)
        End If
        If Vorschau Then
          aDetail$(0, m%, zeile%) = Format(a%, "##")
        Else
          aDetail$(0, m%, zeile%) = Format(DateAdd("d", d% - 1, vonAuswD), "dd")
        End If
        'links
        If Basis#(d%, 1, a%) <> 0 Then
          aDetail$(1, m%, zeile%) = Format(Basis#(d%, 1, a%), " ###,###")
        End If
        If Zusatz#(d%, 1, a%) <> 0 Then
          aDetail$(2, m%, zeile%) = Format(Zusatz#(d%, 1, a%), " ###,###")
        End If
        If Sonder#(d%, 1, a%) <> 0 Then
          aDetail$(3, m%, zeile%) = Format(Sonder#(d%, 1, a%), " ###,###")
        End If
        
        If AnzKd!(d%, 1, a%) <> 0 Then
          aDetail$(4, m%, zeile%) = Format(AnzKd!(d%, 1, a%), " ##,###")
        End If
        
        If AnzRp!(d%, 1, a%) <> 0 Then
          aDetail$(5, m%, zeile%) = Format(AnzRp!(d%, 1, a%), " ##,###")
        End If
  
        If AnzKd!(d%, 1, a%) <> 0 Then
          aDetail$(6, m%, zeile%) = Format(AnzPKd!(d%, 1, a%) * 100! / AnzKd!(d%, 1, a%), " ##0.0")
        End If
        
        zw! = AnzKd!(d%, 1, a%) - AnzRezKd!(d%, 1, a%)
        If zw! <> 0 And AnzKd!(d%, 1, a%) <> 0 Then
          aDetail$(7, m%, zeile%) = Format(zw! * 100 / AnzKd!(d%, 1, a%), " ##0.0")
        End If
        If AnzRezKd!(d%, 1, a%) <> 0 And AnzKd!(d%, 1, a%) <> 0 Then
          aDetail$(8, m%, zeile%) = Format(AnzRezKd!(d%, 1, a%) * 100! / AnzKd!(d%, 1, a%), " ##0.0")
        End If
        
        If erstKd%(d%, 1, a%) > 0 Then
          s$ = Right$("    " + Str$(erstKd%(d%, 1, a%)), 4)
          s$ = Left$(s$, 2) + ":" + Right$(s$, 2)
          aDetail$(9, m%, zeile%) = s$
          If erstKd%(d%, 1, 0) = 0 Or erstKd%(d%, z%, 0) > erstKd%(d%, 1, a%) Then
            erstKd%(d%, 1, 0) = erstKd%(d%, 1, a%)
          End If
        End If
        If letztKd%(d%, 1, a%) > 0 Then
          s$ = Right$("    " + Str$(letztKd%(d%, 1, a%)), 4)
          s$ = Left$(s$, 2) + ":" + Right$(s$, 2)
          aDetail$(10, m%, zeile%) = s$
          ArbeitsZeit& = ArbeitsZeit& + fMinuten%(letztKd%(d%, 1, a%)) - fMinuten%(erstKd%(d%, 1, a%))
          If letztKd%(d%, 1, 0) = 0 Or letztKd%(d%, z%, 0) < letztKd%(d%, 1, a%) Then
            letztKd%(d%, 1, 0) = letztKd%(d%, 1, a%)
          End If
        End If

        If Pause1Anz%(d%, 1, a%) > 0 Then
          aDetail$(11, m%, zeile%) = Format(Pause1Anz%(d%, 1, a%), "## ")
          aDetail$(12, m%, zeile%) = Format(Pause1Zeit(d%, 1, a%), "####")
        End If
        
        If Pause2Anz%(d%, 1, a%) > 0 Then
          aDetail$(13, m%, zeile%) = Format(Pause2Anz%(d%, 1, a%), "##")
          aDetail$(14, m%, zeile%) = Format(Pause2Zeit(d%, 1, a%), "####")
        End If
        If Vorschau Then
          If NORMTAGZEIT% <> 0 Then
            If ((ArbeitsZeit& - Pause2Zeit(0, 1, a%)) / NORMTAGZEIT%) > 0 Then
              aDetail$(15, m%, zeile%) = Format(AnzKd!(d%, 1, a%) / ((ArbeitsZeit& - Pause2Zeit(0, 1, a%)) / NORMTAGZEIT%), " #####")
            End If
          End If
        End If
      Next d%
      
      If m% > 0 Then
        If NORMTAGZEIT% <> 0 Then
          AnzahlNormTage! = (ArbeitsZeit& - Pause2Zeit(0, 1, a%)) / NORMTAGZEIT%
        End If
        AnzNT(m%, 0) = AnzahlNormTage!
        AnzNT(m%, 1) = m%
        If AnzahlNormTage! > 0 Then
          gAnzNormTage! = gAnzNormTage! + AnzahlNormTage!
        End If
      End If
      zeile% = aTage%

      If Not Vorschau Then
          'Summen: links
          If Basis#(0, 1, a%) <> 0 Then
            aDetail$(1, m%, zeile%) = Format(Basis#(0, 1, a%), " ###,###")
          End If
          If Zusatz#(0, 1, a%) <> 0 Then
            aDetail$(2, m%, zeile%) = Format(Zusatz#(0, 1, a%), " ###,###")
          End If
          If Sonder#(0, 1, a%) <> 0 Then
            aDetail$(3, m%, zeile%) = Format(Sonder#(0, 1, a%), " ###,###")
          End If
          
          If AnzKd!(0, 1, a%) <> 0 Then
            aDetail$(4, m%, zeile%) = Format(AnzKd!(0, 1, a%), " ##,###")
          End If
          If AnzRp!(0, 1, a%) <> 0 Then
            aDetail$(5, m%, zeile%) = Format(AnzRp!(0, 1, a%), " ##,###")
          End If
          If AnzKd!(0, 1, a%) <> 0 Then
            aDetail$(6, m%, zeile%) = Format(AnzPKd!(0, 1, a%) * 100! / AnzKd!(0, 1, a%), " ##0.0")
          End If
          zw! = AnzKd!(0, 1, a%) - AnzRezKd!(0, 1, a%)
          If zw! <> 0 And AnzKd!(0, 1, a%) <> 0 Then
            aDetail$(7, m%, zeile%) = Format(zw! * 100 / AnzKd!(0, 1, a%), " ##0.0")
          End If
          If AnzRezKd!(0, 1, a%) <> 0 And AnzKd!(0, 1, a%) <> 0 Then
            aDetail$(8, m%, zeile%) = Format(AnzRezKd!(0, 1, a%) * 100 / AnzKd!(0, 1, a%), " ##0.0")
          End If
          
'          s$ = ""
'          If ArbeitsZeit& <> 0& Then
'            s$ = Format(Int(ArbeitsZeit& / 60), "###") + ":"
'            s$ = s$ + Format((ArbeitsZeit& Mod 60), "##")
'          End If
'          aDetail$(10, m%, zeile%) = s$
                
          If Pause1Anz%(0, 1, a%) > 0 Then
            aDetail$(11, m%, zeile%) = Format(Pause1Anz%(0, 1, a%), "## ")
            aDetail$(12, m%, zeile%) = Format(Pause1Zeit(0, 1, a%), "####")
          End If
                
          If Pause2Anz%(0, 1, a%) > 0 Then
            aDetail$(13, m%, zeile%) = Format(Pause2Anz%(0, 1, a%), "##")
            aDetail$(14, m%, zeile%) = Format(Pause2Zeit(0, 1, a%), "####")
          End If

'              'Monatssummen
          NT! = AnzahlNormTage!
          If m% = 0 Then
            NT! = gAnzNormTage!
          End If
          GoSub SummenZeilen
      Else
          NT! = AnzahlNormTage!
          If m% = 0 Then
            NT! = gAnzNormTage!
          End If
      End If      'If not Vorschau
    End If        'If Detail
    '-------------------------------------------------------------------
    If geraete And Not Vorschau Then
'            If Detail% Then Seite$ = "Seite 2 von 2" Else Seite$ = "Seite 1 von 1"
'            GoSub AuswertungDruckenKopf
'            'Geräte
'            Print #pr, "Gerät Kunden  -% Rezepte     Grundprämien";
'            Print #pr, "     Sonderprämien     Zusatzprämien"
'            Print #pr, String$(77, "-")
'
'            For g% = 0 To 9
'                    If ganzKd!(g%, 1, a%) <> 0 Then
'                            If (ganzKd!(10, 0, a%) > 0) Then
'                                    kdProz!(g%, 1) = ganzKd!(g%, 1, a%) / (ganzKd!(10, 0, a%) / 100)
'                            Else
'                                    kdProz!(g%, 1) = 0
'                            End If
'                            kdProz!(10, 1) = kdProz!(10, 1) + kdProz!(g%, 1)
'                            ikd1% = g%
'                    End If
'                    If ganzKd!(g%, 2, a%) <> 0 Then
'                            If ganzKd!(10, 0, a%) > 0 Then
'                                    kdProz!(g%, 2) = ganzKd!(g%, 2, a%) / (ganzKd!(10, 0, a%) / 100) + 0.05
'                            Else
'                                    kdProz!(g%, 2) = 0
'                            End If
'                            kdProz!(10, 2) = kdProz!(10, 2) + kdProz!(g%, 2)
'                            ikd2% = g%
'                    End If
'            Next g%
'
'            If kdProz!(10, 1) <> 100! Then
'            '  kdProz!(ikd1%, 1) = kdProz!(ikd1%, 1) + 100! - kdProz!(10, 1)
'            End If
'            If kdProz!(10, 2) <> 100! Then
'            '  kdProz!(ikd2%, 2) = kdProz!(ikd2%, 2) + 100! - kdProz!(10, 2)
'            End If
'
'            For g% = 0 To 9
'                    'links
'                    Print #pr, USING; "#/L"; g%;
'                    If ganzKd!(g%, 1, a%) <> 0 Then
'                            Print #pr, USING; "  ##### ###.#"; ganzKd!(g%, 1, a%); kdProz!(g%, 1);
'                    Else
'                            Print #pr, Space$(13);
'                    End If
'                    If ganzRp!(g%, 1, a%) <> 0 Then
'                            Print #pr, USING; "  #####"; ganzRp!(g%, 1, a%);
'                    Else
'                            Print #pr, Space$(7);
'                    End If
'                    If gAnzPB!(g%, 1, a%) <> 0 Then
'                            Print #pr, USING; "  ##### ###,###.##"; gAnzPB!(g%, 1, a%); gPB#(g%, 1, a%)
'                    Else
'                            Print #pr, Space$(18)
'                    End If
'                    '  ##### ###,###.##  ##### ###,###.##
'
'                    'rechts
'                    Print #pr, " /R";
'                    If ganzKd!(g%, 2, a%) <> 0 Then
'                            Print #pr, USING; "  ##### ###.#"; ganzKd!(g%, 2, a%); kdProz!(g%, 2);
'                    Else
'                            Print #pr, Space$(13);
'                    End If
'                    If ganzRp!(g%, 2, a%) <> 0 Then
'                            Print #pr, USING; "  #####"; ganzRp!(g%, 2, a%);
'                    Else
'                            Print #pr, Space$(7);
'                    End If
'                    If gAnzPB!(g%, 2, a%) <> 0 Then
'                            Print #pr, USING; "  ##### ###,###.##"; gAnzPB!(g%, 2, a%); gPB#(g%, 2, a%)
'                    Else
'                            Print #pr, Space$(18)
'                    End If
'            Next g%
'            Print #pr, String$(77, "-")
'
'            'Summe
'            If ganzKd!(10, 0, a%) <> 0 Then
'                    Print #pr, USING; "     ##### ###.#"; ganzKd!(10, 0, a%); 100!;
'            Else
'                    Print #pr, Space$(13);
'            End If
'            If ganzRp!(10, 0, a%) <> 0 Then
'                    Print #pr, USING; "  #####"; ganzRp!(10, 0, a%);
'            Else
'                    Print #pr, Space$(7);
'            End If
'            If gAnzPB!(10, 0, a%) <> 0 Then
'                    Print #pr, USING; "  ##### ###,###.##"; gAnzPB!(10, 0, a%); gPB#(10, 0, a%)
'            Else
'                    Print #pr, Space$(18)
'            End If
'            Print #pr, "   "; String$(74, "-")
'
'            Print #pr, ff$   '1.11
    End If
  End If
  
  dInfo$(0, 0, m%) = "Kunden mit Prämie"
  If Vorschau Then
    dInfo$(0, 1, m%) = Format(AnzPKd!(0, 1, a%), "##,###")
  Else
    dInfo$(0, 1, m%) = aInfo$(0, 1, m%)
  End If
  dInfo$(0, 2, m%) = " "
  dInfo$(1, 0, m%) = "Kunden mit Prämie %"
  If Vorschau Then
    dInfo$(1, 1, m%) = aDetail$(6, m%, 1)
  Else
    dInfo$(1, 1, m%) = aDetail$(6, m%, aTage%)
  End If
  dInfo$(1, 2, m%) = "#"
  
  
  dInfo$(2, 0, m%) = "Kunden mit Zusatzverkauf"
  If Vorschau Then
    dInfo$(2, 1, m%) = Format(AnzZusKd!(0, 1, a%), "##,###")
  Else
    dInfo$(2, 1, m%) = aInfo$(1, 1, m%)
  End If
  dInfo$(2, 2, m%) = " "
  dInfo$(3, 0, m%) = "Kunden mit Sonderprämie"
  If Vorschau Then
    dInfo$(3, 1, m%) = Format(AnzSPrmKd!(0, 1, a%), "##,###")
  Else
    dInfo$(3, 1, m%) = aInfo$(2, 1, m%)
  End If
  dInfo$(3, 2, m%) = " "
  dInfo$(4, 0, m%) = "Privatkunden"
  If Vorschau Then
    dInfo$(4, 1, m%) = Format(AnzKd!(0, 1, a%) - AnzRezKd!(0, 1, a%), "##,###")
  Else
    dInfo$(4, 1, m%) = aInfo$(3, 1, m%)
  End If
  dInfo$(4, 2, m%) = " "
  dInfo$(5, 0, m%) = "Privatkunden %"
  If Vorschau Then
    dInfo$(5, 1, m%) = aDetail$(7, m%, 1)
  Else
    dInfo$(5, 1, m%) = aDetail$(7, m%, aTage%)
  End If
  dInfo$(5, 2, m%) = "#"
    
  dInfo$(6, 0, m%) = "Rezeptkunden"
  If Vorschau Then
    dInfo$(6, 1, m%) = Format(AnzRezKd!(0, 1, a%), "##,###")
  Else
    dInfo$(6, 1, m%) = aInfo$(4, 1, m%)
  End If
  dInfo$(6, 2, m%) = " "
  dInfo$(7, 0, m%) = "Rezeptkunden %"
  If Vorschau Then
    dInfo$(7, 1, m%) = aDetail$(8, m%, 1)
  Else
    dInfo$(7, 1, m%) = aDetail$(8, m%, aTage%)
  End If
  dInfo$(7, 2, m%) = "#"
  
  
  dInfo$(8, 0, m%) = "Anzahl Rabatte"
  If Vorschau Then
    dInfo$(8, 1, m%) = Format(AnzRab!(a%), "##,###")
  Else
    dInfo$(8, 1, m%) = aInfo$(5, 1, m%)
  End If
  dInfo$(8, 2, m%) = " "
  dInfo$(9, 0, m%) = "Rabattsumme"
  If Vorschau Then
    dInfo$(9, 1, m%) = Format(SumRab#(a%), "##,###.##")
  Else
    dInfo$(9, 1, m%) = aInfo$(6, 1, m%)
  End If
  dInfo$(9, 2, m%) = " "
  
  dInfo$(10, 0, m%) = "Summe Normtage"
  If Vorschau Then
    dInfo$(10, 1, m%) = Format(NT!, "###0.0")
  Else
    dInfo$(10, 1, m%) = aInfo$(0, 3, m%)
  End If
  dInfo$(10, 2, m%) = " "
  dInfo$(11, 0, m%) = "Kunden/Normtag"
  If Vorschau Then
    If NT! <> 0 Then
      dInfo$(11, 1, m%) = Format(AnzKd!(0, 1, a%) / NT!, "##,###")
    End If
  Else
    dInfo$(11, 1, m%) = aInfo$(1, 3, m%)
  End If
  dInfo$(11, 2, m%) = "#"
  dInfo$(12, 0, m%) = "Rezepte/Normtag"
  If Vorschau Then
    If NT! <> 0 Then
      dInfo$(12, 1, m%) = Format(AnzRp!(0, 1, a%) / NT!, "##,###")
    End If
  Else
    dInfo$(12, 1, m%) = aInfo$(2, 3, m%)
  End If
  dInfo$(12, 2, m%) = "#"
  dInfo$(13, 0, m%) = "Sonderpräm.Kunden/Normtag"
  If Vorschau Then
    If NT! <> 0 Then
      dInfo$(13, 1, m%) = Format(AnzSPrmKd!(0, 1, a%) / NT!, "##,###")
    End If
  Else
    dInfo$(13, 1, m%) = aInfo$(3, 3, m%)
  End If
  dInfo$(13, 2, m%) = "#"
  dInfo$(14, 0, m%) = "Zusatzverk.Kunden/Normtag"
  If Vorschau Then
    If NT! <> 0 Then
      dInfo$(14, 1, m%) = Format(AnzZusKd!(0, 1, a%) / NT!, "##,###")
    End If
  Else
    dInfo$(14, 1, m%) = aInfo$(4, 3, m%)
  End If
  dInfo$(14, 2, m%) = "#"
  dInfo$(15, 0, m%) = "% Zusatzverk.Kunden/Rezeptkunden"
  If Vorschau Then
    If AnzRezKd!(0, 1, a%) <> 0 Then
      dInfo$(15, 1, m%) = Format(AnzZusKd!(0, 1, a%) * 100! / AnzRezKd!(0, 1, a%), "##0.0")
    End If
  Else
    dInfo$(15, 1, m%) = aInfo$(5, 3, m%)
  End If
  dInfo$(15, 2, m%) = "#"
  dInfo$(16, 0, m%) = "durchschn. Zusatzverkauf"
  If Vorschau Then
    If AnzZusKd!(0, 1, a%) <> 0 Then
      dInfo$(16, 1, m%) = Format(ZusatzUms#(0, 1, a%) / AnzZusKd!(0, 1, a%), "##,##0.00")
    End If
  Else
    dInfo$(16, 1, m%) = aInfo$(6, 3, m%)
  End If
  dInfo$(16, 2, m%) = "#"

  dInfo$(17, 0, m%) = "durchschn. Barverkauf/Privatkunde"
  If Vorschau Then
    If AnzKd!(0, 1, a%) - AnzRezKd!(0, 1, a%) <> 0 Then
      dInfo$(17, 1, m%) = Format(PrivUms#(0, 1, a%) / (AnzKd!(0, 1, a%) - AnzRezKd!(0, 1, a%)), "##,##0.00")
    End If
  Else
    dInfo$(17, 1, m%) = aInfo$(0, 5, m%)
  End If
  dInfo$(17, 2, m%) = "#"
  dInfo$(18, 0, m%) = "Prämienbasis/Kunde"
  If Vorschau Then
    If AnzKd!(0, 1, a%) <> 0 Then
      dInfo$(18, 1, m%) = Format(Basis#(0, 0, a%) / AnzKd!(0, 1, a%), "##,##0.00")
    End If
  Else
    dInfo$(18, 1, m%) = aInfo$(1, 5, m%)
  End If
  dInfo$(18, 2, m%) = "#"
  dInfo$(19, 0, m%) = "PrämienBasis/PrämienKunde"
  If Vorschau Then
    If AnzPKd!(0, 1, a%) <> 0 Then
      dInfo$(19, 1, m%) = Format(Basis#(0, 0, a%) / AnzPKd!(0, 1, a%), "##,##0.00")
    End If
  Else
    dInfo$(19, 1, m%) = aInfo$(2, 5, m%)
  End If
  dInfo$(19, 2, m%) = "#"
  dInfo$(20, 0, m%) = "Prämie"
  If Vorschau Then
    PraemienSumme# = Basis#(0, 0, a%) / 100 * pb#
    dInfo$(20, 1, m%) = Format(Basis#(0, 0, a%) / 100 * pb#, "###,##0.00")
  Else
    dInfo$(20, 1, m%) = aInfo$(3, 5, m%)
  End If
  dInfo$(20, 2, m%) = " "
  dInfo$(21, 0, m%) = "Zusatzprämie"
  If Vorschau Then
    PraemienSumme# = PraemienSumme# + Zusatz#(0, 0, a%) / 100 * zpb#
    dInfo$(21, 1, m%) = Format(Zusatz#(0, 0, a%) / 100 * zpb#, "###,##0.00")
  Else
    dInfo$(21, 1, m%) = aInfo$(4, 5, m%)
  End If
  dInfo$(21, 2, m%) = " "
  dInfo$(22, 0, m%) = "Sonderprämie"
  If Vorschau Then
    PraemienSumme# = PraemienSumme# + Sonder#(0, 0, a%) / 100 * spb#
    dInfo$(22, 1, m%) = Format(Sonder#(0, 0, a%) / 100 * spb#, "###,##0.00")
  Else
    dInfo$(22, 1, m%) = aInfo$(5, 5, m%)
  End If
  dInfo$(22, 2, m%) = " "
  dInfo$(23, 0, m%) = "Prämiensumme"
  If Vorschau Then
    dInfo$(23, 1, m%) = Format(PraemienSumme#, "###,##0.00")
  Else
    dInfo$(23, 1, m%) = aInfo$(6, 5, m%)
  End If
  dInfo$(23, 2, m%) = " "

  dInfo$(24, 0, m%) = "Anzahl kleine Pausen"
  If Vorschau Then
    dInfo$(24, 1, m%) = aDetail$(11, m%, 1)
  Else
    dInfo$(24, 1, m%) = aDetail$(11, m%, aTage%)
  End If
  dInfo$(24, 2, m%) = " "
  dInfo$(25, 0, m%) = "Dauer kleine Pausen"
  If Vorschau Then
    dInfo$(25, 1, m%) = aDetail$(12, m%, 1)
  Else
    dInfo$(25, 1, m%) = aDetail$(12, m%, aTage%)
  End If
  dInfo$(25, 2, m%) = " "
  dInfo$(26, 0, m%) = "Anzahl große Pausen"
  If Vorschau Then
    dInfo$(26, 1, m%) = aDetail$(13, m%, 1)
  Else
    dInfo$(26, 1, m%) = aDetail$(13, m%, aTage%)
  End If
  dInfo$(26, 2, m%) = " "
  dInfo$(27, 0, m%) = "Dauer große Pausen"
  If Vorschau Then
    dInfo$(27, 1, m%) = aDetail$(14, m%, 1)
  Else
    dInfo$(27, 1, m%) = aDetail$(14, m%, aTage%)
  End If
  dInfo$(27, 2, m%) = " "
  dInfo$(HAT_WERTE_SPALTE%, 0, m%) = "Kundenanzahl"
  If Vorschau Then
    dInfo$(HAT_WERTE_SPALTE%, 1, m%) = aDetail$(4, m%, 1)
  Else
    dInfo$(HAT_WERTE_SPALTE%, 1, m%) = aDetail$(4, m%, aTage%)
  End If
  dInfo$(HAT_WERTE_SPALTE%, 2, m%) = " "
  
  dInfo$(29, 0, m%) = "Dauer kl.Pausen/Normtag"
  If Vorschau Then
    If (NT! = 0) Or (Val(aDetail$(12, m%, 1)) = 0) Then
      dInfo$(29, 1, m%) = ""
    Else
      dInfo$(29, 1, m%) = Format(CLng(aDetail$(12, m%, 1)) / NT!, "0")
    End If
  Else
  
    If (xVal(aInfo$(0, 3, m%)) = 0) Or (Val(aDetail$(12, m%, aTage%)) = 0) Then
      dInfo$(29, 1, m%) = ""
    Else
      dInfo$(29, 1, m%) = Format(CLng(aDetail$(12, m%, aTage%)) / xVal(aInfo$(0, 3, m%)), "0")
    End If
  End If
  dInfo$(29, 2, m%) = " "
  
Next m%

If Vorschau Then    'das war die Summenzeile
    m% = 0
    If Basis#(0, 1, 0) <> 0 Then
      aDetail$(1, m%, zeile%) = Format(Basis#(0, 1, 0), "###,###")
    End If
    
    If Zusatz#(0, 1, 0) <> 0 Then
      aDetail$(2, m%, zeile%) = Format(Zusatz#(0, 1, 0), " ###,###")
    End If
    If Sonder#(0, 1, 0) <> 0 Then
      aDetail$(3, m%, zeile%) = Format(Sonder#(0, 1, 0), " ###,###")
    End If
    
    If AnzKd!(0, 1, 0) <> 0 Then
      aDetail$(4, m%, zeile%) = Format(AnzKd!(0, 1, 0), " ##,###")
    End If
    
    If AnzRp!(0, 1, 0) <> 0 Then
      aDetail$(5, m%, zeile%) = Format(AnzRp!(0, 1, 0), " ##,###")
    End If

    If AnzKd!(0, 1, 0) <> 0 Then
      aDetail$(6, m%, zeile%) = Format(AnzPKd!(0, 1, 0) * 100! / AnzKd!(0, 1, 0), " ##0.0")
    End If
    
    zw! = AnzKd!(0, 1, 0) - AnzRezKd!(0, 1, 0)
    If zw! <> 0 And AnzKd!(0, 1, 0) <> 0 Then
      aDetail$(7, m%, zeile%) = Format(zw! * 100 / AnzKd!(0, 1, 0), " ##0.0")
    End If
    If AnzRezKd!(0, 1, 0) <> 0 And AnzKd!(0, 1, 0) <> 0 Then
      aDetail$(8, m%, zeile%) = Format(AnzRezKd!(0, 1, 0) * 100! / AnzKd!(0, 1, 0), " ##0.0")
    End If
    
    If Pause1Anz%(0, 1, 0) > 0 Then
      aDetail$(11, m%, zeile%) = Format(Pause1Anz%(0, 1, 0), "## ")
      aDetail$(12, m%, zeile%) = Format(Pause1Zeit(0, 1, 0), "####")
    End If
          
    If Pause2Anz%(0, 1, 0) > 0 Then
      aDetail$(13, m%, zeile%) = Format(Pause2Anz%(0, 1, 0), "##")
      aDetail$(14, m%, zeile%) = Format(Pause2Zeit(0, 1, 0), "####")
    End If
        
'Monatssummen für Vorschau
  
    a% = 0
    NT! = gAnzNormTage!
    GoSub SummenZeilen
    

'        Print #pr, Space$(15); "Prämienbasis  Faktor      Prämie"
'        Print #pr, String$(47, "-")
'        Print #pr, "Grundprämie   :  ";
'        BasisSumme# = Basis#(0, 0, 0)
'        If BasisSumme# <> 0 Then
'                Print #pr, USING; "###,###.##"; BasisSumme#;
'                Print #pr, USING; "  ###.##"; pb#;
'                BasisPraemie# = BasisSumme# / 100 * pb#
'                'BasisSumme# = Basis#(0, 0, 0)
'                PraemienSumme# = BasisPraemie#
'                Print #pr, USING; "  ###,###.##"; BasisPraemie#
'        Else
'                Print #pr, ""
'        End If
'
'        Print #pr, "Sonderprämie  :  ";
'        Print #pr, ""
'
'        Print #pr, "Zusatzverkäufe:  ";
'        Print #pr, ""
'        Print #pr, String$(47, "-")
'
'        If BasisSumme# <> 0 Then
'                Print #pr, Space$(8); "Summe: "; USING; "#,###,###.##"; BasisSumme#;
'                Print #pr, Space$(8); USING; "#,###,###.##"; PraemienSumme#
'        Else
'                Print #pr, ""
'        End If
'        Print #pr, Space$(15); String$(32, "-")
'
End If

Call SortArray(AnzNT(), CLng(anzahl%), True)
With frmAction.cboMitarbeiter
  .Clear
  .AddItem "Gesamt"
  For m% = 1 To anzahl%
    h$ = RTrim$(Left$(MitArb$(AnzNT(m%, 1)), 20))
    If (h$ <> "") And iMitArb(AnzNT(m%, 1)) Then
      If Vorschau Then
        '???GS
        ReDim Preserve DiagPnr(UBound(DiagPnr) + 1)
        DiagPnr(UBound(DiagPnr)) = AnzNT(m%, 1)
        aDetail$(0, 0, m%) = CStr(AnzNT(m%, 1))
        If PersonalInitialen(AnzNT(m%, 1)) > "" Then aDetail$(0, 0, m%) = PersonalInitialen(AnzNT(m%, 1))
        For spalte% = 1 To 15
          aDetail$(spalte%, 0, m%) = aDetail$(spalte%, AnzNT(m%, 1), 1)
        Next spalte%
      Else
        .AddItem h$
      End If
    End If
  Next m%
  .ListIndex = 0
End With

Call DefErrPop
Exit Sub

AuswertungDruckenKopf:
'        Text$ = Space$(78)
'        p$ = "Pers-Nr. " + Right$("  " + Str$(a%), 2) + " - " + Left$(MitArb$(m%), 20)
'        p$ = Trim(p$)
'        Mid$(Text$, 1) = "Auswertung für " + Format(von, "dd.mm.yy") + " - " + Format(bis, "dd.mm.yy")
'        Mid$(Text$, 78 - Len(p$)) = p$
'
'        Text$ = Space$(78)
'        Mid$(Text$, 1) = "Prämienbasis: "
'        If p3.PrmBas = "I" Then
'                Mid$(Text$, 16) = "AVP inkl. USt."
'        ElseIf p3.PrmBas = "E" Then
'                Mid$(Text$, 16) = "AVP exkl. USt."
'        Else
'                Mid$(Text$, 16) = "Spanne"
'        End If
'        Mid$(Text$, Len(Text$) - 26) = Left$(heute$, 2) + "." + Mid$(heute$, 3, 2) + "." + Left$(DateLong$(heute$), 4)
'        Mid$(Text$, Len(Text$) - 13) = Seite$
'        Print #pr, ".KOPF"; Text$
'        Print #pr, ".KOPF"; strich$
Return
'----------------------------------------------------------------------------------------------------------------------
AuswertungDruckenKopfXT:

If Vorschau Then
  frmAction.Caption = "Personal-Verkaufstatistik: Vorschau für " + Format(von, "dd.mm.yy")
Else
  frmAction.Caption = "Personal-Verkaufstatistik: Auswertung für " + Format(von, "dd.mm.yy") + " - " + Format(Bis, "dd.mm.yy")
End If
frmAction!lblPrmBasis.Caption = "Prämienbasis: "

If p3.PrmBas = "I" Then
  frmAction!lblPrmBasis.Caption = frmAction!lblPrmBasis.Caption + "AVP inkl. MwSt."
ElseIf p3.PrmBas = "E" Then
  frmAction!lblPrmBasis.Caption = frmAction!lblPrmBasis.Caption + "AVP exkl. MwSt."
Else
  frmAction!lblPrmBasis.Caption = frmAction!lblPrmBasis.Caption + "Spanne"
End If
If Vorschau Then
  frmAction!lblPrmBasis.Caption = frmAction!lblPrmBasis.Caption + " - Vorschau für " + Format(von, "dd.mm.yy")
End If
Return
'----------------------------------------------------------------------------------------------------------------------
SummenZeilen:

aInfo$(0, 0, m%) = "Kunden mit Prämie"
aInfo$(1, 0, m%) = "Kunden mit Zusatzverkauf"
aInfo$(2, 0, m%) = "Kunden mit Sonderprämie"
aInfo$(3, 0, m%) = "Privatkunden"
aInfo$(4, 0, m%) = "Rezeptkunden"
aInfo$(5, 0, m%) = "Anzahl Rabatte"
aInfo$(6, 0, m%) = "Rabattsumme"

aInfo$(0, 2, m%) = "Summe Normtage"
aInfo$(1, 2, m%) = "Kunden/Normtag"
aInfo$(2, 2, m%) = "Rezepte/Normtag"
aInfo$(3, 2, m%) = "Sonderpräm.Kunden/Normtag"
aInfo$(4, 2, m%) = "Zusatzverk.Kunden/Normtag"
aInfo$(5, 2, m%) = "% Zusatzverk.Kunden/Rezeptkunden"
aInfo$(6, 2, m%) = "durchschn.  Zusatzverkauf"

aInfo$(0, 4, m%) = "durchschn. Barverkauf/Privatkunde"
aInfo$(1, 4, m%) = "Prämienbasis/Kunde"
aInfo$(2, 4, m%) = "PrämienBasis/PrämienKunde"
aInfo$(3, 4, m%) = "Prämie"
aInfo$(4, 4, m%) = "Zusatzprämie"
aInfo$(5, 4, m%) = "Sonderprämie"
aInfo$(6, 4, m%) = "Prämiensumme"


'"Kunden mit Prämie"
If AnzPKd!(0, 1, a%) <> 0 Then
  aInfo$(0, 1, m%) = Format(AnzPKd!(0, 1, a%), "##,###")
End If

'"Kunden mit Zusatzverkauf
If AnzZusKd!(0, 1, a%) > 0 Then
  aInfo$(1, 1, m%) = Format(AnzZusKd!(0, 1, a%), "##,###")
End If

'"Kunden mit Sonderprämie
If AnzSPrmKd!(0, 1, a%) > 0 Then
  aInfo$(2, 1, m%) = Format(AnzSPrmKd!(0, 1, a%), "##,###")
End If

'"Privatkunden"
If AnzKd!(0, 1, a%) - AnzRezKd!(0, 1, a%) > 0 Then
  aInfo$(3, 1, m%) = Format(AnzKd!(0, 1, a%) - AnzRezKd!(0, 1, a%), "##,###")
End If

'"Rezeptkunden"
If AnzRezKd!(0, 1, a%) > 0 Then
  aInfo$(4, 1, m%) = Format(AnzRezKd!(0, 1, a%), "##,###")
End If


'"Anzahl Rabatte"
'"Rabattsumme"
If p3.RabAbzg = "N" Then
  aInfo$(5, 0, m%) = "Keine Rabatte berücksichtigt"
  aInfo$(6, 0, m%) = ""
Else
  If AnzRab!(a%) <> 0 Then
    aInfo$(5, 0, m%) = "Anzahl Rabatte"
    aInfo$(5, 1, m%) = Format(AnzRab!(a%), "##,###")
    aInfo$(6, 0, m%) = "Rabattsumme"
    aInfo$(6, 1, m%) = Format(SumRab#(a%), "##,###.##")
  Else
    aInfo$(5, 0, m%) = "Keine Rabatte gewährt"
    aInfo$(6, 0, m%) = ""
  End If
End If

'"Summe Normtage"
If m% = 0 Then
  aInfo$(0, 2, m%) = "Summe Normtage"
Else
  aInfo$(0, 2, m%) = "Anwesende Normtage"
End If
aInfo$(0, 3, m%) = Format(NT!, "###0.0")

'"Kunden/Normtag"
If NT! > 0! Then
  aInfo$(1, 3, m%) = Format(AnzKd!(0, 1, a%) / NT!, "##,###")
End If

'"Rezepte/Normtag"
If NT! > 0! Then aInfo$(2, 3, m%) = Format(AnzRp!(0, 1, a%) / NT!, "##,###")

'"Sonderprämienkunden/Normtag"
If NT! > 0! Then
  aInfo$(3, 3, m%) = Format(AnzSPrmKd!(0, 1, a%) / NT!, "##,###")
End If

'"Zusatzverk.Kunden/Normtag"
If NT! > 0! Then
  aInfo$(4, 3, m%) = Format(AnzZusKd!(0, 1, a%) / NT!, "##,###")
End If

'"% Zusatzverk.Kunden/Rezeptkunden"
If AnzRezKd!(0, 1, a%) > 0! Then
  aInfo$(5, 3, m%) = Format(AnzZusKd!(0, 1, a%) * 100! / AnzRezKd!(0, 1, a%), "##0.0")
End If

'"AVG Zusatzverkauf"
If AnzZusKd!(0, 1, a%) <> 0 Then
  aInfo$(6, 3, m%) = Format(ZusatzUms#(0, 1, a%) / AnzZusKd!(0, 1, a%), "##,##0.00")
End If

'aInfo$(0, 4, m%) = "AVG Losung/Privatkunde"
If AnzKd!(0, 1, a%) - AnzRezKd!(0, 1, a%) > 0 Then
  aInfo$(0, 5, m%) = Format(PrivUms#(0, 1, a%) / (AnzKd!(0, 1, a%) - AnzRezKd!(0, 1, a%)), "##,##0.00")
End If

'aInfo$(1, 4, m%) = "Prämienbasis/Kunde"
If AnzKd!(0, 1, a%) > 0 Then
  aInfo$(1, 5, m%) = Format(Basis#(0, 0, a%) / AnzKd!(0, 1, a%), "##,##0.00")
End If
   
'aInfo$(2, 4, m%) = "PrämienBasis/PrämienKunde"
If AnzPKd!(0, 1, a%) > 0 Then
  aInfo$(2, 5, m%) = Format(Basis#(0, 0, a%) / AnzPKd!(0, 1, a%), "##,##0.00")
End If

aInfo$(3, 4, m%) = "Prämie"
BasisPraemie# = Basis#(0, 0, a%) / 100 * pb#
PraemienSumme# = BasisPraemie#
If BasisPraemie# <> 0# Then
  aInfo$(3, 5, m%) = Format(BasisPraemie#, "###,##0.00")
End If

aInfo$(4, 4, m%) = "Zusatzprämie"
ZusatzPraemie# = Zusatz#(0, 0, a%) / 100 * zpb#
PraemienSumme# = PraemienSumme# + ZusatzPraemie#
If ZusatzPraemie# <> 0# Then
  aInfo$(4, 5, m%) = Format(ZusatzPraemie#, "###,##0.00")
End If

aInfo$(5, 4, m%) = "Sonderprämie"
SonderPraemie# = Sonder#(0, 0, a%) / 100 * spb#
PraemienSumme# = PraemienSumme# + SonderPraemie#
If SonderPraemie# <> 0# Then
  aInfo$(5, 5, m%) = Format(SonderPraemie#, "###,##0.00")
End If

aInfo$(6, 4, m%) = "Prämiensumme"
If PraemienSumme# <> 0# Then
  aInfo$(6, 5, m%) = Format(PraemienSumme#, "###,##0.00")
End If

Return

End Sub


Sub SortArray(Feld() As Single, SortMax&, abst As Boolean)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("SortArray")
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

Dim HilfsFeld(1) As Single
Dim zw(1) As Single
Dim a%, b%
Dim ia&, ib&, ic&, Id&

If abst Then
  HilfsFeld(0) = 3.4 ^ 38
Else
  HilfsFeld(0) = -3.4 ^ 38
End If
If SortMax& > 1 Then
  If SortMax& = 2 Then
    If abst Then
      If Feld(1, 0) < Feld(2, 0) Then
        a% = 1: b% = 2
        GoSub SwapFields
      End If
    Else
      If Feld(1) > Feld(2) Then
        a% = 1: b% = 2
        GoSub SwapFields
      End If
    End If
  Else
    ib& = SortMax& - 1
    While ib& > 1
      ib& = Int(ib& * 0.3 + 0.5)
      For ic& = 1 To ib&
        For Id& = ic& + ib& To SortMax& Step ib&
          HilfsFeld(0) = Feld(Id&, 0)
          HilfsFeld(1) = Feld(Id&, 1)
          For ia& = Id& - ib& To 1 Step -ib&
            If abst Then
              If HilfsFeld(0) <= Feld(ia&, 0) Then Exit For
            Else
              If HilfsFeld(0) >= Feld(ia&, 0) Then Exit For
            End If
            Feld(ia& + ib&, 0) = Feld(ia&, 0)
            Feld(ia& + ib&, 1) = Feld(ia&, 1)
          Next ia&
          zw(0) = Feld(ia& + ib&, 0)
          zw(1) = Feld(ia& + ib&, 1)
          Feld(ia& + ib&, 0) = HilfsFeld(0)
          Feld(ia& + ib&, 1) = HilfsFeld(1)
          HilfsFeld(0) = zw(0)
          HilfsFeld(1) = zw(1)
        Next Id&
      Next ic&
    Wend
  End If
End If
Call DefErrPop
Exit Sub

SwapFields:
zw(0) = Feld(a%, 0)
zw(1) = Feld(a%, 1)
Feld(a%, 0) = Feld(b%, 0)
Feld(a%, 1) = Feld(b%, 1)
Feld(b%, 0) = zw(0)
Feld(b%, 1) = zw(1)
      

Return
End Sub


Sub AuswertungRechnen(gefunden As Boolean, von As Date, Bis As Date)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("AuswertungRechnen")
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
Dim LaufNr%, fout%
Dim buf As String
Dim maxPss As Long, begpss As Long, ptrPss As Long
Dim aktzeit%, dd%, d%, WoTag%, Pause%, g%, s%, LastPnr%
Dim datum As Date
Dim AnyRez As Boolean, KuEnde As Boolean
Dim KdUmsatz#, ZusatzBasis#, Zusatzumsatz#, SonderBasis#
Dim TagAnf%, TagEnd%, zw%
Dim FeierTag As Boolean

Dim SuchDatumC$

'fout% = fopen("G:\user\PSTW.$$$", "O")

buf = String(64, 0)
Get #fpss, 1, buf
maxPss = CVS(Left(buf, 4)) - 1

SuchDatumC$ = MKDatum(von)

Call TagSuchen(gefunden, begpss, SuchDatumC$, maxPss)
begpss = begpss - 1
dd% = 0
KuEnde = True
For ptrPss = begpss To maxPss
  Get #fpss, (ptrPss * 64!) + 1, p
  
  
  'es fehlen zwar einige Kundenenden, aber ohne diese stimmen wenigstens die Zahlen -> nur Merkvariablen löschen
  If p.Pnr <> LastPnr% And Not KuEnde Then
'    Get #fpss, ((ptrPss - 1) * 64!) + 1, p
'    p.KuEnde = 1
''    Put #fpss, ((ptrPss - 1) * 64!) + 1, p
'    GoSub KdRpJeTag
'    GoSub KdRpJeGeraet
'    If p.KZ = "P" Or (p.RezNr = 0 And p.pzn = "9999998") Then GoSub PrmKdJeTag
'    Get #fpss, (ptrPss * 64!) + 1, p
    AnyRez = False
    ZusatzBasis# = 0
    Zusatzumsatz# = 0#
    SonderBasis# = 0
    KdUmsatz# = 0#

  End If
  KuEnde = False
  LastPnr% = p.Pnr
  aktzeit% = p.zeit
  datum = CVDatum(p.datum)
'  If p.LaufNr <> LaufNr% Then TagKd = TagKd + 1
'  If Format(datum, "dd") = "17" And p.LaufNr <> LaufNr% Then
'    Print #fout%, p.LaufNr
'  End If
  LaufNr% = p.LaufNr
  If datum >= vonAuswD And datum <= bisAuswD Then
'      If Format(datum, "dd") = "19" And aktzeit% >= 844 Then Stop
    d% = Val(Format(datum, "dd"))
    If d% > dd% Then  'Datumswechsel
'      TagKd = 1
      dd% = d%
      WoTag% = WeekDay(datum)
      FeierTag = IstFeiertag(datum)
      
      If WoTag% <> vbSunday Then
        TagAnf% = OffenRec(WoTag% - 2).iAbOffen
        If TagAnf% Mod 100 > ToleranzOffen% Then
          TagAnf% = TagAnf% \ 100 + (TagAnf% Mod 100 - ToleranzOffen%)
        Else
          zw% = TagAnf% Mod 100 + 60 - ToleranzOffen%
          TagAnf% = (TagAnf% \ 100 - 1) * 100 + zw%
        End If
        TagEnd% = OffenRec(WoTag% - 2).iBisOffen
        If (TagEnd% Mod 100 + ToleranzOffen%) < 60 Then
          TagEnd% = TagEnd% + ToleranzOffen%
        Else
          TagEnd% = (TagEnd% \ 100 + 1) * 100
        End If
      End If
    End If
    d% = DateValue(datum) - DateValue(vonAuswD) + 1
    
    If WoTag% > vbSunday And Not FeierTag Then
      'If p.Pnr = 23 Then Stop
      If (iMitArb(p.Pnr) Or Vorschau) And aktzeit% >= TagAnf% And aktzeit% <= TagEnd% Then

        If erstKd%(d%, 1, p.Pnr) = 0 Then erstKd%(d%, 1, p.Pnr) = aktzeit%
        If letztKd%(d%, 1, p.Pnr) < aktzeit% Then
          If letztKd%(d%, 1, p.Pnr) <> 0 Then
            Pause% = fMinuten%(aktzeit%) - fMinuten%(letztKd%(d%, 1, p.Pnr))
            If Pause% >= PauseGr% Then
              Pause2Anz%(d%, 1, p.Pnr) = Pause2Anz%(d%, 1, p.Pnr) + 1
              Pause2Zeit(d%, 1, p.Pnr) = Pause2Zeit(d%, 1, p.Pnr) + Pause%
              Pause2Anz%(d%, 1, 0) = Pause2Anz%(d%, 1, 0) + 1
              Pause2Zeit(d%, 1, 0) = Pause2Zeit(d%, 1, 0) + Pause%
              Pause2Anz%(0, 1, 0) = Pause2Anz%(0, 1, 0) + 1
              Pause2Zeit(0, 1, 0) = Pause2Zeit(0, 1, 0) + Pause%
            ElseIf Pause% >= PauseKl% Then
              Pause1Anz%(d%, 1, p.Pnr) = Pause1Anz%(d%, 1, p.Pnr) + 1
              Pause1Zeit(d%, 1, p.Pnr) = Pause1Zeit(d%, 1, p.Pnr) + Pause%
              Pause1Anz%(d%, 1, 0) = Pause1Anz%(d%, 1, 0) + 1
              Pause1Zeit(d%, 1, 0) = Pause1Zeit(d%, 1, 0) + Pause%
              Pause1Anz%(0, 1, 0) = Pause1Anz%(0, 1, 0) + 1
              Pause1Zeit(0, 1, 0) = Pause1Zeit(0, 1, 0) + Pause%
            End If
          End If
          letztKd%(d%, 1, p.Pnr) = aktzeit%
        End If
        g% = p.User
        s% = p.BS + 1
        If p.KZ = "P" Or p.KZ = "S" Then    'da ist eine Prämie fällig
          ZusatzBasis# = ZusatzBasis# + CVD(p.Basis) * p.Multi
          Zusatzumsatz# = Zusatzumsatz# + CVD(p.Preis) * p.Multi
          If p.KZ = "P" Then
            'je Mitarb
'            If Abs(CVD(p.Basis) * p.Multi) > 150 Then Stop
            Basis#(d%, 1, p.Pnr) = Basis#(d%, 1, p.Pnr) + CVD(p.Basis) * p.Multi
            Basis#(d%, 0, p.Pnr) = Basis#(d%, 0, p.Pnr) + CVD(p.Basis) * p.Multi
            Basis#(0, 1, p.Pnr) = Basis#(0, 1, p.Pnr) + CVD(p.Basis) * p.Multi
            Basis#(0, 0, p.Pnr) = Basis#(0, 0, p.Pnr) + CVD(p.Basis) * p.Multi
            'alle Mitarb
            Basis#(d%, 1, 0) = Basis#(d%, 1, 0) + CVD(p.Basis) * p.Multi
            Basis#(d%, 0, 0) = Basis#(d%, 0, 0) + CVD(p.Basis) * p.Multi
            Basis#(0, 1, 0) = Basis#(0, 1, 0) + CVD(p.Basis) * p.Multi
            Basis#(0, 0, 0) = Basis#(0, 0, 0) + CVD(p.Basis) * p.Multi
            'Rabatte
            If p.pFlag <> 0 And CVD(p.RabBetr) <> 0 Then
              AnzRab!(p.Pnr) = AnzRab!(p.Pnr) + 1
              SumRab#(p.Pnr) = SumRab#(p.Pnr) + CVD(p.RabBetr)
              AnzRab!(0) = AnzRab!(0) + 1
              SumRab#(0) = SumRab#(0) + CVD(p.RabBetr)
            End If
          End If

          'Geräte je Mitarb
          gAnzPB!(g%, s%, p.Pnr) = gAnzPB!(g%, s%, p.Pnr) + 1
          gAnzPB!(10, 0, p.Pnr) = gAnzPB!(10, 0, p.Pnr) + 1
          gPB#(g%, s%, p.Pnr) = gPB#(g%, s%, p.Pnr) + CVD(p.Basis) * p.Multi
          gPB#(10, 0, p.Pnr) = gPB#(10, 0, p.Pnr) + CVD(p.Basis) * p.Multi
          'Geräte alle Mitarb
          gAnzPB!(g%, s%, 0) = gAnzPB!(g%, s%, 0) + 1
          gAnzPB!(10, 0, 0) = gAnzPB!(10, 0, 0) + 1
          gPB#(g%, s%, 0) = gPB#(g%, s%, 0) + CVD(p.Basis) * p.Multi
          gPB#(10, 0, 0) = gPB#(10, 0, 0) + CVD(p.Basis) * p.Multi

          'Sonder-/Zusatzprämien
          If CVS(p.sPrm) > 0 Then
            SonderBasis# = SonderBasis# + CVS(p.sPrm) * p.Multi
            'je Mitarb
            Sonder#(d%, 1, p.Pnr) = Sonder#(d%, 1, p.Pnr) + CVS(p.sPrm) * p.Multi
            Sonder#(d%, 0, p.Pnr) = Sonder#(d%, 0, p.Pnr) + CVS(p.sPrm) * p.Multi
            Sonder#(0, 1, p.Pnr) = Sonder#(0, 1, p.Pnr) + CVS(p.sPrm) * p.Multi
            Sonder#(0, 0, p.Pnr) = Sonder#(0, 0, p.Pnr) + CVS(p.sPrm) * p.Multi
            'alle Mitarb
            Sonder#(d%, 1, 0) = Sonder#(d%, 1, 0) + CVS(p.sPrm) * p.Multi
            Sonder#(d%, 0, 0) = Sonder#(d%, 0, 0) + CVS(p.sPrm) * p.Multi
            Sonder#(0, 1, 0) = Sonder#(0, 1, 0) + CVS(p.sPrm) * p.Multi
            Sonder#(0, 0, 0) = Sonder#(0, 0, 0) + CVS(p.sPrm) * p.Multi
          End If
          '??? wozu brauche ich den Umsatz?
          Umsatz#(p.Pnr) = Umsatz#(p.Pnr) + CVD(p.Preis) * p.Multi
          Umsatz#(0) = Umsatz#(0) + CVD(p.Preis) * p.Multi
          GoSub KdRpJeTag
          GoSub KdRpJeGeraet
          If p.KZ = "P" Then GoSub PrmKdJeTag
        ElseIf p.KZ = "T" Then    '
          GoSub KdRpJeTag
          GoSub KdRpJeGeraet
          If p.RezNr = 0 And p.pzn = "9999998" Then GoSub PrmKdJeTag
        End If
      Else
              'break
      End If
    End If
  End If
  If datum > Bis Then Exit For
  LastPnr% = p.Pnr
  'Esc abfragen --> brkFlg=true
Next ptrPss
'Close #fout%
Call DefErrPop
Exit Sub
'----------------------------------------------------------------------------------------------------------------------
KdRpJeTag:
KdUmsatz# = KdUmsatz# + CVD(p.Preis) * p.Multi
If p.RezEnde <> 0 Then AnyRez = True
If p.KuEnde <> 0 Then
  KuEnde = True
  If AnyRez Then
    'je Mitarb
    AnzRezKd!(d%, 1, p.Pnr) = AnzRezKd!(d%, 1, p.Pnr) + p.KuEnde
    AnzRezKd!(d%, 0, p.Pnr) = AnzRezKd!(d%, 0, p.Pnr) + p.KuEnde
    AnzRezKd!(0, 1, p.Pnr) = AnzRezKd!(0, 1, p.Pnr) + p.KuEnde
    'alle Mitarb
    AnzRezKd!(d%, 1, 0) = AnzRezKd!(d%, 1, 0) + p.KuEnde
    AnzRezKd!(d%, 0, 0) = AnzRezKd!(d%, 0, 0) + p.KuEnde
    AnzRezKd!(0, 1, 0) = AnzRezKd!(0, 1, 0) + p.KuEnde
    If ZusatzBasis# <> 0# Then
      'Zusatzverkäufe
      'je Mitarb
      AnzZusKd!(d%, 1, p.Pnr) = AnzZusKd!(d%, 1, p.Pnr) + p.KuEnde
      AnzZusKd!(d%, 0, p.Pnr) = AnzZusKd!(d%, 0, p.Pnr) + p.KuEnde
      AnzZusKd!(0, 1, p.Pnr) = AnzZusKd!(0, 1, p.Pnr) + p.KuEnde
      Zusatz#(d%, 1, p.Pnr) = Zusatz#(d%, 1, p.Pnr) + ZusatzBasis#
      Zusatz#(d%, 0, p.Pnr) = Zusatz#(d%, 0, p.Pnr) + ZusatzBasis#
      Zusatz#(0, 1, p.Pnr) = Zusatz#(0, 1, p.Pnr) + ZusatzBasis#
      Zusatz#(0, 0, p.Pnr) = Zusatz#(0, 0, p.Pnr) + ZusatzBasis#
      ZusatzUms#(d%, 1, p.Pnr) = ZusatzUms#(d%, 1, p.Pnr) + Zusatzumsatz#
      ZusatzUms#(d%, 0, p.Pnr) = ZusatzUms#(d%, 0, p.Pnr) + Zusatzumsatz#
      ZusatzUms#(0, 1, p.Pnr) = ZusatzUms#(0, 1, p.Pnr) + Zusatzumsatz#
      'alle Mitarb
      AnzZusKd!(d%, 1, 0) = AnzZusKd!(d%, 1, 0) + p.KuEnde
      AnzZusKd!(d%, 0, 0) = AnzZusKd!(d%, 0, 0) + p.KuEnde
      AnzZusKd!(0, 1, 0) = AnzZusKd!(0, 1, 0) + p.KuEnde
      Zusatz#(d%, 1, 0) = Zusatz#(d%, 1, 0) + ZusatzBasis#
      Zusatz#(d%, 0, 0) = Zusatz#(d%, 0, 0) + ZusatzBasis#
      Zusatz#(0, 1, 0) = Zusatz#(0, 1, 0) + ZusatzBasis#
      Zusatz#(0, 0, 0) = Zusatz#(0, 0, 0) + ZusatzBasis#
      ZusatzUms#(d%, 1, 0) = ZusatzUms#(d%, 1, 0) + Zusatzumsatz#
      ZusatzUms#(d%, 0, 0) = ZusatzUms#(d%, 0, 0) + Zusatzumsatz#
      ZusatzUms#(0, 1, 0) = ZusatzUms#(0, 1, 0) + Zusatzumsatz#
    End If
  Else
    PrivUms#(d%, 1, p.Pnr) = PrivUms#(d%, 1, p.Pnr) + KdUmsatz#
    PrivUms#(d%, 0, p.Pnr) = PrivUms#(d%, 0, p.Pnr) + KdUmsatz#
    PrivUms#(0, 1, p.Pnr) = PrivUms#(0, 1, p.Pnr) + KdUmsatz#
    'alle Mitarb
    PrivUms#(d%, 1, 0) = PrivUms#(d%, 1, 0) + KdUmsatz#
    PrivUms#(d%, 0, 0) = PrivUms#(d%, 0, 0) + KdUmsatz#
    PrivUms#(0, 1, 0) = PrivUms#(0, 1, 0) + KdUmsatz#
  End If
  AnyRez = False
  ZusatzBasis# = 0
  Zusatzumsatz# = 0#
  KdUmsatz# = 0#
  
  If SonderBasis# <> 0# Then
    AnzSPrmKd!(d%, 1, p.Pnr) = AnzSPrmKd!(d%, 1, p.Pnr) + p.KuEnde
    AnzSPrmKd!(d%, 0, p.Pnr) = AnzSPrmKd!(d%, 0, p.Pnr) + p.KuEnde
    AnzSPrmKd!(0, 1, p.Pnr) = AnzSPrmKd!(0, 1, p.Pnr) + p.KuEnde
    
    AnzSPrmKd!(d%, 1, 0) = AnzSPrmKd!(d%, 1, 0) + p.KuEnde
    AnzSPrmKd!(d%, 0, 0) = AnzSPrmKd!(d%, 0, 0) + p.KuEnde
    AnzSPrmKd!(0, 1, 0) = AnzSPrmKd!(0, 1, 0) + p.KuEnde
  End If
  SonderBasis# = 0#
  
  'je Mitarb
  'If p.KuEnde < 0 Then Stop
  'p.KuEnde = Abs(p.KuEnde)
  AnzKd!(d%, 1, p.Pnr) = AnzKd!(d%, 1, p.Pnr) + p.KuEnde
  AnzKd!(d%, 0, p.Pnr) = AnzKd!(d%, 0, p.Pnr) + p.KuEnde
  AnzKd!(0, 1, p.Pnr) = AnzKd!(0, 1, p.Pnr) + p.KuEnde
  'alle Mitarb
  AnzKd!(d%, 1, 0) = AnzKd!(d%, 1, 0) + p.KuEnde
  AnzKd!(d%, 0, 0) = AnzKd!(d%, 0, 0) + p.KuEnde
  AnzKd!(0, 1, 0) = AnzKd!(0, 1, 0) + p.KuEnde
End If

'je Mitarb
'If p.RezEnde <> 0 Then
''  Stop
'  p.Laufnr = p.Laufnr
'  If p.RezEnde <> 1 Then Stop
'End If
AnzRp!(d%, 1, p.Pnr) = AnzRp!(d%, 1, p.Pnr) + p.RezEnde
AnzRp!(d%, 0, p.Pnr) = AnzRp!(d%, 0, p.Pnr) + p.RezEnde
AnzRp!(0, 1, p.Pnr) = AnzRp!(0, 1, p.Pnr) + p.RezEnde
'alle Mitarb
AnzRp!(d%, 1, 0) = AnzRp!(d%, 1, 0) + p.RezEnde
AnzRp!(d%, 0, 0) = AnzRp!(d%, 0, 0) + p.RezEnde
AnzRp!(0, 1, 0) = AnzRp!(0, 1, 0) + p.RezEnde
Return
'----------------------------------------------------------------------------------------------------------------------
PrmKdJeTag:
'je Mitarb
AnzPKd!(d%, 1, p.Pnr) = AnzPKd!(d%, 1, p.Pnr) + p.KuEnde
AnzPKd!(d%, 0, p.Pnr) = AnzPKd!(d%, 0, p.Pnr) + p.KuEnde
AnzPKd!(0, 1, p.Pnr) = AnzPKd!(0, 1, p.Pnr) + p.KuEnde
'alle Mitarb
AnzPKd!(d%, 1, 0) = AnzPKd!(d%, 1, 0) + p.KuEnde
AnzPKd!(d%, 0, 0) = AnzPKd!(d%, 0, 0) + p.KuEnde
AnzPKd!(0, 1, 0) = AnzPKd!(0, 1, 0) + p.KuEnde
Return
'----------------------------------------------------------------------------------------------------------------------
KdRpJeGeraet:
'Geräte je Mitarb
ganzKd!(g%, s%, p.Pnr) = ganzKd!(g%, s%, p.Pnr) + p.KuEnde
ganzKd!(10, 0, p.Pnr) = ganzKd!(10, 0, p.Pnr) + p.KuEnde
ganzRp!(g%, s%, p.Pnr) = ganzRp!(g%, s%, p.Pnr) + p.RezEnde
ganzRp!(10, 0, p.Pnr) = ganzRp!(10, 0, p.Pnr) + p.RezEnde
'Geräte alle Mitarb
ganzKd!(g%, s%, 0) = ganzKd!(g%, s%, 0) + p.KuEnde
ganzKd!(10, 0, 0) = ganzKd!(10, 0, 0) + p.KuEnde
ganzRp!(g%, s%, 0) = ganzRp!(g%, s%, 0) + p.RezEnde
ganzRp!(10, 0, 0) = ganzRp!(10, 0, 0) + p.RezEnde
Return
'----------------------------------------------------------------------------------------------------------------------

End Sub



Function fMinuten%(Uhrzeit%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("fMinuten")
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

fMinuten% = Int(Uhrzeit% / 100) * 60 + (Uhrzeit% Mod 100)
Call DefErrPop
End Function


Sub TagSuchen(gefunden As Boolean, satz As Long, SuchDatumC$, MaxDS As Long)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("TagSuchen")
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

Dim xLinks As Long, xRechts As Long, xMitte As Long
Dim asatz As Long
Dim SuchDatum As Date

SuchDatum = CVDatum(SuchDatumC$)

gefunden = False
xLinks = 64: xRechts = MaxDS
Do While xLinks <= xRechts
  xMitte = Int((xLinks + xRechts) / 2)
  Get #fpss, (xMitte * 64!) + 1, p
  If CVDatum(p.datum) < SuchDatum Then xLinks = xMitte + 1 Else xRechts = xMitte - 1
Loop
asatz = xMitte: If asatz < 64 Then asatz = 64
Get #fpss, (asatz * 64!) + 1, p
Do While CVDatum(p.datum) < SuchDatum And asatz <= MaxDS And Not gefunden
  Get #fpss, (asatz * 64!) + 1, p
  If CVDatum(p.datum) = SuchDatum Then gefunden = True
  If CVDatum(p.datum) < SuchDatum And Not gefunden Then asatz = asatz + 1
Loop
If asatz > MaxDS Then asatz = MaxDS
If CVDatum(p.datum) = SuchDatum Then gefunden = True
satz = asatz
Call DefErrPop
End Sub


Sub TabelleLoeschen()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("TabelleLoeschen")
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

ReDim Basis#(aTage%, 1, MAXPERSONAL%)            'PrämienBasis
ReDim AnzZusKd!(aTage%, 1, MAXPERSONAL%)          'Anzahl Zusatzverkaufkunden
ReDim AnzSPrmKd!(aTage%, 1, MAXPERSONAL%)          'Anzahl Sonderverkaufkunden
ReDim Zusatz#(aTage%, 1, MAXPERSONAL%)
ReDim Sonder#(aTage%, 1, MAXPERSONAL%)
ReDim AnzKd!(aTage%, 1, MAXPERSONAL%)            'Anzahl Kunden
ReDim AnzPKd!(aTage%, 1, MAXPERSONAL%)           'Anzahl PrämKunden
ReDim AnzRezKd!(aTage%, 1, MAXPERSONAL%)           'Anzahl RezeptKunden
ReDim PrivUms#(aTage%, 1, MAXPERSONAL%)           'Umsatz Privatkunden
ReDim ZusatzUms#(aTage%, 1, MAXPERSONAL%)           'Umsatz Zusatzverkäufe
ReDim AnzRp!(aTage%, 1, MAXPERSONAL%)            'Anzahl Rezepte
ReDim AnzBar!(aTage%, 1, MAXPERSONAL%)           'Anzahl Barverkäufe
ReDim erstKd%(aTage%, 1, MAXPERSONAL%)           'erster Kunde         '1.09
ReDim letztKd%(aTage%, 1, MAXPERSONAL%)          'letzter Kunde        '1.09
ReDim Pause1Anz%(aTage%, 1, MAXPERSONAL%)
ReDim Pause1Zeit(aTage%, 1, MAXPERSONAL%)
ReDim Pause2Anz%(aTage%, 1, MAXPERSONAL%)
ReDim Pause2Zeit(aTage%, 1, MAXPERSONAL%)

  '(Gerät/Seite/Mitarb)
ReDim ganzKd!(aGeräte%, 2, MAXPERSONAL%)           'Anzahl Kunden
ReDim ganzRp!(aGeräte%, 2, MAXPERSONAL%)           'Anzahl Rezepte
ReDim gAnzPB!(aGeräte%, 2, MAXPERSONAL%)           'Anzahl PrämienBasen
ReDim gPB#(aGeräte%, 2, MAXPERSONAL%)              'Summe PrämienBasis
ReDim gAnzSP!(aGeräte%, 2, MAXPERSONAL%)           'Anzahl SonderPrämien
ReDim gSP#(aGeräte%, 2, MAXPERSONAL%)              'SonderPrämie
ReDim gAnzZV!(aGeräte%, 2, MAXPERSONAL%)           'Anzahl ZusatzVerkäufe
ReDim gZV#(aGeräte%, 2, MAXPERSONAL%)              'ZusatzVerkauf
'(MitArb)
ReDim AnzRab!(MAXPERSONAL%)                  'Anzahl Rabatte
ReDim SumRab#(MAXPERSONAL%)                  'Summe Rabatte
ReDim Umsatz#(MAXPERSONAL%)                  'Brutto-Umsatz
       

gAnzNormTage! = 0!
Call DefErrPop
End Sub


Sub MitArbSpeichern(anzahl%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("MitArbSpeichern")
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

Dim ftmp As Long
Dim i%

ftmp = fopen("PSTATIS2.DAT", "O")
If ftmp > 0 Then
  For i% = 1 To UBound(MitArb$)
    Print #ftmp, Ansi2Oem(MitArb$(i%))
  Next i%
  Close #ftmp
End If
Call DefErrPop
End Sub

Sub HoleIniFeiertage()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("HoleIniFeiertage")
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

Dim i%, j%, ind%
Dim l&
Dim wert1$, h$, key$, Aktiv$, mond$, Jahr$
Dim OsterSonntag As Date, advent1 As Date, BetTag As Date

j% = 0
For i% = 1 To MAX_FEIERTAGE
    h$ = Space$(100)
    key$ = "Feiertag" + Format(i%, "00")
    l& = GetPrivateProfileString("Feiertage", key$, " ", h$, 101, CurDir + "\winop.ini")
    h$ = Trim$(Left$(h$, l&))
    If (Len(h$) <= 1) Then Exit For
    
    ind% = InStr(h$, ",")
    If (ind% > 0) Then
        wert1$ = RTrim$(Left$(h$, ind% - 1))
        Call OemToChar(wert1$, wert1$)
        Aktiv$ = Mid$(h$, ind% + 1)
        If (wert1$ <> "") Then
            Feiertage(j%).Name = wert1$
            Feiertage(j%).Aktiv = Aktiv$
            j% = j% + 1
        End If
    End If
Next i%
AnzFeiertage% = j%

If (AnzFeiertage% = 0) Then
    Feiertage(0).Name = "Neujahr"
    Feiertage(1).Name = "Heilige 3 Könige"
    Feiertage(2).Name = "Karfreitag"
    Feiertage(3).Name = "Ostermontag"
    Feiertage(4).Name = "Tag der Arbeit"
    Feiertage(5).Name = "Christi Himmelfahrt"
    Feiertage(6).Name = "Pfingstmontag"
    Feiertage(7).Name = "Fronleichnam"
    Feiertage(8).Name = "Maria Himmelfahrt"
    If (para.Land = "A") Then
        Feiertage(9).Name = "Nationalfeiertag"
    Else
        Feiertage(9).Name = "Tag der Deutschen Einheit"
    End If
    Feiertage(10).Name = "Reformationstag"
    Feiertage(11).Name = "Allerheiligen"
    If (para.Land = "A") Then
        Feiertage(12).Name = "Mariä Empfängnis"
    Else
        Feiertage(12).Name = "Buß- und Bettag"
    End If
    Feiertage(13).Name = "1. Weihnachtsfeiertag"
    Feiertage(14).Name = "2. Weihnachtsfeiertag"
    For i% = 0 To 14
        Feiertage(i%).Aktiv = "J"
    Next i%
    AnzFeiertage% = 15
End If

Call DefErrPop
End Sub

Function IstFeiertag(SuchTag As Date) As Boolean
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("IstFeiertag")
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
Dim ret As Boolean

ret = False

For i% = 0 To (AnzFeiertage% - 1)
  If (Feiertage(i%).Aktiv = "J") Then
    If (Feiertage(i%).KalenderTag = SuchTag) Then
      ret = True
      Exit For
    End If
  End If
Next i%

IstFeiertag = ret
Call DefErrPop
End Function

Sub HoleIniOffen()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("HoleIniOffen")
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
Dim i%, j%, ind%
Dim l&
Dim h$, key$

Dim WochenTag$(6)

WochenTag$(0) = "Montag"
WochenTag$(1) = "Dienstag"
WochenTag$(2) = "Mittwoch"
WochenTag$(3) = "Donnerstag"
WochenTag$(4) = "Freitag"
WochenTag$(5) = "Samstag"
WochenTag$(6) = "Sonntag"
    
For i% = 0 To 5
    h$ = Space$(100)
    key$ = WochenTag$(i%)
    l& = GetPrivateProfileString("PVS", key$, " ", h$, 101, CurDir + "\winop.ini")
    h$ = Trim$(Left$(h$, l&))
    If (Len(h$) <= 1) Then
        h$ = "0800,1800"
    End If
    If (Right$(h$, 1) <> ",") Then
        h$ = h$ + ","
    End If
    
    With OffenRec(i%)
        For j% = 0 To 1
            .von(j%) = 0
            .Bis(j%) = 0
        Next j%
        
        ind% = InStr(h$, ",")
        If (ind% > 0) Then
            .von(0) = Val(RTrim$(Left$(h$, ind% - 1)))
            h$ = Mid$(h$, ind% + 1)
        End If
        ind% = InStr(h$, ",")
        If (ind% > 0) Then
            .Bis(0) = Val(RTrim$(Left$(h$, ind% - 1)))
            h$ = Mid$(h$, ind% + 1)
        End If
        ind% = InStr(h$, ",")
        If (ind% > 0) Then
            .von(1) = Val(RTrim$(Left$(h$, ind% - 1)))
            h$ = Mid$(h$, ind% + 1)
        End If
        ind% = InStr(h$, ",")
        If (ind% > 0) Then
            .Bis(1) = Val(RTrim$(Left$(h$, ind% - 1)))
            h$ = Mid$(h$, ind% + 1)
        End If
    End With
Next i%


h$ = "100"
l& = GetPrivateProfileString("PBA", "PersonalStunden", "100", h$, 4, CurDir + "\winop.ini")
OrgPersonalWochenStunden% = Val(Left$(h$, l&))

h$ = "05"
l& = GetPrivateProfileString("PBA", "ToleranzOffen", "05", h$, 3, CurDir + "\winop.ini")
ToleranzOffen% = Val(Left$(h$, l&))

h$ = "05"
l& = GetPrivateProfileString("PBA", "ToleranzVergleich", "05", h$, 3, CurDir + "\winop.ini")
ToleranzVergleich% = Val(Left$(h$, l&))


Call RechneOffen(OffenRec())

Call DefErrPop
End Sub

Sub SpeicherIniOffen()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("SpeicherIniOffen")
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
Dim i%, j%, ind%
Dim l&
Dim h$, h2$, key$

Dim WochenTag$(6)

WochenTag$(0) = "Montag"
WochenTag$(1) = "Dienstag"
WochenTag$(2) = "Mittwoch"
WochenTag$(3) = "Donnerstag"
WochenTag$(4) = "Freitag"
WochenTag$(5) = "Samstag"
WochenTag$(6) = "Sonntag"
    
For i% = 0 To 5
        key$ = "Personal" + Format(i%, "00")
        h$ = PersonalFarben$(i%) + "," + PersonalInitialen$(i%)
        l& = WritePrivateProfileString("PersonalFarben", key$, h$, CurDir + "\winop.ini")
    
    key$ = WochenTag$(i%)
    
    h$ = ""
    With OffenRec(i%)
        If (.von(0) > 0) Then
            h2 = Right$("0000" + Format(.von(0), "0"), 4)
            h$ = h$ + h2$ + ","
            h2 = Right$("0000" + Format(.Bis(0), "0"), 4)
            h$ = h$ + h2$ + ","
            If (.von(1) > 0) Then
                h2 = Right$("0000" + Format(.von(1), "0"), 4)
                h$ = h$ + h2$ + ","
                h2 = Right$("0000" + Format(.Bis(1), "0"), 4)
                h$ = h$ + h2$ + ","
            End If
        End If
    End With
    
    l& = WritePrivateProfileString("PVS", key$, h$, CurDir + "\winop.ini")
Next i%

Call DefErrPop
End Sub

Sub AuslesenFlexOffen(oRec() As OffenStruct)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("AuslesenFlexOffen")
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
Dim i%, j%
Dim h$
    
For i% = 0 To 5
    With oRec(i%)
        For j% = 0 To 1
            .von(j%) = 0
            .Bis(j%) = 0
        Next j%
    End With
Next i%
With frmWinPvsOptionen.flxOptionen4
    For i% = 1 To 6
        h$ = Trim(.TextMatrix(i%, 1))
        If (h$ <> "") Then
            oRec(i% - 1).von(0) = Val(Left$(h$, 2) + Right$(h$, 2))
            h$ = Trim(.TextMatrix(i%, 2))
            If (h$ <> "") Then
                oRec(i% - 1).Bis(0) = Val(Left$(h$, 2) + Right$(h$, 2))
            End If
            
            h$ = Trim(.TextMatrix(i%, 3))
            If (h$ <> "") Then
                oRec(i% - 1).von(1) = Val(Left$(h$, 2) + Right$(h$, 2))
                h$ = Trim(.TextMatrix(i%, 4))
                If (h$ <> "") Then
                    oRec(i% - 1).Bis(1) = Val(Left$(h$, 2) + Right$(h$, 2))
                End If
            End If
        End If
    Next i%
End With

Call DefErrPop
End Sub

Sub RechneOffen(oRec() As OffenStruct)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("RechneOffen")
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
Dim i%, j%, iOffenTagMinuten%, iAbOffen%, iBisOffen%, iAbMinuten%, iBisMinuten%, iOffenMinuten%
Dim AbOffen%, BisOffen%, AbStunde%, AbMinuten%, BisStunde%, BisMinuten%

gVon% = 2359
gBis% = 0
GesamtOffen% = 0
OffenTagMinuten% = 0
For i% = 0 To 5
'    With OffenRec(i%)
    AbOffen% = 2359
    BisOffen% = 0
    With oRec(i%)
        iOffenTagMinuten% = 0
        For j% = 0 To 1
            iAbOffen% = .von(j%)
            iBisOffen% = .Bis(j%)
            If (iBisOffen% > iAbOffen%) And ((iAbOffen% > 0) Or (iBisOffen% > 0)) Then
                If (iAbOffen% < AbOffen%) Then
                    AbOffen% = iAbOffen%
                End If
                If (iBisOffen% > BisOffen%) Then
                    BisOffen% = iBisOffen%
                End If
                
                If (iAbOffen% < gVon%) Then
                    gVon% = iAbOffen%
                End If
                If (iBisOffen% > gBis%) Then
                    gBis% = iBisOffen%
                End If
                
                iAbMinuten% = (iAbOffen% \ 100) * 60 + (iAbOffen% Mod 100)
                iBisMinuten% = (iBisOffen% \ 100) * 60 + (iBisOffen% Mod 100)
                iOffenMinuten% = iBisMinuten% - iAbMinuten%
                iOffenTagMinuten% = iOffenTagMinuten% + iOffenMinuten%
            End If
        Next j%
        .iAbOffen = AbOffen%
        .iBisOffen = BisOffen%
    End With
    If (i% = 0) Then
        OffenTagMinuten% = iOffenTagMinuten%
    End If
    GesamtOffen% = GesamtOffen% + iOffenTagMinuten%
Next i%


gAbOffen% = gVon% - ToleranzOffen%
AbStunde% = gAbOffen% \ 100
AbMinuten% = gAbOffen% Mod 100
If (AbMinuten% > 59) Then
    gAbOffen% = AbStunde% * 100 + (AbMinuten% - 40)
End If

gBisOffen% = gBis% + ToleranzOffen%
BisStunde% = gBisOffen% \ 100
BisMinuten% = gBisOffen% Mod 100
If (BisMinuten% > 59) Then
    gBisOffen% = (BisStunde% + 1) * 100 + (BisMinuten% - 60)
End If


gVon% = gVon% \ 100
If ((gBis% Mod 100) = 0) Then
    gBis% = gBis% - 1
End If
gBis% = gBis% \ 100


Call DefErrPop
End Sub


Sub MitArbLaden(anzahl%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("MitArbLaden")
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
        
Dim fpwd As Long
Dim buf As String
Dim i%
Dim l As Long

For i% = 1 To MAXPERSONAL%
  iMitArb(i%) = False
Next i%
anzahl% = 0
For i% = 1 To MAXPERSONAL%
  If para.Personal(i%) <> Space(Len(para.Personal(i%))) And para.Personal(i%) <> String(Len(para.Personal(i%)), 0) Then
    anzahl% = anzahl% + 1
    MitArb$(anzahl%) = Left$(para.Personal(i%) + Space(20), 20) + " " + Right("  " + Str(i%), 2)
  End If
Next i%

fpwd = FreeFile
fpwd = fopen("PSTATIS2.DAT", "R")
If fpwd > 0 Then
  l = LOF(fpwd)
  Close fpwd
End If
If l > 0 Then
  'Update PersonalTabelle
  fpwd = fopen("PSTATIS2.DAT", "I")
  If fpwd > 0 Then
    buf = ""
    While Not EOF(fpwd)
      Input #fpwd, buf
      buf = Oem2Ansi(buf)
      If Mid(buf, 21, 1) = "*" Then
        For i% = 1 To anzahl%
          If Left$(MitArb$(i%), 20) = Left$(buf, 20) Then
            Mid$(MitArb$(i%), 21, 1) = "*"
            iMitArb(Val(Mid$(MitArb$(i%), 22, 2))) = True
            Exit For
          End If
        Next i%
      End If
    Wend
    Close fpwd
  End If
End If
Call DefErrPop
End Sub


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

'FabsErrf% = clsfabs.Cmnd("K\20", FabsRecno&)
'If (FabsErrf%) Then
'    erg% = MsgBox("Fehler " + Str$(FabsErrf%) + " beim Schliessen der Statistik-Daten", vbInformation)
'End If
    
On Error Resume Next
wpara.ExitEndSub

TaxeDB.Close

ast.CloseDatei
ass.CloseDatei
v.CloseDatei
vt.CloseDatei
arttext.CloseDatei
besorgt.CloseDatei
nnek.CloseDatei


ast.FreeClass
ass.FreeClass
arttext.FreeClass
besorgt.FreeClass
v.FreeClass
vt.FreeClass
nnek.FreeClass

sp.Close
SprmDB.Close
Set sp = Nothing
Set SprmDB = Nothing
Set ast = Nothing
Set ass = Nothing
Set taxe = Nothing
Set arttext = Nothing
Set besorgt = Nothing
Set v = Nothing
Set vt = Nothing
Set para = Nothing
Set wpara = Nothing
Set nnek = Nothing

If ApoCDBda Then
  ApoControlRec.Close
  Set ApoControlRec = Nothing
  ApoControlTRec.Close
  Set ApoControlTRec = Nothing
  ApoControlWRec.Close
  Set ApoControlWRec = Nothing
  
  ApoControlDB.Close
  Set ApoControlDB = Nothing
  ApoControlWDB.Close
  Set ApoControlWDB = Nothing
End If
Call frmAction.frmActionUnload
    
End
Call DefErrPop
End Sub

Function HoleActBenutzer%(Optional Menu As Boolean)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("HoleActBenutzer%")
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
Dim mPos1 As clsMpos

Menu = False

Set mPos1 = New clsMpos
mPos1.OpenDatei
mPos1.GetRecord (Val(para.User) + 1)
HoleActBenutzer% = mPos1.pwCode
If mPos1.m <> 0 Or mPos1.pwCode <> 0 Or mPos1.z <> 0 Then Menu = True
mPos1.CloseDatei


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

ProgrammNamen$(0) = "Personal-Verkaufstatistik"

ReDim aDetail$(15, 0, aTage%)
ReDim aInfo$(6, 5, 0)
ReDim dInfo$(28, 2, 0)

Call DefErrPop
End Sub



Sub StartAnimation(hForm As Object, Optional text$ = "Aufgabe wird bearbeitet ...")
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("StartAnimation")
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
With hForm
    .MousePointer = vbHourglass
    .aniAnimation.Open "findfile.avi"
    .lblAnimation.Caption = text$
    .picAnimationBack.Left = (.ScaleWidth - .picAnimationBack.Width) / 2
    .picAnimationBack.Top = (.ScaleHeight - .picAnimationBack.Height) / 2
'    .picAnimationBack.Left = .Left + (.ScaleWidth - .picAnimationBack.Width) / 2
'    .picAnimationBack.Top = .Top + (.ScaleHeight - .picAnimationBack.Height) / 2
    .picAnimationBack.Visible = True
    .aniAnimation.Play
    .Refresh
End With
Call DefErrPop
End Sub

Sub StopAnimation(hForm As Object)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("StopAnimation")
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
With hForm
    .aniAnimation.Stop
    .picAnimationBack.Visible = False
    .MousePointer = vbDefault
End With
Call DefErrPop
End Sub

Sub HoleIniPersonalFarben()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("HoleIniPersonalFarben")
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
Dim h$, h2$, key$

For i% = 1 To 50
    h2$ = "000000"
    h$ = Space$(100)
    key$ = "Personal" + Format(i%, "00")
    l& = GetPrivateProfileString("PersonalFarben", key$, h2$, h$, 101, CurDir + "\winop.ini")
    h$ = Left$(h$, l&)
    
    ind% = InStr(h$, ",")
    If (ind% <= 0) Then
        h$ = h$ + ","
        ind% = InStr(h$, ",")
    End If
    PersonalFarben$(i%) = Left$(h$, ind% - 1)
    PersonalInitialen(i%) = Mid$(h$, ind% + 1)
'    iFarbeInfo& = BerechneFarbWert&(h$)
Next i%

Call DefErrPop
End Sub

Sub SpeicherIniPersonalFarben()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("SpeicherIniPersonalFarben")
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
Dim key$, h$

For i% = 1 To 50
    If (PersonalFarben$(i%) <> "000000") Or (PersonalInitialen$(i%) <> "") Then
        key$ = "Personal" + Format(i%, "00")
        h$ = PersonalFarben$(i%) + "," + PersonalInitialen$(i%)
        l& = WritePrivateProfileString("PersonalFarben", key$, h$, CurDir + "\winop.ini")
    End If
Next i%

Call DefErrPop
End Sub

Sub SpeicherParameter()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("SpeicherParameter")
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
Dim i%, row%
Dim s$

Get #fpss, 2 * 64 + 1, p3

's$ = frmWinPvsOptionen!txtOptionen0(0).Text
'i% = InStr(s$, ":")
'Do While i% > 0
'  s$ = Left$(s$, i% - 1) + Mid$(s$, i% + 1)
'  i% = InStr(s$, ":")
'Loop
''TagAnf% = Val(s$)
'
'
's$ = frmWinPvsOptionen!txtOptionen0(1).Text
'i% = InStr(s$, ":")
'Do While i% > 0
'  s$ = Left$(s$, i% - 1) + Mid$(s$, i% + 1)
'  i% = InStr(s$, ":")
'Loop
''TagEnd% = Val(s$)

NORMTAGZEIT% = Val(frmWinPvsOptionen!txtOptionen0(2).text)
PauseKl% = Val(frmWinPvsOptionen!txtOptionen0(3).text)
PauseGr% = Val(frmWinPvsOptionen!txtOptionen0(4).text)

pb# = xVal(frmWinPvsOptionen!txtOptionen0(5).text)
zpb# = xVal(frmWinPvsOptionen!txtOptionen0(6).text)
spb# = xVal(frmWinPvsOptionen!txtOptionen0(7).text)

If frmWinPvsOptionen!chkOptionen(0).Value = 1 Then
  RabAbzug = True
Else
  RabAbzug = False
End If

If frmWinPvsOptionen!chkOptionen(1).Value = 1 Then
  LSAuch = True
Else
  LSAuch = False
End If
If frmWinPvsOptionen!chkOptionen(2).Value = 1 Then
  PrivRez = True
Else
  PrivRez = False
End If


If frmWinPvsOptionen!optBasis(0).Value Then
  PrämBasis$ = "I"
ElseIf frmWinPvsOptionen!optBasis(1).Value Then
  PrämBasis$ = "E"
ElseIf frmWinPvsOptionen!optBasis(2).Value Then
  PrämBasis$ = "S"
End If

p3.GndPrm = fnx(pb#) * 100#
p3.ZVkPrm = fnx(zpb#) * 100#
p3.SndPrm = fnx(spb#) * 100#
'p3.TagAnf = TagAnf%
'p3.TagEnd = TagEnd%
p3.Normtag = NORMTAGZEIT%
p3.PauseKl = PauseKl%
p3.PauseGr = PauseGr%
p3.PrmBas = PrämBasis$
If RabAbzug Then
  p3.RabAbzg = "J"
Else
  p3.RabAbzg = "N"
End If
If LSAuch Then
  p3.Liefauch = "J"
Else
  p3.Liefauch = "N"
End If
If PrivRez Then
  p3.PrivRez = "J"
Else
  p3.PrivRez = "N"
End If
Put #fpss, 2 * 64 + 1, p3

With frmWinPvsOptionen!flxOptionen1(0)
  For row% = 1 To 10
    Get #fpss, (2 + row%) * 64 + 1, p4
    s$ = Left$(Trim(.TextMatrix(row%, 0)), 1)
    If InStr("UON", s$) = 0 Then s$ = ""
    operator$(row%) = s$
    p4.ind = s$
    
    s$ = Trim(.TextMatrix(row%, 1))
    wg$(row%) = s$
    p4.wGrp = s$
    
    s$ = Trim(.TextMatrix(row%, 2))
    tLager$(row%) = s$
    p4.tLager = Left(s$ + Space(Len(p4.tLager)), Len(p4.tLager))
    
    s$ = Trim(.TextMatrix(row%, 3))
    LgCode$(row%) = s$
    p4.LgCode = Left(s$ + Space(Len(p4.LgCode)), Len(p4.LgCode))
    
    s$ = Trim(.TextMatrix(row%, 4))
    Geräte$(row%) = s$
    p4.Geräte = s$
    
    s$ = Trim(.TextMatrix(row%, 5))
    vonAVP#(row%) = Val(s$)
    p4.vonAVP = Val(s$)
    
    s$ = Trim(.TextMatrix(row%, 6))
    bisAVP#(row%) = Val(s$)
    p4.bisAVP = Val(s$)
    
    s$ = Trim(.TextMatrix(row%, 7))
    vonSP#(row%) = Val(s$)
    p4.vonSP = Val(s$)
    
    s$ = Trim(.TextMatrix(row%, 8))
    bisSP#(row%) = Val(s$)
    p4.bisSP = Val(s$)
    
    s$ = Trim(.TextMatrix(row%, 9))
    Rp$(row%) = s$
    p4.RpPfl = s$
    
    Put #fpss, (2 + row%) * 64 + 1, p4
  Next row%
End With
Call AnfangsBedingungen
Call DefErrPop

End Sub


Sub HoleIniDiagramme()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("HoleIniDiagramme")
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

Dim i%, ind%, iTyp%, iWas%
Dim l&
Dim h$, key$

For i% = 0 To 4
    h$ = Space$(100)
    key$ = "Diagramm" + Format(i% + 1, "00")
    l& = GetPrivateProfileString("PVS", key$, h$, h$, 101, CurDir + "\winop.ini")
    h$ = Trim$(Left$(h$, l&))
    
    iTyp% = 0
    If (i% = 0) Then
        iWas% = 0
    Else
        iWas% = i% - 1
    End If
    
    ind% = InStr(h$, ",")
    If (ind% > 0) Then
        iTyp% = Val(Left$(h$, ind% - 1))
        iWas% = Val(Mid$(h$, ind% + 1))
    End If
    
    DiagrammTyp%(i%) = iTyp%
    DiagrammWas%(i%) = iWas%
Next i%

h$ = Space$(100)
key$ = "Legende"
l& = GetPrivateProfileString("PVS", key$, h$, h$, 101, CurDir + "\winop.ini")
h$ = Trim$(Left$(h$, l&))

LegendenPosStr$ = h$

h$ = Space$(100)
key$ = "LeereZeilen"
l& = GetPrivateProfileString("PVS", key$, " ", h$, 101, CurDir + "\winop.ini")
h$ = Trim$(Left$(h$, l&))
If (Len(h$) < 1) Then h$ = "J"
LeereZeilen = True
If UCase(Left$(h$, 1)) = "N" Then LeereZeilen = False
Call DefErrPop
End Sub

Sub SpeicherIniLegendenPos()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("SpeicherIniLegendenPos")
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
Dim key$, h$

key$ = "Legende"
l& = WritePrivateProfileString("PVS", key$, LegendenPosStr$, CurDir + "\winop.ini")

Call DefErrPop
End Sub


Sub SpeicherIniDiagramme(pos%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("SpeicherIniDiagramme")
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
Dim key$, h$

key$ = "Diagramm" + Format(pos% + 1, "00")
h$ = Format(DiagrammTyp%(pos%), "0") + "," + Format(DiagrammWas%(pos%), "0")
l& = WritePrivateProfileString("PVS", key$, h$, CurDir + "\winop.ini")

Call DefErrPop
End Sub


Public Function WGTrennen(wgruppen$) As String
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("WGTrennen")
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
Dim x%, xbis%, xvon%
Dim xneu$, x1$, x2$, x3$, sxvon$

x% = 1
While x% <> 0
  x% = InStr(wgruppen$, " ")
  If x% <> 0 Then wgruppen$ = Left$(wgruppen$, x% - 1) + Mid$(wgruppen$, x% + 1)
Wend
x1$ = wgruppen$ + ","
x2$ = "": xneu$ = ""
While Len(x1$) <> 0
  x3$ = Left$(x1$, 1): x1$ = Mid$(x1$, 2)
  If InStr("0123456789", x3$) = 0 Then
    If x3$ = "-" Then
      If Len(x2$) = 1 Then x2$ = x2$ + "0"
      sxvon$ = x2$: x2$ = ""
    Else
      If sxvon$ <> "" Then
        If Len(x2$) = 1 Then x2$ = x2$ + "9"
        xbis% = Val(x2$): If xbis% = 0 Then xbis% = 99
        For xvon% = Val(sxvon$) To xbis%
          xneu$ = xneu$ + "," + Right$(Str$(xvon%), 2)
          If Len(xneu$) > 252 Then xvon% = xbis%
        Next xvon%
      Else
        If Len(x2$) = 1 Then
          xvon% = Val(x2$) * 10: xbis% = xvon% + 9
          For xvon% = xvon% To xbis%
            xneu$ = xneu$ + "," + Right$(Str$(xvon%), 2)
          Next xvon%
        Else
          xneu$ = xneu$ + "," + x2$
        End If
      End If
      sxvon$ = "": x2$ = ""
    End If
  Else
    x2$ = x2$ + x3$
  End If
Wend
WGTrennen = Mid$(xneu$, 2)

Call DefErrPop
End Function


