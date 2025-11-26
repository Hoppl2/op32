VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsWinRezK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Const INI_SECTION = "Rezeptkontrolle"
Const INFO_SECTION = "Infobereich Rezeptkontrolle"

Const NOCTU_PREIS = 2.5

Const REZEPTART_KASSE = 1
Const REZEPTART_ZUZAFREI = 2
Const REZEPTART_BTM = 3
Const REZEPTART_BTM_ZUZAFREI = 4
Const REZEPTART_PRIVAT = 5
Const REZEPTART_GRÜN = 6
Const REZEPTART_IVF = 7
Const REZEPTART_SPRECHSTUNDENBEDARF = 8

Const MAX_VERK_ZEILEN% = 20

Const REZ_BLINK = &H1
Const REZ_MULTI = &H2
Const REZ_WG_4 = &H4
Const REZ_K_STOP = &H8
Const REZ_PREISEINGABE = &H10
Const REZ_FAKTOR = &H20
Const REZ_MULTI_ZEILE = &H40
Const REZ_UNTAXIERT = &H80

Const AV_DATEI_LEN% = 42

Const EINZEL_BEDINGUNG% = 0
Const ODER_BEDINGUNG% = 1
Const UND_BEDINGUNG% = 2


Const W_GESAMTJAHR = 0
Const W_GESAMT = 1
Const W_FAM = 2
Const W_IMPFÄHIG = 3
Const W_IMPIST = 4
Const W_ABZUG = 5
Const W_SALDO = 6
Const W_REZANZ = 7

Const PRIVATPREIS_STAMMVK = 0
Const PRIVATPREIS_AMPV = 1

Const SONDERPZN_WVO = "06460487"


Private Type REZEPT_ARTIKEL_STRUCT
    pzn As String * 8
    text As String * 36
    Preis As Double
    flag As Byte
    Imp123 As Byte
    Fam As Byte
    AutIdem As Byte
    AutIdemKreuz As Byte
    N0 As Byte
    MagIndex As Byte
    Zuz As Single
    appli As Byte
    Faktor As Integer
    zusatz(1) As String * 50 ' 36
    ScreenPzn As String * 10
    NichtInTaxe As Byte
    IstWg4 As Byte
    ImpErspart As Double
    PlusMehrKosten As Byte
    ZuzahlungsErlass As Byte
    Warenzeichen As Byte
    MietDauer As Integer
    HerstRabattPrivat130Brutto As Single
    VdbPauschale As Byte
    TkkPznDruck As Byte
    VebNr As Long
    PauschaleNr As Long
    Fortsetzung As Byte
    HmAbrechnungsKz As String * 2
    LEGS As String * 7
End Type
    
Private Type FORMULAR_STRUCT
    LeftX As Long
    LeftY As Long
    RightX As Long
    RightY As Long
    Fett As Integer
    Typ As String * 2
    typ2 As Boolean
End Type

'Private Type VERKAUF_STRUCT
'    pzn As String * 7
'    filler1 As String * 1
'    text As String * 36
'    filler2 As String * 1
'    Preis As String * 9
'    filler3 As String * 1
'    flag As Byte
'    Imp123 As Byte
'    Zuz As Single
'    Appli As Byte
'End Type


Private Type REZ_DRUCK_STRUCT
    RezNr As String * 14
    VerkDatum As Long
    DruckDatum As Long
    ErstSatz As Long
    NextSatz As Long
    PrevSatz As Long
    AnzZeilen As Integer
    RezGebAnz As Integer
    VerkZeit As Integer
    RezGebSumme As Single
    RezSumme As Single
    VerkUser As String * 3
    VerkLiRe As Byte
End Type

Private Type REZ_ZEILEN_STRUCT
    RezNr As String * 14
    NextSatz As Long
    PrevSatz As Long
    pzn As String * 8
    text As String * 37
    Preis As String * 10
    flag As Byte
End Type

Private Type REZKONTR2_STRUCT
    ean As String * 13
    dummy As String * 1
End Type
Private Type REZKONTR_STRUCT
    fertig As String * 1
    Preis(9) As Double
    RezEan(9) As REZKONTR2_STRUCT
End Type

Private Type IMPFUNG_STRUCT
    pzn As String * 8
    text As String * 28
End Type

Private Type ABGABE_KOSTEN_STRUCT
    Name As String * 12
    kosten As Double
    ind As Integer
End Type
    


Dim ParentForm As Form
Dim ParentToolbar As clsToolbar
Dim ParentInfo As clsInfoBereich
Dim ParentBereich As clsOpBereiche

Dim GebFrei%
Dim ImpfRez%
Dim PrivatRezept%
Dim BtmRezept%
Dim IvfRezept%
Dim Selbsterklaerung%, SelbstErklaerungHidden$, SelbstErklaerungTabstop$
Dim SchutzMasken%
Dim SchonDa%

Dim ActStueck%, ActMenge%, ActMwst%, ActVmRabatt%
Dim ActMietTage%, ActMietWochen%, ActMietMonate%, ActMietDauer%, ActMietZeit%
Dim ActPatientenAlterMonate%, ActPatientenAlterJahre%
Dim ActWert#, ActAep#, ActAvp#, ActVmPreis#, Act20Preis#
Dim ActWarenzeichen$


Dim RezSumme#, RezSummeDazu#, RezGebSumme#, LieferEngPass#
Dim RezGebRechnen%

Dim VerkUser$
Dim VerkLiRe As Byte

Dim RezKundenInfo$(5)

Dim RezMagFeld&(MAX_VERK_ZEILEN)
Dim RezMagAnzMag%(MAX_VERK_ZEILEN)

Dim RezArtikel(MAX_VERK_ZEILEN) As REZEPT_ARTIKEL_STRUCT

Dim HmNummer$(MAX_VERK_ZEILEN)
Dim HmFaktor%(MAX_VERK_ZEILEN)
Dim HmStückPreis#(MAX_VERK_ZEILEN)
    
'Dim VerkStruct(MAX_VERK_ZEILEN) As VERKAUF_STRUCT
Dim RezDruckSatz As REZ_DRUCK_STRUCT
Dim RezZeilenSatz As REZ_ZEILEN_STRUCT
Dim RezKontrSatz As REZKONTR_STRUCT

Dim RezeptBoxen(60) As FORMULAR_STRUCT
Dim BoxenWert$(60)  '36
Dim MaxPruefBox%


Dim BarVerkauf$(2)
Dim RezAddGesamtBrutto$(3), sBrutto$

Dim VerkPtr&, OrgVerkPtr&
Dim lPrivEnde&

Dim bmVerkPtr As Variant
Dim bmOrgVerkPtr As Variant
Dim bmPrivEnde As Variant

Dim lRezNr&

Dim PersonalNr%, MarsFreigabePersonalNr%


Dim BeepFlag%
Dim DruckDatum$
Dim DruckZeit%
Dim VerkDatum$, OrgVerkDatum$
Dim VerkZeit%

Dim AnzRezeptArtikel%

Dim ZusatzKomponentenNr$, ZusatzKomponentenBasisNr$

Dim AktTabIndex%

Dim AvLaender$()
'Dim AvKassen$()
Dim AvVerordnungen$()

Dim SonderPzn$()
Dim HmNummern$()
Dim BtmSonderPzn$()


Dim AbgabeKosten() As ABGABE_KOSTEN_STRUCT
Dim F7Kassen() As ABGABE_KOSTEN_STRUCT

Dim Impfungen() As IMPFUNG_STRUCT

Dim VmDebugAktiv%

Dim Operand$(2)

Dim ActBundeslandInd%, ActKasseInd%, ActVerordnungInd%

Dim ActKkasse$
Dim ActAVWGik&

Dim ActKiste%

Dim IstTm290%, NeuTm290%, IstTmKombi%, IstTmU950%

Dim KeinPreisDa%

Dim buf$
Dim APTR%, ATXT%, TPTR%, TTXT%, ZPTR%, ZPLT%, UPTR%, UDAT%

Dim RoteKassenDa%, IstRoteKasse%

Dim BereitsGedruckt%

Dim DosDruckerLpt%

Dim RezeptHolenZuz#

Dim AvpRezeptNr$
Dim AvpLaufNr&

Dim TransaktionsID&
Dim rzLieferId$

Dim AMverfügbarkeit$
Dim AMverfügbarkeitsPzn%
Dim WVORabattArtikel$

Dim SollTmName$(MAX_VERK_ZEILEN)
Dim SollTmMenge#(MAX_VERK_ZEILEN)
Dim SollTmId&(MAX_VERK_ZEILEN)

Dim BenutzerNr%

Dim AvAbrechnungsNr$(2)
Dim AvHmPositionsNr$(2)

Dim ActPreis#(2), ActPreis20#(2)
Dim ActAVrabatt%(2)

Dim IkSpezialName$

Dim BtmRezeptAlt%

Private Type ArtlistType
    pzn         As Long
    kurz        As String * 28
    Preis       As String * 10
    kz1         As String * 1
    kz2         As String * 2
    rabKZ       As String * 1
    ImportGruppeok As String * 1
    FüllMenge   As String * 10
    GhVerfügbar As Byte
    GhVerfNr    As Integer
    AusgangsArtikel As Byte
    RoteLinie   As Byte
    Sort        As String * 50
    ArtStatus   As String * 1
End Type
Dim ArtL As ArtlistType
Dim ArtLRec() As ArtlistType

Private Type MatchcodeType
  Key         As String * 50
  satz        As Long
  fertig      As Boolean
  bmk         As Variant
  RabattAutidem As Boolean
  AltPG       As Boolean
  KdStatus    As Byte
  GhVerfNr    As Integer
  AusgangsArtikel As Boolean
  RoteLinie    As Boolean
  text(11)     As Variant
End Type
Dim m() As MatchcodeType

Private Type TaxeType
  vc          As String * 1
  herst_nr    As String * 3
  pzn         As String * 8
  meng        As String * 7
  Meh         As String * 2
  kurz        As String * 28
  Name        As String * 77
  aep         As Double
  AVP         As Double
  gc          As String * 2
  SuchtG      As Boolean
  mw          As String * 1
  kas         As String * 3
  lic         As String * 1
  abl         As String * 1
  abgschl     As String * 1
  Generika    As String * 1
  lac         As String * 1
  wg          As String * 2
  appli       As Boolean
  
  FB          As Double
  Hilfsmittel As Boolean
  Verbandmittel As Boolean
  Import      As String * 1
  kasz1       As Byte           '&H80 ... Hilfsmittelkz.  &H40 .. Verbandmittelkz.   &H8 ... Import
  Negativliste As Boolean
  zuza        As Integer
  herst       As String * 5
  Preislinie  As Double
  PreisGrün   As Boolean
End Type

Private Type aMatchType
  asatz       As Long
  tsatz       As Variant
  ssatz       As Long
'  Artikel     As StammdatenType
  taxesatz    As TaxeType
  'TaxeLangText As String * 80
'  Statistik   As StatistikType
  Streichung  As Boolean
  tname       As String * 77
  tmenge      As String * 7
  tmeh        As String * 2
  SortierName As String * 48
  pzn         As Long
  StreichungohneLager As Boolean
  Messwert    As Boolean
  RSAiKreuz   As Boolean
    Artikel_PZN         As String * 8
  Artikel_zuza  As Long
  Artikel_vc    As String * 1
    Artikel_meng        As String * 5
    Artikel_meh         As String * 2
    Artikel_herst       As String * 5
End Type
Dim amatch() As aMatchType

Const LISTE_IMPORT = 1
Const LISTE_AUTIDEM = 2

Public coRot As Long
Public coKnallrot As Long
Public coGelb As Long
Public coWawiGelb As Long
Public coNewlineGridKopf As Long
Public coNewlineGridHighlight As Long
Public coNewlineMygridHighlight As Long
Public coNewlineGridBack As Long
Public coNewlineGridInfoBack As Long
Public coNewlinePicInfoBack As Long
Public coNewlinePicKasseInaktiv As Long
Public coNewlineStammKundeAktiv As Long
Public coNewlineStammKundeInaktiv As Long
Public coNewlineSattesDunkelgrau As Long
Public coNewlineKassenRez As Long
Public coNewlineHinweis As Long
Public coAusgangsartikel As Long
Public coPreisGrün As Long
Public coKennzeichenGelb As Long


Private Const DefErrModul = "WINREZK.CLS"

Public Property Let SetVerkPtr(NewValue&)
VerkPtr& = NewValue&
End Property
Public Property Get GetPrivEnde&()
GetPrivEnde& = lPrivEnde&
End Property

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
    h = "1717"
'    Call ParentToolbar.InitToolbar(ParentForm, INI_DATEI, INI_SECTION, "1717")  'weil in DLL gelöscht !
    If (para.Newline) Then
        Call ParentToolbar.InitToolbar(ParentForm, App.EXEName, INI_SECTION, h)
    Else
        Call ParentToolbar.InitToolbar(ParentForm, INI_DATEI, INI_SECTION, h)
    End If
    
    .cmdToolbar(0).ToolTipText = "ESC Zurück: Zurückschalten auf vorige Bildschirmmaske"
    .cmdToolbar(1).ToolTipText = "F2 Erfassen zusätzlicher Artikel"
    .cmdToolbar(2).ToolTipText = "F3"
    .cmdToolbar(3).ToolTipText = "F4"
    .cmdToolbar(4).ToolTipText = "F5 Entfernen eines Artikels"
    .cmdToolbar(5).ToolTipText = "F6 Drucken des Rezepts"
    .cmdToolbar(6).ToolTipText = "F7 Automatische Rezeptauswahl"
    .cmdToolbar(7).ToolTipText = "F8"
    .cmdToolbar(8).ToolTipText = "F9 Abmelden"
    .cmdToolbar(9).ToolTipText = "shift+F2 Impfstoff-Rezept"
    .cmdToolbar(10).ToolTipText = "shift+F3 Barverkäufe"
    .cmdToolbar(11).ToolTipText = "shift+F4 A+V Parameter"
    .cmdToolbar(12).ToolTipText = "shift+F5 Durchgriff auf Statistik-Anzeige"
    .cmdToolbar(13).ToolTipText = "shift+F6 Neues Rezept - zuz.frei"
    .cmdToolbar(14).ToolTipText = "shift+F7 Neues Rezept - zuz.pflichtig"
    .cmdToolbar(15).ToolTipText = "shift+F8 Indiv. Rezeptur"
    .cmdToolbar(16).ToolTipText = "shift+F9 A+V Debug"
    'cmdToolbar(19).ToolTipText = "Programm beenden"
    
    Call ParentInfo.InitInfoBereich(.flxInfo(0), INI_DATEI, INFO_SECTION, 1, 2)
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
    
    .flxarbeit(0).Visible = False
    
    Call .ShowToolbar
End With

bmVerkPtr = Null

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
Dim i%, row%, col%, iRow%, iCol%, KommWeg%, iDummy%, sF5Enabled%, sF8Enabled%
Dim dDummy#
Dim pzn$, ch$, ActKontrollen$, actzuordnung$, h$

If (ParentForm.Name <> frmAction.Name) Then Call DefErrPop: Exit Sub

EingabeStr$ = ""

sF5Enabled% = False
sF8Enabled% = False

For i% = 1 To 3
    ParentForm.mnuBearbeitenZusatz(i%).Enabled = False
Next i%
ParentForm.mnuBearbeitenZusatz(7).Enabled = False

With ParentForm.flxInfo(0)
    For i% = 0 To (.Rows - 1)
        .TextMatrix(i%, 0) = ""
    Next i%
End With

If (AktTabIndex% = 54) Then
    Call ZeigeInfoBereich("", False)
ElseIf (AktTabIndex% = 31) Then
    Call ZeigeInfoBereich("", False)
ElseIf (AktTabIndex% >= 3) Then
    row% = (AktTabIndex% - 3) \ 4
    col% = (AktTabIndex - 3) Mod 4
    
    pzn$ = RezArtikel(row%).pzn
    If (Left(pzn, 1) = "-") Then
        pzn = Mid(pzn, 2) + "0"
        pzn = MakePzn(pzn)
    End If
    If (Val(pzn$) = 0) Then
        If (para.Newline) Then
            With ParentForm.flxInfo(0)
                For i = 0 To (.Rows - 1)
                    .row = i
                    .col = 0
                    If (.TextMatrix(i, 0) = "") Then
                        .CellBackColor = vbWhite
                    Else
                        .CellFontBold = False
                        .CellForeColor = vbWhite
                        .CellBackColor = RGB(201, 123, 58)
                    End If
                Next i
            End With
        End If
        Call DefErrPop: Exit Sub
    End If
    
'    If (ArtikelDbOk) Then
'        SQLStr$ = "SELECT * FROM ARTIKEL WHERE PZN = " + pzn
'    '                    SQLStr = SQLStr + " AND LagerKz<>0"
'        FabsErrf = Artikel.OpenRecordset(ArtikelAdoRec, SQLStr, 0)
'    Else
'        FabsErrf% = ass.IndexSearch(0, pzn, FabsRecno&)
'    End If
    FabsErrf = 0
    If (FabsErrf% = 0) Then sF5Enabled% = True
    
    Call ZeigeInfoBereich(pzn$, True)
    For i% = 0 To 1
        ParentForm.flxInfo(0).TextMatrix(i%, 0) = RezArtikel(row%).zusatz(i%)
    Next i%
    
    sF8Enabled% = True
    
    ParentForm.mnuBearbeitenZusatz(0).Enabled = True
'    If (RezArtikel(row%).AutIdem) Then ParentForm.mnuBearbeitenZusatz(1).Enabled = True
    ParentForm.mnuBearbeitenZusatz(1).Enabled = True
    If (RezArtikel(row%).Imp123) Then ParentForm.mnuBearbeitenZusatz(2).Enabled = True
    ParentForm.mnuBearbeitenZusatz(3).Enabled = True
    ParentForm.mnuBearbeitenZusatz(6).Enabled = True
    ParentForm.mnuBearbeitenZusatz(7).Enabled = True
    ParentForm.mnuBearbeitenZusatz(8).Enabled = True
Else
    col% = AktTabIndex% + 4
    
    If (col% = 5) Then
        If (RezGebRechnen% = False) Then
            ParentForm.flxInfo(0).TextMatrix(0, 0) = "Zuzahlung fixiert !"
        End If
    ElseIf (col% = 6) Then
        For i% = 0 To UBound(RezAddGesamtBrutto$)
            With ParentForm.flxInfo(0)
                If (i% < .Rows) Then .TextMatrix(i%, 0) = RezAddGesamtBrutto$(i%)
            End With
        Next i%
    End If
    
    Call ZeigeInfoBereich("", False)
End If

If (Trim(ParentForm.flxInfo(0).TextMatrix(0, 0)) = "") Then
    Call PaintAvParameter
End If

If (Chef = False) Then
    sF5Enabled = False
    sF8Enabled = False
    For i = 0 To 8
        ParentForm.mnuBearbeitenZusatz(i).Enabled = False
    Next i
End If

With ParentForm
    .mnuBearbeitenInd(MENU_SF5).Enabled = sF5Enabled%
    .mnuBearbeitenInd(MENU_SF8).Enabled = sF8Enabled%
    .cmdToolbar(MENU_SF5).Enabled = sF5Enabled%
    .cmdToolbar(MENU_SF8).Enabled = sF8Enabled%
    Call .ShowToolbar
End With

If (para.Newline) Then
    With ParentForm.flxInfo(0)
        For i = 0 To (.Rows - 1)
            .row = i
            .col = 0
            If (.TextMatrix(i, 0) = "") Then
                .CellBackColor = vbWhite
            Else
                .CellFontBold = False
                .CellForeColor = vbWhite
                .CellBackColor = RGB(201, 123, 58)
            End If
        Next i
    End With
End If

Call DefErrPop
End Sub

Sub ZeigeInfoBereich(pzn$, AuchWerte%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("ZeigeInfoBereich")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
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
Dim iRow%, iCol%
    
With ParentForm.flxInfo(0)
    iRow% = .row
    iCol% = .col

    Call ParentInfo.ZeigeInfoBereich(pzn$, AuchWerte%)

    If (iRow% >= .Rows) Then iRow% = .Rows - 1
    .row = iRow%
    .col = iCol%
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
Dim sp&, x1&, y1&
Dim x#
            
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

With ParentForm.picRezept
    .Width = ParentForm.flxarbeit(0).Width
    .Height = .Width * 10.3 / 15
    If (.Height > ParentForm.flxarbeit(0).Height) Then
        .Height = ParentForm.flxarbeit(0).Height
        .Width = .Height * 15 / 10.3
    End If
    .Left = (ParentForm.picBack(0).Width - .Width) / 2
    .Top = ParentForm.flxarbeit(0).Top
    
    Call CalcFormularBoxen
End With

Call DefErrPop
End Sub

Sub CalcFormularBoxen()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("CalcFormularBoxen")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
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
Dim he#, he1#, wi#

Call CalcFormularBox(0, "L1", 125, 8, 24, 6, True)       'Instituts-Kz.

Call CalcFormularBox(1, "R4", 96, 17, 19, 8, False, True)  'Zuzahlung
Call CalcFormularBox(2, "R7", 115, 17, 34, 8, True, True)  'Tax-Betrag

If (AnzRezeptArtikel% <= 3) Then

    For i% = 0 To 6
        ind% = 3 + (i% * 4)
        he1# = 8
        
        If (i% > 2) Then wi# = 1000#
        
        Call CalcFormularBox(ind%, "L0", 13 + wi#, 60 + (i% * 9), 80, 6.3, True) 'Name
        Call CalcFormularBox(ind% + 1, "L1", 96 + wi#, 27 + he#, 32, he1#, True)  'PZN
        Call CalcFormularBox(ind% + 2, "Z0", 128 + wi#, 27 + he#, 6, he1#, True)  'Faktor
        Call CalcFormularBox(ind% + 3, "R0", 134 + wi#, 27 + he#, 15, he1#, True) 'Preis
        
        he# = he# + he1# + 1
    Next i%
    
    MaxPruefBox% = 14

Else
    
    For i% = 0 To 6
        ind% = 3 + (i% * 4)
        
        he1# = 7
        If (i% > 0) Then he1# = he1# + 1
        
        Call CalcFormularBox(ind%, "L0", 13, 54 + (i% * 5), 80, 5, True)  'Name
        
        Call CalcFormularBox(ind% + 1, "L1", 96, 28 + he#, 32, he1#, True)    'PZN
        Call CalcFormularBox(ind% + 2, "Z0", 128, 28 + he#, 6, he1#, True)    'Faktor
        Call CalcFormularBox(ind% + 3, "R0", 134, 28 + he#, 15, he1#, True)   'Preis
        
        he# = he# + he1# + 1
    Next i%
    
    MaxPruefBox% = 30

End If

'Call CalcFormularBox(31, "L1", 134, 83, 15, 5, True)         'Abgabe-Datum
Call CalcFormularBox(31, "L1", 45, 83, 15, 5, True)         'Abgabe-Datum
'MaxPruefBox% = MaxPruefBox% + 1
 
    
'nur Anzeigen
Call CalcFormularBox(32, "", 1.5, 9, 8, 8, False)         'Geb-Frei
Call CalcFormularBox(33, "", 1.5, 19, 8, 8, False)         'Geb-Pflichtig
 
Call CalcFormularBox(34, "L0", 13, 4, 80, 5, True)      'Kassen
Call CalcFormularBox(35, "L0", 13, 9, 80, 6, True)
Call CalcFormularBox(36, "L0", 13, 15, 80, 20, True)       'Versicherter

Call CalcFormularBox(37, "", 13, 35, 80, 9, True)
Call CalcFormularBox(38, "", 13, 44, 80, 9, True)

Call CalcFormularBox(39, "", 35, 41, 0, 3, True)       'senkr. Strich
Call CalcFormularBox(40, "", 69, 41, 0, 3, True)       'senkr. Strich

Call CalcFormularBox(41, "", 43, 50, 0, 3, True)       'senkr. Strich
Call CalcFormularBox(42, "", 66, 50, 0, 3, True)       'senkr. Strich

Call CalcFormularBox(43, "Z0", 1.5, 60, 4.5, 6.3, True)         '1.Aut Idem
Call CalcFormularBox(44, "Z0", 1.5, 69, 4.5, 6.3, True)         '2.Aut Idem
Call CalcFormularBox(45, "Z0", 1.5, 78, 4.5, 6.3, True)         '3.Aut Idem


Call CalcFormularBox(46, "", 105.5, 23, 0, 2, False, True)     'senkr. Strich

Call CalcFormularBox(47, "", 125, 24, 0, 1, True, True)      'senkr. Strich
Call CalcFormularBox(48, "", 139, 23, 0, 2, True, True)      'senkr. Strich

Call CalcFormularBox(49, "", 0, 89, 150, 0, True)      'Strich

Call CalcFormularBox(50, "Z0", 6.7, 60, 4.5, 6.3, True)         '1.Import
Call CalcFormularBox(51, "Z0", 6.7, 69, 4.5, 6.3, True)         '2.Import
If (Selbsterklaerung) And (SchutzMasken = 0) Then
    Call CalcFormularBox(52, "L1", 72, 45, 15, 5, True)         'Abgabe-Datum
Else
    Call CalcFormularBox(52, "Z0", 6.7, 78, 4.5, 6.3, True)         '3.Import
End If

Call CalcFormularBox(53, "Z0", 14, 36, 20, 7, True)
Call CalcFormularBox(54, "Z0", 36, 36, 32, 7, True)
Call CalcFormularBox(55, "Z0", 70, 36, 22, 7, True)

Call CalcFormularBox(56, "Z0", 14, 45, 28, 7, True)
Call CalcFormularBox(57, "Z0", 44, 45, 21, 7, True)
Call CalcFormularBox(58, "Z0", 67, 45, 25, 7, True)

Call CalcFormularBox(59, "Z0", 14, 36, 20, 7, True)               'IK-Nr

Call CalcFormularBox(60, "R0", 13, 4, 80, 5, True)               'Kkassen-Name

With frmAction.txtMarsRezeptAnmerkungen
    .Width = RezeptBoxen(2).RightX - RezeptBoxen(2).LeftX
    .Height = RezeptBoxen(35).RightY - RezeptBoxen(34).LeftY
    .Left = RezeptBoxen(34).RightX - .Width
    .Top = RezeptBoxen(34).LeftY
    .ForeColor = vbWhite
    .BackColor = vbRed
    .FontSize = frmAction.Font.Size
    .MaxLength = 250
End With
'With frmAction.lblMarsAnzahlRezepte
'    .Width = RezeptBoxen(1).RightX - RezeptBoxen(1).LeftX
'    .Height = RezeptBoxen(1).RightY - RezeptBoxen(1).LeftY
'    .Left = RezeptBoxen(1).LeftX
'    .Top = RezeptBoxen(0).LeftY
'    .BackColor = vbGreen
'    .FontSize = frmAction.Font.Size + 4
''    .BorderStyle = 1
'End With

Call DefErrPop
End Sub

Sub CalcFormularBox(pos%, Typ$, x#, y#, wi#, he#, typ2 As Boolean, Optional Fett% = False)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("CalcFormularBox")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
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
Dim x1&, x2&, y1&, Y2&

With ParentForm.picRezept
    x1& = x# * .ScaleWidth / 150&
    y1& = y# * .ScaleHeight / 105&
    
    x2& = (x# + wi#) * .ScaleWidth / 150&
    Y2& = (y# + he#) * .ScaleHeight / 105&
End With

With RezeptBoxen(pos%)
    .LeftX = x1&
    .LeftY = y1&
    .RightX = x2&
    .RightY = Y2&
    .Fett = Fett%
    .Typ = Typ$
    .typ2 = typ2
End With

Call DefErrPop
End Sub

Sub PaintRezept(Optional DirektNachAufruf% = 0)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("PaintRezept")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
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
Dim i%, j%, OrgFontSize%, BlinkFlag%, MinX%, IstRot%, hImp123%, MitVerfügbarkeit%, MaxWi%, ind%
Dim x1&, x2&, y1&, Y2&, lDrawWidth&, CurrX&, CurrY&, lColor&
Dim h$, h2$, SQLStr$
Dim AnzZeilen%

With ParentForm.txtMarsRezeptAnmerkungen
    .Visible = False
    .text = Trim(MarsRezeptAnmerkungen)
    If (.text <> "") Then
        .Visible = True
    End If
End With

Call CalcFormularBoxen

IstRoteKasse% = False
Call PaintAvParameter

If (ParentForm.Name = frmAction.Name) Then
    Call PaintAnzuzeigendeKkassen
    
    With ParentForm
        If (ImpfRez%) Then
            .mnuBearbeitenInd(MENU_SF3).Enabled = False
            .mnuBearbeitenInd(MENU_SF4).Enabled = False
            .mnuBearbeitenInd(MENU_SF5).Enabled = False
            .mnuBearbeitenInd(MENU_SF6).Enabled = False
            .mnuBearbeitenInd(MENU_SF7).Enabled = False
            .mnuBearbeitenInd(MENU_SF8).Enabled = False
            .mnuBearbeitenInd(MENU_SF9).Enabled = False
            .cmdToolbar(MENU_SF3).Enabled = False
            .cmdToolbar(MENU_SF4).Enabled = False
            .cmdToolbar(MENU_SF5).Enabled = False
            .cmdToolbar(MENU_SF6).Enabled = False
            .cmdToolbar(MENU_SF7).Enabled = False
            .cmdToolbar(MENU_SF8).Enabled = False
            .cmdToolbar(MENU_SF9).Enabled = False
        End If
    End With
End If


With ParentForm.picRezept
    .DrawWidth = 1
    
    If (ParentForm.Name = frmAction.Name) Then
        If (para.Newline) Then
            Call wpara.ControlBorderless(ParentForm.picRezept, 2, 2)

            With ParentForm.flxInfo(0)
                x1 = .Left + .Width
            End With
            MinX = x1 - .Width
            
            MinX% = ParentForm.flxKkassen.Left + ParentForm.flxKkassen.Width + 1200
            If ((MinX% + .Width) > x1) Then
                MinX% = x1 - .Width
            Else
                .Width = x1 - MinX
                Call CalcFormularBoxen
                Call wpara.ControlBorderless(ParentForm.picRezept, 2, 2)
            End If
        Else
            If (ParentForm.flxKkassen.Visible) Then
                MinX% = ParentForm.flxKkassen.Left + ParentForm.flxKkassen.Width
            Else
                MinX% = ParentForm.txtRezeptNr.Left + ParentForm.txtRezeptNr.Width
            End If
            MinX% = MinX% + 150
        
            If ((MinX% + .Width) > ParentForm.ScaleWidth) Then
                MinX% = ParentForm.ScaleWidth - .Width
            End If
        End If
        
    Else
        MinX% = frmRezEinzeln.picRezept.Left
    End If
    .Left = MinX%
    
    .Visible = True
    
    .Cls
    
    MitVerfügbarkeit = 0
    If (PrivatRezept%) Then
        lColor& = RezeptFarben&(1)
    ElseIf (BtmRezept%) Then
        lColor& = RezeptFarben&(2)
        MitVerfügbarkeit = True
    ElseIf (SonderBelegRezept%) Then
        lColor& = RezeptFarben&(3)
    ElseIf (Selbsterklaerung) Then
        lColor& = RezeptFarben&(0)
    Else
        lColor& = RezeptFarben&(0)
        MitVerfügbarkeit = True
    End If
    
    
    ParentForm.picRezept.Line (0, 0)-(.ScaleWidth, .ScaleHeight), lColor&, BF
    If (ParentForm.Name = frmAction.Name) And (para.Newline) Then
'    If (para.Newline) Then
        ParentForm.picBack(0).ForeColor = RGB(177, 177, 177)
        ParentForm.picBack(0).FillStyle = vbSolid
        ParentForm.picBack(0).FillColor = lColor
'        RoundRect ParentForm.picBack(0).hdc, (.Left) / Screen.TwipsPerPixelX - 5, (.Top / Screen.TwipsPerPixelY) - 5, (ParentForm.picBack(0).Width - wpara.NlFlexBackY) / Screen.TwipsPerPixelX, (.Top + .Height) / Screen.TwipsPerPixelX + 5, 20, 20
        RoundRect ParentForm.picBack(0).hdc, (.Left) / Screen.TwipsPerPixelX - 5, (.Top / Screen.TwipsPerPixelY) - 5, (.Left + .Width) / Screen.TwipsPerPixelX + 5, (.Top + .Height) / Screen.TwipsPerPixelX + 5, 20, 20
        ParentForm.picBack(0).Refresh
'        With .Container
'            .ForeColor = RGB(177, 177, 177)
'            .FillStyle = vbSolid
'            .FillColor = lColor
'        End With
'        RoundRect .Container.hdc, (.Left) / Screen.TwipsPerPixelX - 5, (.Top / Screen.TwipsPerPixelY) - 5, (.Container.Width - wpara.NlFlexBackY) / Screen.TwipsPerPixelX, (.Top + .Height) / Screen.TwipsPerPixelX + 5, 20, 20
'        .Container.Refresh
    End If
    
    Call PaintRezeptDaten(lColor&)
    
    OrgFontSize% = .Font.Size
    
    If (ParentForm.Name = frmAction.Name) Then
        h$ = "4 - Nichtverfügbarkeit eines Rabatt- oder Importartikels"
        MaxWi% = RezeptBoxen(6).RightX - RezeptBoxen(4).LeftX - 150 ' wpara.FrmScrollHeight
        For i% = OrgFontSize To 7 Step -1
            .Font.Size = i%
            If (.TextWidth(h$) < MaxWi%) Then
                For j% = 0 To 2
                    frmAction.cboVerfügbarkeit(j%).FontSize = i%
                Next j%
                Exit For
            End If
        Next i%
    End If
    
    .Font.Size = OrgFontSize%
    
    For i% = 0 To 52    'UBound(RezeptBoxen)
'        If (PrivatRezept% = 0) Or ((i% <> 12) And (i% <> 15) And (i% <> 16) And (i% <> 17) And (i% <> 29)) Then
        If (RezeptBoxen(i%).typ2) Or ((PrivatRezept% = 0) And (SonderBelegRezept% = 0)) Then
            Call PaintRezeptBox(i%)
        End If
    Next i%
    
    If (ParentForm.Name = frmAction.Name) Then
        For i% = 0 To 2
            With frmAction.cboVerfügbarkeit(i%)
                .Left = RezeptBoxen(3 + 4 * i).RightX + 300
                .Top = RezeptBoxen(3 + 4 * i).LeftY
                .Width = RezeptBoxen(6 + 4 * i).RightX - RezeptBoxen(4 + 4 * i).LeftX
                .Visible = MitVerfügbarkeit
                
                .Clear
                If (iDate(VerkDatum) < iDate("010719")) Then
                    .AddItem "1 - ordnungsgem. Abgabe bzw. leere Verord.Zeile"
                    .AddItem "2 - Nichtverfügbarkeit eines Rabattartikels"
                    .AddItem "3 - Nichtverfügbarkeit eines Importartikels"
                    .AddItem "4 - Nichtverfügbarkeit eines Rabatt- o. Importartikels"
                    .AddItem "5 - Notwendigkeit unverzüglicher Abgabe"
                    .AddItem "6 - Pharmazeutische Bedenken"
                    .AddItem "7 - Wunscharzneimittel"
                Else
                    .AddItem "1 - ordnungsgem. Abgabe bzw. leere Verord.Zeile"
                    .AddItem "2 - Nichtverfügbarkeit aller Rabattartikel"
                    .AddItem "3 - Nichtverf. d. 4 günstigsten Generika bzw. günstiger Importe"
                    .AddItem "4 - Nichverf. aller Rabattart. u. günstiger Generika/Importe"
                    .AddItem "5 - Notwendigkeit unverzüglicher Abgabe (statt Rabattart.)"
                    .AddItem "6 - Notwendigkeit unverzügl. Abgabe (statt Rabattart.u. günstiger Genrika/Importe)"
                    .AddItem "7 - Wunscharzneimittel"
                    .AddItem "8 - Pharmazeutische Bedenken (statt Rabattart.)"
                    .AddItem "9 - Pharmazeutische Bedenken (statt Rabattart.u. günstiger Generika/Importe)"
                End If
                
                If (AMverfügbarkeit$ <> "") Then
                    .ListIndex = Val(Mid$(AMverfügbarkeit$, i% + 1, 1)) - 1
                End If
            End With
        Next i%
    ElseIf (AMverfügbarkeit$ <> "") Then
        For i% = 0 To 2
            j% = Val(Mid$(AMverfügbarkeit$, i% + 1, 1))
            If (j% > 1) Then
                With frmRezEinzeln.picRezept
                    .CurrentX = RezeptBoxen(3 + 4 * i).RightX + 150
                    .CurrentY = RezeptBoxen(3 + 4 * i).LeftY
'                    If (j% = 2) Then
'                        h$ = "2 - Nichtverfügbarkeit eines Rabattartikels"
'                    ElseIf (j% = 3) Then
'                        h$ = "3 - Nichtverfügbarkeit eines Importartikels"
'                    ElseIf (j% = 4) Then
'                        h$ = "4 - Nichtverfügbarkeit eines Rabatt- oder Importartikels"
'                    ElseIf (j% = 5) Then
'                        h$ = "5 - Notwendigkeit unverzüglicher Abgabe"
'                    ElseIf (j% = 6) Then
'                        h$ = "6 - Pharmazeutische Bedenken"
'                    Else
'                        h$ = "7 - Wunscharzneimittel"
'                    End If
                    If (j% = 2) Then
                        h$ = "2 - Nichtverfügbarkeit aller Rabattartikel"
                    ElseIf (j% = 3) Then
                        h$ = "3 - Nichtverf. d. 4 günstigsten Generika bzw. günstiger Importe"
                    ElseIf (j% = 4) Then
                        h$ = "4 - Nichverf. aller Rabattart. u. günstiger Generika/Importe"
                    ElseIf (j% = 5) Then
                        h$ = "5 - Notwendigkeit unverzüglicher Abgabe"
                    ElseIf (j% = 6) Then
                        h$ = "6 - Notwendigkeit unverzügl. Abgabe (statt Rabattart.u. günstiger Genrika/Importe)"
                    ElseIf (j% = 7) Then
                        h$ = "7 - Wunscharzneimittel"
                    ElseIf (j% = 8) Then
                        h$ = "8 - Pharmazeutische Bedenken (statt Rabattart.)"
                    Else
                        h$ = "9 - Pharmazeutische Bedenken (statt Rabattart.u. günstiger Generika/Importe)"
                    End If
                    frmRezEinzeln.picRezept.Print h$;
                End With
            End If
        Next i%
    End If
    
    Call CalcRezeptWerte
    
    Call PaintRezeptWert(0, RezApoNr$)
    AnzZeilen% = AnzRezeptArtikel% - 1      'GS 16.05.02
    If AnzZeilen% > 6 Then AnzZeilen% = 6
    For i% = 0 To AnzZeilen%
        j% = 3 + i% * 4
        BlinkFlag% = RezArtikel(i%).flag And REZ_BLINK
        
        h$ = RezArtikel(i%).pzn
'        If (Left$(h$, 1) = "@") Then h$ = RezArtikel(i%).ScreenPzn
'        If (Left$(h$, 1) = "-") Then h$ = RezArtikel(i%).ScreenPzn
        If (Left(h, 1) = "-") Then
            h = Mid(h, 2) + "0"
            h = MakePzn(h)
        End If
        
        h2 = RezArtikel(i%).text
        
        If (Val(h) > 0) Then
            SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + h$
            'Set TaxeRec = TaxeDB.OpenRecordset(SQLStr$)
            On Error Resume Next
            TaxeRec.Close
            Err.Clear
            On Error GoTo DefErr
            TaxeRec.Open SQLStr, taxeAdoDB.ActiveConn
            If Not (TaxeRec.EOF) Then
                h2$ = Trim(TaxeRec!Name) + TaxeRec!menge + TaxeRec!einheit
            End If
        End If
        If (RezArtikel(i%).N0) Then
            h2 = h2 + " (N0)"
        End If
        
        If (i <= 2) And (Left(WVORabattArtikel, 3) <> "111") Then
            If (Mid(WVORabattArtikel, i + 1, 1) = "2") Then
                h2 = h2 + " (W)"
            End If
        End If
        
        Call PaintRezeptWert(j%, h2, BlinkFlag%)
        
        If (Left$(RezArtikel(i%).pzn, 1) = "-") Then
            h$ = RezArtikel(i%).ScreenPzn
        End If
        Call PaintRezeptWert(j% + 1, h$, BlinkFlag%)
        
        If (RezArtikel(i%).MietDauer > 0) Then
            Call PaintRezeptWert(j% + 2, Format(RezArtikel(i%).MietDauer, "0"), BlinkFlag%)
        Else
            Call PaintRezeptWert(j% + 2, Format(RezArtikel(i%).Faktor, "0"), BlinkFlag%)
        End If
        If (RezArtikel(i).VdbPauschale) Then
            Call PaintRezeptWert(j% + 3, Format(RezArtikel(i%).Preis * 100, "0"), BlinkFlag%)
        Else
            Call PaintRezeptWert(j% + 3, Format(RezArtikel(i%).Preis * 100 * RezArtikel(i%).Faktor, "0"), BlinkFlag%)
        End If
        
        
'Rezeptkontrolle: Kassenrezept bitte lachsfarben
'(wie Hintergrund der Spaltenüberschrift Artikel im Systemmanager),
'das jetzige Rufzeichen auf rotem Grund ändern in weißes Rufzeichen (wie in Icon) auf braunem Grund
'wie in der Kasse,
'das jetzige Rufzeichen auf grünem Grund ändern in  einen Haken (wie im Icon) auf olivgrünem Grund
'wie Rabattartikel in Flexkasse.
        If (i% < 3) Then
            h$ = " "
            If (RezArtikel(i%).AutIdem) Then h$ = "!"
            If (RezArtikel(i%).AutIdemKreuz = 2) Then h$ = Chr(67)
            
            If (para.Newline) And (RezArtikel(i%).AutIdem = 3) Then
                h$ = Chr$(214)
'                Call PaintRezeptBox(43 + i%, RGB(206, 185, 23))
                Call PaintRezeptBox(43 + i%, vbGreen)
                Call PaintRezeptWert(43 + i%, h$, False, 2)
                
            ElseIf (para.Newline) And (RezArtikel(i%).AutIdem = 4) Then
'                Call PaintRezeptBox(43 + i%, RGB(135, 61, 52))
                Call PaintRezeptBox(43 + i%, vbRed)
                Call PaintRezeptWert(43 + i%, h$, False, 2)
                
            ElseIf (RezArtikel(i%).AutIdem = 2) Then
                Call PaintRezeptBox(43 + i%, vbYellow)
                Call PaintRezeptWert(43 + i%, h$, False, True)
            ElseIf (RezArtikel(i%).AutIdem = 3) Then
                Call PaintRezeptBox(43 + i%, vbGreen)
                Call PaintRezeptWert(43 + i%, h$, False, True)
                
            ElseIf (RezArtikel(i%).AutIdem = 4) Then    '2
                Call PaintRezeptBox(43 + i%, vbRed)
                If (RezArtikel(i%).AutIdemKreuz = 2) Then
                    Call PaintRezeptWert(43 + i%, h$, False, 2)
                Else
                    Call PaintRezeptWert(43 + i%, h$, False, True)
                End If
            Else
                Call PaintRezeptWert(43 + i%, h$)
            End If
            
            h$ = " "
            IstRot% = False
            hImp123% = RezArtikel(i%).Imp123
            If (hImp123% > 0) Then
                If (hImp123% And 4) Then
                    h$ = "i"
                Else
                    h$ = "I"
                End If
                If (hImp123% And 1) Then
                    IstRot% = True
                End If
            End If
            If (IstRot%) Then
                Call PaintRezeptBox(50 + i%, vbRed)
                Call PaintRezeptWert(50 + i%, h$, False, True)
            Else
                Call PaintRezeptWert(50 + i%, h$)
            End If
        End If
    Next i%
    
    Call PaintRezeptWert(31, VerkDatum$)
    
    If (PrivatRezept% = 0) And (SonderBelegRezept% = 0) And (Selbsterklaerung = 0) Then
        If (GebFrei%) Then
            Call CheckRezeptBox(32)
        Else
            Call CheckRezeptBox(33)
        End If
    End If
    
    For i% = 0 To UBound(RezAddGesamtBrutto$)
        RezAddGesamtBrutto$(i%) = ""
    Next i%
    
    h$ = RezKundenInfo$(0) + vbCr + RezKundenInfo$(1) + vbCr + RezKundenInfo$(2)
    Call OemToChar(h$, h$)
    Call PaintRezeptWert(36, h$)
    
    With RezeptBoxen(36)
        KundenBoxX = .LeftX
        KundenBoxY = .LeftY
    End With

    If (SonderBelegRezept%) Then
        Call PaintRezeptWert(35, Trim(SonderBelege(SonderBelegRezept% - 1).KkBez))
        Call PaintRezeptWert(53, Trim(SonderBelege(SonderBelegRezept% - 1).KassenId))
        Call PaintRezeptWert(54, RezKundenInfo$(3))
        Call PaintRezeptWert(55, Trim(SonderBelege(SonderBelegRezept% - 1).Status))
        Call PaintRezeptWert(56, RezApoNr$)
        Call PaintRezeptWert(57, Trim(SonderBelege(SonderBelegRezept% - 1).GültigBis))
        Call PaintRezeptWert(58, VerkDatum$)
    End If
    
    
    If (IkSpezialName <> "") Then
        h$ = IkSpezialName
        Call PaintRezeptWert(34, h$)
    ElseIf (ActKkasse$ <> "") Then
        KasseRec.index = "Unique"
        KasseRec.Seek "=", ActKkasse$
        If (KasseRec.NoMatch) Or (KasseRec.EOF) Then
            SQLStr$ = "SELECT * FROM kKassen WHERE IK = " + Str$(Val(ActKkasse$))
'            Set KkassenRec = kKassenDB.OpenRecordset(SQLStr$)
'            If (KkassenRec.RecordCount > 0) Then
            FabsErrf = kKassenDB.OpenRecordset(KkassenRec, SQLStr)
            If Not (KkassenRec.EOF) Then
                h$ = Trim(KkassenRec!Name)
                Call PaintRezeptWert(34, h$)
            End If
        Else
            h$ = Trim(KasseRec!Name)
            Call PaintRezeptWert(34, h$)
        End If
    End If
    
    If (AutIdemIk& > 0) Then
        h$ = Format(AutIdemIk&, String(9, "0"))
        Call PaintRezeptWert(59, h$)
    End If

    Call PaintRezeptWert(60, RezKundenInfo$(5))
    
    .Font.Size = OrgFontSize%
    
    If (ParentForm.Name = frmAction.Name) Then
'        MsgBox (CStr(eRezeptVerordnungsTyp))
        If (eRezeptVerordnungsTyp = VerordnungsTyp_OP_Rezepturverordnung) Then  'Or (eRezeptVerordnungsTyp = VerordnungsTyp_OP_Freitextverordnung) Then
            If (FDok = 0) Then
                Set FD_OP = New TI_Back.Fachdienst_OP
                If (FD_OP.Init(FD_OP.NetWindowHandle(ParentForm.hWnd))) Then
                    FDok = True
                End If
            End If
        
            Dim eRezept2 As New TI_Back.eRezept_OP
            Call eRezept2.New2(eRezeptTaskId)
            
            With ParentForm.flxVerordnung
                .Rows = 0   ' .FixedRows
                
                h2 = eRezept2.Anzahl + "x" + vbTab
                With eRezept2.Verordnung
                    h2 = h2 + .sPackungsgroesse + vbTab + .einheit + vbTab + .Darreichungsform
                    If (.Verordnungstext <> "") Then
                        h2 = h2 + ": " + .Verordnungstext
                    End If
                    h2 = h2 + vbTab
                    ParentForm.flxVerordnung.AddItem (h2)
                    For i = 1 To .AnzRezepturBestandteile
                        With .RezepturBestandteil(i - 1)
'                            h = h + .pzn + " " + .Name + " " + .sMenge + " " + .Einheit + " " + .Darreichungsform + " " + .MengeFreitext + vbCrLf
                            h2 = vbTab + .sMenge + vbTab + .einheit + vbTab + .Name + vbTab + .pzn ' " " + .sMenge + " " + .Einheit + " " + .Darreichungsform + " " + .MengeFreitext + vbCrLf
                            ParentForm.flxVerordnung.AddItem (h2)
        '                    h = h + .Name + " " + vbCrLf
                        End With
                    Next i
                    If (.Herstellanweisung <> "") Then
                        h2 = vbTab + vbTab + vbTab + "(" + .Herstellanweisung + ")" + vbTab
                        ParentForm.flxVerordnung.AddItem (h2)
                    End If
                End With
                If (eRezept2.RezepturGebrauchsanweisung <> "") Then
                    h2 = vbTab + vbTab + vbTab + eRezept2.RezepturGebrauchsanweisung + vbTab
                    ParentForm.flxVerordnung.AddItem (h2)

                    .FillStyle = flexFillRepeat
                    .row = .Rows - 1
                    .col = 0
                    .RowSel = .row
                    .ColSel = .Cols - 1
                    .CellFontBold = True
                    .FillStyle = flexFillSingle
                End If
                
                .FillStyle = flexFillRepeat
                .row = 0
                .col = 0
                .RowSel = .row
                .ColSel = .Cols - 1
                .CellFontBold = True
                .FillStyle = flexFillSingle
                
                .BackColor = ParentForm.picToolTip.BackColor
                
                
                .Top = ParentForm.picRezept.Top + KundenBoxY + 60 ' wpara.TitelY
                .Left = ParentForm.picRezept.Left + KundenBoxX + 60 ' wpara.LinksX
                .Height = .Rows * .RowHeight(0) + 90
                
                .Visible = True
            End With
        End If
   
        If (Chef) Then
            ParentForm.mnuBearbeitenZusatz(4).Enabled = (PrivatRezept% = False) And (SonderBelegRezept% = 0)
            ParentForm.mnuBearbeitenZusatz(5).Enabled = (PrivatRezept% = False) And (SonderBelegRezept% = 0)
            
    '        ParentForm.mnuBearbeitenInd(MENU_F6).Enabled = True
            ParentForm.cmdToolbar(MENU_SF6).Enabled = True
            
        End If
        ParentForm.mnuLennartz.Enabled = (LennartzPfad <> "")
        ParentForm.ShowToolbar
        .SetFocus
    End If
End With

If (DirektNachAufruf) And (BtmRezept) Then
    Call iMsgBox("Stimmt das Abgabedatum des BTM mit dem zu druckenden Belieferungsdatum des Rezeptes überein?", vbOKOnly Or vbQuestion, "BTM-Rezept")
End If

If (Selbsterklaerung) Then
    If (SchutzMasken) Then
        With ParentForm.picRezept
            '.ForeColor = vbRed
            .FontBold = True
            .FontSize = .FontSize + 6
            .CurrentX = RezeptBoxen(41).LeftX
            .CurrentY = RezeptBoxen(3).LeftY - 90
'            ParentForm.picRezept.Print "Schutzmasken";
            ParentForm.picRezept.Print frmAction.mnuSchutzmaskenInd(SchutzMasken - 1).Caption;
            .FontBold = False
            .FontSize = .FontSize - 6
        End With
    Else
        h$ = "Nacht- und Notdienstfonds des DAV"
        Call PaintRezeptWert(34, h$)
        
        With ParentForm.picRezept
            .ForeColor = vbRed
            .FontBold = True
            .FontSize = .FontSize + 2
            .CurrentX = RezeptBoxen(3).LeftX
            .CurrentY = RezeptBoxen(3).LeftY
            ParentForm.picRezept.Print "Kein Rezept";
            
            .FontBold = False
            .CurrentX = RezeptBoxen(7).LeftX
            .CurrentY = RezeptBoxen(7).LeftY
            ParentForm.picRezept.Print "Selbsterklärung zur Förderung der Sicherstellung .....";
            
            h = "01." + Mid(VerkDatum, 3, 2) + ".20" + Mid(VerkDatum, 5, 2) + " 00:00"
            h = Format(h, "mmmm yyyy")
            WinVkText = "Grüne-/Privat-Rezepte RX aus Verkaufsdatei " + h + ":   "
            
            .FontBold = False
            WinVkX = RezeptBoxen(4).LeftX
            WinVkY = RezeptBoxen(3).LeftY
    '        .CurrentX = WinVkX
    '        .CurrentY = WinVkY
    '        ParentForm.picRezept.Print WinVkText
            
            With ParentForm.lblWinVk
                .FontSize = ParentForm.picRezept.FontSize
                x1 = ParentForm.picRezept.TextWidth(WinVkText + String(8, "9"))
                .Left = RezeptBoxen(6).RightX - x1
                .Top = WinVkY
                .Width = x1
                .Height = 3 * ParentForm.picRezept.TextHeight("Äg")
                .ForeColor = vbRed
                .Caption = WinVkText + vbCrLf + "... bitte warten ..."
                .Visible = True
            End With
            
            .ForeColor = vbBlack
            .FontSize = .FontSize - 2
        End With
    End If
        
    If (SchutzMasken) Then
        Call TabIndexChange(5, False)
    Else
        h = BoxenWert(52)
        Call PaintRezeptWert(52, h$)
        Call TabIndexChange(6, False)
    End If
    
Else
    Call TabIndexChange(3, False)
End If

Call DefErrPop
End Sub

Sub CalcRezeptWerte()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("CalcRezeptWerte")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
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
Dim i%, k%, ind%, iFaktor%, AnzBlink%, BlinkFlag%, AnzArtikel%
Dim dWert#, GesHerstRabattPrivat130Brutto#
Static InSub%

If (InSub) Then
    Call DefErrPop: Exit Sub
End If

InSub = True

LieferEngPass = 0

If (ParentForm.Name = frmAction.Name) Then
    If (ActPauschaleNr > 0) Then
        Dim j%, SollHmNr$, pzn$
        SollHmNr = ""
        For i% = 0 To (AnzRezeptArtikel% - 1)
            pzn$ = RezArtikel(i%).pzn
            If (Left(pzn, 1) = "-") Then
                pzn = Mid(pzn, 2) + "0"
                pzn = MakePzn(pzn)
            End If
            
            If (Val(pzn) > 0) Then
                SQLStr$ = "SELECT * FROM VdbArtikel WHERE PZN = " + pzn$
                'Set VdbPznRec = AplusVDB.OpenRecordset(SQLStr$)
                FabsErrf = AplusVDB.OpenRecordset(VdbPznRec, SQLStr)
                If Not (VdbPznRec.EOF) Then
                    If (SollHmNr = "") Then
                        SollHmNr = Left(Trim(CheckNullStr(VdbPznRec!HmPositionsNr1)), 7)
                    ElseIf (Left(Trim(CheckNullStr(VdbPznRec!HmPositionsNr1)), 7) = SollHmNr) Then
                        RezArtikel(i).Preis = 0
                        RezArtikel(i).Zuz = 0
                        
                        j% = 3 + i * 4
                        Call PaintRezeptBox(j% + 3)
                        Call PaintRezeptWert(j% + 3, Format(RezArtikel(i%).Preis * 100 * RezArtikel(i%).Faktor, "0"), 0)
                    End If
                End If
            End If
        Next i
    End If
    

    RezSumme# = RezSummeDazu#
    'For i% = 0 To UBound(RezArtikel)
    AnzArtikel% = (AnzRezeptArtikel% - 1)     'GS 16.05.02
'    If AnzArtikel% > 10 Then AnzArtikel% = 10

    ind = -1
    GesHerstRabattPrivat130Brutto = 0
    For i% = 0 To AnzArtikel%
        If (RezArtikel(i).pzn = "02567739") Then
            ind = i
        Else
            If (RezArtikel(i).VdbPauschale) Then
                RezSumme# = RezSumme# + (RezArtikel(i%).Preis)
            Else
                RezSumme# = RezSumme# + (RezArtikel(i%).Preis * RezArtikel(i%).Faktor)
            End If
            GesHerstRabattPrivat130Brutto = GesHerstRabattPrivat130Brutto + (RezArtikel(i%).HerstRabattPrivat130Brutto * RezArtikel(i%).Faktor)
        End If
    Next i%
    
'*** Ver 5.3.0 am 14.8.23
    Dim sVerfügbarkeit2$
    sVerfügbarkeit2 = CheckVerfügbarkeit
    If (sVerfügbarkeit2 <> "111") Then
        For i = 1 To 3
            If (InStr("234", Mid(sVerfügbarkeit2, i, 1)) > 0) Then
                LieferEngPass = 0.6
                RezSumme = RezSumme + LieferEngPass
            End If
        Next i
    End If
'***

    If (PrivatRezept) Then
        If (GesHerstRabattPrivat130Brutto > 0) Then
            RezSumme = RezSumme - GesHerstRabattPrivat130Brutto
            
            If (ind < 0) Then
                AnzRezeptArtikel = AnzRezeptArtikel + 1
                ind = (AnzRezeptArtikel% - 1)
                Call InitRezeptArtikel(ind)
            End If
            RezArtikel(ind).pzn = "02567739"
            RezArtikel(ind).text = "Erstattungsbeträge §130b"
            RezArtikel(ind).Faktor = 1
            RezArtikel(ind).Preis = -GesHerstRabattPrivat130Brutto
            RezArtikel(ind).Zuz = 0
            Call PaintArtikel(ind, 0)
        ElseIf (ind >= 0) Then
            For i% = ind% To AnzArtikel
                RezArtikel(i%) = RezArtikel(i% + 1)
                Call PaintArtikel(i, 0)
            Next i%
            Call InitRezeptArtikel(UBound(RezArtikel))
            
            AnzRezeptArtikel% = AnzRezeptArtikel% - 1
'            Call PaintRezept
'            Call TabIndexChange(3 + (ind% * 4), True)
        End If
    End If
End If
Call PaintRezeptBox(2)
Call PaintRezeptBox(47)
Call PaintRezeptBox(48)
Call PaintRezeptWert(2, Format(RezSumme# * 100, "0"))
    
If (ParentForm.Name = frmAction.Name) Then
    If (RezGebRechnen%) Then
        RezGebSumme# = 0
        If Not (GebFrei%) Then
    '        For i% = 0 To UBound(RezArtikel)
            For i% = 0 To (AnzRezeptArtikel% - 1)
                iFaktor% = RezArtikel(i%).Faktor
'                If (iFaktor% > 1) And (RezArtikel(i%).flag And REZ_WG_4) Then
'                    If (RezArtikel(i%).IstWg4) Then
'                        iFaktor% = 1
'                    End If
'                End If
                If (iFaktor% > 1) And ((RezArtikel(i%).IstWg4) Or (RezArtikel(i).VdbPauschale)) Then
                    iFaktor% = 1
                End If
                dWert# = (RezArtikel(i%).Zuz * iFaktor%)
                If (iFaktor% > 1) Then
                    If (dWert# > 9.989) And (dWert# < 10#) Then
                        dWert# = 10
                    End If
                End If
                RezGebSumme# = RezGebSumme# + dWert#
                
'                MsgBox ("<" + RezArtikel(i).pzn + ">" + " <" + RezArtikel(i).ScreenPzn + ">")
            Next i%
        End If
    End If
End If
If Not (PrivatRezept%) And (SonderBelegRezept% = 0) And (Selbsterklaerung = 0) Then
    Call PaintRezeptBox(1)
    Call PaintRezeptBox(46)
    Call PaintRezeptWert(1, Format(RezGebSumme# * 100, "0"))
End If
                    
If (SchutzMasken) Then
    RezGebSumme = IIf(SchutzMasken <= 2, 2, 0) * RezArtikel(0).Faktor
    Call PaintRezeptBox(1)
    Call PaintRezeptWert(1, Format(RezGebSumme# * 100, "0"))
End If

AnzBlink% = 0
'If (lRezNr& <> 0&) And (PrivatRezept% = 0) Then
If (PrivatRezept% = 0) And (SonderBelegRezept% = 0) Then
    For i% = 0 To (AnzRezeptArtikel% - 1)
        BlinkFlag% = RezArtikel(i%).flag And REZ_BLINK
        If (BlinkFlag%) Then AnzBlink% = AnzBlink% + 1
    Next i%
End If
If (Trim(RezApoNr$) = "") Then AnzBlink% = 1

If (ParentForm.Name = frmAction.Name) Then
    ParentForm.mnuBearbeitenInd(MENU_F6).Enabled = (AnzBlink% = 0)
    ParentForm.cmdToolbar(MENU_F6 + 1).Enabled = (AnzBlink% = 0)
    Call ParentForm.ShowToolbar
End If
        
InSub = False

Call DefErrPop
End Sub

Sub PaintRezeptBox(pos%, Optional color& = vbWhite)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("PaintRezeptBox")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
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
Dim lDrawWidth&, CurrX&, CurrY&

If (pos% >= 50) And (pos% <= 52) And (ImpFlag% = False) And (Selbsterklaerung = 0) Then Call DefErrPop: Exit Sub

If (Selbsterklaerung) And (InStr(SelbstErklaerungHidden, "," + CStr(pos) + ",") > 0) Then
    Call DefErrPop: Exit Sub
End If

With ParentForm.picRezept
    CurrX& = .CurrentX
    CurrY& = .CurrentY
    
    If (RezeptBoxen(pos%).Fett) Then
        lDrawWidth& = .DrawWidth
        .DrawWidth = lDrawWidth& * 2
    End If
End With

If (para.Newline) Then
    If (color = vbGreen) Then
        color = wpara.nlGreen
    ElseIf (color = vbRed) Then
        If (Knallrot = 0) Then
            color = wpara.nlRed
        End If
    ElseIf (color = vbYellow) Then
        color = wpara.nlYellow
    End If
End If


With RezeptBoxen(pos%)
    ParentForm.picRezept.Line (.LeftX, .LeftY)-(.RightX, .RightY), color&, BF
    If (pos% <= 52) Then
        ParentForm.picRezept.Line (.LeftX, .LeftY)-(.RightX, .RightY), vbBlack, B
    End If
End With
    
With ParentForm.picRezept
    If (RezeptBoxen(pos%).Fett) Then
        .DrawWidth = lDrawWidth&
    End If

    .CurrentX = CurrX&
    .CurrentY = CurrY&
End With

Call DefErrPop
End Sub

Sub PaintRezeptDaten(Optional lColor& = -1)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("PaintRezeptDaten")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
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
Dim i%, OrgFontSize%, iFontSize%, CurrX%
Dim diff&, MaxWidth&, MaxHeight&
Dim h$, h2$, Wert$, OrgFontName$
Dim iRec As Recordset
Dim oFont As Font

If (Selbsterklaerung) Then
    Call DefErrPop: Exit Sub
End If

h$ = "Abgabe:   " + OrgVerkDatum$
If (VerkZeit% > 0) Then
    h$ = h$ + "  " + Format(VerkZeit \ 100, "00") + ":" + Format(VerkZeit Mod 100, "00")
Else
    h$ = h$ + Space$(7)
End If
If (VerkUser$ <> "") Then
    h$ = h$ + "   Computer " + VerkUser$ + Mid$("LR", VerkLiRe + 1, 1)
Else
    h$ = h$ + Space$(8)
End If
Wert$ = h$

If (PersonalNr% > 0) Then
    h$ = h$ + vbCrLf + "Benutzer:  " + para.Personal(PersonalNr%)
End If
If (SchonDa%) Then
    h$ = h$ + vbCrLf + "Druck:  " + ""
End If

'Wert$ = h$ + vbCrLf + "Benutzer: "
'Wert$ = Wert$ + vbCrLf + "Druck: "

With ParentForm.picRezept
    OrgFontName$ = .Font.Name
    OrgFontSize% = .Font.Size

    MaxWidth& = .Width - 45
    MaxHeight& = .Height - RezeptBoxen(49).LeftY - 45
    
    h2$ = ""
    For i = 1 To 5
        h2 = h2 + "1" + vbCrLf
    Next i
    Do
        diff& = MaxHeight& - .TextHeight(h2$)
        If (diff& >= 0) Then Exit Do
        
        iFontSize% = .Font.Size - 1
        If (iFontSize% < 8) Then
            .Font.Name = "Small Fonts"
        End If
        .Font.Size = iFontSize%
    Loop

    Do
'        diff& = MaxWidth& - .TextWidth(h$)
        diff& = MaxWidth& - .TextWidth(String(100, "X"))
        If (diff& >= 0) Then Exit Do
        iFontSize% = .Font.Size - 1
        If (iFontSize% < 8) Then
            .Font.Name = "Small Fonts"
        End If
        .Font.Size = iFontSize%
    Loop
    
    CurrX% = .Width - .TextWidth(h$) - 90
    .CurrentX = CurrX%
    .CurrentY = RezeptBoxen(49).LeftY + 45
    ParentForm.picRezept.Print Wert$
    
    If (PersonalNr% > 0) Then
        .CurrentX = CurrX%
        ParentForm.picRezept.Print "Benutzer:  " + para.Personal(PersonalNr%)
    End If
    If (MarsFreigabePersonalNr >= 0) Then
        .CurrentX = CurrX%
        .CurrentY = .CurrentY + 30
        ParentForm.picRezept.Print "Freigabe:  " + Mid(DruckDatum, 5, 2) + Mid(DruckDatum, 3, 2) + Left(DruckDatum, 2) + " " + Format(DruckZeit \ 100, "00") + ":" + Format(DruckZeit Mod 100, "00")
        .CurrentX = CurrX%
        ParentForm.picRezept.Print "Benutzer:  " + para.Personal(MarsFreigabePersonalNr%)
    End If
    If (SchonDa%) Then
        .CurrentX = CurrX%
        .CurrentY = .CurrentY + 30
        h$ = DruckDatum$
        ParentForm.picRezept.Print "Druck:      " + Right$(h$, 2) + Mid$(h$, 3, 2) + Left$(h$, 2)
        
        SQLStr$ = "SELECT * FROM AbrechnungsMeldungen WHERE RezepteUnique='" + RezepteRec!Unique + "'"
        SQLStr$ = SQLStr$ + " ORDER BY LaufNr"
        Set iRec = RezSpeicherDB.OpenRecordset(SQLStr$)
        If Not (iRec.EOF) Then
            CurrX% = 90
            .CurrentX = CurrX%
            .CurrentY = RezeptBoxen(49).LeftY + 45
            .ForeColor = vbRed
            ParentForm.picRezept.Print "Rückmeldung(en) der Abrechnungsstelle:"
    
            Do
                If (iRec.EOF) Then
                    Exit Do
                End If
                
                h = iRec!fStatus + ":  " + CheckNullStr(iRec!fKommentar)
                If (CheckNullStr(iRec!fKurzText) <> "") Then
                    If (CheckNullStr(iRec!fKommentar) <> iRec!fKurzText) Then
                        h = h + "  " + iRec!fKurzText
                    End If
                End If
                If (iRec!PosNr > 0) Then
                    h = h + " (Zeile " + CStr(iRec!PosNr) + ")"
                End If
                .CurrentX = CurrX%
                Do
                    If (h = "") Then
                        Exit Do
                    End If
                    ParentForm.picRezept.Print Left$(h$, 80)
                    h = Mid$(h, 81)
                    .CurrentX = CurrX% + 600
                Loop
                
                iRec.MoveNext
            Loop
            .ForeColor = vbBlack
        End If
    End If
    If (CurrX > 90) Then
        CurrX% = 90
        .Font.Name = "Arial"
        .CurrentX = CurrX%
        .CurrentY = RezeptBoxen(49).LeftY + 45
        ParentForm.picRezept.Print "Tausch der Zeilen mit  Shift+";
        .Font.Name = "Symbol"
        ParentForm.picRezept.Print Chr$(173);
        .Font.Name = "Arial"
        ParentForm.picRezept.Print " oder Shift+";
        .Font.Name = "Symbol"
        ParentForm.picRezept.Print Chr$(175)
'        ParentForm.picRezept.Print "Tausch der Zeilen mit  Shift+" + Chr$(24) + " oder Shift+" + Chr$(25);
        .Font.Name = "Arial"
        .CurrentX = CurrX%
        ParentForm.picRezept.Print "Aut-Idem mit ALT+G, erweitertes Aut-Idem mit STRG+G, Aufruf Kundenkartei mit STRG+S"
        .CurrentX = CurrX%
        ParentForm.picRezept.Print "Aufruf Artikelstamm+X mit STRG+X, Nichtverfügbarkeitsabfrage mit Strg+N"
    
    End If
    
    .Font.Name = OrgFontName$
    .Font.Size = OrgFontSize%
End With



'.TextMatrix(0, 1) = h$
'
'If (PersonalNr% > 0) Then
'    .TextMatrix(1, 1) = para.Personal(PersonalNr%)
'End If
'
'
'
'With ParentForm.flxRezeptDaten
'    If (lColor& > 0) Then .BackColor = lColor&
'    For i% = 0 To (.Rows - 1)
'        .TextMatrix(i%, 1) = ""
'        .TextMatrix(i%, 3) = ""
'    Next i%
'
'    h$ = VerkDatum$
'    If (VerkZeit% > 0) Then
'        h$ = h$ + "  " + Format(VerkZeit / 100, "00") + ":" + Format(VerkZeit Mod 100, "00")
'    Else
'        h$ = h$ + Space$(7)
'    End If
'    If (VerkUser$ <> "") Then
'        h$ = h$ + "   Comp: " + VerkUser$ + Mid$("LR", VerkLiRe + 1, 1)
'    Else
'        h$ = h$ + Space$(8)
'    End If
'    .TextMatrix(0, 1) = h$
'
'    If (PersonalNr% > 0) Then
'        .TextMatrix(1, 1) = para.Personal(PersonalNr%)
'    End If
'
'    .Left = ParentForm.picRezept.Left + ParentForm.picRezept.Width - 60 - .Width
'    .Top = ParentForm.picRezept.Top + RezeptBoxen(49).LeftY + 60
'    .Visible = True
'End With

Call DefErrPop
End Sub

Sub CheckRezeptBox(pos%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("CheckRezeptBox")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
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
Dim lDrawWidth&

With RezeptBoxen(pos%)
    lDrawWidth& = ParentForm.picRezept.DrawWidth
    ParentForm.picRezept.DrawWidth = lDrawWidth& * 2
    ParentForm.picRezept.Line (.LeftX + 15, .RightY - 15)-(.RightX - 15, .LeftY + 15), vbBlack
    ParentForm.picRezept.Line (.LeftX + 15, .LeftY + 15)-(.RightX - 15, .RightY - 15), vbBlack
    ParentForm.picRezept.DrawWidth = lDrawWidth&
End With
    
Call DefErrPop
End Sub

Sub PaintRezeptWert(pos%, OrgWert$, Optional BlinkFlag% = False, Optional Invers% = False)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("PaintRezeptWert")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
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
Dim i%, l%, wi%, typ2%, iFontSize%, OrgFontSize%, OrgFontBold%, ind%
Dim MaxWidth&, MaxHeight&, EchtWidth&, diff&, iColor&
Dim CurrX&, CurrY&, x&, OrgForeColor&
Dim Typ1$, ch$, Wert$, OrgFontName$

Wert$ = OrgWert$
If (pos% = 0) Then
    If (Trim(Wert$) <> "") Then
        Wert$ = "+" + Wert$ + "+"
    End If
End If

BoxenWert$(pos%) = Wert$

If (pos = 43) And (Wert = Chr(67)) Then
    Call CheckRezeptBox(pos)
    Call DefErrPop: Exit Sub
End If

With RezeptBoxen(pos%)
    Typ1$ = Left$(.Typ, 1)
    typ2% = Val(Right$(.Typ, 1))
    EchtWidth& = (.RightX - .LeftX)
    MaxWidth& = EchtWidth& - 45
    MaxHeight& = (.RightY - .LeftY) - 30
End With

With ParentForm.picRezept
    OrgFontName$ = .Font.Name
    OrgFontSize% = .Font.Size
    OrgForeColor& = .ForeColor
    OrgFontBold% = .FontBold

    iColor& = .ForeColor
    If (BlinkFlag%) Then
        .ForeColor = vbRed
    ElseIf (Invers%) Then
        .ForeColor = vbBlack ' vbWhite
    End If
    
    If (Invers% = 2) Then
        .FontName = "Symbol"
        .ForeColor = vbWhite
        .FontBold = True
        .Font.Size = OrgFontSize + 2
        diff& = MaxHeight& - .TextHeight(Wert$)
        CurrY& = RezeptBoxen(pos%).LeftY + diff& / 2 + 30
        
        diff& = MaxWidth& - .TextWidth(Wert$)
        .CurrentX = RezeptBoxen(pos%).LeftX + diff& / 2 + 45
        .CurrentY = CurrY&
        ParentForm.picRezept.Print Wert$;
    Else
        If (pos = 43) Or (pos = 46) Or (pos = 49) Or (pos = 50) Or (pos = 53) Or (pos = 56) Then
            .FontBold = True
        End If
        
        Do
            diff& = MaxHeight& - .TextHeight(Wert$)
            If (diff& >= 0) Then Exit Do
            
            iFontSize% = .Font.Size - 1
            If (iFontSize% < 8) Then
                .Font.Name = "Small Fonts"
            End If
            .Font.Size = iFontSize%
        Loop
        CurrY& = RezeptBoxen(pos%).LeftY + diff& / 2 + 30
        
        Select Case Typ1$
            Case "L"
                If (typ2% = 0) Then
                    CurrX& = RezeptBoxen(pos%).LeftX + 45
                    Do
                        diff& = MaxWidth& - .TextWidth(Wert$)
                        If (diff& >= 0) Then Exit Do
                        iFontSize% = .Font.Size - 1
                        If (iFontSize% < 8) Then
                            .Font.Name = "Small Fonts"
                        End If
                        .Font.Size = iFontSize%
                        If iFontSize% >= 1 Then Exit Do
                    Loop
                    .CurrentY = CurrY&
                    Do
                        .CurrentX = CurrX&
                        ind% = InStr(Wert$, vbCr)
                        If (ind% > 0) Then
                            ParentForm.picRezept.Print Left$(Wert$, ind% - 1)
                            Wert$ = Mid$(Wert$, ind% + 1)
                        Else
                            ParentForm.picRezept.Print Wert$;
                            Exit Do
                        End If
                    Loop
                ElseIf (typ2% = 1) Then
                    If (pos% = 0) And (OrgWert$ <> OrgRezApoNr$) Then .ForeColor = vbRed
                    l% = Len(Wert$)
                    wi% = EchtWidth& / l%
                    For i% = 1 To l%
                        ch$ = Mid$(Wert$, i%, 1)
                        CurrX& = (i% - 1) * wi% + (wi% - .TextWidth(ch$)) / 2
                        .CurrentX = RezeptBoxen(pos%).LeftX + CurrX&
                        .CurrentY = CurrY&
                        ParentForm.picRezept.Print ch$;
                    Next i%
                End If
            
            Case "Z"
                If (typ2% = 0) Then
                    If (pos% >= 43) Then .Font.Name = "Times New Roman"
                    Do
                        diff& = MaxWidth& - .TextWidth(Wert$)
                        If (diff& >= 0) Then Exit Do
                        iFontSize% = .Font.Size - 1
                        If (iFontSize% < 8) Then
                            .Font.Name = "Small Fonts"
                        End If
                        .Font.Size = iFontSize%
                        If iFontSize% <= 1 Then Exit Do
                    Loop
                    .CurrentX = RezeptBoxen(pos%).LeftX + diff& / 2 + 45
                    .CurrentY = CurrY&
                    ParentForm.picRezept.Print Wert$;
                End If
                
            Case "R"
                If (typ2% = 0) Then
                    Do
                        diff& = MaxWidth& - .TextWidth(Wert$)
                        If (diff& >= 0) Then Exit Do
                        iFontSize% = .Font.Size - 1
                        If (iFontSize% < 8) Then
                            .Font.Name = "Small Fonts"
                        End If
                        .Font.Size = iFontSize%
                        If iFontSize% <= 1 Then Exit Do
                    Loop
                    .CurrentX = RezeptBoxen(pos%).RightX - .TextWidth(Wert$) - 45
                    .CurrentY = CurrY&
                    ParentForm.picRezept.Print Wert$;
                ElseIf (typ2% > 1) Then
                    If (Len(Wert$) > typ2%) Then
                        typ2% = Len(Wert$)
                    End If
                    
                    Wert$ = Right$(Space$(typ2%) + Wert$, typ2%)
                    
                    wi% = EchtWidth& / typ2%
                    For i% = 1 To typ2%
                        ch$ = Mid$(Wert$, i%, 1)
                        
                        CurrX& = (i% - 1) * wi% + (wi% - .TextWidth(ch$)) / 2
                        .CurrentX = RezeptBoxen(pos%).LeftX + CurrX&
                        .CurrentY = CurrY&
                        ParentForm.picRezept.Print ch$;
                    Next i%
                End If
        End Select
    End If

'    If (BlinkFlag%) Then
        .ForeColor = iColor&
'    End If

    .ForeColor = OrgForeColor&
    .Font.Name = OrgFontName$
    .Font.Size = OrgFontSize%
    .FontBold = OrgFontBold%
End With

'With frmAction.picRezept
'    wi% = .TextWidth(wert$)
'End With
'
'With RezeptBoxen(pos%)
'    CurrX& = .RightX - 45 - wi%
'    CurrY& = .LeftY + 45
'End With
'
'With frmAction.picRezept
'    .CurrentX = CurrX&
'    .CurrentY = CurrY&
'    frmAction.picRezept.Print wert$
'End With

Call DefErrPop
End Sub

Sub KeyDown(KeyCode%, Shift%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("KeyDown")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
Dim NeuIndex%, ind%

NeuIndex% = AktTabIndex%
If (Selbsterklaerung) Then
    If (KeyCode% = vbKeyRight) Or (KeyCode% = vbKeyDown) Then
        Do
            NeuIndex% = NeuIndex% + 1
            If (NeuIndex% > UBound(RezeptBoxen)) Then
                NeuIndex% = 0
            End If
            If (InStr(SelbstErklaerungTabstop, "," + CStr(NeuIndex) + ",") > 0) Then
                Exit Do
            End If
        Loop
        If (NeuIndex <> AktTabIndex%) Then Call TabIndexChange(NeuIndex%)
        Call DefErrPop: Exit Sub
    ElseIf (KeyCode% = vbKeyLeft) Or (KeyCode% = vbKeyUp) Then
        Do
            NeuIndex% = NeuIndex% - 1
            If (NeuIndex% < 0) Then
                NeuIndex% = UBound(RezeptBoxen)
            End If
            If (InStr(SelbstErklaerungTabstop, "," + CStr(NeuIndex) + ",") > 0) Then
                Exit Do
            End If
        Loop
        If (NeuIndex <> AktTabIndex%) Then Call TabIndexChange(NeuIndex%)
        Call DefErrPop: Exit Sub
    End If
End If

If (KeyCode% = vbKeyRight) Then
    Do
        NeuIndex% = NeuIndex% + 1
        If (NeuIndex% = 55) Then
            NeuIndex% = 0
        ElseIf (NeuIndex% = 32) Then
            If (SonderBelegRezept% > 0) Then
                NeuIndex% = 54
            Else
                NeuIndex% = 0
            End If
        ElseIf (NeuIndex% > MaxPruefBox%) Then
            NeuIndex% = 31
        End If
        If (RezeptBoxen(NeuIndex%).typ2) Or ((PrivatRezept% = 0) And (SonderBelegRezept% = 0)) Then
            Exit Do
        End If
    Loop
ElseIf (KeyCode% = vbKeyLeft) Then
    Do
        NeuIndex% = NeuIndex% - 1
        If (NeuIndex% < 0) Then
            If (SonderBelegRezept% > 0) Then
                NeuIndex% = 54
            Else
                NeuIndex% = 31
            End If
        ElseIf (NeuIndex% = 53) Then
            NeuIndex% = 31
        ElseIf (NeuIndex% = 30) Then
            NeuIndex% = MaxPruefBox%
        End If
        If (RezeptBoxen(NeuIndex%).typ2) Or ((PrivatRezept% = 0) And (SonderBelegRezept% = 0)) Then
            Exit Do
        End If
    Loop
ElseIf (KeyCode% = vbKeyDown) Then
    If (Shift And vbShiftMask) Then
        If (NeuIndex >= 3) And (NeuIndex <= MaxPruefBox - 4) Then
            ind = (NeuIndex - 3) \ 4
            
            RezMagFeld&(MAX_VERK_ZEILEN) = RezMagFeld&(ind + 1)
            RezMagAnzMag%(MAX_VERK_ZEILEN) = RezMagAnzMag%(ind + 1)
            RezArtikel(MAX_VERK_ZEILEN) = RezArtikel(ind + 1)
            HmNummer$(MAX_VERK_ZEILEN) = HmNummer$(ind + 1)
            HmFaktor%(MAX_VERK_ZEILEN) = HmFaktor%(ind + 1)
            HmStückPreis#(MAX_VERK_ZEILEN) = HmStückPreis#(ind + 1)
                
            RezMagFeld&(ind + 1) = RezMagFeld&(ind)
            RezMagAnzMag%(ind + 1) = RezMagAnzMag%(ind)
            RezArtikel(ind + 1) = RezArtikel(ind)
            HmNummer$(ind + 1) = HmNummer$(ind)
            HmFaktor%(ind + 1) = HmFaktor%(ind)
            HmStückPreis#(ind + 1) = HmStückPreis#(ind)
                
            RezMagFeld&(ind) = RezMagFeld&(MAX_VERK_ZEILEN)
            RezMagAnzMag%(ind) = RezMagAnzMag%(MAX_VERK_ZEILEN)
            RezArtikel(ind) = RezArtikel(MAX_VERK_ZEILEN)
            HmNummer$(ind) = HmNummer$(MAX_VERK_ZEILEN)
            HmFaktor%(ind) = HmFaktor%(MAX_VERK_ZEILEN)
            HmStückPreis#(ind) = HmStückPreis#(MAX_VERK_ZEILEN)
    
            AMverfügbarkeit$ = CheckVerfügbarkeit
            If (AMverfügbarkeit$ <> "") And (ind% <= 2) Then
                If (ind = 0) Then
                    AMverfügbarkeit$ = Mid$(AMverfügbarkeit$, ind + 2, 1) + Mid$(AMverfügbarkeit$, ind + 1, 1) + Mid$(AMverfügbarkeit$, ind + 3)
                ElseIf (ind = 1) Then
                    AMverfügbarkeit$ = Left(AMverfügbarkeit, 1) + Mid$(AMverfügbarkeit$, ind + 2, 1) + Mid$(AMverfügbarkeit$, ind + 1, 1)
                End If
            End If
            
            NeuIndex% = NeuIndex% + 4
            
            Call PaintRezept
        End If
    ElseIf (NeuIndex% = 0) Then
        NeuIndex% = 2
    ElseIf (NeuIndex% <= 2) Then
        NeuIndex% = 6
    ElseIf (NeuIndex% <= (MaxPruefBox% - 4)) Then
        NeuIndex% = NeuIndex% + 4
    ElseIf (NeuIndex% = 31) Then
        If (SonderBelegRezept% > 0) Then
            NeuIndex% = 54
        Else
            NeuIndex% = 0
        End If
    ElseIf (NeuIndex% = 54) Then
        NeuIndex% = 0
    Else
        NeuIndex% = 31
    End If
ElseIf (KeyCode% = vbKeyUp) Then
    If (Shift And vbShiftMask) Then
        If (NeuIndex >= 7) And (NeuIndex <= MaxPruefBox) Then
            ind = (NeuIndex - 3) \ 4
            
            RezMagFeld&(MAX_VERK_ZEILEN) = RezMagFeld&(ind - 1)
            RezMagAnzMag%(MAX_VERK_ZEILEN) = RezMagAnzMag%(ind - 1)
            RezArtikel(MAX_VERK_ZEILEN) = RezArtikel(ind - 1)
            HmNummer$(MAX_VERK_ZEILEN) = HmNummer$(ind - 1)
            HmFaktor%(MAX_VERK_ZEILEN) = HmFaktor%(ind - 1)
            HmStückPreis#(MAX_VERK_ZEILEN) = HmStückPreis#(ind - 1)
                
            RezMagFeld&(ind - 1) = RezMagFeld&(ind)
            RezMagAnzMag%(ind - 1) = RezMagAnzMag%(ind)
            RezArtikel(ind - 1) = RezArtikel(ind)
            HmNummer$(ind - 1) = HmNummer$(ind)
            HmFaktor%(ind - 1) = HmFaktor%(ind)
            HmStückPreis#(ind - 1) = HmStückPreis#(ind)
                
            RezMagFeld&(ind) = RezMagFeld&(MAX_VERK_ZEILEN)
            RezMagAnzMag%(ind) = RezMagAnzMag%(MAX_VERK_ZEILEN)
            RezArtikel(ind) = RezArtikel(MAX_VERK_ZEILEN)
            HmNummer$(ind) = HmNummer$(MAX_VERK_ZEILEN)
            HmFaktor%(ind) = HmFaktor%(MAX_VERK_ZEILEN)
            HmStückPreis#(ind) = HmStückPreis#(MAX_VERK_ZEILEN)
    
    
            AMverfügbarkeit$ = CheckVerfügbarkeit
            If (AMverfügbarkeit$ <> "") And (ind% <= 2) Then
                If (ind = 1) Then
                    AMverfügbarkeit$ = Mid$(AMverfügbarkeit$, ind + 1, 1) + Mid$(AMverfügbarkeit$, ind, 1) + Mid$(AMverfügbarkeit$, ind + 2)
                ElseIf (ind = 2) Then
                    AMverfügbarkeit$ = Left(AMverfügbarkeit, 1) + Mid$(AMverfügbarkeit$, ind + 1, 1) + Mid$(AMverfügbarkeit$, ind, 1)
                End If
            End If
            
            NeuIndex% = NeuIndex% - 4
            
            Call PaintRezept
        End If
    ElseIf (NeuIndex% = 0) Then
        If (SonderBelegRezept% > 0) Then
            NeuIndex% = 54
        Else
            NeuIndex% = 31
        End If
    ElseIf (NeuIndex% = 54) Then
        NeuIndex% = 31
    ElseIf (NeuIndex% = 31) Then
        NeuIndex% = MaxPruefBox%
    ElseIf (NeuIndex% <= 2) Then
        NeuIndex% = 0
    ElseIf (NeuIndex% <= 6) Then
        NeuIndex% = 2
    Else
        NeuIndex% = NeuIndex% - 4
    End If
ElseIf (KeyCode% = vbKeyPageUp) Then
    If (RezeptHolen(-1)) Then
        Call PaintRezept
    Else
        Beep
    End If
ElseIf (KeyCode% = vbKeyPageDown) Then
    If (RezeptHolen(1)) Then
        Call PaintRezept
    Else
        Beep
    End If
End If


If (NeuIndex <> AktTabIndex%) Then Call TabIndexChange(NeuIndex%)

Call DefErrPop
End Sub

Function MouseDown%(x!, y!)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("MouseDown%")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
Dim i%, NeuIndex%, ret%, bis%

ret% = False
NeuIndex = AktTabIndex%

bis% = UBound(RezeptBoxen)
If (SonderBelegRezept% = 0) Then
    bis% = 52
End If

For i% = 0 To bis%
    If (i% <> 37) And (i% <> 38) Then
        With RezeptBoxen(i%)
            If (x! >= .LeftX) And (x! <= .RightX) And (y! >= .LeftY) And (y! <= .RightY) Then
                If (.typ2) Or ((PrivatRezept% = 0) And (SonderBelegRezept% = 0)) Then
                    If (Selbsterklaerung) And (InStr(SelbstErklaerungTabstop, "," + CStr(i) + ",") <= 0) Then
                    Else
                        NeuIndex% = i%
                    End If
                End If
                ret% = True
                Exit For
            End If
        End With
    End If
Next i%

If (ret% = 0) Then
    i% = 59
    With RezeptBoxen(i%)
        If (x! >= .LeftX) And (x! <= .RightX) And (y! >= .LeftY) And (y! <= .RightY) Then
            If (.typ2) Or ((PrivatRezept% = 0) And (SonderBelegRezept% = 0)) Then
                NeuIndex% = i%
            End If
            ret% = True
        End If
    End With
End If

If (Chef = False) Then
    If (NeuIndex <> AktTabIndex%) Then
        Call TabIndexChange(NeuIndex%)
    End If
    MouseDown = False
    Call DefErrPop: Exit Function
End If

If (NeuIndex% = 32) Then    'frei
    ret% = False
    Call PaintRezeptBox(33)
    Call CheckRezeptBox(32)
    GebFrei% = True
    Call CalcRezeptWerte
ElseIf (NeuIndex% = 33) Then    'pflichtig
    ret% = False
    Call PaintRezeptBox(32)
    Call CheckRezeptBox(33)
    GebFrei% = False
    Call CalcRezeptWerte
ElseIf (NeuIndex% >= 43) And (NeuIndex% <= 45) Then
    NeuIndex% = 3 + ((NeuIndex% - 43) * 4)
    Call TabIndexChange(NeuIndex%)
    Call ImporteOrAutIdem(False)
ElseIf (NeuIndex% >= 50) And (NeuIndex% <= 52) And (Selbsterklaerung = 0) Then
    NeuIndex% = 3 + ((NeuIndex% - 50) * 4)
    Call TabIndexChange(NeuIndex%)
    Call ImporteOrAutIdem(True)
ElseIf (NeuIndex <> AktTabIndex%) Then
    Call TabIndexChange(NeuIndex%)
End If

MouseDown% = ret%

Call DefErrPop
End Function

Sub TabIndexChange(NeuIndex%, Optional AuchAlterIndex% = True)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("TabIndexChange")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:

Dim lColor As Long
lColor = vbHighlight
lColor = RGB(0, 0, 200)

If (AuchAlterIndex%) Then
    With RezeptBoxen(AktTabIndex%)
        ParentForm.picRezept.DrawMode = 10
        ParentForm.picRezept.DrawWidth = 2
'        If (AktTabIndex% <= 52) Then
            ParentForm.picRezept.Line (.LeftX + 15, .LeftY + 15)-(.RightX - 15, .RightY - 15), lColor, BF
'        Else
'            ParentForm.picRezept.Line (.LeftX + 75, .LeftY + 75)-(.RightX - 60, .RightY - 60), vbHighlight, BF
'        End If
        ParentForm.picRezept.DrawWidth = 1
        ParentForm.picRezept.DrawMode = 13
    End With
End If
    
AktTabIndex% = NeuIndex%

With RezeptBoxen(AktTabIndex%)
    ParentForm.picRezept.DrawMode = 10
    ParentForm.picRezept.DrawWidth = 2
    ParentForm.picRezept.Line (.LeftX + 15, .LeftY + 15)-(.RightX - 15, .RightY - 15), lColor, BF
    ParentForm.picRezept.DrawWidth = 1
    ParentForm.picRezept.DrawMode = 13
End With

Call EchtKurzInfo
    
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
Dim i%, l%, ind%, lInd%, rInd%, EditCol%, aRow%, m%, aMeng%, iLief%, IstBesorger%
Dim j%, hWg%, BlinkFlag%, erg%, iActVerordnung%, ABG_HANDLE%
Dim iAdd%, iAdd2%
Dim spBreite&
Dim dPreis#, dZuz#, dZuzDos#, dZuzPreis#
Dim h$, h2$, pzn$, txt$, mErg$, SQLStr$, ch$
Dim c As Control

Dim row%

If (Chef = False) Then
    Call DefErrPop: Exit Sub
End If
            
frmEdit.txtEdit.MaxLength = 0

If (AutIdemSonderregelOk%) And (AktTabIndex% = 59) Then
    frmKkMatchcode.Show 1
    Call TabIndexChange(AktTabIndex%, False)
    Call DefErrPop: Exit Sub
End If


If (AktTabIndex% = 54) Then
    EditCol% = 8
ElseIf (AktTabIndex% = 31) Then
    EditCol% = 7
ElseIf (AktTabIndex% = 52) Then
    EditCol% = 99
ElseIf (AktTabIndex% >= 3) Then
    row% = (AktTabIndex% - 3) \ 4
    EditCol% = (AktTabIndex - 3) Mod 4
    
    If (row% > AnzRezeptArtikel%) Then
        Beep
        Call DefErrPop: Exit Sub
    End If
Else
    EditCol% = AktTabIndex% + 4
End If

If (EditCol% = 0) Then
    erg% = False
    If (ImpfRez%) Then
        EditErg% = 0
        If (Len(EingabeStr$) = 8) Then
            For i% = 0 To UBound(Impfungen)
                If (EingabeStr$ = Impfungen(i%).pzn) Then
                    EditErg% = True
                    EditTxt$ = Impfungen(i%).text + vbTab + Impfungen(i%).pzn
                    Exit For
                End If
            Next i%
        End If
        
        If (EditErg% = 0) Then
            With frmEdit.flxEdit
                .Left = 0
                .Top = 0
                
                .Rows = 2
                .FixedRows = 1
                .FormatString = "<Name|<Pzn"
                .Rows = 1
                .Cols = 2
                
                .ColWidth(0) = frmEdit.TextWidth(String(28, "X"))
                .ColWidth(1) = frmEdit.TextWidth("9999999.99")
                
                For i% = 0 To UBound(Impfungen)
                    .AddItem Impfungen(i%).text + vbTab + Impfungen(i%).pzn
                Next i%
                
                .Height = .RowHeight(0) * .Rows + 90
                
                spBreite& = 0
                For i% = 0 To (.Cols - 1)
                    spBreite& = spBreite& + .ColWidth(i%)
                Next i%
                .Width = spBreite& + 90
                
                If (para.Newline) Then
                    .ScrollBars = flexScrollBarNone
                    .BorderStyle = 0
                    .Width = .Width - 90
                    .Height = .Height - 90
                    .GridLines = flexGridFlat
                    .GridLinesFixed = flexGridFlat
                    .GridColorFixed = .GridColor
                    .BackColor = wpara.nlFlexBackColor    'vbWhite
                    .BackColorBkg = wpara.nlFlexBackColor    'vbWhite
                    .BackColorFixed = wpara.nlFlexBackColorFixed   ' RGB(199, 176, 123)
                    .BackColorSel = wpara.nlFlexBackColorSel  ' RGB(232, 217, 172)
                    .ForeColorSel = vbBlack
                    
'                    iAdd = wpara.NlFlexBackY
'                    .Left = .Left + iAdd
'                    .Top = .Top + iAdd
'                    frmEdit.Height = .Height + 2 * iAdd
'                    frmEdit.Width = .Width + 2 * iAdd
                End If
        
                frmEdit.Height = .Height '+ 60
                frmEdit.Width = .Width '+ 60
                
                frmEdit.Left = ParentForm.Left + (ParentForm.Width - frmEdit.Width) / 2
                frmEdit.Top = ParentForm.Top + (ParentForm.Height - frmEdit.Height) / 2
            
                
                .row = 1
                .col = 0
                .ColSel = .Cols - 1
                .Visible = True
            End With
            
            frmEdit.Show 1
        End If
        
        If (EditErg%) Then
            ind% = InStr(EditTxt$, vbTab)
            txt$ = Left$(EditTxt$, ind% - 1)
            pzn$ = Mid$(EditTxt$, ind% + 1)
            
            Call InitRezeptArtikel(row%)
            RezArtikel(row%).pzn = pzn$
            RezArtikel(row%).text = txt$
            RezArtikel(row%).Faktor = 1
            
            If (row% >= AnzRezeptArtikel%) Then
                AnzRezeptArtikel% = row% + 1
            End If
            
            erg% = True
        End If
    ElseIf (PrivatRezept% = 0) And (SonderBelegRezept% = 0) And (RezArtikel(row%).flag And REZ_WG_4) And (RezArtikel(row%).flag And REZ_BLINK) Then
        erg% = Verbandmittel%
    Else
'        If (Len(EingabeStr$) = 8) Then EingabeStr$ = Left$(EingabeStr$, 7)
'        If (Len(EingabeStr$) = 7) Then
'            mErg$ = EingabeStr$
    
        If (Len(EingabeStr$) = 8) Then
            mErg$ = EingabeStr$
        Else
            txt$ = Trim(RezArtikel(row%).text)
            txt$ = Left$(txt$, 25)
            If (Left$(txt$, 13) = "KONTROLL-STOP") Then txt$ = ""
            If (UCase(Left$(txt$, 8)) = "REZEPTUR") Then txt$ = ""
            
            If (txt <> "") Then
                pzn$ = RezArtikel(row%).pzn
                If (Left(pzn, 1) = "-") Then
                    pzn = Mid(pzn, 2) + "0"
                    pzn = MakePzn(pzn)
                End If
                
                If (Val(pzn) > 0) Then
                    SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + pzn
                    'Set TaxeRec = TaxeDB.OpenRecordset(SQLStr$)
                    On Error Resume Next
                    TaxeRec.Close
                    Err.Clear
                    On Error GoTo DefErr
                    TaxeRec.Open SQLStr, taxeAdoDB.ActiveConn
                    If Not (TaxeRec.EOF) Then
                        txt = UCase(Trim(TaxeRec!Name)) '+ TaxeRec!menge + TaxeRec!einheit
                    End If
                End If
            
                h$ = txt$
                txt$ = ""
                l% = Len(h$)
                For i% = 1 To l%
                    ch$ = Mid$(h$, i%, 1)
                    If ((ch$ >= "A") And (ch$ <= "Z")) Or ((ch$ >= "0") And (ch$ <= "9")) Then
                        txt$ = txt$ + ch$
                    End If
                Next i%
            End If
            
            If (txt$ <> "") Then
'                txt$ = "PZN" + RezArtikel(row%).pzn + txt$
                txt$ = "PZN" + pzn + txt$
                txt = Left(txt, 19)
            End If
            
            mErg$ = Matchcode(0, pzn$, txt$, (txt$ <> ""), True)
        End If
        
        If (mErg$ <> "") Then
'            ind% = InStr(mErg$, "@")
'            pzn$ = Left$(mErg$, ind% - 1)
            pzn$ = mErg$

            Call NeuerArtikel(row%, pzn$)
            For i = 0 To 1
                If (RezArtikel(row%).flag And REZ_WG_4) Then erg% = Verbandmittel%
                If (erg = 0) And ((RezArtikel(row).VebNr > 0) Or (RezArtikel(row).PauschaleNr > 0)) Then
                    ActVebNr = 0
                    ActPauschaleNr = 0
                    RezArtikel(row).VebNr = 0
                    RezArtikel(row).PauschaleNr = 0
                Else
                    Exit For
                End If
            Next i
            Call CheckVerordnungMDB
            If (RezArtikel(row%).Preis = 0#) And (KeinPreisDa%) Then
                Call iMsgBox("Kein Preis vorhanden !")
            End If
            erg% = True
        End If
    End If
    
    If (erg%) Then
        Call PaintArtikel(row%)
        If (row% < AnzRezeptArtikel%) And (row% < 2) Then
            Call TabIndexChange(AktTabIndex% + 4, True)
        End If
    End If
ElseIf (EditCol% = 1) Then
'    If (Left$(RezArtikel(row%).pzn, 1) = "@") Then
    If (Left$(RezArtikel(row%).pzn, 1) = "-") Then
        PznAuswahlWert$ = RezArtikel(row%).ScreenPzn
    Else
        PznAuswahlWert$ = RezArtikel(row%).pzn
    End If
    frmPznAuswahl.Show 1
    
    If (PznAuswahlName$ <> "") Then
        If (PznAuswahlWert$ = "HMEingabe") Then
            EditModus% = 0
                        
            Load frmEdit
            
            With frmEdit
                .Left = ParentForm.picBack(0).Left + ParentForm.picRezept.Left + RezeptBoxen(AktTabIndex%).LeftX + 45
                .Left = .Left + ParentForm.Left + wpara.FrmBorderHeight
                .Top = ParentForm.picBack(0).Top + ParentForm.picRezept.Top + RezeptBoxen(AktTabIndex%).LeftY + 45
                .Top = .Top + ParentForm.Top + wpara.FrmBorderHeight + wpara.FrmCaptionHeight + wpara.FrmMenuHeight
                .Width = RezeptBoxen(AktTabIndex%).RightX - RezeptBoxen(AktTabIndex%).LeftX
                .Height = RezeptBoxen(AktTabIndex%).RightY - RezeptBoxen(AktTabIndex%).LeftY
            End With
            With frmEdit.txtEdit
                .Width = frmEdit.ScaleWidth
'                .Height = frmEdit.Height
                .Left = 0
                .Top = 0
                    
                h2$ = ""
                .MaxLength = 10
                
                .text = h2$
                .BackColor = vbWhite
                .Visible = True
            End With
            
            frmEdit.Show 1
            
            h$ = ""
            If (EditErg%) Then
                h$ = Trim(EditTxt$)
            End If
            If (h$ = "") Then
                Call DefErrPop: Exit Sub
            End If
            
            PznAuswahlWert$ = h$
        ElseIf (Val(PznAuswahlWert) = 2567053) Or (Val(PznAuswahlWert) = 2566993) Then
            frmAusEinzelung.Show 1
            RezArtikel(row%).Faktor = 1
            RezArtikel(row%).Preis = AuseinzelungPreisGesamt
            RezArtikel(row%).Zuz = CalcZuzaNeu(AuseinzelungPreisGesamt)
        ElseIf (Val(PznAuswahlWert) = 6460688) Then
            RezArtikel(row%).Preis = 4.26   '2.91
        End If
        
        RezArtikel(row%).text = PznAuswahlName$
        
        pzn$ = PznAuswahlWert$
        l% = Len(pzn$)
        If (l% > 8) Then
'            pzn$ = "@" + Chr$(l%) + Ascii2bcd(pzn$, 10)
            pzn$ = "-9999999"
            RezArtikel(row%).ScreenPzn = PznAuswahlWert$
        End If
        
        If (Val(pzn$) = 9999063) And (VmFlag%) Then
            iActVerordnung% = ActVerordnung%
            ActVerordnung% = 6
            erg% = Verbandmittel
            ActVerordnung% = iActVerordnung%
        End If
        
        RezArtikel(row%).pzn = pzn$
        
        If (RezArtikel(row%).Faktor = 0) Then RezArtikel(row%).Faktor = 1
        
        j% = 3 + row% * 4
        For i% = 0 To 3
            Call PaintRezeptBox(j% + i%)
        Next i%
        
        RezArtikel(row%).flag = RezArtikel(row%).flag And (REZ_BLINK Xor &HFF)
        
        BlinkFlag% = RezArtikel(row%).flag And REZ_BLINK
        Call PaintRezeptWert(j%, RezArtikel(row%).text, BlinkFlag%)
        h$ = RezArtikel(row%).pzn
'        If (Left$(h$, 1) = "@") Then h$ = RezArtikel(row%).ScreenPzn
        If (Left$(h$, 1) = "-") Then h$ = RezArtikel(row%).ScreenPzn
        Call PaintRezeptWert(j% + 1, h$, BlinkFlag%)
        Call PaintRezeptWert(j% + 2, Format(RezArtikel(row%).Faktor, "0"), BlinkFlag%)
        Call PaintRezeptWert(j% + 3, Format(RezArtikel(row%).Preis * 100 * RezArtikel(row%).Faktor, "0"), BlinkFlag%)
        
        If (row% >= AnzRezeptArtikel%) Then
            AnzRezeptArtikel% = row% + 1
        End If
            
'        If (Val(pzn$) = 9999063) Then
'            AnzRezeptArtikel% = AnzRezeptArtikel% + 1
'            row% = AnzRezeptArtikel% - 1
'            Call InitRezeptArtikel(row%)
'            RezArtikel(row%).text = BtmName$(ind%)
'            RezArtikel(row%).pzn = "02567001" ' "8150012"
'            RezArtikel(row%).Faktor = 1
'            RezArtikel(row%).Preis = BtmKosten#(ind%)
'        End If
    
        
        
        Call CalcRezeptWerte
        
        Call TabIndexChange(AktTabIndex%, False)

''''''''''''
        Dim iTabIndex%
        If (CheckBtmSonderPzn%(pzn) > 0) Then
            BtmRezept% = True
            iTabIndex% = AktTabIndex%
            Call PruefeBtmKosten
            sBrutto$ = RezAddGesamtBrutto$(0)
            Call PaintRezept
            Call RepaintBtmGebühr
            Call TabIndexChange(iTabIndex%, True)
        End If
''''''''''''
    End If
ElseIf (EditCol% = 6) Then
    With frmGesamtBrutto.flxEdit
        .Left = 0
        .Top = 0
        
        .Rows = 2
        .FixedRows = 1
        .FormatString = "<Abgabeart|>Preis"
        .Rows = 1
        .Cols = 2
        
        .ColWidth(0) = frmGesamtBrutto.TextWidth(String(30, "X"))
        .ColWidth(1) = frmGesamtBrutto.TextWidth("99999999.99")
        
        For i% = 0 To UBound(AbgabeKosten)
            dPreis# = AbgabeKosten(i%).kosten
            If (dPreis# = 0#) Then
                h2$ = ""
            Else
                h2$ = Format(dPreis#, "0.00")
            End If
            .AddItem AbgabeKosten(i%).Name + vbTab + h2$
        Next i%
        
        .Height = .RowHeight(0) * .Rows + 90
        
        spBreite& = 0
        For i% = 0 To (.Cols - 1)
            spBreite& = spBreite& + .ColWidth(i%)
        Next i%
        .Width = spBreite& + 90
        
        frmGesamtBrutto.Height = .Height
        frmGesamtBrutto.Width = .Width
        frmGesamtBrutto.Left = ParentForm.Left + (ParentForm.Width - frmGesamtBrutto.Width) / 2
        frmGesamtBrutto.Top = ParentForm.Top + (ParentForm.Height - frmGesamtBrutto.Height) / 2
    
        
        .row = 1
        If (.Rows > 2) Then .row = 2
        .col = 0
'        .ColSel = .Cols - 1
        .Visible = True
    End With
    
    If (para.Newline) Then
        iAdd = wpara.NlFlexBackY
        iAdd2 = wpara.NlCaptionY
        
        With frmGesamtBrutto.flxEdit
            .ScrollBars = flexScrollBarNone
            .BorderStyle = 0
            .Width = .Width - 90
            .Height = .Height - 90
            .GridLines = flexGridFlat
            .GridLinesFixed = flexGridFlat
            .GridColorFixed = .GridColor
            .BackColor = wpara.nlFlexBackColor    'vbWhite
            .BackColorBkg = wpara.nlFlexBackColor    'vbWhite
            .BackColorFixed = wpara.nlFlexBackColorFixed   ' RGB(199, 176, 123)
            .BackColorSel = wpara.nlFlexBackColorSel  ' RGB(232, 217, 172)
            .ForeColorSel = vbBlack
            
            .Left = .Left + 2 * iAdd
            .Top = .Top + 2 * iAdd
        End With
        
        With frmGesamtBrutto
            .Width = .Width + 4 * iAdd
            .Height = .Height + 4 * iAdd
    
            .flxEdit.Top = .flxEdit.Top + iAdd2
            
            With .nlcmdOk
                .Init
                .Left = (frmGesamtBrutto.ScaleWidth - (.Width * 2 + 300)) / 2
                .Top = frmGesamtBrutto.flxEdit.Top + frmGesamtBrutto.flxEdit.Height + iAdd + 600
                .Caption = frmGesamtBrutto.cmdOk.Caption
                .TabIndex = frmGesamtBrutto.cmdOk.TabIndex
                .Enabled = frmGesamtBrutto.cmdOk.Enabled
                .Default = frmGesamtBrutto.cmdOk.Default
                .Cancel = frmGesamtBrutto.cmdOk.Cancel
                .Visible = True
            End With
            .cmdOk.Visible = False
        
            With .nlcmdEsc
                .Init
                .Left = frmGesamtBrutto.nlcmdOk.Left + .Width + 300
                .Top = frmGesamtBrutto.nlcmdOk.Top
                .Caption = frmGesamtBrutto.cmdEsc.Caption
                .TabIndex = frmGesamtBrutto.cmdEsc.TabIndex
                .Enabled = frmGesamtBrutto.cmdEsc.Enabled
                .Default = frmGesamtBrutto.cmdEsc.Default
                .Cancel = frmGesamtBrutto.cmdEsc.Cancel
                .Visible = True
            End With
            .cmdEsc.Visible = False

            .Height = .nlcmdOk.Top + .nlcmdOk.Height + wpara.FrmCaptionHeight + 450

            Call wpara.NewLineWindow(frmGesamtBrutto, .nlcmdOk.Top)
            RoundRect .hdc, (.flxEdit.Left - iAdd) / Screen.TwipsPerPixelX, (.flxEdit.Top - iAdd) / Screen.TwipsPerPixelY, (.flxEdit.Left + .flxEdit.Width + iAdd) / Screen.TwipsPerPixelX, (.flxEdit.Top + .flxEdit.Height + iAdd) / Screen.TwipsPerPixelY, 20, 20
        End With
    Else
        frmGesamtBrutto.nlcmdOk.Visible = False
        frmGesamtBrutto.nlcmdEsc.Visible = False
    End If
    
    frmGesamtBrutto.Show 1
    
    If (EditErg%) Then
        ind% = InStr(EditTxt$, vbTab)
        txt$ = Left$(EditTxt$, ind% - 1)
        
'        If (UCase(Left$(txt$, 10)) = "ZUSCHLAG=0") Then
        If (UCase(Left$(txt$, 12)) = "ZURÜCKSETZEN") Then
            RezSummeDazu# = 0#
            
            Call CalcRezeptWerte
            Call TabIndexChange(AktTabIndex%, False)
            
            For i% = 0 To UBound(RezAddGesamtBrutto$)
                RezAddGesamtBrutto$(i%) = ""
            Next i%
            Call EchtKurzInfo
        Else
            h2$ = Mid$(EditTxt$, ind% + 1)
            If (h2$ = "") Then
                dPreis# = 0#
            Else
                dPreis# = CDbl(h2$)
            End If
            
            For i% = 0 To UBound(AbgabeKosten)
                With AbgabeKosten(i%)
                    If (.Name = txt$) Then
                        If (.kosten <> dPreis#) Then
                            .kosten = dPreis#
                                
                            ABG_HANDLE% = FileOpen("abgpr.dat", "RW", "B")
                            If (ABG_HANDLE% > 0) Then
                                h$ = .Name + Right$(Space$(4) + Format(dPreis# * 100#, "0"), 4)
                                Put #ABG_HANDLE%, (.ind * 16) + 1, h$
                                Close #ABG_HANDLE%
                            End If
                        End If
                        
                        Exit For
                    End If
                End With
            Next i%

        
            
            pzn$ = ""
            If (BtmAlsZeile%) Then
                If (Left$(txt$, 5) = "Besch") Then
                    pzn$ = "09999637"    '"8150006"
                ElseIf (Left$(txt$, 3) = "BTM") Then
                    pzn$ = "02567001"      ' "8150012"
                End If
            End If
            
            If (pzn$ <> "") Then
                AnzRezeptArtikel% = AnzRezeptArtikel% + 1
                row% = AnzRezeptArtikel% - 1
                Call InitRezeptArtikel(row%)
                RezArtikel(row%).text = txt$
                RezArtikel(row%).pzn = pzn$
                RezArtikel(row%).Faktor = 1
                RezArtikel(row%).Preis = dPreis#
                Call PaintRezept
                Call TabIndexChange(2, True)
            Else
                RezSummeDazu# = RezSummeDazu# + dPreis#
                
                Call CalcRezeptWerte
                Call TabIndexChange(AktTabIndex%, False)
                
                For i% = 0 To UBound(RezAddGesamtBrutto$)
                    If (RezAddGesamtBrutto$(i%) = "") Then
                        RezAddGesamtBrutto$(i%) = txt$ + Format(dPreis#, "   (0.00)")
                        Call EchtKurzInfo
                        Exit For
                    End If
                Next i%
            End If
        End If
    End If

Else
    EditModus% = 1
                
    Load frmEdit
    
    With frmEdit
        .Left = ParentForm.picBack(0).Left + ParentForm.picRezept.Left + RezeptBoxen(AktTabIndex%).LeftX + 45
        .Left = .Left + ParentForm.Left + wpara.FrmBorderHeight
        .Top = ParentForm.picBack(0).Top + ParentForm.picRezept.Top + RezeptBoxen(AktTabIndex%).LeftY + 45
        .Top = .Top + ParentForm.Top + wpara.FrmBorderHeight + wpara.FrmCaptionHeight + wpara.FrmMenuHeight
        .Width = RezeptBoxen(AktTabIndex%).RightX - RezeptBoxen(AktTabIndex%).LeftX
        .Height = RezeptBoxen(AktTabIndex%).RightY - RezeptBoxen(AktTabIndex%).LeftY
    End With
    With frmEdit.txtEdit
        .Width = frmEdit.ScaleWidth
    '            .Height = frmEdit.ScaleHeight
        .Left = 0
        .Top = 0
            
        If (EditCol% = 2) Then
            h2$ = RezArtikel(row%).Faktor
        ElseIf (EditCol% = 3) Then
            h2$ = Format(RezArtikel(row%).Preis * 100#, "0")
        ElseIf (EditCol% = 5) Then
            h2$ = Format(RezGebSumme# * 100#, "0")
        ElseIf (EditCol% = 4) Then
            h2$ = RezApoNr$
        ElseIf (EditCol% = 7) Then
            EditModus% = 10
            h2$ = VerkDatum$
        ElseIf (EditCol% = 8) Then
            EditModus% = 0
            h2$ = RezKundenInfo$(3)
            .MaxLength = 12
        ElseIf (EditCol% = 99) Then
            EditModus% = 10
            h2$ = BoxenWert(52)
        End If
        
'        h2$ = BoxenWert$(AktTabIndex%)
        .text = h2$
        .BackColor = vbWhite
        .Visible = True
    End With
    
    frmEdit.Show 1
    
    If (EditErg%) Then
        h$ = Trim(EditTxt$)
        
        If (h$ <> BoxenWert$(AktTabIndex%)) Then
            If (EditCol% = 2) Or (EditCol% = 3) Then
                If (EditCol% = 2) Then
                    If (Val(h$) >= 0) And (Val(h$) <= 9999) Then RezArtikel(row%).Faktor = Val(h$)
                    If (RezArtikel(row%).flag And REZ_WG_4) Then
                        erg% = Verbandmittel%
                    ElseIf (RezArtikel(row%).IstWg4) Then
                        SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + RezArtikel(row%).pzn
                        'Set TaxeRec = TaxeDB.OpenRecordset(SQLStr$)
                        On Error Resume Next
                        TaxeRec.Close
                        Err.Clear
                        On Error GoTo DefErr
                        TaxeRec.Open SQLStr, taxeAdoDB.ActiveConn
                        If (TaxeRec.EOF = False) Then
                            If (TaxeRec!ZuzFrei = 0) Then
                                dZuzPreis# = RezArtikel(row%).Preis * RezArtikel(row%).Faktor
                                dZuz# = CalcZuzaNeu(dZuzPreis#, TaxeRec!HimiVerbrauch, RezArtikel(row%).Faktor)
                                RezArtikel(row%).Zuz = dZuz#
                            End If
                        End If
                    End If
                    If (UCase(Left$(RezArtikel(row%).text, 3)) = "BTM") Then
                        BtmFaktorManuell = True
                    End If
                Else
                    RezArtikel(row%).Preis = Val(h$) / 100
                    RezArtikel(row%).flag = RezArtikel(row%).flag And (REZ_BLINK Xor &HFF)
                End If
            
                j% = 3 + row% * 4
                For i% = 0 To 3
                    Call PaintRezeptBox(j% + i%)
                Next i%
                BlinkFlag% = RezArtikel(row%).flag And REZ_BLINK
                Call PaintRezeptWert(j%, RezArtikel(row%).text, BlinkFlag%)
                h$ = RezArtikel(row%).pzn
'                If (Left$(h$, 1) = "@") Then h$ = RezArtikel(row%).ScreenPzn
                If (Left$(h$, 1) = "-") Then h$ = RezArtikel(row%).ScreenPzn
                Call PaintRezeptWert(j% + 1, h$, BlinkFlag%)
                Call PaintRezeptWert(j% + 2, Format(RezArtikel(row%).Faktor, "0"), BlinkFlag%)
                If (RezArtikel(row).VdbPauschale) Then
                    Call PaintRezeptWert(j% + 3, Format(RezArtikel(row%).Preis * 100, "0"), BlinkFlag%)
                Else
                    Call PaintRezeptWert(j% + 3, Format(RezArtikel(row%).Preis * 100 * RezArtikel(row%).Faktor, "0"), BlinkFlag%)
                End If
                
                Call TabIndexChange(AktTabIndex%, False)
                
                Call CalcRezeptWerte
                Call CheckVerordnungMDB
            ElseIf (EditCol% = 5) Then
                If (Val(h$) >= 0) And (Val(h$) <= 999999) Then RezGebSumme# = Val(h$) / 100
                RezGebRechnen% = False
                
                Call CalcRezeptWerte
                Call TabIndexChange(AktTabIndex%, False)
        
'                Call PaintRezeptBox(AktTabIndex%)
'                Call PaintRezeptWert(AktTabIndex%, Format(RezGebSumme# * 100, "0"))
'                Call TabIndexChange(AktTabIndex%, False)
            ElseIf (EditCol% = 4) Then
                RezApoNr$ = Right$(Space$(7) + h$, 7)
                ParentForm.mnuBearbeitenInd(MENU_F6).Enabled = (Trim(RezApoNr$) <> "")
                ParentForm.cmdToolbar(MENU_F6 + 1).Enabled = (Trim(RezApoNr$) <> "")
                Call ParentForm.ShowToolbar
                Call PaintRezeptBox(AktTabIndex%)
                Call PaintRezeptWert(AktTabIndex%, RezApoNr$)
                If (SonderBelegRezept% > 0) Then
                    Call PaintRezeptBox(56)
                    Call PaintRezeptWert(56, h$, False)
                End If
                Call TabIndexChange(AktTabIndex%, False)
            ElseIf (EditCol% = 7) Then
                VerkDatum$ = h$
                Call PaintRezeptDaten
                Call PaintRezeptBox(AktTabIndex%)
                Call PaintRezeptWert(AktTabIndex%, h$, False)
                If (SonderBelegRezept% > 0) Then
                    Call PaintRezeptBox(58)
                    Call PaintRezeptWert(58, h$, False)
                End If
                Call TabIndexChange(AktTabIndex%, False)
            ElseIf (EditCol% = 8) Then
                RezKundenInfo$(3) = h$
                Call PaintRezeptDaten
                Call PaintRezeptBox(AktTabIndex%)
                Call PaintRezeptWert(AktTabIndex%, h$, False)
                Call TabIndexChange(AktTabIndex%, False)
            ElseIf (EditCol% = 99) Then
                BoxenWert(52) = h$
                Call PaintRezeptDaten
                
                With RezeptBoxen(AktTabIndex)
                    ParentForm.picRezept.ForeColor = vbWhite
                    ParentForm.picRezept.Line (.LeftX, .LeftY)-(.RightX, .RightY), vbWhite, BF
                    ParentForm.picRezept.ForeColor = vbBlack
                End With
'                Call PaintRezeptBox(AktTabIndex%)
                
                Call PaintRezeptWert(AktTabIndex%, h$, False)
                Call TabIndexChange(AktTabIndex%, False)
            End If
        End If
    End If

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

'Call AuslesenBestellung(True, False, True)

Call DefErrPop
End Sub

Sub MenuBearbeiten(index As Integer)
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
Dim i%, row%, col%, ret%, iTabIndex%, rInd%, erg%
Dim l&
Dim h$, pzn$, txt$
Dim IstSek, StartSek

Select Case index

    Case MENU_F2
        Call InsertArtikel
        
    Case MENU_F4
        Call ZeigeInhaltsStoffe
        
    Case MENU_F5
        Call LoescheArtikel
        
    Case MENU_F6
        If (MarsModus = MARS_REZEPT_KONTROLLE) Then
            Call frmAction.MarsSaveRezeptAnmerkungen
            Call WriteRezeptSpeicher
            If (MarsAutomaticModus) Then
                Call frmAction.MarsNextRezept
            Else
                ParentForm.cmdEsc(0).Value = True
            End If
        Else
            Call DruckeRezept
        End If

    
    Case MENU_F7
'        Call AuswahlF7Kassen
    
    Case MENU_F8
'        With ParentForm.flxarbeit(0)
'            If (.TextMatrix(.row, 15) = "v") And (InStr(para.Benutz, "Y") > 0) Then
'                Call ZeigeBesorgerDaten
'            Else
'                Call ZusatzFenster(ZUSATZ_ARTIKEL, KorrPzn$, KorrTxt$)
'            End If
'        End With
    
    Case MENU_SF2
        If (eRezeptTaskId <> "") Then
            If (FDok = 0) Then
                Set FD_OP = New TI_Back.Fachdienst_OP
                If (FD_OP.Init(FD_OP.NetWindowHandle(ParentForm.hWnd))) Then
                    FDok = True
                End If
            End If
        
            AMverfügbarkeit$ = CheckVerfügbarkeit
'            If (Val(RezArtikel(0).pzn) <> Val(eRezeptPzn)) Or (eRezeptAMverfügbarkeit <> Mid(AMverfügbarkeit, 1, 1)) Then
'                Call eRezeptSpeichern
'            End If
            Dim eRezeptGeändert%
            Dim eRezeptStatusWerte2$
            Dim eRezeptOrg As TI_Back.eRezept_OP
            
            Dim lPzn&
            For i% = 0 To (AnzRezeptArtikel% - 1)
                lPzn = Val(RezArtikel(i%).pzn)
                If (lPzn = 9999637) Or (lPzn = 6461110) Then  'Beschaffungskosten 'Botendienst
                    eRezeptSonderPZN = True
                    Exit For
                End If
            Next i

            
            eRezeptGeändert = 0
            eRezeptStatusWerte2 = RezArtikel(0).pzn + Mid(AMverfügbarkeit, 1, 1) + Format(RezArtikel(0).Preis, "0.00") + Format(RezArtikel(0).Zuz, "0.00") + VerkDatum
            'MsgBox (eRezeptStatusWerte + vbCrLf + eRezeptStatusWerte2)
            If (eRezeptStatusWerte <> eRezeptStatusWerte2) Or (eRezeptSonderPZN) Then
                eRezeptGeändert = True
                Set eRezeptOrg = New TI_Back.eRezept_OP
                Call eRezeptOrg.New2(eRezeptTaskId)
                Call eRezeptSpeichern
            End If
            
            frm_eRezeptEdit.Show 1
            If (eRezeptGespeichert) Then
                eRezeptStatusWerte = eRezeptStatusWerte2
            ElseIf (eRezeptGespeichert = 0) And (eRezeptGeändert) Then
                Call eRezeptRestore(eRezeptOrg)
            End If
            
'            wbVerordnung.Visible = Not (wbVerordnung.Visible)
'            wbAbgabedaten.Visible = Not (wbAbgabedaten.Visible)
'            Call EchtKurzInfo
        Else
            Call InitRezept
            GebFrei% = True
            ImpfRez% = True
            ActVerordnung% = 2
            Call PaintRezept
        End If
        
        
    Case MENU_SF3
'        frmBarVerkauf.Show 1
        Call InitRezept
        PrivatRezept% = True
        GebFrei% = True
        Call PaintRezept
    
    Case MENU_SF4
        iTabIndex% = AktTabIndex%
        frmAvAuswahl.Show 1
        Call PaintAvParameter
        With ParentForm
            If (.picRezept.Visible) Then
                .picRezept.SetFocus
                For i% = 0 To (AnzRezeptArtikel% - 1)
                    If (RezArtikel(i%).flag And REZ_WG_4) Then
                        AktTabIndex% = 3 + i% * 4
                        Call Verbandmittel
                    End If
                Next i%
                Call PaintRezept
                Call TabIndexChange(iTabIndex%, True)
            Else
                .txtRezeptNr.SetFocus
            End If
        End With
    
    Case MENU_SF5
        ZeigeStatistik
        
    Case MENU_SF6
        If (eRezeptTaskId <> "") Then
            If (FDok = 0) Then
                Set FD_OP = New TI_Back.Fachdienst_OP
                If (FD_OP.Init(FD_OP.NetWindowHandle(ParentForm.hWnd))) Then
                    FDok = True
                End If
            End If
        
            If (Val(RezArtikel(0).pzn) <> Val(eRezeptPzn)) Or (eRezeptAMverfügbarkeit <> Mid(AMverfügbarkeit, 1, 1)) Then
                Call eRezeptSpeichern
            End If
            
            bEinzelErgebnis = True
        
            Dim bErgebnis As Boolean
            Dim eRezept2 As New TI_Back.eRezept_OP
            Call eRezept2.New2(eRezeptTaskId)
'            If (FD_OP.TI_RezeptAbrechnen(eRezept2, (eRezeptOpStatus = 2))) Then
''                    MsgBox (IIf(SollStatus = 2, "Vorabprüfung: ", "Einreichung: ") + CStr(VorabPruefung(sTaskId, (SollStatus = 2))))
'                bErgebnis = VorabPruefung(eRezeptTaskId, (eRezeptOpStatus = 2))
            If (FD_OP.TI_RezeptAbrechnen(eRezept2, True)) Then
'                    MsgBox (IIf(SollStatus = 2, "Vorabprüfung: ", "Einreichung: ") + CStr(VorabPruefung(sTaskId, (SollStatus = 2))))
                bErgebnis = VorabPruefung(eRezeptTaskId, True)
            Else
                Call MessageBox("Problem beim Erstellen der Abgabedaten", vbCritical)
            End If
        Else
            Call InitRezept
            GebFrei% = True
            ActVerordnung% = 2
            Call PaintRezept
        End If
        
    Case MENU_SF7
        Call InitRezept
        ActVerordnung% = 1
        Call PaintRezept
    
    Case MENU_SF8
        Call IndividuelleRezeptur
        
    Case MENU_SF9
        VmDebugAktiv% = True
        Call Verbandmittel
        VmDebugAktiv% = False
        
End Select

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
'Dim Lm%, BuchM%, rInd%
'
If (Chef = False) Then
    Call DefErrPop: Exit Sub
End If
            
Call MenuBearbeiten(MENU_SF4)
    
'    With flxInfo(0)
'        row% = .row
'        col% = .col
'        h$ = RTrim(.text)
'    End With
'    If (col% = 0) Then
'        Call ActProgram.cmdOkClick(h$)
'    End If

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
'Dim row%, AbholNr%
'Dim h$
'
'With ParentForm.flxarbeit(0)
'    row% = .row
'    h$ = RTrim$(.TextMatrix(row%, 11))
'    AbholNr% = Val(h$)
'End With
'
'Call BesorgerInfo(AbholNr%, KorrPzn$, KorrTxt$)

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
If (para.Newline) Then
    frmnlRezkOptionen.Show 1
Else
    frmRezkOptionen.Show 1
End If
If (OptionenNeu%) Then
    Call PaintAvParameter
    If (ParentForm.picRezept.Visible) Then Call PaintRezept
End If
Call DefErrPop
End Sub

Sub mnuBearbeitenZusatzClick(index As Integer)
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

Select Case index
    Case 0
        frmBarVerkauf.Show 1
    Case 1
'        Call ImporteOrAutIdem(False, False)
        Call ImporteOrAutIdemAlt(False, False)
    Case 2
        Call ImporteOrAutIdem(True)
    Case 3
        Call ZeigeInhaltsStoffe
    Case 4      'zuz.frei
        Call PaintRezeptBox(33)
        Call CheckRezeptBox(32)
        GebFrei% = True
        Call CalcRezeptWerte
    Case 5      'zuz.pfl.
        Call PaintRezeptBox(32)
        Call CheckRezeptBox(33)
        GebFrei% = False
        Call CalcRezeptWerte
    Case 6
        Dim lKuNr&, TaskId&
        lKuNr = Val(RezKundenInfo$(4))
        If (lKuNr > 0) Then
            TaskId& = Shell("WinKunds.exe " + CStr(lKuNr), vbNormalFocus)
        End If
    Case 7
        Call ZeigePlusX
    Case 8
        Dim row%
        Dim sPzn As String
        sPzn = ""
        If (AktTabIndex% >= 3) Then
            row% = (AktTabIndex% - 3) \ 4
            sPzn$ = RezArtikel(row%).pzn
        End If
        TaskId& = Shell("WinMsv3.exe ABSAGEN" + sPzn, vbNormalFocus)
End Select

Call DefErrPop
End Sub

Function RezeptHolen%(Optional PrivatModus% = 0)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("RezeptHolen%")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
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
Dim i%, j%, k%, l%, ind%, PrivGef%, iUser%, iVerkMult%, hhRez%, PrivatBesorger%, IstDrüber%, ind2%
Dim pTag%, pMonat%, pJahr%, PrivDatum%
Dim PrivUser As Byte, PrivLiRe As Byte, pGebuehren As Byte
Dim lAnzVkSaetze&, lPrivLauf&, iKuNr&
Dim dZuz#, dZuzDos#, dPreis#, dAvp#, PrivatBesorgerPreis#
Dim ch$, RezEan$, hStr$, LetztArtikel$, AktArtikel$, wg$, SQLStr$, Unique$, h$, sPrivDatum$
Dim aktavp As Double, OrgAvp As Double
Dim iiDat As Date
Dim AvpRec As Recordset

Static fOrgRezNr!
'Static lPrivEnde&

If (VerkaufDbOk%) Then
    If (VerkaufAdoDB.SqlServerDB) Then
        RezeptHolen = RezeptHolenDB(PrivatModus%)
    Else
        RezeptHolen = RezeptHolenMDB(PrivatModus%)
    End If
    Call DefErrPop: Exit Function
End If

h$ = RezNr$

Call InitRezept(1)

RezNr$ = Trim(h$)
l% = Len(RezNr)

If (UCase$(Left$(RezNr, 1)) = "X") Then
    If (AvpTeilnahme%) Then
        RezNr$ = Mid$(RezNr$, 2, 11)
'        h$ = Mid$(RezNr$, 7, 4)
'        h$ = Mid$(h$, 3, 2) + Left$(h$, 2)
        SQLStr$ = "SELECT * FROM Rezepte WHERE AvpRezeptNr='" + RezNr$ + "'"
        Set AvpRec = RezSpeicherDB.OpenRecordset(SQLStr$)
        If (AvpRec.RecordCount > 0) Then
            RezNr$ = Trim(AvpRec!RezeptNr)
            l% = Len(RezNr)
        End If
    End If
End If

If (UCase$(Left$(RezNr, 1)) = "P") Then
    RezNr$ = UCase$(RezNr$)

    PrivUser = 0
    PrivLiRe = 0
    PrivDatum% = 0
    sPrivDatum$ = ""
    If (l% = 5) Then
        sPrivDatum$ = Mid$(RezNr$, 2)
        For i% = 2 To 5
            If (InStr("0123456789", Mid$(RezNr$, i%, 1)) = 0) Then
                sPrivDatum$ = ""
                Exit For
            End If
        Next i%
        If (sPrivDatum$ <> "") Then
            pTag% = Val(Left$(sPrivDatum$, 2))
            If (pTag% = 0) Or (pTag% > 31) Then
                sPrivDatum$ = ""
            End If
        End If
        If (sPrivDatum$ <> "") Then
            pMonat% = Val(Mid$(sPrivDatum$, 3, 2))
            If (pMonat% = 0) Or (pMonat% > 12) Then
                sPrivDatum$ = ""
            End If
        End If
        If (sPrivDatum$ <> "") Then
            pJahr% = Year(Now) - 2000
            iiDat = Format(pTag, "00") + "." + Format(pMonat%, "00") + "." + Format(Now, "yyyy")
            If (DateDiff("d", iiDat, Now) < 0) Then
                pJahr% = pJahr% - 1
            End If
            h$ = Format(pTag, "00") + Format(pMonat%, "00") + Format(pJahr%, "00")
            PrivDatum% = iDate(h$)
        End If
    ElseIf (l% >= 2) Then
        ch$ = Mid$(RezNr, 2, 1)
        If (InStr("0123456789", ch$) > 0) Then
            PrivUser = Val(ch$)
            If (l% >= 3) Then
                ind% = InStr("LR", Mid$(RezNr, 3, 1))
                If (ind% > 0) Then PrivLiRe = ind%
            End If
        End If
    End If


    OrgVerkPtr& = VerkPtr&
    lAnzVkSaetze& = vk.DateiLen / vk.RecordLen

    Select Case PrivatModus%
        Case 0
            VerkPtr& = lAnzVkSaetze&
        Case -1
            VerkPtr& = VerkPtr& - 1
        Case 1
            VerkPtr& = lPrivEnde& + 1
        Case 2
            'VerkPtr& wurde von aussen gesetzt
    End Select
'
'    Call StartAnimation(frmAction, "Suche nach Privatrezept: ")

    pGebuehren = 0
    lPrivLauf& = 0
    Do
        Call vk.GetRecord(VerkPtr&)

        RezEan$ = vk.RezEan

        If (pGebuehren) Then
            If (vk.EndeT And &H7) And (PrivatModus% <> 1) Then
                VerkPtr& = VerkPtr& + 1
                Exit Do
            End If

            If (vk.gebuehren <> pGebuehren) Or (Left$(RezEan, 13) <> Space$(13)) Then
                VerkPtr& = VerkPtr& + 1
                Exit Do
            End If
        End If

        If (vk.gebuehren > 0) And (pGebuehren = 0) And (Left$(RezEan, 13) = Space$(13)) Then
            PrivGef% = True
            If (PrivDatum% > 0) Then
                If (vk.Datum > PrivDatum%) Then
                    PrivGef% = False
                End If
            ElseIf (PrivUser) Then
                iUser% = vk.User
                If (PrivUser <> iUser%) Then
                    PrivGef% = False
                Else
                    If (PrivLiRe) Then
                        VerkLiRe = vk.bs
                        If (VerkLiRe <> (PrivLiRe - 1)) Then PrivGef% = False
                    End If
                End If
            End If

            If (PrivGef%) Then
                pGebuehren = vk.gebuehren
                If (PrivatModus% = 1) Then Exit Do
            End If
        End If

        If (PrivatModus% = 1) Then
            VerkPtr& = VerkPtr& + 1
            If (VerkPtr& >= lAnzVkSaetze&) Then Exit Do
        ElseIf (VerkPtr& <= 2&) Then
            Exit Do
        Else
            VerkPtr& = VerkPtr& - 1
        End If

        lPrivLauf& = lPrivLauf& + 1
    Loop

    If (pGebuehren = 0) Then
        Call iMsgBox("keine weiteren passenden Privat-Rezepte gespeichert !")
        VerkPtr& = OrgVerkPtr&
    End If

    lPrivEnde& = VerkPtr&
    lRezNr& = -1&
    PrivatRezept% = True
    GebFrei% = True

Else

    i% = 1
    Do
        If (i% > l%) Then Exit Do

        ch$ = Mid$(RezNr$, i%, 1)
        If (InStr("0123456789", ch$) > 0) Then
            i% = i% + 1
        Else
            RezNr$ = Left$(RezNr$, i% - 1) + Mid$(RezNr$, i% + 1)
            l% = Len(RezNr$)
        End If
    Loop
    If (l% > 13) Then RezNr$ = Left$(RezNr$, 13)

    lRezNr& = 1&
    If (l% < 13) Then
        If (l% <= 5) Then
            lRezNr& = Val(RezNr$)
        Else
            lRezNr& = 0
        End If
        If (lRezNr& <= 0) Or (lRezNr& > 32000) Then
            RezeptHolen% = False
            Call DefErrPop: Exit Function
        End If
        hStr$ = "9999998" + Format(lRezNr&, "00000")
        Call PruefZiffer(hStr$)
        RezNr$ = hStr$
    End If

'    lRezNr& = CLng(RezNr$)

    ParentForm.txtRezeptNr.text = RezNr$
    
    VerkPtr& = 0&
    
    FabsErrf% = RezTab.IndexSearch(0, RezNr$, FabsRecno&)
    If (FabsErrf% = 0) Then
        Call RezTab.GetRecord(FabsRecno&)
        VerkPtr& = RezTab.VerkPtr + 1
    End If

    If (RezeptHuellen%) Then
        SchonDa% = False
    ElseIf (RezSpeicherOK%) Then
        Set RezepteRec = RezSpeicherDB.OpenRecordset("Rezepte", dbOpenTable)
        With RezepteRec
            .index = "RezeptNr"
            .Seek ">=", RezNr$
            If (.NoMatch = False) Then
                If (RezepteRec!RezeptNr = RezNr$) Then
                    DruckDatum$ = RezepteRec!DruckDatum
                    SchonDa% = True
                End If
            End If
        End With
    End If
    
    If (SchonDa% = False) And (VerkPtr& <= 0) Then
        RezeptHolen% = False
        Call DefErrPop: Exit Function
    End If

End If


If (SchonDa%) Then
    Call HoleAusRezeptSpeicher
Else
    Call InitRezept(2)

    LetztArtikel$ = ""
    k% = -1
    
    Call vk.GetRecord(VerkPtr&)
    If (PrivatRezept%) Then pGebuehren = vk.gebuehren
    If (vk.PersCode > 0) Then PersonalNr% = vk.PersCode

    IvfRezept% = (vk.RezeptArt = 7)
    
    PrivatBesorger% = 0
    Do
        If (Mid$(vk.text, 15, 4) <> "PZN:") Then
            AktArtikel$ = vk.pzn + vk.text + vk.Preis
            If (Left$(vk.text, 5) = "u n t") Then
            ElseIf (vk.pzn = "9999999") And (PrivatRezept% = False) And (Left$(UCase(vk.text), 10) = "IK-NUMMER:") Then
                h$ = Trim(vk.KKNr)
                If (AutIdemSonderregelOk%) And (Val(h$) > 0) Then
                    SQLStr$ = "SELECT * FROM kKassen WHERE IK = " + Str$(Val(h$))
'                    Set KkassenRec = kKassenDB.OpenRecordset(SQLStr$)
'                    If (KkassenRec.RecordCount > 0) Then
                    FabsErrf = kKassenDB.OpenRecordset(KkassenRec, SQLStr)
                    If Not (KkassenRec.EOF) Then
                        AutIdemIk& = CheckNullLong(KkassenRec!IK)
                        AutIdemKbvNr& = CheckNullLong(KkassenRec!KbvNr)
                        ActAVWGik& = AutIdemIk&
                    End If
                End If
            ElseIf (PrivatBesorger% = 1) Then
                k% = k% + 1
                RezArtikel(k%).pzn = vk.pzn
                RezArtikel(k%).text = vk.text
                RezArtikel(k%).Preis = Val(vk.Preis)
                RezArtikel(k%).flag = 0
                If Chr(vk.AVP) = "A" Then RezArtikel(k%).appli = vk.AVP
                RezArtikel(k%).Zuz = 0
                RezArtikel(k%).Imp123 = 0
                RezArtikel(k%).Fam = 0
                RezArtikel(k%).AutIdem = 0
                RezArtikel(k%).N0 = 0
                RezArtikel(k%).MagIndex = 0
                RezArtikel(k%).Faktor = 1
                RezArtikel(k%).NichtInTaxe = 0
                RezArtikel(k%).ImpErspart = 0
                
                SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + vk.pzn
                'Set TaxeRec = TaxeDB.OpenRecordset(SQLStr$)
                On Error Resume Next
                TaxeRec.Close
                Err.Clear
                On Error GoTo DefErr
                TaxeRec.Open SQLStr, taxeAdoDB.ActiveConn
                If Not (TaxeRec.EOF) Then
                    RezArtikel(k%).pzn = Format(TaxeRec!pzn, String(7, "0"))
                    
                    hStr$ = Left$(TaxeRec!Name + Space$(29), 29) + Mid$(TaxeRec!menge, 3) + TaxeRec!einheit
                    RezArtikel(k%).text = hStr$
                    
                    RezArtikel(k%).Preis = TaxeRec!vk / 100#
                End If
                
                If (Left$(vk.text, 1) = "(") Then
                    RezArtikel(k%).Preis = PrivatBesorgerPreis#
                End If
            
                PrivatBesorger% = 2
            ElseIf (PrivatRezept%) And (Left$(vk.text, 6) = "ACONTO") Then
                PrivatBesorger% = 1
                PrivatBesorgerPreis# = Val(vk.Preis)
            ElseIf (Left$(vk.text, 5) = "Abhol") Then
                If ((PrivatRezept%) And ((vk.gebuehren <> pGebuehren) Or (Left$(vk.RezEan, 13) <> Space$(13)))) Then
                    k% = k% + 1
                Else
                    RezArtikel(k% + 1).zusatz(0) = vk.text
                End If
                PrivatBesorger% = 0
            ElseIf (PrivatRezept%) And (Left$(vk.text, 1) = "(") Then
                RezArtikel(k%).zusatz(1) = vk.text
                PrivatBesorger% = 0
            ElseIf (AktArtikel$ = LetztArtikel$) And (vk.pzn <> "9999999") And _
                ((PrivatRezept% = False) Or ((vk.gebuehren = pGebuehren) And (Left$(vk.RezEan, 13) = Space$(13)))) Then
                RezArtikel(k%).Faktor = RezArtikel(k%).Faktor + 1
                
                SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + RezArtikel(k%).pzn
                'Set TaxeRec = TaxeDB.OpenRecordset(SQLStr$)
                On Error Resume Next
                TaxeRec.Close
                Err.Clear
                On Error GoTo DefErr
                TaxeRec.Open SQLStr, taxeAdoDB.ActiveConn
                If (TaxeRec.EOF = False) Then
                    If (TaxeRec!HimiVerbrauch) Then
                        RezArtikel(k%).Zuz = RezArtikel(k%).Zuz + vk.gebsumme
                    End If
                End If
            ElseIf (Left$(vk.text, 1) = " ") And (Mid$(vk.text, 20, 1) = "x") Then
                iVerkMult% = Mid$(vk.text, 16, 3) - 1
                RezArtikel(k%).Faktor = RezArtikel(k%).Faktor + iVerkMult%
                If (PrivatBesorger% = 2) Then
                    RezArtikel(k%).Preis = RezArtikel(k%).Preis / RezArtikel(k%).Faktor
                Else
                    RezArtikel(k%).Preis = Val(Mid$(vk.text, 21, 9))
                End If
                PrivatBesorger% = 0
'                If (IvfRezept) Then
'                    RezArtikel(k%).Zuz = RezArtikel(k%).Zuz / RezArtikel(k%).Faktor
'                End If
            Else
                k% = k% + 1
                RezArtikel(k%).pzn = vk.pzn
                RezArtikel(k%).text = vk.text
                RezArtikel(k%).Preis = Val(vk.Preis)
                RezArtikel(k%).flag = 0
                If Chr(vk.AVP) = "A" Then RezArtikel(k%).appli = vk.AVP
                RezArtikel(k%).Zuz = 0
'                If (IvfRezept) Then
'                    RezArtikel(k%).Zuz = VK.gebsumme
'                End If
                RezArtikel(k%).Imp123 = 0
                RezArtikel(k%).Fam = 0
                RezArtikel(k%).AutIdem = 0
                RezArtikel(k%).N0 = 0
                RezArtikel(k%).MagIndex = 0
                RezArtikel(k%).Faktor = 1
                RezArtikel(k%).NichtInTaxe = 0
                RezArtikel(k%).ImpErspart = 0
                
                If (Left$(vk.text, 5) = "KONTR") Then
                    RezArtikel(k%).flag = RezArtikel(k%).flag Or (REZ_K_STOP Or REZ_BLINK)
                    Mid$(RezArtikel(k%).text, 21, 6) = "(PZN?)"
                    BeepFlag% = True
                End If
                If (Left$(vk.text, 12) = "PREISEINGABE") Then
                    Mid$(RezArtikel(k%).text, 21, 6) = "(PZN?)"
                End If
    
                If (RezArtikel(k%).pzn = "9999999") Then
                    If (RezArtikel(k%).flag And REZ_K_STOP) Then
                    ElseIf (RezArtikel(k%).Preis > 0) Then
                        RezArtikel(k%).appli = (vk.gebsumme = 0)
                    End If
                End If
                
                RezArtikel(k%).IstWg4 = 0
                RezArtikel(k%).PlusMehrKosten = 0
                RezArtikel(k%).ZuzahlungsErlass = 0
                RezArtikel(k%).Warenzeichen = 0
                RezArtikel(k%).MietDauer = 0
                
'                wg$ = VK.wg
'                If (Left$(wg$, 1) = "4") Then
'                    RezArtikel(k%).flag = RezArtikel(k%).flag Or (REZ_WG_4 Or REZ_BLINK)
'                    RezArtikel(k%).IstWg4 = 1
'                ElseIf (VmFlag%) Then
'                    FabsErrf% = VmPzn.IndexSearch(0, VK.pzn, FabsRecno&)
'                    If (FabsErrf% = 0) Then
'                        RezArtikel(k%).flag = RezArtikel(k%).flag Or (REZ_WG_4 Or REZ_BLINK)
'                    End If
'                End If
                
                SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + RezArtikel(k%).pzn
                'Set TaxeRec = TaxeDB.OpenRecordset(SQLStr$)
                On Error Resume Next
                TaxeRec.Close
                Err.Clear
                On Error GoTo DefErr
                TaxeRec.Open SQLStr, taxeAdoDB.ActiveConn
                If (TaxeRec.EOF = False) Then
                    RezArtikel(k%).IstWg4 = CheckZuzZeilenwert
                    If (TaxeRec!VerbandKz) Then
                        RezArtikel(k%).flag = RezArtikel(k%).flag Or (REZ_WG_4 Or REZ_BLINK)
                    ElseIf (VmFlag%) Then
                        If (AplusVOk) Then
                            SQLStr$ = "SELECT * FROM VdbArtikel WHERE PZN = " + vk.pzn
'                            Set VdbPznRec = AplusVDB.OpenRecordset(SQLStr$)
'                            FabsErrf = Abs(VdbPznRec.EOF)
                            FabsErrf = AplusVDB.OpenRecordset(VdbPznRec, SQLStr)
                            VdbPznRec.Close
'                        Else
'                            FabsErrf% = VmPzn.IndexSearch(0, vk.pzn, FabsRecno&)
                        End If
                        If (FabsErrf% = 0) Then
                            RezArtikel(k%).flag = RezArtikel(k%).flag Or (REZ_WG_4 Or REZ_BLINK)
                        End If
                    End If
                    
                    If (TaxeRec!HimiVerbrauch) Then
                        RezArtikel(k%).Zuz = vk.gebsumme
                    End If
                    
                    If (Asc(vk.ZuzaFlag) = 101) Then
                        RezArtikel(k%).PlusMehrKosten = 1
                    End If
                    If (Asc(vk.ZuzaFlag) = 102) Then
                        RezArtikel(k%).ZuzahlungsErlass = 1
                        RezArtikel(k%).Zuz = vk.gebsumme
                    End If
                End If
                
                PrivatBesorger% = 0
            End If
        End If
    
        LetztArtikel$ = AktArtikel$

        If (vk.EndeT And &H7) Or _
            ((PrivatRezept%) And ((vk.gebuehren <> pGebuehren) Or (Left$(vk.RezEan, 13) <> Space$(13)))) Then
            
            VerkDatum$ = sDate(vk.Datum)
            OrgVerkDatum$ = VerkDatum$

'          DruckDatum = 0L;

            VerkZeit% = vk.zeit
            
            ActVerordnung% = 1
            If (vk.RezeptArt = 1) Then ActVerordnung% = 2

            iUser% = vk.User
            VerkUser$ = Format(iUser%, " 0")
            VerkLiRe = vk.bs
    
            iKuNr = vk.knr
            If (iKuNr > 0) Then
                Call HoleKundenInfo(iKuNr)
            End If


            If (vk.EndeT And &H2) Then
                RezGebSumme# = 0#
                GebFrei% = True
            End If

            If (vk.KKTyp > 0) Then ActKasse% = vk.KKTyp
            
            ActKkasse$ = Trim(vk.KKNr)
            
            If (PrivatRezept) Then
                ActKasse% = 0
                ActKkasse$ = ""
            End If
            
'            If (AutIdemSonderregelOk%) And (Val(ActKkasse$) > 0) Then
'                SQLStr$ = "SELECT * FROM kKassen WHERE IK = " + Str$(Val(ActKkasse$))
'                Set KkassenRec = kKassenDB.OpenRecordset(SQLStr$)
'                If (KkassenRec.RecordCount > 0) Then
'                    AutIdemIk& = CheckNullLong(KkassenRec!IK)
'                    AutIdemKbvNr& = CheckNullLong(KkassenRec!KbvNr)
'                End If
'            End If


            If ((PrivatRezept% = False) Or ((vk.gebuehren = pGebuehren) And (Left$(vk.RezEan, 13) = Space$(13)))) Then
                k% = k% + 1
'                i++;
            Else
                lPrivEnde& = lPrivEnde& - 1
            End If
            
            Exit Do
        End If

        lPrivEnde& = lPrivEnde& + 1
        
        Call vk.GetRecord
    Loop
    
    AnzRezeptArtikel% = k%
    
    SonderBelegRezept% = False
    If (AnzRezeptArtikel% = 1) Then
        For i% = 1 To AnzSonderBelege%
            If (SonderBelege(i% - 1).pzn = RezArtikel(0).pzn) Then
                SonderBelegRezept% = i%
                PrivatRezept% = 0
                BtmRezept% = 0
                IvfRezept% = 0
                RezKundenInfo$(1) = ""
                RezKundenInfo$(2) = "SONDERBELEG"
                Exit For
            End If
        Next i%
    End If


    If Not (vk.EndeT And &H4) And Not (PrivatRezept%) And (SonderBelegRezept% = 0) Then
        For i% = 0 To UBound(BarVerkauf$)
            Call vk.GetRecord
    
            hhRez% = vk.RezNr
            If (hhRez% = 0) Then
                hStr$ = vk.text + " " + vk.Preis
                BarVerkauf$(i%) = hStr$
                
                With ParentForm
'                    .mnuBearbeitenInd(MENU_SF3).Enabled = True
'                    .cmdToolbar(MENU_SF3).Enabled = True
                End With
            End If
    
            If (hhRez% <> 0) Or (vk.EndeT And &H4) Then Exit For
        Next i%
    End If

    If (PrivatRezept%) Or (SonderBelegRezept%) Then
    Else
        AMverfügbarkeit = Left$(String(AnzRezeptArtikel%, "1") + "000", 3)
    End If

    For i% = 0 To (AnzRezeptArtikel% - 1)
        dZuz# = 0#
        dAvp# = 0
        
        If (PrivatRezept%) Or (SonderBelegRezept%) Then
            RezArtikel(i%).flag = RezArtikel(i%).flag And (REZ_BLINK Xor &HFF)
        Else
        
            If (RezArtikel(i%).pzn = "9999999") Then
                If (RezArtikel(i%).flag And REZ_K_STOP) Then
                    dZuz# = 5   'ZuzFeld!(3)
                    RezArtikel(i%).flag = RezArtikel(i%).flag Or REZ_BLINK
                ElseIf (RezArtikel(i%).Preis > 0) Then
                    If (RezArtikel(i%).appli = 0) Then
                        dZuz# = CalcZuzaNeu(RezArtikel(i%).Preis)
                    End If
                    
                    RezArtikel(i%).flag = RezArtikel(i%).flag Or REZ_BLINK
                    
                    ind2% = InStr(RezArtikel(i%).text, "(PZN?)")
                    If (ind2% <= 0) Then
                        ch$ = Trim(RezArtikel(i%).text) + "   " + "(PZN?)"
                        RezArtikel(i%).text = Left$(ch$ + Space$(36), 36)
                    End If
                End If
            Else
                SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + RezArtikel(i%).pzn
                'Set TaxeRec = TaxeDB.OpenRecordset(SQLStr$)
                On Error Resume Next
                TaxeRec.Close
                Err.Clear
                On Error GoTo DefErr
                TaxeRec.Open SQLStr, taxeAdoDB.ActiveConn
                If (TaxeRec.EOF) Then
                    RezArtikel(i%).NichtInTaxe = 1
                    If (CheckSonderPzn%(RezArtikel(i%).pzn) > 0) Then
'                        dZuz# = ZuzFeld!(3)
                        dZuz# = CalcZuzaNeu(RezArtikel(i%).Preis)
                    End If
                Else
'                    If (TaxeRec!Warengruppe = 3) Then
'                        RezArtikel(i%).IstWg4 = 1
'                    End If
                    
                    dAvp# = TaxeRec!vk / 100#
                    
                    dPreis# = RezArtikel(i%).Preis
                    If (RezArtikel(i%).PlusMehrKosten) Then
                        If (TaxeRec!FESTBETRAG > 0) And (TaxeRec!FESTBETRAG < TaxeRec!vk) Then
                            dPreis# = (TaxeRec!FESTBETRAG / 100#)
                        Else
                            RezArtikel(i%).PlusMehrKosten = 0
                        End If
                    End If
                    
                    If (RezArtikel(i%).ZuzahlungsErlass) Then
                        dZuzDos# = CalcTaxeZuzahlungDOS(dPreis#, RezArtikel(i%).Zuz + 200)
                    Else
                        dZuzDos# = CalcTaxeZuzahlungDOS(dPreis#)
                    End If
                    
                    If (TaxeRec!Negliste) Then dZuzDos# = 1#
                    If (RezArtikel(i%).appli) Then dZuzDos# = 0#
                    
                    If (RezArtikel(i%).PlusMehrKosten) Then
                        RezArtikel(i%).Preis = TaxeRec!vk / 100#
                    End If
                    
    '                If (Val(Left$(TaxeRec!Zuzahlung, 1)) = 9) Then dZuzDOS# = 0#
                    
                    If (dZuzDos# >= 1000#) Then
                        dZuz# = dZuzDos# - 1000#
                    ElseIf Not (GebFrei%) And (dZuzDos# > 0#) Then
                        RezArtikel(i%).Preis = 0#
                        If (dZuzDos# = 3#) Then
    '                        Call iMsgBox("Artikel unter Zuzahlung, Geldrückgabe ?")
                        Else
                            RezArtikel(i%).flag = RezArtikel(i%).flag And (REZ_BLINK Xor &HFF)
                        End If
                    End If
                    If (TaxeRec!HimiVerbrauch) Then
                        dZuz# = RezArtikel(i%).Zuz / RezArtikel(i%).Faktor
                    End If
                
                    If (RezArtikel(i%).ZuzahlungsErlass) Then
                        dZuz# = RezArtikel(i%).Zuz
                    End If
                
                    If (TaxeRec!BtmKz) Then
                        BtmRezept% = True
                    End If
                    
                    hStr$ = Format$(TaxeRec!Abgabebest, "0")
                    If (InStr("01245", hStr$) <> 0) Then
                        RezArtikel(i%).Fam = 1
                    End If
                
                    Call CheckAutIdem(i%)
                    SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + RezArtikel(i%).pzn
                    'Set TaxeRec = TaxeDB.OpenRecordset(SQLStr$)
                    On Error Resume Next
                    TaxeRec.Close
                    Err.Clear
                    On Error GoTo DefErr
                    TaxeRec.Open SQLStr, taxeAdoDB.ActiveConn
                    
                    
                    If (Left$(TaxeRec!Zuzahlung, 1) = "9") Then
                        RezArtikel(i%).IstWg4 = 0
                    End If
                    
                    If (ImpFlag% = 0) Then
                        RezArtikel(i%).Imp123 = 0
                    ElseIf (TaxeRec!OriginalPZN = TaxeRec!pzn) Then
                        RezArtikel(i%).Imp123 = 6
                        OrgAvp = TaxeRec!vk / 100#
                        SQLStr$ = "SELECT * FROM TAXE WHERE ORIGINALPZN = " + PznString(TaxeRec!OriginalPZN)
                        'Set TaxeRec = TaxeDB.OpenRecordset(SQLStr$)
                        On Error Resume Next
                        TaxeRec.Close
                        Err.Clear
                        On Error GoTo DefErr
                        TaxeRec.Open SQLStr, taxeAdoDB.ActiveConn
                        TaxeRec.MoveFirst
                        Do While Not TaxeRec.EOF
                          If (TaxeRec!OriginalPZN <> TaxeRec!pzn) Then
                            aktavp = TaxeRec!vk / 100#
                            If (aktavp <= OrgAvp - 15) Or (aktavp <= OrgAvp * 0.85) Then
                              RezArtikel(i%).Imp123 = 5
                              Exit Do
                            End If
                          End If
                          TaxeRec.MoveNext
                        Loop
                        SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + RezArtikel(i%).pzn
                        'Set TaxeRec = TaxeDB.OpenRecordset(SQLStr$)
                        On Error Resume Next
                        TaxeRec.Close
                        Err.Clear
                        On Error GoTo DefErr
                        TaxeRec.Open SQLStr, taxeAdoDB.ActiveConn
                    
                    ElseIf (TaxeRec!OriginalPZN > 0) Then
                        aktavp = TaxeRec!vk / 100
                        SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + PznString(TaxeRec!OriginalPZN)
                        'Set TaxeRec = TaxeDB.OpenRecordset(SQLStr$)
                        On Error Resume Next
                        TaxeRec.Close
                        Err.Clear
                        On Error GoTo DefErr
                        TaxeRec.Open SQLStr, taxeAdoDB.ActiveConn
                        If Not (TaxeRec.EOF) Then
                          OrgAvp = TaxeRec!vk / 100
                        End If
                        If (aktavp <= OrgAvp - 15) Or (aktavp <= OrgAvp * 0.85) Then
                          RezArtikel(i%).Imp123 = 2
                          RezArtikel(i%).ImpErspart = (OrgAvp - aktavp)
                        Else
                          RezArtikel(i%).Imp123 = 1
                        End If
                        SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + RezArtikel(i%).pzn
                        'Set TaxeRec = TaxeDB.OpenRecordset(SQLStr$)
                        On Error Resume Next
                        TaxeRec.Close
                        Err.Clear
                        On Error GoTo DefErr
                        TaxeRec.Open SQLStr, taxeAdoDB.ActiveConn
                    
                    End If
                End If
            End If
        End If
        
        If (IvfRezept) Then
            RezArtikel(i%).Zuz = 0
            RezArtikel(i%).Preis = dAvp / 2
        Else
            RezArtikel(i%).Zuz = dZuz#
        End If
    Next i%
    
    If (IvfRezept) Then
        If (RezArtikel(0).pzn <> "9999643") Then
            For i% = UBound(RezArtikel) - 1 To 0 Step -1
                RezArtikel(i% + 1) = RezArtikel(i%)
            Next i%
            Call InitRezeptArtikel(0)
'            Call InitKontrollStop(row%)

            RezArtikel(0).pzn = "9999643"
            RezArtikel(0).text = "IVF-Rezept"
            RezArtikel(0).Faktor = 1
            RezArtikel(0).Preis = 0
            RezArtikel(0).Zuz = 0
                
            AnzRezeptArtikel% = AnzRezeptArtikel% + 1
        End If
    End If
    
    If (BtmRezept%) Then
        For i% = 0 To UBound(RezAddGesamtBrutto$)
            RezAddGesamtBrutto$(i%) = ""
        Next i%
        Call PruefeBtmKosten
        sBrutto$ = RezAddGesamtBrutto$(0)
    End If
    
    Call PruefeRezkontrDat
    
    If (ActKasse% > 0) Then
        For i% = 0 To (AnzRezeptArtikel% - 1)
            If (RezArtikel(i%).flag And REZ_WG_4) Then
                RezeptHolenZuz# = RezArtikel(i%).Zuz
                AktTabIndex% = 3 + i% * 4
                Call Verbandmittel
            End If
        Next i%
    End If
    RezeptHolenZuz# = 0
    
End If

Call CheckVerordnungMDB

'ActVerordnung% = 1

RezeptHolen% = True

Call DefErrPop
End Function

Function RezeptHolenDB%(Optional PrivatModus% = 0)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("RezeptHolenDB%")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
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
Dim i%, j%, k%, l%, ind%, PrivGef%, iUser%, iVerkMult%, hhRez%, IstDrüber%, ind2%, SuchePrivatRezepte%, IstWVORabattArtikel%, bCheckRezeptEnde%
Dim pTag%, pMonat%, pJahr%, PrivDatum%, iVerkDatum%, RezNrOk%, eRezeptSollRezepturen%
Dim PrivUser As Byte, PrivLiRe As Byte, AktPrivRezNr As Byte, eRezeptSollOpStatus As Byte
Dim lAnzVkSaetze&, lIkNr&, iKuNr&, lCheckVerkaufId&
Dim dZuz#, dZuzDos#, dPreis#, dAvp#, dLastZuzahlung#
Dim ch$, RezEan$, hStr$, LetztArtikel$, AktArtikel$, wg$, SQLStr$, Unique$, h$, h2$, h3$, sPrivDatum$, sPzn$, sText$, OrgRezNr$, AMverfügbarkeit2$, RezeptImage$
Dim eRezeptSollTaskId$
Dim aktavp As Double, OrgAvp As Double
Dim iiDat As Date
Dim VerkaufDat As Date
Dim AvpRec As Recordset
Dim AVP#, FESTBETR#, VglWert#
Dim PreisDiff_Log%
Static LastID&
Static LastDatum As Date
Static LastLaufNr&

Static fOrgRezNr!
'Static lPrivEnde&

PreisDiffStr = ""
RezeptImage = ""

If (PrivatModus <> 0) Then
'    MsgBox ("etaskid ....:" + eRezeptTaskId)
    If (eRezeptTaskId <> "") Then
        RezNr = "E" + eRezeptTaskId
        If (eRezeptListe <> "") Then
'            MsgBox (eRezeptListe + vbCrLf + vbCrLf + RezNr)
            h3 = "," + RezNr + ","
            ind = InStr(eRezeptListe, h3)
            If (ind > 0) Then
                If (PrivatModus = -1) Then
                    h2 = Left(eRezeptListe, ind - 1)
                    ind = InStrRev(h2, ",")
                    If (ind > 0) Then
                        RezNr = Mid(h2, ind + 1)
                        ParentForm.txtRezeptNr.text = RezNr
                    End If
                Else
                    h2 = Mid(eRezeptListe, ind + Len(h3))
                    ind = InStr(h2, ",")
                    If (ind > 0) Then
                        RezNr = Left(h2, ind - 1)
                        ParentForm.txtRezeptNr.text = RezNr
                    End If
                End If
                PrivatModus = 0
            End If
        End If
    End If
End If

h$ = RezNr$

'MsgBox ("PrivatModus: " + CStr(PrivatModus))

Call WinRezDebug("RezNr: " + h)

Call InitRezept(1)

RezNr$ = Trim(h$)
l% = Len(RezNr)


MarsRezeptAnmerkungen = ""
If (para.MARS) Then
    Dim TestlRezNr&, TestRezNr$, TestRezNr2$
    
    TestRezNr = RezNr
    i% = 1
    Do
        If (i% > l%) Then Exit Do

        ch$ = Mid$(TestRezNr$, i%, 1)
        If (InStr("0123456789", ch$) > 0) Then
            i% = i% + 1
        Else
            TestRezNr$ = Left$(TestRezNr$, i% - 1) + Mid$(TestRezNr$, i% + 1)
            l% = Len(TestRezNr$)
        End If
    Loop
    If (l% > 13) Then TestRezNr$ = Left$(TestRezNr$, 13)

    TestlRezNr& = 1&
    If (l% < 13) Then
        If (l% <= 5) Then
            TestlRezNr& = Val(TestRezNr$)
        Else
            TestlRezNr& = 0
        End If
'        If ((lRezNr& <= 0) Or (lRezNr& > 32000)) Then
        If (TestlRezNr& <= 0) Then
            RezeptHolenDB% = False
            Call DefErrPop: Exit Function
        End If
        hStr$ = "9999998" + Format(TestlRezNr&, "00000")
        Call PruefZiffer(hStr$)
        TestRezNr$ = hStr$
    End If
    
    TestRezNr2 = ""
    If (Left(TestRezNr, 3) = "989") Then
        hStr$ = "9999998" + Mid(RezNr, 8, 5)
        Call PruefZiffer(hStr$)
        TestRezNr2$ = hStr$
    End If
    

    
    
    SQLStr = "SELECT * FROM Anmerkungen WHERE RezeptNr='" + TestRezNr + "'"
    Set MarsAnmerkungenRec = MarsRezSpeicherDB.OpenRecordset(SQLStr$)
    If (MarsAnmerkungenRec.EOF) And (TestRezNr2 <> "") Then
        SQLStr = "SELECT * FROM Anmerkungen WHERE RezeptNr='" + TestRezNr2 + "'"
        Set MarsAnmerkungenRec = MarsRezSpeicherDB.OpenRecordset(SQLStr$)
    End If
    If (MarsAnmerkungenRec.RecordCount > 0) Then
        MarsRezeptAnmerkungen = Trim(CheckNullStr(MarsAnmerkungenRec!Anmerkungen))
    End If

    SQLStr = "SELECT * FROM Rezepte WHERE RezeptNr='" + TestRezNr + "'"
    Set RezepteRec = RezSpeicherDB.OpenRecordset(SQLStr$)
    If (RezepteRec.EOF) And (TestRezNr2 <> "") Then
        SQLStr = "SELECT * FROM Rezepte WHERE RezeptNr='" + TestRezNr2 + "'"
        Set RezepteRec = RezSpeicherDB.OpenRecordset(SQLStr$)
    End If
    If (RezepteRec.RecordCount > 0) Then
        h = CheckNullStr(RezepteRec!DruckDatum)
        Call MessageBox("Achtung:" + vbCrLf + vbCrLf + "Rezept wurde BEREITS BEDRUCKT (" + Mid(h, 5, 2) + "." + Mid(h, 3, 2) + ".20" + Left(h, 2) + ") !", vbInformation)
    End If

    SQLStr = "SELECT * FROM Rezepte WHERE RezeptNr='" + TestRezNr + "'"
    Set MarsRezepteRec = MarsRezSpeicherDB.OpenRecordset(SQLStr$)
    If (MarsRezepteRec.EOF) And (TestRezNr2 <> "") Then
        SQLStr = "SELECT * FROM Rezepte WHERE RezeptNr='" + TestRezNr2 + "'"
        Set MarsRezepteRec = MarsRezSpeicherDB.OpenRecordset(SQLStr$)
    End If
    If (MarsRezepteRec.RecordCount > 0) Then
        If (MarsModus = MARS_REZEPT_KONTROLLE) Then
            Call MessageBox("Achtung:" + vbCrLf + vbCrLf + "Rezept wurde BEREITS FREIGEGEBEN!", vbInformation)
        End If
    Else
        If (MarsModus = MARS_REZEPT_DRUCK) Then
            Call MessageBox("Achtung:" + vbCrLf + vbCrLf + "Rezept wurde NOCH NICHT FREIGEGEBEN!", vbCritical)
            RezeptHolenDB% = False
            Call DefErrPop: Exit Function
        End If
    End If
End If



If (UCase$(Left$(RezNr, 1)) = "X") Then
    If (AvpTeilnahme%) Then
        RezNr$ = Mid$(RezNr$, 2, 11)
'        h$ = Mid$(RezNr$, 7, 4)
'        h$ = Mid$(h$, 3, 2) + Left$(h$, 2)
        SQLStr$ = "SELECT * FROM Rezepte WHERE AvpRezeptNr='" + RezNr$ + "'"
        Set AvpRec = RezSpeicherDB.OpenRecordset(SQLStr$)
        If (AvpRec.RecordCount > 0) Then
            RezNr$ = Trim(AvpRec!RezeptNr)
            l% = Len(RezNr)
        End If
    End If
End If

OrgRezNr = RezNr
RezNrOk = 0

'MsgBox (RezNr + "  " + OrgRezNr)
If ((UCase$(Left$(RezNr, 1)) = "P") Or (UCase$(Left$(RezNr, 1)) = "E")) Then
    PrivUser = 0
    PrivLiRe = 0
    PrivDatum% = 0
    sPrivDatum$ = ""
    eRezeptSollOpStatus = 0
    eRezeptSollTaskId = ""
    eRezeptSollRezepturen = 0

    RezNr = Mid$(RezNr, 2)
    l% = Len(RezNr)
    SuchePrivatRezepte = (l = 0)

    If (UCase$(Left$(OrgRezNr, 1)) = "E") Then
'        MsgBox (Str(l) + " <" + RezNr + ">")
        If (l > 0) Then
            If (l > 12) Then
                eRezeptSollTaskId = RezNr
                RezNr = ""
                l = 0
                SuchePrivatRezepte = (l = 0)
            ElseIf (InStr("12345", Left(RezNr, 1)) > 0) Then
                eRezeptSollOpStatus = Val(Left(RezNr, 1)) + 1
'                MsgBox (CStr(eRezeptSollOpStatus))
                RezNr = Mid$(RezNr, 2)
                l% = Len(RezNr)
                SuchePrivatRezepte = (l = 0)
'                MsgBox ("suche...: " + CStr(SuchePrivatRezepte))
            End If
            
            If (UCase(Left(RezNr, 1)) = "R") Then
                eRezeptSollRezepturen = True
                RezNr = Mid$(RezNr, 2)
                l% = Len(RezNr)
                SuchePrivatRezepte = (l = 0)
            End If
        End If
    End If
    
    If (InStr(UCase(RezNr), "C") > 0) Then
        SuchePrivatRezepte = True
    End If
    If (InStr(UCase(RezNr), "D") > 0) Then
        SuchePrivatRezepte = True
    End If
    If (SuchePrivatRezepte) And (l > 0) Then
        h = UCase(RezNr)
        Do
            ch = UCase(Left(h, 1))
            If (ch = "D") Then
                h = Mid(h, 2) + "C"
                ind = InStr(h, "C")
                If (ind > 0) Then
                    h2 = Left(h, ind - 1)
                    h = Mid(h, ind)
                    h = Left(h, Len(h) - 1)

                    If (Len(h2) = 4) Then
                        sPrivDatum$ = h2
                        For i% = 1 To 4
                            If (InStr("0123456789", Mid$(h2, i%, 1)) = 0) Then
                                sPrivDatum$ = ""
                                Exit For
                            End If
                        Next i%
                        If (sPrivDatum$ <> "") Then
                            pTag% = Val(Left$(sPrivDatum$, 2))
                            If (pTag% = 0) Or (pTag% > 31) Then
                                sPrivDatum$ = ""
                            End If
                        End If
                        If (sPrivDatum$ <> "") Then
                            pMonat% = Val(Mid$(sPrivDatum$, 3, 2))
                            If (pMonat% = 0) Or (pMonat% > 12) Then
                                sPrivDatum$ = ""
                            End If
                        End If
                        If (sPrivDatum$ <> "") Then
                            pJahr% = Year(Now) - 2000
                            iiDat = Format(pTag, "00") + "." + Format(pMonat%, "00") + "." + Format(Now, "yyyy")
                            If (DateDiff("d", iiDat, Now) < 0) Then
                                pJahr% = pJahr% - 1
                            End If
                            h2$ = Format(pTag, "00") + Format(pMonat%, "00") + Format(pJahr%, "00")
                            PrivDatum% = iDate(h2$)
                        End If
                    End If
                End If
            ElseIf (ch = "C") Then
                h = Mid(h, 2) + "D"
                ind = InStr(h, "D")
                If (ind > 0) Then
                    h2 = Left(h, ind - 1)
                    h = Mid(h, ind)
                    h = Left(h, Len(h) - 1)

                    If (Len(h2) >= 1) Then
                        ch$ = Left(h2, 1)
                        If (InStr("0123456789", ch$) > 0) Then
                            PrivUser = Val(ch$)
                            h2 = Mid(h2, 2)
                            If (h2 <> "") Then
                                ind% = InStr("LR", h2)
                                If (ind% > 0) Then PrivLiRe = ind%
                            End If
                        End If
                    End If
                End If
            Else
                Exit Do
            End If
        Loop
    End If
End If

'If (UCase$(Left$(OrgRezNr, 1)) = "P") And (l = 0) Then
If ((UCase$(Left$(OrgRezNr, 1)) = "P") Or (UCase$(Left$(OrgRezNr, 1)) = "E")) And (SuchePrivatRezepte) Then
Else
    RezNrOk = True

    Call WinRezDebug("RezNr2: " + RezNr)

    i% = 1
    Do
        If (i% > l%) Then Exit Do

        ch$ = Mid$(RezNr$, i%, 1)
        If (InStr("0123456789", ch$) > 0) Then
            i% = i% + 1
        Else
            RezNr$ = Left$(RezNr$, i% - 1) + Mid$(RezNr$, i% + 1)
            l% = Len(RezNr$)
        End If
    Loop
    If (l% > 13) Then RezNr$ = Left$(RezNr$, 13)

    Call WinRezDebug("RezNr3: " + RezNr)

    lRezNr& = 1&
    If (l% < 13) Then
        If (l% <= 5) Then
            lRezNr& = Val(RezNr$)
        Else
            lRezNr& = 0
        End If
'        If ((lRezNr& <= 0) Or (lRezNr& > 32000)) Then
        If (lRezNr& <= 0) Then
            If ((UCase$(Left$(OrgRezNr, 1)) = "P") Or (UCase$(Left$(OrgRezNr, 1)) = "E")) Then
                RezNrOk = 0
            Else
                RezeptHolenDB% = False
                Call DefErrPop: Exit Function
            End If
        End If
        hStr$ = "9999998" + Format(lRezNr&, "00000")
        Call PruefZiffer(hStr$)
        RezNr$ = hStr$
    End If

    Call WinRezDebug("RezNr4: " + RezNr)
    Call WinRezDebug("OrgRezNr4: " + OrgRezNr)


    bmVerkPtr = Null
'    MsgBox ("null")

'    Set VerkaufRec = VerkaufDB.OpenRecordset("Verkauf", dbOpenTable)

    If ((UCase$(Left$(OrgRezNr, 1)) = "P") Or (UCase$(Left$(OrgRezNr, 1)) = "E")) Then
'        VerkaufRec.index = "PrivRezNr"
'        VerkaufRec.Seek ">=", Mid$(OrgRezNr$, 2), Now
'        If (VerkaufRec.NoMatch = False) Then
'            If (VerkaufRec!PrivRezLaufNr = Val(Mid$(OrgRezNr$, 2))) Then
'                bmVerkPtr = VerkaufRec.Bookmark
'                PrivatRezept% = True
'                GebFrei% = True
'            End If
'        End If
        SQLStr = "SELECT * FROM RezeptNummern WHERE RezeptNr='" + CStr(Val(Mid$(OrgRezNr$, 2))) + "'"
'        SQLStr = SQLStr + " ORDER BY Datum DESC,LaufNr DESC,ZeilenNr"
        Call WinRezDebug("SQL1: " + SQLStr)
        FabsErrf = VerkaufAdoDB.OpenRecordset(VerkaufAdoRec, SQLStr, 0)
        If Not (VerkaufAdoRec.EOF) Then
            SQLStr = "SELECT * FROM Verkauf WHERE PrivRezLaufNr=" + CStr(Val(Mid$(OrgRezNr$, 2)))
            SQLStr = SQLStr + " ORDER BY Datum DESC,LaufNr DESC,ZeilenNr"
            Call WinRezDebug("SQL2: " + SQLStr)
            FabsErrf = VerkaufAdoDB.OpenRecordset(VerkaufAdoRec, SQLStr, 0)
            If Not (VerkaufAdoRec.EOF) Then
                bmVerkPtr = VerkaufAdoRec.Bookmark
                Call WinRezDebug("bmVerkPtr: " + CStr(bmVerkPtr))
                PrivatRezept% = True
                GebFrei% = True
            End If
        End If
    Else
        ParentForm.txtRezeptNr.text = RezNr$

'        VerkaufRec.index = "RezeptNr"
'        VerkaufRec.Seek ">=", RezNr$, Now
'        If (VerkaufRec.NoMatch = False) Then
'            If (VerkaufRec!RezeptNr = RezNr$) Then
'                bmVerkPtr = VerkaufRec.Bookmark
'            End If
'        End If
        
'        SQLStr = "SELECT * FROM RezeptNummern WHERE RezeptNr='" + RezNr + "'"
''        SQLStr = SQLStr + " ORDER BY Datum DESC,LaufNr DESC,ZeilenNr"
'        FabsErrf = VerkaufAdoDB.OpenRecordset(VerkaufAdoRec, SQLStr, 0)
'        If Not (VerkaufAdoRec.EOF) Then
'            SQLStr = "SELECT * FROM Verkauf WHERE RezeptNr='" + RezNr + "'"
'            SQLStr = SQLStr + " ORDER BY Datum DESC,LaufNr DESC,ZeilenNr"
'            FabsErrf = VerkaufAdoDB.OpenRecordset(VerkaufAdoRec, SQLStr, 0)
'            If Not (VerkaufAdoRec.EOF) Then
'                bmVerkPtr = VerkaufAdoRec.Bookmark
'            End If
'        End If
    
        SQLStr = "SELECT * FROM Verkauf WHERE RezeptNr='" + RezNr + "'"
        SQLStr = SQLStr + " ORDER BY Datum DESC,LaufNr DESC,ZeilenNr"
        Call WinRezDebug("SQL3: " + SQLStr)
        FabsErrf = VerkaufAdoDB.OpenRecordset(VerkaufAdoRec, SQLStr, 0)
        If (VerkaufAdoRec.EOF) Then
'            SQLStr = "SELECT * FROM Verkauf WHERE PrivRezLaufNr=" + CStr(lRezNr)
            SQLStr = "SELECT * FROM Verkauf WHERE PrivRezLaufNr=" + Mid(RezNr, 8, 5)
            SQLStr = SQLStr + " ORDER BY Datum DESC,LaufNr DESC,ZeilenNr"
            Call WinRezDebug("SQL4: " + SQLStr)
            FabsErrf = VerkaufAdoDB.OpenRecordset(VerkaufAdoRec, SQLStr, 0)
            If Not (VerkaufAdoRec.EOF) Then
                bmVerkPtr = VerkaufAdoRec.Bookmark
                Call WinRezDebug("bmVerkPtr2: " + CStr(bmVerkPtr))
                PrivatRezept% = True
                GebFrei% = True
            End If
        End If
        If Not (VerkaufAdoRec.EOF) Then
            If (CheckNullInt(VerkaufAdoRec!StornoKunde) = 0) Then
                bmVerkPtr = VerkaufAdoRec.Bookmark
                Call WinRezDebug("bmVerkPtr3: " + CStr(bmVerkPtr))
            End If
        End If
    
    End If

    If (RezeptHuellen%) Then
        SchonDa% = False
    ElseIf (RezSpeicherOK%) Then
        Set RezepteRec = RezSpeicherDB.OpenRecordset("Rezepte", dbOpenTable)
        With RezepteRec
            .index = "RezeptNr"
            Call WinRezDebug("RezSpeicher-Seek: " + RezNr)
            .Seek ">=", RezNr$
            If (.NoMatch = False) Then
                If (RezepteRec!RezeptNr = RezNr$) Then
                    DruckDatum$ = RezepteRec!DruckDatum
                    SchonDa% = True
                    Call WinRezDebug("DruckDatum: " + DruckDatum)
                End If
            End If
        End With
    End If

    If (SchonDa% = False) And (IsNull(bmVerkPtr)) Then
'        If (UCase$(Left$(OrgRezNr, 1)) = "P") Then
'            RezNrOk = 0
'        Else
            RezeptHolenDB% = False
            Call DefErrPop: Exit Function
'        End If
    End If
End If

Call WinRezDebug("SchonDa: " + CStr(SchonDa))

bCheckRezeptEnde = 0
lCheckVerkaufId = 0
If ((UCase$(Left$(OrgRezNr, 1)) = "P") Or (UCase$(Left$(OrgRezNr, 1)) = "E")) And (RezNrOk = 0) Then
    RezNr = OrgRezNr
    RezNr$ = UCase$(RezNr$)
    
    bCheckRezeptEnde = True

    bmOrgVerkPtr = bmVerkPtr

    AktPrivRezNr = 0

    If (PrivatModus <> 2) Then
'        Set VerkaufRec = VerkaufDB.OpenRecordset("Verkauf", dbOpenTable)
'        VerkaufRec.index = "Unique"
    End If

    Select Case PrivatModus%
        Case 0
            LastID = 0
            LastDatum = "01.01.2000"
            LastLaufNr = 1
            SQLStr = "SELECT t.* FROM ("
'            SQLStr = SQLStr + "SELECT TOP 100 Verkauf.*, TI_eRezepte.Bundle, TI_eRezepte.OpStatus FROM Verkauf"
            If (UCase$(Left$(OrgRezNr, 1)) = "P") Then
                SQLStr = SQLStr + "SELECT TOP 100 Verkauf.* FROM Verkauf"
                SQLStr = SQLStr + " WHERE PrivRezLaufNr>0"
            Else
                SQLStr = SQLStr + "SELECT TOP 100 Verkauf.*, TI_eRezepte.Bundle, TI_eRezepte.OpStatus, TI_eRezepte.VerordnungsTyp, ZA_11_Schluessel, Botendienst, Beschaffungskosten FROM Verkauf"
'                SQLStr = SQLStr + " WHERE NOT(eRezTaskID IS NULL)"
                SQLStr = SQLStr + " LEFT JOIN TI_eRezepte ON Verkauf.eRezTaskID=TI_eRezepte.TaskId"
                SQLStr = SQLStr + " WHERE LEN(eRezTaskID)>0"
                SQLStr = SQLStr + " AND LEN(RezeptNr)>0"
                If (eRezeptSollOpStatus > 0) Then
                    SQLStr = SQLStr + " AND (OpStatus=" + CStr(eRezeptSollOpStatus) + ")"
                    If (eRezeptSollRezepturen) Then
                        SQLStr = SQLStr + " AND (VerordnungsTyp=3)"
                    End If
                ElseIf (eRezeptSollTaskId <> "") Then
                    SQLStr = SQLStr + " AND (eRezTaskID='" + eRezeptSollTaskId + "')"
                End If
            End If
            If (PrivDatum% > 0) Then
                h = sDate(PrivDatum)
                SQLStr = SQLStr + " AND (Datum<='" + Left(h, 2) + "." + Mid(h, 3, 2) + ".20" + Mid(h, 5, 2) + " 23:59" + "')"
            End If
            If (PrivUser) Then
                SQLStr = SQLStr + " AND (Computer=" + CStr(PrivUser) + ")"
                If (PrivLiRe) Then
                    SQLStr = SQLStr + " AND (Seite=" + CStr(PrivLiRe - 1) + ")"
                End If
            End If
            SQLStr = SQLStr + " ORDER BY Datum DESC,LaufNr DESC,ZeilenNr Desc"
            SQLStr = SQLStr + ") T"
            SQLStr = SQLStr + " ORDER BY Datum,LaufNr,ZeilenNr"
            Clipboard.SetText SQLStr, vbCFText
'            MsgBox (SQLStr)
            FabsErrf = VerkaufAdoDB.OpenRecordset(VerkaufAdoRec, SQLStr, 0)
            If (FabsErrf <> 0) Then
                If (UCase$(Left$(OrgRezNr, 1)) = "P") Then
                    Call iMsgBox("keine passenden Privat-Rezepte gespeichert !")
                Else
                    Call iMsgBox("keine passenden eRezepte gespeichert !")
                End If
                RezeptHolenDB% = False
                Call DefErrPop: Exit Function
            End If
            VerkaufAdoRec.MoveLast
            
'            VerkaufAdoRec.MoveFirst
        Case -1
'            MsgBox ("-1")
            VerkaufAdoRec.Bookmark = bmVerkPtr
'            MsgBox (".bookmark")
            VerkaufAdoRec.MovePrevious
        Case 1
'            MsgBox ("1")
            VerkaufAdoRec.Bookmark = bmPrivEnde
'            MsgBox (".bookmark")
            VerkaufAdoRec.MoveNext
        Case 2
            'VerkPtr& wurde von aussen gesetzt
    End Select
'

    Do
'        MsgBox ("do")
        If (VerkaufAdoRec.EOF) Or (VerkaufAdoRec.BOF) Then
            Exit Do
        End If

        bmVerkPtr = VerkaufAdoRec.Bookmark
'        MsgBox ("gesetzt")

        RezEan$ = CheckNullStr(VerkaufAdoRec!RezeptNr)
'        MsgBox ("rezean: " + RezEan)

        If (AktPrivRezNr > 0) Then
            If ((VerkaufAdoRec!KundenEnde) Or (VerkaufAdoRec!RezeptEnde)) And (PrivatModus% <> 1) Then
                VerkaufAdoRec.MoveNext
                bmVerkPtr = VerkaufAdoRec.Bookmark
'                MsgBox ("gesetzt2")
                Exit Do
            End If

'            If (VerkaufAdoRec!PrivRezNr <> AktPrivRezNr) Or (Val(RezEan) > 0) Then
'                VerkaufAdoRec.MoveNext
'                bmVerkPtr = VerkaufAdoRec.Bookmark
''                MsgBox ("gesetzt3")
'                Exit Do
'            End If
        End If

        If (UCase$(Left$(OrgRezNr, 1)) = "P") Then
            If (VerkaufAdoRec!PrivRezNr <> AktPrivRezNr) Or (Val(RezEan) > 0) Then
                VerkaufAdoRec.MoveNext
                bmVerkPtr = VerkaufAdoRec.Bookmark
'                MsgBox ("gesetzt3")
                Exit Do
            End If
            
            If ((VerkaufAdoRec!RezeptArt = REZEPTART_PRIVAT) Or VerkaufAdoRec!RezeptArt = REZEPTART_GRÜN) And (AktPrivRezNr = 0) And (Val(RezEan) = 0) Then
                PrivGef% = True
                If (VerkaufAdoRec!PrivRezNr = 0) Then
                    PrivGef% = False
                End If
                If (PrivDatum% > 0) Then
                    h$ = Format(VerkaufAdoRec!Datum, "DDMMYY")
                    iVerkDatum% = iDate(h$)
                    If (iVerkDatum > PrivDatum%) Then
                        PrivGef% = False
                    End If
                ElseIf (PrivUser) Then
                    iUser% = VerkaufAdoRec!Computer
                    If (PrivUser <> iUser%) Then
                        PrivGef% = False
                    Else
                        If (PrivLiRe) Then
                            VerkLiRe = VerkaufAdoRec!Seite
                            If (VerkLiRe <> (PrivLiRe - 1)) Then PrivGef% = False
                        End If
                    End If
                End If
    
                If (PrivGef%) Then
                    AktPrivRezNr = VerkaufAdoRec!PrivRezNr
                    If (PrivatModus% = 1) Then Exit Do
                End If
            End If
        Else
            If (AktPrivRezNr = 0) Then
                PrivGef% = True
'                If (VerkaufAdoRec!PrivRezNr = 0) Then
'                    PrivGef% = False
'                End If
                If (PrivDatum% > 0) Then
                    h$ = Format(VerkaufAdoRec!Datum, "DDMMYY")
                    iVerkDatum% = iDate(h$)
                    If (iVerkDatum > PrivDatum%) Then
                        PrivGef% = False
                    End If
                ElseIf (PrivUser) Then
                    iUser% = VerkaufAdoRec!Computer
                    If (PrivUser <> iUser%) Then
                        PrivGef% = False
                    Else
                        If (PrivLiRe) Then
                            VerkLiRe = VerkaufAdoRec!Seite
                            If (VerkLiRe <> (PrivLiRe - 1)) Then PrivGef% = False
                        End If
                    End If
                End If
    
                If (PrivGef%) Then
                    AktPrivRezNr = 1 ' VerkaufAdoRec!PrivRezNr
                    If (PrivatModus% = 1) Then Exit Do
                End If
            End If
        End If

        LastID = VerkaufAdoRec!Id
        LastDatum = VerkaufAdoRec!Datum
        LastLaufNr = VerkaufAdoRec!LaufNr
        If (PrivatModus% = 1) Then
            VerkaufAdoRec.MoveNext
        Else
            VerkaufAdoRec.MovePrevious
        End If
    Loop

    If (AktPrivRezNr = 0) Then
        If (UCase$(Left$(OrgRezNr, 1)) = "P") Then
            Call iMsgBox("keine weiteren passenden Privat-Rezepte gespeichert !")
        Else
            Call iMsgBox("keine weiteren passenden eRezepte gespeichert !")
        End If
        bmVerkPtr = bmOrgVerkPtr
'        MsgBox ("gesetzt4")
        If (IsNull(bmVerkPtr)) Then
'            MsgBox ("false")
            RezeptHolenDB% = False
            Call DefErrPop: Exit Function
        End If
    End If

    bmPrivEnde = bmVerkPtr
    If (UCase$(Left$(OrgRezNr, 1)) = "P") Then
        lRezNr& = -1&
        PrivatRezept% = True
        GebFrei% = True
'    Else
'        eRezeptTaskId = VerkaufAdoRec!eRezTaskID
'        eRezeptBundleKBV = VerkaufAdoRec!Bundle
'        eRezeptOpStatus = VerkaufAdoRec!OpStatus
'        eRezeptPzn = VerkaufAdoRec!pzn
    End If
End If


If (SchonDa%) Then
    Call HoleAusRezeptSpeicher
ElseIf (MarsModus = MARS_REZEPT_DRUCK) Then
    Call HoleAusRezeptSpeicher((RezNr))
Else
    Call InitRezept(2)
    AMverfügbarkeit2 = String(10, "1")

    LetztArtikel$ = ""
    k% = -1

    VerkaufAdoRec.Bookmark = bmVerkPtr
    If (PrivatRezept%) Then
        AktPrivRezNr = VerkaufAdoRec!PrivRezNr
    End If
    If (VerkaufAdoRec!Bediener > 0) Then
        PersonalNr% = VerkaufAdoRec!Bediener
        Call WinRezDebug("PersonalNr: " + CStr(PersonalNr))
    End If
    If (UCase$(Left$(OrgRezNr, 1)) = "E") Then
        eRezeptTaskId = VerkaufAdoRec!eRezTaskID
        eRezeptBundleKBV = VerkaufAdoRec!Bundle
        eRezeptOpStatus = VerkaufAdoRec!OpStatus
        eRezeptPzn = VerkaufAdoRec!pzn
        eRezeptVerordnungsTyp = VerkaufAdoRec!VerordnungsTyp
        eRezeptNoctu = IIf(CheckNullInt(VerkaufAdoRec!ZA_11_Schluessel) = 0, 2.5, 0)
        eRezeptBotendienst = CheckNullDouble(VerkaufAdoRec!Botendienst)
        eRezeptBeschaffungsKosten = CheckNullDouble(VerkaufAdoRec!BeschaffungsKosten)
        
'        MsgBox ("taskid: " + eRezeptTaskId)
        
        RezNr = VerkaufAdoRec!RezeptNr
    End If
    lIkNr = CheckNullLong(VerkaufAdoRec!IkNr)
    Call WinRezDebug("lIkNr: " + CStr(lIkNr))
    If (AutIdemSonderregelOk%) And (lIkNr > 0) Then
        SQLStr$ = "SELECT * FROM kKassen WHERE IK = " + Str$(lIkNr)
        Call WinRezDebug(SQLStr)
'        Set KkassenRec = kKassenDB.OpenRecordset(SQLStr$)
'        If (KkassenRec.RecordCount > 0) Then
        FabsErrf = kKassenDB.OpenRecordset(KkassenRec, SQLStr)
        If Not (KkassenRec.EOF) Then
            AutIdemIk& = CheckNullLong(KkassenRec!IK)
            AutIdemKbvNr& = CheckNullLong(KkassenRec!KbvNr)
            ActAVWGik& = AutIdemIk&
            Call WinRezDebug("AutIdemKbvNr: " + CStr(AutIdemKbvNr))
        End If
    End If

    IvfRezept% = (VerkaufAdoRec!RezeptArt = 7)

    dLastZuzahlung = 0
    Do
        If (VerkaufAdoRec.EOF) Then
            Exit Do
        End If
        
        lCheckVerkaufId = VerkaufAdoRec!Id
           
        sPzn = PznString(VerkaufAdoRec!pzn)
        sText = VerkaufAdoRec!text
        If (VerkaufAdoRec!SonderPreis) Then
            dPreis# = VerkaufAdoRec!vk
            If ((VerkaufAdoRec!FB > 0) And (VerkaufAdoRec!FB < dPreis)) Then
                dPreis# = VerkaufAdoRec!FB
            End If
        Else
            dPreis = VerkaufAdoRec!Preis
        End If
        If (dPreis# = 0) Then
            If (CheckNullLong(VerkaufAdoRec!AbholNr) > 0) Then
                dPreis# = VerkaufAdoRec!vk
                If (VerkaufAdoRec!FB > 0) And (VerkaufAdoRec!FB < dPreis#) Then
                    dPreis# = VerkaufAdoRec!FB
                End If
            End If
        End If
        
'        MsgBox (sPzn + "  " + sText)

        If (RezeptImage = "") Then
            RezeptImage = CheckNullStr(VerkaufAdoRec!rezimage)
        End If
        
        If (Mid$(sText, 15, 4) = "PZN:") Then
            'Stückelung
            'Neue Regelung
            If (InStr(Left(sText, 10), ": ") > 0) Then
                RezArtikel(k).ZuzahlungsErlass = RezArtikel(k).ZuzahlungsErlass + 50
                RezArtikel(k).Zuz = dLastZuzahlung
            End If
        Else
            AktArtikel$ = sPzn + sText + Format(dPreis, "00000.00")
            If (Left$(sText, 5) = "u n t") Then
                If (VerkaufAdoRec!AbholNr > 0) Then
                    If (CheckNullStr(VerkaufAdoRec!AbholerArt) = "R") Then
                        RezArtikel(k%).zusatz(0) = "Abhol-Nr" + Right(Space$(4) + Format(VerkaufAdoRec!AbholNr, "0"), 4)
'                    ElseIf (PrivatRezept) And (CheckNullStr(VerkaufAdoRec!AbholerArt) = "B") Then
'                        RezArtikel(k%).Preis = (VerkaufAdoRec!Aconto + VerkaufAdoRec!Restzahlung) / RezArtikel(k%).Faktor
                    End If
                End If
            ElseIf (AktArtikel$ = LetztArtikel$) And (Val(sPzn) <> 9999999) And _
                ((PrivatRezept% = False) Or ((VerkaufAdoRec!PrivRezNr = AktPrivRezNr) And (Val(VerkaufAdoRec!RezeptNr) = 0))) Then
                RezArtikel(k%).Faktor = RezArtikel(k%).Faktor + VerkaufAdoRec!Anzahl
            Else
                k% = k% + 1
                RezArtikel(k%).pzn = sPzn
                RezArtikel(k%).text = Left$(sText, 28) + " " + CheckNullStr(VerkaufAdoRec!menge) + CheckNullStr(VerkaufAdoRec!einheit)
                RezArtikel(k%).Preis = dPreis
                RezArtikel(k%).Faktor = VerkaufAdoRec!Anzahl
                RezArtikel(k%).flag = 0
                RezArtikel(k%).appli = Abs(InStr(UCase(VerkaufAdoRec!Rezeptzeichen), "A") > 0)  'Abs(VerkaufAdoRec!Hilfsmittel)

                RezArtikel(k%).Zuz = 0
                RezArtikel(k%).HerstRabattPrivat130Brutto = CheckNullDouble(VerkaufAdoRec!HerstRabattPrivat130Brutto)
                RezArtikel(k).VdbPauschale = 0
'                If (IvfRezept) Then
'                    RezArtikel(k%).Zuz = VerkaufAdoRec!Zuzahlung
'                    RezArtikel(k%).Preis = VerkaufAdoRec!VK
'                End If

                RezArtikel(k%).Imp123 = 0
                RezArtikel(k%).Fam = 0
                RezArtikel(k%).AutIdem = 0
                RezArtikel(k%).AutIdemKreuz = CheckNullByte(VerkaufAdoRec!AutIdemKreuz)
                RezArtikel(k%).N0 = 0
                RezArtikel(k%).MagIndex = 0

                RezArtikel(k%).NichtInTaxe = 0
                RezArtikel(k%).ImpErspart = 0

                If (Left$(sText, 5) = "KONTR") Then
                    RezArtikel(k%).flag = RezArtikel(k%).flag Or (REZ_K_STOP Or REZ_BLINK)
                    Mid$(RezArtikel(k%).text, 21, 6) = "(PZN?)"
                    BeepFlag% = True
                End If
                If (Left$(sText, 12) = "PREISEINGABE") Then
                    Mid$(RezArtikel(k%).text, 21, 6) = "(PZN?)"
                End If

                If (Val(RezArtikel(k%).pzn) = 9999999) Then
                    If (RezArtikel(k%).flag And REZ_K_STOP) Then
                    ElseIf (RezArtikel(k%).Preis > 0) Then
                        If (InStr(UCase(VerkaufAdoRec!Kassenzeichen), "A") > 0) Then
                            RezArtikel(k%).appli = 1
                        End If
                    End If
                End If

                RezArtikel(k%).IstWg4 = 0
                RezArtikel(k%).PlusMehrKosten = 0
                RezArtikel(k%).ZuzahlungsErlass = 0
                RezArtikel(k%).Warenzeichen = 0
                RezArtikel(k%).MietDauer = 0
                RezArtikel(k%).TkkPznDruck = 0
                
                RezArtikel(k%).VebNr = CheckNullLong(VerkaufAdoRec!VebNr)
                RezArtikel(k%).PauschaleNr = CheckNullLong(VerkaufAdoRec!PauschaleNr)
                
                If (CheckNullInt(VerkaufAdoRec!Verfügbarkeit) > 1) Then
                    Mid(AMverfügbarkeit2, k + 1, 1) = CStr(VerkaufAdoRec!Verfügbarkeit)
                End If
                If (k = 0) And (UCase$(Left$(OrgRezNr, 1)) = "E") Then
                    eRezeptAMverfügbarkeit = CStr(VerkaufAdoRec!Verfügbarkeit)
                End If

                If (VerkaufAdoRec!AbholNr > 0) Then
                    If (CheckNullStr(VerkaufAdoRec!AbholerArt) = "R") Then
                        RezArtikel(k%).zusatz(0) = "Abhol-Nr" + Right(Space$(4) + Format(VerkaufAdoRec!AbholNr, "0"), 4)
                    ElseIf (PrivatRezept) And (CheckNullStr(VerkaufAdoRec!AbholerArt) = "B") Then
                        If (CheckNullLong(VerkaufAdoRec!ErsetztPzn) = 0) Then
                            RezArtikel(k%).Preis = (VerkaufAdoRec!Aconto + VerkaufAdoRec!Restzahlung) / RezArtikel(k%).Faktor
                        End If
                    End If
                End If

                If (VerkaufAdoRec!Taxmuster) Then
                    SollTmName(k%) = UCase(sText)
                    If (CheckNullLong(VerkaufAdoRec!ParenteralTMId) > 0) Then
                        RezArtikel(k%).Warenzeichen = 98
                        SollTmId(k%) = VerkaufAdoRec!ParenteralTMId
                    Else
                        RezArtikel(k%).Warenzeichen = 99
                        SollTmMenge(k%) = xVal(CheckNullStr(VerkaufAdoRec!menge))
                    End If
                End If

                IstWVORabattArtikel = 0
                On Error Resume Next
'                IstWVORabattArtikel = CheckNullLong(VerkaufAdoRec!WVORabattArtikel)
                IstWVORabattArtikel = CheckNullLong(VerkaufAdoRec!WirkstoffVerordnung)
                On Error GoTo DefErr
                If (IstWVORabattArtikel <> 0) Then
                    Mid(WVORabattArtikel, k + 1, 1) = "2"
                End If
                
                RezArtikel(k%).Fortsetzung = 0
                If (CheckNullLong(VerkaufAdoRec!Fortsetzung)) Then
                    RezArtikel(k%).Fortsetzung = 1
                End If
                
                SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + RezArtikel(k%).pzn
                'Set TaxeRec = TaxeDB.OpenRecordset(SQLStr$)
                On Error Resume Next
                TaxeRec.Close
                Err.Clear
                On Error GoTo DefErr
                TaxeRec.Open SQLStr, taxeAdoDB.ActiveConn
                If (TaxeRec.EOF = False) Then
                    RezArtikel(k%).IstWg4 = CheckZuzZeilenwert
                    If (TaxeRec!VerbandKz) Then
                        RezArtikel(k%).flag = RezArtikel(k%).flag Or (REZ_WG_4 Or REZ_BLINK)
                    ElseIf (VmFlag%) Then
                        If (AplusVOk) Then
                            SQLStr$ = "SELECT * FROM VdbArtikel WHERE PZN = " + sPzn
'                            Set VdbPznRec = AplusVDB.OpenRecordset(SQLStr$)
'                            FabsErrf = Abs(VdbPznRec.EOF)
                            FabsErrf = AplusVDB.OpenRecordset(VdbPznRec, SQLStr)
                            VdbPznRec.Close
                        Else
                            FabsErrf% = VmPzn.IndexSearch(0, sPzn, FabsRecno&)
                        End If
                        If (FabsErrf% = 0) Then
                            RezArtikel(k%).flag = RezArtikel(k%).flag Or (REZ_WG_4 Or REZ_BLINK)
                        End If
                    End If

                    If (TaxeRec!HimiVerbrauch) Then
                        If (CheckNullDouble(VerkaufAdoRec!Zuzages) > 0) Then
                            RezArtikel(k%).Zuz = VerkaufAdoRec!Zuzages / RezArtikel(k%).Faktor
                        Else
                            RezArtikel(k%).Zuz = VerkaufAdoRec!Zuzahlung
                        End If
                    End If

                    If (VerkaufAdoRec!MehrkostenverzichtKZ) Then
                        If (VerkaufAdoRec!Mehrkosten = 0) Then
                            RezArtikel(k%).Preis = VerkaufAdoRec!vk
                        Else
                            RezArtikel(k%).PlusMehrKosten = 1
                        End If
                    End If
                    If (VerkaufAdoRec!ZuzahlungsErlassKZ) Then
                        RezArtikel(k%).ZuzahlungsErlass = 1
                        RezArtikel(k%).Zuz = VerkaufAdoRec!Zuzahlung
                    End If
                    
                    If (VerkaufAdoRec!Mehrkosten > 0) Then
                        If (VerkaufAdoRec!Hilfsmittel) Then
                            RezArtikel(k).zusatz(1) = "MK " + CStr(VerkaufAdoRec!Mehrkosten)
                        End If
                    End If
                    
                    dLastZuzahlung = VerkaufAdoRec!Zuzahlung
                End If

                If (para.MARS) Then
                    Dim SollRezNr$
                    
                    If (VerkaufAdoRec!AbholNr > 0) Then
                        If (AbholerSQL) Then
                            Dim iAdoRec As New ADODB.Recordset
                            
                            If (AbholerConnOpen = 0) Then
                            End If
                            
                            If (PrivatRezept = 0) And (Val(RezNr$) > 0) Then
                                SollRezNr$ = RezNr$
                            Else
                                SollRezNr$ = "P" + Space$(12)
                            End If
                        
                            SQLStr$ = "SELECT * FROM AbholerDetails WHERE AbholerNummer =" + Str$(VerkaufAdoRec!AbholNr)
                            SQLStr = SQLStr + " AND AbholerInd =" + Str(VerkaufAdoRec!AbholID)
                            On Error Resume Next
                            iAdoRec.Close
                            Err.Clear
                            On Error GoTo DefErr
                            iAdoRec.Open SQLStr, AbholerConn    ', adOpenDynamic, adLockOptimistic
                            If Not (iAdoRec.EOF) Then
                                If (CStr(iAdoRec!RezeptNr) = SollRezNr$) Then
                                    If (iAdoRec!Status < 4) Then
                                        RezArtikel(k).flag = RezArtikel(k).flag Or REZ_BLINK
                                        Call MessageBox(RezArtikel(k).text + " (PZN:" + RezArtikel(k).pzn + ")" + vbCrLf + vbCrLf + vbCrLf + " ist noch nicht als ABGEHOLT markiert!" + vbCrLf + vbCrLf + "Das Rezept kann deshalb noch nicht bedruckt werden!", vbInformation, "Problem")
                                    End If
                                End If
                            End If
                            AbholerConn.Close
                        ElseIf (AbholerMdb%) Then
                            Dim iRec As Recordset
                            
                            If (PrivatRezept = 0) And (Val(RezNr$) > 0) Then
                                SollRezNr$ = RezNr$
                            Else
                                SollRezNr$ = "P" + Space$(12)
                            End If
                        
                            Set AbholerDB = OpenDatabase(ABHOLER_DB, False, True)
                            SQLStr$ = "SELECT * FROM AbholerDetails WHERE AbholerNummer =" + Str$(VerkaufAdoRec!AbholNr)
                            SQLStr = SQLStr + " AND AbholerInd =" + Str(VerkaufAdoRec!AbholID)
                            Set iRec = AbholerDB.OpenRecordset(SQLStr$)
                            If Not (iRec.EOF) Then
                                If (CStr(iRec!RezeptNr) = SollRezNr$) Then
                                    If (iRec!Status < 4) Then
                                        RezArtikel(k).flag = RezArtikel(k).flag Or REZ_BLINK
                                        Call MessageBox(RezArtikel(k).text + " (PZN:" + RezArtikel(k).pzn + ")" + vbCrLf + vbCrLf + vbCrLf + " ist noch nicht als ABGEHOLT markiert!" + vbCrLf + vbCrLf + "Das Rezept kann deshalb noch nicht bedruckt werden!", vbInformation, "Problem")
                                    End If
                                End If
                            End If
                            
                            AbholerDB.Close
                        End If
                        
                    End If
                    
                End If
        
'                PrivatBesorger% = 0
            End If
        End If

        LetztArtikel$ = AktArtikel$

'        If ((VerkaufAdoRec!KundenEnde) Or (VerkaufAdoRec!RezeptEnde)) Or _
            ((PrivatRezept%) And ((VerkaufAdoRec!PrivRezNr <> AktPrivRezNr) Or (Val(VerkaufAdoRec!RezeptNr) > 0))) Then
        If (VerkaufAdoRec!RezeptEnde) Then

            VerkDatum$ = Format(VerkaufAdoRec!Datum, "DDMMYY")
            OrgVerkDatum$ = VerkDatum$
            
            VerkaufDat = VerkaufAdoRec!Datum
            MarsAutomaticDatum(0) = VerkaufDat

'          DruckDatum = 0L;

            VerkZeit% = Val(Format(VerkaufAdoRec!Datum, "HHMM"))

            ActVerordnung% = 1
            If (VerkaufAdoRec!RezeptArt = REZEPTART_SPRECHSTUNDENBEDARF) Then ActVerordnung% = 2

            iUser% = VerkaufAdoRec!Computer
            VerkUser$ = Format(iUser%, " 0")
            VerkLiRe = VerkaufAdoRec!Seite

            iKuNr = VerkaufAdoRec!KundenNr
            If (iKuNr > 0) Then
                Call HoleKundenInfo(iKuNr)
            End If

            If (VerkaufAdoRec!RezeptArt = REZEPTART_ZUZAFREI) Or (VerkaufAdoRec!RezeptArt = REZEPTART_BTM_ZUZAFREI) Or (VerkaufAdoRec!RezeptArt = REZEPTART_SPRECHSTUNDENBEDARF) Then
                RezGebSumme# = 0#
                GebFrei% = True
            End If

            If (VerkaufAdoRec!AVKassentyp > 0) Then
                ActKasse% = VerkaufAdoRec!AVKassentyp
            End If

            ActKkasse$ = ""
            If (VerkaufAdoRec!AVKassenNr > 0) Then
                ActKkasse$ = Right$(String(9, "0") + Format(VerkaufAdoRec!AVKassenNr), 9) 'Trim(VK.KKNr)
            End If

            ActVebNr = CheckNullLong(VerkaufAdoRec!VebNr)
            ActPauschaleNr = CheckNullLong(VerkaufAdoRec!PauschaleNr)

            If (PrivatRezept) Then
                ActKasse% = 0
                ActKkasse$ = ""
            End If
            
            bmPrivEnde = VerkaufAdoRec.Bookmark
            k% = k% + 1

            Exit Do
        End If

        VerkaufAdoRec.MoveNext
    Loop

    AnzRezeptArtikel% = k%
    If (AnzRezeptArtikel = 0) Then
'        If (bCheckRezeptEnde) And (lCheckVerkaufId <> 0) Then
'            MsgBox (CStr(lCheckVerkaufId))
'        End If
        
        RezeptHolenDB% = False
        Call DefErrPop: Exit Function
    End If

    
    If (Val(Mid(OrgVerkDatum, 3, 2)) <> Month(Now)) Then
        Dim AltTaxeDatum$
        
        AltTaxeDatum = CheckAlteTaxe
        If (AltTaxeDatum = "") Then
            Call MessageBox("Achtung: die zum Zeitpunkt der Abgabe am " + Left(OrgVerkDatum, 2) + "." + Mid(OrgVerkDatum, 3, 2) + ".20" + Mid(OrgVerkDatum, 5, 2) + " gültige Taxe ist nicht mehr verfügbar!", vbInformation, "Rezeptkontrolle")
        Else
            Call MessageBox("Hinweis:" + vbCrLf + "zum Zeitpunkt der Abgabe am " + Left(OrgVerkDatum, 2) + "." + Mid(OrgVerkDatum, 3, 2) + ".20" + Mid(OrgVerkDatum, 5, 2) + " gültige Taxe:  " + AltTaxeDatum + vbCrLf, vbInformation, "Rezeptkontrolle")
            
            taxeAdoDB.CloseDB
            AplusVDB.CloseDB
            AutIdemDB.CloseDB
            kKassenDB.CloseDB
            
            taxeAdoDB.OpenDB (AltTaxeDatum)
            AplusVDB.OpenDB (AltTaxeDatum)
            AutIdemDB.OpenDB (AltTaxeDatum)
            kKassenDB.OpenDB (AltTaxeDatum)
            
            AlteTaxeAktiv = True
        End If
    End If
    
    SonderBelegRezept% = False
    If (AnzRezeptArtikel% = 1) Then
        For i% = 1 To AnzSonderBelege%
            If (SonderBelege(i% - 1).pzn = RezArtikel(0).pzn) Then
                SonderBelegRezept% = i%
                PrivatRezept% = 0
                BtmRezept% = 0
                IvfRezept% = 0
                RezKundenInfo$(1) = ""
                RezKundenInfo$(2) = "SONDERBELEG"
                Exit For
            End If
        Next i%
    End If


    If (UCase$(Left$(OrgRezNr, 1)) = "E") Then
    ElseIf Not (PrivatRezept%) And (SonderBelegRezept% = 0) Then
'        Set VerkaufRec = VerkaufDB.OpenRecordset("Verkauf", dbOpenTable)
'        VerkaufRec.index = "Unique"

        VerkaufAdoRec.Bookmark = bmPrivEnde
        Do
            If (VerkaufAdoRec.BOF) Then
                Exit Do
            End If
            If (VerkaufAdoRec!KundenEnde) Then
                Exit Do
            End If
            VerkaufAdoRec.MovePrevious
        Loop

        SQLStr = "SELECT * FROM Verkauf WHERE Datum='" + Format(VerkaufDat, "DD.MM.YYYY HH:MM:SS") + "'"
        SQLStr = SQLStr + " AND Computer=" + VerkUser
        SQLStr = SQLStr + " AND Seite=" + CStr(VerkLiRe)
        SQLStr = SQLStr + " ORDER BY ZeilenNr"
        FabsErrf = VerkaufAdoDB.OpenRecordset(VerkaufAdoRec, SQLStr, 0)
            
        i% = 0
        Do
            If (VerkaufAdoRec.EOF) Then
                Exit Do
            End If

            If (VerkaufAdoRec!RezeptArt = 0) And (VerkaufAdoRec!Anzahl <> 0) Then
                hStr$ = VerkaufAdoRec!text + " " + Format(VerkaufAdoRec!Preis, "0.00")
                BarVerkauf$(i%) = hStr$
                i% = i% + 1
                If (i% > UBound(BarVerkauf$)) Then
                    Exit Do
                End If
            End If

            If (VerkaufAdoRec!KundenEnde) Then
                Exit Do
            End If
            
            VerkaufAdoRec.MoveNext
        Loop
    End If

    If (PrivatRezept%) Or (SonderBelegRezept%) Then
    Else
        AMverfügbarkeit = Left$(String(AnzRezeptArtikel%, "1") + "000", 3)
        For i = 1 To AnzRezeptArtikel%
            If (Val(Mid(AMverfügbarkeit2, i, 1)) > 1) Then
                Mid(AMverfügbarkeit, i, 1) = Mid(AMverfügbarkeit2, i, 1)
            End If
        Next i
    End If

'MsgBox (CStr(AnzRezeptArtikel))
    For i% = 0 To (AnzRezeptArtikel% - 1)
        dZuz# = 0#
        dAvp# = 0

        If (PrivatRezept%) Or (SonderBelegRezept%) Then
            RezArtikel(i%).flag = RezArtikel(i%).flag And (REZ_BLINK Xor &HFF)
            If (PrivatRezept) And (Val(RezArtikel(i%).pzn) <> 9999999) Then
                SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + RezArtikel(i%).pzn
                'Set TaxeRec = TaxeDB.OpenRecordset(SQLStr$)
                On Error Resume Next
                TaxeRec.Close
                Err.Clear
                On Error GoTo DefErr
                TaxeRec.Open SQLStr, taxeAdoDB.ActiveConn
                If (TaxeRec.EOF) Then
                    RezArtikel(i%).NichtInTaxe = 1

                    If (ArtikelDbOk) Then
                        SQLStr$ = "SELECT * FROM ARTIKEL WHERE PZN = " + RezArtikel(i%).pzn
    '                    SQLStr = SQLStr + " AND LagerKz<>0"
                        FabsErrf = Artikel.OpenRecordset(ArtikelAdoRec, SQLStr)
                    Else
                        FabsErrf% = ast.IndexSearch(0, RezArtikel(i%).pzn, FabsRecno&)
                        If (FabsErrf% = 0) Then
                            ast.GetRecord (FabsRecno& + 1)
                        End If
                    End If
                    If (FabsErrf% = 0) Then
                        If (Val(ast.wg) = 3) Then
                            RezArtikel(i).pzn = "09999011"
                            RezArtikel(i).text = "Rezeptur"
                        End If
                    End If
                ElseIf (TaxeRec!BtmKz) Then
                    '4.0.52: aktiviert
                    BtmRezept% = True
                End If
            End If
        Else

            If (Val(RezArtikel(i%).pzn) = 9999999) Then
                If (RezArtikel(i%).flag And REZ_K_STOP) Then
                    dZuz# = 5   'ZuzFeld!(3)
                    RezArtikel(i%).flag = RezArtikel(i%).flag Or REZ_BLINK
                ElseIf (RezArtikel(i%).Preis > 0) Then
                    If (RezArtikel(i%).appli = 0) Then
                        dZuz# = CalcZuzaNeu(RezArtikel(i%).Preis)
                    End If

                    RezArtikel(i%).flag = RezArtikel(i%).flag Or REZ_BLINK

                    ind2% = InStr(RezArtikel(i%).text, "(PZN?)")
                    If (ind2% <= 0) Then
                        ch$ = Trim(RezArtikel(i%).text) + "   " + "(PZN?)"
                        RezArtikel(i%).text = Left$(ch$ + Space$(36), 36)
                    End If
                End If
            Else
                SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + RezArtikel(i%).pzn
                'Set TaxeRec = TaxeDB.OpenRecordset(SQLStr$)
                On Error Resume Next
                TaxeRec.Close
                Err.Clear
                On Error GoTo DefErr
                TaxeRec.Open SQLStr, taxeAdoDB.ActiveConn
                If (TaxeRec.EOF) Then
                    RezArtikel(i%).NichtInTaxe = 1
                    
                    If (ArtikelDbOk) Then
                        SQLStr$ = "SELECT * FROM ARTIKEL WHERE PZN = " + RezArtikel(i%).pzn
    '                    SQLStr = SQLStr + " AND LagerKz<>0"
                        FabsErrf = Artikel.OpenRecordset(ArtikelAdoRec, SQLStr)
                    Else
                        FabsErrf% = ast.IndexSearch(0, RezArtikel(i%).pzn, FabsRecno&)
                        If (FabsErrf% = 0) Then
                            ast.GetRecord (FabsRecno& + 1)
                        End If
                    End If
                    
                    If (FabsErrf% = 0) Then
                        If (ast.rez = "SG") Then
                            BtmRezept% = True
                        End If
                    End If
                    
                    If (CheckSonderPzn%(RezArtikel(i%).pzn) >= 0) Then
'                        dZuz# = ZuzFeld!(3)
                        If (Val(RezArtikel(i%).pzn) <> 2567018) Then
                            dZuz# = CalcZuzaNeu(RezArtikel(i%).Preis)
                        End If
                    Else
                        If (FabsErrf% = 0) Then
                            If (Val(ast.wg) = 3) Then
                                RezArtikel(i).pzn = "09999011"
                                RezArtikel(i).text = "Rezeptur"
                                dZuz# = CalcZuzaNeu(RezArtikel(i%).Preis)
                            End If
                        End If
                    End If
                Else
'                    If (TaxeRec!Warengruppe = 3) Then
'                        RezArtikel(i%).IstWg4 = 1
'                    End If

                    dAvp# = TaxeRec!vk / 100#
                    If (IvfRezept) Then
                        If (TaxeRec!FESTBETRAG > 0) And (TaxeRec!FESTBETRAG < TaxeRec!vk) Then
                            dAvp# = (TaxeRec!FESTBETRAG / 100#)
                        End If
                    End If

                    dPreis# = RezArtikel(i%).Preis
                    If (RezArtikel(i%).PlusMehrKosten) Then
                        If (TaxeRec!FESTBETRAG > 0) And (TaxeRec!FESTBETRAG < TaxeRec!vk) Then
                            dPreis# = (TaxeRec!FESTBETRAG / 100#)
                        Else
                            RezArtikel(i%).PlusMehrKosten = 0
                        End If
                    End If

                    If (RezArtikel(i%).ZuzahlungsErlass) Then
                        dZuzDos# = CalcTaxeZuzahlungDOS(dPreis#, RezArtikel(i%).Zuz + 200)
                    Else
                        dZuzDos# = CalcTaxeZuzahlungDOS(dPreis#)
                    End If

                    If (TaxeRec!Negliste) Then dZuzDos# = 1#
                    If (RezArtikel(i%).appli) Then dZuzDos# = 0#

                    If (RezArtikel(i%).PlusMehrKosten) Then
                        RezArtikel(i%).Preis = TaxeRec!vk / 100#
                    End If

                    If (dZuzDos# >= 1000#) Then
                        dZuz# = dZuzDos# - 1000#
                    ElseIf Not (GebFrei%) And (dZuzDos# > 0#) Then
                        RezArtikel(i%).Preis = 0#
                        If (dZuzDos# = 3#) Then
    '                        Call iMsgBox("Artikel unter Zuzahlung, Geldrückgabe ?")
                        Else
                            RezArtikel(i%).flag = RezArtikel(i%).flag And (REZ_BLINK Xor &HFF)
                        End If
                    End If
                    If (TaxeRec!HimiVerbrauch) Then
                        dZuz# = RezArtikel(i%).Zuz '/ RezArtikel(i%).Faktor
                    End If

                    If (RezArtikel(i%).ZuzahlungsErlass) Then
                        If (RezArtikel(i).ZuzahlungsErlass >= 50) Then
                            RezArtikel(i).ZuzahlungsErlass = RezArtikel(i).ZuzahlungsErlass - 50
                        End If
                        dZuz# = RezArtikel(i%).Zuz
                    End If

                    If (TaxeRec!BtmKz) Then
                        BtmRezept% = True
                    End If

                    hStr$ = Format$(TaxeRec!Abgabebest, "0")
                    If (InStr("01245", hStr$) <> 0) Then
                        RezArtikel(i%).Fam = 1
                    End If

                    Call CheckAutIdem(i%)
                    SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + RezArtikel(i%).pzn
                    'Set TaxeRec = TaxeDB.OpenRecordset(SQLStr$)
                    On Error Resume Next
                    TaxeRec.Close
                    Err.Clear
                    On Error GoTo DefErr
                    TaxeRec.Open SQLStr, taxeAdoDB.ActiveConn

                    If (Left$(TaxeRec!Zuzahlung, 1) = "9") Then
                        RezArtikel(i%).IstWg4 = 0
                    End If

                    If (ImpFlag% = 0) Then
                        RezArtikel(i%).Imp123 = 0
                    ElseIf (TaxeRec!OriginalPZN = TaxeRec!pzn) Then
                        RezArtikel(i%).Imp123 = 6
                        OrgAvp = TaxeRec!vk / 100#
                        SQLStr$ = "SELECT * FROM TAXE WHERE ORIGINALPZN = " + PznString(TaxeRec!OriginalPZN)
                        'Set TaxeRec = TaxeDB.OpenRecordset(SQLStr$)
                        On Error Resume Next
                        TaxeRec.Close
                        Err.Clear
                        On Error GoTo DefErr
                        TaxeRec.Open SQLStr, taxeAdoDB.ActiveConn
                        TaxeRec.MoveFirst
                        Do While Not TaxeRec.EOF
                          If (TaxeRec!OriginalPZN <> TaxeRec!pzn) Then
                            aktavp = TaxeRec!vk / 100#
                            If (aktavp <= OrgAvp - 15) Or (aktavp <= OrgAvp * 0.85) Then
                              RezArtikel(i%).Imp123 = 5
                              Exit Do
                            End If
                          End If
                          TaxeRec.MoveNext
                        Loop
                        SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + RezArtikel(i%).pzn
                        'Set TaxeRec = TaxeDB.OpenRecordset(SQLStr$)
                        On Error Resume Next
                        TaxeRec.Close
                        Err.Clear
                        On Error GoTo DefErr
                        TaxeRec.Open SQLStr, taxeAdoDB.ActiveConn

                    ElseIf (TaxeRec!OriginalPZN > 0) Then
                        aktavp = TaxeRec!vk / 100
                        SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + PznString(TaxeRec!OriginalPZN)
                        'Set TaxeRec = TaxeDB.OpenRecordset(SQLStr$)
                        On Error Resume Next
                        TaxeRec.Close
                        Err.Clear
                        On Error GoTo DefErr
                        TaxeRec.Open SQLStr, taxeAdoDB.ActiveConn
                        If Not (TaxeRec.EOF) Then
                          OrgAvp = TaxeRec!vk / 100
                        End If
                        If (aktavp <= OrgAvp - 15) Or (aktavp <= OrgAvp * 0.85) Then
                          RezArtikel(i%).Imp123 = 2
                          RezArtikel(i%).ImpErspart = (OrgAvp - aktavp)
                        Else
                          RezArtikel(i%).Imp123 = 1
                        End If
                        SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + RezArtikel(i%).pzn
                        'Set TaxeRec = TaxeDB.OpenRecordset(SQLStr$)
                        On Error Resume Next
                        TaxeRec.Close
                        Err.Clear
                        On Error GoTo DefErr
                        TaxeRec.Open SQLStr, taxeAdoDB.ActiveConn
                    End If
                    
                    If (PreisDiffAktiv) Then
                        AVP# = TaxeRec!vk / 100
                        FESTBETR# = TaxeRec!FESTBETRAG / 100
                        If ((TaxeRec!FestKz) And (FESTBETR# < AVP#)) Then
                            VglWert# = FESTBETR#
                        Else
                            VglWert# = AVP#
                        End If
                        If (VglWert > 0) Then
                            If (RezArtikel(i%).Preis <> VglWert) Then
                                h = RezArtikel(i).pzn + "  " + RezArtikel(i).text + "  " + Format(RezArtikel(i).Preis, "0.00") + "  " + Format(VglWert, "0.00")
                                PreisDiff_Log% = FreeFile
                                Open "PreisDif.log" For Append As #PreisDiff_Log
                                Print #PreisDiff_Log%, Format(Now, "DD.MM.YYYY HH:MM:SS  ") + RezNr + "  " + h$
                                Close #PreisDiff_Log%
                                
                                PreisDiffStr = PreisDiffStr + h + vbCrLf
                            End If
                        End If
                    End If
                    
                End If
            End If
        End If

        If (IvfRezept) Then
            RezArtikel(i%).Zuz = 0
            RezArtikel(i%).Preis = FNX(dAvp / 2)
        Else
            RezArtikel(i%).Zuz = dZuz#
        End If
    Next i%

    If (IvfRezept) Then
        If (Val(RezArtikel(0).pzn) <> 9999643) Then
            For i% = UBound(RezArtikel) - 1 To 0 Step -1
                RezArtikel(i% + 1) = RezArtikel(i%)
            Next i%
            Call InitRezeptArtikel(0)
'            Call InitKontrollStop(row%)

            RezArtikel(0).pzn = "09999643"
            RezArtikel(0).text = "IVF-Rezept"
            RezArtikel(0).Faktor = 1
            RezArtikel(0).Preis = 0
            RezArtikel(0).Zuz = 0

            AnzRezeptArtikel% = AnzRezeptArtikel% + 1
        End If
    End If

    If (BtmRezept%) Then
        If (PrivatRezept) Then
            '4.0.52: unter Kommentar gesetzt
'            PrivatRezept = 0
        End If
        For i% = 0 To UBound(RezAddGesamtBrutto$)
            RezAddGesamtBrutto$(i%) = ""
        Next i%
        Call PruefeBtmKosten
        sBrutto$ = RezAddGesamtBrutto$(0)
    End If

    Call PruefeRezkontrDat

    If (ActVebNr = 0) And (ActPauschaleNr = 0) Then
        If (AutIdemIk& > 0) Then
            SQLStr = "SELECT count(*) AS iAnz FROM VdbVVI LEFT JOIN VdbVVL ON VdbVVI.VebNr=VdbVVL.VebNr"
            SQLStr = SQLStr + " WHERE VdbVVI.Ik=" + CStr(AutIdemIk)
            SQLStr = SQLStr + " AND VdbVVL.LandNr=" + CStr(ActBundesland)
            FabsErrf = AplusVDB.OpenRecordset(VdbBedRec, SQLStr)
            If Not (VdbBedRec.EOF) Then
                If (CheckNullLong(VdbBedRec!iAnz) = 1) Then
                    SQLStr = "SELECT VdbVVI.VebNr FROM VdbVVI LEFT JOIN VdbVVL ON VdbVVI.VebNr=VdbVVL.VebNr"
                    SQLStr = SQLStr + " WHERE VdbVVI.Ik=" + CStr(AutIdemIk)
                    SQLStr = SQLStr + " AND VdbVVL.LandNr=" + CStr(ActBundesland)
                    FabsErrf = AplusVDB.OpenRecordset(VdbBedRec, SQLStr)
                    If Not (VdbBedRec.EOF) Then
                        ActVebNr = CheckNullLong(VdbBedRec!VebNr)
                    End If
                End If
            End If
        End If
        If (ActVebNr = 0) And (ActKasse% > 0) Then
            SQLStr = "SELECT Top 1* FROM VdbBedingungen"
            SQLStr = SQLStr + " WHERE KostentrNr=" + CStr(ActKasse)
            SQLStr = SQLStr + " AND LandNr=" + CStr(ActBundesland)
            FabsErrf = AplusVDB.OpenRecordset(VdbBedRec, SQLStr)
            If Not (VdbBedRec.EOF) Then
                ActVebNr = CheckNullLong(VdbBedRec!VebNr)
            End If
        End If
    End If
    If (ActKasse% < 0) And (AutIdemIk& > 0) Then
        SQLStr = "SELECT Top 1* FROM VdbVIK"
        SQLStr = SQLStr + " WHERE IK=" + CStr(AutIdemIk)
        SQLStr = SQLStr + " ORDER BY KostentrNr"
        FabsErrf = AplusVDB.OpenRecordset(VdbBedRec, SQLStr)
        If Not (VdbBedRec.EOF) Then
            ActKasse = CheckNullLong(VdbBedRec!KostentrNr)
        End If
    End If
    
    If (ActVebNr > 0) Then
        For j% = 0 To UBound(AvVereinbarungen$)
            h2 = AvVereinbarungen(j)
            If (Val(Left(h2, 10)) = ActVebNr) Then
                If (InStr(UCase(Trim(Mid$(h2, 12, 100))), "BERUFSGENOSSENSCHAFT") > 0) Then
                    IkSpezialName = Trim(Mid$(h2, 12, 100))
                End If
                
                Exit For
            End If
        Next j%
    End If

    If (ActKasse% > 0) Or (ActVebNr > 0) Or (ActPauschaleNr > 0) Then
        For i% = 0 To (AnzRezeptArtikel% - 1)
            If (RezArtikel(i%).flag And REZ_WG_4) Then
                RezeptHolenZuz# = RezArtikel(i%).Zuz
                AktTabIndex% = 3 + i% * 4
                Call Verbandmittel
            
                If (PreisDiffAktiv) Then
                    SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + RezArtikel(i%).pzn
                    'Set TaxeRec = TaxeDB.OpenRecordset(SQLStr$)
                    On Error Resume Next
                    TaxeRec.Close
                    Err.Clear
                    On Error GoTo DefErr
                    TaxeRec.Open SQLStr, taxeAdoDB.ActiveConn
                    If Not (TaxeRec.EOF) Then
                        AVP# = TaxeRec!vk / 100
                        FESTBETR# = TaxeRec!FESTBETRAG / 100
                        If ((TaxeRec!FestKz) And (FESTBETR# < AVP#)) Then
                            VglWert# = FESTBETR#
                        Else
                            VglWert# = AVP#
                        End If
'                        If (VglWert > 0) Then
                            If (RezArtikel(i%).Preis <> VglWert) Then
                                h = RezArtikel(i).pzn + "  " + RezArtikel(i).text + "  " + Format(RezArtikel(i).Preis, "0.00") + "  " + Format(VglWert, "0.00")
                                PreisDiff_Log% = FreeFile
                                Open "PreisDif.log" For Append As #PreisDiff_Log
                                Print #PreisDiff_Log%, Format(Now, "DD.MM.YYYY HH:MM:SS  ") + RezNr + "  " + h$ + "  A+V"
                                Close #PreisDiff_Log%
                                
'                                PreisDiffStr = PreisDiffStr + h + vbCrLf
                            End If
'                        End If
                    End If
                End If
            End If
        Next i%
    End If
    RezeptHolenZuz# = 0
End If

Call WinRezDebug("Ende RezeptHolenDB: PersonalNr: " + CStr(PersonalNr))
Call WinRezDebug("Ende RezeptHolenDB: AutIdemKbvNr: " + CStr(AutIdemKbvNr))

Call CheckVerordnungMDB

'ActVerordnung% = 1

If (MarsModus = MARS_REZEPT_KONTROLLE) Then
'    Shell CurDir + "\OCR2016.EXE " + "ABC" + " 11", vbMinimizedNoFocus
    If (RezeptImage <> "") Then
        RezeptImage = CurDir + "\KUNSCAN\" + RezeptImage
'        Shell CurDir + "\OCR2016.EXE " + RezeptImage + " 19", vbMinimizedNoFocus
        Shell CurDir + "\OpRezept " + RezeptImage + ",1", vbNormalNoFocus
        DoEvents
        SeriellPause
        AppActivate ParentForm.Caption
        DoEvents
    End If
End If

If (eRezeptTaskId <> "") Then
    eRezeptStatusWerte = RezArtikel(0).pzn + eRezeptAMverfügbarkeit + Format(RezArtikel(0).Preis, "0.00") + Format(RezArtikel(0).Zuz, "0.00") + VerkDatum
    eRezeptSonderPZN = 0
    Dim lPzn&
    For i% = 0 To (AnzRezeptArtikel% - 1)
        lPzn = Val(RezArtikel(i%).pzn)
        If (lPzn = 9999637) Or (lPzn = 6461110) Then  'Beschaffungskosten 'Botendienst
            eRezeptSonderPZN = True
            Exit For
        End If
    Next i
    
    Dim iNoctuVoran As Integer
    iNoctuVoran = -1
    For i% = 0 To (AnzRezeptArtikel% - 2)
        lPzn = Val(RezArtikel(i%).pzn)
        If (lPzn = 2567018) Then 'Noctu
            iNoctuVoran = 0
            Exit For
        End If
    Next i
    If (iNoctuVoran >= 0) Then
        Dim iRezArtikel As REZEPT_ARTIKEL_STRUCT
        iRezArtikel = RezArtikel(iNoctuVoran)
        Dim sAMverfügbarkeit As String
        sAMverfügbarkeit = Mid(AMverfügbarkeit2, iNoctuVoran + 1, 1)
        For i% = (iNoctuVoran + 1) To (AnzRezeptArtikel% - 1)
            RezArtikel(i - 1) = RezArtikel(i)
            Mid(AMverfügbarkeit, i, 1) = Mid(AMverfügbarkeit, i + 1, 1)
        Next i
        RezArtikel(AnzRezeptArtikel% - 1) = iRezArtikel
        Mid(AMverfügbarkeit, AnzRezeptArtikel%, 1) = sAMverfügbarkeit
    End If
    
    If (eRezeptNoctu > 0) Then
        Call CheckPznVorhanden(2567018, "Notdienstgebühr", 2.5)
    End If
    If (eRezeptBotendienst > 0) Then
        Call CheckPznVorhanden(6461110, "Botendienst", eRezeptBotendienst)
    End If
    If (eRezeptBeschaffungsKosten > 0) Then
        Call CheckPznVorhanden(9999637, "Beschaffungskosten", eRezeptBeschaffungsKosten)
    End If
    

    With ParentToolbar
        .ImageInd(5) = 18
        .LoadToolbarButton (5)
        .ImageInd(9) = 19
        .LoadToolbarButton (9)
        .ImageInd(10) = 17
        .LoadToolbarButton (10)
        .ImageInd(13) = 20
        .LoadToolbarButton (13)
        .ImageInd(14) = 17
        .LoadToolbarButton (14)
        .ShowToolbar (.Size)
    End With
    With ParentForm
        .cmdToolbar(5).ToolTipText = "F6 Dispensieren und Einreichen eRezept"
        .cmdToolbar(9).ToolTipText = "shift+F2 Info zu eRezept"
        .cmdToolbar(13).ToolTipText = "shift+F6 Vorabprüfung"
        .cmdToolbar(14).ToolTipText = "shift+F7 "
    End With
        
    With ParentForm.flxarbeit(0)
        Dim x&
        
        If (eRezeptOpStatus >= 3) Then
            Dim lColor&
            
            lColor = RGB(255, 216, 0)
            If (eRezeptOpStatus = 5) Then
                lColor = vbRed
            ElseIf (eRezeptOpStatus = 6) Then
                lColor = vbGreen
            End If
            
            ParentForm.picBack(0).FillColor = lColor ' RGB(232, 217, 172)
            x = ParentForm.txtRezeptNr.Left + ParentForm.txtRezeptNr.Width + 300
            RoundRect ParentForm.picBack(0).hdc, (.Left) / Screen.TwipsPerPixelX, (.Top / Screen.TwipsPerPixelY) - 5, (.Left + x) / Screen.TwipsPerPixelX, (.Top + .Height) / Screen.TwipsPerPixelX + 5, 20, 20
    '        Call wpara.FillGradient(ParentForm.picBack(0), (.Left) / Screen.TwipsPerPixelX + 1, ((.Top + .Height / 2) / Screen.TwipsPerPixelY), (.Left + x) / Screen.TwipsPerPixelX - 2, (.Top + .Height) / Screen.TwipsPerPixelY - 60, RGB(232, 217, 172), RGB(242, 227, 182))
            Call wpara.FillGradient(ParentForm.picBack(0), (.Left) / Screen.TwipsPerPixelX + 1, ((.Top + .Height / 2) / Screen.TwipsPerPixelY), (.Left + x) / Screen.TwipsPerPixelX - 2, (.Top + .Height) / Screen.TwipsPerPixelY - 60, lColor, lColor)
        End If
    End With
    
'    With ParentForm.flxInfo(0)
'        .Cols = 6
'        .Rows = 0
'        .Rows = 2
'        .FixedRows = 1
'
'        Dim sFormatStr$
'    '    sFormatStr = "|<Gescannt|<TaskId|<PrescriptionId|<AnzahlVerordnetePackungen|<Pzn|<Packung|Packungsgroesse|<Einheit|<Darreichungsform|^N123|^AI#|<KK|<KostentraegerName|"
'        sFormatStr = "^|<Datum|<Aktion|<Ergebnis||"
'        .FormatString = sFormatStr
'
'        .GridLines = flexGridFlat
'        .SelectionMode = flexSelectionByRow
'        .HighLight = flexHighlightWithFocus
'
'        .ScrollBars = flexScrollBarVertical
'
'        Dim bOk%, iAktion%
'        Dim sAktion$, sTag$
'        .Rows = .FixedRows
'        SQLStr = "SELECT * FROM TI_Aktionen"
'        SQLStr = SQLStr + " WHERE (TaskId='" + ParentForm.flxarbeit(0).TextMatrix(.row, 2) + "')"
'        SQLStr = SQLStr + " ORDER BY AnlageDatum"
'        FabsErrf = VerkaufAdoDB.OpenRecordset(VerkaufAdoRec, SQLStr, 0)
'        'If (FabsErrf <> 0) Then
'        '    Call iMsgBox("keine passenden Rezepte gespeichert !")
'        '    Call DefErrPop: Exit Sub
'        'End If
'        Do
'            If (VerkaufAdoRec.EOF) Then
'                Exit Do
'            End If
'
'            sAktion = CheckNullStr(VerkaufAdoRec!Aktion)
'
'            bOk = True
'            If (UCase(sAktion) = "ABGEBEN") Then
'                bOk = (CheckNullStr(VerkaufAdoRec!DispenseXml) <> "") And (CheckNullStr(VerkaufAdoRec!ResultXml) <> "")
'            ElseIf (UCase(sAktion) = "ABRUF") Then
'                bOk = (CheckNullStr(VerkaufAdoRec!BundleXml) <> "")
'            End If
'
'            h = IIf(bOk, Chr(214), "")
'            h = h + vbTab + Format(CheckNullDate(VerkaufAdoRec!Anlagedatum), "DD.MM.YY HH:mm")
'            h = h + vbTab + sAktion
'            h = h + vbTab
'            h = h + vbTab + Format(CheckNullDate(VerkaufAdoRec!Anlagedatum), "YYMMDDHHmmss") + Format(VerkaufAdoRec!Id, "000000")
'            .AddItem h, 1
'
'            VerkaufAdoRec.MoveNext
'        Loop
'        VerkaufAdoRec.Close
'
'        SQLStr = "SELECT * FROM TI_FiveRx" '
'        SQLStr = SQLStr + " WHERE (TaskId='" + ParentForm.flxarbeit(0).TextMatrix(.row, 2) + "')"
'        SQLStr = SQLStr + " ORDER BY AnlageDatum"
'        FabsErrf = VerkaufAdoDB.OpenRecordset(VerkaufAdoRec, SQLStr, 0)
'        Do
'            If (VerkaufAdoRec.EOF) Then
'                Exit Do
'            End If
'
'            iAktion = CheckNullInt(VerkaufAdoRec!AktionInt)
'            If (iAktion = 0) Or (iAktion = 2) Or (iAktion = 4) Then
'                If (iAktion = 0) Then
'                    sAktion = "Vorab"
'                ElseIf (iAktion = 2) Then
'                    sAktion = "Einreichung"
'                Else
'                    sAktion = "Ergebnis"
'                End If
'
'                h = ""
'                h = h + vbTab + Format(CheckNullDate(VerkaufAdoRec!Anlagedatum), "DD.MM.YY HH:mm")
'                h = h + vbTab + sAktion
'                h = h + vbTab
'                h = h + vbTab + Format(CheckNullDate(VerkaufAdoRec!Anlagedatum), "YYMMDDHHmmss") + Format(VerkaufAdoRec!Id, "000000")
'                .AddItem h, 1
'            Else
'                sAktion = CheckNullStr(VerkaufAdoRec!FiveRxXml)
'                .TextMatrix(1, 5) = sAktion
'                If (iAktion = 1) Or (iAktion = 5) Then
'                    sTag = "STATUS"
'                Else
'                    sTag = "RZLIEFERID"
'                End If
'        '        ind = InStr(UCase(sAktion), "<" + sTag + ">")
'        '        If (ind > 0) Then
'        '            sAktion = Mid(sAktion, ind + Len(sTag) + 1 + 1)
'        '            ind = InStr(UCase(sAktion), "</" + sTag + ">")
'        '            If (ind > 0) Then
'        '                sAktion = Left(sAktion, ind - 1)
'        '                flxInfo(0).TextMatrix(1, 3) = sAktion
'        '
'        '                bOk = True
'        '                If (iAktion = 1) Or (iAktion = 5) Then
'        '                    bOk = (UCase(sAktion) <> "FEHLER")
'        '                End If
'        '                If (bOk) Then
'        '                    flxInfo(0).TextMatrix(1, 0) = Chr(214)
'        '                End If
'        '            End If
'        '        End If
'                sAktion = XmlAbschnitt(sAktion, sTag)
'                If (sAktion <> "") Then
'                    .TextMatrix(1, 3) = sAktion
'
'                    bOk = True
'                    If (iAktion = 1) Or (iAktion = 5) Then
'                        bOk = (UCase(sAktion) <> "FEHLER")
'                    End If
'                    If (bOk) Then
'                        .TextMatrix(1, 0) = Chr(214)
'                    End If
'                End If
'            End If
'
'        '    bOk = True
'        ''    If (UCase(sAktion) = "ABGEBEN") Then
'        ''        bOk = (CheckNullStr(VerkaufAdoRec!DispenseXml) <> "") And (CheckNullStr(VerkaufAdoRec!ResultXml) <> "")
'        ''    ElseIf (UCase(sAktion) = "ABRUF") Then
'        ''        bOk = (CheckNullStr(VerkaufAdoRec!BundleXml) <> "")
'        ''    End If
'        '
'        '    h = IIf(bOk, Chr(214), "")
'        '    h = h + vbTab + Format(CheckNullDate(VerkaufAdoRec!Anlagedatum), "DD.MM.YY HH:mm")
'        '    h = h + vbTab + sAktion
'        '    flxInfo(0).AddItem h, 1
'
'            VerkaufAdoRec.MoveNext
'        Loop
'        VerkaufAdoRec.Close
'
'        If (.Rows > .FixedRows) Then
'            .FillStyle = flexFillRepeat
'            .row = 1
'            .col = 0
'            .RowSel = .Rows - 1
'            .ColSel = .col
'            .CellFontName = "Symbol"
'
'            .row = .FixedRows
'            .col = 4
'            .RowSel = .Rows - 1
'            .ColSel = .col
'            .Sort = flexSortStringDescending ' 4   'Zahlen absteigend
'            .FillStyle = flexFillSingle
'
'            .Enabled = True
'            .row = .FixedRows
'            .col = 0
'            .ColSel = .Cols - 1
'
'            .ScrollBars = flexScrollBarVertical
'
'            .BackColor = wpara.nlFlexBackColor 'vbWhite
'            .BackColorBkg = wpara.nlFlexBackColor  'vbWhite
'            .BackColorFixed = wpara.nlFlexBackColorFixed   ' RGB(199, 176, 123)
'    '        If (.SelectionMode = flexSelectionFree) Then
'    '            .BackColorSel = RGB(135, 61, 52)
'    '            .ForeColorSel = vbWhite '.ForeColor
'    '        Else
'                .BackColorSel = wpara.nlFlexBackColorSel  ' RGB(232, 217, 172)
'                .ForeColorSel = .ForeColor
'    '        End If
'            .Appearance = 0
'        End If
'    End With
End If

RezeptHolenDB% = True

Call DefErrPop
End Function

Function RezeptHolenMDB%(Optional PrivatModus% = 0)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("RezeptHolenMDB%")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
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
Dim i%, j%, k%, l%, ind%, PrivGef%, iUser%, iVerkMult%, hhRez%, IstDrüber%, ind2%
Dim pTag%, pMonat%, pJahr%, PrivDatum%, iVerkDatum%, RezNrOk%
Dim PrivUser As Byte, PrivLiRe As Byte, AktPrivRezNr As Byte
Dim lAnzVkSaetze&, lIkNr&, iKuNr&
Dim dZuz#, dZuzDos#, dPreis#, dAvp#
Dim ch$, RezEan$, hStr$, LetztArtikel$, AktArtikel$, wg$, SQLStr$, Unique$, h$, sPrivDatum$, sPzn$, sText$, OrgRezNr$
Dim aktavp As Double, OrgAvp As Double
Dim iiDat As Date
Dim AvpRec As Recordset
Dim AVP#, FESTBETR#, VglWert#
Dim PreisDiff_Log%

Static fOrgRezNr!
'Static lPrivEnde&


PreisDiffStr = ""

h$ = RezNr$

Call InitRezept(1)

RezNr$ = Trim(h$)
l% = Len(RezNr)

If (UCase$(Left$(RezNr, 1)) = "X") Then
    If (AvpTeilnahme%) Then
        RezNr$ = Mid$(RezNr$, 2, 11)
'        h$ = Mid$(RezNr$, 7, 4)
'        h$ = Mid$(h$, 3, 2) + Left$(h$, 2)
        SQLStr$ = "SELECT * FROM Rezepte WHERE AvpRezeptNr='" + RezNr$ + "'"
        Set AvpRec = RezSpeicherDB.OpenRecordset(SQLStr$)
        If (AvpRec.RecordCount > 0) Then
            RezNr$ = Trim(AvpRec!RezeptNr)
            l% = Len(RezNr)
        End If
    End If
End If

OrgRezNr = RezNr
RezNrOk = 0

If (UCase$(Left$(RezNr, 1)) = "P") Then
    RezNr = Mid$(RezNr, 2)
    l% = Len(RezNr)
End If

If (UCase$(Left$(OrgRezNr, 1)) = "P") And (l = 0) Then
Else
    RezNrOk = True

    i% = 1
    Do
        If (i% > l%) Then Exit Do

        ch$ = Mid$(RezNr$, i%, 1)
        If (InStr("0123456789", ch$) > 0) Then
            i% = i% + 1
        Else
            RezNr$ = Left$(RezNr$, i% - 1) + Mid$(RezNr$, i% + 1)
            l% = Len(RezNr$)
        End If
    Loop
    If (l% > 13) Then RezNr$ = Left$(RezNr$, 13)

    lRezNr& = 1&
    If (l% < 13) Then
        If (l% <= 5) Then
            lRezNr& = Val(RezNr$)
        Else
            lRezNr& = 0
        End If
        If ((lRezNr& <= 0) Or (lRezNr& > 32000)) Then
            If (UCase$(Left$(OrgRezNr, 1)) = "P") Then
                RezNrOk = 0
            Else
                RezeptHolenMDB% = False
                Call DefErrPop: Exit Function
            End If
        End If
        hStr$ = "9999998" + Format(lRezNr&, "00000")
        Call PruefZiffer(hStr$)
        RezNr$ = hStr$
    End If

'    lRezNr& = CLng(RezNr$)

    bmVerkPtr = Null

    Set VerkaufRec = VerkaufDB.OpenRecordset("Verkauf", dbOpenTable)

    If (UCase$(Left$(OrgRezNr, 1)) = "P") Then
'        ParentForm.txtRezeptNr.text = "P" + RezNr$

        VerkaufRec.index = "PrivRezNr"
        VerkaufRec.Seek ">=", Mid$(OrgRezNr$, 2), Now
        If (VerkaufRec.NoMatch = False) Then
            If (VerkaufRec!PrivRezlaufNr = Val(Mid$(OrgRezNr$, 2))) Then
                bmVerkPtr = VerkaufRec.Bookmark
                PrivatRezept% = True
                GebFrei% = True
            End If
        End If
    Else
        ParentForm.txtRezeptNr.text = RezNr$

    '    VerkaufRec.Seek "=", RezNr$
    '    If (VerkaufRec.NoMatch = False) Then
    '        bmVerkPtr = VerkaufRec.Bookmark
    '    End If
        VerkaufRec.index = "RezeptNr"
        VerkaufRec.Seek ">=", RezNr$, Now
        If (VerkaufRec.NoMatch = False) Then
            If (VerkaufRec!RezeptNr = RezNr$) Then
                bmVerkPtr = VerkaufRec.Bookmark
            End If
        End If
    End If

    If (RezeptHuellen%) Then
        SchonDa% = False
    ElseIf (RezSpeicherOK%) Then
        Set RezepteRec = RezSpeicherDB.OpenRecordset("Rezepte", dbOpenTable)
        With RezepteRec
            .index = "RezeptNr"
            .Seek ">=", RezNr$
            If (.NoMatch = False) Then
                If (RezepteRec!RezeptNr = RezNr$) Then
                    DruckDatum$ = RezepteRec!DruckDatum
                    SchonDa% = True
                End If
            End If
        End With
    End If

    If (SchonDa% = False) And (IsNull(bmVerkPtr)) Then
        If (UCase$(Left$(OrgRezNr, 1)) = "P") Then
            RezNrOk = 0
        Else
            RezeptHolenMDB% = False
            Call DefErrPop: Exit Function
        End If
    End If
End If

If (UCase$(Left$(OrgRezNr, 1)) = "P") And (RezNrOk = 0) Then
    RezNr = OrgRezNr
    RezNr$ = UCase$(RezNr$)

    PrivUser = 0
    PrivLiRe = 0
    PrivDatum% = 0
    sPrivDatum$ = ""
    If (l% = 5) Then
        sPrivDatum$ = Mid$(RezNr$, 2)
        For i% = 2 To 5
            If (InStr("0123456789", Mid$(RezNr$, i%, 1)) = 0) Then
                sPrivDatum$ = ""
                Exit For
            End If
        Next i%
        If (sPrivDatum$ <> "") Then
            pTag% = Val(Left$(sPrivDatum$, 2))
            If (pTag% = 0) Or (pTag% > 31) Then
                sPrivDatum$ = ""
            End If
        End If
        If (sPrivDatum$ <> "") Then
            pMonat% = Val(Mid$(sPrivDatum$, 3, 2))
            If (pMonat% = 0) Or (pMonat% > 12) Then
                sPrivDatum$ = ""
            End If
        End If
        If (sPrivDatum$ <> "") Then
            pJahr% = Year(Now) - 2000
            iiDat = Format(pTag, "00") + "." + Format(pMonat%, "00") + "." + Format(Now, "yyyy")
            If (DateDiff("d", iiDat, Now) < 0) Then
                pJahr% = pJahr% - 1
            End If
            h$ = Format(pTag, "00") + Format(pMonat%, "00") + Format(pJahr%, "00")
            PrivDatum% = iDate(h$)
        End If
    ElseIf (l% >= 2) Then
        ch$ = Mid$(RezNr, 2, 1)
        If (InStr("0123456789", ch$) > 0) Then
            PrivUser = Val(ch$)
            If (l% >= 3) Then
                ind% = InStr("LR", Mid$(RezNr, 3, 1))
                If (ind% > 0) Then PrivLiRe = ind%
            End If
        End If
    End If


    bmOrgVerkPtr = bmVerkPtr
'    lAnzVkSaetze& = VK.DateiLen / VK.RecordLen


    AktPrivRezNr = 0

    If (PrivatModus <> 2) Then
        Set VerkaufRec = VerkaufDB.OpenRecordset("Verkauf", dbOpenTable)
        VerkaufRec.index = "Unique"
    End If

    Select Case PrivatModus%
        Case 0
            VerkaufRec.MoveLast
'            VerkPtr& = lAnzVkSaetze&
        Case -1
            VerkaufRec.Bookmark = bmVerkPtr
            VerkaufRec.MovePrevious
'            VerkPtr& = VerkPtr& - 1
        Case 1
            VerkaufRec.Bookmark = bmPrivEnde
            VerkaufRec.MoveNext
'            VerkPtr& = lPrivEnde& + 1
        Case 2
            'VerkPtr& wurde von aussen gesetzt
    End Select
'

    Do
        If (VerkaufRec.EOF) Or (VerkaufRec.BOF) Then
            Exit Do
        End If

        bmVerkPtr = VerkaufRec.Bookmark
'        Call VK.GetRecord(VerkPtr&)

        RezEan$ = CheckNullStr(VerkaufRec!RezeptNr)

        If (AktPrivRezNr > 0) Then
            If ((VerkaufRec!KundenEnde) Or (VerkaufRec!RezeptEnde)) And (PrivatModus% <> 1) Then
                VerkaufRec.MoveNext
                bmVerkPtr = VerkaufRec.Bookmark
'                VerkPtr& = VerkPtr& + 1
                Exit Do
            End If

            If (VerkaufRec!PrivRezNr <> AktPrivRezNr) Or (Val(RezEan) > 0) Then
                VerkaufRec.MoveNext
                bmVerkPtr = VerkaufRec.Bookmark
'                VerkPtr& = VerkPtr& + 1
                Exit Do
            End If
        End If

        If ((VerkaufRec!RezeptArt = REZEPTART_PRIVAT) Or VerkaufRec!RezeptArt = REZEPTART_GRÜN) And (AktPrivRezNr = 0) And (Val(RezEan) = 0) Then
            PrivGef% = True
            If (VerkaufRec!PrivRezNr = 0) Then
                PrivGef% = False
            End If
            If (PrivDatum% > 0) Then
                h$ = Format(VerkaufRec!Datum, "DDMMYY")
                iVerkDatum% = iDate(h$)
                If (iVerkDatum > PrivDatum%) Then
                    PrivGef% = False
                End If
            ElseIf (PrivUser) Then
                iUser% = VerkaufRec!Computer
                If (PrivUser <> iUser%) Then
                    PrivGef% = False
                Else
                    If (PrivLiRe) Then
                        VerkLiRe = VerkaufRec!Seite
                        If (VerkLiRe <> (PrivLiRe - 1)) Then PrivGef% = False
                    End If
                End If
            End If

            If (PrivGef%) Then
                AktPrivRezNr = VerkaufRec!PrivRezNr
                If (PrivatModus% = 1) Then Exit Do
            End If
        End If

        If (PrivatModus% = 1) Then
            VerkaufRec.MoveNext
'            VerkPtr& = VerkPtr& + 1
'            If (VerkPtr& >= lAnzVkSaetze&) Then Exit Do
'        ElseIf (VerkPtr& <= 2&) Then
'            Exit Do
        Else
            VerkaufRec.MovePrevious
'            VerkPtr& = VerkPtr& - 1
        End If
    Loop

    If (AktPrivRezNr = 0) Then
        Call iMsgBox("keine weiteren passenden Privat-Rezepte gespeichert !")
        bmVerkPtr = bmOrgVerkPtr
        If (IsNull(bmVerkPtr)) Then
            RezeptHolenMDB% = False
            Call DefErrPop: Exit Function
        End If
    End If

    bmPrivEnde = bmVerkPtr
    lRezNr& = -1&
    PrivatRezept% = True
    GebFrei% = True
'Else
'
'    Set VerkaufRec = VerkaufDB.OpenRecordset("Verkauf", dbOpenTable)
'    VerkaufRec.index = "RezeptNr"
'
'    i% = 1
'    Do
'        If (i% > l%) Then Exit Do
'
'        ch$ = Mid$(RezNr$, i%, 1)
'        If (InStr("0123456789", ch$) > 0) Then
'            i% = i% + 1
'        Else
'            RezNr$ = Left$(RezNr$, i% - 1) + Mid$(RezNr$, i% + 1)
'            l% = Len(RezNr$)
'        End If
'    Loop
'    If (l% > 13) Then RezNr$ = Left$(RezNr$, 13)
'
'    lRezNr& = 1&
'    If (l% < 13) Then
'        If (l% <= 5) Then
'            lRezNr& = Val(RezNr$)
'        Else
'            lRezNr& = 0
'        End If
'        If (lRezNr& <= 0) Or (lRezNr& > 32000) Then
'            RezeptHolenMDB% = False
'            Call DefErrPop: Exit Function
'        End If
'        hStr$ = "9999998" + Format(lRezNr&, "00000")
'        Call PruefZiffer(hStr$)
'        RezNr$ = hStr$
'    End If
'
''    lRezNr& = CLng(RezNr$)
'
'    ParentForm.txtRezeptNr.text = RezNr$
'
'    bmVerkPtr = Null
'
''    VerkaufRec.Seek "=", RezNr$
''    If (VerkaufRec.NoMatch = False) Then
''        bmVerkPtr = VerkaufRec.Bookmark
''    End If
'    VerkaufRec.Seek ">=", RezNr$, Now
'    If (VerkaufRec.NoMatch = False) Then
'        If (VerkaufRec!RezeptNr = RezNr$) Then
'            bmVerkPtr = VerkaufRec.Bookmark
'        End If
'    End If
'
'    If (RezeptHuellen%) Then
'        SchonDa% = False
'    ElseIf (RezSpeicherOK%) Then
'        Set RezepteRec = RezSpeicherDB.OpenRecordset("Rezepte", dbOpenTable)
'        With RezepteRec
'            .index = "RezeptNr"
'            .Seek ">=", RezNr$
'            If (.NoMatch = False) Then
'                If (RezepteRec!RezeptNr = RezNr$) Then
'                    DruckDatum$ = RezepteRec!DruckDatum
'                    SchonDa% = True
'                End If
'            End If
'        End With
'    End If
'
'    If (SchonDa% = False) And (IsNull(bmVerkPtr)) Then
'        RezeptHolenMDB% = False
'        Call DefErrPop: Exit Function
'    End If

End If


If (SchonDa%) Then
    Call HoleAusRezeptSpeicher
Else
    Call InitRezept(2)

    LetztArtikel$ = ""
    k% = -1

    VerkaufRec.Bookmark = bmVerkPtr
'    Call VK.GetRecord(VerkPtr&)
    If (PrivatRezept%) Then
        AktPrivRezNr = VerkaufRec!PrivRezNr
    End If
    If (VerkaufRec!Bediener > 0) Then
        PersonalNr% = VerkaufRec!Bediener
    End If

    lIkNr = CheckNullLong(VerkaufRec!IkNr)
    If (AutIdemSonderregelOk%) And (lIkNr > 0) Then
        SQLStr$ = "SELECT * FROM kKassen WHERE IK = " + Str$(lIkNr)
'        Set KkassenRec = kKassenDB.OpenRecordset(SQLStr$)
'        If (KkassenRec.RecordCount > 0) Then
        FabsErrf = kKassenDB.OpenRecordset(KkassenRec, SQLStr)
        If Not (KkassenRec.EOF) Then
            AutIdemIk& = CheckNullLong(KkassenRec!IK)
            AutIdemKbvNr& = CheckNullLong(KkassenRec!KbvNr)
            ActAVWGik& = AutIdemIk&
        End If
    End If

    IvfRezept% = (VerkaufRec!RezeptArt = 7)

    Do
        sPzn = PznString(VerkaufRec!pzn)
        sText = VerkaufRec!text
        If (VerkaufRec!SonderPreis) Then
            dPreis# = VerkaufRec!vk
            If ((VerkaufRec!FB > 0) And (VerkaufRec!FB < dPreis)) Then
                dPreis# = VerkaufRec!FB
            End If
        Else
            dPreis = VerkaufRec!Preis
        End If
        If (dPreis# = 0) Then
            If (CheckNullLong(VerkaufRec!AbholNr) > 0) Then
                dPreis# = VerkaufRec!vk
                If (VerkaufRec!FB > 0) And (VerkaufRec!FB < dPreis#) Then
                    dPreis# = VerkaufRec!FB
                End If
            End If
        End If

        If (Mid$(sText, 15, 4) = "PZN:") Then
            'Stückelung
        Else
            AktArtikel$ = sPzn + sText + Format(dPreis, "00000.00")
            If (Left$(sText, 5) = "u n t") Then
                If (VerkaufRec!AbholNr > 0) Then
                    If (CheckNullStr(VerkaufRec!AbholerArt) = "R") Then
                        RezArtikel(k%).zusatz(0) = "Abhol-Nr" + Right(Space$(4) + Format(VerkaufRec!AbholNr, "0"), 4)
'                    ElseIf (PrivatRezept) And (CheckNullStr(VerkaufRec!AbholerArt) = "B") Then
'                        RezArtikel(k%).Preis = (VerkaufRec!Aconto + VerkaufRec!Restzahlung) / RezArtikel(k%).Faktor
                    End If
                End If

'            ElseIf (sPzn = "9999999") And (Left$(UCase(sText), 10) = "IK-NUMMER:") Then
'                h$ = Trim(VK.KKNr)
'                If (AutIdemSonderregelOk%) And (Val(h$) > 0) Then
'                    SQLStr$ = "SELECT * FROM kKassen WHERE IK = " + Str$(Val(h$))
'                    Set KkassenRec = kKassenDB.OpenRecordset(SQLStr$)
'                    If (KkassenRec.RecordCount > 0) Then
'                        AutIdemIk& = CheckNullLong(KkassenRec!IK)
'                        AutIdemKbvNr& = CheckNullLong(KkassenRec!KbvNr)
'                        ActAVWGik& = AutIdemIk&
'                    End If
'                End If
'            ElseIf (Left$(sText, 5) = "Abhol") Then
'                If ((PrivatRezept%) And ((VerkaufRec!PrivRezNr <> AktPrivRezNr) Or (Val(VerkaufRec!RezeptNr) > 0))) Then
'                    k% = k% + 1
'                Else
'                    RezArtikel(k% + 1).zusatz(0) = sText
'                End If
'                PrivatBesorger% = 0
'            ElseIf (PrivatRezept%) And (Left$(sText, 1) = "(") Then
'                RezArtikel(k%).zusatz(1) = sText
'                PrivatBesorger% = 0
            ElseIf (AktArtikel$ = LetztArtikel$) And (Val(sPzn) <> 9999999) And _
                ((PrivatRezept% = False) Or ((VerkaufRec!PrivRezNr = AktPrivRezNr) And (Val(VerkaufRec!RezeptNr) = 0))) Then
                RezArtikel(k%).Faktor = RezArtikel(k%).Faktor + VerkaufRec!Anzahl

' weg mit 2.0.17
'                If (VerkaufRec!HimiVerbrauch) Then
'                    RezArtikel(k%).Zuz = RezArtikel(k%).Zuz + VerkaufRec!Zuzahlung
'                End If

'            ElseIf (Left$(sText, 1) = " ") And (Mid$(sText, 20, 1) = "x") Then
'                iVerkMult% = Mid$(sText, 16, 3) - 1
'                RezArtikel(k%).Faktor = RezArtikel(k%).Faktor + iVerkMult%
'                If (PrivatBesorger% = 2) Then
'                    RezArtikel(k%).Preis = RezArtikel(k%).Preis / RezArtikel(k%).Faktor
'                Else
'                    RezArtikel(k%).Preis = Val(Mid$(sText, 21, 9))
'                End If
'                PrivatBesorger% = 0
            Else
                k% = k% + 1
                RezArtikel(k%).pzn = sPzn
                RezArtikel(k%).text = Left$(sText, 28) + " " + CheckNullStr(VerkaufRec!menge) + CheckNullStr(VerkaufRec!einheit)
                RezArtikel(k%).Preis = dPreis
                RezArtikel(k%).Faktor = VerkaufRec!Anzahl
                RezArtikel(k%).flag = 0
                RezArtikel(k%).appli = Abs(InStr(UCase(VerkaufRec!Rezeptzeichen), "A") > 0)  'Abs(VerkaufRec!Hilfsmittel)

                RezArtikel(k%).Zuz = 0
                RezArtikel(k%).HerstRabattPrivat130Brutto = CheckNullDouble(VerkaufRec!HerstRabattPrivat130Brutto)
                RezArtikel(k).VdbPauschale = 0
'                If (IvfRezept) Then
'                    RezArtikel(k%).Zuz = VerkaufRec!Zuzahlung
'                    RezArtikel(k%).Preis = VerkaufRec!VK
'                End If

                RezArtikel(k%).Imp123 = 0
                RezArtikel(k%).Fam = 0
                RezArtikel(k%).AutIdem = 0
                RezArtikel(k%).N0 = 0
                RezArtikel(k%).MagIndex = 0

                RezArtikel(k%).NichtInTaxe = 0
                RezArtikel(k%).ImpErspart = 0

                If (Left$(sText, 5) = "KONTR") Then
                    RezArtikel(k%).flag = RezArtikel(k%).flag Or (REZ_K_STOP Or REZ_BLINK)
                    Mid$(RezArtikel(k%).text, 21, 6) = "(PZN?)"
                    BeepFlag% = True
                End If
                If (Left$(sText, 12) = "PREISEINGABE") Then
                    Mid$(RezArtikel(k%).text, 21, 6) = "(PZN?)"
                End If

                If (Val(RezArtikel(k%).pzn) = 9999999) Then
                    If (RezArtikel(k%).flag And REZ_K_STOP) Then
                    ElseIf (RezArtikel(k%).Preis > 0) Then
                        If (InStr(UCase(VerkaufRec!Kassenzeichen), "A") > 0) Then
                            RezArtikel(k%).appli = 1
                        End If
                    End If
                End If

                RezArtikel(k%).IstWg4 = 0
                RezArtikel(k%).PlusMehrKosten = 0
                RezArtikel(k%).ZuzahlungsErlass = 0
                RezArtikel(k%).Warenzeichen = 0
                RezArtikel(k%).MietDauer = 0

                If (VerkaufRec!AbholNr > 0) Then
                    If (CheckNullStr(VerkaufRec!AbholerArt) = "R") Then
                        RezArtikel(k%).zusatz(0) = "Abhol-Nr" + Right(Space$(4) + Format(VerkaufRec!AbholNr, "0"), 4)
                    ElseIf (PrivatRezept) And (CheckNullStr(VerkaufRec!AbholerArt) = "B") Then
                        If (CheckNullLong(VerkaufRec!ErsetztPzn) = 0) Then
                            RezArtikel(k%).Preis = (VerkaufRec!Aconto + VerkaufRec!Restzahlung) / RezArtikel(k%).Faktor
                        End If
                    End If
                End If

                If (VerkaufRec!Taxmuster) Then
                    RezArtikel(k%).Warenzeichen = 99
                    SollTmName(k%) = UCase(sText)
                    SollTmMenge(k%) = xVal(CheckNullStr(VerkaufRec!menge))
                End If
'                wg$ = VK.wg
'                If (Left$(wg$, 1) = "4") Then
'                    RezArtikel(k%).flag = RezArtikel(k%).flag Or (REZ_WG_4 Or REZ_BLINK)
'                    RezArtikel(k%).IstWg4 = 1
'                ElseIf (VmFlag%) Then
'                    FabsErrf% = VmPzn.IndexSearch(0, VK.pzn, FabsRecno&)
'                    If (FabsErrf% = 0) Then
'                        RezArtikel(k%).flag = RezArtikel(k%).flag Or (REZ_WG_4 Or REZ_BLINK)
'                    End If
'                End If

                SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + RezArtikel(k%).pzn
                'Set TaxeRec = TaxeDB.OpenRecordset(SQLStr$)
                On Error Resume Next
                TaxeRec.Close
                Err.Clear
                On Error GoTo DefErr
                TaxeRec.Open SQLStr, taxeAdoDB.ActiveConn
                If (TaxeRec.EOF = False) Then
                    RezArtikel(k%).IstWg4 = CheckZuzZeilenwert
                    If (TaxeRec!VerbandKz) Then
                        RezArtikel(k%).flag = RezArtikel(k%).flag Or (REZ_WG_4 Or REZ_BLINK)
                    ElseIf (VmFlag%) Then
                        If (AplusVOk) Then
                            SQLStr$ = "SELECT * FROM VdbArtikel WHERE PZN = " + sPzn
'                            Set VdbPznRec = AplusVDB.OpenRecordset(SQLStr$)
'                            FabsErrf = Abs(VdbPznRec.EOF)
                            FabsErrf = AplusVDB.OpenRecordset(VdbPznRec, SQLStr)
                            VdbPznRec.Close
                        Else
                            FabsErrf% = VmPzn.IndexSearch(0, sPzn, FabsRecno&)
                        End If
                        If (FabsErrf% = 0) Then
                            RezArtikel(k%).flag = RezArtikel(k%).flag Or (REZ_WG_4 Or REZ_BLINK)
                        End If
                    End If

                    If (TaxeRec!HimiVerbrauch) Then
                        If (CheckNullDouble(VerkaufRec!Zuzages) > 0) Then
                            RezArtikel(k%).Zuz = VerkaufRec!Zuzages / RezArtikel(k%).Faktor
                        Else
                            RezArtikel(k%).Zuz = VerkaufRec!Zuzahlung
                        End If
                    End If

                    If (VerkaufRec!MehrkostenverzichtKZ) Then
                        RezArtikel(k%).PlusMehrKosten = 1
                    End If
                    If (VerkaufRec!ZuzahlungsErlassKZ) Then
                        RezArtikel(k%).ZuzahlungsErlass = 1
                        RezArtikel(k%).Zuz = VerkaufRec!Zuzahlung
                    End If
                End If

'                PrivatBesorger% = 0
            End If
        End If

        LetztArtikel$ = AktArtikel$

'        If ((VerkaufRec!KundenEnde) Or (VerkaufRec!RezeptEnde)) Or _
            ((PrivatRezept%) And ((VerkaufRec!PrivRezNr <> AktPrivRezNr) Or (Val(VerkaufRec!RezeptNr) > 0))) Then
        If (VerkaufRec!RezeptEnde) Then

            VerkDatum$ = Format(VerkaufRec!Datum, "DDMMYY")
            OrgVerkDatum$ = VerkDatum$

'          DruckDatum = 0L;

            VerkZeit% = Val(Format(VerkaufRec!Datum, "HHMM"))

            ActVerordnung% = 1
            If (VerkaufRec!RezeptArt = REZEPTART_SPRECHSTUNDENBEDARF) Then ActVerordnung% = 2

            iUser% = VerkaufRec!Computer
            VerkUser$ = Format(iUser%, " 0")
            VerkLiRe = VerkaufRec!Seite

            iKuNr = VerkaufRec!KundenNr
            If (iKuNr > 0) Then
                Call HoleKundenInfo(iKuNr)
            End If

            If (VerkaufRec!RezeptArt = REZEPTART_ZUZAFREI) Or (VerkaufRec!RezeptArt = REZEPTART_BTM_ZUZAFREI) Or (VerkaufRec!RezeptArt = REZEPTART_SPRECHSTUNDENBEDARF) Then
                RezGebSumme# = 0#
                GebFrei% = True
            End If

            If (VerkaufRec!AVKassentyp > 0) Then
                ActKasse% = VerkaufRec!AVKassentyp
            End If

            ActKkasse$ = ""
            If (VerkaufRec!AVKassenNr > 0) Then
                ActKkasse$ = Right$(String(9, "0") + Format(VerkaufRec!AVKassenNr), 9) 'Trim(VK.KKNr)
            End If

            ActVebNr = CheckNullLong(VerkaufRec!VebNr)
            ActPauschaleNr = CheckNullLong(VerkaufRec!PauschaleNr)

            If (PrivatRezept) Then
                ActKasse% = 0
                ActKkasse$ = ""
            End If

'            If (AutIdemSonderregelOk%) And (Val(ActKkasse$) > 0) Then
'                SQLStr$ = "SELECT * FROM kKassen WHERE IK = " + Str$(Val(ActKkasse$))
'                Set KkassenRec = kKassenDB.OpenRecordset(SQLStr$)
'                If (KkassenRec.RecordCount > 0) Then
'                    AutIdemIk& = CheckNullLong(KkassenRec!IK)
'                    AutIdemKbvNr& = CheckNullLong(KkassenRec!KbvNr)
'                End If
'            End If


'            If ((PrivatRezept% = False) Or ((VerkaufRec!PrivRezNr = AktPrivRezNr) And (Val(VerkaufRec!RezeptNr) = 0))) Then
'                k% = k% + 1
'            Else
'                lPrivEnde& = lPrivEnde& - 1
'            End If

            bmPrivEnde = VerkaufRec.Bookmark
            k% = k% + 1

            Exit Do
        End If

        VerkaufRec.MoveNext
    Loop

    AnzRezeptArtikel% = k%

    SonderBelegRezept% = False
    If (AnzRezeptArtikel% = 1) Then
        For i% = 1 To AnzSonderBelege%
            If (SonderBelege(i% - 1).pzn = RezArtikel(0).pzn) Then
                SonderBelegRezept% = i%
                PrivatRezept% = 0
                BtmRezept% = 0
                IvfRezept% = 0
                RezKundenInfo$(1) = ""
                RezKundenInfo$(2) = "SONDERBELEG"
                Exit For
            End If
        Next i%
    End If


'    If Not (VerkaufRec!KundenEnde) And Not (PrivatRezept%) And (SonderBelegRezept% = 0) Then
'        Set VerkaufRec = VerkaufDB.OpenRecordset("Verkauf", dbOpenTable)
'        VerkaufRec.index = "Unique"
'
'        VerkaufRec.Bookmark = bmVerkPtr ' bmPrivEnde
'        For i% = 0 To UBound(BarVerkauf$)
'            VerkaufRec.MoveNext
'            If (VerkaufRec.EOF) Then
'                Exit For
'            End If
'
'            If (VerkaufRec!RezeptArt = 0) Then
'                hStr$ = VerkaufRec!text + " " + Format(VerkaufRec!Preis, "0.00")
'                BarVerkauf$(i%) = hStr$
'            End If
'
'            If (VerkaufRec!KundenEnde) Then
'                Exit For
'            End If
'        Next i%
'    End If

    If Not (PrivatRezept%) And (SonderBelegRezept% = 0) Then
        Set VerkaufRec = VerkaufDB.OpenRecordset("Verkauf", dbOpenTable)
        VerkaufRec.index = "Unique"

        VerkaufRec.Bookmark = bmPrivEnde
        Do
            If (VerkaufRec.BOF) Then
                Exit Do
            End If
            If (VerkaufRec!KundenEnde) Then
                Exit Do
            End If
            VerkaufRec.MovePrevious
        Loop

        i% = 0
        Do
            VerkaufRec.MoveNext
            If (VerkaufRec.EOF) Then
                Exit Do
            End If

            If (VerkaufRec!RezeptArt = 0) And (VerkaufRec!Anzahl <> 0) Then
                hStr$ = VerkaufRec!text + " " + Format(VerkaufRec!Preis, "0.00")
                BarVerkauf$(i%) = hStr$
                i% = i% + 1
                If (i% > UBound(BarVerkauf$)) Then
                    Exit Do
                End If
            End If

            If (VerkaufRec!KundenEnde) Then
                Exit Do
            End If
        Loop
    End If

    If (PrivatRezept%) Or (SonderBelegRezept%) Then
    Else
        AMverfügbarkeit = Left$(String(AnzRezeptArtikel%, "1") + "000", 3)
    End If

    For i% = 0 To (AnzRezeptArtikel% - 1)
        dZuz# = 0#
        dAvp# = 0

        If (PrivatRezept%) Or (SonderBelegRezept%) Then
            RezArtikel(i%).flag = RezArtikel(i%).flag And (REZ_BLINK Xor &HFF)
            If (PrivatRezept) And (Val(RezArtikel(i%).pzn) <> 9999999) Then
                SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + RezArtikel(i%).pzn
                'Set TaxeRec = TaxeDB.OpenRecordset(SQLStr$)
                On Error Resume Next
                TaxeRec.Close
                Err.Clear
                On Error GoTo DefErr
                TaxeRec.Open SQLStr, taxeAdoDB.ActiveConn
                If (TaxeRec.EOF) Then
                    RezArtikel(i%).NichtInTaxe = 1

                    If (ArtikelDbOk) Then
                        SQLStr$ = "SELECT * FROM ARTIKEL WHERE PZN = " + RezArtikel(i%).pzn
    '                    SQLStr = SQLStr + " AND LagerKz<>0"
                        FabsErrf = Artikel.OpenRecordset(ArtikelAdoRec, SQLStr)
                    Else
                        FabsErrf% = ast.IndexSearch(0, RezArtikel(i%).pzn, FabsRecno&)
                        If (FabsErrf% = 0) Then
                            ast.GetRecord (FabsRecno& + 1)
                        End If
                    End If
                    If (FabsErrf% = 0) Then
                        If (Val(ast.wg) = 3) Then
                            RezArtikel(i).pzn = "09999011"
                            RezArtikel(i).text = "Rezeptur"
                        End If
                    End If
                ElseIf (TaxeRec!BtmKz) Then
                    BtmRezept% = True
                End If
            End If
        Else

            If (Val(RezArtikel(i%).pzn) = 9999999) Then
                If (RezArtikel(i%).flag And REZ_K_STOP) Then
                    dZuz# = 5   'ZuzFeld!(3)
                    RezArtikel(i%).flag = RezArtikel(i%).flag Or REZ_BLINK
                ElseIf (RezArtikel(i%).Preis > 0) Then
                    If (RezArtikel(i%).appli = 0) Then
                        dZuz# = CalcZuzaNeu(RezArtikel(i%).Preis)
                    End If

                    RezArtikel(i%).flag = RezArtikel(i%).flag Or REZ_BLINK

                    ind2% = InStr(RezArtikel(i%).text, "(PZN?)")
                    If (ind2% <= 0) Then
                        ch$ = Trim(RezArtikel(i%).text) + "   " + "(PZN?)"
                        RezArtikel(i%).text = Left$(ch$ + Space$(36), 36)
                    End If
                End If
            Else
                SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + RezArtikel(i%).pzn
                'Set TaxeRec = TaxeDB.OpenRecordset(SQLStr$)
                On Error Resume Next
                TaxeRec.Close
                Err.Clear
                On Error GoTo DefErr
                TaxeRec.Open SQLStr, taxeAdoDB.ActiveConn
                If (TaxeRec.EOF) Then
                    RezArtikel(i%).NichtInTaxe = 1
                    If (CheckSonderPzn%(RezArtikel(i%).pzn) >= 0) Then
'                        dZuz# = ZuzFeld!(3)
                        dZuz# = CalcZuzaNeu(RezArtikel(i%).Preis)
                    Else
                        If (ArtikelDbOk) Then
                            SQLStr$ = "SELECT * FROM ARTIKEL WHERE PZN = " + RezArtikel(i%).pzn
        '                    SQLStr = SQLStr + " AND LagerKz<>0"
                            FabsErrf = Artikel.OpenRecordset(ArtikelAdoRec, SQLStr)
                        Else
                            FabsErrf% = ast.IndexSearch(0, RezArtikel(i%).pzn, FabsRecno&)
                            If (FabsErrf% = 0) Then
                                ast.GetRecord (FabsRecno& + 1)
                            End If
                        End If
                        If (FabsErrf% = 0) Then
                            If (Val(ast.wg) = 3) Then
                                RezArtikel(i).pzn = "09999011"
                                RezArtikel(i).text = "Rezeptur"
                                dZuz# = CalcZuzaNeu(RezArtikel(i%).Preis)
                            End If
                        End If
                    End If
                Else
'                    If (TaxeRec!Warengruppe = 3) Then
'                        RezArtikel(i%).IstWg4 = 1
'                    End If

                    dAvp# = TaxeRec!vk / 100#

                    dPreis# = RezArtikel(i%).Preis
                    If (RezArtikel(i%).PlusMehrKosten) Then
                        If (TaxeRec!FESTBETRAG > 0) And (TaxeRec!FESTBETRAG < TaxeRec!vk) Then
                            dPreis# = (TaxeRec!FESTBETRAG / 100#)
                        Else
                            RezArtikel(i%).PlusMehrKosten = 0
                        End If
                    End If

                    If (RezArtikel(i%).ZuzahlungsErlass) Then
                        dZuzDos# = CalcTaxeZuzahlungDOS(dPreis#, RezArtikel(i%).Zuz + 200)
                    Else
                        dZuzDos# = CalcTaxeZuzahlungDOS(dPreis#)
                    End If

                    If (TaxeRec!Negliste) Then dZuzDos# = 1#
                    If (RezArtikel(i%).appli) Then dZuzDos# = 0#

                    If (RezArtikel(i%).PlusMehrKosten) Then
                        RezArtikel(i%).Preis = TaxeRec!vk / 100#
                    End If

    '                If (Val(Left$(TaxeRec!Zuzahlung, 1)) = 9) Then dZuzDOS# = 0#

                    If (dZuzDos# >= 1000#) Then
                        dZuz# = dZuzDos# - 1000#
                    ElseIf Not (GebFrei%) And (dZuzDos# > 0#) Then
                        RezArtikel(i%).Preis = 0#
                        If (dZuzDos# = 3#) Then
    '                        Call iMsgBox("Artikel unter Zuzahlung, Geldrückgabe ?")
                        Else
                            RezArtikel(i%).flag = RezArtikel(i%).flag And (REZ_BLINK Xor &HFF)
                        End If
                    End If
                    If (TaxeRec!HimiVerbrauch) Then
                        dZuz# = RezArtikel(i%).Zuz '/ RezArtikel(i%).Faktor
                    End If

                    If (RezArtikel(i%).ZuzahlungsErlass) Then
                        dZuz# = RezArtikel(i%).Zuz
                    End If

                    If (TaxeRec!BtmKz) Then
                        BtmRezept% = True
                    End If

                    hStr$ = Format$(TaxeRec!Abgabebest, "0")
                    If (InStr("01245", hStr$) <> 0) Then
                        RezArtikel(i%).Fam = 1
                    End If

                    Call CheckAutIdem(i%)
                    SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + RezArtikel(i%).pzn
                    'Set TaxeRec = TaxeDB.OpenRecordset(SQLStr$)
                    On Error Resume Next
                    TaxeRec.Close
                    Err.Clear
                    On Error GoTo DefErr
                    TaxeRec.Open SQLStr, taxeAdoDB.ActiveConn

                    If (Left$(TaxeRec!Zuzahlung, 1) = "9") Then
                        RezArtikel(i%).IstWg4 = 0
                    End If

                    If (ImpFlag% = 0) Then
                        RezArtikel(i%).Imp123 = 0
                    ElseIf (TaxeRec!OriginalPZN = TaxeRec!pzn) Then
                        RezArtikel(i%).Imp123 = 6
                        OrgAvp = TaxeRec!vk / 100#
                        SQLStr$ = "SELECT * FROM TAXE WHERE ORIGINALPZN = " + PznString(TaxeRec!OriginalPZN)
                        'Set TaxeRec = TaxeDB.OpenRecordset(SQLStr$)
                        On Error Resume Next
                        TaxeRec.Close
                        Err.Clear
                        On Error GoTo DefErr
                        TaxeRec.Open SQLStr, taxeAdoDB.ActiveConn
                        TaxeRec.MoveFirst
                        Do While Not TaxeRec.EOF
                          If (TaxeRec!OriginalPZN <> TaxeRec!pzn) Then
                            aktavp = TaxeRec!vk / 100#
                            If (aktavp <= OrgAvp - 15) Or (aktavp <= OrgAvp * 0.85) Then
                              RezArtikel(i%).Imp123 = 5
                              Exit Do
                            End If
                          End If
                          TaxeRec.MoveNext
                        Loop
                        SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + RezArtikel(i%).pzn
                        'Set TaxeRec = TaxeDB.OpenRecordset(SQLStr$)
                        On Error Resume Next
                        TaxeRec.Close
                        Err.Clear
                        On Error GoTo DefErr
                        TaxeRec.Open SQLStr, taxeAdoDB.ActiveConn

                    ElseIf (TaxeRec!OriginalPZN > 0) Then
                        aktavp = TaxeRec!vk / 100
                        SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + PznString(TaxeRec!OriginalPZN)
                        'Set TaxeRec = TaxeDB.OpenRecordset(SQLStr$)
                        On Error Resume Next
                        TaxeRec.Close
                        Err.Clear
                        On Error GoTo DefErr
                        TaxeRec.Open SQLStr, taxeAdoDB.ActiveConn
                        If Not (TaxeRec.EOF) Then
                          OrgAvp = TaxeRec!vk / 100
                        End If
                        If (aktavp <= OrgAvp - 15) Or (aktavp <= OrgAvp * 0.85) Then
                          RezArtikel(i%).Imp123 = 2
                          RezArtikel(i%).ImpErspart = (OrgAvp - aktavp)
                        Else
                          RezArtikel(i%).Imp123 = 1
                        End If
                        SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + RezArtikel(i%).pzn
                        'Set TaxeRec = TaxeDB.OpenRecordset(SQLStr$)
                        On Error Resume Next
                        TaxeRec.Close
                        Err.Clear
                        On Error GoTo DefErr
                        TaxeRec.Open SQLStr, taxeAdoDB.ActiveConn

                    End If
                    
                    If (PreisDiffAktiv) Then
                        AVP# = TaxeRec!vk / 100
                        FESTBETR# = TaxeRec!FESTBETRAG / 100
                        If ((TaxeRec!FestKz) And (FESTBETR# < AVP#)) Then
                            VglWert# = FESTBETR#
                        Else
                            VglWert# = AVP#
                        End If
                        If (VglWert > 0) Then
                            If (RezArtikel(i%).Preis <> VglWert) Then
                                h = RezArtikel(i).pzn + "  " + RezArtikel(i).text + "  " + Format(RezArtikel(i).Preis, "0.00") + "  " + Format(VglWert, "0.00")
                                PreisDiff_Log% = FreeFile
                                Open "PreisDif.log" For Append As #PreisDiff_Log
                                Print #PreisDiff_Log%, Format(Now, "DD.MM.YYYY HH:MM:SS  ") + RezNr + "  " + h$
                                Close #PreisDiff_Log%
                                
                                PreisDiffStr = PreisDiffStr + h + vbCrLf
                            End If
                        End If
                    End If
                End If
            End If
        End If

        If (IvfRezept) Then
            RezArtikel(i%).Zuz = 0
            RezArtikel(i%).Preis = FNX(dAvp / 2)
        Else
            RezArtikel(i%).Zuz = dZuz#
        End If
    Next i%

    If (IvfRezept) Then
        If (Val(RezArtikel(0).pzn) <> 9999643) Then
            For i% = UBound(RezArtikel) - 1 To 0 Step -1
                RezArtikel(i% + 1) = RezArtikel(i%)
            Next i%
            Call InitRezeptArtikel(0)
'            Call InitKontrollStop(row%)

            RezArtikel(0).pzn = "09999643"
            RezArtikel(0).text = "IVF-Rezept"
            RezArtikel(0).Faktor = 1
            RezArtikel(0).Preis = 0
            RezArtikel(0).Zuz = 0

            AnzRezeptArtikel% = AnzRezeptArtikel% + 1
        End If
    End If

    If (BtmRezept%) Then
        If (PrivatRezept) Then
            PrivatRezept = 0
        End If
        For i% = 0 To UBound(RezAddGesamtBrutto$)
            RezAddGesamtBrutto$(i%) = ""
        Next i%
        Call PruefeBtmKosten
        sBrutto$ = RezAddGesamtBrutto$(0)
    End If

    Call PruefeRezkontrDat
    
    If (ActVebNr = 0) And (ActPauschaleNr = 0) Then
        If (AutIdemIk& > 0) Then
            SQLStr = "SELECT count(*) AS iAnz FROM VdbVVI LEFT JOIN VdbVVL ON VdbVVI.VebNr=VdbVVL.VebNr"
            SQLStr = SQLStr + " WHERE VdbVVI.Ik=" + CStr(AutIdemIk)
            SQLStr = SQLStr + " AND VdbVVL.LandNr=" + CStr(ActBundesland)
            FabsErrf = AplusVDB.OpenRecordset(VdbBedRec, SQLStr)
            If Not (VdbBedRec.EOF) Then
                If (CheckNullLong(VdbBedRec!iAnz) = 1) Then
                    SQLStr = "SELECT VdbVVI.VebNr FROM VdbVVI LEFT JOIN VdbVVL ON VdbVVI.VebNr=VdbVVL.VebNr"
                    SQLStr = SQLStr + " WHERE VdbVVI.Ik=" + CStr(AutIdemIk)
                    SQLStr = SQLStr + " AND VdbVVL.LandNr=" + CStr(ActBundesland)
                    FabsErrf = AplusVDB.OpenRecordset(VdbBedRec, SQLStr)
                    If Not (VdbBedRec.EOF) Then
                        ActVebNr = CheckNullLong(VdbBedRec!VebNr)
                    End If
                End If
            End If
        End If
        If (ActVebNr = 0) And (ActKasse% > 0) Then
            SQLStr = "SELECT Top 1* FROM VdbBedingungen"
            SQLStr = SQLStr + " WHERE KostentrNr=" + CStr(ActKasse)
            SQLStr = SQLStr + " AND LandNr=" + CStr(ActBundesland)
            FabsErrf = AplusVDB.OpenRecordset(VdbBedRec, SQLStr)
            If Not (VdbBedRec.EOF) Then
                ActVebNr = CheckNullLong(VdbBedRec!VebNr)
            End If
        End If
    End If
    If (ActKasse% < 0) And (AutIdemIk& > 0) Then
        SQLStr = "SELECT Top 1* FROM VdbVIK"
        SQLStr = SQLStr + " WHERE IK=" + CStr(AutIdemIk)
        SQLStr = SQLStr + " ORDER BY KostentrNr"
        FabsErrf = AplusVDB.OpenRecordset(VdbBedRec, SQLStr)
        If Not (VdbBedRec.EOF) Then
            ActKasse = CheckNullLong(VdbBedRec!KostentrNr)
        End If
    End If

    If (ActKasse% > 0) Or (ActVebNr > 0) Or (ActPauschaleNr > 0) Then
        For i% = 0 To (AnzRezeptArtikel% - 1)
            If (RezArtikel(i%).flag And REZ_WG_4) Then
                RezeptHolenZuz# = RezArtikel(i%).Zuz
                AktTabIndex% = 3 + i% * 4
                Call Verbandmittel
            
                If (PreisDiffAktiv) Then
                    SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + RezArtikel(i%).pzn
                    'Set TaxeRec = TaxeDB.OpenRecordset(SQLStr$)
                    On Error Resume Next
                    TaxeRec.Close
                    Err.Clear
                    On Error GoTo DefErr
                    TaxeRec.Open SQLStr, taxeAdoDB.ActiveConn
                    If Not (TaxeRec.EOF) Then
                        AVP# = TaxeRec!vk / 100
                        FESTBETR# = TaxeRec!FESTBETRAG / 100
                        If ((TaxeRec!FestKz) And (FESTBETR# < AVP#)) Then
                            VglWert# = FESTBETR#
                        Else
                            VglWert# = AVP#
                        End If
'                        If (VglWert > 0) Then
                            If (RezArtikel(i%).Preis <> VglWert) Then
                                h = RezArtikel(i).pzn + "  " + RezArtikel(i).text + "  " + Format(RezArtikel(i).Preis, "0.00") + "  " + Format(VglWert, "0.00")
                                PreisDiff_Log% = FreeFile
                                Open "PreisDif.log" For Append As #PreisDiff_Log
                                Print #PreisDiff_Log%, Format(Now, "DD.MM.YYYY HH:MM:SS  ") + RezNr + "  " + h$ + "  A+V"
                                Close #PreisDiff_Log%
                                
'                                PreisDiffStr = PreisDiffStr + h + vbCrLf
                            End If
'                        End If
                    End If
                End If
            End If
        Next i%
    End If
    RezeptHolenZuz# = 0

End If

Call CheckVerordnungMDB

'ActVerordnung% = 1

RezeptHolenMDB% = True

Call DefErrPop
End Function

Sub CheckAutIdem(pos%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("CheckAutIdem")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
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
Dim IstDrüber%, Sonderregelnr%, ok%, AusnahmeErsetzungFl
Dim BtmDerivatMenge#
Dim SQLStr$, SQLStr2$, SQLStr3$, SqlSonderregelNr$, h$
Dim iRec As New ADODB.Recordset
'Dim tRec2 As Recordset
Dim tRec2 As New ADODB.Recordset
Dim RabattAutidem As Boolean
        
RezArtikel(pos%).N0 = Abs(TaxeRec!NGrösse = "N0")

Call WinRezDebug("CheckAutIdem: " + RezArtikel(pos).pzn)

If (AutIdemOk%) Then
        Call WinRezDebug("CheckAutIdem: " + "AutIdemOk")
    If (TaxeRec!AutIdemPflicht) Or (TaxeRec!AutIdemNeu) Then
        AusnahmeErsetzungFl = 0
        On Error Resume Next
        AusnahmeErsetzungFl = (CheckNullLong(TaxeRec!AusnahmeErsetzungFl) = 1)
        On Error GoTo DefErr
        If (AusnahmeErsetzungFl) Then
            Call WinRezDebug("CheckAutIdem: " + "Ausnahme...")
            Beep
        Else
            RezArtikel(pos%).AutIdem = 1
            If (TaxeRec!AutIdemPflicht) Then
                RezArtikel(pos%).AutIdem = 2
            End If
            
            Call WinRezDebug("CheckAutIdem: " + "AutIdem=" + CStr(RezArtikel(pos).AutIdem))
            If (AutIdemKbvNr& > 0) Then
                SqlSonderregelNr$ = ""
                SQLStr$ = "SELECT * FROM AutIdemSonderregelKrankenkasse WHERE KbvNr = " + Str$(AutIdemKbvNr&)
                Call WinRezDebug("CheckAutIdem: " + SQLStr)
    '            Set iRec = AutIdemDB.OpenRecordset(SQLStr$)
                iRec.Open SQLStr, AutIdemDB.ActiveConn
                Do
                    If (iRec.EOF) Then
                        Exit Do
                    End If
                    
                    Sonderregelnr% = Format(CheckNullInt(iRec!AutIdemSonderregelNr), "0")
                    If (SqlSonderregelNr$ = "") Then
                        SqlSonderregelNr$ = "SELECT * FROM AutIdemSonderregelArtikel WHERE ("
                    Else
                        SqlSonderregelNr$ = SqlSonderregelNr$ + " OR"
                    End If
                    SqlSonderregelNr$ = SqlSonderregelNr$ + " AutIdemSonderregelNr = " + Str$(Sonderregelnr%)
                    
                    iRec.MoveNext
                Loop
                iRec.Close
                
                If (SqlSonderregelNr$ <> "") Then
                    SQLStr$ = "SELECT * FROM PZNSUBST WHERE PZN = " + Str$(TaxeRec!pzn)
                    Call WinRezDebug("CheckAutIdem: " + SQLStr)
    '                Set AutIdemRec = AutIdemDB.OpenRecordset(SQLStr$)
                    AutIdemRec.Open SQLStr, AutIdemDB.ActiveConn
                    Do
                        If (AutIdemRec.EOF) Then
                            Exit Do
                        End If
                        
    '                    If (AutIdemRec!pzn = AutIdemRec!Substpzn) Then
    '                        Beep
    '                    Else
                            SQLStr$ = SqlSonderregelNr$ + ")"
    '                        SQLStr$ = "SELECT * FROM AutIdemSonderregelArtikel"
    '                        SQLStr$ = SQLStr$ + " WHERE AutIdemSonderregelNr = " + Str$(SonderregelNr%)
                            SQLStr$ = SQLStr$ + " AND NrKz='P' AND ZoNr=" + Str$(AutIdemRec!Substpzn)
                            
    ''''''''''''''
                            BtmDerivatMenge# = 0
                            If (BtmDerivatMengeIgnorieren) Then
                            Else
                                On Error Resume Next
                                BtmDerivatMenge# = TaxeRec!BtmDerivatMenge
                                On Error GoTo DefErr
                                If (BtmDerivatMenge >= 0.01) Then
                                    SQLStr3$ = "SELECT * FROM M2Inhalt WHERE M2Nr = " + Mid(TaxeRec!M2, 2, 9)
                    '                Set AutIdemRec = AutIdemDB.OpenRecordset(SQLStr$)
                                    On Error Resume Next
                                    tRec2.Close
                                    Err.Clear
                                    On Error GoTo DefErr
                                    tRec2.Open SQLStr3, AutIdemDB.ActiveConn
        '                            If Not tRec2.NoMatch Then
                                    If Not (tRec2.EOF) Then
                                        If (InStr(UCase(CheckNullStr(tRec2!Inhalt)), "HYDROMORPHON") > 0) Then
                                            BtmDerivatMenge = 0
                                        End If
                                    End If
                                    tRec2.Close
                                End If
                            End If
                            
                            ok = True
                            If (BtmDerivatMenge >= 0.01) Then
                                SQLStr2$ = "SELECT * FROM TAXE WHERE PZN = " + Str(AutIdemRec!Substpzn)
                                'Set tRec2 = TaxeDB.OpenRecordset(SQLStr2$)
                                On Error Resume Next
                                tRec2.Close
                                Err.Clear
                                On Error GoTo DefErr
                                tRec2.Open SQLStr2, taxeAdoDB.ActiveConn
    '                            If Not tRec2.NoMatch Then
                                If Not (tRec2.EOF) Then
                                    If dCheckNull(tRec2!BtmDerivatMenge) <> BtmDerivatMenge Then
                                        ok = False
                                    End If
                                End If
                            End If
                            If ok Then
    ''''''''''''''
    '                            Set iRec = AutIdemDB.OpenRecordset(SQLStr$)
    '                            If (iRec.RecordCount > 0) Then
                                iRec.Open SQLStr, AutIdemDB.ActiveConn
                                If Not (iRec.EOF) Then
                                    If (TaxeRec!pzn = AutIdemRec!Substpzn) Then
                                        RezArtikel(pos%).AutIdem = 3
                                        Exit Do
                                    Else
                                        RezArtikel(pos%).AutIdem = 4    '2
                                    End If
                                End If
                                iRec.Close
                            End If
    '                    End If
                        
                        AutIdemRec.MoveNext
                    Loop
                    AutIdemRec.Close
                End If
            End If
        End If
        Call WinRezDebug("CheckAutIdem: " + "Ende AutIdem=" + CStr(RezArtikel(pos).AutIdem))
    Else
        If (TaxeRec!Identgruppe > 0&) Then
            If (GibtsOpAutIdem%) Then RezArtikel(pos%).AutIdem = 1
        End If
            
        If (AutIdemKbvNr& > 0) Then
            SqlSonderregelNr$ = ""
            SQLStr$ = "SELECT * FROM AutIdemSonderregelKrankenkasse WHERE KbvNr = " + Str$(AutIdemKbvNr&)
'            Set iRec = AutIdemDB.OpenRecordset(SQLStr$)
            iRec.Open SQLStr, AutIdemDB.ActiveConn
            Do
                If (iRec.EOF) Then
                    Exit Do
                End If
                
                Sonderregelnr% = Format(CheckNullInt(iRec!AutIdemSonderregelNr), "0")
                If (SqlSonderregelNr$ = "") Then
                    SqlSonderregelNr$ = "SELECT * FROM AutIdemSonderregelArtikel WHERE ("
                Else
                    SqlSonderregelNr$ = SqlSonderregelNr$ + " OR"
                End If
                SqlSonderregelNr$ = SqlSonderregelNr$ + " AutIdemSonderregelNr = " + Str$(Sonderregelnr%)
                
                iRec.MoveNext
            Loop
            iRec.Close
            
            If (SqlSonderregelNr$ <> "") Then
                SQLStr$ = SqlSonderregelNr$ + ")"
                SQLStr$ = SQLStr$ + " AND NrKz='P' AND ZoNr=" + Str$(TaxeRec!pzn)
                iRec.Open SQLStr, AutIdemDB.ActiveConn
                If Not (iRec.EOF) Then
                    RezArtikel(pos%).AutIdem = 3
                End If
                iRec.Close
            End If
        End If
    End If

    If (PrivatRezept% = 0) And (kKassenOk%) And (AutIdemKbvNr > 0) And ((RezArtikel(pos%).flag And REZ_WG_4) = 0) Then
        RabattAutidem = CheckRabattAutIdem(TaxeRec!pzn)
        If Not RabattAutidem And Not IstRabattAutIdem(TaxeRec!pzn) Then
            If (CheckRabattAutIdem(TaxeRec!pzn, 1)) Then
                Call iMsgBox("Rabattartikel mit abweichender DAR oder N-Größe vorhanden", vbInformation, Trim(RezArtikel(pos).text))
            End If
        End If
    End If
ElseIf (TaxeRec!Identgruppe > 0&) Then
    IstDrüber% = False
    If (TaxeRec!FestKz) And (TaxeRec!FESTBETRAG < TaxeRec!vk) Then IstDrüber% = True
    
    If (GibtsOpAutIdem%) Then
        RezArtikel(pos%).AutIdem = 1
        If (IstDrüber% And (TaxeRec!vk <= TaxeRec!FESTBETRAG)) Then
            RezArtikel(pos%).AutIdem = 2
        End If
    End If
End If
       
Call DefErrPop
End Sub

Function GibtsOpAutIdem%()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("GibtsOpAutIdem%")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
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
Dim i%, ret%, ok%, ind%
Dim h$, h2$, SQLStr$, OrgPznIndik$, PznIndik$
Dim iRec As Recordset
        
ret% = False

'SQLStr$ = "SELECT * FROM TAXE WHERE IDENTGRUPPE = " + Str$(TaxeRec!IdentGruppe)
'SQLStr$ = SQLStr$ + " and Menge = """ + TaxeRec!menge + """ and Einheit = """ + TaxeRec!einheit + """"
'SQLStr$ = SQLStr$ + " and Festbetrag = " + Str$(TaxeRec!FESTBETRAG)
'SQLStr$ = SQLStr$ + " ORDER BY VK"
'Set iRec = TaxeDB.OpenRecordset(SQLStr$)
'If Not (iRec.EOF) Then
'    iRec.MoveLast
'End If
'If (iRec.RecordCount > 1) Then
'    iRec.MoveFirst
'    OrgPznIndik$ = GetPznIndik(TaxeRec!pzn)
'    If (OrgPznIndik$ = "") Then
'        ret% = True
'    Else
'        iRec.MoveFirst
'        Do
'            If (iRec.EOF) Then
'                Exit Do
'            End If
'
'            If (iRec!pzn <> TaxeRec!pzn) Then
'                PznIndik$ = GetPznIndik(iRec!pzn)
'                If (PznIndik$ <> "") Then
'                    PznIndik$ = "," + PznIndik$
'                End If
'
'                ok% = True
'                h$ = OrgPznIndik$
'                Do
'                    ind% = InStr(h$, ",")
'                    If (ind% = 0) Then
'                        Exit Do
'                    Else
'                        h2$ = "," + Left$(h$, ind%)
'                        h$ = Mid$(h$, ind% + 1)
'                        If (InStr(PznIndik$, h2$) = 0) Then
'                            ok% = 0
'                            Exit Do
'                        End If
'                    End If
'                Loop
'                If (ok%) Then
'                    ret% = True
'                    Exit Do
'                End If
'            End If
'
'            iRec.MoveNext
'        Loop
'    End If
'End If

GibtsOpAutIdem% = ret%
       
Call DefErrPop
End Function

Function HoleOpAutIdem$()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("HoleOpAutIdem$")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
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
Dim i%, ok%, ind%, tmp%
Dim ret$, h$, h2$, SQLStr$, OrgPznIndik$, PznIndik$
'Dim iRec As Recordset
Dim iRec As New ADODB.Recordset
        
ret$ = ""

SQLStr$ = "SELECT * FROM TAXE WHERE IDENTGRUPPE = " + Str$(TaxeRec!Identgruppe)
'SQLStr$ = SQLStr$ + " and Menge = """ + TaxeRec!menge + """ and Einheit = """ + TaxeRec!einheit + """"
SQLStr$ = SQLStr$ + " and Menge = '" + TaxeRec!menge + "' and Einheit = '" + TaxeRec!einheit + "'"
SQLStr$ = SQLStr$ + " and Festbetrag = " + Str$(TaxeRec!FESTBETRAG)
SQLStr$ = SQLStr$ + " ORDER BY VK"
'Set iRec = TaxeDB.OpenRecordset(SQLStr$)
On Error Resume Next
iRec.Close
Err.Clear
On Error GoTo DefErr
iRec.Open SQLStr, taxeAdoDB.ActiveConn
If Not (iRec.EOF) Then
    iRec.MoveLast
End If
If (iRec.RecordCount > 0) Then
    OrgPznIndik$ = GetPznIndik(TaxeRec!pzn)
    
    i% = 1
    ret$ = "SELECT * FROM TAXE WHERE PZN = "
    
    iRec.MoveFirst
    Do
        If (iRec.EOF) Then
            Exit Do
        End If
        
        ok% = True
        If (OrgPznIndik$ <> "") Then
            PznIndik$ = GetPznIndik(iRec!pzn)
            If (PznIndik$ <> "") Then
                PznIndik$ = "," + PznIndik$
            End If
            
            h$ = OrgPznIndik$
            Do
                ind% = InStr(h$, ",")
                If (ind% = 0) Then
                    Exit Do
                Else
                    h2$ = "," + Left$(h$, ind%)
                    h$ = Mid$(h$, ind% + 1)
                    If (InStr(PznIndik$, h2$) = 0) Then
                        ok% = 0
                        Exit Do
                    End If
                End If
            Loop
        End If
        If (ok%) Then
            If (i% > 1) Then ret$ = ret$ + " OR PZN = "
            ret$ = ret$ + Str$(iRec!pzn)
            i% = i% + 1
        End If
        
        iRec.MoveNext
    Loop
    ret$ = ret$ + " ORDER BY VK"
End If

HoleOpAutIdem$ = ret$
       
Call DefErrPop
End Function

Function GetPznIndik$(pzn&)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("GetPznIndik$")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
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
Dim ret$, SQLStr$
Dim iRec As New ADODB.Recordset

ret$ = ""
    
SQLStr$ = "SELECT * FROM PZNINDIK WHERE PZN = " + Str$(pzn&)
SQLStr$ = SQLStr$ + " ORDER BY INDIKBEREICHNR"
'Set iRec = AutIdemDB.OpenRecordset(SQLStr$)
iRec.Open SQLStr, AutIdemDB.ActiveConn
Do
    If (iRec.EOF) Then
        Exit Do
    End If
    
    ret$ = ret$ + Format(iRec!IndikbereichNr, "0") + ","
    
    iRec.MoveNext
Loop
iRec.Close
'If (ret$ <> "") Then
'    ret$ = ret$ + ","
'End If

GetPznIndik$ = ret$
       
Call DefErrPop
End Function

Sub HoleAusRezeptSpeicher(Optional MarsRezNr$ = "")
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("HoleAusRezeptSpeicher")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
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
Dim i%, j%, k%, IstDrüber%, hWg%
Dim lRezNr2&, iKuNr&
Dim h$, h2$, hStr$, Unique$, SQLStr$, sVerfügbarkeit$, MarsRezNr2$
Dim aktavp As Double, OrgAvp As Double
Dim iRec As Recordset
    
lRezNr2& = lRezNr&
Call InitRezept
lRezNr& = lRezNr2&
SchonDa% = True

'    Set RezepteRec = RezSpeicherDB.OpenRecordset("Rezepte", dbOpenTable)
If (MarsRezNr <> "") Then
    Set ArtikelRec = MarsRezSpeicherDB.OpenRecordset("Artikel", dbOpenTable)
    Set AuswertungRec = MarsRezSpeicherDB.OpenRecordset("Auswertung", dbOpenTable)
    SQLStr = "SELECT * FROM Rezepte WHERE RezeptNr='" + MarsRezNr + "'"
    Set RezepteRec = MarsRezSpeicherDB.OpenRecordset(SQLStr$)
    If (RezepteRec.EOF) Then
        MarsRezNr2 = ""
        If (Left(MarsRezNr, 3) = "989") Then
            hStr$ = "9999998" + Mid(MarsRezNr, 8, 5)
            Call PruefZiffer(hStr$)
            MarsRezNr2$ = hStr$
            SQLStr = "SELECT * FROM Rezepte WHERE RezeptNr='" + MarsRezNr2 + "'"
            Set RezepteRec = MarsRezSpeicherDB.OpenRecordset(SQLStr$)
        End If
    End If
    

Else
    Set ArtikelRec = RezSpeicherDB.OpenRecordset("Artikel", dbOpenTable)
    Set AuswertungRec = RezSpeicherDB.OpenRecordset("Auswertung", dbOpenTable)
End If

AnzRezeptArtikel% = RezepteRec!AnzArtikel

Unique$ = RezepteRec!Unique

ActVebNr = CheckNullLong(RezepteRec!VebNr)
ActPauschaleNr = CheckNullLong(RezepteRec!PauschaleNr)
                            
ArtikelRec.index = "Unique"
For j% = 0 To (AnzRezeptArtikel% - 1)
    
    
    h$ = Unique$ + Format(j% + 1, "0")
    If j% >= 9 And j% <= 243 Then
        h$ = Unique$ + Chr$(j% + 65)   'für Privatrezepte
    End If
    
    ArtikelRec.Seek "=", h$

    If (ArtikelRec.NoMatch = False) Then
        RezArtikel(j%).appli = 0
        RezArtikel(j%).Fam = 0
        RezArtikel(j%).AutIdem = 0
        RezArtikel(j%).N0 = 0
        RezArtikel(j%).MagIndex = 0
        RezArtikel(j%).NichtInTaxe = 0
''
        If (ArtikelRec!pzn < 0) Then
            RezArtikel(j%).pzn = "-" + Format(Abs(ArtikelRec!pzn), "0000000")
        Else
            RezArtikel(j%).pzn = PznString(ArtikelRec!pzn)
        End If

        RezArtikel(j%).text = ArtikelRec!text
        RezArtikel(j%).flag = ArtikelRec!flag
        RezArtikel(j%).Zuz = ArtikelRec!Zuz
        RezArtikel(j%).HerstRabattPrivat130Brutto = CheckNullDouble(ArtikelRec!HerstRabattPrivat130Brutto)
        RezArtikel(j).VdbPauschale = 0
        RezArtikel(j%).Imp123 = ArtikelRec!Imp
        RezArtikel(j%).ImpErspart = ArtikelRec!ImpErspart

        RezArtikel(j%).Faktor = ArtikelRec!Faktor
        If IsNull(ArtikelRec!ePreis) Then
            RezArtikel(j%).Preis = ArtikelRec!Preis
            If (RezArtikel(j%).Faktor > 1) Then
                RezArtikel(j%).Preis = RezArtikel(j%).Preis / RezArtikel(j%).Faktor
            End If
        Else
            RezArtikel(j%).Preis = ArtikelRec!ePreis
        End If

        RezArtikel(j%).ScreenPzn = ArtikelRec!ScreenPzn
        RezArtikel(j%).TkkPznDruck = CheckNullDouble(ArtikelRec!TkkPznDruck)
'        RezArtikel(j%).AutIdemKreuz = CheckNullDouble(ArtikelRec!AutIdemKreuz)
        
        RezArtikel(j%).VebNr = CheckNullLong(ArtikelRec!VebNr)
        RezArtikel(j%).PauschaleNr = CheckNullLong(ArtikelRec!PauschaleNr)
        If (RezArtikel(j).VebNr = 0) And (RezArtikel(j).PauschaleNr = 0) Then
            RezArtikel(j%).VebNr = ActVebNr
            RezArtikel(j%).PauschaleNr = ActPauschaleNr
        End If
    

'Neuer Bereich ?????
        RezArtikel(j%).AutIdemKreuz = CheckNullByte(ArtikelRec!AutIdemKreuz)
        RezArtikel(j%).Imp123 = CheckNullByte(ArtikelRec!Imp)
        RezArtikel(j%).ImpErspart = CheckNullDouble(ArtikelRec!ImpErspart)
        RezArtikel(j%).Warenzeichen = CheckNullByte(ArtikelRec!Warenzeichen)
        
        RezArtikel(j).Fam = CheckNullByte(ArtikelRec!Fam)
        RezArtikel(j).AutIdem = CheckNullByte(ArtikelRec!AutIdem)
        RezArtikel(j).N0 = CheckNullByte(ArtikelRec!N0)
        RezArtikel(j).appli = CheckNullByte(ArtikelRec!appli)
        
        RezArtikel(j).zusatz(0) = CheckNullStr(ArtikelRec!zusatz1)
        RezArtikel(j).zusatz(1) = CheckNullStr(ArtikelRec!Zusatz2) + CheckNullStr(ArtikelRec!Zusatz2a)
        
        RezArtikel(j).NichtInTaxe = CheckNullByte(ArtikelRec!NichtInTaxe)
        RezArtikel(j).IstWg4 = CheckNullByte(ArtikelRec!IstWg4)
        RezArtikel(j).PlusMehrKosten = CheckNullByte(ArtikelRec!PlusMehrKosten)
        RezArtikel(j).ZuzahlungsErlass = CheckNullByte(ArtikelRec!ZuzahlungsErlass)
        RezArtikel(j).MietDauer = CheckNullInt(ArtikelRec!MietDauer)
        RezArtikel(j).VdbPauschale = CheckNullByte(ArtikelRec!VdbPauschale)
        
        RezArtikel(j).Fortsetzung = CheckNullByte(ArtikelRec!Fortsetzung)
        
        RezArtikel(j).HmAbrechnungsKz = CheckNullStr(ArtikelRec!HmAbrechnungsKz)
        RezArtikel(j).LEGS = CheckNullStr(ArtikelRec!LEGS)
    End If
Next j%

For j% = 0 To (AnzRezeptArtikel% - 1)
    If (RezArtikel(j).pzn = SONDERPZN_WVO) Then
        WVORabattArtikel = Right("111" + CStr(RezArtikel(j).Faktor), 3)
        For k% = (j + 1) To (AnzRezeptArtikel% - 1)
            RezArtikel(k - 1) = RezArtikel(k)
        Next k
        AnzRezeptArtikel = AnzRezeptArtikel - 1
        Exit For
    End If
Next

RezNr$ = RezepteRec!RezeptNr
ActKkasse$ = RezepteRec!Kkasse
GebFrei% = RezepteRec!GebFrei
RezGebSumme# = RezepteRec!RezGebSumme
RezSumme# = RezepteRec!RezSumme

ActBundesland% = RezepteRec!Bundesland
If (ActBundesland% >= 1000) Then
    IvfRezept% = True
    ActBundesland% = ActBundesland% - 1000
End If

ActKasse% = RezepteRec!Kasse
ActVerordnung% = RezepteRec!Verordnung

PrivatRezept% = (Left$(ActKkasse$, 6) = "Privat")
If (ActKkasse$ = "Unbekannt") Then ActKkasse$ = ""

'If (AutIdemSonderregelOk%) And (Val(ActKkasse$) > 0) Then
'    SQLStr$ = "SELECT * FROM kKassen WHERE IK = " + Str$(Val(ActKkasse$))
'    Set KkassenRec = kKassenDB.OpenRecordset(SQLStr$)
'    If (KkassenRec.RecordCount > 0) Then
'        AutIdemIk& = CheckNullLong(KkassenRec!IK)
'        AutIdemKbvNr& = CheckNullLong(KkassenRec!KbvNr)
'    End If
'End If

h$ = RezepteRec!VerkDatum
VerkDatum$ = Right$(h$, 2) + Mid$(h$, 3, 2) + Left$(h$, 2)
OrgVerkDatum$ = VerkDatum$

VerkZeit% = RezepteRec!VerkZeit
VerkUser$ = RezepteRec!VerkComp

MarsAutomaticDatum(0) = Left(VerkDatum, 2) + "." + Mid(VerkDatum, 3, 2) + ".20" + Mid(VerkDatum, 5, 2) + " " + Format(VerkZeit \ 100, "00") + ":" + Format(VerkZeit Mod 100, "00")

PersonalNr% = RezepteRec!PersonalNr
    
DruckDatum$ = RezepteRec!DruckDatum
If (MarsRezNr <> "") Then
    DruckZeit% = RezepteRec!DruckZeit
    MarsFreigabePersonalNr% = RezepteRec!DruckPersonalNr
End If
RezApoNr$ = RezepteRec!InstitutsKz
'    RezepteRec!fam = RezGesamtFAM#
'    RezepteRec!ImpFähig = RezImpFähig#
'    RezepteRec!ImpIst = RezImpIst#
'    RezepteRec!DruckZeit = Mid$(Unique$, 7, 4)

'h$ = Format(RezepteRec!AvpMonat, "0000")
'AvpRezeptNr$ = Format(RezepteRec!AvpLaufNr, "000000") + Mid$(h$, 3, 2) + Left$(h$, 2)
AvpLaufNr& = CheckNullLong(RezepteRec!AvpLaufNr)
If (AvpLaufNr& > 0) Then
    AvpRezeptNr$ = RezepteRec!AvpRezeptNr
End If

If (IsNull(RezepteRec!Verfügbarkeit)) Then
    AMverfügbarkeit$ = ""
Else
    AMverfügbarkeit$ = RezepteRec!Verfügbarkeit
End If

TransaktionsID = CheckNullLong(RezepteRec!TransaktionsID)
ParenteralHash = CheckNullStr(RezepteRec!ParenteralHash)
ParenteralRezept = Abs(ParenteralHash <> "") - 1 'Abs(TransaktionsId > 0) - 1
rzLieferId = CheckNullStr(RezepteRec!rzLieferId)
    
RezeptMitHilfsmittelNr% = CheckNullInt(RezepteRec!P302)
    
ActVebNr = CheckNullLong(RezepteRec!VebNr)
ActPauschaleNr = CheckNullLong(RezepteRec!PauschaleNr)
                            
AutIdemIk& = CheckNullLong(RezepteRec!kKassenIk)

VerkLiRe = CheckNullByte(RezepteRec!VerkLiRe)
AutIdemKbvNr = CheckNullLong(RezepteRec!AutIdemKbvNr)
ActAVWGik = CheckNullLong(RezepteRec!AVWGik)

If (ParenteralRezept >= 0) Then
    SQLStr$ = "SELECT * FROM ParenteralTaxierungen WHERE RezepteUnique='" + Unique + "'"
    SQLStr$ = SQLStr$ + " ORDER BY AnfMagInd"
    Set iRec = RezSpeicherDB.OpenRecordset(SQLStr$)
    If Not (iRec.EOF) Then
        MAG_SPEICHER% = FileOpen("MagTax" + Trim(para.User), "RW", "B")
        If (MAG_SPEICHER% > 0) Then
            MagSpeicherIndex% = LOF(MAG_SPEICHER%) / Len(TaxierRec) + 1
            
            j% = 0
            Seek #MAG_SPEICHER%, ((MagSpeicherIndex% - 1) * CLng(Len(TaxierRec))) + 1
            Do
                If (iRec.EOF) Then
                    Exit Do
                End If
                
                If (j = 0) Then
                    For i% = 0 To UBound(ParenteralPzn)
                        If (iRec!pzn = Val(ParenteralPzn(i))) Then
                            ParenteralRezept = i
                            Exit For
                        End If
                    Next i
                End If
                TaxierRec.ActPreis = iRec!ActPreis
                TaxierRec.ActMenge = iRec!ActMenge
                TaxierRec.Meh = iRec!einheit
                TaxierRec.kurz = iRec!text
                TaxierRec.pzn = PznString(iRec!pzn)
                TaxierRec.flag = iRec!flag
                TaxierRec.kp = iRec!kp
                TaxierRec.GStufe = iRec!GStufe
                TaxierRec.Verwurf = CheckNullByte(iRec!Verwurf)
                
                If (Trim(TaxierRec.kurz) <> "") Then
                    Put #MAG_SPEICHER%, , TaxierRec
                    j% = j% + 1
                    
                    If (ParenteralRezept > 15) And (ParenteralRezept < 24) Then
                        If (InStr(UCase(TaxierRec.kurz), "E-DOSIS") > 0) Then
                            SubstitutionsMg = TaxierRec.GStufe
                            SubstitutionsEinzelpreis = TaxierRec.kp
                            
                            Dim SubstitutionsAbgabenStr As Integer
                            SubstitutionsAbgabenStr = TaxierRec.ActMenge
                            For i = 2 To 0 Step -1
                                SubstitutionsAbgaben(i) = SubstitutionsAbgabenStr Mod 10
                                SubstitutionsAbgabenStr = (SubstitutionsAbgabenStr \ 10)
                            Next i
                        End If
                    End If
                End If
                
                iRec.MoveNext
            Loop
            
            For k% = (j% + 1) To 100
                TaxierRec.flag = 255
                Put #MAG_SPEICHER%, , TaxierRec
            Next k%
            
            Close #MAG_SPEICHER%
            
            RezArtikel(0).MagIndex = MagSpeicherIndex%
        End If
    End If
End If


If (ActVebNr = 0) And (ActPauschaleNr = 0) Then
    If (AutIdemIk& > 0) Then
        SQLStr = "SELECT count(*) AS iAnz FROM VdbVVI LEFT JOIN VdbVVL ON VdbVVI.VebNr=VdbVVL.VebNr"
        SQLStr = SQLStr + " WHERE VdbVVI.Ik=" + CStr(AutIdemIk)
        SQLStr = SQLStr + " AND VdbVVL.LandNr=" + CStr(ActBundesland)
        FabsErrf = AplusVDB.OpenRecordset(VdbBedRec, SQLStr)
        If Not (VdbBedRec.EOF) Then
            If (CheckNullLong(VdbBedRec!iAnz) = 1) Then
                SQLStr = "SELECT VdbVVI.VebNr FROM VdbVVI LEFT JOIN VdbVVL ON VdbVVI.VebNr=VdbVVL.VebNr"
                SQLStr = SQLStr + " WHERE VdbVVI.Ik=" + CStr(AutIdemIk)
                SQLStr = SQLStr + " AND VdbVVL.LandNr=" + CStr(ActBundesland)
                FabsErrf = AplusVDB.OpenRecordset(VdbBedRec, SQLStr)
                If Not (VdbBedRec.EOF) Then
                    ActVebNr = CheckNullLong(VdbBedRec!VebNr)
                End If
            End If
        End If
    End If
    If (ActVebNr = 0) And (ActKasse% > 0) Then
        SQLStr = "SELECT Top 1* FROM VdbBedingungen"
        SQLStr = SQLStr + " WHERE KostentrNr=" + CStr(ActKasse)
        SQLStr = SQLStr + " AND LandNr=" + CStr(ActBundesland)
        FabsErrf = AplusVDB.OpenRecordset(VdbBedRec, SQLStr)
        If Not (VdbBedRec.EOF) Then
            ActVebNr = CheckNullLong(VdbBedRec!VebNr)
        End If
    End If
End If

If (ActVebNr > 0) Then
    For j% = 0 To UBound(AvVereinbarungen$)
        h2 = AvVereinbarungen(j)
        If (Val(Left(h2, 10)) = ActVebNr) Then
            If (InStr(UCase(Trim(Mid$(h2, 12, 100))), "BERUFSGENOSSENSCHAFT") > 0) Then
                IkSpezialName = Trim(Mid$(h2, 12, 100))
            End If
            
            Exit For
        End If
    Next j%
End If

For i% = 0 To (AnzRezeptArtikel% - 1)
    If (RezArtikel(i%).NichtInTaxe = 0) Then
        h$ = RezArtikel(i%).pzn
        If (Left(h, 1) = "-") Then
            h = Mid(h, 2) + "0"
            h = MakePzn(h)
        End If
        
        SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + h$
        'Set TaxeRec = TaxeDB.OpenRecordset(SQLStr$)
        On Error Resume Next
        TaxeRec.Close
        Err.Clear
        On Error GoTo DefErr
        TaxeRec.Open SQLStr, taxeAdoDB.ActiveConn
        If Not (TaxeRec.EOF) Then
            If (TaxeRec!BtmKz) Then
                BtmRezept% = True
            End If
        End If
    End If
Next i%

SonderBelegRezept% = False
If (AnzRezeptArtikel% = 1) Then
    For i% = 1 To AnzSonderBelege%
        If (SonderBelege(i% - 1).pzn = RezArtikel(0).pzn) Then
            SonderBelegRezept% = i%
            PrivatRezept% = 0
            BtmRezept% = 0
            IvfRezept% = 0
            RezKundenInfo$(1) = ""
            RezKundenInfo$(2) = "SONDERBELEG"
            Exit For
        End If
    Next i%
End If
'If (SonderBelegRezept% > 0) Then
    On Error Resume Next
    iKuNr = RezepteRec!knr
    On Error GoTo DefErr
    If (iKuNr > 0) Then Call HoleKundenInfo(iKuNr)
'End If


Call DefErrPop
End Sub

Sub ZeigeStatistik()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("ZeigeStatistik")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
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
Dim row%
Dim TaskId&
Dim pzn$, ch$


row% = AktTabIndex% \ 4
pzn$ = RezArtikel(row%).pzn
If (Val(pzn$) = 0) Then Call DefErrPop: Exit Sub

'Call ZeigeStatbild(pzn$, ParentForm)
'AppActivate ParentForm.Caption

TaskId& = Shell("WinStat.exe " + pzn$, vbNormalFocus)

Call DefErrPop
End Sub

Sub InitRezept(Optional InitModus% = 3)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("InitRezept")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
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

If (AlteTaxeAktiv) Then
    taxeAdoDB.CloseDB
    AplusVDB.CloseDB
    AutIdemDB.CloseDB
    kKassenDB.CloseDB
    
    taxeAdoDB.OpenDB
    AplusVDB.OpenDB
    AutIdemDB.OpenDB
    kKassenDB.OpenDB
    
    AlteTaxeAktiv = False
End If


If (InitModus% And &H1) Then
    GebFrei% = False
    ImpfRez% = False
    ActBundesland% = OrgBundesland%
    ActKasse% = -1
    ActKkasse$ = ""
    ActAVWGik& = -1
    PrivatRezept% = False
    eRezeptTaskId = ""
    eRezeptStatusWerte = ""
    eRezeptPzn = ""
    eRezeptAMverfügbarkeit = ""
    eRezeptOpStatus = 0
    eRezeptVerordnungsTyp = 0
    BtmRezept% = False
    IvfRezept% = False
    SonderBelegRezept% = False
    Selbsterklaerung = False
    SchutzMasken = 0
    SchonDa% = False
    
'    MsgBox ("taskid weg")
    
    BereitsGedruckt% = False
    
    ActKiste% = 0
    
    On Error Resume Next
    Kill "MagTax" + Trim(para.User)
    ParentForm.lblWinVk.Visible = 0
    On Error GoTo DefErr
    
    RezNr$ = ""
    lRezNr& = 0
    DruckDatum$ = Space$(6)
    
    sBrutto$ = ""
    
    AvpRezeptNr$ = ""
    AutIdemKbvNr& = 0
    AutIdemIk& = 0
    
    AMverfügbarkeit$ = ""
    RezeptMitHilfsmittelNr% = -1
    BtmFaktorManuell = 0
    
    TransaktionsID& = 0
    ParenteralRezept = -1
    Parenteral_AOK_LosGebiet% = 0
    Parenteral_AOK_NordOst% = 0
    ParenteralHash = ""
    rzLieferId = ""
    pCharge = ""
    
    AuseinzelungPzn(0) = ""
    AuseinzelungBtm = 0
    
    ActVebNr = 0
    ActPauschaleNr = 0
    
    IkSpezialName = ""
    
    WVORabattArtikel = String(10, "1")
    
    MarsFreigabePersonalNr = -1
    
    ZusatzKomponentenNr = ""
    ZusatzKomponentenBasisNr = ""
    
    SubstitutionsMg = 0
    For i = 0 To 2
        SubstitutionsAbgaben(i) = 0
    Next
    SubstitutionsEinzelpreis = 0

    LennartzPzn = 0
End If

If (InitModus% And &H2) Then
    For i% = 0 To MAX_VERK_ZEILEN%
        Call InitRezeptArtikel(i%)
    Next i%
        
    For i% = 0 To UBound(BarVerkauf$)
        BarVerkauf$(i%) = ""
    Next i%
    
    RezGebSumme# = 0
    RezSummeDazu# = 0
    RezGebRechnen% = True
    
    VerkDatum$ = Format(Now, "DDMMYY")
    OrgVerkDatum$ = VerkDatum$
    VerkZeit% = Val(Format(Now, "HHMM"))   ' 0
    VerkUser$ = para.User   ' ""
    VerkLiRe = 0
    PersonalNr% = 0
    
    For i% = 0 To 5
        RezKundenInfo$(i%) = ""
    Next i%
    
    AnzRezeptArtikel% = 0

    LennartzPzn = 0
End If
    
Call DefErrPop
End Sub

Function RezKontrInit%()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("RezKontrInit%")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
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
Dim i%, PARAM_DAT%, DosDruckerInd%, erg%
Dim l&
Dim h$
  
'  ProgrammTyp = REZ_KONTR;
'  StartMagZeile = TAX_START_MAG_ZEILE;
  
    RezSpeicherOK% = OpenRezeptSpeicher%
    If (RezSpeicherOK% = False) Then
        Call iMsgBox("Rezeptspeicher momentan nicht verwendbar !")
    End If
'    ParentForm.Caption = "Rezeptkontrolle - Abrechnungsmonat " + Format(CDate("01." + Mid(AbrechMonat$, 3, 2) + "." + Left(AbrechMonat$, 2)), "mmmm") + " bis " + Mid(AbrechAbgabeDatum$, 5, 2) + "." + Mid(AbrechAbgabeDatum$, 3, 2) + "." + Left(AbrechAbgabeDatum$, 2)
    If (InStr(para.Benutz, "V") > 0) And (para.Land = "D") Then
        AllesFlag% = (InStr(para.Benutz, "R") > 0)
        
        ImpFlag% = (InStr(para.Benutz, "!") > 0)
        'neu ab Ver: 4.0.33
        ImpFlag = True
        
        EasyImpFlag% = (InStr(para.Benutz, "i") > 0)
        If (EasyImpFlag%) Then ImpFlag% = True
        VmFlag% = (InStr(para.Benutz, "?") > 0)
        FiveRxFlag% = (InStr(para.Benutz, "Z") > 0) Or (InStr(para.Benutz, "z") > 0)
    Else
        RezKontrInit% = False: Call DefErrPop: Exit Function
    End If
    
'  ret = TaxierenInit();
'  RezeptTyp = 1;
'
'  if (ret == TRUE)  {
'
'    sprintf(CrLf, "%c%c", 13, 10);
'    RezZeilenSatzLen = sizeof(REZ_ZEILEN_STRUCT);
'    TaxeSatzLen = sizeof(TAXE_STRUCT);
'    AnzVerkZeilen = REZ_STOP_ZEILE - REZ_START_ZEILE + 1;
'
'    if (AllesFlag == TRUE)  {
'
'      if (VmFlag == TRUE)
'        HolAbgabePreise();

    Call HolVerbandmittel
    
    Call LadeSonderPzn("sondrpzn.dat", SonderPzn$)
    Call LadeSonderPzn("hmnummer.dat", HmNummern$)
    Call LadeSonderPzn("BtmPzns.dat", BtmSonderPzn$)
    
    Call LadeAbgabePreise
    
    Call LadeImpfungen
    
    Call HoleFiveRxPzns
    
    Call OpenArbEmb
    
'
'    /* neu in 1.84 */
'    /*  öffnen der Datei mit den Importen */
'    if (ImpFlag == TRUE)  {
'      IMPORT_DAT = OpenDatei("import.dat", O_RDONLY|O_BINARY|O_DENYNONE);
'      if (IMPORT_DAT < 0)
'        ImpFlag = FALSE;
'    } /* if */


    h$ = "0"
    l& = GetPrivateProfileString("Allgemein", "RezeptHuellen", "0", h$, 2, "\user\dp.ini")
    h$ = Left$(h$, l&)
    If (Val(h$) <> 0) Then
        RezeptHuellen% = True
    Else
        RezeptHuellen% = False
    End If

    h$ = "0"
    l& = GetPrivateProfileString("AlleKassen", "Knallrot", "0", h$, 2, "\user\vkpara.ini")
    h$ = Left$(h$, l&)
    If (Val(h$) <> 0) Then
        Knallrot% = True
    Else
        Knallrot% = False
    End If
    
    h$ = "0"
    i = GetPrivateProfileString("AlleKassen", "PrivRezDruckHoch", "0", h, 2, "\user\vkpara.ini")
    h$ = Left$(h$, l&)
    If (Val(h$) <> 0) Then
        PrivRezDruckHoch% = True
    Else
        PrivRezDruckHoch% = False
    End If

    h$ = "1"
    l& = GetPrivateProfileString("AlleKassen", "BTMDerivatmengeIgnorieren", "1", h$, 2, "\user\vkpara.ini")
    h$ = Left$(h$, l&)
    If (Val(h$) <> 0) Then
        BtmDerivatMengeIgnorieren = True
    Else
        BtmDerivatMengeIgnorieren = False
    End If
    
'    h$ = "0"
'    l& = GetPrivateProfileString("Mars", "Mars", "0", h$, 2, "\user\vkpara.ini")
'    h$ = Left$(h$, l&)
'    MarsAktiv = (Val(h$) = 1)
    
    

    h$ = "0"
    l& = GetPrivateProfileString("Allgemein", "DSKPrivatrezpreis", "0", h$, 2, "\user\dp.ini")
    h$ = Left$(h$, l&)
    PrivatPreisTyp = Val(h$)

    
    PARAM_DAT% = FileOpen("param.dat", "I")
    For i% = 0 To 22
        If (EOF(PARAM_DAT%)) Then Exit For
        
        Line Input #PARAM_DAT%, h$
        
        If (i% = 8) Then
            DosDruckerInd% = Val(h$)
        ElseIf (i% = 16) Then
            RezApoName$(0) = Trim(h$)
        ElseIf (i% = 17) Then
            RezApoName$(1) = Trim(h$)
        ElseIf (i% = 18) Then
            If (RezApoNr$ = "") Then
                RezApoNr$ = Trim(h$)
                OrgRezApoNr$ = RezApoNr$
            End If
        ElseIf (i% = 19) Then
            ZuzFeld!(4) = Val(h$)
        ElseIf (i% = 20) Then
            ZuzFeld!(3) = Val(h$)
        ElseIf (i% = 21) Then
            ZuzFeld!(5) = Val(h$)
        ElseIf (i% = 22) Then
            ZuzFeld!(6) = Val(h$)
        End If
    Next i%
    Close #PARAM_DAT%
    
    
    
    If (RezApoDruckName$ = "") Then
        RezApoDruckName$ = RezApoName$(0) + " " + RezApoName$(1)
        Call OemToChar(RezApoDruckName$, RezApoDruckName$)
        BtmRezDruckName$ = RezApoDruckName$
    End If
    
    If (RezeptDrucker$ = "") And (DosDruckerInd% > 0) Then
        erg% = GetGeraetParameter(DosDruckerInd%, RezeptDrucker$, RezeptDruckerPara$)
        If (RezeptDrucker$ = "WINPRINT") Then RezeptDrucker$ = ""
        If (UCase(Left$(RezeptDruckerPara$, 3)) = "LPT") Then RezeptDrucker$ = ""
    End If
    If (RezeptDrucker$ = "") Then RezeptDrucker$ = Printer.DeviceName
    StandardDrucker = Printer.DeviceName
    Call ParentForm.ErzeugeDruckerAuswahl
    
    Call PaintAvParameter
    
    ParentForm.mnuDateiInd(0).Enabled = DarfRezSpeicher
    ParentForm.cmdDatei(0).Enabled = DarfRezSpeicher

    ParentForm.mnuDateiInd(1).Enabled = EasyImpFlag%
    ParentForm.cmdDatei(1).Enabled = EasyImpFlag%

    AbholerMdb% = 0
    AbholerSqlDatabase = "Abholer" + sqlop.SqlGetSuffix
    AbholerSQL = sqlop.SqlCheckDatabase(AbholerSqlDatabase)
    If (AbholerSQL) Then
        AbholerMdb% = True
    Else
        AbholerMdb% = (InStr(para.Benutz, "Y") > 0) Or (InStr(para.Benutz, "y") > 0)
        If (AbholerMdb%) Then
            AbholerMdb% = (Dir$(ABHOLER_DB) <> "")
        End If
        If (AbholerMdb%) Then
            On Error Resume Next
            Err.Clear
            Set AbholerDB = OpenDatabase(ABHOLER_DB, False, True)
            If (Err.Number = 0) Then
                AbholerMdb% = True
                AbholerDB.Close
            End If
            Err.Clear
            On Error GoTo DefErr
        End If
        If (AbholerMdb%) Then
            h$ = "winabhol.exe"
            AbholerMdb% = (Dir$(h$) <> "")
        End If
    End If
    
    ParentForm.mnuDateiInd(5).Enabled = AbholerMdb%
    ParentForm.cmdDatei(5).Enabled = AbholerMdb%

    ParentForm.mnuDateiInd(6).Enabled = AvpTeilnahme%
    ParentForm.cmdDatei(6).Enabled = AvpTeilnahme%

    ParentForm.mnuDateiInd(7).Enabled = FiveRxFlag
    ParentForm.cmdDatei(7).Enabled = FiveRxFlag
    
    ParentForm.mnuDateiInd(8).Enabled = FileExist("ti_back.exe")
    ParentForm.cmdDatei(8).Enabled = ParentForm.mnuDateiInd(8).Enabled
    
'  TesteHuffman();
  
RezKontrInit% = True

Call DefErrPop
End Function

Sub HolVerbandmittel()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("HolVerbandmittel")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
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
Dim i%, erg%, VERBAND%, ubou%
Dim h$, h2$
Dim avRec As Recordset

VERBAND% = FileOpen("verbandm.dat", "RW", "B")

h$ = String(6, 0)
Seek #VERBAND%, 9
Get #VERBAND%, , h$
VmRabattFaktor# = Val(h$)

If (VmFlag%) Then

    If (AplusVOk) Then
        erg% = True
        If (erg%) Then erg% = LadeAvDateiMDB%("VdbLaender", AvLaender$)
        If (erg%) Then erg% = LadeAvDateiMDB%("VdbKostentraeger", AvKassen$)
        If (erg%) Then erg% = LadeAvVereinbarungenMDB%("VdbVereinbarungen", AvVereinbarungen$)
        If (erg%) Then erg% = LadeAvPauschalenMDB%("VdbPauschalen", AvPauschalen$)
        If (erg%) Then erg% = LadeAvDateiMDB%("VdbVerordnungstypen", AvVerordnungen$)
        
    Else
        erg% = True
        If (erg%) Then erg% = LadeAvDatei%("vm-land.dat", AvLaender$)
        If (erg%) Then erg% = LadeAvDatei%("vm-kasse.dat", AvKassen$)
        If (erg%) Then erg% = LadeAvDatei%("vm-vord.dat", AvVerordnungen$)
    End If

    If (erg% = 0) Then
        Call iMsgBox("Dateien des Artikelstamms+V nicht ladbar !")
    End If

    ActBundesland% = 0
    
    If (VERBAND% > 0) Then
        Seek #VERBAND%, 7
        h$ = String(2, 0)
        Get #VERBAND%, , h$
        ActBundesland% = CVI(h$)
    End If


    If (ActBundesland% > UBound(AvLaender$)) Then ActBundesland% = 1
    OrgBundesland% = ActBundesland%
End If

ReDim F7Kassen(0)
ubou% = -1

h$ = String(16, 0)
Get #VERBAND%, 1, h$
For i% = 0 To 9
    Get #VERBAND%, , h$
 
    Call OemToChar(h$, h$)
    
    h2$ = Trim(h$)
    If (h2$ <> "") Then
        ubou% = ubou% + 1
        ReDim Preserve F7Kassen(ubou%)
        With F7Kassen(ubou%)
            .Name = Left$(h2$, 12)
            .kosten = Val(Mid$(h2$, 13))
            .ind = i%
        End With
    End If
Next i%

Close #VERBAND%

Call DefErrPop
End Sub

Function LadeAvDatei%(SrcDatei$, AvFeld$())
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("LadeAvDatei%")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
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
Dim AV_HANDLE%, ret%, ubou%
Dim h$
  
ret% = False

ReDim AvFeld$(0)
ubou% = -1

h$ = para.TaxeLw + ":\taxe\" + SrcDatei$
AV_HANDLE% = FileOpen(h$, "R", "B")

If (AV_HANDLE% > 0) Then
    ret% = True
    
    h$ = String(AV_DATEI_LEN%, 0)
    Do
        Get #AV_HANDLE%, , h$
        If (EOF(AV_HANDLE%)) Then Exit Do
        
        Call OemToChar(h$, h$)
        
        ubou% = ubou% + 1
        ReDim Preserve AvFeld$(ubou%)
        AvFeld$(ubou%) = Trim(h$)
    Loop
    
    Close #AV_HANDLE%
End If

LadeAvDatei% = ret%
    
Call DefErrPop
End Function

Function LadeAvDateiMDB%(SrcTabelle$, AvFeld$())
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("LadeAvDateiMDB%")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
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
Dim AV_HANDLE%, ret%, ubou%
Dim h$
Dim avRec As New ADODB.Recordset
  
ret% = False

ReDim AvFeld$(0)
ubou% = -1

SQLStr$ = "SELECT * FROM " + SrcTabelle + " WHERE (Id<>0) ORDER BY Id"
'Set avRec = AplusVDB.OpenRecordset(SQLStr$)
avRec.Open SQLStr, AplusVDB.ActiveConn
Do
    If (avRec.EOF) Then
        Exit Do
    End If
    
    ret% = True
    ubou% = ubou% + 1
    ReDim Preserve AvFeld$(ubou%)
    h$ = Right$(String(2, "0") + CStr(avRec!Id), 2)
    h$ = h$ + CheckNullStr(avRec!Name)
    AvFeld$(ubou%) = h$
    
    avRec.MoveNext
Loop
avRec.Close

LadeAvDateiMDB% = ret%
    
Call DefErrPop
End Function

Function LadeAvVereinbarungenMDB%(SrcTabelle$, AvFeld$())
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("LadeAvVereinbarungenMDB%")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
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
Dim AV_HANDLE%, ret%, ubou%
Dim h$
Dim avRec As New ADODB.Recordset
  
ret% = False

ReDim AvFeld$(0)
ubou% = -1

SQLStr$ = "SELECT * FROM " + SrcTabelle + " WHERE (VebNr<>0) ORDER BY VebNr"
'Set avRec = AplusVDB.OpenRecordset(SQLStr$)
avRec.Open SQLStr, AplusVDB.ActiveConn
Do
    If (avRec.EOF) Then
        Exit Do
    End If
    
    ret% = True
    ubou% = ubou% + 1
    ReDim Preserve AvFeld$(ubou%)
    h$ = Right$(String(10, "0") + CStr(avRec!VebNr), 10)
    h$ = h$ + Format(avRec!NutzungsBedingungKz, "0")
'    h$ = h$ + CheckNullStr(avRec!Name)
    h$ = h$ + Left(CheckNullStr(avRec!Name) + Space(100), 100)
    h$ = h$ + CheckNullStr(avRec!Stand)
    AvFeld$(ubou%) = h$
    
    avRec.MoveNext
Loop
avRec.Close

LadeAvVereinbarungenMDB% = ret%
    
Call DefErrPop
End Function

Function LadeAvPauschalenMDB%(SrcTabelle$, AvFeld$())
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("LadeAvPauschalenMDB%")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
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
Dim AV_HANDLE%, ret%, ubou%
Dim h$
Dim avRec As New ADODB.Recordset
  
ret% = False

ReDim AvFeld$(0)
ubou% = -1

SQLStr$ = "SELECT * FROM " + SrcTabelle + " WHERE (PauschaleNr<>0) ORDER BY PauschaleNr"
'Set avRec = AplusVDB.OpenRecordset(SQLStr$)
avRec.Open SQLStr, AplusVDB.ActiveConn
Do
    If (avRec.EOF) Then
        Exit Do
    End If
    
    ret% = True
    ubou% = ubou% + 1
    ReDim Preserve AvFeld$(ubou%)
    h$ = Right$(String(10, "0") + CStr(avRec!PauschaleNr), 10)
    h$ = h$ + Format(avRec!NutzungsBedingungKz, "0")
    h$ = h$ + Left("Pauschale: " + CheckNullStr(avRec!Name) + Space(100), 100)
    h$ = h$ + Right$(String(10, "0") + CStr(avRec!VebNr), 10)
    AvFeld$(ubou%) = h$
    
    avRec.MoveNext
Loop
avRec.Close

LadeAvPauschalenMDB% = ret%
    
Call DefErrPop
End Function

Function Verbandmittel%()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("Verbandmittel%")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
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
Dim i%, row%, ind%, erg%, l%
Dim dZuz#, dZuzPreis#
Dim pzn$, SQLStr$, ActGruppe$, h$, ActBedKey$
  
'  int i, GefFlag, hWg;
'  float hZuz;
'  double hPreis;
'  long AltOffs;
'  char h[51], zw2[11], ActBedKey[31], ActGruppe[19];
'  char fStr[11], *p, ch;
'  int fAusw, erg, l;

If (AplusVOk%) Then
    Verbandmittel = VerbandmittelMDB
    Call DefErrPop: Exit Function
End If

Verbandmittel% = False

If (VmFlag% = False) Then
'    Call AltVerbandMittel
    Call DefErrPop: Exit Function
End If

row% = (AktTabIndex% - 3) \ 4
pzn$ = RezArtikel(row%).pzn
If (Val(pzn$) = 0) Then Call DefErrPop: Exit Function

ActWarenzeichen$ = " "
ActMietTage% = -1
ActMietWochen% = -1
ActMietMonate% = -1
ActMietDauer = -1
ActPatientenAlterMonate = -1
ActPatientenAlterJahre = -1
  
ActWert# = -1#

'VmPreisOk% = False

ActMenge% = RezArtikel(row%).Faktor

FabsErrf% = VmPzn.IndexSearch(0, pzn$, FabsRecno&)
If (FabsErrf% <> 0) Then
    Call DefErrPop: Exit Function
End If

Call VmPzn.GetRecord(FabsRecno&)

SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + pzn$
'Set TaxeRec = TaxeDB.OpenRecordset(SQLStr$)
On Error Resume Next
TaxeRec.Close
Err.Clear
On Error GoTo DefErr
TaxeRec.Open SQLStr, taxeAdoDB.ActiveConn
If (TaxeRec.EOF) Then
    Call DefErrPop: Exit Function
End If
  
If (TaxeRec!MwStKZ) Then
    ActMwst% = para.Mwst(1)
Else
    ActMwst% = para.Mwst(2)
End If

ActAep# = TaxeRec!EK / 100
ActAvp# = TaxeRec!vk / 100


h$ = Trim(VmPzn.menge)

If (h$ <> "") Then
    If (Val(h$) = 0) Then h$ = ""
End If

If (h$ = "") Then h$ = TaxeRec!menge

ind% = InStr(h$, "X")
If (ind% > 0) Then
    ActStueck% = Val(Left$(h$, ind% - 1)) * Val(Mid$(h$, ind% + 1))
Else
    ActStueck% = Val(h$)
End If


  
If (VmDebugAktiv%) Then
    Load frmAvDebug
    
    With frmAvDebug.flxAvDebug(0)
        .AddItem "PZN" + vbTab + pzn$
        .AddItem "HM-Nr" + vbTab + VmPzn.HmPositionsNr1
        
        ind% = VmPzn.HilfsmittelKz
        If (ind% = 0) Then
          h$ = "keine Angabe (0)"
        ElseIf (ind% = 1) Then
          h$ = "nicht im HM-Verzeichnis (1)"
        ElseIf (ind% = 2) Then
          h$ = "ist im HM-Verzeichnis (2)"
        ElseIf (ind% = 3) Then
          h$ = "kommt ins HM-Verzeichnis (3)"
        ElseIf (ind% = 4) Then
          h$ = "ist kein HM (4)"
        End If
        .AddItem "HM-Kz  " + vbTab + h$
        .AddItem "GRUPPE" + vbTab + Trim(VmPzn.Gruppe)
        .AddItem ""
    
        .AddItem "EK" + vbTab + Format(ActAep#, "0.00")
        .AddItem "VK" + vbTab + Format(ActAvp#, "0.00")
        .AddItem "ST" + vbTab + Format(ActStueck%, "0")
        .AddItem "MENGE" + vbTab + Format(ActMenge%, "0")
        .AddItem ""

        .AddItem "LAND" + vbTab + AvLaender$(ActBundeslandInd%)
        .AddItem "KASSE" + vbTab + AvKassen$(ActKasseInd%)
        .AddItem "VERORD" + vbTab + AvVerordnungen$(ActVerordnungInd%)
    End With
End If

dZuz# = CalcTaxeZuzahlung

If (ActKasse% < 0) Then
    frmAvAuswahl.Show 1
    Call PaintAvParameter
End If
If (ActKasse% < 0) Then
    Call DefErrPop: Exit Function
End If




ActGruppe$ = VmPzn.Gruppe

Do
    ActBedKey$ = ActGruppe$ + Format(ActBundesland, "00") + Format(ActKasse, "00") + Format(ActVerordnung, "00")
    erg% = SucheInGruppe%(ActBedKey$)
    If (erg%) Then Exit Do

    ActBedKey$ = ActGruppe$ + Format(18, "00") + Format(ActKasse, "00") + Format(ActVerordnung, "00")
    erg% = SucheInGruppe%(ActBedKey$)
    If (erg%) Then Exit Do

    h$ = ActGruppe$
    h$ = Trim(h$)
    l% = Len(h$)
    If (l% <= 2) Then Exit Do
    
'    ActGruppe$ = Left$(Left$(h$, l% - 2) + Space$(18), 18)
    ActGruppe$ = Left$(Left$(h$, l% - 2) + Space$(30), 30)
Loop


If (erg%) Then
    If (VmBed.GenehmigungsPflicht = "J") Then
        Call iMsgBox("Genehmigungspflichtig, Preisbildung überdenken !", 0, "A+V Taxierung")
'        erg% = False
    End If
    If (VmBed.KostenVoranschlag = "J") Then
        Call iMsgBox("Vor der Abgabe ist der Kasse ein Kostenvoranschlag vorzulegen !", 0, "A+V Taxierung")
'        erg% = False
    End If
    If (VmBed.NegKnz = "J") Then
        Call iMsgBox("Artikel der Negativ-Liste !", 0, "A+V Taxierung")
    End If
End If

If (erg%) Then
    erg% = MachBerechnung%(VmBed.KeyBed)
    If (erg% = 0) Then
        RezArtikel(row%).Zuz = 0
    End If
End If

If (VmDebugAktiv%) Then
    For i% = 0 To 2
        With frmAvDebug.flxAvDebug(i%)
            .FillStyle = flexFillRepeat
            If (.Rows = 1) Then .Rows = 2
            If (i% = 0) Then
                .row = 0
            Else
                .row = 1
            End If
            .col = 1
            .RowSel = .Rows - 1
            .ColSel = .Cols - 1
            .CellBackColor = vbWhite
            .FillStyle = flexFillSingle
            
            .col = 0
            .ColSel = .Cols - 1
        End With
    Next i%
    frmAvDebug.Show 1
    erg% = False
End If


If (erg%) Then
    
    RezArtikel(row%).flag = RezArtikel(row%).flag And (REZ_BLINK Xor &HFF)

    If (Left$(TaxeRec!Zuzahlung, 1) = "9") Then
        dZuz# = Act20Preis# * 0.2
        dZuz# = FNX(dZuz#)
        If (ActKasse% = 7) Then
            dZuz# = 0#
        End If
    End If

    RezArtikel(row%).Preis = ActVmPreis#
    If (RezArtikel(row%).IstWg4) Then
        ActVmPreis# = ActVmPreis# * ActMenge%
    End If
    
    If (TaxeRec!ZuzFrei = 0) Then
'        dZuz# = CalcZuzaNeu(RezArtikel(row%).Preis, TaxeRec!HimiVerbrauch)
        dZuzPreis# = ActVmPreis#
        If (ActVmRabatt% = 1) Or (ActVmRabatt% = 2) Then
            dZuzPreis# = dZuzPreis# * 0.94
        ElseIf (ActVmRabatt% = 8) Then
            dZuzPreis# = dZuzPreis# * 0.95
        End If
        dZuz# = CalcZuzaNeu(dZuzPreis#, TaxeRec!HimiVerbrauch, RezArtikel(row%).Faktor)
    End If
    
    RezArtikel(row%).Zuz = dZuz#
    
    RezArtikel(row%).MietDauer = ActMietDauer

    Verbandmittel% = True
    
'    If Not (GebFrei%) And (dZuz# > (RezArtikel(row%).Preis / VmRabattFaktor#)) Then
    If Not (GebFrei%) And (dZuz# > (ActVmPreis# / VmRabattFaktor#)) Then
        RezArtikel(row%).Preis = 0
        RezArtikel(row%).Zuz = 0
    End If

    If (RezArtikel(row%).appli) And (Val(Left$(TaxeRec!Zuzahlung, 1)) <> 9) Then
        RezArtikel(row%).Zuz = 0
    End If
    
    If (ActWarenzeichen$ = "J") Then
        RezArtikel(row%).Warenzeichen = 1
    End If
End If

Call DefErrPop
End Function

Function VerbandmittelMDB%()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("VerbandmittelMDB%")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
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
Dim i%, row%, iRow%, ind%, erg%, l%, iGenehmPflicht%, VebOk%, ZweiteZeile%, ZweiteZeileDa%, Paragraph302%, AbrechnungsverfahrenKz%
Dim dZuz#, dZuzPreis#, VmAddPreis#
Dim VerweisVebNr&
Dim pzn$, SQLStr$, ActGruppe$, h$, h2$, ActBedKey$, hPosNr$
Dim VdbVOKGRec As New ADODB.Recordset
Dim tRec As New ADODB.Recordset
  
VerbandmittelMDB% = False

If (VmFlag% = False) Then
    Call DefErrPop: Exit Function
End If

If (RezArtikel(row).Fortsetzung) Then
    RezArtikel(row%).flag = RezArtikel(row%).flag And (REZ_BLINK Xor &HFF)
    VerbandmittelMDB% = True
    Call DefErrPop: Exit Function
End If

row% = (AktTabIndex% - 3) \ 4
pzn$ = RezArtikel(row%).pzn
If (Left(pzn, 1) = "-") Then
    pzn = Mid(pzn, 2) + "0"
    pzn = MakePzn(pzn)
End If
If (Val(pzn$) = 0) Then Call DefErrPop: Exit Function

If (AutIdemIk > 0) Then
    SQLStr$ = "SELECT * FROM VdbKrankenkasseHinweis WHERE IK = " + CStr(AutIdemIk)
    FabsErrf = AplusVDB.OpenRecordset(VdbHinweiseRec, SQLStr)
    If Not (VdbHinweiseRec.EOF) Then
        VdbHinweiseRec.Close
        h = "Problem: Verordnungen zu Lasten dieses IK dürfen nicht (mehr) beliefert werden!"
        Call iMsgBox(h$, vbCritical, "A+V Taxierung")
        Call DefErrPop: Exit Function
    End If
End If


ActWarenzeichen$ = " "
ActMietZeit = -1
ActMietTage% = -1
ActMietWochen% = -1
ActMietMonate% = -1
ActMietDauer = -1
ActPatientenAlterMonate = -1
ActPatientenAlterJahre = -1
  
ActWert# = -1#

'VmPreisOk% = False

ActMenge% = RezArtikel(row%).Faktor

SQLStr$ = "SELECT * FROM VdbArtikel WHERE PZN = " + pzn$
'Set VdbPznRec = AplusVDB.OpenRecordset(SQLStr$)
FabsErrf = AplusVDB.OpenRecordset(VdbPznRec, SQLStr)
If (VdbPznRec.EOF) Then
    VdbPznRec.Close
    Call DefErrPop: Exit Function
End If

SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + pzn$
'Set TaxeRec = TaxeDB.OpenRecordset(SQLStr$)
On Error Resume Next
TaxeRec.Close
Err.Clear
On Error GoTo DefErr
TaxeRec.Open SQLStr, taxeAdoDB.ActiveConn
If (TaxeRec.EOF) Then
    Call DefErrPop: Exit Function
End If
  
If (TaxeRec!MwStKZ) Then
    ActMwst% = para.Mwst(1)
Else
    ActMwst% = para.Mwst(2)
End If

ActAep# = TaxeRec!EK / 100
ActAvp# = TaxeRec!vk / 100


h$ = Trim(CheckNullStr(VdbPznRec!menge))

If (h$ <> "") Then
    If (Val(h$) = 0) Then h$ = ""
End If

If (h$ = "") Then h$ = TaxeRec!menge

ind% = InStr(h$, "X")
If (ind% > 0) Then
    ActStueck% = Val(Left$(h$, ind% - 1)) * Val(Mid$(h$, ind% + 1))
Else
    ActStueck% = Val(h$)
End If

dZuz# = CalcTaxeZuzahlung

If (RezArtikel(row).VebNr > 0) Or (RezArtikel(row).PauschaleNr > 0) Then
    ActVebNr = RezArtikel(row).VebNr
    ActPauschaleNr = RezArtikel(row).PauschaleNr
End If
If (ActVebNr = 0) And (ActPauschaleNr = 0) Then
    frmAvAuswahl.Show 1
    Call PaintAvParameter
End If
If (ActVebNr = 0) And (ActPauschaleNr = 0) Then
    Call DefErrPop: Exit Function
End If

RezArtikel(row).VebNr = ActVebNr
RezArtikel(row).PauschaleNr = ActPauschaleNr

If (ActPauschaleNr > 0) Then
    SQLStr = "SELECT * FROM VdbPauschalen WHERE PauschaleNr=" + CStr(ActPauschaleNr)
    FabsErrf = AplusVDB.OpenRecordset(VdbPauschaleRec, SQLStr)
    If (VdbPauschaleRec.EOF) Then
        VdbPauschaleRec.Close
        Call DefErrPop: Exit Function
    End If
    ActVebNr = CheckNullLong(VdbPauschaleRec!VebNr)
    
    If (VdbPauschaleRec!NutzungsBedingungKz > 0) Then
        If (InStr(BeigetreteneVereinbarungen, "," + "-" + CStr(ActPauschaleNr) + ",") = 0) Then
            h = "Problem: Pauschale '" + Trim(CheckNullStr(VdbPauschaleRec!Name)) + "' ist BEITRITTSPFLICHTIG!"
            Call iMsgBox(h$, vbCritical, "A+V Taxierung")
            Call DefErrPop: Exit Function
        End If
    End If
End If

If (ActVebNr > 0) Then
    SQLStr = "SELECT * FROM VdbVereinbarungen WHERE VebNr=" + CStr(ActVebNr)
    FabsErrf = AplusVDB.OpenRecordset(VdbVebRec, SQLStr)
    If (VdbVebRec.EOF) Then
        VdbVebRec.Close
        Call DefErrPop: Exit Function
    End If

    If (VdbVebRec!NutzungsBedingungKz > 0) Then
        If (InStr(BeigetreteneVereinbarungen, "," + CStr(ActVebNr) + ",") = 0) Then
            h = "Problem: Vereinbarung '" + Trim(CheckNullStr(VdbVebRec!Name)) + "' ist BEITRITTSPFLICHTIG!"
            Call iMsgBox(h$, vbCritical, "A+V Taxierung")
            Call DefErrPop: Exit Function
        End If
    End If
End If
  
  
If (VmDebugAktiv%) Then
    Load frmAvDebug
    
    With frmAvDebug.flxAvDebug(0)
        .AddItem "PZN" + vbTab + pzn$
        
        h$ = "HM-Nr" + vbTab + Trim(CheckNullStr(VdbPznRec!HmPositionsNr1))
        If (Trim(CheckNullStr(VdbPznRec!HmPositionsNr1)) <> "") Then
            h = h + " / " + Trim(CheckNullStr(VdbPznRec!HmPositionsNr2))
        End If
        .AddItem h$
        
        ind% = CheckNullInt(VdbPznRec!HilfsmittelKz)
        If (ind% = 0) Then
          h$ = "keine Angabe (0)"
        ElseIf (ind% = 1) Then
          h$ = "nicht im HM-Verzeichnis (1)"
        ElseIf (ind% = 2) Then
          h$ = "ist im HM-Verzeichnis (2)"
        ElseIf (ind% = 3) Then
          h$ = "kommt ins HM-Verzeichnis (3)"
        ElseIf (ind% = 4) Then
          h$ = "ist kein HM (4)"
        End If
        .AddItem "HM-Kz  " + vbTab + h$
        .AddItem "GRUPPE" + vbTab + Trim(CheckNullStr(VdbPznRec!GruppeNr))
        .AddItem ""
    
        .AddItem "EK" + vbTab + Format(ActAep#, "0.00")
        .AddItem "VK" + vbTab + Format(ActAvp#, "0.00")
        .AddItem "ST" + vbTab + Format(ActStueck%, "0")
        .AddItem "MENGE" + vbTab + Format(ActMenge%, "0")
        .AddItem ""

        .AddItem "LAND" + vbTab + AvLaender$(ActBundeslandInd%)
'        .AddItem "KASSE" + vbTab + AvKassen$(ActKasseInd%)
        .AddItem "VERORD" + vbTab + AvVerordnungen$(ActVerordnungInd%)
        
        If (ActVebNr > 0) Then
            .AddItem "Vereinbarung" + vbTab + CStr(VdbVebRec!VebNr) + CheckNullStr(VdbVebRec!Name)
            .AddItem "Nutzungsbed." + vbTab + CStr(CheckNullLong(VdbVebRec!NutzungsBedingungKz))
            .AddItem "Stand" + vbTab + CheckNullStr(VdbVebRec!Stand)
            .AddItem "Verweis" + vbTab + CStr(CheckNullLong(VdbVebRec!VerweisVebNr))
        End If
        If (ActPauschaleNr > 0) Then
            .AddItem "Pauschale" + vbTab + CStr(VdbPauschaleRec!PauschaleNr) + CheckNullStr(VdbPauschaleRec!Name)
            .AddItem "Nutzungsbed." + vbTab + CStr(CheckNullLong(VdbPauschaleRec!NutzungsBedingungKz))
'            .AddItem "Zuzahlung" + vbTab + CStr(CheckNullLong(VdbPauschaleRec!ZuzahlungKz))
'            .AddItem "Verweis" + vbTab + CStr(CheckNullLong(VdbPauschaleRec!VebNr))
        End If
    End With
End If

'dZuz# = CalcTaxeZuzahlung
'
'If (ActVebNr <= 0) Then
'    frmAvAuswahl.Show 1
'    Call PaintAvParameter
'End If
'If (ActVebNr <= 0) Then
'    Call DefErrPop: Exit Function
'End If




ActGruppe$ = CheckNullStr(VdbPznRec!GruppeNr)

Do
    erg% = SucheInGruppeMDB%(ActGruppe$, VdbVebRec!VebNr, ActPauschaleNr)
    If (erg%) Then Exit Do

    If (VdbVebRec!VerweisVebNr > 0) Then
        erg% = SucheInGruppeMDB%(ActGruppe$, VdbVebRec!VerweisVebNr, ActPauschaleNr)
        If (erg%) Then
            For i% = 0 To UBound(AvVereinbarungen$)
                h2 = AvVereinbarungen(i)
                If (Val(Left(h2, 10)) = VdbVebRec!VerweisVebNr) Then
                    If (Val(Mid(h2, 11, 1)) > 0) Then
                        If (InStr(BeigetreteneVereinbarungen, "," + CStr(VdbVebRec!VerweisVebNr) + ",") = 0) Then
                            h = "Problem: Vereinbarung '" + Trim(Mid$(h2, 12)) + "' ist BEITRITTSPFLICHTIG!"
                            Call iMsgBox(h$, vbCritical, "A+V Taxierung")
                            erg = 0
                        End If
                    End If
                    Exit For
                End If
            Next i%
        End If
        If (erg%) Then Exit Do
    End If

    h$ = ActGruppe$
    h$ = Trim(h$)
    l% = Len(h$)
    If (l% <= 2) Then Exit Do
    
'    ActGruppe$ = Left$(Left$(h$, l% - 2) + Space$(18), 18)
    ActGruppe$ = Left$(Left$(h$, l% - 2) + Space$(30), 30)
Loop

If (erg) And (ActPauschaleNr > 0) Then
    If (VdbPauschaleRec!GenehmigungsPflichtKz > 0) Then
        If (VdbPauschaleRec!GenehmigungsPflichtKz = 2) Then
            h$ = "Genehmigungspflichtig bei Erstverordnung"
        Else
            h$ = "Genehmigungspflichtig"
        End If
        h$ = h$ + ", Preisbildung überdenken !"
        Call iMsgBox(h$, 0, "A+V Taxierung")
'        erg% = False
    End If
    
    ActVmPreis = CheckNullLong(VdbPauschaleRec!Wert) / 100#
    If (Abs(VdbPauschaleRec!MehrwertSteuerFl) > 0) Then
        ActVmPreis# = ActVmPreis# * (1# + (ActMwst) / 100#)
    End If
    ActVmPreis# = CDbl(Int(Abs(ActVmPreis#) * 100# + 0.501) / 100#)
    
    If (VmDebugAktiv%) Then
        With frmAvDebug.flxAvDebug(2)
            ind = CheckNullLong(VdbPauschaleRec!AbrechnungszeitraumKz)
            h2 = CStr(ind)
            If (ind = 1) Then
                h2 = h2 + " (30 Tage)"
            ElseIf (ind = 2) Then
                h2 = h2 + " (Kalendermonat)"
            End If
            .AddItem "Abrech-Zeitraum" + vbTab + h2$
            .AddItem "Abrechnungs-Nr" + vbTab + CheckNullStr(VdbPauschaleRec!AbrechnungsNr)
            .AddItem "Genehmigung" + vbTab + CStr(VdbPauschaleRec!GenehmigungsPflichtKz)
            
            h2 = ""
            If (CheckNullLong(VdbPauschaleRec!MehrwertSteuerFl) = 0) Then
                h2 = "0%"
            ElseIf (CheckNullLong(VdbPauschaleRec!MwStKZ) = 0) Then
                h2 = CStr(para.Mwst(2)) + "%"
            Else
                h2 = CStr(para.Mwst(1)) + "%"
            End If
            .AddItem "MwStKz" + vbTab + h2
            .AddItem "Zuzahlung" + vbTab + CStr(VdbPauschaleRec!ZuzahlungKz)
            .AddItem "Pauschale" + vbTab + CStr(CheckNullLong(VdbPauschaleRec!Wert) / 100#)
            .AddItem "MwSt" + vbTab + CStr(ActVmPreis)
        End With
    End If

    AvAbrechnungsNr(0) = Trim(CheckNullStr(VdbPauschaleRec!AbrechnungsNr))
        
    If (VmDebugAktiv%) Then
        For i% = 0 To 2
            With frmAvDebug.flxAvDebug(i%)
                .FillStyle = flexFillRepeat
                If (.Rows = 1) Then .Rows = 2
                If (i% = 0) Then
                    .row = 0
                Else
                    .row = 1
                End If
                .col = 1
                .RowSel = .Rows - 1
                .ColSel = .Cols - 1
                .CellBackColor = vbWhite
                .FillStyle = flexFillSingle
                
                .col = 0
                .ColSel = .Cols - 1
            End With
        Next i%
        frmAvDebug.Show 1
        erg% = False
    End If

    RezArtikel(row%).flag = RezArtikel(row%).flag And (REZ_BLINK Xor &HFF)

'    If (Left$(TaxeRec!Zuzahlung, 1) = "9") Then
'        dZuz# = Act20Preis# * 0.2
'        dZuz# = FNX(dZuz#)
'        If (ActKasse% = 7) Then
'            dZuz# = 0#
'        End If
'    End If

    RezArtikel(row%).VdbPauschale = 1
    RezArtikel(row%).Preis = ActVmPreis#
'    If (RezArtikel(row%).IstWg4) Then
'        ActVmPreis# = ActVmPreis# * ActMenge%
'    End If
    
    dZuz = 0
    If (TaxeRec!ZuzFrei = 0) And (VdbPauschaleRec!ZuzahlungKz) Then
        dZuzPreis# = ActVmPreis#
'        dZuz# = CalcZuzaNeu(dZuzPreis#, TaxeRec!HimiVerbrauch, RezArtikel(row%).Faktor)
        RezeptHolenZuz = 0
        dZuz# = CalcZuzaNeu(dZuzPreis#, TaxeRec!HimiVerbrauch, 1)
    End If
    RezArtikel(row%).Zuz = dZuz#
    
    RezArtikel(row%).MietDauer = ActMietDauer

    VerbandmittelMDB% = True
    
'    If Not (GebFrei%) And (dZuz# > (RezArtikel(row%).Preis / VmRabattFaktor#)) Then
    If Not (GebFrei%) And (dZuz# > (ActVmPreis# / VmRabattFaktor#)) Then
        RezArtikel(row%).Preis = 0
        RezArtikel(row%).Zuz = 0
    End If

    If (RezArtikel(row%).appli) And (Val(Left$(TaxeRec!Zuzahlung, 1)) <> 9) Then
        RezArtikel(row%).Zuz = 0
    End If
    
    If (ActWarenzeichen$ = "J") Then
        RezArtikel(row%).Warenzeichen = 1
    End If

    If (AvAbrechnungsNr(0) = "PZN") Then
        RezArtikel(row%).TkkPznDruck = 1
    ElseIf (Val(AvAbrechnungsNr(0)) <> 0) Then
        h = RezArtikel(row%).pzn
        If (Left(h, 1) <> "-") Then
            RezArtikel(row%).pzn = "-" + Left(h, 7)
            RezArtikel(row%).ScreenPzn = AvAbrechnungsNr(0) 'Format(AvAbrechnungsNr(0), String(10, "0"))
        End If
    End If
    
    Call DefErrPop: Exit Function
End If



If (erg%) Then
'    If (VdbBedRec!GenehmigungspflichtFl) Then
    If (VdbBedRec!GenehmigungsPflicht > 0) Then
        If (VdbBedRec!GenehmigungsPflicht = 3) Then
            h$ = "Lieferausschluss für diesen Artikel"
        ElseIf (VdbBedRec!GenehmigungsPflicht = 2) Then
            h$ = "Genehmigungspflichtig bei Erstverordnung"
        Else
            h$ = "Genehmigungspflichtig"
        End If
        iGenehmPflicht = VdbBedRec!Genehmpflicht_Zusatz
        If (iGenehmPflicht > 0) Then
            h$ = h$ + ", betrifft nur "
            If (iGenehmPflicht = 1) Then
                h$ = h$ + "Pflegeheimbewohner"
            ElseIf (iGenehmPflicht = 2) Then
                h$ = h$ + "BVG-Rezepte"
            Else
                h$ = h$ + "Pflegeheimbewohner und BVG-Rezepte"
            End If
        End If
        h$ = h$ + ", Preisbildung überdenken !"
        Call iMsgBox(h$, 0, "A+V Taxierung")
'        erg% = False
    End If
        
    If (VdbBedRec!KostenvoranschlagFl) Then
        Call iMsgBox("Vor der Abgabe ist der Kasse ein Kostenvoranschlag vorzulegen !", 0, "A+V Taxierung")
'        erg% = False
    End If
    
    If (VdbBedRec!NegativlisteFl) Then
        Call iMsgBox("Artikel der Negativ-Liste !", 0, "A+V Taxierung")
    End If
End If

For i% = 0 To 2
    AvAbrechnungsNr(i) = ""
    AvHmPositionsNr(i) = ""
Next i

If (erg%) Then
    RezArtikel(row).HmAbrechnungsKz = HmAbrechnungsKz
    RezArtikel(row).LEGS = LEGS
    
    erg% = MachBerechnungMDB%(VdbBedRec!bednr)
    If (erg% = 0) Then
        RezArtikel(row%).Zuz = 0
    End If
End If

If (VmDebugAktiv%) Then
    For i% = 0 To 2
        With frmAvDebug.flxAvDebug(i%)
            .FillStyle = flexFillRepeat
            If (.Rows = 1) Then .Rows = 2
            If (i% = 0) Then
                .row = 0
            Else
                .row = 1
            End If
            .col = 1
            .RowSel = .Rows - 1
            .ColSel = .Cols - 1
            .CellBackColor = vbWhite
            .FillStyle = flexFillSingle
            
            .col = 0
            .ColSel = .Cols - 1
        End With
    Next i%
    frmAvDebug.Show 1
    erg% = False
End If


If (erg%) Then
    
    VmAddPreis = 0
    If (VdbBedRec!WertErmittlungKz = 2) Then
        SQLStr = "SELECT * FROM VdbVOKG WHERE VdbVOKG.VebNr=" + CStr(ActVebNr)
        SQLStr = SQLStr + " AND VdbVOKG.VOrdTypNr=" + CStr(ActVerordnung)
        SQLStr = SQLStr + " AND VdbVOKG.Folgeausdruck='ABR302'"
        SQLStr = SQLStr + " ORDER BY VdbVOKG.Key_VOKG"
        'Set VdbVOKGRec = AplusVDB.OpenRecordset(SQLStr)
        VdbVOKGRec.Open SQLStr, AplusVDB.ActiveConn
        If (VdbVOKGRec.EOF) Then
            VerweisVebNr = 0
            SQLStr = "SELECT * FROM VdbVereinbarungen WHERE VebNr=" + CStr(ActVebNr)
            FabsErrf = AplusVDB.OpenRecordset(tRec, SQLStr)
            If Not (tRec.EOF) Then
                VerweisVebNr = CheckNullLong(tRec!VerweisVebNr)
            End If
            tRec.Close
        
            If (VerweisVebNr > 0) Then
                VdbVOKGRec.Close
        
                SQLStr = "SELECT * FROM VdbVOKG WHERE VdbVOKG.VebNr=" + CStr(VerweisVebNr)
                SQLStr = SQLStr + " AND VdbVOKG.VOrdTypNr=" + CStr(ActVerordnung)
                SQLStr = SQLStr + " AND VdbVOKG.Folgeausdruck='ABR302'"
                SQLStr = SQLStr + " ORDER BY VdbVOKG.Key_VOKG"
                'Set VdbVOKGRec = AplusVDB.OpenRecordset(SQLStr)
                VdbVOKGRec.Open SQLStr, AplusVDB.ActiveConn
            End If
        End If
        Paragraph302 = Not (VdbVOKGRec.EOF)
        
        'neu in 4.0.62
        AbrechnungsverfahrenKz = 0
        On Error Resume Next
        AbrechnungsverfahrenKz = VdbBedRec!AbrechnungsverfahrenKz
        On Error GoTo DefErr
        Paragraph302 = (AbrechnungsverfahrenKz = 2)
        
        ZweiteZeile = 0
        For i = 0 To 2
            If (AvHmPositionsNr(i) = "") Then
                ActVmPreis = ActPreis(i)
                If (Paragraph302) Then
                    h = RezArtikel(row%).pzn
                    RezArtikel(row%).pzn = "-" + Left(h, 7)
                    RezArtikel(row%).ScreenPzn = CheckNullStr(VdbPznRec!HmPositionsNr1)
                End If
                If (ZweiteZeile) Then
                    Exit For
                End If
                ZweiteZeile = True
'                Exit For
            ElseIf (AvHmPositionsNr(i) = VdbPznRec!HmPositionsNr2) Then
                ZweiteZeileDa = 0
                If (row < AnzRezeptArtikel - 1) Then
                    If (Trim(RezArtikel(row + 1).text) = "Zusatzkomponente") Then
                        ZweiteZeileDa = True
                    End If
                End If
                If (ZweiteZeileDa = 0) Then
                    AnzRezeptArtikel% = AnzRezeptArtikel% + 1
                End If
                iRow% = AnzRezeptArtikel% - 1
                Call InitRezeptArtikel(iRow%)
                RezArtikel(iRow%).text = "Zusatzkomponente"
'                RezArtikel(iRow%).pzn = "-ZUSG_HM"
                If (Paragraph302) Then
                    RezArtikel(iRow%).pzn = "-9999999"
                    RezArtikel(iRow%).ScreenPzn = AvHmPositionsNr(i)
                    ZusatzKomponentenNr = AvHmPositionsNr(i)
                    ZusatzKomponentenBasisNr = RezArtikel(row).ScreenPzn
                Else
                    RezArtikel(iRow%).pzn = "        "
                    RezArtikel(iRow%).flag = RezArtikel(iRow%).flag Or (REZ_BLINK)
                End If
                RezArtikel(iRow%).Faktor = RezArtikel(row%).Faktor '1
                RezArtikel(iRow%).Preis = ActPreis(i)
                VmAddPreis = ActPreis(i)
                
'                dZuzPreis# = ActPreis(i)
'                If (ActVmRabatt% = 1) Or (ActVmRabatt% = 2) Then
'                    dZuzPreis# = dZuzPreis# * 0.94
'                ElseIf (ActVmRabatt% = 8) Then
'                    dZuzPreis# = dZuzPreis# * 0.95
'                End If
'                dZuz# = CalcZuzaNeu(dZuzPreis#, 0, 1)
'                RezArtikel(iRow%).Zuz = dZuz#
                RezArtikel(iRow%).Zuz = 0
                
                VerbandmittelMDB% = True
                
                If (ActWarenzeichen$ = "J") Then
                    RezArtikel(iRow%).Warenzeichen = 1
                End If
                
                Call PaintArtikel(iRow)
                
                If (ZweiteZeile) Then
                    Exit For
                End If
                ZweiteZeile = True
            End If
        Next i
    End If
    
    AbrechnungsverfahrenKz = 0
    On Error Resume Next
    AbrechnungsverfahrenKz = VdbBedRec!AbrechnungsverfahrenKz
    On Error GoTo DefErr
    If (AbrechnungsverfahrenKz = 2) Then
        If (AvAbrechnungsNr(0) = "PZN") Then
        Else
            RezeptMitHilfsmittelNr% = 1
    
            hPosNr$ = Trim(Left$(CheckNullStr(VdbPznRec!HmPositionsNr1), 10))
            If (hPosNr$ <> "") Then
                RezArtikel(row).pzn = "-" + Left(pzn, 7)
                RezArtikel(row).ScreenPzn = hPosNr
            End If
        End If
    End If
    
    RezArtikel(row%).flag = RezArtikel(row%).flag And (REZ_BLINK Xor &HFF)

    If (Left$(TaxeRec!Zuzahlung, 1) = "9") Then
        dZuz# = Act20Preis# * 0.2
        dZuz# = FNX(dZuz#)
        If (ActKasse% = 7) Then
            dZuz# = 0#
        End If
    End If

    If (AbrechnungsVerfahren_2_Aktiv) And (AbrechnungsverfahrenKz = 2) Then
        RezArtikel(row%).VdbPauschale = 1
        RezArtikel(row%).Preis = ActVmPreis#
        If (RezArtikel(row%).IstWg4 = 0) Then
            ActVmPreis = ActVmPreis / ActMenge
        End If
    Else
        RezArtikel(row%).Preis = ActVmPreis#
        If (RezArtikel(row%).IstWg4) Then
            ActVmPreis# = ActVmPreis# * ActMenge%
        End If
    End If
    
    dZuz = 0
    If (TaxeRec!ZuzFrei = 0) And (VdbBedRec!ZuzahlungFl) Then
'        dZuz# = CalcZuzaNeu(RezArtikel(row%).Preis, TaxeRec!HimiVerbrauch)
        dZuzPreis# = ActVmPreis# + VmAddPreis
        If (ActVmRabatt% = 1) Or (ActVmRabatt% = 2) Then
            dZuzPreis# = dZuzPreis# * 0.94
        ElseIf (ActVmRabatt% = 8) Then
            dZuzPreis# = dZuzPreis# * 0.95
        End If
        dZuz# = CalcZuzaNeu(dZuzPreis#, TaxeRec!HimiVerbrauch, RezArtikel(row%).Faktor)
    
        If (AbrechnungsVerfahren_2_Aktiv) And (AbrechnungsverfahrenKz = 2) Then
            If (RezArtikel(row%).IstWg4 = 0) Then
                dZuz = dZuz * ActMenge
            End If
        End If
    End If
    
    RezArtikel(row%).Zuz = dZuz#
    
    RezArtikel(row%).MietDauer = ActMietDauer

    VerbandmittelMDB% = True
    
'    If Not (GebFrei%) And (dZuz# > (RezArtikel(row%).Preis / VmRabattFaktor#)) Then
    If Not (GebFrei%) And (dZuz# > (ActVmPreis# / VmRabattFaktor#)) Then
        RezArtikel(row%).Preis = 0
        RezArtikel(row%).Zuz = 0
    End If

    If (RezArtikel(row%).appli) And (Val(Left$(TaxeRec!Zuzahlung, 1)) <> 9) Then
        RezArtikel(row%).Zuz = 0
    End If
    
    If (ActWarenzeichen$ = "J") Then
        RezArtikel(row%).Warenzeichen = 1
    End If

    If (AvAbrechnungsNr(0) = "PZN") Then
        RezArtikel(row%).TkkPznDruck = 1
    ElseIf (Val(AvAbrechnungsNr(0)) <> 0) Then
        h = RezArtikel(row%).pzn
        If (Left(h, 1) <> "-") Then
            RezArtikel(row%).pzn = "-" + Left(h, 7)
            RezArtikel(row%).ScreenPzn = AvAbrechnungsNr(0) 'Format(AvAbrechnungsNr(0), String(10, "0"))
        End If
    End If
End If

Call DefErrPop
End Function

Function CheckVerordnungMDB%()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("CheckVerordnungMDB%")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
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
Dim i%, j%, k%, row%, ind%, ind2%, erg%, l%, iKeyPkl%, AnzKeyPkl%(1), iTabIndex%
Dim KeyVOKG&, VerweisVebNr&
Dim dZuz#, dZuzPreis#, dWert#(1)
Dim pzn$, SQLStr$, ActGruppe$, ActBedingung$, h$, h2$, ActBedKey$, PznPkl$(10), PznTreffer$(10), ActOperator$, VdbZeilen$, OrgVdbZeilen$
Dim MsgStr$, hPosNr$, SollKeyPkl$

Dim VdbPKLRec As New ADODB.Recordset
Dim VdbVAPRec As New ADODB.Recordset
Dim VdbVOKGRec As New ADODB.Recordset
Dim VdbVOKZRec As New ADODB.Recordset
'Dim tRec As Recordset
Dim tRec As New ADODB.Recordset
  
'  int i, GefFlag, hWg;
'  float hZuz;
'  double hPreis;
'  long AltOffs;
'  char h[51], zw2[11], ActBedKey[31], ActGruppe[19];
'  char fStr[11], *p, ch;
'  int fAusw, erg, l;

If (BtmRezept) And (BtmFaktorManuell = 0) Then
    j = 0
    ind = -1
    For i% = 0 To (AnzRezeptArtikel% - 1)
        If (Val(RezArtikel(i%).pzn) = 2567001) Then
            ind = i
        ElseIf (RezArtikel(i).Faktor > 0) Then
            SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + RezArtikel(i%).pzn
            'Set tRec = TaxeDB.OpenRecordset(SQLStr$)
            On Error Resume Next
            tRec.Close
            Err.Clear
            On Error GoTo DefErr
            tRec.Open SQLStr, taxeAdoDB.ActiveConn
            If (tRec.EOF) Then
                If (CheckBtmSonderPzn%(RezArtikel(i%).pzn) >= 0) Then
                    j = j + 1
                End If
            Else
                If (tRec!BtmKz) Then
                    j = j + 1
                End If
            End If
        End If
    Next i%
    If (ind >= 0) Then
        RezArtikel(ind).Faktor = j
        
        iTabIndex% = AktTabIndex%
        Call PaintArtikel(ind%)
        Call TabIndexChange(iTabIndex%, False)
    End If
End If


If (AplusVOk% = 0) Then
    CheckVerordnungMDB = 0
    Call DefErrPop: Exit Function
End If

'If (ActKasse = 14) Or (ActKasse = 31) Or (ActKasse = 36) Or (ActKasse = 41) Or (ActKasse = 64) Then
'    RezeptMitHilfsmittelNr% = 1
'    CheckVerordnungMDB = 0
'    Call DefErrPop: Exit Function
'End If


''''''''''
SQLStr = "SELECT * FROM VdbVOKG INNER JOIN VdbVOKZ ON VdbVOKZ.Key_VOKG=VdbVOKG.Key_VOKG"
SQLStr = SQLStr + " WHERE VdbVOKG.VebNr=" + CStr(ActVebNr)
SQLStr = SQLStr + " AND VdbVOKG.VOrdTypNr=" + CStr(ActVerordnung)
SQLStr = SQLStr + " AND VdbVOKG.Folgeausdruck='ABR302'"
SQLStr = SQLStr + " ORDER BY VdbVOKG.Key_VOKG"
'Set VdbVOKGRec = AplusVDB.OpenRecordset(SQLStr)
VdbVOKGRec.Open SQLStr, AplusVDB.ActiveConn
If (VdbVOKGRec.EOF) Then
    VerweisVebNr = 0
    SQLStr = "SELECT * FROM VdbVereinbarungen WHERE VebNr=" + CStr(ActVebNr)
    FabsErrf = AplusVDB.OpenRecordset(tRec, SQLStr)
    If Not (tRec.EOF) Then
        VerweisVebNr = CheckNullLong(tRec!VerweisVebNr)
    End If
    tRec.Close

    If (VerweisVebNr > 0) Then
        VdbVOKGRec.Close

        SQLStr = "SELECT * FROM VdbVOKG INNER JOIN VdbVOKZ ON VdbVOKZ.Key_VOKG=VdbVOKG.Key_VOKG"
        SQLStr = SQLStr + " WHERE VdbVOKG.VebNr=" + CStr(VerweisVebNr)
        SQLStr = SQLStr + " AND VdbVOKG.VOrdTypNr=" + CStr(ActVerordnung)
        SQLStr = SQLStr + " AND VdbVOKG.Folgeausdruck='ABR302'"
        SQLStr = SQLStr + " ORDER BY VdbVOKG.Key_VOKG"
        'Set VdbVOKGRec = AplusVDB.OpenRecordset(SQLStr)
        VdbVOKGRec.Open SQLStr, AplusVDB.ActiveConn
    End If
End If
If Not (VdbVOKGRec.EOF) Then
    SollKeyPkl = ","
    Do
        If (VdbVOKGRec.EOF) Then
            Exit Do
        End If
        
        iKeyPkl = CheckNullInt(VdbVOKGRec!Key_PKL)
        If (iKeyPkl > 0) Then
            SollKeyPkl = SollKeyPkl + CStr(iKeyPkl) + ","
        End If
        
        VdbVOKGRec.MoveNext
    Loop

    For i% = 0 To (AnzRezeptArtikel% - 1)
        If (RezArtikel(i%).TkkPznDruck) Then
        Else
            h = RezArtikel(i).pzn
            If (h = "-9999999") Then
                RezeptMitHilfsmittelNr% = -1
            ElseIf (Left(h, 1) <> "-") Then
                SQLStr = "SELECT * FROM VdbVAP"
                SQLStr = SQLStr + " WHERE PZN=" + h
                SQLStr = SQLStr + " ORDER BY Key_PKL"
                VdbVAPRec.Open SQLStr, AplusVDB.ActiveConn
                If Not (VdbVAPRec.EOF) Then
                    Do
                        If (VdbVAPRec.EOF) Then
                            Exit Do
                        End If
                        
                        If (InStr(SollKeyPkl, "," + CStr(CheckNullInt(VdbVAPRec!Key_PKL)) + ",") > 0) Then
                            RezeptMitHilfsmittelNr% = 1
                
                            SQLStr$ = "SELECT * FROM VdbArtikel WHERE PZN = " + h
                            FabsErrf = AplusVDB.OpenRecordset(VdbPznRec2, SQLStr)
                            If Not (VdbPznRec2.EOF) Then
                                hPosNr$ = Trim(Left$(CheckNullStr(VdbPznRec2!HmPositionsNr1), 10))
                                If (hPosNr$ <> "") Then
                                    RezArtikel(i).pzn = "-" + Left(h, 7)
                                    RezArtikel(i).ScreenPzn = hPosNr
                                End If
                            End If
                            VdbPznRec2.Close
                            
                            Exit Do
                        End If
                        
                        VdbVAPRec.MoveNext
                    Loop
                End If
                VdbVAPRec.Close
            End If
        End If
    Next i
End If
VdbVOKGRec.Close
'''''''''


SQLStr = "SELECT * FROM VdbVOKG"
'SQLStr = SQLStr + " WHERE KostentrNr=" + CStr(ActKasse)
'SQLStr = SQLStr + " AND LandNr=" + CStr(ActBundesland)
SQLStr = SQLStr + " WHERE VebNr=" + CStr(ActVebNr)
SQLStr = SQLStr + " AND VOrdTypNr=" + CStr(ActVerordnung)
SQLStr = SQLStr + " AND Folgeausdruck<>'ABR302'"
SQLStr = SQLStr + " ORDER BY Key_VOKG"
'Set VdbVOKGRec = AplusVDB.OpenRecordset(SQLStr)
VdbVOKGRec.Open SQLStr, AplusVDB.ActiveConn
If (VdbVOKGRec.EOF) Then
    VerweisVebNr = 0
    SQLStr = "SELECT * FROM VdbVereinbarungen WHERE VebNr=" + CStr(ActVebNr)
    FabsErrf = AplusVDB.OpenRecordset(tRec, SQLStr)
    If Not (tRec.EOF) Then
        VerweisVebNr = CheckNullLong(tRec!VerweisVebNr)
    End If
    tRec.Close

    If (VerweisVebNr > 0) Then
        VdbVOKGRec.Close
    
        SQLStr = "SELECT * FROM VdbVOKG"
    '    SQLStr = SQLStr + " WHERE KostentrNr=" + CStr(ActKasse)
    '    SQLStr = SQLStr + " AND LandNr=" + CStr(18)
        SQLStr = SQLStr + " WHERE VebNr=" + CStr(VerweisVebNr)
        SQLStr = SQLStr + " AND VOrdTypNr=" + CStr(ActVerordnung)
        SQLStr = SQLStr + " AND Folgeausdruck<>'ABR302'"
        SQLStr = SQLStr + " ORDER BY Key_VOKG"
    '    Set VdbVOKGRec = AplusVDB.OpenRecordset(SQLStr)
        VdbVOKGRec.Open SQLStr, AplusVDB.ActiveConn
    End If
End If

If (VdbVOKGRec.EOF) Then
    CheckVerordnungMDB = 0
    Call DefErrPop: Exit Function
'ElseIf (CheckNullStr(VdbVOKGRec!Folgeausdruck) = "ABR302") Then
'    RezeptMitHilfsmittelNr% = 1
'    CheckVerordnungMDB = 0
'
'    For i% = 0 To (AnzRezeptArtikel% - 1)
'        h = RezArtikel(i).pzn
'        If (h = "-9999999") Then
'            RezeptMitHilfsmittelNr% = -1
'        ElseIf (Left(h, 1) <> "-") Then
'            SQLStr$ = "SELECT * FROM VdbArtikel WHERE PZN = " + h
'            FabsErrf = AplusVDB.OpenRecordset(VdbPznRec2, SQLStr)
'            If Not (VdbPznRec2.EOF) Then
'                hPosNr$ = Trim(Left$(CheckNullStr(VdbPznRec2!HmPositionsNr1), 10))
'                If (hPosNr$ <> "") Then
'                    RezArtikel(i).pzn = "-" + Left(h, 7)
'                    RezArtikel(i).ScreenPzn = hPosNr
'                End If
'            End If
'            VdbPznRec2.Close
'        End If
'    Next i
'
'    Call DefErrPop: Exit Function
End If

OrgVdbZeilen = ""
For i% = 0 To (AnzRezeptArtikel% - 1)
    pzn$ = RezArtikel(i).pzn
    If (Left(pzn, 1) = "-") Then
        pzn = Mid(pzn, 2) + "0"
        pzn = MakePzn(pzn)
    End If
    
    h$ = ""
    If (Val(pzn) > 0) Then
        SQLStr$ = "SELECT * FROM VdbArtikel WHERE PZN = " + pzn ' RezArtikel(i%).pzn
    '    Set VdbPznRec = AplusVDB.OpenRecordset(SQLStr$)
        FabsErrf = AplusVDB.OpenRecordset(VdbPznRec2, SQLStr)
        If Not (VdbPznRec2.EOF) Then
            OrgVdbZeilen = OrgVdbZeilen + "," + CStr(i) + ","
        End If
        VdbPznRec2.Close
    
        SQLStr = "SELECT * FROM VdbVAP"
        SQLStr = SQLStr + " WHERE PZN=" + pzn   'RezArtikel(i%).pzn
        SQLStr = SQLStr + " ORDER BY Key_PKL"
    '    Set VdbVAPRec = AplusVDB.OpenRecordset(SQLStr)
        VdbVAPRec.Open SQLStr, AplusVDB.ActiveConn
        If Not (VdbVAPRec.EOF) Then
            h$ = ","
            Do
                If (VdbVAPRec.EOF) Then
                    Exit Do
                End If
                
                h$ = h$ + CStr(CheckNullInt(VdbVAPRec!Key_PKL)) + ","
                
                VdbVAPRec.MoveNext
            Loop
        End If
        VdbVAPRec.Close
    End If
    
    PznPkl(i) = h$
    PznTreffer(i) = ""
Next i

Do
    If (VdbVOKGRec.EOF) Then
        Exit Do
    End If
    
    VdbZeilen = OrgVdbZeilen
    
    SQLStr = "SELECT * FROM VdbVOKZ"
    SQLStr = SQLStr + " WHERE Key_VOKG=" + CStr(VdbVOKGRec!Key_VOKG)
    SQLStr = SQLStr + " ORDER BY Zaehler"
'    Set VdbVOKZRec = AplusVDB.OpenRecordset(SQLStr)
    FabsErrf = AplusVDB.OpenRecordset(VdbVOKZRec, SQLStr)
    If Not (VdbVOKZRec.EOF) Then
        VdbZeilen = ""
        Do
            If (VdbVOKZRec.EOF) Then
                Exit Do
            End If
            
            iKeyPkl = CheckNullInt(VdbVOKZRec!Key_PKL)
            If (iKeyPkl > 0) Then
                AnzKeyPkl(0) = 0
                For i% = 0 To (AnzRezeptArtikel% - 1)
                    If (InStr(PznPkl(i), "," + CStr(iKeyPkl) + ",") > 0) Then
                        AnzKeyPkl(0) = AnzKeyPkl(0) + 1
                    End If
                Next i
            
                AnzKeyPkl(1) = Val(CheckNullStr(VdbVOKZRec!Anz_Zeilen_Zahl))
                ActOperator$ = CheckNullStr(VdbVOKZRec!Anz_Zeilen_Operator)
                
                If (((ActOperator = "<") And (AnzKeyPkl(0) < AnzKeyPkl(1))) Or _
                    ((ActOperator = "=") And (AnzKeyPkl(0) = AnzKeyPkl(1))) Or _
                    ((ActOperator = ">") And (AnzKeyPkl(0) > AnzKeyPkl(1))) Or _
                    ((ActOperator = "<=") And (AnzKeyPkl(0) <= AnzKeyPkl(1))) Or _
                    ((ActOperator = "<>") And (AnzKeyPkl(0) <> AnzKeyPkl(1))) Or _
                    ((ActOperator = ">=") And (AnzKeyPkl(0) >= AnzKeyPkl(1))) _
                   ) Then
                    For i% = 0 To (AnzRezeptArtikel% - 1)
                        If (InStr(PznPkl(i), "," + CStr(iKeyPkl) + ",") > 0) Then
                            VdbZeilen = VdbZeilen + "," + CStr(i) + ","
                        End If
                    Next i
                Else
                    CheckVerordnungMDB = 0
                    Call DefErrPop: Exit Function
                End If
            End If
            
            VdbVOKZRec.MoveNext
        Loop
    End If
    
    ActBedingung$ = Trim(CheckNullStr(VdbVOKGRec!Bedingungsausdruck))
    If (UCase(Left(ActBedingung$, 8)) = "VO_SUMME") Then
        dWert(0) = 0
        For i% = 0 To (AnzRezeptArtikel% - 1)
            If (InStr(VdbZeilen, "," + CStr(i) + ",") > 0) Then
                dWert(0) = dWert(0) + (RezArtikel(i%).Preis * RezArtikel(i%).Faktor)
            End If
        Next i
        
        ActBedingung$ = Mid$(ActBedingung$, 9)
        ActOperator = Left$(ActBedingung, 1)
        ActBedingung$ = Mid$(ActBedingung$, 2)
        If (Left$(ActBedingung, 1) = "=") Then
            ActOperator = ActOperator + "="
            ActBedingung$ = Mid$(ActBedingung$, 2)
        End If
        dWert(1) = xVal(ActBedingung)
        
        If (((ActOperator = "<") And (dWert(0) < dWert(1))) Or _
            ((ActOperator = "=") And (dWert(0) = dWert(1))) Or _
            ((ActOperator = ">") And (dWert(0) > dWert(1))) Or _
            ((ActOperator = "<=") And (dWert(0) <= dWert(1))) Or _
            ((ActOperator = "<>") And (dWert(0) <> dWert(1))) Or _
            ((ActOperator = ">=") And (dWert(0) >= dWert(1))) _
           ) Then
        Else
            CheckVerordnungMDB = 0
            Call DefErrPop: Exit Function
        End If
    End If
    
    For i% = 0 To (AnzRezeptArtikel% - 1)
        If (InStr(VdbZeilen, "," + CStr(i) + ",") > 0) Then
            PznTreffer(i) = PznTreffer(i) + CStr(VdbVOKGRec!Key_VOKG) + ","
        End If
    Next i

'    If (VmDebugAktiv%) Then
'        h$ = "SATZ"
'        h$ = h$ + vbTab + VdbBedRec!GruppeNr + Right$("00" + CStr(VdbBedRec!LandNr), 2) + Right$("00" + CStr(VdbBedRec!KostentrNr), 2) + Right$("00" + CStr(VdbBedRec!VOrdTypNr), 2)
'
'        h$ = h$ + vbTab + CheckNullStr(VdbBedRec!Warenzeichen)
'        h$ = h$ + vbTab + CStr(Abs(VdbBedRec!RundungstypFl))
'
'        h2$ = "N"
'        If (VdbBedRec!ZuzahlungFl) Then
'            h2$ = "J"
'        End If
'        h$ = h$ + vbTab + h2$
'
'        h2$ = "N"
'        If (VdbBedRec!NegativlisteFl) Then
'            h2$ = "J"
'        End If
'        h$ = h$ + vbTab + h2$
'
'        h2$ = "0"
'        If (VdbBedRec!WertErmittlungFl) Then
'            h2$ = "1"
'        End If
'        h$ = h$ + vbTab + h2$
'
'        h2$ = "N"
'        If (VdbBedRec!KostenvoranschlagFl) Then
'            h2$ = "J"
'        End If
'        h$ = h$ + vbTab + h2$
'
'        h2$ = "N"
'        If (VdbBedRec!GenehmigungspflichtFl) Then
'            h2$ = "J"
'        End If
'        h$ = h$ + vbTab + h2$
'
'        With frmAvDebug.flxAvDebug(1)
'            .AddItem h$
'            .AddItem "BEDING" + vbTab + "(" + CStr(VdbBedRec!BedNr) + ")" + ActBedingung$
'        End With
'    End If

    VdbVOKGRec.MoveNext
Loop

erg = 0
For i% = 0 To (AnzRezeptArtikel% - 1)
    If (PznTreffer(i) <> "") Then
        erg = True
        Exit For
    End If
Next i

If (erg = 0) Then
    CheckVerordnungMDB = 0
    Call DefErrPop: Exit Function
End If


InhaltsstoffeModus = 1
Load frmInhaltsStoffe
frmInhaltsStoffe.Caption = "Zeilenübergreifende Regelungen"

With frmInhaltsStoffe.flxInhaltsStoffe
    .Rows = .FixedRows
    For i% = 0 To (AnzRezeptArtikel% - 1)
        If (PznTreffer(i) <> "") Then
            k = 0
            .AddItem RezArtikel(i).text
            
            h = PznTreffer(i)
            Do
                ind% = InStr(h, ",")
                If (ind > 0) Then
                    KeyVOKG& = Val(Left(h, ind - 1))
    
                    SQLStr = "SELECT * FROM VdbVOKG"
                    SQLStr = SQLStr + " WHERE Key_VOKG=" + CStr(KeyVOKG)
'                    Set VdbVOKGRec = AplusVDB.OpenRecordset(SQLStr)
                    FabsErrf = AplusVDB.OpenRecordset(VdbVOKGRec, SQLStr)
                    If Not (VdbVOKGRec.EOF) Then
                        h2$ = CheckNullStr(VdbVOKGRec!Folgeausdruck)
                        If (h2 = "G") Then
                            MsgStr = "Genehmigungspflicht"
                        ElseIf (h2 = "K") Then
                            MsgStr = "Kostenvoranschlag"
                        End If
                        If (k > 0) Then
                            .AddItem ""
                        End If
                        .TextMatrix(.Rows - 1, 1) = MsgStr
                        .row = .Rows - 1
                        .FillStyle = flexFillRepeat
                        .col = 0
                        .ColSel = .Cols - 1
                        .CellFontBold = True
                        .FillStyle = flexFillSingle
                        k = k + 1
        
                        ActBedingung$ = Trim(CheckNullStr(VdbVOKGRec!Bedingungsausdruck))
                        If (ActBedingung <> "") Then
                            h2 = " bei "
                            If (UCase(Left(ActBedingung$, 8)) = "VO_SUMME") Then
                                h2 = h2 + "Verordnungssumme" + Mid$(ActBedingung, 9)
                            Else
                                h2 = h2 + ActBedingung
                            End If
                            .AddItem vbTab + h2
                        End If
                    
                        SQLStr = "SELECT * FROM VdbVOKZ"
                        SQLStr = SQLStr + " WHERE Key_VOKG=" + CStr(VdbVOKGRec!Key_VOKG)
                        SQLStr = SQLStr + " ORDER BY Zaehler"
'                        Set VdbVOKZRec = AplusVDB.OpenRecordset(SQLStr)
                        FabsErrf = AplusVDB.OpenRecordset(VdbVOKZRec, SQLStr)
                        If Not (VdbVOKZRec.EOF) Then
                            j = 0
                            h2 = ""
                            Do
                                If (VdbVOKZRec.EOF) Then
                                    Exit Do
                                End If
                                
                                h2 = ""
                                If (j = 0) Then
                                    h2 = h2 + "wenn "
                                Else
                                    h2 = h2 + "und "
                                End If
                                ActOperator = CheckNullStr(VdbVOKZRec!Anz_Zeilen_Operator)
                                AnzKeyPkl(0) = Val(CheckNullStr(VdbVOKZRec!Anz_Zeilen_Zahl))
                                h2 = h2 + ActOperator
                                h2 = h2 + " " + CStr(AnzKeyPkl(0))
                                h2 = h2 + " Zeilen mit Produktklasse "
                                
                                iKeyPkl = CheckNullInt(VdbVOKZRec!Key_PKL)
                                SQLStr = "SELECT * FROM VdbPKL"
                                SQLStr = SQLStr + " WHERE Key_PKL=" + CStr(iKeyPkl)
'                                Set VdbPKLRec = AplusVDB.OpenRecordset(SQLStr)
                                FabsErrf = AplusVDB.OpenRecordset(VdbPKLRec, SQLStr)
                                If (VdbPKLRec.EOF) Then
                                    h2 = h2 + "(" + CStr(iKeyPkl) + ")"
                                Else
                                    h2 = h2 + CheckNullStr(VdbPKLRec!Name)
                                End If
                                .AddItem vbTab + h2
                                
                                j = j + 1
                                VdbVOKZRec.MoveNext
                            Loop
                            If (j = 1) Then
                                ind2 = InStr(h2, "> 0")
                                If (ind2 = 0) Then
                                    ind2 = InStr(h2, ">= 1")
                                End If
                                If (ind2 > 0) Then
                                    h2 = "(Produktklasse "
                                    If (VdbPKLRec.EOF) Then
                                        h2 = h2 + CStr(iKeyPkl)
                                    Else
                                        h2 = h2 + CheckNullStr(VdbPKLRec!Name)
                                    End If
                                    h2 = h2 + ")"
                                    .Rows = .Rows - 1
                                    .AddItem vbTab + h2
                                    End If
                            End If
                        End If
                    End If
                    
                    h = Mid$(h, ind + 1)
                Else
                    Exit Do
                End If
            Loop
        End If
    Next i
    .row = 1
    .col = 0
    .ColSel = .Cols - 1
End With

frmInhaltsStoffe.Show 1

CheckVerordnungMDB = True

Call DefErrPop
End Function



Function CalcTaxeZuzahlung#()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("CalcTaxeZuzahlung#")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
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
Dim i%, zz%
Dim ret#
Dim ch$

'ret# = 0#
'
'ch$ = Left$(TaxeRec!Zuzahlung, 1)
'If (TaxeRec!Warengruppe = 14) Then ch$ = "1"
'If (ch$ = "9") Then
''    MachTaxeAuswahlZuzahlung$ = "20%"
'ElseIf (ch$ = "0") Then
''    MachTaxeAuswahlZuzahlung$ = "kA"
'ElseIf (ch$ = "1") Then
''    MachTaxeAuswahlZuzahlung$ = "nb"
'ElseIf (ch$ = "2") Then
''    MachTaxeAuswahlZuzahlung$ = "ng"
'Else
'    For i% = 1 To 5
'        zz% = Val(Mid$(TaxeRec!Zuzahlung, i%, 1))
'        ret# = ret# + ZuzFeld(zz%)
'    Next i%
'End If
'
'CalcTaxeZuzahlung# = ret#

ret# = 0
If (TaxeRec!ZuzFrei = 0) And (TaxeRec!ErstattungsFähig < 3) Then
    ret# = TaxeRec!ZuzaNeu / 100
End If
CalcTaxeZuzahlung# = ret#

Call DefErrPop
End Function

Function CalcTaxeZuzahlungDOS#(Optional dPreis# = -1, Optional dZuzErlass# = -1)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("CalcTaxeZuzahlungDOS#")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
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
Dim ret#, VglWert#, AVP#, FESTBETR#


'ind% = Val(Left$(TaxeRec!Zuzahlung, 1))
'If (TaxeRec!Warengruppe = 14) Or (ind% = 9) Then
'    ret# = 1
'ElseIf (ind% >= 0) And (ind% <= 2) Then
'    ret# = ind%
'Else
'    ret# = 0
'    For i% = 1 To 5
'        ind% = Val(Mid$(TaxeRec!Zuzahlung, i%, 1))
'        ret# = ret# + ZuzFeld(ind%)
'    Next i%
'
'    AVP# = TaxeRec!VK / 100
'    FESTBETR# = TaxeRec!FESTBETRAG / 100
'    If ((TaxeRec!FestKz) And (FESTBETR# < AVP#)) Then
'        VglWert# = FESTBETR#
'    Else
'        VglWert# = AVP#
'    End If
'
'    If ((VglWert# > 0#) And (ret# > (VglWert# / VmRabattFaktor#))) Then
'        ret# = 3
'    Else
'        ret# = ret# + 1000#
'    End If
'End If


'bis inkl. Ver 1.0.61
'ret# = 0
'If (TaxeRec!ZuzFrei = 0) And (TaxeRec!ErstattungsFähig < 3) Then
'    If (dPreis# >= 0) Then
'        ret# = CalcZuzaNeu(dPreis#)
'    Else
'        ret# = TaxeRec!ZuzaNeu / 100
'    End If
'End If
'
'avp# = TaxeRec!VK / 100
'FESTBETR# = TaxeRec!FESTBETRAG / 100
'If ((TaxeRec!FestKz) And (FESTBETR# < avp#)) Then
'    VglWert# = FESTBETR#
'Else
'    VglWert# = avp#
'End If
'
'If ((VglWert# > 0#) And (ret# > (VglWert# / VmRabattFaktor#))) Then
'    ret# = 3
'Else
'    ret# = ret# + 1000#
'End If
'Ende bis inkl. Ver 1.0.61

'Ab Ver 1.0.62
ret# = 0
If (TaxeRec!ErstattungsFähig = 3) And (TaxeRec!BedErstattungsFähig = 0) Then
    ret# = 1
Else
    If (TaxeRec!BedErstattungsFähig = 2) Then
        If (dPreis# < 0) Then
            Call iMsgBox("Achtung: Artikel 'BEDINGT erstattungsfähig'", vbInformation)
        End If
    End If
    
    If (TaxeRec!ZuzFrei = 0) Then
        If (dPreis# >= 0) Then
            ret# = CalcZuzaNeu(dPreis#)
        Else
            ret# = TaxeRec!ZuzaNeu / 100
        End If
    
        If (dZuzErlass# >= 200#) Then
            ret# = dZuzErlass# - 200
        ElseIf (dZuzErlass# = 100#) Then
            ret# = 0
        ElseIf (dZuzErlass# > 0#) Then
            ret# = ret# * (100 - dZuzErlass#) / 100#
        End If
    End If
    
    AVP# = TaxeRec!vk / 100
    FESTBETR# = TaxeRec!FESTBETRAG / 100
    If ((TaxeRec!FestKz) And (FESTBETR# < AVP#)) Then
        VglWert# = FESTBETR#
    Else
        VglWert# = AVP#
    End If
    
    If ((VglWert# > 0#) And (ret# > (VglWert# / VmRabattFaktor#))) Then
        ret# = 3
    Else
        ret# = ret# + 1000#
    End If
End If

CalcTaxeZuzahlungDOS# = ret#

Call DefErrPop
End Function

Function SucheInGruppe%(BedKey$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("SucheInGruppe%")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
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
Dim ErstFlag%, erg%
Dim ActRecNo&
Dim ActBedingung$, h$
  
'  int GefFlag, ErstFlag, erg, erg2;
'  long ActRecno;
'  char h[25], jnStr[2], ActBedingung[101];

If (VmDebugAktiv%) Then
    frmAvDebug.flxAvDebug(1).AddItem "KEY" + vbTab + BedKey$
End If


ErstFlag% = True
Do

    If (ErstFlag%) Then
        FabsErrf% = VmBed.IndexSearch(0, BedKey$, FabsRecno&)
        ErstFlag% = False
    Else
        FabsErrf% = VmBed.IndexNext(0, ActRecNo&, BedKey$, FabsRecno&)
    End If
    
    If (FabsErrf% <> 0) Then
        SucheInGruppe% = False: Call DefErrPop: Exit Function
    End If
    
    Call VmBed.GetRecord(FabsRecno&)

    ActRecNo& = FabsRecno&

'    h$ = GetLastKey(24)
    h$ = GetLastKey(36)
    
    If (h$ <> BedKey$) Then
        SucheInGruppe% = False: Call DefErrPop: Exit Function
    End If

    ActBedingung$ = Trim(VmBed.Bedingung)

    If (VmDebugAktiv%) Then
        With VmBed
            h$ = "SATZ"
            h$ = h$ + vbTab + .Gruppe + .Land + .Kasse + .Verordnung
            h$ = h$ + vbTab + .Warenzeichen + vbTab + .Rundung + vbTab + .Anteil
            h$ = h$ + vbTab + .NegKnz + vbTab + .WertErmittlung + vbTab + .KostenVoranschlag
            h$ = h$ + vbTab + .GenehmigungsPflicht
        End With
        With frmAvDebug.flxAvDebug(1)
            .AddItem h$
            .AddItem "BEDING" + vbTab + "(" + VmBed.KeyBed + ")" + ActBedingung$
        End With
    End If

    erg% = PruefeBedingung%(ActBedingung$)

    If (erg%) Then
        If (InStr("JN", VmBed.Warenzeichen) > 0) Then
            If (ActWarenzeichen$ = " ") Then
                erg% = iMsgBox("Wurde das Produkt als Markenprodukt verordnet ?", vbYesNo Or vbInformation, "A+V Taxierung")
                If (erg% = vbYes) Then
                    ActWarenzeichen$ = "J"
                Else
                    ActWarenzeichen$ = "N"
                End If
            End If
            
            If (VmBed.Warenzeichen <> ActWarenzeichen$) Then
                erg% = False
            End If
        End If
    End If

    If (erg%) Then Exit Do

Loop

SucheInGruppe% = True
    
Call DefErrPop
End Function

Function SucheInGruppeMDB%(sGruppe$, iVebNr&, iPauschaleNr&)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("SucheInGruppeMDB%")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
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
Dim i%, ind%, ErstFlag%, erg%, iWert%, AnzHmAbrechnungsKz%
Dim ActRecNo&
Dim ActBedingung$, h$, h2$, BedKey$, SollHmAbrechnungsKz$
  
'  int GefFlag, ErstFlag, erg, erg2;
'  long ActRecno;
'  char h[25], jnStr[2], ActBedingung[101];

'If (iBundesland = 0) Then
'    iBundesland = ActBundesland
'End If
'BedKey$ = sGruppe$ + Format(iBundesland, "00") + Format(ActKasse, "00") + Format(ActVerordnung, "00")
BedKey$ = sGruppe$ + Format(iVebNr, "000000") + Format(ActVerordnung, "00")

If (VmDebugAktiv%) Then
    frmAvDebug.flxAvDebug(1).AddItem "KEY" + vbTab + BedKey$
End If

'HmAbrechnungsKz = Trim(CheckNullStr(VdbBedRec!HilfsmittelAbrechnungsKz))
'If (HmAbrechnungsKz <> "") Then
    HmAbrechnungsKz = ""
    AnzHmAbrechnungsKz = 0
    
    SQLStr = "SELECT GruppeNr,VebNr,HilfsmittelAbrechnungsKz,count(*) as iAnz FROM VdbBedingungen"
    SQLStr = SQLStr + " WHERE GruppeNr='" + sGruppe + "'"
    SQLStr = SQLStr + " AND VebNr=" + CStr(iVebNr)
    SQLStr = SQLStr + " AND VOrdTypNr=" + CStr(ActVerordnung)
    SQLStr = SQLStr + " group by GruppeNr,VebNr,hilfsmittelabrechnungskz"
    SQLStr = SQLStr + " order by  GruppeNr,VebNr,hilfsmittelabrechnungskz"
    FabsErrf = AplusVDB.OpenRecordset(VdbBedRec2, SQLStr)
    If Not (VdbBedRec2.EOF) Then
        Do
            If (VdbBedRec2.EOF) Then
                Exit Do
            End If
            
            If (Trim(CheckNullStr(VdbBedRec2!HilfsmittelAbrechnungsKz)) <> "") Then
                HmAbrechnungsKz = HmAbrechnungsKz + CheckNullStr(VdbBedRec2!HilfsmittelAbrechnungsKz) + ","
                AnzHmAbrechnungsKz = AnzHmAbrechnungsKz + 1
            End If
    
            VdbBedRec2.MoveNext
        Loop
    End If
    If (AnzHmAbrechnungsKz = 1) Then
        ind = InStr(HmAbrechnungsKz, ",")
        If (ind > 0) Then
            HmAbrechnungsKz = Left(HmAbrechnungsKz, ind - 1)
        End If
    ElseIf (HmAbrechnungsKz <> "") Then
        frmHmAbrechnungsKz.Show 1
    End If
'End If
    

SQLStr = "SELECT * FROM VdbBedingungen WHERE GruppeNr='" + sGruppe + "'"
'SQLStr = SQLStr + " AND KostentrNr=" + CStr(ActKasse)
'SQLStr = SQLStr + " AND LandNr=" + CStr(iBundesland)
SQLStr = SQLStr + " AND VebNr=" + CStr(iVebNr)
SQLStr = SQLStr + " AND VOrdTypNr=" + CStr(ActVerordnung)
'If (iPauschaleNr > 0) Then
    SQLStr = SQLStr + " AND PauschaleNr=" + CStr(iPauschaleNr)
'End If
FabsErrf = AplusVDB.OpenRecordset(VdbBedRec, SQLStr)
Do
    If (VdbBedRec.EOF) Then
        SucheInGruppeMDB% = 0
        Call DefErrPop: Exit Function
    End If
    
    If (VdbVebRec!NutzungsBedingungKz = 0) And (VdbBedRec!VerweisBedNr > 0) Then
        SQLStr = "SELECT * FROM VdbBedingungen WHERE BedNr=" + CStr(VdbBedRec!VerweisBedNr)
        FabsErrf = AplusVDB.OpenRecordset(VdbBedRec2, SQLStr)
        If Not (VdbBedRec2.EOF) Then
            For i% = 0 To UBound(AvVereinbarungen$)
                h2 = AvVereinbarungen(i)
                If (Val(Left(h2, 10)) = VdbBedRec2!VebNr) Then
                    If (Val(Mid(h2, 11, 1)) > 0) Then
                        If (InStr(BeigetreteneVereinbarungen, "," + CStr(VdbBedRec2!VebNr) + ",") = 0) Then
                            h = "Problem: Vereinbarung '" + Trim(Mid$(h2, 12)) + "' ist BEITRITTSPFLICHTIG!"
                            Call iMsgBox(h$, vbCritical, "A+V Taxierung")
                            erg = 0
                        End If
                    End If
                    Exit For
                End If
            Next i%
        
        End If
    End If
    
    ActBedingung$ = Trim(CheckNullStr(VdbBedRec!Bedingungsausdruck))

    If (VmDebugAktiv%) Then
        h$ = "SATZ"
'        h$ = h$ + vbTab + VdbBedRec!GruppeNr + Right$("00" + CStr(VdbBedRec!LandNr), 2) + Right$("00" + CStr(VdbBedRec!KostentrNr), 2) + Right$("00" + CStr(VdbBedRec!VOrdTypNr), 2)
        h$ = h$ + vbTab + VdbBedRec!GruppeNr + Format(VdbBedRec!VebNr, "0000000") + Right$("00" + CStr(VdbBedRec!VOrdTypNr), 2)
        
        h$ = h$ + vbTab + CheckNullStr(VdbBedRec!Warenzeichen)
        h$ = h$ + vbTab + CStr(Abs(VdbBedRec!RundungstypFl))
    
        h2$ = "N"
        If (VdbBedRec!ZuzahlungFl) Then
            h2$ = "J"
        End If
        h$ = h$ + vbTab + h2$
        
        h2$ = "N"
        If (VdbBedRec!NegativlisteFl) Then
            h2$ = "J"
        End If
        h$ = h$ + vbTab + h2$
        
'        h2$ = "0"
'        If (VdbBedRec!WertErmittlungFl) Then
'            h2$ = "1"
'        End If
        h2 = CStr(Abs(VdbBedRec!WertErmittlungKz))
        h$ = h$ + vbTab + h2$
        
        h2$ = "N"
        If (VdbBedRec!KostenvoranschlagFl) Then
            h2$ = "J"
        End If
        h$ = h$ + vbTab + h2$
        
'        h2$ = "N"
'        If (VdbBedRec!GenehmigungspflichtFl) Then
'            h2$ = "J"
'        End If
        h2 = CStr(Abs(VdbBedRec!GenehmigungsPflicht))
        h$ = h$ + vbTab + h2$
        
        h$ = h$ + vbTab + CStr(VdbBedRec!PauschaleNr)
        
        h2$ = "N"
        If (VdbBedRec!ZuzahlungFl) Then
            h2$ = "J"
        End If
        h$ = h$ + vbTab + h2$
        
        h2 = ""
        On Error Resume Next
        h2 = CStr(VdbBedRec!AbrechnungsverfahrenKz)
        On Error GoTo DefErr
        h$ = h$ + vbTab + h2
        
        h2 = VdbBedRec!HilfsmittelAbrechnungsKz
        h$ = h$ + vbTab + h2
        
        h2$ = ""
        If (VdbBedRec!VerweisBedNr > 0) Then
            h2$ = CStr(VdbBedRec!VerweisBedNr)
        End If
        h$ = h$ + vbTab + h2$
        
        With frmAvDebug.flxAvDebug(1)
            .AddItem h$
            .AddItem "BEDING" + vbTab + "(" + CStr(VdbBedRec!bednr) + ")" + ActBedingung$
            .TextMatrix(.Rows - 1, 12) = Trim(CheckNullStr(VdbBedRec!Leistungserbringergruppenschluessel))
        End With
    End If

    erg% = PruefeBedingung%(ActBedingung$)

    If (erg%) Then
        SollHmAbrechnungsKz = Trim(CheckNullStr(VdbBedRec!HilfsmittelAbrechnungsKz))
        If (SollHmAbrechnungsKz <> "") Then
            erg = (SollHmAbrechnungsKz = HmAbrechnungsKz)
        End If
    End If
        
    If (erg%) Then
        If (InStr("JN", VdbBedRec!Warenzeichen) > 0) Then
            If (ActWarenzeichen$ = " ") Then
                erg% = iMsgBox("Wurde das Produkt als Markenprodukt verordnet ?", vbYesNo Or vbInformation, "A+V Taxierung")
                If (erg% = vbYes) Then
                    ActWarenzeichen$ = "J"
                Else
                    ActWarenzeichen$ = "N"
                End If
            End If
            
            If (VdbBedRec!Warenzeichen <> ActWarenzeichen$) Then
                erg% = False
            End If
        End If
    End If

    If (erg%) Then
        If (SollHmAbrechnungsKz <> "") Then
            LEGS = Trim(CheckNullStr(VdbBedRec!Leistungserbringergruppenschluessel))
        End If
    End If
    
    If (erg%) Then Exit Do

    VdbBedRec.MoveNext
Loop

SucheInGruppeMDB% = True
    
Call DefErrPop
End Function

Function ParseAusdruck%(AusdruckStr$, SuchOp$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("ParseAusdruck%")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
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
Dim ret%, ind%, l%

ret% = False
  
ind% = InStr(AusdruckStr$, SuchOp$)
If (ind% > 0) Then
    l% = Len(SuchOp$)
    Operand$(0) = Trim(Left$(AusdruckStr$, ind% - 1))
    Operand$(1) = Trim(Mid$(AusdruckStr$, ind% + l%))

    ret% = True
End If

ParseAusdruck% = ret%
    
Call DefErrPop
End Function

Function PruefeEinzelBedingung%(EinzelBedingung$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("PruefeEinzelBedingung%")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
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
Dim erg%, ret%
Dim dWert1#, dWert2#
Dim ActOperator$
  
ret% = False

If (EinzelBedingung$ = "") Then
    PruefeEinzelBedingung% = True: Call DefErrPop: Exit Function
End If

ActOperator$ = "<="
erg% = ParseAusdruck%(EinzelBedingung$, ActOperator$)
If (erg% = False) Then
    ActOperator$ = ">="
    erg% = ParseAusdruck%(EinzelBedingung$, ActOperator$)
End If
If (erg% = False) Then
    ActOperator$ = "="
    erg% = ParseAusdruck%(EinzelBedingung$, ActOperator$)
End If
If (erg% = False) Then
    ActOperator$ = "<"
    erg% = ParseAusdruck%(EinzelBedingung$, ActOperator$)
End If
If (erg% = False) Then
    ActOperator$ = ">"
    erg% = ParseAusdruck%(EinzelBedingung$, ActOperator$)
End If
  
If (erg%) Then
    dWert1# = MachEinzelBerechnung#(Operand$(0))
    dWert2# = MachEinzelBerechnung#(Operand$(1))
    
    If (((ActOperator = "<") And (dWert1# < dWert2#)) Or _
        ((ActOperator = "=") And (dWert1# = dWert2#)) Or _
        ((ActOperator = ">") And (dWert1# > dWert2#)) Or _
        ((ActOperator = "<=") And (dWert1# <= dWert2#)) Or _
        ((ActOperator = "<>") And (dWert1# <> dWert2#)) Or _
        ((ActOperator = ">=") And (dWert1# >= dWert2#)) _
       ) Then
       ret% = True
    End If
End If

PruefeEinzelBedingung% = ret%
    
Call DefErrPop
End Function

Function PruefeBedingung%(Bedingung$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("PruefeBedingung%")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
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
Dim ret%, ActVerknuepfung%, ind%, ind2%
Dim ActBedingung$
  
ret% = False

If (Bedingung$ = "") Then
    PruefeBedingung% = True: Call DefErrPop: Exit Function
End If

Do
    ActVerknuepfung% = EINZEL_BEDINGUNG
    
    ind% = InStr(Bedingung$, "ODER")
    If (ind% > 0) Then
        ActVerknuepfung% = ODER_BEDINGUNG
        ind2% = ind% + 4
    Else
        ind% = InStr(Bedingung$, "UND")
        If (ind% > 0) Then
            ActVerknuepfung% = UND_BEDINGUNG
            ind2% = ind% + 3
        End If
    End If
    
    If (ind% > 0) Then
        ActBedingung$ = Left$(Bedingung$, ind% - 1)
    Else
        ActBedingung$ = Bedingung$
    End If
    
    ret% = PruefeEinzelBedingung%(ActBedingung$)
    
    If (ActVerknuepfung% = EINZEL_BEDINGUNG) Then Exit Do
    
    If (ret%) And (ActVerknuepfung% = ODER_BEDINGUNG) Then Exit Do
    
    If (ret% = False) And (ActVerknuepfung% = UND_BEDINGUNG) Then Exit Do
    
    Bedingung$ = Mid$(Bedingung$, ind2%)
Loop

PruefeBedingung% = ret%
    
Call DefErrPop
End Function

Function ParseOperand#(Operand$, Operator$, Optional RechStr$ = "")
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("ParseOperand#")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
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
Dim l%
Dim dWert#, dAep#
Dim AnzStr$, h$

dWert# = 0#
AnzStr$ = ""

If (Operand$ = "TAGE") And (ActMietTage% < 0) Then
    AnzStr$ = "Tage"
ElseIf (Operand$ = "WOCHEN") And (ActMietWochen% < 0) Then
    AnzStr$ = "Wochen"
ElseIf (Operand$ = "MONATE") And (ActMietMonate% < 0) Then
    AnzStr$ = "Monate"
End If

If (AnzStr$ <> "") Then
    h$ = MyInputBox("Dauer der Miete (in " + AnzStr$ + "): ", "A+V Taxierung", "1")
    l% = Val(h$)
    ActMietDauer = l%
    
    If (Operand = "TAGE") Then
        ActMietTage% = l%
    ElseIf (Operand$ = "WOCHEN") Then
        ActMietWochen% = l%
    ElseIf (Operand$ = "MONATE") Then
        ActMietMonate% = l%
    End If
End If

AnzStr$ = ""
If (Operand$ = "PATIENTENALTER(M)") And (ActPatientenAlterMonate% < 0) Then
    AnzStr$ = "Monaten"
ElseIf (Operand$ = "PATIENTENALTER(J)") And (ActPatientenAlterJahre < 0) Then
    AnzStr$ = "Jahren"
End If

If (AnzStr$ <> "") Then
    h$ = MyInputBox("Patientenalter (in " + AnzStr$ + "): ", "A+V Taxierung", "1")
    l% = Val(h$)
    
    If (Operand$ = "PATIENTENALTER(M)") Then
        ActPatientenAlterMonate% = l
    Else
        ActPatientenAlterJahre% = l
    End If
End If

  
If (InStr(Operand$, "AMPV") > 0) Then
    dAep# = ActAep#
    If (RechStr$ <> "") Then
        If (InStr(RechStr$, "MENGE*EK+AMPV")) Then
            dAep# = dAep# * ActMenge%
        End If
    End If
End If

If (Operand$ = "TAGE") Then
    dWert# = ActMietTage%
ElseIf (Operand$ = "WOCHEN") Then
    dWert# = ActMietWochen%
ElseIf (Operand$ = "MONATE") Then
    dWert# = ActMietMonate%
ElseIf (Operand$ = "MENGE") Then
    dWert# = ActMenge%
ElseIf (Operand$ = "ST") Then
    dWert# = ActStueck%
ElseIf (Operand$ = "EK") Then
    dWert# = ActAep#
ElseIf (Operand$ = "GEK") Then
    dWert# = ActAep# * ActMenge%
ElseIf (Operand$ = "VK") Then
    dWert# = ActAvp#
ElseIf (Operand$ = "VP") Or (Operand$ = "FB") Then
    If (ActWert# < 0) Then
        If (AplusVOk) Then
            SQLStr = "SELECT * FROM VdbBerechnungsformeln WHERE BedNr=" + CStr(VdbBedRec!bednr)
            SQLStr = SQLStr + " AND FormelNr=1"
'            Set VdbBerRec = AplusVDB.OpenRecordset(SQLStr)
            FabsErrf = AplusVDB.OpenRecordset(VdbBerRec, SQLStr)
            If Not (VdbBerRec.EOF) Then
                ActWert# = VdbBerRec!Wert / 100#
            End If
        Else
            AnzStr$ = VmBed.KeyBed + Format(1, "00000")
            
            FabsErrf% = VmRech.IndexSearch(0, AnzStr$, FabsRecno&)
            If (FabsErrf% = 0) Then
                Call VmRech.GetRecord(FabsRecno&)
                ActWert# = VmRech.RechWert / 100#
            End If
        End If
    End If
    dWert# = ActWert#
ElseIf (Operand$ = "AMPV") Or (Operand$ = "AMPV_ALT") Then
'    dWert# = CalcAMPV(ActAep#, ActMwst%) / (1# + (ActMwst%) / 100#) - ActAep#
    dWert# = CalcAMPV(dAep#, ActMwst%) / (1# + (ActMwst%) / 100#) - dAep#
ElseIf (Operand$ = "AMPV_NEU") Then
'    dWert# = 8.1 + (dAep# * 0.03)
'    dWert# = 8.35 + (dAep# * 0.03)   'ab 1.1.2013
'    dWert# = 8.51 + (dAep# * 0.03)   'ab 1.8.2013
'    dWert# = 8.56 + (dAep# * 0.03)   'ab 31.7.2020
    dWert# = 8.76 + (dAep# * 0.03)   'ab 19.7.2023; jetzt inkl 0.20 EUR für 'Finanzierung zusätzlicher pharmazeutischer Dienstleistungen nach § 129 Absatz 5e des Fünften Buches SGB'
ElseIf (Operand$ = "PATIENTENALTER(M)") Then
    dWert# = ActPatientenAlterMonate
ElseIf (Operand$ = "PATIENTENALTER(J)") Then
    dWert# = ActPatientenAlterJahre
Else
    l% = Len(Operand$)
'    for (j=0; j<l; j++)  {
'      if (Operand[j] == ',')
'        Operand[j] = '.';
'    } /* for */

    If (Right$(Operand$, 1) = "%") Then
        Operand$ = Left$(Operand$, l% - 1)
        If (Operator$ = "-") Then
            dWert# = 1# - (xVal(Operand$) / 100#)
        Else
            dWert# = 1# + (xVal(Operand$) / 100#)
        End If
        Operator$ = "*"
    Else
        dWert# = xVal(Operand$)
    End If
End If

ParseOperand# = dWert#
    
Call DefErrPop
End Function

Function APVparse(zeile As String, RechenFlag As Integer, Berechnung As String) As Double

'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("APVparse")
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
If (Err.Number = (999 + vbObjectError)) Then End
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
'On Error GoTo 0
'Err.Raise 999 + vbObjectError, "DSK.VBP", "Fehler in DLL"
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim i As Integer
Dim b As String
Dim p As Integer
Dim klammer As Integer
Dim c As String
Dim minop As Integer
Dim op As Integer
Dim opWert As Integer
Dim oppos As Integer
Dim oplen As Integer
Dim op2 As Integer
Dim wert1 As Double
Dim wert2 As Double
Dim erg As Double
Dim h As String
Dim iFunktion As Integer
Dim apvTmp As New ADODB.Recordset

b$ = zeile
b$ = LTrim(RTrim(b$))

'"Euro" wegschneiden - hat keine Funktion
Do
  i = InStr(UCase(b$), "EURO")
  If i > 0 Then
    b$ = Left(b$, i - 1) + Mid(b$, i + 4)
  End If
Loop Until i = 0

'Klammern entfernen
If b$ = "" Then APVparse# = 0#: Call DefErrPop: Exit Function
If Left$(b$, 1) = "(" And Right$(b$, 1) = ")" Then
  p% = 1: klammer% = 0
  Do
    c$ = Mid$(b$, p%, 1)
    If c$ = "(" Then
      klammer% = klammer% + 1
    ElseIf c$ = ")" Then
      klammer% = klammer% - 1
      If klammer% = 0 Then
        If p% = Len(b$) Then b$ = Mid$(b$, 2, Len(b$) - 2)
        Exit Do
      End If
    End If
    p% = p% + 1
  Loop While p% <= Len(b$)
End If

'höchstwertigen Operator suchen

p% = 1: klammer% = 0
minop% = 0
Do
  c$ = Mid$(b$, p%, 1)
  If c$ = "(" Then
    klammer% = klammer% + 1
  ElseIf c$ = ")" Then
    klammer% = klammer% - 1
  Else
    op% = InStr(";:*/+-<=>|&", c$)
    If klammer% = 0 And op% > 0 Then
      opWert% = op%
      
'       GS 11.06.08 die Wahrheit liegt in der Mitte: es geht zwar Punktrechnung vor Strichrechnung, aber z.B. bei / und * gilt die Reihenfolge innerhabl der Berechnungsvorschrift (nicht kommutativ!)
      If op% = 3 Or op% = 4 Then
        opWert% = 4
      End If
      If op% = 5 Or op% = 6 Then
        opWert% = 6
      End If
      
'      'GS 11.3.08: seit 1.1.08 gilt offenbar doch "Punktrechnung vor Strichrechnung"
      
'      If op% >= 3 And op% <= 6 Then
'        opWert% = 6
'        'bei A+V-Berechnung ist Wertigkeit der math. OP wurst: es wird von vorne nach hinten abgearbeitet.
'        'damit das funkt, werden alle math. OP auf gleichen Wert gesetzt und auch bei Gleichheit der neue als minop genommen
'      End If
      
 '     oplen% = 1
      If op% >= 7 Then      'Vgl- oder logischer Op.
        If op% >= 10 Then   'AND = 20, OR = 21
          op% = op% + 10
          opWert% = op%
        Else                'Vergleichsoperatoren
          op2% = InStr("<=>", Mid$(b$, p% + 1, 1))
          If op2% > 0 Then
            op% = op% + 3 + op2%
            opWert% = op%
            '<= ... 12
            '=< ... 12
            '<> ... 13
            '>= ... 14
            '=> ... 14
'            oplen% = 2
            p% = p% + 1
          End If
        End If
      End If
      If minop% <= opWert% Then
        minop% = op%
        If op% >= 12 Then oplen% = 2 Else oplen% = 1    '1.2.95
        oppos% = p% - (oplen% - 1)
      End If
    End If
  End If
  p% = p% + 1
Loop While p% <= Len(b$)
If minop% > 0 Then
  wert1# = APVparse#(Left$(b$, oppos% - 1), RechenFlag, Berechnung)
  wert2# = APVparse#(Mid$(b$, oppos% + oplen%), RechenFlag, Berechnung)
End If

Select Case minop%
Case 1
  erg# = wert1#: If wert1# <= 0# Then erg# = wert2#
Case 2
  erg# = wert1#: If wert2# < wert1# Then erg# = wert2#
Case 3
  erg# = wert1# * wert2#
Case 4
  erg# = wert1# / wert2#
Case 5
  erg# = wert1# + wert2#
Case 6
  erg# = wert1# - wert2#

Case 7
  erg# = (wert1# < wert2#)
Case 8
  erg# = (wert1# = wert2#)
Case 9
  erg# = (wert1# > wert2#)
Case 12
  erg# = (wert1# <= wert2#)
Case 13
  erg# = (wert1# <> wert2#)
Case 14
  erg# = (wert1# >= wert2#)

Case 20
  erg# = wert1# Or wert2#
Case 21
  erg# = wert1# And wert2#
  
Case Else
  'If minop% <> 0 Then Stop
  'Variable oder Konstante
  Select Case b$
  Case "ST"
    erg# = ActStueck%
  Case "EK"
    erg# = ActAep#
  Case "AMPV", "AMPV_ALT"
    If InStr(UCase(Berechnung), "MENGE*EK+AMPV") > 0 Then
      erg# = CalcAMPV(ActAep# * ActMenge, ActMwst%) / (1# + (ActMwst%) / 100#) - (ActAep# * ActMenge) 'APVampv#(ap.EK * ap.menge)
    Else
      erg# = CalcAMPV(ActAep#, ActMwst%) / (1# + (ActMwst%) / 100#) - ActAep#   'APVampv#(ap.EK)
    End If
  Case "AMPV_NEU"
'    erg# = 8.51 + (ActAep# * 0.03)  'ab 1.8.2013
    erg# = 8.56 + (ActAep# * 0.03)  'ab 31.7.2020
    erg# = 8.76 + (ActAep# * 0.03)   'ab 19.7.2023; jetzt inkl 0.20 EUR für 'Finanzierung zusätzlicher pharmazeutischer Dienstleistungen nach § 129 Absatz 5e des Fünften Buches SGB'
  Case "VP", "FB"
    If (ActWert# < 0) Then
        SQLStr = "SELECT * FROM VdbBerechnungsformeln WHERE BedNr=" + CStr(VdbBedRec!bednr)
        SQLStr = SQLStr + " AND FormelNr=1"
        FabsErrf = AplusVDB.OpenRecordset(VdbBerRec, SQLStr)
        If Not (VdbBerRec.EOF) Then
            ActWert# = VdbBerRec!Wert / 100#
        End If
    End If
    erg# = ActWert#
  Case "MENGE"
    erg# = ActMenge%
  Case "VK"
    erg# = ActAvp#
  Case "GEK"
    erg# = ActAep# * ActMenge%
  Case "TAGE", "WOCHEN", "MONATE"
        
    If ActMietZeit < 0 Then
'      If FragenAnzeigen Then
        h = InputBox("Dauer der Miete (in " + FirstLettersUcase(b$) + "):")
        ActMietZeit = Val(h)
'      Else
'        ActMietZeit = 0
'      End If
    End If
    erg# = ActMietZeit
      
  Case Else
    If UCase(Left(b$, 6)) = "RUNDE(" Or UCase(Left(b$, 9)) = "AMPREISV(" Or UCase(Left(b$, 5)) = "MWST(" Then
      If UCase(Left(b$, 6)) = "RUNDE(" Then
        b$ = Mid$(b$, 7)
        iFunktion = 1
      ElseIf UCase(Left(b$, 9)) = "AMPREISV(" Then
        b$ = Mid$(b$, 10)
        iFunktion = 2
      ElseIf UCase(Left(b$, 5)) = "MWST(" Then
        b$ = Mid$(b$, 6)
        iFunktion = 3
      End If
      p% = 1: klammer% = 0
      Do
        c$ = Mid$(b$, p%, 1)
        If c$ = "(" Then
          klammer% = klammer% + 1
        ElseIf c$ = ")" Then
          klammer% = klammer% - 1
        ElseIf (klammer = 0) Then
          If (c$ = ";") Then
              wert1# = APVparse#(Mid$(b$, 1, p - 1), RechenFlag, Berechnung)
              wert2# = Val(Mid(b$, p + 1, Len(b$) - 1 - p))
'              Call DebugFile("minop=0: Wert1: " + CStr(wert1) + ", Wert2: " + CStr(wert2))
              If (iFunktion = 1) Then
                  erg# = fnxAE(wert1#, CInt(wert2#))
              ElseIf (iFunktion = 2) Then
                  If (wert2 = 1#) Then
'                      erg# = wert1 + (8.51 + (wert1 * 0.03))
'                      erg# = wert1 + (8.56 + (wert1 * 0.03))    'ab 31.7.2020
                      erg# = wert1 + (8.76 + (wert1 * 0.03))    'ab 19.7.2023; jetzt inkl 0.20 EUR für 'Finanzierung zusätzlicher pharmazeutischer Dienstleistungen nach § 129 Absatz 5e des Fünften Buches SGB'
                  Else
                      erg# = wert1 + CalcAMPV(wert1, ActMwst%) / (1# + (ActMwst%) / 100#) - wert1   'APVampv#(ap.EK)   'APVampv#(wert1)
                  End If
              ElseIf iFunktion = 3 Then
                Select Case wert2#
                Case 1
                  erg# = wert1# * (1 + ActMwst / 100#)
                Case 2
                  erg# = wert1# * (1 + para.Mwst(2) / 100#)
                Case 3
                  erg# = wert1# * (1 + para.Mwst(1) / 100#)
                Case 4
                  erg# = wert1# / (1 + ActMwst / 100#)
                Case 5
                  erg# = wert1# / (1 + para.Mwst(2) / 100#)
                Case 6
                  erg# = wert1# / (1 + para.Mwst(1) / 100#)
                End Select
              End If
              Exit Do
          End If
        End If
        p% = p% + 1
      Loop While p% <= Len(b$)
      
    Else
      'Komma auf Punkt tauschen
      Do
        i% = InStr(b$, ","): If i% > 0 Then Mid$(b$, i%, 1) = "."
      Loop While i% > 0
      erg# = Val(b$)
    End If
  End Select
End Select
APVparse# = erg#
'If minop% <> 0 Then Call DebugFile("Parse Zwischenergebnis: " + CStr(erg#))
Call DefErrPop
End Function

Function fnxAE(x As Double, NachKomma As Integer) As Double
Dim dPower As Double

dPower = 10 ^ NachKomma
fnxAE = Sgn(x) * CDbl(Int(Abs(x) * dPower + 0.501) / dPower)

End Function

Function MachEinzelBerechnung#(EinzelBerechnung$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("MachEinzelBerechnung#")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
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
Dim i%, l%
Dim EinzelPreis#, dWert#
Dim hBerechnung$, ActOperand$, ActOperator$, ch$
  
EinzelPreis# = 0#
hBerechnung$ = Trim(EinzelBerechnung$)

l% = Len(hBerechnung$)
If (l% = 0) Then
    MachEinzelBerechnung# = 0#: Call DefErrPop: Exit Function
End If


hBerechnung$ = hBerechnung$ + "@"
l% = l% + 1

ActOperator$ = ""
  
For i% = 1 To l%
    ch$ = Mid$(hBerechnung$, i%, 1)
    If (InStr("+-*/", ch$) > 0) Or (ch$ = "@") Then
        dWert# = ParseOperand#(ActOperand$, ActOperator$, hBerechnung$)
        Select Case ActOperator$
            Case "+"
                EinzelPreis# = EinzelPreis# + dWert#
            Case "-"
                EinzelPreis# = EinzelPreis# - dWert#
            Case "*"
                EinzelPreis# = EinzelPreis# * dWert#
            Case "/"
                EinzelPreis# = EinzelPreis# / dWert#
            Case Else
                EinzelPreis# = dWert#
        End Select
        
        ActOperand$ = ""
        ActOperator$ = ch$
    Else
        ActOperand$ = ActOperand$ + ch$
    End If
Next i%

MachEinzelBerechnung# = FNX(EinzelPreis#)
    
Call DefErrPop
End Function

Function MachBerechnung%(KeyBed$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("MachBerechnung%")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
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
Dim i%, AnzGef%, VkBerechnung%, ind%, ind2%, ActAVrabatt%(2)
Dim ActPreis#(2), ActPreis20#(2), EinzelPfennig#, KlammerPreis#, dPreis#, diff#
Dim ActKeyBed$, ActBerechnung$, ActBer2$, h$, ch$
  
If (VmDebugAktiv%) Then
    With frmAvDebug.flxAvDebug(2)
        .AddItem "BERECH"  '18
        .AddItem "WERT"    '19
        .AddItem "EINZEL"    '20
        .AddItem "MWST"    '21
        .AddItem "RABATT"    '22
        .AddItem "RUNDEN"    '23
    End With
End If


AnzGef% = 0

For i% = 1 To 1000

    ActKeyBed$ = KeyBed$ + Format(i%, "00000")
    FabsErrf% = VmRech.IndexSearch(0, ActKeyBed$, FabsRecno&)
    If (FabsErrf% > 0) Then Exit For
        
    Call VmRech.GetRecord(FabsRecno&)
    
    ActWert# = VmRech.RechWert / 100#
    
    ActBerechnung$ = Trim(VmRech.Berechnung)
'    ActBerechnung$ = "(EK+25%)*0,95"
'    ActBerechnung$ = "EK+ST*0,4"
'    ActBerechnung$ = "EK-15%"
    
    h$ = ActBerechnung$
    ind% = InStr(h$, "+")
    If (ind% <= 0) Then
        ind% = InStr(h$, "-")
    End If
    If (ind% > 0) Then
        ind% = InStr(h$, "*")
        If (ind% > 1) Then
            If (Mid$(h$, ind% - 1, 1) <> ")") Then
                Do
                    If (ind% = 1) Then
                        ind% = 0
                        Exit Do
                    End If
    
                    ind% = ind% - 1
                    ch$ = Mid$(h$, ind%, 1)
                    If (ch$ = "(") Then
                        ind% = -1
                        Exit Do
                    End If
                    If (InStr("+-*/", ch$) > 0) Or (ch$ = "@") Then
                        Exit Do
                    End If
                Loop
    
                If (ind% >= 0) Then
                    ActBerechnung$ = Left$(h$, ind%) + "(" + Mid$(h$, ind% + 1) + ")"
                End If
            End If
        End If
    End If
    
    If (VmDebugAktiv%) Then
        With frmAvDebug.flxAvDebug(2)
            ind% = AnzGef% * 3 + 1
            .TextMatrix(1, ind%) = ActBerechnung$
            .TextMatrix(1, ind% + 1) = VmRech.Mwst
            .TextMatrix(1, ind% + 2) = VmRech.Rabatt
            If (ActWert# <> 0#) Then
                .TextMatrix(2, ind%) = Format(ActWert#, "0.00")
            End If
        End With
    End If

    
    Do
        ind% = InStr(ActBerechnung$, "(")
        If (ind% = 0) Then Exit Do
        
        
        ind2% = InStr(ind% + 1, ActBerechnung$, ")")
        If (ind% = 0) Then
            MachBerechnung% = False: Call DefErrPop: Exit Function
        End If
        
        ActBer2$ = Mid$(ActBerechnung$, ind% + 1, ind2% - ind% - 1)
        
        KlammerPreis# = MachEinzelBerechnung#(ActBer2$)
        
        ActBerechnung$ = Left$(ActBerechnung$, ind% - 1) + Format(KlammerPreis#, "    0.0000") + Mid$(ActBerechnung$, ind2% + 1)
    Loop
    
    
    ActPreis#(AnzGef%) = MachEinzelBerechnung#(ActBerechnung$)
    
    If (VmDebugAktiv%) Then
        ind% = AnzGef% * 3 + 1
        frmAvDebug.flxAvDebug(2).TextMatrix(3, ind%) = Format(ActPreis#(AnzGef%), "0.00")
    End If
    
    VkBerechnung% = False
    If (InStr(VmRech.Berechnung, "VK") > 0) Then VkBerechnung% = True
    If (VmRech.Mwst = "J") And (VkBerechnung% = False) Then
        ActPreis#(AnzGef%) = ActPreis#(AnzGef%) * (1# + (ActMwst) / 100#)
    End If
    
    If (VmDebugAktiv%) Then
        ind% = AnzGef% * 3 + 1
        frmAvDebug.flxAvDebug(2).TextMatrix(4, ind%) = Format(ActPreis#(AnzGef%), "0.00")
    End If
    
    If (VmPzn.HilfsmittelKz = 2) Or (Left$(TaxeRec!Zuzahlung, 1) = "9") Then
'        If (ActKasse% = 2) Then VmRech.Rabatt = "J"
        If (ActKasse% = 2) Then VmRech.Rabatt = "0"
    End If
    
    ActPreis20#(AnzGef%) = ActPreis#(AnzGef%)

'    If (VmRech.Rabatt = "J") And (ActKasse% <> 2) Then
    If (VmRech.Rabatt = "0") Then
        If (ActKasse% <> 2) Then
            ActPreis20#(AnzGef%) = ActPreis20#(AnzGef%) / VmRabattFaktor
        End If
    ElseIf (VmRech.Rabatt = "5") Then
        If (TaxeRec!ApoRabattKz = 1) Then
            VmRech.Rabatt = "6"
        ElseIf (TaxeRec!ApoRabattKz = 2) Then
            VmRech.Rabatt = "7"
        Else
            VmRech.Rabatt = "4"
        End If
    End If
    
    ActAVrabatt%(AnzGef%) = Val(VmRech.Rabatt)
        
'    If (VmRech.Rabatt = "N") Then
    If (VmRech.Rabatt = "1") Or (VmRech.Rabatt = "2") Then
        ActPreis#(AnzGef%) = ActPreis#(AnzGef%) * VmRabattFaktor
        
        If (VmDebugAktiv%) Then
            ind% = AnzGef% * 3 + 1
            frmAvDebug.flxAvDebug(2).TextMatrix(5, ind%) = Format(ActPreis#(AnzGef%), "0.00")
        End If
    ElseIf (VmRech.Rabatt = "3") Then
        dPreis# = ActPreis20#(AnzGef%)
        Select Case dPreis#
            Case 0# To 52.46
                dPreis# = dPreis# * 0.94
            Case 52.47 To 54.8
                dPreis# = 49.32
            Case 54.81 To 820.22
                dPreis# = dPreis# * 0.9
            Case Else
                diff# = dPreis# - 820.22
                dPreis# = dPreis# * 0.9 - (diff# * 0.06)
        End Select
        ActPreis20#(AnzGef%) = dPreis#
    ElseIf (VmRech.Rabatt = "6") Then
        ActPreis20#(AnzGef%) = ActPreis20#(AnzGef%) - 2#
    ElseIf (VmRech.Rabatt = "7") Then
        ActPreis20#(AnzGef%) = ActPreis20#(AnzGef%) * 0.95
    ElseIf (VmRech.Rabatt = "8") Then
        ActPreis#(AnzGef%) = ActPreis#(AnzGef%) * 1.05263
        
        If (VmDebugAktiv%) Then
            ind% = AnzGef% * 3 + 1
            frmAvDebug.flxAvDebug(2).TextMatrix(5, ind%) = Format(ActPreis#(AnzGef%), "0.00")
        End If
    End If

    
    AnzGef% = i%

Next i%

If (AnzGef% = 0) Then
    MachBerechnung% = False: Call DefErrPop: Exit Function
End If


ActVmPreis# = 9999999.99
If (VmBed.WertErmittlung = "1") And (ActPreis#(0) > 0#) Then
    ActVmPreis# = ActPreis#(0)
    Act20Preis# = ActPreis20#(0)
    ActVmRabatt% = ActAVrabatt%(0)
Else
    For i% = 0 To (AnzGef% - 1)
        If (ActPreis#(i%) > 0) And (ActPreis#(i%) < ActVmPreis#) Then
            ActVmPreis# = ActPreis#(i%)
            Act20Preis# = ActPreis20#(i%)
            ActVmRabatt% = ActAVrabatt%(i%)
        End If
    Next i%
End If


If (ActVmPreis# = 9999999.99) Then
    MachBerechnung% = False: Call DefErrPop: Exit Function
End If


If (InStr("01", VmBed.Rundung) > 0) Then
    If (VmBed.Rundung = "0") Then
'        ActVmPreis# = (ActVmPreis# * 100#) '+ 0.501
'        ActVmPreis# = CLng(ActVmPreis#)
'        ActVmPreis# = ActVmPreis# / 100#
        ActVmPreis# = CDbl(Int(Abs(ActVmPreis#) * 100# + 0.501) / 100#)
    Else
        ActVmPreis# = ActVmPreis# * 10#
        
        EinzelPfennig# = ActVmPreis# - Fix(ActVmPreis#)
        If (EinzelPfennig# < 0.25) Then
            EinzelPfennig# = 0#
        ElseIf (EinzelPfennig# < 0.75) Then
            EinzelPfennig# = 0.5
        Else
            EinzelPfennig# = 1#
        End If
        
        ActVmPreis# = Fix(ActVmPreis#) + EinzelPfennig#
        ActVmPreis# = ActVmPreis# / 10#
    End If
    
    If (VmDebugAktiv%) Then
        frmAvDebug.flxAvDebug(2).TextMatrix(6, 1) = Format(ActVmPreis#, "0.00")
    End If
End If

MachBerechnung% = True
    
Call DefErrPop
End Function

Function MachBerechnungMDB%(lBedNr&)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("MachBerechnungMDB%")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
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
Dim i%, AnzGef%, VkBerechnung%, ind%, ind2%, ActAVrabatt%(2), iRabatt%, iWert%, p1%, p2%, AbrechnungsverfahrenKz%
Dim EinzelPfennig#, KlammerPreis#, dPreis#, diff#, erg#
Dim ActKeyBed$, ActBerechnung$, ActBer2$, h$, h2$, ch$, c$, x$
  
If (VmDebugAktiv%) Then
    With frmAvDebug.flxAvDebug(2)
        .AddItem "BERECH"  '18
        .AddItem "WERT"    '19
        .AddItem "EINZEL"    '20
        .AddItem "MWST"    '21
        .AddItem "RABATT"    '22
        .AddItem "RUNDEN"    '23
        .AddItem "ABRECH-Nr/HM-Nr"    '24
    End With
End If


AnzGef% = 0

SQLStr = "SELECT * FROM VdbBerechnungsformeln WHERE BedNr=" + CStr(lBedNr&)
SQLStr = SQLStr + " ORDER BY FormelNr"
'Set VdbBerRec = AplusVDB.OpenRecordset(SQLStr)
FabsErrf = AplusVDB.OpenRecordset(VdbBerRec, SQLStr)
For i% = 1 To 1000
    If (VdbBerRec.EOF) Then
        Exit For
    End If

    ActWert# = CheckNullLong(VdbBerRec!Wert) / 100#
    
    AvAbrechnungsNr(AnzGef) = Trim(CheckNullStr(VdbBerRec!AbrechnungsNr))
    AvHmPositionsNr(AnzGef) = Trim(CheckNullStr(VdbBerRec!HmPositionsNr))
    
    ActBerechnung$ = Trim(CheckNullStr(VdbBerRec!Formelausdruck))
    
'    h$ = ActBerechnung$
'    ind% = InStr(h$, "+")
'    If (ind% <= 0) Then
'        ind% = InStr(h$, "-")
'    End If
'    If (ind% > 0) Then
'        ind% = InStr(h$, "*")
'        If (ind% > 1) Then
'            If (Mid$(h$, ind% - 1, 1) <> ")") Then
'                Do
'                    If (ind% = 1) Then
'                        ind% = 0
'                        Exit Do
'                    End If
'
'                    ind% = ind% - 1
'                    ch$ = Mid$(h$, ind%, 1)
'                    If (ch$ = "(") Then
'                        ind% = -1
'                        Exit Do
'                    End If
'                    If (InStr("+-*/", ch$) > 0) Or (ch$ = "@") Then
'                        Exit Do
'                    End If
'                Loop
'
'                If (ind% >= 0) Then
'                    ActBerechnung$ = Left$(h$, ind%) + "(" + Mid$(h$, ind% + 1) + ")"
'                End If
'            End If
'        End If
'    End If
    
    If (VmDebugAktiv%) Then
        With frmAvDebug.flxAvDebug(2)
            ind% = AnzGef% * 3 + 1
            .TextMatrix(1, ind%) = ActBerechnung$
            
            iWert% = Abs(VdbBerRec!MehrwertSteuerFl)
            If (iWert% = 1) Then
                h2$ = "J"
            Else
                h2$ = "N"
            End If
            .TextMatrix(1, ind% + 1) = h2$
            
            .TextMatrix(1, ind% + 2) = CStr(VdbBerRec!Kassenrabatt)
            
            If (ActWert# <> 0#) Then
                .TextMatrix(2, ind%) = Format(ActWert#, "0.00")
            End If
        End With
    End If

    
'    Do
'        ind% = InStr(ActBerechnung$, "(")
'        If (ind% = 0) Then Exit Do
'
'
'        ind2% = InStr(ind% + 1, ActBerechnung$, ")")
'        If (ind% = 0) Then
'            MachBerechnungMDB% = False: Call DefErrPop: Exit Function
'        End If
'
'        ActBer2$ = Mid$(ActBerechnung$, ind% + 1, ind2% - ind% - 1)
'
'        KlammerPreis# = MachEinzelBerechnung#(ActBer2$)
'
'        ActBerechnung$ = Left$(ActBerechnung$, ind% - 1) + Format(KlammerPreis#, "    0.0000") + Mid$(ActBerechnung$, ind2% + 1)
'    Loop
'
'
'    ActPreis#(AnzGef%) = MachEinzelBerechnung#(ActBerechnung$)
    
    'alle +x oder -x durch *(1+x/100) oder *(1-x/100) ersetzen
    p1% = InStr(ActBerechnung$, "%")
    If p1% > 0 Then
      p2% = p1%
      Do While p1% > 1
        p1% = p1% - 1
        c$ = Mid$(ActBerechnung$, p1%, 1)
        If c$ = "+" Or c$ = "-" Then Exit Do
      Loop
      If p1% > 1 Then
        x$ = Mid$(ActBerechnung$, p1% + 1, p2% - p1% - 1)
        ActBerechnung$ = Left$(ActBerechnung, p1% - 1) + "*" + "(1" + c$ + "(" + x$ + "/100))" + Mid(ActBerechnung$, p2% + 1)
      End If
    End If
    
    erg# = APVparse#(ActBerechnung$, True, ActBerechnung$)
    ActPreis#(AnzGef%) = FNX(erg#)
  
    AbrechnungsverfahrenKz = 0
    On Error Resume Next
    AbrechnungsverfahrenKz = VdbBedRec!AbrechnungsverfahrenKz
    On Error GoTo DefErr
    If (AbrechnungsVerfahren_2_Aktiv) And (AbrechnungsverfahrenKz = 2) Then
        ActPreis#(AnzGef%) = ActPreis#(AnzGef%) * ActMenge%
    End If
    
    
    If (VmDebugAktiv%) Then
        ind% = AnzGef% * 3 + 1
        frmAvDebug.flxAvDebug(2).TextMatrix(3, ind%) = Format(ActPreis#(AnzGef%), "0.00")
    End If
    
    VkBerechnung% = False
    If (InStr(CheckNullStr(VdbBerRec!Formelausdruck), "VK") > 0) Then VkBerechnung% = True
    If (Abs(VdbBerRec!MehrwertSteuerFl) > 0) And (VkBerechnung% = False) Then
        ActPreis#(AnzGef%) = ActPreis#(AnzGef%) * (1# + (ActMwst) / 100#)
    End If
    ActPreis#(AnzGef%) = FNX(ActPreis#(AnzGef%))
    
    If (VmDebugAktiv%) Then
        ind% = AnzGef% * 3 + 1
        frmAvDebug.flxAvDebug(2).TextMatrix(4, ind%) = Format(ActPreis#(AnzGef%), "0.00")
    End If
    
    iRabatt% = VdbBerRec!Kassenrabatt
    
    If (VdbPznRec!HilfsmittelKz = 2) Or (Left$(TaxeRec!Zuzahlung, 1) = "9") Then
'        If (ActKasse% = 2) Then VmRech.Rabatt = "J"
        If (ActKasse% = 2) Then iRabatt = 0
    End If
    
    ActPreis20#(AnzGef%) = ActPreis#(AnzGef%)

'    If (VmRech.Rabatt = "J") And (ActKasse% <> 2) Then
    If (iRabatt = 0) Then
        If (ActKasse% <> 2) Then
            ActPreis20#(AnzGef%) = ActPreis20#(AnzGef%) / VmRabattFaktor
        End If
    ElseIf (iRabatt = 5) Then
        If (TaxeRec!ApoRabattKz = 1) Then
            iRabatt = 6
        ElseIf (TaxeRec!ApoRabattKz = 2) Then
            iRabatt = 7
        Else
            iRabatt = 4
        End If
    End If
    
    ActAVrabatt%(AnzGef%) = Val(iRabatt)
        
'    If (VmRech.Rabatt = "N") Then
    If (iRabatt = 1) Or (iRabatt = 2) Then
        ActPreis#(AnzGef%) = ActPreis#(AnzGef%) * VmRabattFaktor
        
        If (VmDebugAktiv%) Then
            ind% = AnzGef% * 3 + 1
            frmAvDebug.flxAvDebug(2).TextMatrix(5, ind%) = Format(ActPreis#(AnzGef%), "0.00")
        End If
    ElseIf (iRabatt = 3) Then
        dPreis# = ActPreis20#(AnzGef%)
        Select Case dPreis#
            Case 0# To 52.46
                dPreis# = dPreis# * 0.94
            Case 52.47 To 54.8
                dPreis# = 49.32
            Case 54.81 To 820.22
                dPreis# = dPreis# * 0.9
            Case Else
                diff# = dPreis# - 820.22
                dPreis# = dPreis# * 0.9 - (diff# * 0.06)
        End Select
        ActPreis20#(AnzGef%) = dPreis#
    ElseIf (iRabatt = 6) Then
        ActPreis20#(AnzGef%) = ActPreis20#(AnzGef%) - 2#
    ElseIf (iRabatt = 7) Then
        ActPreis20#(AnzGef%) = ActPreis20#(AnzGef%) * 0.95
    ElseIf (iRabatt = 8) Then
        ActPreis#(AnzGef%) = ActPreis#(AnzGef%) * 1.05263
        
        If (VmDebugAktiv%) Then
            ind% = AnzGef% * 3 + 1
            frmAvDebug.flxAvDebug(2).TextMatrix(5, ind%) = Format(ActPreis#(AnzGef%), "0.00")
        End If
    End If
    ActPreis#(AnzGef%) = FNX(ActPreis#(AnzGef%))
    
    If (VmDebugAktiv%) Then
        If (AvAbrechnungsNr(AnzGef) <> "") Or (AvHmPositionsNr(AnzGef) <> "") Then
            ind% = AnzGef% * 3 + 1
            
            h2 = ""
            If (AvAbrechnungsNr(AnzGef) <> "") Then
                h2 = AvAbrechnungsNr(AnzGef)
            End If
            h2 = h2 + " / "
            If (AvHmPositionsNr(AnzGef) <> "") Then
                h2 = h2 + AvHmPositionsNr(AnzGef)
            End If
            frmAvDebug.flxAvDebug(2).TextMatrix(7, ind%) = h2$
        End If
    End If
    
    If (VdbBedRec!RundungstypFl) Then
        ActPreis#(AnzGef) = ActPreis#(AnzGef) * 10#
        
        EinzelPfennig# = ActPreis#(AnzGef) - Fix(ActPreis#(AnzGef))
        If (EinzelPfennig# < 0.25) Then
            EinzelPfennig# = 0#
        ElseIf (EinzelPfennig# < 0.75) Then
            EinzelPfennig# = 0.5
        Else
            EinzelPfennig# = 1#
        End If
        
        ActPreis#(AnzGef) = Fix(ActPreis#(AnzGef)) + EinzelPfennig#
        ActPreis#(AnzGef) = ActPreis#(AnzGef) / 10#
    Else
        ActPreis#(AnzGef) = CDbl(Int(Abs(ActPreis#(AnzGef)) * 100# + 0.501) / 100#)
    End If

    
    AnzGef% = i%
    
    VdbBerRec.MoveNext
Next i%

If (AnzGef% = 0) Then
    MachBerechnungMDB% = False: Call DefErrPop: Exit Function
End If


ActVmPreis# = 9999999.99
If (VdbBedRec!WertErmittlungFl) And (ActPreis#(0) > 0#) Then
    ActVmPreis# = ActPreis#(0)
    Act20Preis# = ActPreis20#(0)
    ActVmRabatt% = ActAVrabatt%(0)
Else
    For i% = 0 To (AnzGef% - 1)
        If (ActPreis#(i%) > 0) And (ActPreis#(i%) < ActVmPreis#) Then
            ActVmPreis# = ActPreis#(i%)
            Act20Preis# = ActPreis20#(i%)
            ActVmRabatt% = ActAVrabatt%(i%)
        End If
    Next i%
End If


If (ActVmPreis# = 9999999.99) Then
    MachBerechnungMDB% = False: Call DefErrPop: Exit Function
End If


'If (VdbBedRec!RundungstypFl) Then
'    ActVmPreis# = ActVmPreis# * 10#
'
'    EinzelPfennig# = ActVmPreis# - Fix(ActVmPreis#)
'    If (EinzelPfennig# < 0.25) Then
'        EinzelPfennig# = 0#
'    ElseIf (EinzelPfennig# < 0.75) Then
'        EinzelPfennig# = 0.5
'    Else
'        EinzelPfennig# = 1#
'    End If
'
'    ActVmPreis# = Fix(ActVmPreis#) + EinzelPfennig#
'    ActVmPreis# = ActVmPreis# / 10#
'Else
'    ActVmPreis# = CDbl(Int(Abs(ActVmPreis#) * 100# + 0.501) / 100#)
'End If

If (VmDebugAktiv%) Then
    frmAvDebug.flxAvDebug(2).TextMatrix(6, 1) = Format(ActVmPreis#, "0.00")
End If

MachBerechnungMDB% = True
    
Call DefErrPop
End Function

'Function MachBerechnungMDB%(lBedNr&)
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Call DefErrFnc("MachBerechnungMDB%")
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
'Dim i%, AnzGef%, VkBerechnung%, ind%, ind2%, ActAVrabatt%(2), iRabatt%, iWert%
'Dim EinzelPfennig#, KlammerPreis#, dPreis#, diff#
'Dim ActKeyBed$, ActBerechnung$, ActBer2$, h$, h2$, ch$
'
'If (VmDebugAktiv%) Then
'    With frmAvDebug.flxAvDebug(2)
'        .AddItem "BERECH"  '18
'        .AddItem "WERT"    '19
'        .AddItem "EINZEL"    '20
'        .AddItem "MWST"    '21
'        .AddItem "RABATT"    '22
'        .AddItem "RUNDEN"    '23
'        .AddItem "ABRECH-Nr/HM-Nr"    '24
'    End With
'End If
'
'
'AnzGef% = 0
'
'SQLStr = "SELECT * FROM VdbBerechnungsformeln WHERE BedNr=" + CStr(lBedNr&)
'SQLStr = SQLStr + " ORDER BY FormelNr"
''Set VdbBerRec = AplusVDB.OpenRecordset(SQLStr)
'FabsErrf = AplusVDB.OpenRecordset(VdbBerRec, SQLStr)
'For i% = 1 To 1000
'    If (VdbBerRec.EOF) Then
'        Exit For
'    End If
'
'    ActWert# = CheckNullLong(VdbBerRec!Wert) / 100#
'
'    AvAbrechnungsNr(AnzGef) = Trim(CheckNullStr(VdbBerRec!AbrechnungsNr))
'    AvHmPositionsNr(AnzGef) = Trim(CheckNullStr(VdbBerRec!HmPositionsNr))
'
'    ActBerechnung$ = Trim(CheckNullStr(VdbBerRec!Formelausdruck))
''    ActBerechnung$ = "(EK+25%)*0,95"
''    ActBerechnung$ = "EK+ST*0,4"
''    ActBerechnung$ = "EK-15%"
'
'    h$ = ActBerechnung$
'    ind% = InStr(h$, "+")
'    If (ind% <= 0) Then
'        ind% = InStr(h$, "-")
'    End If
'    If (ind% > 0) Then
'        ind% = InStr(h$, "*")
'        If (ind% > 1) Then
'            If (Mid$(h$, ind% - 1, 1) <> ")") Then
'                Do
'                    If (ind% = 1) Then
'                        ind% = 0
'                        Exit Do
'                    End If
'
'                    ind% = ind% - 1
'                    ch$ = Mid$(h$, ind%, 1)
'                    If (ch$ = "(") Then
'                        ind% = -1
'                        Exit Do
'                    End If
'                    If (InStr("+-*/", ch$) > 0) Or (ch$ = "@") Then
'                        Exit Do
'                    End If
'                Loop
'
'                If (ind% >= 0) Then
'                    ActBerechnung$ = Left$(h$, ind%) + "(" + Mid$(h$, ind% + 1) + ")"
'                End If
'            End If
'        End If
'    End If
'
'    If (VmDebugAktiv%) Then
'        With frmAvDebug.flxAvDebug(2)
'            ind% = AnzGef% * 3 + 1
'            .TextMatrix(1, ind%) = ActBerechnung$
'
'            iWert% = Abs(VdbBerRec!MehrwertSteuerFl)
'            If (iWert% = 1) Then
'                h2$ = "J"
'            Else
'                h2$ = "N"
'            End If
'            .TextMatrix(1, ind% + 1) = h2$
'
'            .TextMatrix(1, ind% + 2) = CStr(VdbBerRec!Kassenrabatt)
'
'            If (ActWert# <> 0#) Then
'                .TextMatrix(2, ind%) = Format(ActWert#, "0.00")
'            End If
'        End With
'    End If
'
'
'    Do
'        ind% = InStr(ActBerechnung$, "(")
'        If (ind% = 0) Then Exit Do
'
'
'        ind2% = InStr(ind% + 1, ActBerechnung$, ")")
'        If (ind% = 0) Then
'            MachBerechnungMDB% = False: Call DefErrPop: Exit Function
'        End If
'
'        ActBer2$ = Mid$(ActBerechnung$, ind% + 1, ind2% - ind% - 1)
'
'        KlammerPreis# = MachEinzelBerechnung#(ActBer2$)
'
'        ActBerechnung$ = Left$(ActBerechnung$, ind% - 1) + Format(KlammerPreis#, "    0.0000") + Mid$(ActBerechnung$, ind2% + 1)
'    Loop
'
'
'    ActPreis#(AnzGef%) = MachEinzelBerechnung#(ActBerechnung$)
'
'    If (VmDebugAktiv%) Then
'        ind% = AnzGef% * 3 + 1
'        frmAvDebug.flxAvDebug(2).TextMatrix(3, ind%) = Format(ActPreis#(AnzGef%), "0.00")
'    End If
'
'    VkBerechnung% = False
'    If (InStr(CheckNullStr(VdbBerRec!Formelausdruck), "VK") > 0) Then VkBerechnung% = True
'    If (Abs(VdbBerRec!MehrwertSteuerFl) > 0) And (VkBerechnung% = False) Then
'        ActPreis#(AnzGef%) = ActPreis#(AnzGef%) * (1# + (ActMwst) / 100#)
'    End If
'
'    If (VmDebugAktiv%) Then
'        ind% = AnzGef% * 3 + 1
'        frmAvDebug.flxAvDebug(2).TextMatrix(4, ind%) = Format(ActPreis#(AnzGef%), "0.00")
'    End If
'
'    iRabatt% = VdbBerRec!Kassenrabatt
'
'    If (VdbPznRec!HilfsmittelKz = 2) Or (Left$(TaxeRec!Zuzahlung, 1) = "9") Then
''        If (ActKasse% = 2) Then VmRech.Rabatt = "J"
'        If (ActKasse% = 2) Then iRabatt = 0
'    End If
'
'    ActPreis20#(AnzGef%) = ActPreis#(AnzGef%)
'
''    If (VmRech.Rabatt = "J") And (ActKasse% <> 2) Then
'    If (iRabatt = 0) Then
'        If (ActKasse% <> 2) Then
'            ActPreis20#(AnzGef%) = ActPreis20#(AnzGef%) / VmRabattFaktor
'        End If
'    ElseIf (iRabatt = 5) Then
'        If (TaxeRec!ApoRabattKz = 1) Then
'            iRabatt = 6
'        ElseIf (TaxeRec!ApoRabattKz = 2) Then
'            iRabatt = 7
'        Else
'            iRabatt = 4
'        End If
'    End If
'
'    ActAVrabatt%(AnzGef%) = Val(iRabatt)
'
''    If (VmRech.Rabatt = "N") Then
'    If (iRabatt = 1) Or (iRabatt = 2) Then
'        ActPreis#(AnzGef%) = ActPreis#(AnzGef%) * VmRabattFaktor
'
'        If (VmDebugAktiv%) Then
'            ind% = AnzGef% * 3 + 1
'            frmAvDebug.flxAvDebug(2).TextMatrix(5, ind%) = Format(ActPreis#(AnzGef%), "0.00")
'        End If
'    ElseIf (iRabatt = 3) Then
'        dPreis# = ActPreis20#(AnzGef%)
'        Select Case dPreis#
'            Case 0# To 52.46
'                dPreis# = dPreis# * 0.94
'            Case 52.47 To 54.8
'                dPreis# = 49.32
'            Case 54.81 To 820.22
'                dPreis# = dPreis# * 0.9
'            Case Else
'                diff# = dPreis# - 820.22
'                dPreis# = dPreis# * 0.9 - (diff# * 0.06)
'        End Select
'        ActPreis20#(AnzGef%) = dPreis#
'    ElseIf (iRabatt = 6) Then
'        ActPreis20#(AnzGef%) = ActPreis20#(AnzGef%) - 2#
'    ElseIf (iRabatt = 7) Then
'        ActPreis20#(AnzGef%) = ActPreis20#(AnzGef%) * 0.95
'    ElseIf (iRabatt = 8) Then
'        ActPreis#(AnzGef%) = ActPreis#(AnzGef%) * 1.05263
'
'        If (VmDebugAktiv%) Then
'            ind% = AnzGef% * 3 + 1
'            frmAvDebug.flxAvDebug(2).TextMatrix(5, ind%) = Format(ActPreis#(AnzGef%), "0.00")
'        End If
'    End If
'
'    If (VmDebugAktiv%) Then
'        If (AvAbrechnungsNr(AnzGef) <> "") Or (AvHmPositionsNr(AnzGef) <> "") Then
'            ind% = AnzGef% * 3 + 1
'
'            h2 = ""
'            If (AvAbrechnungsNr(AnzGef) <> "") Then
'                h2 = AvAbrechnungsNr(AnzGef)
'            End If
'            h2 = h2 + " / "
'            If (AvHmPositionsNr(AnzGef) <> "") Then
'                h2 = h2 + AvHmPositionsNr(AnzGef)
'            End If
'            frmAvDebug.flxAvDebug(2).TextMatrix(7, ind%) = h2$
'        End If
'    End If
'
'    If (VdbBedRec!RundungstypFl) Then
'        ActPreis#(AnzGef) = ActPreis#(AnzGef) * 10#
'
'        EinzelPfennig# = ActPreis#(AnzGef) - Fix(ActPreis#(AnzGef))
'        If (EinzelPfennig# < 0.25) Then
'            EinzelPfennig# = 0#
'        ElseIf (EinzelPfennig# < 0.75) Then
'            EinzelPfennig# = 0.5
'        Else
'            EinzelPfennig# = 1#
'        End If
'
'        ActPreis#(AnzGef) = Fix(ActPreis#(AnzGef)) + EinzelPfennig#
'        ActPreis#(AnzGef) = ActPreis#(AnzGef) / 10#
'    Else
'        ActPreis#(AnzGef) = CDbl(Int(Abs(ActPreis#(AnzGef)) * 100# + 0.501) / 100#)
'    End If
'
'
'    AnzGef% = i%
'
'    VdbBerRec.MoveNext
'Next i%
'
'If (AnzGef% = 0) Then
'    MachBerechnungMDB% = False: Call DefErrPop: Exit Function
'End If
'
'
'ActVmPreis# = 9999999.99
'If (VdbBedRec!WertErmittlungFl) And (ActPreis#(0) > 0#) Then
'    ActVmPreis# = ActPreis#(0)
'    Act20Preis# = ActPreis20#(0)
'    ActVmRabatt% = ActAVrabatt%(0)
'Else
'    For i% = 0 To (AnzGef% - 1)
'        If (ActPreis#(i%) > 0) And (ActPreis#(i%) < ActVmPreis#) Then
'            ActVmPreis# = ActPreis#(i%)
'            Act20Preis# = ActPreis20#(i%)
'            ActVmRabatt% = ActAVrabatt%(i%)
'        End If
'    Next i%
'End If
'
'
'If (ActVmPreis# = 9999999.99) Then
'    MachBerechnungMDB% = False: Call DefErrPop: Exit Function
'End If
'
'
''If (VdbBedRec!RundungstypFl) Then
''    ActVmPreis# = ActVmPreis# * 10#
''
''    EinzelPfennig# = ActVmPreis# - Fix(ActVmPreis#)
''    If (EinzelPfennig# < 0.25) Then
''        EinzelPfennig# = 0#
''    ElseIf (EinzelPfennig# < 0.75) Then
''        EinzelPfennig# = 0.5
''    Else
''        EinzelPfennig# = 1#
''    End If
''
''    ActVmPreis# = Fix(ActVmPreis#) + EinzelPfennig#
''    ActVmPreis# = ActVmPreis# / 10#
''Else
''    ActVmPreis# = CDbl(Int(Abs(ActVmPreis#) * 100# + 0.501) / 100#)
''End If
'
'If (VmDebugAktiv%) Then
'    frmAvDebug.flxAvDebug(2).TextMatrix(6, 1) = Format(ActVmPreis#, "0.00")
'End If
'
'MachBerechnungMDB% = True
'
'Call DefErrPop
'End Function


Sub ImporteOrAutIdem(IstImportModus%, Optional IstAutIdemModus% = True)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("ImporteOrAutIdem")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
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
'Dim i%, ind%, row%, AnzGilt%, AutIdemNeu%, Sonderregelnr%, NeueKbvNr%, iAnzZ%, iTabIndex%
Dim iAdd%, iAdd2%
'Dim spBreite&, MaxHe&
'Dim OrgAVP#, OrgKp#, avp#, avp2#, unten#, oben#, grenze#, BtmDerivatMenge#, FESTBETR#
'Dim s$, SQLStr$, kz$, OrgPZN$, PZN$, Titel$, SqlSonderregelNr$, h$, rabKZ$
'Dim iRec As New ADODB.Recordset
'Dim sTaxeRec As New ADODB.Recordset
'Dim tRec2 As New ADODB.Recordset
Dim row%, NeueKbvNr%, ind%, ind2%
Dim pzn$, h$

If (AktTabIndex% < 3) Then
    Call DefErrPop: Exit Sub
End If
    
row% = (AktTabIndex% - 3) \ 4
pzn$ = RezArtikel(row%).pzn
IdentPzn$ = pzn$

If (row >= AnzRezeptArtikel%) Then
    Call DefErrPop: Exit Sub
End If

SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + pzn$
'Set TaxeRec = TaxeDB.OpenRecordset(SQLStr$)
On Error Resume Next
TaxeRec.Close
Err.Clear
On Error GoTo DefErr
TaxeRec.Open SQLStr, taxeAdoDB.ActiveConn
If (TaxeRec.EOF) Then
    Call DefErrPop: Exit Sub
End If

If (PrivatRezept% = 0) And (AutIdemKbvNr& = 0) Then
    frmKkMatchcode.Show 1
    If (AutIdemKbvNr& > 0) Then
        NeueKbvNr% = True
    End If
End If

ReDim ArtLRec(0)
Dim iArtLRec%
iArtLRec = -1


If (IstImportModus) Then
'    Dim ast1 As StammdatenType
    Dim OrgPZN As String
    Dim OrgAvp As Double
    Dim AVP As Double
    Dim testavp As Double
    Dim fhImp As Integer
    Dim xLinks As Long
    Dim xRechts As Long
    Dim xMitte As Long
    Dim gefunden As Long
    Dim test As String
    Dim impPzn As String
    Dim Anzahl As Integer
    Dim erg As Integer
    Dim billiger As Boolean
    Dim billigerAbsolut As Boolean
    Dim NurBilligere As Boolean
    Dim SuchLagernde As Boolean
    Dim Preis As Double
    Dim ImportGruppeNr As Long
    Dim tRec As New ADODB.Recordset
    Dim SQL As String
    Dim vkt As Double
    Dim Durchlauf As Integer
    Dim IstBilliger As Boolean
    Dim i, j As Long
    Dim RoteLinieDa As Boolean
    Dim AusgangsPreis As Double
    Dim FB As Double
    Dim PZNDatei As String
    Dim fhout As Long
    Dim AnzVerfügbarkeitsabfrage As Integer
    Dim artPreis As Double
    Dim NoOriginal As Boolean
    Dim ImRezept As Boolean
    Dim ImportAlt As Boolean
    
    ImportAlt = False
    
    OrgPZN = pzn
'    If Val(Rezeptnummer) > 0 And Not Sprechstundenbedarf Then ImRezept = True
    ImRezept = True
    
    testavp = TaxeRec!vk / 100#
    AusgangsPreis = testavp
    'DrS 9.4.2014 Anzeige immer VK - Herstellerrabatt
    
    '10.6.19 DrS: günstigkeit eines Imports immer nach AVP unabhängig Herstellerrabatt
'    If Val(Rezeptnummer) > 0 And TaxeRec!HerstRabattKZ <> 0 Then
    If TaxeRec!HerstRabattKZ <> 0 Then
        AusgangsPreis = AusgangsPreis - HerstRabattBrutto(TaxeRec)
    End If
        
    FB = dCheckNull(TaxeRec!Rabwert_130a_2_SGB) / 100#
    If FB > 0 Then    'Impfstoffe: wenn Rabatt nach §130a günstiger als Herstellerrabatt, dann muss dieser abgezogen werden
        If (TaxeRec!vk / 100#) - FB < AusgangsPreis Then AusgangsPreis = (TaxeRec!vk / 100#) - FB
    End If

    If ImRezept Then
      FB = dCheckNull(TaxeRec!FESTBETRAG) / 100#
      If FB > 0 And FB < AusgangsPreis Then AusgangsPreis = FB    '1.4.49
    End If
    
    NoOriginal = True
    If TaxeRec!OriginalPZN <> 0 Then
      NoOriginal = False
      pzn = PznString(TaxeRec!OriginalPZN)
      If pzn <> OrgPZN Then
        SQLStr = "SELECT * FROM Taxe WHERE PZN=" + pzn
        On Error Resume Next
        TaxeRec.Close
        Err.Clear
        On Error GoTo DefErr
        TaxeRec.Open SQLStr, taxeAdoDB.ActiveConn
        If (TaxeRec.EOF) Then
            Call DefErrPop: Exit Sub
        End If
        
        Preis = TaxeRec!vk
        Preis = Preis / 100#
        
        If TaxeRec!HerstRabattKZ <> 0 Then
          Preis = Preis - HerstRabattBrutto(TaxeRec)
        End If
          
        FB = dCheckNull(TaxeRec!Rabwert_130a_2_SGB) / 100#
        If FB > 0 Then    'Impfstoffe: wenn Rabatt nach §130a günstiger als Herstellerrabatt, dann muss dieser abgezogen werden
          If (TaxeRec!vk / 100#) - FB < Preis Then Preis = (TaxeRec!vk / 100#) - FB
        End If
        
        FB = dCheckNull(TaxeRec!FESTBETRAG) / 100#
        If FB > 0 And FB < Preis Then Preis = FB    '1.4.49

      Else
            Preis = AusgangsPreis   ' testavp
      End If
      ImportGruppeNr = TaxeRec!ImportGruppeNr
    Else
      ImportGruppeNr = TaxeRec!ImportGruppeNr
    End If
    
    Call Taxe2ast(pzn$)    'OriginalArtikel. Wenn es keine OriginalPZN gibt, dann Ausgangsartikel
   
    OrgAvp = ast.AVP
    artPreis = OrgAvp
    '10.6.19 DrS: günstigkeit eines Imports immer nach AVP unabhängig Herstellerrabatt
'    If TaxeRec!HerstRabattKZ <> 0 And Val(Rezeptnummer) > 0 Then
    If TaxeRec!HerstRabattKZ <> 0 Then
      artPreis = FNX(artPreis - HerstRabattBrutto(TaxeRec))
    End If
    
    FB = dCheckNull(TaxeRec!Rabwert_130a_2_SGB) / 100#
    If FB > 0 Then    'Impfstoffe: wenn Rabatt nach §130a günstiger als Herstellerrabatt, dann muss dieser abgezogen werden
      If (TaxeRec!vk / 100#) - FB < artPreis Then artPreis = (TaxeRec!vk / 100#) - FB
    End If
        
    If ImRezept Then
      FB = dCheckNull(TaxeRec!FESTBETRAG) / 100#
      If FB > 0 And FB < artPreis Then artPreis = FB    '1.4.49
    End If
    
    ArtL.pzn = Val(ast.pzn)
    ArtL.kurz = ast.kurz
    ArtL.Preis = uFormat(artPreis, Space(Len(ArtL.Preis) - 4) + "0.00")
    Anzahl = 1
    If TaxeRec!OriginalPZN > 0 Then
      ArtL.kz1 = "o"
    Else
      ArtL.kz1 = " "    'Importgruppe, aber keine OriginalPZN
    End If
    ArtL.kz2 = ".-"
    ArtL.GhVerfügbar = 0
    ArtL.rabKZ = " "
    If (Val(ast.pzn) = Val(OrgPZN)) Then ArtL.AusgangsArtikel = 1 Else ArtL.AusgangsArtikel = 0
    If IstRabattAutIdem(Val(ArtL.pzn)) Then ArtL.rabKZ = "r"
    ArtL.RoteLinie = 0
    
    ArtL.ArtStatus = TaxeRec!ArtStatus
    
    iArtLRec = iArtLRec + 1
    ReDim Preserve ArtLRec(iArtLRec)
    ArtLRec(iArtLRec) = ArtL
'    PutDataRecord dbArtL, CLng(anzahl)
'
    pzn = ast.pzn
'
    For Durchlauf = 1 To 2
      If Durchlauf = 1 And Not ImportAlt And ImportGruppeNr > 0 Then
        SQLStr = "SELECT * FROM Taxe WHERE ImportgruppeNr=" + CStr(ImportGruppeNr)
      ElseIf (Durchlauf = 2 And Not ImportAlt) Or (Durchlauf = 1 And ImportGruppeNr = 0) Then
        If Anzahl > 1 Then Exit For

        SQLStr = "SELECT * FROM Taxe WHERE OriginalPZN=" + pzn
      Else
        'alte Variante und schon im ersten Durchlauf erledigt
        Exit For
      End If
      
    On Error Resume Next
    TaxeRec.Close
    Err.Clear
    On Error GoTo DefErr
    TaxeRec.Open SQLStr, taxeAdoDB.ActiveConn
      
      Do While Not TaxeRec.EOF
        If TaxeRec!pzn <> Val(pzn) Then
          Anzahl = Anzahl + 1
          Call Taxe2ast(PznString(TaxeRec!pzn))
          AVP = ast.AVP
          artPreis = AVP
          'DrS 9.4.2014 Anzeige immer VK - Herstellerrabatt
          '10.6.19 DrS: günstigkeit eines Imports immer nach AVP unabhängig Herstellerrabatt --> artpreis <> AVP
          If TaxeRec!HerstRabattKZ <> 0 Then
            artPreis = FNX(artPreis - HerstRabattBrutto(TaxeRec))
          End If
            If ImRezept Then
              FB = dCheckNull(TaxeRec!FESTBETRAG) / 100#
              If FB > 0 And FB < artPreis Then artPreis = FB '1.4.49
            End If
          ArtL.pzn = Val(ast.pzn)
          ArtL.kurz = ast.kurz
          ArtL.Preis = uFormat(artPreis, Space(Len(ArtL.Preis) - 4) + "0.00")
          ArtL.kz1 = " "
          If ArtL.pzn = Val(pzn) Then
            ArtL.kz2 = "="
          Else
            IstBilliger = False
            If Not NoOriginal Then
              IstBilliger = IstGünstigerImport(artPreis, Preis)
            End If
            If IstBilliger Then
              ArtL.kz2 = " "
              billiger = True
              billigerAbsolut = True
            Else
              ArtL.kz2 = ".-"
              If Not NoOriginal And AVP < OrgAvp Then billigerAbsolut = True
            End If


            'If (avp < orgavp - 1) And (avp < (orgavp * 0.9)) Then
    '        If (avp <= (orgavp * 0.85) Or avp <= (orgavp - 15)) Then
    '          artl.kz2 = " "
    '          billiger = True
    '          billigerAbsolut = True
    '        Else
    '          artl.kz2 = ".-"
    '          If avp < orgavp Then billigerAbsolut = True
    '        End If
          End If
          ArtL.GhVerfügbar = 0
          If Val(ast.pzn) = Val(OrgPZN) Then ArtL.AusgangsArtikel = 1 Else ArtL.AusgangsArtikel = 0
          ArtL.rabKZ = " "
          If IstRabattAutIdem(Val(ArtL.pzn)) Then ArtL.rabKZ = "r"
          ArtL.RoteLinie = 0
          
          ArtL.ArtStatus = TaxeRec!ArtStatus
    

            iArtLRec = iArtLRec + 1
            ReDim Preserve ArtLRec(iArtLRec)
            ArtLRec(iArtLRec) = ArtL
'          PutDataRecord dbArtL, CLng(anzahl)
        End If
        TaxeRec.MoveNext
      Loop
    Next Durchlauf
'
    If billigerAbsolut = False Then
'      GetDataRecord dbArtL, 1
'      ArtL.kz2 = "  "
'      PutDataRecord dbArtL, 1
      ArtLRec(0).kz2 = "  "
    End If
    If Anzahl > 1 Then
    '    Rabkz, Preis, Kz1 und 1.Stelle Kz2, Kurz
'      Call SortOpenFile(dbArtL.fh, dbArtL.RecLen, 1, CLng(anzahl), "-46,1,33,10,-43,2,5,28", "")
        For i = 1 To Anzahl
            With ArtLRec(i - 1)
                h = Format(999 - Asc(.rabKZ), "000") + .Preis + Format(999 - Asc(.kz1), "000") + Format(999 - Asc(Left(.kz2, 1)), "000") + .kurz '+ vbTab + Str(i)
                .Sort = h
            End With
        Next i
    
        Dim MinInd%
        For i = 1 To Anzahl
            MinInd = i
            For j = (i + 1) To Anzahl
                If (ArtLRec(j - 1).Sort < ArtLRec(MinInd - 1).Sort) Then
                    MinInd = j
                End If
            Next j
            ArtL = ArtLRec(i - 1)
            ArtLRec(i - 1) = ArtLRec(MinInd - 1)
            ArtLRec(MinInd - 1) = ArtL
    '            h = .List(i - 1)
    '            ind = InStr(h, vbTab)
    '            If (ind > 0) Then
    '                ind2 = Val(Mid(h, ind + 1))
    '                ArtL = ArtLRec(i - 1)
    '                ArtLRec(i - 1) = ArtLRec(ind2 - 1)
    '                ArtLRec(ind2 - 1) = ArtL
    '            End If
        Next i

        'RabKZ, Preis, Kurz
'      ImporteNeu = True
    End If
    If Anzahl > 0 Then
'      If AutoVerfügbarkeitsabfrage Then
'        PZNDatei = CurDir + "\P" + Format(Now, "nnss") + CStr(ComputerNummer) + Mid("LR", BSSeite + 1, 1) + ".$$$"
'        fhout = fopen(PZNDatei, "o")
'      End If

      For i = 1 To Anzahl
'        GetDataRecord dbArtL, i
            ArtL = ArtLRec(i - 1)
        If ArtL.rabKZ = "r" Then
'          If fhout > 0 And AutoVerfügbarkeitsabfrage Then
'            Print #fhout, PznString(ArtL.PZN)
'            AnzVerfügbarkeitsabfrage = AnzVerfügbarkeitsabfrage + 1
'          End If
        ElseIf Not RoteLinieDa Then

          'If xVal(artl.Preis) > AusgangsPreis Or artl.kz2 = ".-" Then
          'Günstigkeit eines Imports hat nichts mit der Preislinie zu tun! Wenn kein Rabattvertrag, darf sogar teuerster Import abgegeben werden
          'wenn verordnet und es keine nicht-Importe in der Alt+G-Linie gibt
          If xVal(ArtL.Preis) > xVal(uFormat(AusgangsPreis, "########0.00")) Then

            ArtL.RoteLinie = 1
'            PutDataRecord dbArtL, i
            ArtLRec(i - 1) = ArtL
            RoteLinieDa = True
          Else
'            If fhout > 0 And AutoVerfügbarkeitsabfrage Then
'              Print #fhout, PznString(ArtL.PZN)
'              AnzVerfügbarkeitsabfrage = AnzVerfügbarkeitsabfrage + 1
'            End If
          End If
        Else
          Exit For
        End If
      Next i
    End If
    If Anzahl > 1 And NurBilligere Then    '1.2.64
      If billiger Then
      Else
'        ImporteNeu = False
        Anzahl = 0
     End If
    End If
    If Anzahl <= 1 Then
'      erg = dbClose(dbArtL)
'      Kill RTrim(dbArtL.FileName)
    Else
'      ArtikelListenArt = LISTE_IMPORT
    End If
Else
    ''''''''''''''''''''''''''''
Dim satz As Long
Dim recno As Long
'Dim menge As Single
'Dim Füllmenge As Single
Dim AnyGrün As Boolean
Dim billigst As Double
Dim teuerst As Double
Dim Anz4billigst As Integer
Dim ViertBilligstPreis As Double
Dim PreisVorher As Double
Dim grenze As Double
Dim OriginalKP As Double
Dim OriginalAVP As Double
'Dim art1 As StammdatenType
Dim AnzGilt As Integer
Dim t(3) As Double
Dim b(3) As Double
Dim zw As Double
'Dim SQLStr As String
Dim iSqlStr As String, tSQLStr As String
Dim TaxeTmpRec As New ADODB.Recordset
Dim tmpRec As New ADODB.Recordset
Dim Sonderregelnr As Long
Dim PznSubst As Boolean
Dim WirkStoff As String
Dim PG As String
Dim PZNok As Boolean
Dim UnterFB As Boolean
Dim raus As Boolean
Dim asatz As Long
Dim TaxeVK As Double
Dim GenImpZähler As Integer
Dim ok As Boolean
Dim ret As Boolean
Dim OrgPznIndik As String
Dim PznIndik As String
Dim h2 As String
'Dim h As String
'Dim ind As Integer
Dim SqlSonderregelNr As String
Dim BtmDerivatMenge As Double
'Dim Preis As Double
Dim AktHMNr As String
Dim NoGenerika As Boolean
'Dim PZNDatei As String
'Dim fhout As Long
'Dim AnzVerfügbarkeitsabfrage As Integer
'Dim RoteLinieDa As Boolean
'Dim AusgangsPreis As Double
Dim AnyRVDa As Boolean
Dim ImportGr As Long
Dim nurImporte As Boolean
Dim OrgNgr As String
Dim OrgMenge As Long  'String
Dim iOk As Boolean

    nurImporte = True
    
    Anzahl = 0
    ImRezept = True
    
    OrgNgr = TaxeRec!NGrösse
    
'    OrgMenge = ""
'    If (TaxeRec!BtmKz) Then
'        OrgMenge = CheckNullStr(TaxeRec!menge)
'    End If
    OrgMenge = 0
    If (TaxeRec!BtmKz) Then
        OrgMenge = CheckNullLong(TaxeRec!StdMenge)
    End If
    
    ImportGr = TaxeRec!ImportGruppeNr
'    If raus Then
'        Call Message("zu diesem Artikel konnten keine Generika gefunden werden!", vbExclamation)
'      Call DefErrPop
'      Exit Sub
'    End If
    Call Taxe2ast(pzn)
    
      OriginalKP = ast.kp
      OriginalAVP = ast.AVP
      If ast.kp > 0 And ast.kp < OriginalAVP Then OriginalAVP = ast.kp
      'DrS 28.1.2014: bei AutIdems Reihung und Anzeige mit AVP. Bei Importen weiterhin abzügl. Herstellerrabatt
      'Rahmenvertrag ab 1.7.2019: Herstellerrabatt abziehen
        If TaxeRec!HerstRabattKZ <> 0 Then
          OriginalAVP = OriginalAVP - HerstRabattBrutto(TaxeRec)
        End If
    
    If Not TaxeRec.EOF Then
'      GenerikaAusgangsArtikel = Trim(TaxeRec!name) + " " + Trim(TaxeRec!menge) + " " + Trim(TaxeRec!einheit) + " (" + Trim(TaxeRec!HerstellerKB) + ")"
    End If
    BtmDerivatMenge# = 0
    If (BtmDerivatMengeIgnorieren) Then
    Else
        On Error Resume Next
        BtmDerivatMenge# = TaxeRec!BtmDerivatMenge
        If (BtmDerivatMenge >= 0.01) Then
            SQLStr$ = "SELECT * FROM M2Inhalt WHERE M2Nr = " + Mid(TaxeRec!M2, 2, 9)
            On Error Resume Next
            AutIdemRec.Close
            Err.Clear
            On Error GoTo DefErr
            AutIdemRec.Open SQLStr, AutIdemDB.ActiveConn
            If Not (AutIdemRec.EOF) Then
                If (InStr(UCase(CheckNullStr(AutIdemRec!Inhalt)), "HYDROMORPHON") > 0) Then
                    BtmDerivatMenge = 0
                End If
            End If
            AutIdemRec.Close
        End If
        On Error GoTo DefErr
    End If
    
    SQLStr = ""
      If TaxeRec!AutIdemNeu > 0 Then
        PznSubst = True
        SQLStr = "SELECT * FROM PZNSUBST WHERE PZN = " + CStr(TaxeRec!pzn)
        On Error Resume Next
        AutIdemRec.Close
        On Error GoTo DefErr
        AutIdemRec.Open SQLStr, AutIdemDB.ActiveConn
        SQLStr = ""
        If Not AutIdemRec.EOF Then
          AutIdemRec.MoveFirst
          i = 1
          SQLStr = "SELECT * FROM TAXE WHERE (PZN = "
          Do While Not (AutIdemRec.EOF)
            If (i > 1) Then SQLStr = SQLStr + " OR PZN = "
            SQLStr = SQLStr + CStr(AutIdemRec!Substpzn)
            AutIdemRec.MoveNext
            i = i + 1
            If i > 250 Then Exit Do
          Loop
          If (BtmDerivatMenge >= 0.01) Then
              SQLStr = SQLStr + ") AND (BtmDerivatMenge=" + Str(BtmDerivatMenge)
          End If
          
        End If
        AutIdemRec.Close
        Set AutIdemRec = Nothing
        If SQLStr > "" Then
'          If AktKundensatz > 0 And Newline Then BezugSQL = "SELECT TOP 2 * FROM Bezug WHERE Kundennr=" + CStr(AktKundensatz) + " AND " + Mid(SQLStr, 26) + ") ORDER BY Datum DESC"
          SQLStr = SQLStr + ") ORDER BY VK,NAME"
        End If
      End If
      If TaxeRec!AutIdemNeu = 0 Then
'       If ImporteNeu(ast.PZN) Then
        Call ImporteOrAutIdem(True)
          Anzahl = 2
          Call DefErrPop
          Exit Sub
'        End If
      End If
    
    If SQLStr > "" Then
        On Error Resume Next
        TaxeTmpRec.Close
        Err.Clear
        On Error GoTo DefErr
        TaxeTmpRec.Open SQLStr, taxeAdoDB.ActiveConn
    Else        'If art <> 0 Then
        Call MsgBox("zu diesem Artikel konnten keine Generika gefunden werden!", vbExclamation)
      Call DefErrPop
      Exit Sub
    End If
    
'    If PznSubst Then
'      ArtikelListenArt = LISTE_AUTIDEM
'   End If
    
    If Not (TaxeTmpRec.EOF) Then
        TaxeTmpRec.MoveLast
    End If
    If (TaxeTmpRec.RecordCount > 1) Then
      TaxeTmpRec.MoveFirst
    End If
    
    For GenImpZähler = 1 To 2
      '1 ... normale Generikasuche
      
      '2 Importe:
      If GenImpZähler = 2 Then
        If TaxeRec!OriginalPZN = 0 And TaxeRec!ImportGruppeNr = 0 Then Exit For
        iSqlStr = "SELECT * FROM TAXE WHERE (PZN = "
        i = 0
        If TaxeRec!ImportGruppeNr > 0 Then
          tSQLStr = "SELECT * FROM Taxe WHERE ImportgruppeNr=" + CStr(TaxeRec!ImportGruppeNr)
        
            On Error Resume Next
            TaxeTmpRec.Close
            Err.Clear
            On Error GoTo DefErr
            TaxeTmpRec.Open tSQLStr, taxeAdoDB.ActiveConn
          Do While Not TaxeTmpRec.EOF
            If InStr(SQLStr, "PZN = " + CStr(TaxeTmpRec!pzn)) = 0 Then
              If i > 0 Then iSqlStr = iSqlStr + " OR PZN = "
              iSqlStr = iSqlStr + CStr(TaxeTmpRec!pzn)
              i = i + 1
            End If
            TaxeTmpRec.MoveNext
          Loop
        End If
        If TaxeRec!OriginalPZN > 0 Then
          tSQLStr = "SELECT * FROM Taxe WHERE OriginalPZN=" + CStr(TaxeRec!OriginalPZN)
          
            On Error Resume Next
            TaxeTmpRec.Close
            Err.Clear
            On Error GoTo DefErr
            TaxeTmpRec.Open tSQLStr, taxeAdoDB.ActiveConn
          
          Do While Not TaxeTmpRec.EOF
            If InStr(SQLStr, "PZN = " + CStr(TaxeTmpRec!pzn)) = 0 And InStr(iSqlStr, "PZN = " + CStr(TaxeTmpRec!pzn)) = 0 Then
              If i > 0 Then iSqlStr = iSqlStr + " OR PZN = "
              iSqlStr = iSqlStr + CStr(TaxeTmpRec!pzn)
              i = i + 1
            End If
            TaxeTmpRec.MoveNext
          Loop
        End If
        If i = 0 Then Exit For   'keine Importe, die nicht eh schon in der Liste wären
        iSqlStr = iSqlStr + ")"
        
            On Error Resume Next
            TaxeTmpRec.Close
            Err.Clear
            On Error GoTo DefErr
            TaxeTmpRec.Open iSqlStr, taxeAdoDB.ActiveConn
      End If
      
      Do While Not TaxeTmpRec.EOF
        'If TaxeTmpRec!ArtStatus <> "S" Then      'außer Handel-Artikel sollen in die Liste hinein! (RV2019)
          If Not PznSubst Then
            OrgPznIndik$ = GetPznIndik(TaxeRec!pzn)
            ret = True
            If (OrgPznIndik$ > "") And TaxeTmpRec!pzn <> TaxeRec!pzn Then
          
              PznIndik$ = GetPznIndik(TaxeTmpRec!pzn)
              If (PznIndik$ <> "") Then PznIndik$ = "," + PznIndik$
              
              h$ = OrgPznIndik$
              Do
                ind% = InStr(h$, ",")
                If (ind% = 0) Then
                  Exit Do
                Else
                  h2$ = "," + Left$(h$, ind%)
                  h$ = Mid$(h$, ind% + 1)
                  If (InStr(PznIndik$, h2$) = 0) Then
                    ret = False
                    Exit Do
                  End If
                End If
              Loop
            End If
'            If (OrgMenge <> "") Then
'                If TaxeTmpRec!menge <> OrgMenge Then ret = False
            If (OrgMenge <> 0) Then
                If TaxeTmpRec!StdMenge <> OrgMenge Then ret = False
            Else
                If TaxeTmpRec!NGrösse <> OrgNgr Then ret = False
            End If
  
            If ret Then
              ArtL.pzn = TaxeTmpRec!pzn
              If ArtL.pzn = ast.pzn Then ArtL.AusgangsArtikel = 1 Else ArtL.AusgangsArtikel = 0
              ArtL.kurz = TaxeTmpRec!Name
              TaxeVK = TaxeTmpRec!vk / 100#
              Preis = TaxeTmpRec!vk
              ArtL.Preis = uFormat(Preis / 100#, Space(Len(ArtL.Preis) - 4) + "0.00")
              If ArtL.AusgangsArtikel = 1 Then AusgangsPreis = xVal(ArtL.Preis)
              If TaxeTmpRec!ImportGruppeNr = 0 Then nurImporte = False
              If (ImportGr > 0) And (TaxeTmpRec!ImportGruppeNr <> ImportGr) Then nurImporte = False
              If ImportGr = 0 Or TaxeTmpRec!ImportGruppeNr = ImportGr Then ArtL.ImportGruppeok = "1" Else ArtL.ImportGruppeok = "9"
              Anzahl = Anzahl + 1
              ArtL.kz1 = " "
              If TaxeVK > 0 Then
                If OriginalKP > 0 Then
                  If TaxeVK > OriginalKP Then
                    ArtL.kz1 = "+"
                  Else
                    ArtL.kz1 = "-"
                  End If
                Else
                  ArtL.kz1 = " "
                End If
              End If
              ArtL.GhVerfügbar = 0
              ArtL.RoteLinie = 0
              
              ArtL.ArtStatus = TaxeTmpRec!ArtStatus
    

                iArtLRec = iArtLRec + 1
                ReDim Preserve ArtLRec(iArtLRec)
                ArtLRec(iArtLRec) = ArtL
'              PutDataRecord dbArtL, CLng(anzahl)
            End If
          Else
'           If (OrgMenge <> "") Then
'                iOk = (TaxeTmpRec!menge = OrgMenge)
            If (OrgMenge <> 0) Then
                iOk = (TaxeTmpRec!StdMenge = OrgMenge)
            Else
                iOk = (TaxeTmpRec!NGrösse = OrgNgr)
            End If
'              If TaxeTmpRec!NGrösse = OrgNgr Then
              If (iOk) Then
                ArtL.pzn = TaxeTmpRec!pzn
                If ArtL.pzn = ast.pzn Then ArtL.AusgangsArtikel = 1 Else ArtL.AusgangsArtikel = 0
                ArtL.kurz = TaxeTmpRec!Name
                Preis = TaxeTmpRec!vk
                TaxeVK = TaxeTmpRec!vk / 100#
          'DrS 28.1.2014: bei AutIdems Reihung und Anzeige mit AVP. Bei Importen weiterhin abzügl. Herstellerrabatt
          'neuer Rahmenvertrag 1.7.2019: Herstellerrabatt abziehen
                  If TaxeTmpRec!HerstRabattKZ <> 0 Then
                    Preis = Preis - (HerstRabattBrutto(TaxeTmpRec) * 100#)
                  End If
                ArtL.Preis = uFormat(Preis / 100#, Space(Len(ArtL.Preis) - 4) + "0.00")
                If ArtL.AusgangsArtikel = 1 Then AusgangsPreis = xVal(ArtL.Preis)
                If ImportGr = 0 Or TaxeTmpRec!ImportGruppeNr = ImportGr Then ArtL.ImportGruppeok = "1" Else ArtL.ImportGruppeok = "9"
                If TaxeTmpRec!ImportGruppeNr = 0 Then nurImporte = False
                If (ImportGr > 0) And (TaxeTmpRec!ImportGruppeNr <> ImportGr) Then nurImporte = False
                Anzahl = Anzahl + 1
                ArtL.kz1 = " "
                If PznSubst Then
                  If OriginalKP > 0 And TaxeVK > 0 Then
                    If TaxeVK > OriginalKP Then
                      ArtL.kz1 = "+"
                    ElseIf TaxeVK < OriginalKP Then
                      ArtL.kz1 = "-"
                    End If
                  End If
                  ArtL.kz2 = " "
                  If Val(ArtL.Preis) > grenze And grenze > 0 Then
                    ArtL.kz2 = "."
                  Else
    '                If SuchPzn And ArtL.PZN = Val(art1.PZN) Then
    '                  PZNok = True
    '                End If
                  End If
                End If
                ArtL.GhVerfügbar = 0
                ArtL.RoteLinie = 0
                
'                ArtL.ArtStatus = sCheckNull(TaxeTmpRec!ArtStatus)
                ArtL.ArtStatus = TaxeTmpRec!ArtStatus
    

                    iArtLRec = iArtLRec + 1
                    ReDim Preserve ArtLRec(iArtLRec)
                    ArtLRec(iArtLRec) = ArtL
    '            PutDataRecord dbArtL, CLng(anzahl)
              End If
          End If
          If Anzahl > 500 Then Exit Do
        'End If
        TaxeTmpRec.MoveNext
      Loop
    Next GenImpZähler
    
    If Anzahl = 0 Then
'      erg = dbClose(dbArtL)
'      If SuchPzn Then
'        SuchPzn = False
'      ElseIf Not KeineMeldung Then
        Call MsgBox("zu diesem Artikel konnten keine Generika gefunden werden!", vbExclamation)
'      End If
    Else
      If nurImporte Then    'nur Importe: es gelten nur Regeln für Importrelevanten Markt: Rabattvertrag und alles was nicht teurer ist als verordneter Artikel
'        erg = dbClose(dbArtL)
'        If ImporteNeu(ast.PZN) Then
        Call ImporteOrAutIdem(True)
          Anzahl = 2
          Call DefErrPop
          Exit Sub
'        End If
      End If
      Sonderregelnr = 0
      SqlSonderregelNr = ""
      If AutIdemKbvNr > 0 Then
        SQLStr$ = "SELECT * FROM AutIdemSonderregelKrankenkasse WHERE KbvNr = " + CStr(AutIdemKbvNr)
        
       tmpRec.Open SQLStr, AutIdemDB.ActiveConn
        If Not (tmpRec.EOF) Then
          tmpRec.MoveFirst
          Do While Not tmpRec.EOF
            Sonderregelnr = dCheckNull(tmpRec!AutIdemSonderregelNr)
            If (Sonderregelnr > 0) Then
              If (SqlSonderregelNr = "") Then
                  SqlSonderregelNr = "SELECT * FROM AutIdemSonderregelArtikel WHERE ("
              Else
                  SqlSonderregelNr = SqlSonderregelNr + " OR"
              End If
              SqlSonderregelNr = SqlSonderregelNr + " AutIdemSonderregelNr = " + Format(dCheckNull(tmpRec!AutIdemSonderregelNr), "##0")
            End If
            tmpRec.MoveNext
          Loop
          tmpRec.Close
          Set tmpRec = Nothing
        End If
      End If
      For satz = 1 To Anzahl
'        GetDataRecord dbArtL, satz
            ArtL = ArtLRec(satz - 1)
        ArtL.rabKZ = " "
        If AutIdemKbvNr > 0 And (SqlSonderregelNr > "") Then
          SQLStr$ = SqlSonderregelNr + ")"
          SQLStr$ = SQLStr$ + " AND NrKz='P' AND ZoNr=" + CStr(ArtL.pzn)
          tmpRec.Open SQLStr, AutIdemDB.ActiveConn
          If Not (tmpRec.EOF) Then
            'artl.kz2 = Left(artl.kz2, 1) + "r"
            ArtL.rabKZ = "r"
            AnyGrün = True
          End If
          tmpRec.Close
          Set tmpRec = Nothing
        End If
'        PutDataRecord dbArtL, satz
            ArtLRec(satz - 1) = ArtL
      Next satz
      
            '10.6.19 DrS: Importgruppe doch egal für Günstigkeite ("generischer Markt")
    
    '    'Importgr,Preis, Kurz
'        Call SortOpenFile(dbArtL.fh, dbArtL.RecLen, 1, CLng(anzahl), "33,10,5,28", "")
        For i = 1 To Anzahl
            With ArtLRec(i - 1)
                h = .Preis + .kurz '+ vbTab + Str(i)
                .Sort = h
            End With
        Next i
    
'        Dim MinInd%
        For i = 1 To Anzahl
            MinInd = i
            For j = (i + 1) To Anzahl
                If (ArtLRec(j - 1).Sort < ArtLRec(MinInd - 1).Sort) Then
                    MinInd = j
                End If
            Next j
            ArtL = ArtLRec(i - 1)
            ArtLRec(i - 1) = ArtLRec(MinInd - 1)
            ArtLRec(MinInd - 1) = ArtL
    '            h = .List(i - 1)
    '            ind = InStr(h, vbTab)
    '            If (ind > 0) Then
    '                ind2 = Val(Mid(h, ind + 1))
    '                ArtL = ArtLRec(i - 1)
    '                ArtLRec(i - 1) = ArtLRec(ind2 - 1)
    '                ArtLRec(ind2 - 1) = ArtL
    '            End If
        Next i
                
        
        Anz4billigst = 0
        PreisVorher = 0
        
        For satz = 1 To Anzahl
'          GetDataRecord dbArtL, satz
          ArtL = ArtLRec(satz - 1)
          '10.6.19 DrS: Importgruppe doch egal für Günstigkeite ("generischer Markt")
    '      If artl.ImportGruppeok <> "1" Then Exit For   'wenn Ausgangsartikel Imp.Gr. hat, müssen 4 billigste aus gleicher Imp.Gr. ermittelt werden
    '      '                                        'wenn nicht, haben eh alle DS ImoprgruppeOk=1
          '7.6.2019 DrS: bis 4 Artikel zählen und dann noch alle, die gleichen Preis wie der 4. haben!!!
          'vorher 4 billigste Preise zählen
          
          If (ArtL.ArtStatus <> "S") Then
            Anz4billigst = Anz4billigst + 1
'          Else
'            Beep
          End If
          PreisVorher = Val(ArtL.Preis)
          If Anz4billigst = 4 Then
            ViertBilligstPreis = Val(ArtL.Preis)
    '        satz = satz + 1 'für Schleife nachher wichtig! Bei diesem Artikel darf Importgruppeok nocht gelöscht werden
            Exit For
          End If
          
          
        Next satz
        If ViertBilligstPreis = 0 And Anz4billigst > 0 Then ViertBilligstPreis = PreisVorher
        
        'RabKZ, ImportGrOk, Preis, Kurz
 '       Call SortOpenFile(dbArtL.fh, dbArtL.RecLen, 1, CLng(anzahl), "-46,1,33,10,5,28", "")
        For i = 1 To Anzahl
            With ArtLRec(i - 1)
                h = Format(999 - Asc(.rabKZ), "000") + .Preis + .kurz '+ vbTab + Str(i)
                .Sort = h
            End With
        Next i
    
'        Dim MinInd%
        For i = 1 To Anzahl
            MinInd = i
            For j = (i + 1) To Anzahl
                If (ArtLRec(j - 1).Sort < ArtLRec(MinInd - 1).Sort) Then
                    MinInd = j
                End If
            Next j
            ArtL = ArtLRec(i - 1)
            ArtLRec(i - 1) = ArtLRec(MinInd - 1)
            ArtLRec(MinInd - 1) = ArtL
    '            h = .List(i - 1)
    '            ind = InStr(h, vbTab)
    '            If (ind > 0) Then
    '                ind2 = Val(Mid(h, ind + 1))
    '                ArtL = ArtLRec(i - 1)
    '                ArtLRec(i - 1) = ArtLRec(ind2 - 1)
    '                ArtLRec(ind2 - 1) = ArtL
    '            End If
        Next i
      
'      If ArtikelListenArt = LISTE_AUTIDEM And ImRezept And anzahl > 0 And AutoVerfügbarkeitsabfrage Then
'        PZNDatei = CurDir + "\P" + Format(Now, "nnss") + CStr(ComputerNummer) + Mid("LR", BSSeite + 1, 1) + ".$$$"
'        fhout = fopen(PZNDatei, "o")
'      End If
      
'      If SuchPzn Then
'        SuchPzn = (AusgangsPreis <= ViertBilligstPreis)
'        erg = dbClose(dbArtL)
'        Call DefErrPop
'        Exit Sub
'      End If
    'ganz oben Rabattvertragsartikel, dann die vier billigsten (der gleichen Importruppe), dann alle anderen
      For satz = 1 To Anzahl
'        GetDataRecord dbArtL, satz
        ArtL = ArtLRec(satz - 1)
        ArtL.kz2 = " "
        
        If Val(ArtL.Preis) <= ViertBilligstPreis Then 'And artl.ImportGruppeok = "1" Then
'          If SuchPzn And ArtL.PZN = Val(art.PZN) Then PZNok = True
          If Not RoteLinieDa And ArtL.rabKZ <> "r" And xVal(ArtL.Preis) > xVal(uFormat(AusgangsPreis, "#######0.00")) Then
            RoteLinieDa = True
            ArtL.RoteLinie = 1
          ElseIf Not AnyRVDa Then
'            If fhout > 0 And AutoVerfügbarkeitsabfrage Then
'              Print #fhout, PznString(ArtL.PZN)
'              AnzVerfügbarkeitsabfrage = AnzVerfügbarkeitsabfrage + 1
'            End If
          End If
        Else
          ArtL.kz2 = "."
          If Not RoteLinieDa And ArtL.rabKZ <> "r" Then
            RoteLinieDa = True
            ArtL.RoteLinie = 1
          End If
        End If
        If ArtL.rabKZ = "r" Then
          'artl.kz2 = Left(artl.kz2, 1) + "r"
          AnyRVDa = True
'          If fhout > 0 And AutoVerfügbarkeitsabfrage Then
'            Print #fhout, PznString(ArtL.PZN)
'            AnzVerfügbarkeitsabfrage = AnzVerfügbarkeitsabfrage + 1
'          End If
        ElseIf Not RoteLinieDa Then
          If Val(ArtL.Preis) > ViertBilligstPreis Then ' Or artl.ImportGruppeok <> "1" Then
            RoteLinieDa = True
            ArtL.RoteLinie = 1
          End If
        End If
'        If SuchPzn And Val(ArtL.Preis) <= OriginalKP And ArtL.PZN <> Val(art1.PZN) Then UnterFB = True
    
'        If ArtikelListenArt <> LISTE_AUTIDEM Or Val(Rezeptnummer) = 0 Then
'          dbAst.ActiveIndex = 1
'          asatz = db_Seek(dbAst, "=", PznString(ArtL.PZN))
'          If asatz > 0 Then
'            GetDataRecord dbAst, asatz
'            ArtL.Preis = uFormat(CVD(ast.avp), Space(Len(ArtL.Preis) - 4) + "0.00")
'          End If
'        End If
'        PutDataRecord dbArtL, satz
            ArtLRec(satz - 1) = ArtL
      Next satz
      
      If Not AnyGrün Then
'        If CheckRabattAutIdem(Val(LastPzn), 1) Then
        If CheckRabattAutIdem(Val(pzn), 1) Then
          'GenerikaRabAIAndereDAR = True
        End If
      End If
      
'      If SuchPzn Then
'        SuchPzn = Not PZNok
'        erg = dbClose(dbArtL)
'      End If
    End If
    
    TaxeTmpRec.Close
    Set TaxeTmpRec = Nothing
    
'    If AnzVerfügbarkeitsabfrage > 0 And AutoVerfügbarkeitsabfrage Then
'      Close #fhout
'      Call ImportVerfügbar(False, PZNDatei)
'    End If
    
    ''''''''''''''''''''''''''''
End If
    
    ReDim Preserve m(Anzahl - 1)
    ReDim Preserve amatch(Anzahl - 1)
    Call BelegeSatz
    Call BelegeWork(IstImportModus)
    
With frmImportOrIdent
    .Caption = IIf(IstImportModus, "Auswahl IMPORTE", "Auswahl AutIdem")
    .cmdF2.Enabled = (IstImportModus% = 0)

    If (para.Newline) Then
        iAdd = wpara.NlFlexBackY
        iAdd2 = wpara.NlCaptionY

        With frmImportOrIdent.flxEdit
'            .ScrollBars = flexScrollBarNone
            .BorderStyle = 0
            .Width = .Width - 90
            .Height = .Height - 75
            .GridLines = flexGridFlat
            .GridLinesFixed = flexGridFlat
            .GridColorFixed = .GridColor
            .BackColor = wpara.nlFlexBackColor    'vbWhite
            .BackColorBkg = wpara.nlFlexBackColor    'vbWhite
            .BackColorFixed = wpara.nlFlexBackColorFixed   ' RGB(199, 176, 123)
            .BackColorSel = wpara.nlFlexBackColorSel  ' RGB(232, 217, 172)
            .ForeColorSel = vbBlack

            .Left = .Left + 2 * iAdd
            .Top = .Top + 2 * iAdd

'            For ii% = 1 To (.Rows - 1)
'                .row = ii
'
'                .col = 1
'                If (.CellBackColor = vbGreen) Then
'                    .CellBackColor = wpara.nlGreen
'                End If
'
'                .col = 2
'                If (.CellForeColor = vbRed) Then
'                    .CellForeColor = vbWhite
''                    If (Knallrot) Then
''                        .CellBackColor = vbRed
''                    Else
'                        .CellBackColor = wpara.nlRed
''                    End If
'                End If
'
'                .col = 4
'                If (.CellForeColor = vbRed) Then
'                    .CellForeColor = vbWhite
''                    If (Knallrot) Then
''                        .CellBackColor = vbRed
''                    Else
'                        .CellBackColor = wpara.nlRed
''                    End If
'                End If
'
'                .col = 7
'                If (.CellForeColor = vbRed) Then
'                    .CellForeColor = vbWhite
''                    If (Knallrot) Then
''                        .CellBackColor = vbRed
''                    Else
'                        .CellBackColor = wpara.nlRed
''                    End If
'                End If
'            Next i

            .row = 0
            .col = 0
            .ColSel = .Cols - 1
            
            .Redraw = True
        End With

        With frmImportOrIdent
            .Width = .Width + 4 * iAdd
            .Height = .Height + 4 * iAdd

            .flxEdit.Top = .flxEdit.Top + iAdd2

            With .nlcmdOk
                .Init
                .Left = (frmImportOrIdent.ScaleWidth - (.Width * 2 + 300)) / 2
                .Top = frmImportOrIdent.flxEdit.Top + frmImportOrIdent.flxEdit.Height + iAdd + 600
                .Caption = frmImportOrIdent.cmdOk.Caption
                .TabIndex = frmImportOrIdent.cmdOk.TabIndex
                .Enabled = frmImportOrIdent.cmdOk.Enabled
                .Default = frmImportOrIdent.cmdOk.Default
                .Cancel = frmImportOrIdent.cmdOk.Cancel
                .Visible = True
            End With
            .cmdOk.Visible = False

            With .nlcmdEsc
                .Init
                .Left = frmImportOrIdent.nlcmdOk.Left + .Width + 300
                .Top = frmImportOrIdent.nlcmdOk.Top
                .Caption = frmImportOrIdent.cmdEsc.Caption
                .TabIndex = frmImportOrIdent.cmdEsc.TabIndex
                .Enabled = frmImportOrIdent.cmdEsc.Enabled
                .Default = frmImportOrIdent.cmdEsc.Default
                .Cancel = frmImportOrIdent.cmdEsc.Cancel
                .Visible = True
            End With
            .cmdEsc.Visible = False

            .Height = .nlcmdOk.Top + .nlcmdOk.Height + wpara.FrmCaptionHeight + 450

            Call wpara.NewLineWindow(frmImportOrIdent, .nlcmdOk.Top)
            RoundRect .hdc, (.flxEdit.Left - iAdd) / Screen.TwipsPerPixelX, (.flxEdit.Top - iAdd) / Screen.TwipsPerPixelY, (.flxEdit.Left + .flxEdit.Width + iAdd) / Screen.TwipsPerPixelX, (.flxEdit.Top + .flxEdit.Height + iAdd) / Screen.TwipsPerPixelY, 20, 20
        End With
    Else
        .nlcmdOk.Visible = False
        .nlcmdEsc.Visible = False
    End If

    .Left = frmAction.Left + (frmAction.Width - .Width) / 2
    .Top = frmAction.Top + (frmAction.Height - .Height) / 2

    .Show 1
End With

    
'    Close #fhout
'    If AnzVerfügbarkeitsabfrage > 0 And AutoVerfügbarkeitsabfrage Then
'      Call ImportVerfügbar(False, PZNDatei)
'    End If
'    lagernd = SuchLagernde

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'If (IstImportModus% And (RezArtikel(row%).Imp123 = 0)) Or ((IstImportModus% = 0) And (IstAutIdemModus) And (RezArtikel(row%).AutIdem = 0)) Then
'    Call DefErrPop: Exit Sub
'End If
'
'Sonderregelnr% = 0
'NeueKbvNr% = 0
'If (AutIdemSonderregelOk%) Then
'    If (PrivatRezept% = 0) And (AutIdemKbvNr& = 0) Then
'        frmKkMatchcode.Show 1
'        If (AutIdemKbvNr& > 0) Then
'            NeueKbvNr% = True
'        End If
'    End If
'    If (AutIdemKbvNr& > 0) Then
'        SqlSonderregelNr$ = ""
'        SQLStr$ = "SELECT * FROM AutIdemSonderregelKrankenkasse WHERE KbvNr = " + Str$(AutIdemKbvNr&)
''        Set iRec = AutIdemDB.OpenRecordset(SQLStr$)
'        iRec.open SQLStr, AutIdemDB.ActiveConn
'        Do
'            If (iRec.EOF) Then
'                Exit Do
'            End If
'
'            Sonderregelnr% = Format(CheckNullInt(iRec!AutIdemSonderregelNr), "0")
'            If (SqlSonderregelNr$ = "") Then
'                SqlSonderregelNr$ = "SELECT * FROM AutIdemSonderregelArtikel WHERE ("
'            Else
'                SqlSonderregelNr$ = SqlSonderregelNr$ + " OR"
'            End If
'            SqlSonderregelNr$ = SqlSonderregelNr$ + " AutIdemSonderregelNr = " + Str$(Sonderregelnr%)
'
'            iRec.MoveNext
'        Loop
'        iRec.Close
'    End If
'End If
'
'SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + PZN$
''Set TaxeRec = TaxeDB.OpenRecordset(SQLStr$)
'On Error Resume Next
'TaxeRec.Close
'Err.Clear
'On Error GoTo DefErr
'TaxeRec.open SQLStr, taxeAdoDB.ActiveConn
'If (TaxeRec.EOF) Then
'    Call DefErrPop: Exit Sub
'End If
'
'If (IstImportModus%) Then
'    Titel$ = "Auswahl IMPORTE"
'    If (TaxeRec!OriginalPZN = 0&) Then
'        Call DefErrPop: Exit Sub
'    End If
'
'    PZN$ = PznString(TaxeRec!OriginalPZN)
'    SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + PZN$
'    'Set TaxeRec = TaxeDB.OpenRecordset(SQLStr$)
'    On Error Resume Next
'    TaxeRec.Close
'    Err.Clear
'    On Error GoTo DefErr
'    TaxeRec.open SQLStr, taxeAdoDB.ActiveConn
'    If (TaxeRec.EOF) Then
'        Call DefErrPop: Exit Sub
'    End If
'
'    OrgAVP = TaxeRec!vk / 100#
'    FESTBETR = TaxeRec!FESTBETRAG / 100#
''    If ((TaxeRec!FestKz) And (FESTBETR < OrgAvp)) Then
''        OrgAvp = FESTBETR
''    End If
'    OrgAVP = OrgAVP - FNX(TaxeRec!HerstRabatt / 100#)
'Else
'    If (IstAutIdemModus) Then
'        Titel$ = "Auswahl AUT-IDEM"
'    Else
'        Titel$ = "Auswahl IDENTE"
'        If (TaxeRec!IdentGruppe = 0) Then
'            Call DefErrPop: Exit Sub
'        End If
'    End If
'    If (AutIdemOk%) Then
'        If (TaxeRec!AutIdemNeu) Then
'            AutIdemNeu% = True
'        End If
'    End If
'
'    If (AutIdemNeu%) Then
'        OrgKp# = TaxeRec!FESTBETRAG / 100#
''        OrgKp# = TaxeRec!PreisLinie / 100#
''        Titel$ = Titel$ + " (Preislinie: " + Format(OrgKp#, "0.00") + ")"
'    Else
'        OrgKp# = TaxeRec!FESTBETRAG / 100#
'    End If
'
'    OrgKp# = TaxeRec!vk / 100#
'End If
'
'With frmImportOrIdent.flxEdit
'    .Left = 0
'    .Top = 0
'
'    .Rows = 2
'    .FixedRows = 1
'    .FormatString = "<Pzn|<Name|>Menge|<Eh|^|^Kz|<Hersteller|>Preis||"
'    .Rows = 1
'    .Cols = 11
'
'    .ColWidth(0) = frmImportOrIdent.TextWidth(String(10, "9"))
'    .ColWidth(1) = frmImportOrIdent.TextWidth(String(28, "X"))
'    .ColWidth(2) = frmImportOrIdent.TextWidth(String(6, "X"))
'    .ColWidth(3) = frmImportOrIdent.TextWidth(String(3, "X"))
'    .ColWidth(4) = frmImportOrIdent.TextWidth("Ww")
'    .ColWidth(5) = frmImportOrIdent.TextWidth("Ww")
'    .ColWidth(6) = frmImportOrIdent.TextWidth(String(9, "X"))
'    .ColWidth(7) = frmImportOrIdent.TextWidth("9999999.99")
'    .ColWidth(8) = 0
'    .ColWidth(9) = 0    'frmImportOrIdent.TextWidth("Ww")
'    .ColWidth(10) = wpara.FrmScrollHeight
'
'
'    If (IstImportModus%) Then
'        avp# = TaxeRec!vk
'        FESTBETR = TaxeRec!FESTBETRAG
''        If ((TaxeRec!FestKz) And (FESTBETR < AVP)) Then
''            AVP = FESTBETR
''        End If
'        If (TaxeRec!HerstRabattKZ <> 0) Then
'            avp = avp - TaxeRec!HerstRabatt
'        End If
'        avp# = avp / 100#
'
'        s$ = PznString(TaxeRec!PZN) + vbTab + Trim(TaxeRec!Name) + vbTab
'        s$ = s$ + TaxeRec!menge + vbTab + TaxeRec!einheit + vbTab + "i" + vbTab
'        s$ = s$ + vbTab + TaxeRec!HerstellerKB + vbTab + Format(avp#, "### ##0.00") + vbTab
'        s$ = s$ + "zzz" + Format(avp#, "0000000.00") + Trim(TaxeRec!Name) + "," + TaxeRec!HerstellerKB + vbTab + TaxeRec!ArtStatus
'        .AddItem s$
'
'        .row = .Rows - 1
'        .col = 4
'        .CellForeColor = vbRed
'
'        SQLStr$ = "SELECT * FROM TAXE WHERE ORIGINALPZN = " + Str$(TaxeRec!OriginalPZN)
'    ElseIf (IstAutIdemModus = 0) Then
'        BtmDerivatMenge# = 0
'        On Error Resume Next
'        BtmDerivatMenge# = TaxeRec!BtmDerivatMenge
'        On Error GoTo 0
'
'        On Error GoTo DefErr
'        If (BtmDerivatMenge >= 0.01) Then
'            SQLStr$ = "SELECT * FROM M2Inhalt WHERE M2Nr = " + Mid(TaxeRec!M2, 2, 9)
''                Set AutIdemRec = AutIdemDB.OpenRecordset(SQLStr$)
'            On Error Resume Next
'            tRec2.Close
'            Err.Clear
'            On Error GoTo DefErr
'            tRec2.open SQLStr, AutIdemDB.ActiveConn
''                            If Not tRec2.NoMatch Then
'            If Not (tRec2.EOF) Then
'                If (InStr(UCase(CheckNullStr(tRec2!Inhalt)), "HYDROMORPHON") > 0) Then
'                    BtmDerivatMenge = 0
'                End If
'            End If
'            tRec2.Close
'        End If
'
'        SQLStr$ = "SELECT * FROM TAXE WHERE IDENTGRUPPE = " + Str$(TaxeRec!IdentGruppe)
''        SQLStr$ = SQLStr$ + " And einheit = """ + TaxeRec!einheit + """"
'        SQLStr$ = SQLStr$ + " And einheit = '" + TaxeRec!einheit + "'"
'        If (BtmDerivatMenge >= 0.01) Then
'            SQLStr$ = SQLStr$ + " AND (BtmDerivatMenge=" + Str(BtmDerivatMenge) + ")"
'        End If
'        SQLStr$ = SQLStr$ + " ORDER BY Menge,VK"
'    ElseIf (AutIdemNeu%) Then
'        BtmDerivatMenge# = 0
'        On Error Resume Next
'        BtmDerivatMenge# = TaxeRec!BtmDerivatMenge
'        On Error GoTo 0
'
'        On Error GoTo DefErr
'        If (BtmDerivatMenge >= 0.01) Then
'            SQLStr$ = "SELECT * FROM M2Inhalt WHERE M2Nr = " + Mid(TaxeRec!M2, 2, 9)
''                Set AutIdemRec = AutIdemDB.OpenRecordset(SQLStr$)
'            On Error Resume Next
'            tRec2.Close
'            Err.Clear
'            On Error GoTo DefErr
'            tRec2.open SQLStr, AutIdemDB.ActiveConn
''                            If Not tRec2.NoMatch Then
'            If Not (tRec2.EOF) Then
'                If (InStr(UCase(CheckNullStr(tRec2!Inhalt)), "HYDROMORPHON") > 0) Then
'                    BtmDerivatMenge = 0
'                End If
'            End If
'            tRec2.Close
'        End If
'
'        SQLStr$ = "SELECT * FROM PZNSUBST WHERE PZN = " + Str$(TaxeRec!PZN)
''        Set AutIdemRec = AutIdemDB.OpenRecordset(SQLStr$)
'        AutIdemRec.open SQLStr, AutIdemDB.ActiveConn
'        If (AutIdemRec.EOF) Then
'            SQLStr$ = ""
'        Else
'            i% = 1
'            SQLStr$ = "SELECT * FROM TAXE WHERE (PZN ="
'            Do
'                If (AutIdemRec.EOF) Then
'                    Exit Do
'                End If
'
'                If (i% > 1) Then SQLStr$ = SQLStr$ + " OR PZN ="
'                SQLStr$ = SQLStr$ + Str$(AutIdemRec!Substpzn)
'                AutIdemRec.MoveNext
'
'                i% = i% + 1
'                If (i% > 250) Then
'                    Exit Do
'                End If
'            Loop
'            If (BtmDerivatMenge >= 0.01) Then
'                SQLStr$ = SQLStr$ + ") AND (BtmDerivatMenge=" + Str(BtmDerivatMenge)
'            End If
''            SQLStr$ = SQLStr$ + ") ORDER BY VK-HerstRabatt,NAME"
'            SQLStr$ = SQLStr$ + ") ORDER BY VK,NAME"
'        End If
'        AutIdemRec.Close
'        PZN$ = ""
'    Else
'        SQLStr = ""
''        If (TaxeRec!IdentGruppe > 0&) Then
''            SQLStr$ = HoleOpAutIdem$
''            pzn$ = ""
''        End If
'    End If
'
'    If (SQLStr = "") Then
'        Call DefErrPop: Exit Sub
'    End If
'
'    'Set TaxeRec = TaxeDB.OpenRecordset(SQLStr$)
'    On Error Resume Next
'    TaxeRec.Close
'    Err.Clear
'    On Error GoTo DefErr
'    TaxeRec.open SQLStr, taxeAdoDB.ActiveConn
'    If (TaxeRec.EOF) Then
'        Call DefErrPop: Exit Sub
'    End If
'
'
'    While Not (TaxeRec.EOF)
'
''        If (TaxeRec!ArtStatus <> "S") Then
'            avp# = TaxeRec!vk
'            FESTBETR = TaxeRec!FESTBETRAG
''            If (IstImportModus) And ((TaxeRec!FestKz) And (FESTBETR < AVP)) Then
''                AVP = FESTBETR
''            End If
'            If (IstImportModus) And (TaxeRec!HerstRabattKZ <> 0) Then
'                avp = avp - TaxeRec!HerstRabatt
'            End If
'            avp# = avp / 100#
'
'            If (IstAutIdemModus = 0) Or (TaxeRec!PZN <> Val(PZN$)) Then
'                If (IstImportModus%) Then
'                    If (OrgAVP# - 15) >= avp# Or (OrgAVP# * 0.85) >= avp# Then
'                    'If (avp# < (OrgAvp# - 1)) And (avp# < (OrgAvp# * 0.9)) Then
'                        kz$ = ""
'                    Else
'                        kz$ = "-"
'                    End If
'                Else
'                    kz$ = ""
'                    If (IstAutIdemModus) And (OrgKp# > 0) Then
'                        If (avp# > OrgKp#) Then
'                            kz$ = "+"
'                        Else
'                            kz$ = "-"
'                        End If
'                    End If
'                End If
'
'                s$ = PznString(TaxeRec!PZN) + vbTab + Trim(TaxeRec!Name) + vbTab
'                s$ = s$ + TaxeRec!menge + vbTab + TaxeRec!einheit + vbTab + vbTab
'                s$ = s$ + kz$ + vbTab + TaxeRec!HerstellerKB + vbTab + Format(avp#, "### ##0.00") + vbTab
'                If (IstAutIdemModus = 0) Then
'                    s$ = s$ + TaxeRec!menge + "," + Format(avp#, "0000000.00") + Trim(TaxeRec!Name) + "," + TaxeRec!HerstellerKB
'                Else
'                    s$ = s$ + Format(avp#, "0000000.00") + Trim(TaxeRec!Name) + "," + TaxeRec!HerstellerKB
'                End If
'                s$ = s$ + vbTab + TaxeRec!ArtStatus
'    '            If Not (IstImportModus%) Or kz$ = "" Then
'                    .AddItem s$
'
'                    If (ArtikelDbOk) Then
'                        SQLStr$ = "SELECT * FROM ARTIKEL WHERE PZN = " + PznString(TaxeRec!PZN)
'                        SQLStr = SQLStr + " AND LagerKz<>0"
'                        FabsErrf = Artikel.OpenRecordset(ArtikelAdoRec, SQLStr, 0)
'                    Else
'                        FabsErrf% = ass.IndexSearch(0, Format(TaxeRec!PZN, "0000000"), FabsRecno&)
'                    End If
'                    If (FabsErrf% = 0) Then
'                        .FillStyle = flexFillRepeat
'                        .row = .Rows - 1
'                        .col = 0
'                        .RowSel = .row
'                        .ColSel = .Cols - 1
'                        .CellFontBold = True
'                        .FillStyle = flexFillSingle
'                    End If
'
'                    rabKZ = "zzz"
'                    If (SqlSonderregelNr$ <> "") Then
'                        SQLStr$ = SqlSonderregelNr$ + ")"
'                        SQLStr$ = SQLStr$ + " AND NrKz='P' AND ZoNr=" + Str$(TaxeRec!PZN)
''                        Set iRec = AutIdemDB.OpenRecordset(SQLStr$)
''                        If (iRec.RecordCount > 0) Then
'                        iRec.open SQLStr, AutIdemDB.ActiveConn
'                        If Not (iRec.EOF) Then
'                            rabKZ = "   "
'                            .FillStyle = flexFillRepeat
'                            .row = .Rows - 1
'                            .col = 1
'                            .RowSel = .row
'                            .ColSel = .col
'                            .CellBackColor = vbGreen
'                            .FillStyle = flexFillSingle
'                        End If
'                        iRec.Close
'                    End If
'                    .TextMatrix(.Rows - 1, 8) = rabKZ + .TextMatrix(.Rows - 1, 8)
'
'                    If (Val(TaxeRec!AlternativPackung) > 0) Then
'                        .FillStyle = flexFillRepeat
'                        .row = .Rows - 1
'                        .col = 2
'                        .RowSel = .row
'                        .ColSel = .col
'                        .CellForeColor = vbRed
'                        .FillStyle = flexFillSingle
'                    End If
'
'                    .FillStyle = flexFillRepeat
'                    .row = .Rows - 1
'                    .col = 4
'                    .RowSel = .row
'                    If (TaxeRec!OriginalPZN = TaxeRec!PZN) Then
'                        .text = "i"
'                        .CellForeColor = vbRed
'                    ElseIf (TaxeRec!OriginalPZN > 0) Then
'                        .text = "I"
'
''                        Set sTaxeRec = TaxeDB.OpenRecordset("SELECT * FROM TAXE WHERE PZN = " + CStr(TaxeRec!OriginalPZN))
'                        On Error Resume Next
'                        sTaxeRec.Close
'                        Err.Clear
'                        On Error GoTo DefErr
'                        sTaxeRec.open "SELECT * FROM TAXE WHERE PZN = " + CStr(TaxeRec!OriginalPZN), taxeAdoDB.ActiveConn
'                        If Not (sTaxeRec.EOF) Then
'                            avp2# = sTaxeRec!vk
'                            FESTBETR = sTaxeRec!FESTBETRAG
''                            If (IstImportModus) And ((sTaxeRec!FestKz) And (FESTBETR < avp2)) Then
''                                avp2 = FESTBETR
''                            End If
'                            If (IstImportModus) And (sTaxeRec!HerstRabattKZ <> 0) Then
'                                avp2 = avp2 - sTaxeRec!HerstRabatt
'                            End If
'                            avp2# = avp2 / 100#
'
'                            If (avp2) - 15 >= (avp) Or (avp2) * 0.85 >= avp Then
'                            Else
'                                .CellForeColor = vbRed
'                            End If
'                        End If
'                    End If
'                    .FillStyle = flexFillSingle
'    '            End If
'            End If
''        End If
'
'        TaxeRec.MoveNext
'    Wend
'
'    .row = 1
'    .col = 8
'    .RowSel = .Rows - 1
'    .ColSel = .col
''    If (IstImportModus%) Then
''        .Sort = 8
''    Else
''        .Sort = 1
''    End If
'    .Sort = flexSortGenericAscending
'    .col = 1
'    .ColSel = .Cols - 1
'
'    .Height = .RowHeight(0) * .Rows + 90
'    If (para.Newline) Then
'        MaxHe = wpara.WorkAreaHeight - 2400
'    Else
'        MaxHe = wpara.WorkAreaHeight - 150
'    End If
'    If (.Height > MaxHe) Then
''        .Height = wpara.WorkAreaHeight - 150
'        .Height = MaxHe
'        .Height = ((.Height - 90) \ .RowHeight(0)) * .RowHeight(0) + 90
'    End If
'
'    spBreite& = 0
'    For i% = 0 To (.Cols - 1)
'        spBreite& = spBreite& + .ColWidth(i%)
'    Next i%
'    .Width = spBreite& + 90
'
'    With frmImportOrIdent
'        .Caption = Titel$
'        .Height = .flxEdit.Height + wpara.FrmCaptionHeight + 60
'        .Width = .flxEdit.Width + 60
'        .Left = ParentForm.Left + (ParentForm.Width - .Width) / 2
'        .Top = ParentForm.Top + (ParentForm.Height - .Height) / 2
'    End With
'
''    If (IstImportModus% = 0) And (IstAutIdemModus) Then
'    If (IstImportModus% = 0) Then
'        If (AutIdemNeu%) Then
'            grenze# = OrgKp#
'        ElseIf (.Rows > 6) Then
'            For i% = 1 To 3
'                unten# = unten# + xVal(.TextMatrix(i%, 7))
'            Next i%
'            unten# = unten# / 3
'            For i% = (.Rows - 3) To (.Rows - 1)
'                oben# = oben# + xVal(.TextMatrix(i%, 7))
'            Next i%
'            oben# = oben# / 3
'            grenze# = xVal(.TextMatrix(1, 7)) + (oben# - unten#) / 3 'unteres Drittel
'        Else
'            If (.Rows <= 4) Then
'                i% = 1
'            Else
'                i% = .Rows - 3
'            End If
'            grenze# = xVal(.TextMatrix(i%, 7))
'        End If
'
'        For i% = 1 To (.Rows - 1)
'            avp# = xVal(.TextMatrix(i%, 7))
'
'            .row = i%
'            .col = 7
'            .RowSel = .row
'            .ColSel = .col
'            '.CellFontBold = (avp# > grenze#)
'            If AnzGilt% < 3 Or avp# = grenze# Then
'              '.CellFontBold = (avp# > grenze#)
'              .CellFontBold = False
'              grenze# = avp#
'            Else
'              .CellFontBold = True
'                If (IstAutIdemModus) Then
'                    .CellForeColor = vbRed
'                End If
'            End If
'            If AnzGilt% <= 3 Then grenze# = avp#
'            If (avp# <= grenze#) Then
'                AnzGilt% = AnzGilt% + 1
'            End If
'        Next i%
'
'        If (AutIdemNeu% = 0) And (AnzGilt% < 5) And (.Rows > 6) Then 'wenn weniger als 5 im 1. Drittel -> 5 zulassen
'            For i% = (AnzGilt + 1) To 5
'                .row = i%
'                .col = 7
'                .RowSel = .row
'                .ColSel = .col
'                .CellFontBold = False
'            Next i%
'            grenze# = xVal(.TextMatrix(5, 7))
'            Do
'                If (i% >= .Rows) Then Exit Do
'                If (xVal(.TextMatrix(i%, 7)) = grenze#) Then
'                    .row = i%
'                    .col = 7
'                    .RowSel = .row
'                    .ColSel = .col
'                    .CellFontBold = False
'                    i% = i% + 1
'                Else
'                    Exit Do
'                End If
'            Loop
'        End If
'
'        .row = 1
'    Else
''      .row = 1
''      .col = 7
''      .RowSel = .row
''      .ColSel = .col
''      .CellFontBold = True
''      .CellForeColor = vbRed
''
''      If .Rows > 2 Then
''        .row = 2
''      End If
'
'        For i = 1 To .Rows - 1
'            .row = i
'            .col = 4
'            If (.text <> "i") And (.CellForeColor = vbRed) Then
'                .col = 7
'                .CellForeColor = vbRed
'            End If
'        Next i
'        .row = 1
'
'    End If
'
'    .col = 0
'    .ColSel = .Cols - 1
'    .Visible = True
'End With
'
'With frmImportOrIdent
'    .cmdF2.Enabled = (IstImportModus% = 0)
'
'    If (para.Newline) Then
'        iAdd = wpara.NlFlexBackY
'        iAdd2 = wpara.NlCaptionY
'
'        With frmImportOrIdent.flxEdit
''            .ScrollBars = flexScrollBarNone
'            .BorderStyle = 0
'            .Width = .Width - 90
'            .Height = .Height - 75
'            .GridLines = flexGridFlat
'            .GridLinesFixed = flexGridFlat
'            .GridColorFixed = .GridColor
'            .BackColor = wpara.nlFlexBackColor    'vbWhite
'            .BackColorBkg = wpara.nlFlexBackColor    'vbWhite
'            .BackColorFixed = wpara.nlFlexBackColorFixed   ' RGB(199, 176, 123)
'            .BackColorSel = wpara.nlFlexBackColorSel  ' RGB(232, 217, 172)
'            .ForeColorSel = vbBlack
'
'            .Left = .Left + 2 * iAdd
'            .Top = .Top + 2 * iAdd
'
'            For i% = 1 To (.Rows - 1)
'                .row = i
'
'                .col = 1
'                If (.CellBackColor = vbGreen) Then
'                    .CellBackColor = wpara.nlGreen
'                End If
'
'                .col = 2
'                If (.CellForeColor = vbRed) Then
'                    .CellForeColor = vbWhite
''                    If (Knallrot) Then
''                        .CellBackColor = vbRed
''                    Else
'                        .CellBackColor = wpara.nlRed
''                    End If
'                End If
'
'                .col = 4
'                If (.CellForeColor = vbRed) Then
'                    .CellForeColor = vbWhite
''                    If (Knallrot) Then
''                        .CellBackColor = vbRed
''                    Else
'                        .CellBackColor = wpara.nlRed
''                    End If
'                End If
'
'                .col = 7
'                If (.CellForeColor = vbRed) Then
'                    .CellForeColor = vbWhite
''                    If (Knallrot) Then
''                        .CellBackColor = vbRed
''                    Else
'                        .CellBackColor = wpara.nlRed
''                    End If
'                End If
'            Next i
'
'            .row = 1
'            .col = 0
'            .ColSel = .Cols - 1
'        End With
'
'        With frmImportOrIdent
'            .Width = .Width + 4 * iAdd
'            .Height = .Height + 4 * iAdd
'
'            .flxEdit.Top = .flxEdit.Top + iAdd2
'
'            With .nlcmdOk
'                .Init
'                .Left = (frmImportOrIdent.ScaleWidth - (.Width * 2 + 300)) / 2
'                .Top = frmImportOrIdent.flxEdit.Top + frmImportOrIdent.flxEdit.Height + iAdd + 600
'                .Caption = frmImportOrIdent.cmdOk.Caption
'                .TabIndex = frmImportOrIdent.cmdOk.TabIndex
'                .Enabled = frmImportOrIdent.cmdOk.Enabled
'                .Default = frmImportOrIdent.cmdOk.Default
'                .Cancel = frmImportOrIdent.cmdOk.Cancel
'                .Visible = True
'            End With
'            .cmdOk.Visible = False
'
'            With .nlcmdEsc
'                .Init
'                .Left = frmImportOrIdent.nlcmdOk.Left + .Width + 300
'                .Top = frmImportOrIdent.nlcmdOk.Top
'                .Caption = frmImportOrIdent.cmdEsc.Caption
'                .TabIndex = frmImportOrIdent.cmdEsc.TabIndex
'                .Enabled = frmImportOrIdent.cmdEsc.Enabled
'                .Default = frmImportOrIdent.cmdEsc.Default
'                .Cancel = frmImportOrIdent.cmdEsc.Cancel
'                .Visible = True
'            End With
'            .cmdEsc.Visible = False
'
'            .Height = .nlcmdOk.Top + .nlcmdOk.Height + wpara.FrmCaptionHeight + 450
'
'            Call wpara.NewLineWindow(frmImportOrIdent, .nlcmdOk.Top)
'            RoundRect .hdc, (.flxEdit.Left - iAdd) / Screen.TwipsPerPixelX, (.flxEdit.Top - iAdd) / Screen.TwipsPerPixelY, (.flxEdit.Left + .flxEdit.Width + iAdd) / Screen.TwipsPerPixelX, (.flxEdit.Top + .flxEdit.Height + iAdd) / Screen.TwipsPerPixelY, 20, 20
'        End With
'    Else
'        .nlcmdOk.Visible = False
'        .nlcmdEsc.Visible = False
'    End If
'
'    .Left = frmAction.Left + (frmAction.Width - .Width) / 2
'    .Top = frmAction.Top + (frmAction.Height - .Height) / 2
'
'    .Show 1
'End With
'
If (EditErg%) Then
    ind% = InStr(EditTxt$, vbTab)
    pzn$ = Left$(EditTxt$, ind% - 1)

    Call NeuerArtikel(row%, pzn$)
    Call PaintArtikel(row%)
End If
'
'
'If (NeueKbvNr%) Then
'    iAnzZ% = AnzRezeptArtikel% - 1      'GS 16.05.02
'    If (iAnzZ% > 6) Then
'        iAnzZ% = 6
'    End If
'    iTabIndex% = AktTabIndex%
'    For i% = 0 To iAnzZ%
'        RezArtikel(i%).AutIdem = 0
'        RezArtikel(i%).N0 = 0
'        SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + RezArtikel(i%).PZN
'        'Set TaxeRec = TaxeDB.OpenRecordset(SQLStr$)
'        On Error Resume Next
'        TaxeRec.Close
'        Err.Clear
'        On Error GoTo DefErr
'        TaxeRec.open SQLStr, taxeAdoDB.ActiveConn
'        Call CheckAutIdem(i%)
'        Call PaintArtikel(i%)
'    Next i%
'    Call TabIndexChange(iTabIndex%, False)
'End If

Call DefErrPop
End Sub

Sub ImporteOrAutIdemAlt(IstImportModus%, Optional IstAutIdemModus% = True)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("ImporteOrAutIdemAlt")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
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
Dim i%, ind%, row%, AnzGilt%, AutIdemNeu%, Sonderregelnr%, NeueKbvNr%, iAnzZ%, iTabIndex%
Dim iAdd%, iAdd2%
Dim spBreite&, MaxHe&
Dim OrgAvp#, OrgKp#, AVP#, avp2#, unten#, oben#, grenze#, BtmDerivatMenge#, FESTBETR#
Dim s$, SQLStr$, kz$, OrgPZN$, pzn$, Titel$, SqlSonderregelNr$, h$, rabKZ$
Dim iRec As New ADODB.Recordset
Dim sTaxerec As New ADODB.Recordset
Dim tRec2 As New ADODB.Recordset
Dim ind2%

If (AktTabIndex% < 3) Then
    Call DefErrPop: Exit Sub
End If
    
row% = (AktTabIndex% - 3) \ 4
pzn$ = RezArtikel(row%).pzn
IdentPzn$ = pzn$

If (row >= AnzRezeptArtikel%) Then
    Call DefErrPop: Exit Sub
End If


If (IstImportModus% And (RezArtikel(row%).Imp123 = 0)) Or ((IstImportModus% = 0) And (IstAutIdemModus) And (RezArtikel(row%).AutIdem = 0)) Then
    Call DefErrPop: Exit Sub
End If

Sonderregelnr% = 0
NeueKbvNr% = 0
If (AutIdemSonderregelOk%) Then
    If (PrivatRezept% = 0) And (AutIdemKbvNr& = 0) Then
        frmKkMatchcode.Show 1
        If (AutIdemKbvNr& > 0) Then
            NeueKbvNr% = True
        End If
    End If
    If (AutIdemKbvNr& > 0) Then
        SqlSonderregelNr$ = ""
        SQLStr$ = "SELECT * FROM AutIdemSonderregelKrankenkasse WHERE KbvNr = " + Str$(AutIdemKbvNr&)
'        Set iRec = AutIdemDB.OpenRecordset(SQLStr$)
        iRec.Open SQLStr, AutIdemDB.ActiveConn
        Do
            If (iRec.EOF) Then
                Exit Do
            End If

            Sonderregelnr% = Format(CheckNullInt(iRec!AutIdemSonderregelNr), "0")
            If (SqlSonderregelNr$ = "") Then
                SqlSonderregelNr$ = "SELECT * FROM AutIdemSonderregelArtikel WHERE ("
            Else
                SqlSonderregelNr$ = SqlSonderregelNr$ + " OR"
            End If
            SqlSonderregelNr$ = SqlSonderregelNr$ + " AutIdemSonderregelNr = " + Str$(Sonderregelnr%)

            iRec.MoveNext
        Loop
        iRec.Close
    End If
End If

SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + pzn$
'Set TaxeRec = TaxeDB.OpenRecordset(SQLStr$)
On Error Resume Next
TaxeRec.Close
Err.Clear
On Error GoTo DefErr
TaxeRec.Open SQLStr, taxeAdoDB.ActiveConn
If (TaxeRec.EOF) Then
    Call DefErrPop: Exit Sub
End If

If (IstImportModus%) Then
    Titel$ = "Auswahl IMPORTE"
    If (TaxeRec!OriginalPZN = 0&) Then
        Call DefErrPop: Exit Sub
    End If

    pzn$ = PznString(TaxeRec!OriginalPZN)
    SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + pzn$
    'Set TaxeRec = TaxeDB.OpenRecordset(SQLStr$)
    On Error Resume Next
    TaxeRec.Close
    Err.Clear
    On Error GoTo DefErr
    TaxeRec.Open SQLStr, taxeAdoDB.ActiveConn
    If (TaxeRec.EOF) Then
        Call DefErrPop: Exit Sub
    End If

    OrgAvp = TaxeRec!vk / 100#
    FESTBETR = TaxeRec!FESTBETRAG / 100#
'    If ((TaxeRec!FestKz) And (FESTBETR < OrgAvp)) Then
'        OrgAvp = FESTBETR
'    End If
    OrgAvp = OrgAvp - FNX(TaxeRec!HerstRabatt / 100#)
Else
    If (IstAutIdemModus) Then
        Titel$ = "Auswahl AUT-IDEM"
    Else
        Titel$ = "Auswahl IDENTE"
        If (TaxeRec!Identgruppe = 0) Then
            Call DefErrPop: Exit Sub
        End If
    End If
    If (AutIdemOk%) Then
        If (TaxeRec!AutIdemNeu) Then
            AutIdemNeu% = True
        End If
    End If

    If (AutIdemNeu%) Then
        OrgKp# = TaxeRec!FESTBETRAG / 100#
'        OrgKp# = TaxeRec!PreisLinie / 100#
'        Titel$ = Titel$ + " (Preislinie: " + Format(OrgKp#, "0.00") + ")"
    Else
        OrgKp# = TaxeRec!FESTBETRAG / 100#
    End If

    OrgKp# = TaxeRec!vk / 100#
End If

With frmImportOrIdent.flxEdit
    .Left = 0
    .Top = 0

    .Rows = 2
    .FixedRows = 1
    .FormatString = "<Pzn|<Name|>Menge|<Eh|^|^Kz|<Hersteller|>Preis||"
    .Rows = 1
    .Cols = 11

    .ColWidth(0) = frmImportOrIdent.TextWidth(String(10, "9"))
    .ColWidth(1) = frmImportOrIdent.TextWidth(String(28, "X"))
    .ColWidth(2) = frmImportOrIdent.TextWidth(String(6, "X"))
    .ColWidth(3) = frmImportOrIdent.TextWidth(String(3, "X"))
    .ColWidth(4) = frmImportOrIdent.TextWidth("Ww")
    .ColWidth(5) = frmImportOrIdent.TextWidth("Ww")
    .ColWidth(6) = frmImportOrIdent.TextWidth(String(9, "X"))
    .ColWidth(7) = frmImportOrIdent.TextWidth("9999999.99")
    .ColWidth(8) = 0
    .ColWidth(9) = 0    'frmImportOrIdent.TextWidth("Ww")
    .ColWidth(10) = wpara.FrmScrollHeight


    If (IstImportModus%) Then
        AVP# = TaxeRec!vk
        FESTBETR = TaxeRec!FESTBETRAG
'        If ((TaxeRec!FestKz) And (FESTBETR < AVP)) Then
'            AVP = FESTBETR
'        End If
        If (TaxeRec!HerstRabattKZ <> 0) Then
            AVP = AVP - TaxeRec!HerstRabatt
        End If
        AVP# = AVP / 100#

        s$ = PznString(TaxeRec!pzn) + vbTab + Trim(TaxeRec!Name) + vbTab
        s$ = s$ + TaxeRec!menge + vbTab + TaxeRec!einheit + vbTab + "i" + vbTab
        s$ = s$ + vbTab + TaxeRec!HerstellerKB + vbTab + Format(AVP#, "### ##0.00") + vbTab
        s$ = s$ + "zzz" + Format(AVP#, "0000000.00") + Trim(TaxeRec!Name) + "," + TaxeRec!HerstellerKB + vbTab + TaxeRec!ArtStatus
        .AddItem s$

        .row = .Rows - 1
        .col = 4
        .CellForeColor = vbRed

        SQLStr$ = "SELECT * FROM TAXE WHERE ORIGINALPZN = " + Str$(TaxeRec!OriginalPZN)
    ElseIf (IstAutIdemModus = 0) Then
        BtmDerivatMenge# = 0
        On Error Resume Next
        BtmDerivatMenge# = TaxeRec!BtmDerivatMenge
        On Error GoTo 0

        On Error GoTo DefErr
        If (BtmDerivatMenge >= 0.01) Then
            SQLStr$ = "SELECT * FROM M2Inhalt WHERE M2Nr = " + Mid(TaxeRec!M2, 2, 9)
'                Set AutIdemRec = AutIdemDB.OpenRecordset(SQLStr$)
            On Error Resume Next
            tRec2.Close
            Err.Clear
            On Error GoTo DefErr
            tRec2.Open SQLStr, AutIdemDB.ActiveConn
'                            If Not tRec2.NoMatch Then
            If Not (tRec2.EOF) Then
                If (InStr(UCase(CheckNullStr(tRec2!Inhalt)), "HYDROMORPHON") > 0) Then
                    BtmDerivatMenge = 0
                End If
            End If
            tRec2.Close
        End If

        SQLStr$ = "SELECT * FROM TAXE WHERE IDENTGRUPPE = " + Str$(TaxeRec!Identgruppe)
'        SQLStr$ = SQLStr$ + " And einheit = """ + TaxeRec!einheit + """"
        SQLStr$ = SQLStr$ + " And einheit = '" + TaxeRec!einheit + "'"
        If (BtmDerivatMenge >= 0.01) Then
            SQLStr$ = SQLStr$ + " AND (BtmDerivatMenge=" + Str(BtmDerivatMenge) + ")"
        End If
        SQLStr$ = SQLStr$ + " ORDER BY Menge,VK"
    ElseIf (AutIdemNeu%) Then
        BtmDerivatMenge# = 0
        On Error Resume Next
        BtmDerivatMenge# = TaxeRec!BtmDerivatMenge
        On Error GoTo 0

        On Error GoTo DefErr
        If (BtmDerivatMenge >= 0.01) Then
            SQLStr$ = "SELECT * FROM M2Inhalt WHERE M2Nr = " + Mid(TaxeRec!M2, 2, 9)
'                Set AutIdemRec = AutIdemDB.OpenRecordset(SQLStr$)
            On Error Resume Next
            tRec2.Close
            Err.Clear
            On Error GoTo DefErr
            tRec2.Open SQLStr, AutIdemDB.ActiveConn
'                            If Not tRec2.NoMatch Then
            If Not (tRec2.EOF) Then
                If (InStr(UCase(CheckNullStr(tRec2!Inhalt)), "HYDROMORPHON") > 0) Then
                    BtmDerivatMenge = 0
                End If
            End If
            tRec2.Close
        End If

        SQLStr$ = "SELECT * FROM PZNSUBST WHERE PZN = " + Str$(TaxeRec!pzn)
'        Set AutIdemRec = AutIdemDB.OpenRecordset(SQLStr$)
        AutIdemRec.Open SQLStr, AutIdemDB.ActiveConn
        If (AutIdemRec.EOF) Then
            SQLStr$ = ""
        Else
            i% = 1
            SQLStr$ = "SELECT * FROM TAXE WHERE (PZN ="
            Do
                If (AutIdemRec.EOF) Then
                    Exit Do
                End If

                If (i% > 1) Then SQLStr$ = SQLStr$ + " OR PZN ="
                SQLStr$ = SQLStr$ + Str$(AutIdemRec!Substpzn)
                AutIdemRec.MoveNext

                i% = i% + 1
                If (i% > 250) Then
                    Exit Do
                End If
            Loop
            If (BtmDerivatMenge >= 0.01) Then
                SQLStr$ = SQLStr$ + ") AND (BtmDerivatMenge=" + Str(BtmDerivatMenge)
            End If
'            SQLStr$ = SQLStr$ + ") ORDER BY VK-HerstRabatt,NAME"
            SQLStr$ = SQLStr$ + ") ORDER BY VK,NAME"
        End If
        AutIdemRec.Close
        pzn$ = ""
    Else
        SQLStr = ""
'        If (TaxeRec!IdentGruppe > 0&) Then
'            SQLStr$ = HoleOpAutIdem$
'            pzn$ = ""
'        End If
    End If

    If (SQLStr = "") Then
        Call DefErrPop: Exit Sub
    End If

    'Set TaxeRec = TaxeDB.OpenRecordset(SQLStr$)
    On Error Resume Next
    TaxeRec.Close
    Err.Clear
    On Error GoTo DefErr
    TaxeRec.Open SQLStr, taxeAdoDB.ActiveConn
    If (TaxeRec.EOF) Then
        Call DefErrPop: Exit Sub
    End If


    While Not (TaxeRec.EOF)

'        If (TaxeRec!ArtStatus <> "S") Then
            AVP# = TaxeRec!vk
            FESTBETR = TaxeRec!FESTBETRAG
'            If (IstImportModus) And ((TaxeRec!FestKz) And (FESTBETR < AVP)) Then
'                AVP = FESTBETR
'            End If
            If (IstImportModus) And (TaxeRec!HerstRabattKZ <> 0) Then
                AVP = AVP - TaxeRec!HerstRabatt
            End If
            AVP# = AVP / 100#

            If (IstAutIdemModus = 0) Or (TaxeRec!pzn <> Val(pzn$)) Then
                If (IstImportModus%) Then
                    If (OrgAvp# - 15) >= AVP# Or (OrgAvp# * 0.85) >= AVP# Then
                    'If (avp# < (OrgAvp# - 1)) And (avp# < (OrgAvp# * 0.9)) Then
                        kz$ = ""
                    Else
                        kz$ = "-"
                    End If
                Else
                    kz$ = ""
                    If (IstAutIdemModus) And (OrgKp# > 0) Then
                        If (AVP# > OrgKp#) Then
                            kz$ = "+"
                        Else
                            kz$ = "-"
                        End If
                    End If
                End If

                s$ = PznString(TaxeRec!pzn) + vbTab + Trim(TaxeRec!Name) + vbTab
                s$ = s$ + TaxeRec!menge + vbTab + TaxeRec!einheit + vbTab + vbTab
                s$ = s$ + kz$ + vbTab + TaxeRec!HerstellerKB + vbTab + Format(AVP#, "### ##0.00") + vbTab
                If (IstAutIdemModus = 0) Then
                    s$ = s$ + TaxeRec!menge + "," + Format(AVP#, "0000000.00") + Trim(TaxeRec!Name) + "," + TaxeRec!HerstellerKB
                Else
                    s$ = s$ + Format(AVP#, "0000000.00") + Trim(TaxeRec!Name) + "," + TaxeRec!HerstellerKB
                End If
                s$ = s$ + vbTab + TaxeRec!ArtStatus
    '            If Not (IstImportModus%) Or kz$ = "" Then
                    .AddItem s$

                    If (ArtikelDbOk) Then
                        SQLStr$ = "SELECT * FROM ARTIKEL WHERE PZN = " + PznString(TaxeRec!pzn)
                        SQLStr = SQLStr + " AND LagerKz<>0"
                        FabsErrf = Artikel.OpenRecordset(ArtikelAdoRec, SQLStr, 0)
                    Else
                        FabsErrf% = ass.IndexSearch(0, Format(TaxeRec!pzn, "0000000"), FabsRecno&)
                    End If
                    If (FabsErrf% = 0) Then
                        .FillStyle = flexFillRepeat
                        .row = .Rows - 1
                        .col = 0
                        .RowSel = .row
                        .ColSel = .Cols - 1
                        .CellFontBold = True
                        .FillStyle = flexFillSingle
                    End If

                    rabKZ = "zzz"
                    If (SqlSonderregelNr$ <> "") Then
                        SQLStr$ = SqlSonderregelNr$ + ")"
                        SQLStr$ = SQLStr$ + " AND NrKz='P' AND ZoNr=" + Str$(TaxeRec!pzn)
'                        Set iRec = AutIdemDB.OpenRecordset(SQLStr$)
'                        If (iRec.RecordCount > 0) Then
                        iRec.Open SQLStr, AutIdemDB.ActiveConn
                        If Not (iRec.EOF) Then
                            rabKZ = "   "
                            .FillStyle = flexFillRepeat
                            .row = .Rows - 1
                            .col = 1
                            .RowSel = .row
                            .ColSel = .col
                            .CellBackColor = vbGreen
                            .FillStyle = flexFillSingle
                        End If
                        iRec.Close
                    End If
                    .TextMatrix(.Rows - 1, 8) = rabKZ + .TextMatrix(.Rows - 1, 8)

                    If (Val(TaxeRec!AlternativPackung) > 0) Then
                        .FillStyle = flexFillRepeat
                        .row = .Rows - 1
                        .col = 2
                        .RowSel = .row
                        .ColSel = .col
                        .CellForeColor = vbRed
                        .FillStyle = flexFillSingle
                    End If

                    .FillStyle = flexFillRepeat
                    .row = .Rows - 1
                    .col = 4
                    .RowSel = .row
                    If (TaxeRec!OriginalPZN = TaxeRec!pzn) Then
                        .text = "i"
                        .CellForeColor = vbRed
                    ElseIf (TaxeRec!OriginalPZN > 0) Then
                        .text = "I"

'                        Set sTaxeRec = TaxeDB.OpenRecordset("SELECT * FROM TAXE WHERE PZN = " + CStr(TaxeRec!OriginalPZN))
                        On Error Resume Next
                        sTaxerec.Close
                        Err.Clear
                        On Error GoTo DefErr
                        sTaxerec.Open "SELECT * FROM TAXE WHERE PZN = " + CStr(TaxeRec!OriginalPZN), taxeAdoDB.ActiveConn
                        If Not (sTaxerec.EOF) Then
                            avp2# = sTaxerec!vk
                            FESTBETR = sTaxerec!FESTBETRAG
'                            If (IstImportModus) And ((sTaxeRec!FestKz) And (FESTBETR < avp2)) Then
'                                avp2 = FESTBETR
'                            End If
                            If (IstImportModus) And (sTaxerec!HerstRabattKZ <> 0) Then
                                avp2 = avp2 - sTaxerec!HerstRabatt
                            End If
                            avp2# = avp2 / 100#

                            If (avp2) - 15 >= (AVP) Or (avp2) * 0.85 >= AVP Then
                            Else
                                .CellForeColor = vbRed
                            End If
                        End If
                    End If
                    .FillStyle = flexFillSingle
    '            End If
            End If
'        End If

        TaxeRec.MoveNext
    Wend

    .row = 1
    .col = 8
    .RowSel = .Rows - 1
    .ColSel = .col
'    If (IstImportModus%) Then
'        .Sort = 8
'    Else
'        .Sort = 1
'    End If
    .Sort = flexSortGenericAscending
    .col = 1
    .ColSel = .Cols - 1

    .Height = .RowHeight(0) * .Rows + 90
    If (para.Newline) Then
        MaxHe = wpara.WorkAreaHeight - 2400
    Else
        MaxHe = wpara.WorkAreaHeight - 150
    End If
    If (.Height > MaxHe) Then
'        .Height = wpara.WorkAreaHeight - 150
        .Height = MaxHe
        .Height = ((.Height - 90) \ .RowHeight(0)) * .RowHeight(0) + 90
    End If

    spBreite& = 0
    For i% = 0 To (.Cols - 1)
        spBreite& = spBreite& + .ColWidth(i%)
    Next i%
    .Width = spBreite& + 90

    With frmImportOrIdent
        .Caption = Titel$
        .Height = .flxEdit.Height + wpara.FrmCaptionHeight + 60
        .Width = .flxEdit.Width + 60
        .Left = ParentForm.Left + (ParentForm.Width - .Width) / 2
        .Top = ParentForm.Top + (ParentForm.Height - .Height) / 2
    End With

'    If (IstImportModus% = 0) And (IstAutIdemModus) Then
    If (IstImportModus% = 0) Then
        If (AutIdemNeu%) Then
            grenze# = OrgKp#
        ElseIf (.Rows > 6) Then
            For i% = 1 To 3
                unten# = unten# + xVal(.TextMatrix(i%, 7))
            Next i%
            unten# = unten# / 3
            For i% = (.Rows - 3) To (.Rows - 1)
                oben# = oben# + xVal(.TextMatrix(i%, 7))
            Next i%
            oben# = oben# / 3
            grenze# = xVal(.TextMatrix(1, 7)) + (oben# - unten#) / 3 'unteres Drittel
        Else
            If (.Rows <= 4) Then
                i% = 1
            Else
                i% = .Rows - 3
            End If
            grenze# = xVal(.TextMatrix(i%, 7))
        End If

        For i% = 1 To (.Rows - 1)
            AVP# = xVal(.TextMatrix(i%, 7))

            .row = i%
            .col = 7
            .RowSel = .row
            .ColSel = .col
            '.CellFontBold = (avp# > grenze#)
            If AnzGilt% < 3 Or AVP# = grenze# Then
              '.CellFontBold = (avp# > grenze#)
              .CellFontBold = False
              grenze# = AVP#
            Else
              .CellFontBold = True
                If (IstAutIdemModus) Then
                    .CellForeColor = vbRed
                End If
            End If
            If AnzGilt% <= 3 Then grenze# = AVP#
            If (AVP# <= grenze#) Then
                AnzGilt% = AnzGilt% + 1
            End If
        Next i%

        If (AutIdemNeu% = 0) And (AnzGilt% < 5) And (.Rows > 6) Then 'wenn weniger als 5 im 1. Drittel -> 5 zulassen
            For i% = (AnzGilt + 1) To 5
                .row = i%
                .col = 7
                .RowSel = .row
                .ColSel = .col
                .CellFontBold = False
            Next i%
            grenze# = xVal(.TextMatrix(5, 7))
            Do
                If (i% >= .Rows) Then Exit Do
                If (xVal(.TextMatrix(i%, 7)) = grenze#) Then
                    .row = i%
                    .col = 7
                    .RowSel = .row
                    .ColSel = .col
                    .CellFontBold = False
                    i% = i% + 1
                Else
                    Exit Do
                End If
            Loop
        End If

        .row = 1
    Else
'      .row = 1
'      .col = 7
'      .RowSel = .row
'      .ColSel = .col
'      .CellFontBold = True
'      .CellForeColor = vbRed
'
'      If .Rows > 2 Then
'        .row = 2
'      End If

        For i = 1 To .Rows - 1
            .row = i
            .col = 4
            If (.text <> "i") And (.CellForeColor = vbRed) Then
                .col = 7
                .CellForeColor = vbRed
            End If
        Next i
        .row = 1

    End If

    .col = 0
    .ColSel = .Cols - 1
    .Visible = True
End With

With frmImportOrIdent
    .cmdF2.Enabled = (IstImportModus% = 0)

    If (para.Newline) Then
        iAdd = wpara.NlFlexBackY
        iAdd2 = wpara.NlCaptionY

        With frmImportOrIdent.flxEdit
'            .ScrollBars = flexScrollBarNone
            .BorderStyle = 0
            .Width = .Width - 90
            .Height = .Height - 75
            .GridLines = flexGridFlat
            .GridLinesFixed = flexGridFlat
            .GridColorFixed = .GridColor
            .BackColor = wpara.nlFlexBackColor    'vbWhite
            .BackColorBkg = wpara.nlFlexBackColor    'vbWhite
            .BackColorFixed = wpara.nlFlexBackColorFixed   ' RGB(199, 176, 123)
            .BackColorSel = wpara.nlFlexBackColorSel  ' RGB(232, 217, 172)
            .ForeColorSel = vbBlack

            .Left = .Left + 2 * iAdd
            .Top = .Top + 2 * iAdd

            For i% = 1 To (.Rows - 1)
                .row = i

                .col = 1
                If (.CellBackColor = vbGreen) Then
                    .CellBackColor = wpara.nlGreen
                End If

                .col = 2
                If (.CellForeColor = vbRed) Then
                    .CellForeColor = vbWhite
'                    If (Knallrot) Then
'                        .CellBackColor = vbRed
'                    Else
                        .CellBackColor = wpara.nlRed
'                    End If
                End If

                .col = 4
                If (.CellForeColor = vbRed) Then
                    .CellForeColor = vbWhite
'                    If (Knallrot) Then
'                        .CellBackColor = vbRed
'                    Else
                        .CellBackColor = wpara.nlRed
'                    End If
                End If

                .col = 7
                If (.CellForeColor = vbRed) Then
                    .CellForeColor = vbWhite
'                    If (Knallrot) Then
'                        .CellBackColor = vbRed
'                    Else
                        .CellBackColor = wpara.nlRed
'                    End If
                End If
            Next i

            .row = 1
            .col = 0
            .ColSel = .Cols - 1
        End With

        With frmImportOrIdent
            .Width = .Width + 4 * iAdd
            .Height = .Height + 4 * iAdd

            .flxEdit.Top = .flxEdit.Top + iAdd2

            With .nlcmdOk
                .Init
                .Left = (frmImportOrIdent.ScaleWidth - (.Width * 2 + 300)) / 2
                .Top = frmImportOrIdent.flxEdit.Top + frmImportOrIdent.flxEdit.Height + iAdd + 600
                .Caption = frmImportOrIdent.cmdOk.Caption
                .TabIndex = frmImportOrIdent.cmdOk.TabIndex
                .Enabled = frmImportOrIdent.cmdOk.Enabled
                .Default = frmImportOrIdent.cmdOk.Default
                .Cancel = frmImportOrIdent.cmdOk.Cancel
                .Visible = True
            End With
            .cmdOk.Visible = False

            With .nlcmdEsc
                .Init
                .Left = frmImportOrIdent.nlcmdOk.Left + .Width + 300
                .Top = frmImportOrIdent.nlcmdOk.Top
                .Caption = frmImportOrIdent.cmdEsc.Caption
                .TabIndex = frmImportOrIdent.cmdEsc.TabIndex
                .Enabled = frmImportOrIdent.cmdEsc.Enabled
                .Default = frmImportOrIdent.cmdEsc.Default
                .Cancel = frmImportOrIdent.cmdEsc.Cancel
                .Visible = True
            End With
            .cmdEsc.Visible = False

            .Height = .nlcmdOk.Top + .nlcmdOk.Height + wpara.FrmCaptionHeight + 450

            Call wpara.NewLineWindow(frmImportOrIdent, .nlcmdOk.Top)
            RoundRect .hdc, (.flxEdit.Left - iAdd) / Screen.TwipsPerPixelX, (.flxEdit.Top - iAdd) / Screen.TwipsPerPixelY, (.flxEdit.Left + .flxEdit.Width + iAdd) / Screen.TwipsPerPixelX, (.flxEdit.Top + .flxEdit.Height + iAdd) / Screen.TwipsPerPixelY, 20, 20
        End With
    Else
        .nlcmdOk.Visible = False
        .nlcmdEsc.Visible = False
    End If

    .Left = frmAction.Left + (frmAction.Width - .Width) / 2
    .Top = frmAction.Top + (frmAction.Height - .Height) / 2

    .Show 1
End With

If (EditErg%) Then
    ind% = InStr(EditTxt$, vbTab)
    pzn$ = Left$(EditTxt$, ind% - 1)

    Call NeuerArtikel(row%, pzn$)
    Call PaintArtikel(row%)
End If


If (NeueKbvNr%) Then
    iAnzZ% = AnzRezeptArtikel% - 1      'GS 16.05.02
    If (iAnzZ% > 6) Then
        iAnzZ% = 6
    End If
    iTabIndex% = AktTabIndex%
    For i% = 0 To iAnzZ%
        RezArtikel(i%).AutIdem = 0
        RezArtikel(i%).N0 = 0
        SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + RezArtikel(i%).pzn
        'Set TaxeRec = TaxeDB.OpenRecordset(SQLStr$)
        On Error Resume Next
        TaxeRec.Close
        Err.Clear
        On Error GoTo DefErr
        TaxeRec.Open SQLStr, taxeAdoDB.ActiveConn
        Call CheckAutIdem(i%)
        Call PaintArtikel(i%)
    Next i%
    Call TabIndexChange(iTabIndex%, False)
End If

Call DefErrPop
End Sub

Function HerstRabattBrutto(t As ADODB.Recordset) As Double
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("HerstRabattBrutto")
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
If (Err.Number = (999 + vbObjectError)) Then End
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
'On Error GoTo 0
'Err.Raise 999 + vbObjectError, "DSK.VBP", "Fehler in DLL"
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'DrS 22.4.2013: Leider gibt es so bescheuerte Vereinbarungen, dass der normale Herstellerrabatt NETTO vom Bruttopreis abgezogen werden muss,
'der 130B-Rabatt aber brutto vom Bruttopreis abgezogen wird.
'Im taxefeld herstRabatt sind alle Rabatte netto addiert, nur der 130B-Rabatt wird brutto addiert --> passt!

Dim MwCode As Integer
Dim r As Double, r130B As Double

'MwCode = 2
'If t!MwStKZ <> 0 Then MwCode = 1
r = dCheckNull(t!HerstRabatt) / 100#
'r130B = dCheckNull(t!HerstRabattPrivat130) / 100#
'If r130B > 0 Then
'  r = r - r130B
'  r130B = fnx(r130B * ((100# + MwProzent(MwCode)) / 100#))
'  r = r + r130B
'End If

Dim Rabwert_130a_2_SGB As Double
Rabwert_130a_2_SGB = 0
On Error Resume Next
Rabwert_130a_2_SGB = CheckNullLong(t!Rabwert_130a_2_SGB) / 100#
On Error GoTo DefErr
If (Rabwert_130a_2_SGB > r) Then
    r = Rabwert_130a_2_SGB
End If

HerstRabattBrutto = FNX(r)
Call DefErrPop
End Function

Sub MachMCImportKZ(sTaxerec As ADODB.Recordset, i As Integer, SpalteImp As Integer)

Dim OriginalAVP As Double, ArtVK As Double, FB As Double
Dim SQLStr As String
Dim t As New ADODB.Recordset

m(i).text(SpalteImp) = ""
If sTaxerec!OriginalPZN = sTaxerec!pzn Then
  m(i).text(SpalteImp) = "i."
ElseIf sTaxerec!ImportGruppeNr > 0 Then
  m(i).text(SpalteImp) = "I."
  If sTaxerec!OriginalPZN > 0 Then
    ArtVK = sTaxerec!vk
    
    SQLStr = "SELECT * FROM TAXE WHERE PZN = " + CStr(sTaxerec!OriginalPZN)
    t.Open SQLStr, taxeAdoDB.ActiveConn
    If Not t.EOF Then
      OriginalAVP = t!vk
        If IstGünstigerImport(ArtVK, OriginalAVP) Then
          m(i).text(SpalteImp) = Left(m(i).text(SpalteImp), 1)
        End If
    End If
  End If
End If


End Sub

Function IstGünstigerImport(vk As Double, originalvk As Double) As Boolean

Dim erg As Boolean

erg = False
If originalvk <= 100 Then    'bis 100€ mind. 15% billiger
  If originalvk * 0.85 >= vk Then erg = True
ElseIf originalvk <= 300 Then    'bis 300€ mind. 10€
  If originalvk - 15 >= vk Then erg = True
Else    'mind. 5%
  If originalvk * 0.95 >= vk Then erg = True
End If

IstGünstigerImport = erg
End Function

Sub BelegeSatz()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("BelegeSatz")
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
If (Err.Number = (999 + vbObjectError)) Then End
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
'On Error GoTo 0
'Err.Raise 999 + vbObjectError, "DSK.VBP", "Fehler in DLL"
DefErrEnd:
Dim i%
Dim MCSpalte_ImportKZ As Integer
Dim MCSpalte_NGrösse As Integer
Dim MCSpalte_Herst As Integer
Dim Ngr As Byte
Dim MKV As Byte
Dim ZzFreiKz As Boolean
Dim ArtVK As Double
Dim OriginalAVP As Double

MCSpalte_ImportKZ = 5
MCSpalte_NGrösse = 6
MCSpalte_Herst = 7


For i = 0 To UBound(m) '- 1
    m(i).fertig = False
    m(i).RabattAutidem = False
    m(i).AltPG = False
    m(i).KdStatus = 0
    m(i).satz = i + 1 ' satz
    
        ArtL = ArtLRec(i)
      m(i).AusgangsArtikel = False
        
        amatch(i).ssatz = 0
        amatch(i).Streichung = False
        amatch(i).StreichungohneLager = False
        amatch(i).Artikel_vc = ""
        Ngr = 0
        MKV = False
        ZzFreiKz = False
        m(i).text(11) = ""
        m(i).RoteLinie = False
        ArtVK = 0
        If ArtL.pzn = 0 Then
          ArtL.pzn = 9999999
        Else
          If ArtL.RoteLinie = 1 Then
            m(i).RoteLinie = True
          End If
          SQLStr = "SELECT * FROM Taxe WHERE PZN=" + CStr(ArtL.pzn)
            On Error Resume Next
            TaxeRec.Close
            Err.Clear
            On Error GoTo DefErr
            TaxeRec.Open SQLStr, taxeAdoDB.ActiveConn
          If Not TaxeRec.EOF Then
            amatch(i).tsatz = TaxeRec.Bookmark
            Call TaxeMDB2Type(i)
            amatch(i).tname = TaxeRec!Name
            amatch(i).SortierName = TaxeRec!SortierName
            Call Taxe2ast(CStr(ArtL.pzn))
            If TaxeRec!ArtStatus = "S" Then ast.vc = "S"
'            amatch(i).Artikel = ast
            amatch(i).Artikel_zuza = ast.zuza
            amatch(i).Artikel_vc = ast.vc
            amatch(i).Artikel_PZN = ast.pzn
            amatch(i).Artikel_meng = ast.meng
            amatch(i).Artikel_meh = ast.Meh
            amatch(i).Artikel_herst = ast.herst

            
            amatch(i).pzn = TaxeRec!pzn
            Ngr = Val(Mid(TaxeRec!NGrösse, 2, 1))
            If Ngr = 0 And Left(TaxeRec!NGrösse, 2) = "N0" Then Ngr = 9
            If TaxeRec!MehrkostenverzichtKZ > 0 Then MKV = True
            If TaxeRec!ZzFreiFBKz > 0 Or TaxeRec!ZzFreiTeststreifenKz > 0 Then ZzFreiKz = True
            If Val(TaxeRec!AlternativPackung) > 0 Then m(i).AltPG = True
            m(i).text(11) = TaxeRec!DarForm
            ArtVK = TaxeRec!vk
          End If
        End If
        
        m(i).text(9) = ""
        
        SQLStr$ = "SELECT * FROM ARTIKEL WHERE PZN = " + PznString(ArtL.pzn)
        FabsErrf = Artikel.OpenRecordset(ArtikelAdoRec, SQLStr)
        amatch(i).asatz = Abs(FabsErrf = 0)
        If amatch(i).asatz > 0 Then
          amatch(i).ssatz = -1
'          GetDataRecord dbAss, Abs(amatch(i).ssatz)
           m(i).text(9) = CStr(ass.poslag)
'          amatch(i).Artikel = ast
          If ArtikelAdoRec!LagerKz <> 0 Then
            If ass.halt = "S" Then amatch(i).Streichung = True
            If ass.halt = "S" And ass.poslag <= 0 Then amatch(i).StreichungohneLager = True
          Else
            amatch(i).ssatz = 0
          End If
          
        amatch(i).Artikel_zuza = ast.zuza
        amatch(i).Artikel_vc = ast.vc
        amatch(i).Artikel_PZN = ast.pzn
        amatch(i).Artikel_meng = ast.meng
        amatch(i).Artikel_meh = ast.Meh
        amatch(i).Artikel_herst = ast.herst
          
          'so würde langer Taxe-Name abgeschnitten: wenn erste 28 Stellen gleich, dann bleibt Taxetext
          If Not IsEmpty(amatch(i).tsatz) And Not IsNull(amatch(i).tsatz) Then
            If Left(amatch(i).tname, Len(ast.kurz)) <> ast.kurz Then amatch(i).tname = ast.kurz
          End If
          If Ngr = 0 Then Ngr = N_Grösse(CSng(amatch(i).Artikel_zuza))
        End If
        m(i).text(1) = ""
        
'        If BesorgtMarkierung And amatch(i).ssatz = 0 Then
'          If FrüherBesorgt(amatch(i).PZN) Then
'            m(i).Text(1) = m(i).Text(1) + MC_Besorgt
'          End If
'        End If
'        If AktKundensatz > 0 Then Call CheckPharmBedenken(m(i).Text(1), amatch(i).PZN)
'        If ZuletztAbgegebenPZN > 0 And ZuletztAbgegebenPZN = amatch(i).PZN Then
'          m(i).Text(1) = MC_Zuletztabgegeben + m(i).Text(1)
'        End If
        
        
'        If PartnerDbDa Then
'
'          FremdPznRec.Seek "=", ArtL.PZN
'          If (FremdPznRec.NoMatch = False) Then
'            'm(z).Text(1) = m(z).Text(1) + Chr(187)
'            m(i).Text(1) = m(i).Text(1) + MC_Partner
'          End If
'        End If
'        If SPrmDBda Then      'PVS - Prämienartikel
'
'          sprm.index = "Unique"
'          sprm.Seek "=", ArtL.PZN
'          If Not sprm.NoMatch Then
'            'm(z).Text(1) = m(z).Text(1) + Chr(137)
'            m(i).Text(1) = m(i).Text(1) + MC_SonderPrämie  ''Daumen
'          End If
'        End If
'        If ArtL.GhVerfügbar <> 0 Then
'          m(i).Text(1) = "ü" + Trim(m(i).Text(1)) 'Hakerl
'        End If
        
'        If amatch(i).ssatz <> 0 Then
'          Pool.Seek "=", ArtL.PZN
'          If Not Pool.NoMatch Then
'            On Error Resume Next
'            If m(i).Text(1) = "" Then
'              m(i).Text(1) = MC_Pool
'            Else
'            If Left(m(i).Text(1), 1) = MC_Besorgt Then
'              m(i).Text(1) = MC_Besorgt + MC_Pool
'            Else
'              m(i).Text(1) = MC_Pool      'Pool-Zeichen ersetzt andere
'            End If
'            End If
'            On Error GoTo DefErr
'          End If
'        End If
        If ArtL.AusgangsArtikel Then
          m(i).AusgangsArtikel = True
        End If
        
'        If MARS And MarsArtikelGesperrt(PznString(ArtL.PZN), True) Then m(i).Text(1) = MC_MarsSperr
        
        If amatch(i).Streichung Or amatch(i).Artikel_vc = "S" Then
            m(i).text(1) = Trim(m(i).text(1)) + "V"    'V wird als Kreus dargestellt
        End If
        m(i).text(2) = amatch(i).tname
        m(i).text(3) = Trim(amatch(i).Artikel_meng)
        If Not TaxeRec.EOF Then
          If TaxeRec!pzn = Val(amatch(i).Artikel_PZN) Then m(i).text(3) = TaxeRec!menge
        End If
        m(i).text(4) = amatch(i).Artikel_meh
        m(i).text(MCSpalte_Herst) = amatch(i).Artikel_herst
          m(i).text(MCSpalte_NGrösse) = ArtL.kz1 + ArtL.kz2
        m(i).text(MCSpalte_ImportKZ) = ""  '(Ii)
        If ArtL.rabKZ = "r" Then
          m(i).RabattAutidem = True
        End If
        
'        If Not IsEmpty(amatch(i).tsatz) And Not IsNull(amatch(i).tsatz) Then
'          m(i).text(MCSpalte_ImportKZ) = amatch(i).taxesatz.Import
'          If (m(i).text(MCSpalte_ImportKZ) = "I" Or m(i).text(MCSpalte_ImportKZ) = "i") Then
'            m(i).text(MCSpalte_ImportKZ) = m(i).text(MCSpalte_ImportKZ) + "."
'            If TaxeRec!pzn <> TaxeRec!OriginalPZN And TaxeRec!OriginalPZN > 0 Then
'
'              SQLStr = "SELECT * FROM Taxe WHERE PZN=" + CStr(TaxeRec!OriginalPZN)
'              On Error Resume Next
'              TaxeRec.Close
'              Err.Clear
'              On Error GoTo DefErr
'              TaxeRec.open SQLStr, taxeAdoDB.ActiveConn
'              If Not TaxeRec.EOF Then
'                OriginalAVP = TaxeRec!vk
''                If Not RV2019 And TaxeRec!HerstRabattKZ <> 0 And Val(Rezeptnummer) > 0 Then
''                  OriginalAVP = OriginalAVP - (HerstRabattBrutto(TaxeRec) * 100#)
''                End If
''                If RV2019 Then
'                  If IstGünstigerImport(ArtVK, OriginalAVP) Then
'                    m(i).text(MCSpalte_ImportKZ) = Left(m(i).text(MCSpalte_ImportKZ), 1)
'                  End If
''                ElseIf (OriginalAVP / 100#) - 15 >= Val(ArtL.Preis) Or (OriginalAVP / 100#) * 0.85 >= Val(ArtL.Preis) Then
''                  m(i).Text(MCSpalte_ImportKZ) = Left(m(i).Text(MCSpalte_ImportKZ), 1)
''                End If
'              End If
'            End If
'          End If
'        End If
    If Not TaxeRec.EOF Then
      Call MachMCImportKZ(TaxeRec, i, MCSpalte_ImportKZ)
    End If
        
        m(i).text(8) = ""
        If Ngr > 0 Then
          If Ngr <= 9 Then
            If Ngr = 9 Then
              m(i).text(8) = "N0"
            Else
              m(i).text(8) = "N" + CStr(Ngr)
            End If
            If MKV Then m(i).text(8) = "." + m(i).text(8)
          End If
        End If
        
        m(i).text(10) = ArtL.Preis
        If ZzFreiKz Then m(i).text(10) = "." + m(i).text(10)
Next

Call DefErrPop
End Sub

Sub TaxeMDB2Type(pos%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("TaxeMDB2Type")
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
If (Err.Number = (999 + vbObjectError)) Then End
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
'On Error GoTo 0
'Err.Raise 999 + vbObjectError, "DSK.VBP", "Fehler in DLL"
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim s As String
Dim wg As String
Dim NeuWg As Integer
Dim i As Integer
Dim t As TaxeType

On Error Resume Next      'einige neue Taxerec!Felder sind ev. noch nicht vorhanden

With amatch(pos).taxesatz

    .vc = "V"
    If TaxeRec!ArtStatus = "S" Then .vc = "S"
    .herst_nr = TaxeRec!HerstellerNr
    .pzn = PznString(TaxeRec!pzn)
    .meng = TaxeRec!menge
    .Meh = TaxeRec!einheit
    .kurz = Left(TaxeRec!Name, 28)
    .Name = TaxeRec!Name
    If TaxeRec!RufKZ > 0 Then
      .kas = "RW "
    ElseIf TaxeRec!Negliste > 0 Then
      .kas = "*  "
    Else
      .kas = "   "
    End If
    .aep = TaxeRec!EK / 100#
    .AVP = TaxeRec!vk / 100#
    s = TaxeRec!gueltig
    .gc = Mid("0123456789OEZ", Val(Mid(s, 3, 2)) + 1, 1) + Right(s, 1)
    If TaxeRec!BtmKz = 1 Then
      .SuchtG = True
    Else
      .SuchtG = False
    End If
    .mw = "2"
    If TaxeRec!MwStKZ <> 0 Then .mw = "1"
    .lic = " "
    If TaxeRec!UeberGH = 2 Then
      .lic = "K"
    ElseIf TaxeRec!UeberGH = 1 Then
      .lic = "D"
    End If
    If TaxeRec!VerfallArt = 1 Then
      .abl = "A"
    Else
      .abl = " "
    End If
    
    .abgschl = TaxeRec!Abgabebest
    
    If Val(TaxeRec!AlternativPackung) > 0 Then .Generika = "g"
    If TaxeRec!Lagerung > 0 Then .lac = Format(TaxeRec!Lagerung, "0")
    
    wg = TaxeRec!Warengruppe
    .appli = False
    'If Val(wg) >= 10 And Val(wg) <= 19 Then
    ''  .appli = True
    '  wg = Str$(Val(wg) - 10)
    'End If
    'GoSub WGconvert
    .wg = wg
    
    .FB = TaxeRec!FESTBETRAG / 100#
    .Hilfsmittel = False
    If TaxeRec!HilfsmittelKz > 0 Then .Hilfsmittel = True
    
    .Verbandmittel = False
    'If TaxeRec!VerbandKz <> 0 Then .Verbandmittel = True
    If .wg = 4 Then .Verbandmittel = True
    .Negativliste = False
    If TaxeRec!Negliste <> 0 Then .Negativliste = True
    
    .Import = " "
    If TaxeRec!OriginalPZN = TaxeRec!pzn Then
      .Import = "i"
    ElseIf TaxeRec!ImportGruppeNr <> 0 Then
      .Import = "I"
    End If
    
    '.zuza = ZuzaInt(TaxeRec!Zuzahlung)
    .herst = TaxeRec!HerstellerKB
    .Preislinie = TaxeRec!Preislinie / 100#
    
    If TaxeRec!ZzFreiTeststreifenKz > 0 Or TaxeRec!ZzFreiFBKz > 0 Then .PreisGrün = True
End With

Call DefErrPop: Exit Sub

'WGconvert:
'
'NeuWg = 0
's = "," + Mid(Str(Val(wg)), 2) + ","
'For i = 1 To 9
'  If NeuWg = 0 Then
'    If InStr(WGconv(i), s) > 0 Then NeuWg = i
'  End If
'Next i
'If NeuWg = 0 Then NeuWg = 9
'wg = Right(Str(NeuWg), 1) + " "
'Return

Call DefErrPop
End Sub

Function ZuzaStr(ByVal ZuzaBits As Single) As String

'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("ZuzaStr")
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
If (Err.Number = (999 + vbObjectError)) Then End
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
'On Error GoTo 0
'Err.Raise 999 + vbObjectError, "DSK.VBP", "Fehler in DLL"
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim z As Integer
Dim z1 As Integer
Dim z2 As Integer
Dim Zuzahl As String

If ZuzaBits < 0 Then ZuzaBits = ZuzaBits + 65536!
'* 1 x shift right
z = ZuzaBits \ 2
For z1 = 1 To 5
  z2 = (z Mod 8)
  If z2 = 7 Then z2 = 9
  '* 3 x shift right
  z = z \ 8
  Zuzahl = Right(Str$(z2), 1) + Zuzahl
Next z1

ZuzaStr = Zuzahl

Call DefErrPop
End Function


Function N_Grösse(ZuzaBits As String) As Byte
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("N_Grösse")
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
If (Err.Number = (999 + vbObjectError)) Then End
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
'On Error GoTo 0
'Err.Raise 999 + vbObjectError, "DSK.VBP", "Fehler in DLL"
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim z1 As Integer
Dim z2 As Integer
Dim Zuzahl As String

Zuzahl = ZuzaStr(ZuzaBits)
For z1 = 1 To 5
  z2 = Val(Mid(Zuzahl, z1, 1))
  If z2 = 3 Then
    '* Zuschlag 1 Arzneimittel
    N_Grösse = 1
  ElseIf z2 = 4 Then
    '* Zuschlag 1 Verbandstoffe
    N_Grösse = 1
  ElseIf z2 = 5 Then
    '* Zuschlag 2 Arzneimittel
    N_Grösse = 2
  ElseIf z2 = 6 Then
    '* Zuschlag 3 Arzneimittel
    N_Grösse = 3
  ElseIf z2 = 9 Then
    N_Grösse = 9
  End If
Next z1

Call DefErrPop
End Function

Sub BelegeWork(IstImportModus%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("BelegeWork")
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
If (Err.Number = (999 + vbObjectError)) Then End
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
'On Error GoTo 0
'Err.Raise 999 + vbObjectError, "DSK.VBP", "Fehler in DLL"
DefErrEnd:
Dim i%, z%
Dim spBreite&, MaxHe&
Dim RabattVertragDa As Integer
Dim AusgangsArtikelDa As Integer
Dim RedLineDa As Integer

Dim ArtikelListenArt%
ArtikelListenArt = IIf(IstImportModus, LISTE_IMPORT, ArtikelListenArt = LISTE_AUTIDEM)

Dim MCSpalte_ImportKZ As Integer
Dim MCSpalte_NGrösse As Integer
Dim MCSpalte_Herst As Integer
MCSpalte_ImportKZ = 5
MCSpalte_NGrösse = 6
MCSpalte_Herst = 7


'  coKasseBackColor = RGB(220, 220, 220)
'  coPreisForeColor = RGB(135, 61, 52)
'  coKennzeichenGelb = RGB(201, 123, 58)
  coRot = RGB(135, 61, 52)
'  coKnallrot = coRot
'  s = Space(255)
'  i = GetPrivateProfileString("AlleKassen", "Knallrot", "", s, 255, INI_DATEI)
'  If i > 0 Then
'    If Val(Left(s, 1)) > 0 Then coKnallrot = vbRed
'  End If
'  coGelb = RGB(201, 123, 58)
'  coWawiGelb = RGB(255, 255, 100)
  coPreisGrün = RGB(206, 185, 23)
'  coNewlineGridKopf = RGB(112, 91, 53)
  coNewlineGridHighlight = RGB(199, 176, 123)
'  coNewlineMygridHighlight = RGB(217, 130, 24)     'RGB(206, 185, 83)
'  coNewlineGridInfoBack = RGB(232, 217, 172)
'  coNewlineGridBack = RGB(255, 246, 217)
'  coNewlinePicInfoBack = RGB(233, 233, 233)
'  coNewlinePicKasseInaktiv = RGB(115, 115, 115)
'  coNewlineStammKundeAktiv = RGB(216, 161, 84)
'  coNewlineStammKundeInaktiv = coNewlinePicKasseInaktiv
'  coNewlineSattesDunkelgrau = RGB(64, 64, 64)
'  coNewlineKassenRez = RGB(221, 191, 153)
  coNewlineHinweis = RGB(255, 131, 62)
  coAusgangsartikel = coNewlineGridHighlight

    Dim l&
    Dim h$
    h$ = Space(50)
    l& = GetPrivateProfileString(UserSection, "FarbeAusgangsartikel", "", h$, 50, "\user\vkpara.ini")
    If (l > 0) Then
       coAusgangsartikel = Val(Left(h, l))
    End If


With frmImportOrIdent.flxEdit
    .Left = 0
    .Top = 0

    .Rows = 2
    .Cols = 12
    .FixedRows = 0
'    .FormatString = "<Pzn|<Name|>Menge|<Eh|^|^Kz|<Hersteller|>Preis||"
'    .Rows = 1
'    .Cols = 11
'
'    .ColWidth(0) = frmImportOrIdent.TextWidth(String(10, "9"))
'    .ColWidth(1) = frmImportOrIdent.TextWidth(String(28, "X"))
'    .ColWidth(2) = frmImportOrIdent.TextWidth(String(6, "X"))
'    .ColWidth(3) = frmImportOrIdent.TextWidth(String(3, "X"))
'    .ColWidth(4) = frmImportOrIdent.TextWidth("Ww")
'    .ColWidth(5) = frmImportOrIdent.TextWidth("Ww")
'    .ColWidth(6) = frmImportOrIdent.TextWidth(String(9, "X"))
'    .ColWidth(7) = frmImportOrIdent.TextWidth("9999999.99")
'    .ColWidth(8) = 0
'    .ColWidth(9) = 0    'frmImportOrIdent.TextWidth("Ww")
'    .ColWidth(10) = wpara.FrmScrollHeight

    '.FormatString = "|<Kz|<Name|>Menge|<Eh|<5|<6|<7|<8|<9|<10|<11|"
    For i = 0 To 10
        .ColWidth(i) = frmImportOrIdent.TextWidth(String(10, "X"))
    Next i
'    .ColWidth(0) = 0
    .ColWidth(2) = frmImportOrIdent.TextWidth(String(40, "X"))
    .ColWidth(9) = 0
    .ColWidth(11) = frmImportOrIdent.TextWidth(String(3, "X"))
    .ColAlignment(3) = flexAlignRightCenter
    .ColAlignment(10) = flexAlignRightCenter

    .Rows = UBound(m) + 1
    .Height = .RowHeight(0) * (.Rows + 1) + 90
    MaxHe = wpara.WorkAreaHeight - 2400
    If (.Height > MaxHe) Then
    '        .Height = wpara.WorkAreaHeight - 150
        .Height = MaxHe
        .Height = ((.Height - 90) \ .RowHeight(0)) * .RowHeight(0) + 90
    End If
    
    spBreite& = 0
    For i% = 0 To (.Cols - 1)
        spBreite& = spBreite& + .ColWidth(i%)
    Next i%
    .Width = spBreite& + 90
    
End With

With frmImportOrIdent
    .Caption = "ABCDE"  ' Titel$
    .Height = .flxEdit.Height + wpara.FrmCaptionHeight + 60
    .Width = .flxEdit.Width + 60
    .Left = ParentForm.Left + (ParentForm.Width - .Width) / 2
    .Top = ParentForm.Top + (ParentForm.Height - .Height) / 2
End With


  
Call McMalenAnfang

With frmImportOrIdent.flxEdit
    For z = 0 To UBound(m) '- 1
    '    mgdGrid(Status).row = z
         
    '        If amatch(z).asatz > 0 Then
    '          'If Not MARS And (amatch(z).Artikel.lac = "!" Or (RowaCompNr >= 0 And InStr(RowaLac + RowaLacKuehl, amatch(z).Artikel.lac) > 0)) Then
    '          If (amatch(z).Artikel.lac = "!" Or (RowaCompNr >= 0 And InStr(RowaLac + RowaLacKuehl, amatch(z).Artikel.lac) > 0)) Then
    '            mgdGrid(Status).col = 0 'ganze Zeile
    '            'mgdGrid(Status).BackColor = vbInactiveTitleBar
    '            mgdGrid(Status).BackColor = FarbeAnfordernBack
    '            'mgdGrid(Status).ForeColor = vbWindowBackground
    '            mgdGrid(Status).ForeColor = FarbeAnfordernFore
    '          End If
    '        End If
          
            .row = z '+ 1
            .TextMatrix(.row, 0) = amatch(z).Artikel_PZN
            .col = 1 'Kennzeichen-Spalte
            .CellFontName = "WingDings"
            If amatch(z).ssatz <> 0 And Not amatch(z).StreichungohneLager Then
              .col = 0
              .CellFontBold = True
            End If
            If (Mid(m(z).text(6), 2, 1) = ".") Or (ArtikelListenArt = LISTE_IMPORT And Left(m(z).text(MCSpalte_ImportKZ), 1) = "o") Then
              .col = MCSpalte_ImportKZ
              .CellFontBold = True
              .col = 10
              .CellFontBold = True
                .CellForeColor = vbWhite
                .CellBackColor = coRot
            End If
            
            .col = MCSpalte_ImportKZ
            If Mid(m(z).text(MCSpalte_ImportKZ), 2, 1) = "." Then
                .CellBackColor = coRot
                .CellForeColor = vbWhite
                .CellFontBold = True
            Else
              .CellFontBold = False
            End If
            .CellFontName = "Times New Roman"
          
            If m(z).RabattAutidem Then
              .col = 2
              .CellBackColor = coPreisGrün
              .col = 3
              .CellBackColor = coPreisGrün
              RabattVertragDa = z
            End If
            If m(z).AusgangsArtikel Then
              .col = 2
              .CellBackColor = coAusgangsartikel
              AusgangsArtikelDa = z
            End If
'            If (ArtikelListenArt = LISTE_IMPORT Or ArtikelListenArt = LISTE_AUTIDEM) Then   'And RV2019 Then
              If m(z).RoteLinie Then
'                mgdGrid(Status).RoteLinie = z
                RedLineDa = z
              End If
'            End If
          
          For i = 1 To .Cols - 1
            .col = i
            If i = 1 Then
                If Trim(m(z).text(1)) > "" Then
'                  If Left(m(z).Text(1), 1) = MC_Besorgt Then
'                    'mgdGrid(Status).Text = Mid(m(z).Text(1), 2)
'                    mgdGrid(Status).col = 2
'                    mgdGrid(Status).FontItalic = True
'                    'mgdGrid(Status).FontBold = True
'                    mgdGrid(Status).FontBold = False
'                    mgdGrid(Status).col = 1
'                    If Newline Then mgdGrid(Status).ForeColor = coNewlineHinweis
'                  End If
                  .text = m(z).text(1)
'                  If InStr(m(z).Text(1), MC_Pool) > 0 Or InStr(m(z).Text(1), MC_MarsSperr) > 0 Then    'nur bei Pool-Artikeln und Mars-Sperrliste
'                    mgdGrid(Status).BackColor = coKennzeichenGelb
'                  End If
                End If
                
                If amatch(z).Artikel_vc = "S" Then
                  'mgdGrid(Status).Text = Trim(mgdGrid(Status).Text) + "V"
                  'mgdGrid(Status).ForeColor = coRot
                  'mgdGrid(Status).ForeColor = coRot
                  'If NewLine Then
                    .CellForeColor = vbBlack
                    'mgdGrid(Status).FontBold = True
                  'End If
                ElseIf amatch(z).Streichung Then
                    .CellForeColor = coNewlineHinweis
                End If
            Else
              If i = 11 Then
'                If mgdGrid(Status).ColWidth(11) = 1 Then Exit For
                Exit For
              End If
              
                    .text = CStr(m(z).text(i))
                    If ArtikelListenArt = LISTE_IMPORT And i = 6 Then
                      If Mid(.text, 2, 1) = "." Then
                        .text = Left(.text, 1) + " " + Mid(.text, 3)
                      End If
                    End If
                  If i = 2 Then
                      .text = FirstLettersUcase(CStr(m(z).text(i)))
                  End If
            End If
            If (i = 3) Then
              If m(z).AltPG And Not m(z).RabattAutidem Then
                .CellBackColor = coRot
                .CellForeColor = vbWhite     '1.1.94
              End If
            End If
            If (i = MCSpalte_ImportKZ Or i = MCSpalte_NGrösse) Then
              If i = MCSpalte_ImportKZ Then .text = Trim(.text)
              If Right(.text, 1) = "r" Then
                .text = Left(.text, Len(.text) - 1)
              End If
              If Right(Trim(.text), 1) = "." Then
                .text = Left(.text, Len(Trim(.text)) - 1)
              End If
              '1.2.53 billigste 3 unter Newline nicht sichtbar -> Soalte mit NGröße Hintergrund grün
              If i = MCSpalte_NGrösse And Mid(m(z).text(6), 2, 1) = " " Then
                .col = 8
                .CellBackColor = coPreisGrün
                .col = 6
              End If
            ElseIf (i = 8) Then   'Mehrkostenverzicht -> gelb
              If Left(m(z).text(i), 1) = "." Then
                'mgdGrid(Status).BackColor = coGelb
                .text = Mid(m(z).text(i), 2)
              End If
            ElseIf i = 10 Then  'ZuZaverzicht -> grün
              If Left(m(z).text(i), 1) = "." Then
                .CellBackColor = coPreisGrün
                .CellForeColor = vbBlack
                .text = Mid(m(z).text(i), 2)
              End If
            End If
          Next i
'        ElseIf (ArtikelListenArt = LISTE_IMPORT Or ArtikelListenArt = LISTE_AUTIDEM) Then   'And RV2019 Then
'          If RedLineDa = 0 Then
''            mgdGrid(Status).RoteLinie = z
'            RedLineDa = z
'          End If
'        End If
          
        
'        If Not MusterSuche Then GoSub CheckEvents
      Next z
              
    If (RedLineDa = 0) Then
'                mgdGrid(Status).RoteLinie = z
      RedLineDa = .Rows - 1
      .Rows = .Rows + 1
      .TextMatrix(.Rows - 1, 2) = "Preislinie"
        .FillStyle = flexFillRepeat
        .row = .Rows - 1
        .col = 0
        .RowSel = .row
        .ColSel = .Cols - 1
        .CellBackColor = vbRed
        .FillStyle = flexFillSingle
    Else
        .AddItem vbTab + vbTab + "Preislinie", RedLineDa
        .FillStyle = flexFillRepeat
        .row = RedLineDa
        .col = 0
        .RowSel = .row
        .ColSel = .Cols - 1
        .CellBackColor = vbRed
        .FillStyle = flexFillSingle
    End If
        
'        z = mgdGrid(Status).SelectedRow(1)
'        If z = 0 Then z = 1
'        If Status = ST_ARTLISTE And (ArtikelListenArt = LISTE_AUTIDEM Or ArtikelListenArt = LISTE_IMPORT) And AusgangsArtikelDa > 0 Then
'          If AusgangsArtikelDa <= RedLineDa And (RabattVertragDa = 0 Or RabattVertragDa >= AusgangsArtikelDa) Then z = AusgangsArtikelDa
'        End If
'        mgdGrid(Status).SelectRows(z, z) = True
'
'        If Not (Suchanzeige) And Not Unfertig Then Call AktiveZeile(z)
'    '    If Status = ST_KUNDEN Or Status = ST_BEZÜGE Then Call ZeigeHinweiseAE(4)      'nur Info über MC-Art
'      'mgdGrid(Status).Autoredraw = True
'    '  mgdGrid(Status).Refresh
'      mgdGrid(Status).Redraw
'
'        z = mgdGrid(Status).SelectedRow(1)
'        If z = 0 Then
'          z = 1
'          mgdGrid(Status).SelectRows(z, z) = True
'        End If
'        If Status = ST_ARTLISTE Or Status = ST_INHALTSSTOFF Then mgdGrid(Status).Redraw
'
'        If Not Unfertig Then Call AktiveZeile(z)
End With
  
Call DefErrPop
End Sub

Sub McMalenAnfang()
With frmImportOrIdent.flxEdit
    .row = 0 'alle Zeilen
    .col = 0 'alle Spalten
'    mgdGrid(Status).RoteLinie = 0
    .text = ""
    .CellFontBold = False
    .CellFontUnderline = False
    .CellFontItalic = False
      .CellForeColor = vbBlack
      .CellBackColor = coNewlineGridBack
    
    .row = 0 'alle Zeilen
    .col = 1 'Kennzeichen-Spalte
    .CellFontName = "WingDings"
End With

End Sub



'Sub ImporteOrAutIdem(IstImportModus%, Optional IstAutIdemModus% = True)
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Call DefErrFnc("ImporteOrAutIdem")
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
'Dim i%, ind%, row%, AnzGilt%, AutIdemNeu%, Sonderregelnr%, NeueKbvNr%, iAnzZ%, iTabIndex%
'Dim iAdd%, iAdd2%
'Dim spBreite&, MaxHe&
'Dim OrgAvp#, OrgKp#, AVP#, avp2#, unten#, oben#, grenze#, BtmDerivatMenge#, FESTBETR#
'Dim s$, SQLStr$, kz$, OrgPZN$, pzn$, Titel$, SqlSonderregelNr$, h$, RabKz$
'Dim iRec As New ADODB.Recordset
''Dim sTaxeRec As Recordset
'Dim sTaxeRec As New ADODB.Recordset
'Dim tRec2 As New ADODB.Recordset
'
'If (AktTabIndex% < 3) Then
'    Call DefErrPop: Exit Sub
'End If
'
'row% = (AktTabIndex% - 3) \ 4
'pzn$ = RezArtikel(row%).pzn
'IdentPzn$ = pzn$
'
'If (row >= AnzRezeptArtikel%) Then
'    Call DefErrPop: Exit Sub
'End If
'
'If (IstImportModus% And (RezArtikel(row%).Imp123 = 0)) Or ((IstImportModus% = 0) And (IstAutIdemModus) And (RezArtikel(row%).AutIdem = 0)) Then
'    Call DefErrPop: Exit Sub
'End If
'
'Sonderregelnr% = 0
'NeueKbvNr% = 0
'If (AutIdemSonderregelOk%) Then
'    If (PrivatRezept% = 0) And (AutIdemKbvNr& = 0) Then
'        frmKkMatchcode.Show 1
'        If (AutIdemKbvNr& > 0) Then
'            NeueKbvNr% = True
'        End If
'    End If
'    If (AutIdemKbvNr& > 0) Then
'        SqlSonderregelNr$ = ""
'        SQLStr$ = "SELECT * FROM AutIdemSonderregelKrankenkasse WHERE KbvNr = " + Str$(AutIdemKbvNr&)
''        Set iRec = AutIdemDB.OpenRecordset(SQLStr$)
'        iRec.open SQLStr, AutIdemDB.ActiveConn
'        Do
'            If (iRec.EOF) Then
'                Exit Do
'            End If
'
'            Sonderregelnr% = Format(CheckNullInt(iRec!AutIdemSonderregelNr), "0")
'            If (SqlSonderregelNr$ = "") Then
'                SqlSonderregelNr$ = "SELECT * FROM AutIdemSonderregelArtikel WHERE ("
'            Else
'                SqlSonderregelNr$ = SqlSonderregelNr$ + " OR"
'            End If
'            SqlSonderregelNr$ = SqlSonderregelNr$ + " AutIdemSonderregelNr = " + Str$(Sonderregelnr%)
'
'            iRec.MoveNext
'        Loop
'        iRec.Close
'    End If
'End If
'
'SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + pzn$
''Set TaxeRec = TaxeDB.OpenRecordset(SQLStr$)
'On Error Resume Next
'TaxeRec.Close
'Err.Clear
'On Error GoTo DefErr
'TaxeRec.open SQLStr, taxeAdoDB.ActiveConn
'If (TaxeRec.EOF) Then
'    Call DefErrPop: Exit Sub
'End If
'
'If (IstImportModus%) Then
'    Titel$ = "Auswahl IMPORTE"
'    If (TaxeRec!OriginalPZN = 0&) Then
'        Call DefErrPop: Exit Sub
'    End If
'
'    pzn$ = PznString(TaxeRec!OriginalPZN)
'    SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + pzn$
'    'Set TaxeRec = TaxeDB.OpenRecordset(SQLStr$)
'    On Error Resume Next
'    TaxeRec.Close
'    Err.Clear
'    On Error GoTo DefErr
'    TaxeRec.open SQLStr, taxeAdoDB.ActiveConn
'    If (TaxeRec.EOF) Then
'        Call DefErrPop: Exit Sub
'    End If
'
'    OrgAvp = TaxeRec!vk / 100#
'    FESTBETR = TaxeRec!FESTBETRAG / 100#
''    If ((TaxeRec!FestKz) And (FESTBETR < OrgAvp)) Then
''        OrgAvp = FESTBETR
''    End If
'    OrgAvp = OrgAvp - FNX(TaxeRec!HerstRabatt / 100#)
'Else
'    If (IstAutIdemModus) Then
'        Titel$ = "Auswahl AUT-IDEM"
'    Else
'        Titel$ = "Auswahl IDENTE"
'        If (TaxeRec!IdentGruppe = 0) Then
'            Call DefErrPop: Exit Sub
'        End If
'    End If
'    If (AutIdemOk%) Then
'        If (TaxeRec!AutIdemNeu) Then
'            AutIdemNeu% = True
'        End If
'    End If
'
'    If (AutIdemNeu%) Then
'        OrgKp# = TaxeRec!FESTBETRAG / 100#
''        OrgKp# = TaxeRec!PreisLinie / 100#
''        Titel$ = Titel$ + " (Preislinie: " + Format(OrgKp#, "0.00") + ")"
'    Else
'        OrgKp# = TaxeRec!FESTBETRAG / 100#
'    End If
'
'    OrgKp# = TaxeRec!vk / 100#
'End If
'
'With frmImportOrIdent.flxEdit
'    .Left = 0
'    .Top = 0
'
'    .Rows = 2
'    .FixedRows = 1
'    .FormatString = "<Pzn|<Name|>Menge|<Eh|^|^Kz|<Hersteller|>Preis||"
'    .Rows = 1
'    .Cols = 11
'
'    .ColWidth(0) = frmImportOrIdent.TextWidth(String(10, "9"))
'    .ColWidth(1) = frmImportOrIdent.TextWidth(String(28, "X"))
'    .ColWidth(2) = frmImportOrIdent.TextWidth(String(6, "X"))
'    .ColWidth(3) = frmImportOrIdent.TextWidth(String(3, "X"))
'    .ColWidth(4) = frmImportOrIdent.TextWidth("Ww")
'    .ColWidth(5) = frmImportOrIdent.TextWidth("Ww")
'    .ColWidth(6) = frmImportOrIdent.TextWidth(String(9, "X"))
'    .ColWidth(7) = frmImportOrIdent.TextWidth("9999999.99")
'    .ColWidth(8) = 0
'    .ColWidth(9) = 0    'frmImportOrIdent.TextWidth("Ww")
'    .ColWidth(10) = wpara.FrmScrollHeight
'
'
'    If (IstImportModus%) Then
'        AVP# = TaxeRec!vk
'        FESTBETR = TaxeRec!FESTBETRAG
''        If ((TaxeRec!FestKz) And (FESTBETR < AVP)) Then
''            AVP = FESTBETR
''        End If
'        If (TaxeRec!HerstRabattKZ <> 0) Then
'            AVP = AVP - TaxeRec!HerstRabatt
'        End If
'        AVP# = AVP / 100#
'
'        s$ = PznString(TaxeRec!pzn) + vbTab + Trim(TaxeRec!Name) + vbTab
'        s$ = s$ + TaxeRec!menge + vbTab + TaxeRec!einheit + vbTab + "i" + vbTab
'        s$ = s$ + vbTab + TaxeRec!HerstellerKB + vbTab + Format(AVP#, "### ##0.00") + vbTab
'        s$ = s$ + "zzz" + Format(AVP#, "0000000.00") + Trim(TaxeRec!Name) + "," + TaxeRec!HerstellerKB + vbTab + TaxeRec!ArtStatus
'        .AddItem s$
'
'        .row = .Rows - 1
'        .col = 4
'        .CellForeColor = vbRed
'
'        SQLStr$ = "SELECT * FROM TAXE WHERE ORIGINALPZN = " + Str$(TaxeRec!OriginalPZN)
'    ElseIf (IstAutIdemModus = 0) Then
'        BtmDerivatMenge# = 0
'        On Error Resume Next
'        BtmDerivatMenge# = TaxeRec!BtmDerivatMenge
'        On Error GoTo 0
'
'        On Error GoTo DefErr
'        If (BtmDerivatMenge >= 0.01) Then
'            SQLStr$ = "SELECT * FROM M2Inhalt WHERE M2Nr = " + Mid(TaxeRec!M2, 2, 9)
''                Set AutIdemRec = AutIdemDB.OpenRecordset(SQLStr$)
'            On Error Resume Next
'            tRec2.Close
'            Err.Clear
'            On Error GoTo DefErr
'            tRec2.open SQLStr, AutIdemDB.ActiveConn
''                            If Not tRec2.NoMatch Then
'            If Not (tRec2.EOF) Then
'                If (InStr(UCase(CheckNullStr(tRec2!Inhalt)), "HYDROMORPHON") > 0) Then
'                    BtmDerivatMenge = 0
'                End If
'            End If
'            tRec2.Close
'        End If
'
'        SQLStr$ = "SELECT * FROM TAXE WHERE IDENTGRUPPE = " + Str$(TaxeRec!IdentGruppe)
''        SQLStr$ = SQLStr$ + " And einheit = """ + TaxeRec!einheit + """"
'        SQLStr$ = SQLStr$ + " And einheit = '" + TaxeRec!einheit + "'"
'        If (BtmDerivatMenge >= 0.01) Then
'            SQLStr$ = SQLStr$ + " AND (BtmDerivatMenge=" + Str(BtmDerivatMenge) + ")"
'        End If
'        SQLStr$ = SQLStr$ + " ORDER BY Menge,VK"
'    ElseIf (AutIdemNeu%) Then
'        BtmDerivatMenge# = 0
'        On Error Resume Next
'        BtmDerivatMenge# = TaxeRec!BtmDerivatMenge
'        On Error GoTo 0
'
'        On Error GoTo DefErr
'        If (BtmDerivatMenge >= 0.01) Then
'            SQLStr$ = "SELECT * FROM M2Inhalt WHERE M2Nr = " + Mid(TaxeRec!M2, 2, 9)
''                Set AutIdemRec = AutIdemDB.OpenRecordset(SQLStr$)
'            On Error Resume Next
'            tRec2.Close
'            Err.Clear
'            On Error GoTo DefErr
'            tRec2.open SQLStr, AutIdemDB.ActiveConn
''                            If Not tRec2.NoMatch Then
'            If Not (tRec2.EOF) Then
'                If (InStr(UCase(CheckNullStr(tRec2!Inhalt)), "HYDROMORPHON") > 0) Then
'                    BtmDerivatMenge = 0
'                End If
'            End If
'            tRec2.Close
'        End If
'
'        SQLStr$ = "SELECT * FROM PZNSUBST WHERE PZN = " + Str$(TaxeRec!pzn)
''        Set AutIdemRec = AutIdemDB.OpenRecordset(SQLStr$)
'        AutIdemRec.open SQLStr, AutIdemDB.ActiveConn
'        If (AutIdemRec.EOF) Then
'            SQLStr$ = ""
'        Else
'            i% = 1
'            SQLStr$ = "SELECT * FROM TAXE WHERE (PZN ="
'            Do
'                If (AutIdemRec.EOF) Then
'                    Exit Do
'                End If
'
'                If (i% > 1) Then SQLStr$ = SQLStr$ + " OR PZN ="
'                SQLStr$ = SQLStr$ + Str$(AutIdemRec!Substpzn)
'                AutIdemRec.MoveNext
'
'                i% = i% + 1
'                If (i% > 250) Then
'                    Exit Do
'                End If
'            Loop
'            If (BtmDerivatMenge >= 0.01) Then
'                SQLStr$ = SQLStr$ + ") AND (BtmDerivatMenge=" + Str(BtmDerivatMenge)
'            End If
''            SQLStr$ = SQLStr$ + ") ORDER BY VK-HerstRabatt,NAME"
'            SQLStr$ = SQLStr$ + ") ORDER BY VK,NAME"
'        End If
'        AutIdemRec.Close
'        pzn$ = ""
'    Else
'        SQLStr = ""
''        If (TaxeRec!IdentGruppe > 0&) Then
''            SQLStr$ = HoleOpAutIdem$
''            pzn$ = ""
''        End If
'    End If
'
'    If (SQLStr = "") Then
'        Call DefErrPop: Exit Sub
'    End If
'
'    'Set TaxeRec = TaxeDB.OpenRecordset(SQLStr$)
'    On Error Resume Next
'    TaxeRec.Close
'    Err.Clear
'    On Error GoTo DefErr
'    TaxeRec.open SQLStr, taxeAdoDB.ActiveConn
'    If (TaxeRec.EOF) Then
'        Call DefErrPop: Exit Sub
'    End If
'
'
'    While Not (TaxeRec.EOF)
'
''        If (TaxeRec!ArtStatus <> "S") Then
'            AVP# = TaxeRec!vk
'            FESTBETR = TaxeRec!FESTBETRAG
''            If (IstImportModus) And ((TaxeRec!FestKz) And (FESTBETR < AVP)) Then
''                AVP = FESTBETR
''            End If
'            If (IstImportModus) And (TaxeRec!HerstRabattKZ <> 0) Then
'                AVP = AVP - TaxeRec!HerstRabatt
'            End If
'            AVP# = AVP / 100#
'
'            If (IstAutIdemModus = 0) Or (TaxeRec!pzn <> Val(pzn$)) Then
'                If (IstImportModus%) Then
'                    If (OrgAvp# - 15) >= AVP# Or (OrgAvp# * 0.85) >= AVP# Then
'                    'If (avp# < (OrgAvp# - 1)) And (avp# < (OrgAvp# * 0.9)) Then
'                        kz$ = ""
'                    Else
'                        kz$ = "-"
'                    End If
'                Else
'                    kz$ = ""
'                    If (IstAutIdemModus) And (OrgKp# > 0) Then
'                        If (AVP# > OrgKp#) Then
'                            kz$ = "+"
'                        Else
'                            kz$ = "-"
'                        End If
'                    End If
'                End If
'
'                s$ = PznString(TaxeRec!pzn) + vbTab + Trim(TaxeRec!Name) + vbTab
'                s$ = s$ + TaxeRec!menge + vbTab + TaxeRec!einheit + vbTab + vbTab
'                s$ = s$ + kz$ + vbTab + TaxeRec!HerstellerKB + vbTab + Format(AVP#, "### ##0.00") + vbTab
'                If (IstAutIdemModus = 0) Then
'                    s$ = s$ + TaxeRec!menge + "," + Format(AVP#, "0000000.00") + Trim(TaxeRec!Name) + "," + TaxeRec!HerstellerKB
'                Else
'                    s$ = s$ + Format(AVP#, "0000000.00") + Trim(TaxeRec!Name) + "," + TaxeRec!HerstellerKB
'                End If
'                s$ = s$ + vbTab + TaxeRec!ArtStatus
'    '            If Not (IstImportModus%) Or kz$ = "" Then
'                    .AddItem s$
'
'                    If (ArtikelDbOk) Then
'                        SQLStr$ = "SELECT * FROM ARTIKEL WHERE PZN = " + PznString(TaxeRec!pzn)
'                        SQLStr = SQLStr + " AND LagerKz<>0"
'                        FabsErrf = Artikel.OpenRecordset(ArtikelAdoRec, SQLStr, 0)
'                    Else
'                        FabsErrf% = ass.IndexSearch(0, Format(TaxeRec!pzn, "0000000"), FabsRecno&)
'                    End If
'                    If (FabsErrf% = 0) Then
'                        .FillStyle = flexFillRepeat
'                        .row = .Rows - 1
'                        .col = 0
'                        .RowSel = .row
'                        .ColSel = .Cols - 1
'                        .CellFontBold = True
'                        .FillStyle = flexFillSingle
'                    End If
'
'                    RabKz = "zzz"
'                    If (SqlSonderregelNr$ <> "") Then
'                        SQLStr$ = SqlSonderregelNr$ + ")"
'                        SQLStr$ = SQLStr$ + " AND NrKz='P' AND ZoNr=" + Str$(TaxeRec!pzn)
''                        Set iRec = AutIdemDB.OpenRecordset(SQLStr$)
''                        If (iRec.RecordCount > 0) Then
'                        iRec.open SQLStr, AutIdemDB.ActiveConn
'                        If Not (iRec.EOF) Then
'                            RabKz = "   "
'                            .FillStyle = flexFillRepeat
'                            .row = .Rows - 1
'                            .col = 1
'                            .RowSel = .row
'                            .ColSel = .col
'                            .CellBackColor = vbGreen
'                            .FillStyle = flexFillSingle
'                        End If
'                        iRec.Close
'                    End If
'                    .TextMatrix(.Rows - 1, 8) = RabKz + .TextMatrix(.Rows - 1, 8)
'
'                    If (Val(TaxeRec!AlternativPackung) > 0) Then
'                        .FillStyle = flexFillRepeat
'                        .row = .Rows - 1
'                        .col = 2
'                        .RowSel = .row
'                        .ColSel = .col
'                        .CellForeColor = vbRed
'                        .FillStyle = flexFillSingle
'                    End If
'
'                    .FillStyle = flexFillRepeat
'                    .row = .Rows - 1
'                    .col = 4
'                    .RowSel = .row
'                    If (TaxeRec!OriginalPZN = TaxeRec!pzn) Then
'                        .text = "i"
'                        .CellForeColor = vbRed
'                    ElseIf (TaxeRec!OriginalPZN > 0) Then
'                        .text = "I"
'
''                        Set sTaxeRec = TaxeDB.OpenRecordset("SELECT * FROM TAXE WHERE PZN = " + CStr(TaxeRec!OriginalPZN))
'                        On Error Resume Next
'                        sTaxeRec.Close
'                        Err.Clear
'                        On Error GoTo DefErr
'                        sTaxeRec.open "SELECT * FROM TAXE WHERE PZN = " + CStr(TaxeRec!OriginalPZN), taxeAdoDB.ActiveConn
'                        If Not (sTaxeRec.EOF) Then
'                            avp2# = sTaxeRec!vk
'                            FESTBETR = sTaxeRec!FESTBETRAG
''                            If (IstImportModus) And ((sTaxeRec!FestKz) And (FESTBETR < avp2)) Then
''                                avp2 = FESTBETR
''                            End If
'                            If (IstImportModus) And (sTaxeRec!HerstRabattKZ <> 0) Then
'                                avp2 = avp2 - sTaxeRec!HerstRabatt
'                            End If
'                            avp2# = avp2 / 100#
'
'                            If (avp2) - 15 >= (AVP) Or (avp2) * 0.85 >= AVP Then
'                            Else
'                                .CellForeColor = vbRed
'                            End If
'                        End If
'                    End If
'                    .FillStyle = flexFillSingle
'    '            End If
'            End If
''        End If
'
'        TaxeRec.MoveNext
'    Wend
'
'    .row = 1
'    .col = 8
'    .RowSel = .Rows - 1
'    .ColSel = .col
''    If (IstImportModus%) Then
''        .Sort = 8
''    Else
''        .Sort = 1
''    End If
'    .Sort = flexSortGenericAscending
'    .col = 1
'    .ColSel = .Cols - 1
'
'    .Height = .RowHeight(0) * .Rows + 90
'    If (para.Newline) Then
'        MaxHe = wpara.WorkAreaHeight - 2400
'    Else
'        MaxHe = wpara.WorkAreaHeight - 150
'    End If
'    If (.Height > MaxHe) Then
''        .Height = wpara.WorkAreaHeight - 150
'        .Height = MaxHe
'        .Height = ((.Height - 90) \ .RowHeight(0)) * .RowHeight(0) + 90
'    End If
'
'    spBreite& = 0
'    For i% = 0 To (.Cols - 1)
'        spBreite& = spBreite& + .ColWidth(i%)
'    Next i%
'    .Width = spBreite& + 90
'
'    With frmImportOrIdent
'        .Caption = Titel$
'        .Height = .flxEdit.Height + wpara.FrmCaptionHeight + 60
'        .Width = .flxEdit.Width + 60
'        .Left = ParentForm.Left + (ParentForm.Width - .Width) / 2
'        .Top = ParentForm.Top + (ParentForm.Height - .Height) / 2
'    End With
'
''    If (IstImportModus% = 0) And (IstAutIdemModus) Then
'    If (IstImportModus% = 0) Then
'        If (AutIdemNeu%) Then
'            grenze# = OrgKp#
'        ElseIf (.Rows > 6) Then
'            For i% = 1 To 3
'                unten# = unten# + xVal(.TextMatrix(i%, 7))
'            Next i%
'            unten# = unten# / 3
'            For i% = (.Rows - 3) To (.Rows - 1)
'                oben# = oben# + xVal(.TextMatrix(i%, 7))
'            Next i%
'            oben# = oben# / 3
'            grenze# = xVal(.TextMatrix(1, 7)) + (oben# - unten#) / 3 'unteres Drittel
'        Else
'            If (.Rows <= 4) Then
'                i% = 1
'            Else
'                i% = .Rows - 3
'            End If
'            grenze# = xVal(.TextMatrix(i%, 7))
'        End If
'
'        For i% = 1 To (.Rows - 1)
'            AVP# = xVal(.TextMatrix(i%, 7))
'
'            .row = i%
'            .col = 7
'            .RowSel = .row
'            .ColSel = .col
'            '.CellFontBold = (avp# > grenze#)
'            If AnzGilt% < 3 Or AVP# = grenze# Then
'              '.CellFontBold = (avp# > grenze#)
'              .CellFontBold = False
'              grenze# = AVP#
'            Else
'              .CellFontBold = True
'                If (IstAutIdemModus) Then
'                    .CellForeColor = vbRed
'                End If
'            End If
'            If AnzGilt% <= 3 Then grenze# = AVP#
'            If (AVP# <= grenze#) Then
'                AnzGilt% = AnzGilt% + 1
'            End If
'        Next i%
'
'        If (AutIdemNeu% = 0) And (AnzGilt% < 5) And (.Rows > 6) Then 'wenn weniger als 5 im 1. Drittel -> 5 zulassen
'            For i% = (AnzGilt + 1) To 5
'                .row = i%
'                .col = 7
'                .RowSel = .row
'                .ColSel = .col
'                .CellFontBold = False
'            Next i%
'            grenze# = xVal(.TextMatrix(5, 7))
'            Do
'                If (i% >= .Rows) Then Exit Do
'                If (xVal(.TextMatrix(i%, 7)) = grenze#) Then
'                    .row = i%
'                    .col = 7
'                    .RowSel = .row
'                    .ColSel = .col
'                    .CellFontBold = False
'                    i% = i% + 1
'                Else
'                    Exit Do
'                End If
'            Loop
'        End If
'
'        .row = 1
'    Else
''      .row = 1
''      .col = 7
''      .RowSel = .row
''      .ColSel = .col
''      .CellFontBold = True
''      .CellForeColor = vbRed
''
''      If .Rows > 2 Then
''        .row = 2
''      End If
'
'        For i = 1 To .Rows - 1
'            .row = i
'            .col = 4
'            If (.text <> "i") And (.CellForeColor = vbRed) Then
'                .col = 7
'                .CellForeColor = vbRed
'            End If
'        Next i
'        .row = 1
'
'    End If
'
'    .col = 0
'    .ColSel = .Cols - 1
'    .Visible = True
'End With
'
'With frmImportOrIdent
'    .cmdF2.Enabled = (IstImportModus% = 0)
'
'    If (para.Newline) Then
'        iAdd = wpara.NlFlexBackY
'        iAdd2 = wpara.NlCaptionY
'
'        With frmImportOrIdent.flxEdit
''            .ScrollBars = flexScrollBarNone
'            .BorderStyle = 0
'            .Width = .Width - 90
'            .Height = .Height - 75
'            .GridLines = flexGridFlat
'            .GridLinesFixed = flexGridFlat
'            .GridColorFixed = .GridColor
'            .BackColor = wpara.nlFlexBackColor    'vbWhite
'            .BackColorBkg = wpara.nlFlexBackColor    'vbWhite
'            .BackColorFixed = wpara.nlFlexBackColorFixed   ' RGB(199, 176, 123)
'            .BackColorSel = wpara.nlFlexBackColorSel  ' RGB(232, 217, 172)
'            .ForeColorSel = vbBlack
'
'            .Left = .Left + 2 * iAdd
'            .Top = .Top + 2 * iAdd
'
'            For i% = 1 To (.Rows - 1)
'                .row = i
'
'                .col = 1
'                If (.CellBackColor = vbGreen) Then
'                    .CellBackColor = wpara.nlGreen
'                End If
'
'                .col = 2
'                If (.CellForeColor = vbRed) Then
'                    .CellForeColor = vbWhite
''                    If (Knallrot) Then
''                        .CellBackColor = vbRed
''                    Else
'                        .CellBackColor = wpara.nlRed
''                    End If
'                End If
'
'                .col = 4
'                If (.CellForeColor = vbRed) Then
'                    .CellForeColor = vbWhite
''                    If (Knallrot) Then
''                        .CellBackColor = vbRed
''                    Else
'                        .CellBackColor = wpara.nlRed
''                    End If
'                End If
'
'                .col = 7
'                If (.CellForeColor = vbRed) Then
'                    .CellForeColor = vbWhite
''                    If (Knallrot) Then
''                        .CellBackColor = vbRed
''                    Else
'                        .CellBackColor = wpara.nlRed
''                    End If
'                End If
'            Next i
'
'            .row = 1
'            .col = 0
'            .ColSel = .Cols - 1
'        End With
'
'        With frmImportOrIdent
'            .Width = .Width + 4 * iAdd
'            .Height = .Height + 4 * iAdd
'
'            .flxEdit.Top = .flxEdit.Top + iAdd2
'
'            With .nlcmdOk
'                .Init
'                .Left = (frmImportOrIdent.ScaleWidth - (.Width * 2 + 300)) / 2
'                .Top = frmImportOrIdent.flxEdit.Top + frmImportOrIdent.flxEdit.Height + iAdd + 600
'                .Caption = frmImportOrIdent.cmdOk.Caption
'                .TabIndex = frmImportOrIdent.cmdOk.TabIndex
'                .Enabled = frmImportOrIdent.cmdOk.Enabled
'                .Default = frmImportOrIdent.cmdOk.Default
'                .Cancel = frmImportOrIdent.cmdOk.Cancel
'                .Visible = True
'            End With
'            .cmdOk.Visible = False
'
'            With .nlcmdEsc
'                .Init
'                .Left = frmImportOrIdent.nlcmdOk.Left + .Width + 300
'                .Top = frmImportOrIdent.nlcmdOk.Top
'                .Caption = frmImportOrIdent.cmdEsc.Caption
'                .TabIndex = frmImportOrIdent.cmdEsc.TabIndex
'                .Enabled = frmImportOrIdent.cmdEsc.Enabled
'                .Default = frmImportOrIdent.cmdEsc.Default
'                .Cancel = frmImportOrIdent.cmdEsc.Cancel
'                .Visible = True
'            End With
'            .cmdEsc.Visible = False
'
'            .Height = .nlcmdOk.Top + .nlcmdOk.Height + wpara.FrmCaptionHeight + 450
'
'            Call wpara.NewLineWindow(frmImportOrIdent, .nlcmdOk.Top)
'            RoundRect .hdc, (.flxEdit.Left - iAdd) / Screen.TwipsPerPixelX, (.flxEdit.Top - iAdd) / Screen.TwipsPerPixelY, (.flxEdit.Left + .flxEdit.Width + iAdd) / Screen.TwipsPerPixelX, (.flxEdit.Top + .flxEdit.Height + iAdd) / Screen.TwipsPerPixelY, 20, 20
'        End With
'    Else
'        .nlcmdOk.Visible = False
'        .nlcmdEsc.Visible = False
'    End If
'
'    .Left = frmAction.Left + (frmAction.Width - .Width) / 2
'    .Top = frmAction.Top + (frmAction.Height - .Height) / 2
'
'    .Show 1
'End With
'
'If (EditErg%) Then
'    ind% = InStr(EditTxt$, vbTab)
'    pzn$ = Left$(EditTxt$, ind% - 1)
'
'    Call NeuerArtikel(row%, pzn$)
'    Call PaintArtikel(row%)
'End If
'
'
'If (NeueKbvNr%) Then
'    iAnzZ% = AnzRezeptArtikel% - 1      'GS 16.05.02
'    If (iAnzZ% > 6) Then
'        iAnzZ% = 6
'    End If
'    iTabIndex% = AktTabIndex%
'    For i% = 0 To iAnzZ%
'        RezArtikel(i%).AutIdem = 0
'        RezArtikel(i%).N0 = 0
'        SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + RezArtikel(i%).pzn
'        'Set TaxeRec = TaxeDB.OpenRecordset(SQLStr$)
'        On Error Resume Next
'        TaxeRec.Close
'        Err.Clear
'        On Error GoTo DefErr
'        TaxeRec.open SQLStr, taxeAdoDB.ActiveConn
'        Call CheckAutIdem(i%)
'        Call PaintArtikel(i%)
'    Next i%
'    Call TabIndexChange(iTabIndex%, False)
'End If
'
'Call DefErrPop
'End Sub

Sub FlxAvAuswahlBefuellen()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("FlxAvAuswahlBefuellen")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
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
Dim IstVebNr&, SollVebNr&
Dim h$, h2$
Dim VdbVvlRec As New ADODB.Recordset

With frmAvAuswahl.flxAvAuswahl(0)
    .Rows = 0
    .Cols = 6
    For i% = 0 To UBound(AvLaender$)
        h$ = Mid$(AvLaender$(i%), 3) + vbTab + Left$(AvLaender$(i%), 2)
        .AddItem h$
    Next i%
    .row = 0
    For i% = 0 To (.Rows - 1)
        If (Val(.TextMatrix(i%, 1)) = ActBundesland%) Then
            .row = i%
            Exit For
        End If
    Next i%
    ActBundeslandInd% = .row
End With

With frmAvAuswahl.flxAvAuswahl(1)
'    .Rows = 0
'    .Cols = 5
''    For i% = 0 To UBound(AvKassen$)
''        h$ = Mid$(AvKassen$(i%), 3) + vbTab + Left$(AvKassen$(i%), 2)
''        .AddItem h$
''    Next i%
''    .row = 0
''    For i% = 0 To (.Rows - 1)
''        If (Val(.TextMatrix(i%, 1)) = ActKasse%) Then
''            .row = i%
''            Exit For
''        End If
''    Next i%
''    ActKasseInd% = .row
'
'    SQLStr = "SELECT * FROM VdbVVL"
''    SQLStr = SQLStr + " LEFT JOIN VdbVereinbarungen ON VdbVVL.VebNr=VdbVereinbarungen.VebNr"
'    SQLStr = SQLStr + " WHERE VdbVVL.LandNr=" + CStr(ActBundesland) + " OR VdbVVL.LandNr=18"
'    FabsErrf = AplusVDB.OpenRecordset(VdbVvlRec, SQLStr)
'    Do
'        If (VdbVvlRec.EOF) Then
'            Exit Do
'        End If
'
'        SollVebNr = CheckNullLong(VdbVvlRec!VebNr)
'        h$ = "(leer)" + vbTab + "0"
'        For i% = 0 To UBound(AvVereinbarungen$)
'            h2 = AvVereinbarungen(i)
'            IstVebNr = Val(Left(h2, 10))
'            If (IstVebNr = SollVebNr) Then
'                h$ = Trim(Mid$(h2, 11)) + vbTab + CStr(IstVebNr)
'                Exit For
'            End If
'        Next i%
'        .AddItem h$
'
'        VdbVvlRec.MoveNext
'    Loop
'    VdbVvlRec.Close
'
'    .row = 1
'    ActKasseInd% = .row
End With

With frmAvAuswahl.flxAvAuswahl(2)
    .Rows = 0
    .Cols = 6
    For i% = 0 To UBound(AvVerordnungen$)
        h$ = Mid$(AvVerordnungen$(i%), 3) + vbTab + Left$(AvVerordnungen$(i%), 2)
        .AddItem h$
    Next i%
    .row = 0
    For i% = 0 To (.Rows - 1)
        If (Val(.TextMatrix(i%, 1)) = ActVerordnung%) Then
            .row = i%
            Exit For
        End If
    Next i%
    ActVerordnungInd% = .row
End With
    
Call DefErrPop
End Sub

Sub FlxOptionenBefuellen(iFlex As MSFlexGrid)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("FlxOptionenBefuellen")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
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

With iFlex
    .Rows = 0
    .Cols = 3
    For i% = 0 To UBound(AvLaender$)
        h$ = Mid$(AvLaender$(i%), 3) + vbTab + Left$(AvLaender$(i%), 2)
        .AddItem h$
    Next i%
    .row = 0
    For i% = 0 To (.Rows - 1)
        If (Val(.TextMatrix(i%, 1)) = ActBundesland%) Then
            .row = i%
            Exit For
        End If
    Next i%
    ActBundeslandInd% = .row
End With

Call DefErrPop
End Sub

Sub LadeSonderPzn(SrcDatei$, SonderFeld$())
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("LadeSonderPzn")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
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
Dim SONDER_HANDLE%, ubou%
Dim h$
  
ReDim SonderFeld$(0)
ubou% = -1

h$ = "\user\" + SrcDatei$
If (Dir$(h$) <> "") Then
    SONDER_HANDLE% = FileOpen(h$, "I")
    
    If (SONDER_HANDLE% > 0) Then
        
        Do
            If (EOF(SONDER_HANDLE%)) Then Exit Do
            Line Input #SONDER_HANDLE%, h$
            
            Call OemToChar(h$, h$)
            
            If (Trim(h) <> "") Then
                ubou% = ubou% + 1
                ReDim Preserve SonderFeld$(ubou%)
                SonderFeld$(ubou%) = Trim(h$)
            End If
        Loop
        
        Close #SONDER_HANDLE%
    End If
End If

Call DefErrPop
End Sub

Sub FlxPznAuswahlBefuellen(ind%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("FlxPznAuswahlBefuellen")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
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

With frmPznAuswahl.flxPznAuswahl(1)
    .Rows = 1
    If (ind% = 0) Then
        .AddItem "Direkteingabe HM-Nummer" + vbTab + "HMEingabe"
    ElseIf (ind% = 1) Then
        For i% = 0 To UBound(HmNummern$)
            h$ = Mid$(HmNummern$(i%), 11) + vbTab + Left$(HmNummern$(i%), 10)
            .AddItem h$
        Next i%
    Else
        For i% = 0 To UBound(SonderPzn$)
            h$ = Mid$(SonderPzn$(i%), 9) + vbTab + Left$(SonderPzn$(i%), 8)
            .AddItem h$
        Next i%
    End If
    .row = 0
End With

Call DefErrPop
End Sub

Public Sub LadeAbgabePreise()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("LadeAbgabePreise")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
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
Dim i%, ABG_HANDLE%, ubou%
Dim dPreis#
Dim h$, h2$
  
ReDim AbgabeKosten(0)
ubou% = -1

ubou% = ubou% + 1
ReDim Preserve AbgabeKosten(ubou%)
With AbgabeKosten(ubou%)
    .Name = "Zurücksetzen"
    .kosten = 0
    .ind = 0
End With

If (Dir("abgpr.dat") = "") Then Call DefErrPop: Exit Sub

ABG_HANDLE% = FileOpen("abgpr.dat", "R", "B")

If (ABG_HANDLE% > 0) Then
    h$ = String(16, 0)
    Get #ABG_HANDLE%, , h$
    For i% = 1 To 100
        Get #ABG_HANDLE%, (i% * 16) + 1, h$
        If (EOF(ABG_HANDLE%)) Then Exit For
        
        Call OemToChar(h$, h$)
        
        h2$ = Trim(h$)
        If (h2$ <> "") Then
            ubou% = ubou% + 1
            ReDim Preserve AbgabeKosten(ubou%)
            With AbgabeKosten(ubou%)
                .Name = Left$(h2$, 12)
                .kosten = Val(Mid$(h2$, 13)) / 100#
                .ind = i%
            End With
        End If
    Next i%
    
    Close #ABG_HANDLE%
End If

Call DefErrPop
End Sub

Sub LadeImpfungen()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("LadeImpfungen")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
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
Dim IMPF_HANDLE%, ubou%
Dim h$, h2$, pzn$
  
ImpfstoffeDa% = False

ReDim Impfungen(0)
ubou% = -1

If (Dir("impfung.dat") = "") Then Call DefErrPop: Exit Sub

IMPF_HANDLE% = FileOpen("impfung.dat", "R", "B")

If (IMPF_HANDLE% > 0) Then

    Do
        h$ = String(38, 0)
        Get #IMPF_HANDLE%, , h$
        If (EOF(IMPF_HANDLE%)) Then Exit Do
        
        h$ = Left$(h$, 36)
        Call OemToChar(h$, h$)
        
        pzn$ = Mid$(h$, 29, 8)
        h$ = Trim(Left$(h$, 28))
        
        If (h$ <> "") Then
            ubou% = ubou% + 1
            ReDim Preserve Impfungen(ubou%)
            Impfungen(ubou%).pzn = pzn$
            Impfungen(ubou%).text = h$
        End If
    Loop
    
    Close #IMPF_HANDLE%
End If

ImpfstoffeDa% = True

Call DefErrPop
End Sub

Sub PaintAvParameter()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("PaintAvParameter")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
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
Dim Name$, Key$, h$

ActBundeslandInd% = 0
If (ActBundesland% <= 0) Then
    Name$ = ""
Else
    Key$ = Format(ActBundesland%, "00")
    For i% = 0 To UBound(AvLaender$)
        h$ = AvLaender$(i%)
        If (Left$(h$, 2) = Key$) Then
            Name$ = Mid$(h$, 3)
            ActBundeslandInd% = i%
            Exit For
        End If
    Next i%
End If
ParentForm.lblAv(0).Caption = Name$


ActVerordnungInd% = 0
If (ActVerordnung% <= 0) Then
    Name$ = ""
Else
    Key$ = Format(ActVerordnung%, "00")
    For i% = 0 To UBound(AvVerordnungen$)
        h$ = AvVerordnungen$(i%)
        If (Left$(h$, 2) = Key$) Then
            Name$ = Mid$(h$, 3)
            ActVerordnungInd% = i%
            Exit For
        End If
    Next i%
End If
ParentForm.lblAv(1).Caption = Name$


ActKasseInd% = 0
If (ActKasse% <= 0) Then
    Name$ = ""
Else
    Key$ = Format(ActKasse%, "00")
    For i% = 0 To UBound(AvKassen$)
        h$ = AvKassen$(i%)
        If (Left$(h$, 2) = Key$) Then
            Name$ = Mid$(h$, 3)
            ActKasseInd% = i%
            Exit For
        End If
    Next i%
End If
ParentForm.lblAv(2).Caption = Name$

With ParentForm
    For i% = 0 To 2
        If (i% >= .flxInfo(0).Rows) Then Exit For
        .flxInfo(0).TextMatrix(i%, 0) = .lblAv(i%).Caption
    Next i%
    If (RezeptMitHilfsmittelNr% = 1) Then
        If (ParentForm.Name = frmAction.Name) Then
            .flxInfo(0).TextMatrix(3, 0) = "§302"
        End If
    ElseIf (ActKkasse$ <> "") And (.flxInfo(0).Rows >= 4) Then
        KasseRec.index = "Unique"
        KasseRec.Seek "=", ActKkasse$
        If (KasseRec.NoMatch) Or (KasseRec.EOF) Then
            h$ = "(" + ActKkasse$ + ")"
        Else
            h$ = Trim(KasseRec!Name)
            IstRoteKasse% = KasseRec!importkz
        End If
        If (ParentForm.Name = frmAction.Name) Then
            .flxInfo(0).TextMatrix(3, 0) = h$
        End If
    End If
End With

Call DefErrPop
End Sub

Sub NeuerArtikel(row%, pzn$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("NeuerArtikel")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
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
Dim i%, hWg%, iTabIndex%, IstDrüber%, IstSonderPzn%, MsgErg%, iMkVerzicht%, iZuzErlass%
Dim dPreis#, dZuz#, dZuzDos#, dZuzErlass#
Dim aktavp As Double, OrgAvp As Double
Dim SQLStr$, txt$, hStr$, h$

KeinPreisDa% = False

IstSonderPzn% = CheckSonderPzn%(pzn$)

If (IstSonderPzn% > 0) Then
    RezArtikel(row%).pzn = pzn$
    RezArtikel(row%).text = Mid$(SonderPzn$(IstSonderPzn%), 9)
    RezArtikel(row%).Faktor = 1
    RezArtikel(row%).Preis = 0#
    RezArtikel(row%).flag = 0
    RezArtikel(row%).Zuz = ZuzFeld!(3)
    RezArtikel(row%).HerstRabattPrivat130Brutto = 0
    RezArtikel(row).VdbPauschale = 0
    RezArtikel(row%).Imp123 = 0
    RezArtikel(row%).ImpErspart = 0
    RezArtikel(row%).Fam = 0
    RezArtikel(row%).AutIdem = 0
    RezArtikel(row%).AutIdemKreuz = 0
    RezArtikel(row%).N0 = 0
    RezArtikel(row%).MagIndex = 0
    RezArtikel(row%).TkkPznDruck = 0
    RezArtikel(row%).VebNr = 0
    RezArtikel(row%).PauschaleNr = 0
    RezArtikel(row%).Fortsetzung = 0
    RezArtikel(row%).HmAbrechnungsKz = Space(2)
    RezArtikel(row%).LEGS = Space(7)

    If (pzn = "02567018") Then
        SQLStr$ = "SELECT * FROM ARTIKEL WHERE PZN = " + pzn
        SQLStr = SQLStr + " AND LagerKz<>0"
        FabsErrf = Artikel.OpenRecordset(ArtikelAdoRec, SQLStr)
        If (FabsErrf = 0) Then
            RezArtikel(row%).Preis = ast.AVP
        End If
    End If

    If (row% >= AnzRezeptArtikel%) Then
        AnzRezeptArtikel% = row% + 1
    End If

    Call DefErrPop: Exit Sub
End If



'''''''''''''

SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + pzn$
'Set TaxeRec = TaxeDB.OpenRecordset(SQLStr$)
On Error Resume Next
TaxeRec.Close
Err.Clear
On Error GoTo DefErr
TaxeRec.Open SQLStr, taxeAdoDB.ActiveConn
If (TaxeRec.EOF) Then
    Call DefErrPop: Exit Sub
End If
  
RezArtikel(row%).pzn = PznString(TaxeRec!pzn)

txt$ = Left$(TaxeRec!Name + Space$(29), 29) + Mid$(TaxeRec!menge, 3) + TaxeRec!einheit
RezArtikel(row%).text = txt$

RezArtikel(row%).PlusMehrKosten = 0
RezArtikel(row%).ZuzahlungsErlass = 0
RezArtikel(row%).Warenzeichen = 0
RezArtikel(row%).MietDauer = 0

RezArtikel(row%).zusatz(0) = Space$(Len(RezArtikel(row%).zusatz(0)))
RezArtikel(row%).zusatz(1) = Space$(Len(RezArtikel(row%).zusatz(1)))

RezArtikel(row%).AutIdemKreuz = 0

RezArtikel(row%).HerstRabattPrivat130Brutto = FNX(CheckNullDouble(TaxeRec!HerstRabattPrivat130) * (1# + (para.Mwst(2)) / 100#) / 100#)

RezArtikel(row%).Faktor = 1

RezArtikel(row%).Fortsetzung = 0
RezArtikel(row%).HmAbrechnungsKz = Space(2)
RezArtikel(row%).LEGS = Space(7)

If (PrivatRezept%) Then
    dPreis = PrivatRezeptPreis
ElseIf (PrivatRezept% = 0) And (TaxeRec!FestKz) And (TaxeRec!FESTBETRAG < TaxeRec!vk) And (ActKasse% <> 7) Then
    dPreis# = TaxeRec!FESTBETRAG
Else
    dPreis# = TaxeRec!vk
End If

dZuzErlass# = 0
If (PrivatRezept% = 0) And (kKassenOk%) Then
    iMkVerzicht% = 0
    iZuzErlass% = 0
    On Error Resume Next
    iMkVerzicht% = CheckNullInt(TaxeRec!MehrkostenverzichtKZ)
    iZuzErlass% = CheckNullInt(TaxeRec!ZuzahlungsErlassKZ)
    On Error GoTo 0
    If (iMkVerzicht%) Or (iZuzErlass%) Then
        If (ActAVWGik& < 0) Then
            Do
                h$ = MyInputBox("IK der Krankenkasse: ", "Prüfung auf Mehrkosten-Verzicht/Zuzahlungs-Erlass")
                h$ = UCase(Trim(h$))
                If (h$ = "") Then
                    Exit Do
                ElseIf (Val(h$) > 0) Then
                    SQLStr$ = "SELECT * FROM kKassen"
                    SQLStr = SQLStr + " WHERE IK = " + Str$(Val(h$))
'                    Set KkassenRec = kKassenDB.OpenRecordset(SQLStr$)
'                    If (KkassenRec.RecordCount > 0) Then
                    FabsErrf = kKassenDB.OpenRecordset(KkassenRec, SQLStr)
                    If Not (KkassenRec.EOF) Then
                        ActAVWGik& = Val(h$)
                        Exit Do
                    End If
                End If
            Loop
        End If
        If (ActAVWGik& > 0) Then
            If (iMkVerzicht%) Then
                SQLStr$ = "SELECT * FROM MehrkostenVerzicht"
                SQLStr = SQLStr + " WHERE PZN = " + Str$(TaxeRec!pzn)
                SQLStr = SQLStr + " AND IK = " + Str$(ActAVWGik&)
'                Set KkassenRec = kKassenDB.OpenRecordset(SQLStr$)
'                If (KkassenRec.RecordCount > 0) Then
                FabsErrf = kKassenDB.OpenRecordset(KkassenRec, SQLStr)
                If Not (KkassenRec.EOF) Then
                    RezArtikel(row%).PlusMehrKosten = 1
                    dPreis# = TaxeRec!vk
                End If
            End If
            If (iZuzErlass%) Then
                SQLStr$ = "SELECT * FROM ZuzahlungsErlass"
                SQLStr = SQLStr + " WHERE PZN = " + Str$(TaxeRec!pzn)
                SQLStr = SQLStr + " AND IK = " + Str$(ActAVWGik&)
'                Set KkassenRec = kKassenDB.OpenRecordset(SQLStr$)
'                If (KkassenRec.RecordCount > 0) Then
                FabsErrf = kKassenDB.OpenRecordset(KkassenRec, SQLStr)
                If Not (KkassenRec.EOF) Then
                    RezArtikel(row%).ZuzahlungsErlass = 1
                    dZuzErlass# = KkassenRec!Prozent
                End If
            End If
        End If
    End If
    
'    If (RezArtikel(row%).PlusMehrKosten = 0) Then
'        Dim dMehrkosten#
'        If TaxeRec!FESTBETRAG > 0 And TaxeRec!FESTBETRAG < TaxeRec!vk Then
'            dMehrkosten = (TaxeRec!vk - TaxeRec!FESTBETRAG) / 100#
'            RezArtikel(row).zusatz(1) = "MK " + CStr(dMehrkosten)
'        End If
'    End If
End If

RezArtikel(row%).Preis = dPreis# / 100#

If (RezArtikel(row%).Preis = 0#) Then KeinPreisDa% = True

RezArtikel(row%).flag = RezArtikel(row%).flag And ((REZ_WG_4 Or REZ_BLINK) Xor &HFF)

RezArtikel(row%).IstWg4 = 0
'hWg% = TaxeRec!Warengruppe
''If (hWg% = 4) Or (hWg% = 14) Then
'If (hWg% = 4) Then
'    RezArtikel(row%).flag = RezArtikel(row%).flag Or (REZ_WG_4 Or REZ_BLINK)
'    If (Left$(TaxeRec!Zuzahlung, 1) <> "9") Then
'        RezArtikel(row%).IstWg4 = 1
'    End If
''ElseIf (hWg% = 3) Then
''    RezArtikel(row%).IstWg4 = 1
'End If

RezArtikel(row%).TkkPznDruck = 0


RezArtikel(row%).IstWg4 = CheckZuzZeilenwert
If (PrivatRezept%) Then
ElseIf (TaxeRec!VerbandKz) Then
    RezArtikel(row%).flag = RezArtikel(row%).flag Or (REZ_WG_4 Or REZ_BLINK)
ElseIf (VmFlag%) Then
    If (AplusVOk) Then
        SQLStr$ = "SELECT * FROM VdbArtikel WHERE PZN = " + pzn
'        Set VdbPznRec = AplusVDB.OpenRecordset(SQLStr$)
'        FabsErrf = Abs(VdbPznRec.EOF)
        FabsErrf = AplusVDB.OpenRecordset(VdbPznRec, SQLStr)
    Else
        FabsErrf% = VmPzn.IndexSearch(0, pzn, FabsRecno&)
    End If
    If (FabsErrf% = 0) Then
        RezArtikel(row%).flag = RezArtikel(row%).flag Or (REZ_WG_4 Or REZ_BLINK)
    End If
End If

dZuz# = 0#
If (PrivatRezept%) Then
    dZuzDos# = 0
    RezArtikel(row%).Zuz = dZuz#
    RezArtikel(row%).Imp123 = 0
    RezArtikel(row%).AutIdem = 0
    RezArtikel(row%).N0 = 0
ElseIf (IvfRezept%) Then
    RezArtikel(row%).Zuz = RezArtikel(row%).Preis / 2
    RezArtikel(row%).Imp123 = 0
    RezArtikel(row%).AutIdem = 0
    RezArtikel(row%).N0 = 0
Else
    If (RezArtikel(row%).ZuzahlungsErlass) Then
        dZuzDos# = CalcTaxeZuzahlungDOS(-1, dZuzErlass#)
    Else
        dZuzDos# = CalcTaxeZuzahlungDOS
    End If
    If (TaxeRec!Negliste) Then dZuzDos# = 1#
    If (Val(Left$(TaxeRec!Zuzahlung, 1)) = 9) Then dZuzDos# = 0#

    If (dZuzDos# = 1#) Then
        Call iMsgBox("Achtung: Artikel mit Zuzahlung 'Nicht betroffen'", vbInformation)
    ElseIf (dZuzDos# >= 1000#) Then
        dZuz# = dZuzDos# - 1000#
    ElseIf Not (GebFrei%) And (dZuzDos# > 0#) Then
        RezArtikel(row%).Preis = 0#
        If (dZuzDos# = 3#) Then
            Call iMsgBox("Artikel unter Zuzahlung, Geldrückgabe ?")
            KeinPreisDa% = False
        Else
            RezArtikel(row%).flag = RezArtikel(row%).flag And (REZ_BLINK Xor &HFF)
        End If
    End If

'    If (RezArtikel(row%).ZuzahlungsErlass) Then
'        If (dZuzErlass# = 100#) Then
'            dZuz# = 0
'        ElseIf (dZuzErlass# > 0#) Then
'            dZuz# = dZuz# * (100 - dZuzErlass#)
'        End If
'    End If
    
    RezArtikel(row%).Zuz = dZuz#
    
    If (ImpFlag% = 0) Then
        RezArtikel(row%).Imp123 = 0
    ElseIf (TaxeRec!OriginalPZN = TaxeRec!pzn) Then
        RezArtikel(row%).Imp123 = 6
        OrgAvp = TaxeRec!vk / 100#
        SQLStr$ = "SELECT * FROM TAXE WHERE ORIGINALPZN = " + PznString(TaxeRec!OriginalPZN)
        'Set TaxeRec = TaxeDB.OpenRecordset(SQLStr$)
        On Error Resume Next
        TaxeRec.Close
        Err.Clear
        On Error GoTo DefErr
        TaxeRec.Open SQLStr, taxeAdoDB.ActiveConn
        TaxeRec.MoveFirst
        Do While Not TaxeRec.EOF
          If (TaxeRec!OriginalPZN <> TaxeRec!pzn) Then
            aktavp = TaxeRec!vk / 100#
            If (aktavp <= OrgAvp - 15) Or (aktavp <= OrgAvp * 0.85) Then
              RezArtikel(row%).Imp123 = 5
              Exit Do
            End If
          End If
          TaxeRec.MoveNext
        Loop
        SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + pzn$
        'Set TaxeRec = TaxeDB.OpenRecordset(SQLStr$)
        On Error Resume Next
        TaxeRec.Close
        Err.Clear
        On Error GoTo DefErr
        TaxeRec.Open SQLStr, taxeAdoDB.ActiveConn
    
    ElseIf (TaxeRec!OriginalPZN > 0) Then
        'RezArtikel(row%).Imp123 = 2
        
        aktavp = TaxeRec!vk / 100
        SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + PznString(TaxeRec!OriginalPZN)
        'Set TaxeRec = TaxeDB.OpenRecordset(SQLStr$)
        On Error Resume Next
        TaxeRec.Close
        Err.Clear
        On Error GoTo DefErr
        TaxeRec.Open SQLStr, taxeAdoDB.ActiveConn
        If Not (TaxeRec.EOF) Then
          OrgAvp = TaxeRec!vk / 100
        End If
        SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + pzn$
        'Set TaxeRec = TaxeDB.OpenRecordset(SQLStr$)
        On Error Resume Next
        TaxeRec.Close
        Err.Clear
        On Error GoTo DefErr
        TaxeRec.Open SQLStr, taxeAdoDB.ActiveConn
        
        If (aktavp <= OrgAvp - 15) Or (aktavp <= OrgAvp * 0.85) Then
          RezArtikel(row%).Imp123 = 2
          RezArtikel(row%).ImpErspart = (OrgAvp - aktavp)
        Else
          RezArtikel(row%).Imp123 = 1
        End If
    Else
        RezArtikel(row%).Imp123 = 0
    End If
    
                
    hStr$ = Format$(TaxeRec!Abgabebest, "0")
    If (InStr("01245", hStr$) <> 0) Then
        RezArtikel(row%).Fam = 1
    Else
        RezArtikel(row%).Fam = 0
    End If
    
    RezArtikel(row%).AutIdem = 0
    RezArtikel(row%).N0 = 0
    Call CheckAutIdem(row%)
End If

RezArtikel(row%).MagIndex = 0

If (row% >= AnzRezeptArtikel%) Then
    AnzRezeptArtikel% = row% + 1
End If

'If (TaxeRec!BTMKz) And (BtmRezept% = False) And (SonderBelegRezept% = False) And (PrivatRezept% = False) And (IvfRezept% = False) Then
'4.0.52: Prüfung auf PrivatRezept weg
If (TaxeRec!BtmKz) And (BtmRezept% = False) And (SonderBelegRezept% = False) And (IvfRezept% = False) Then
    BtmRezept% = True
    iTabIndex% = AktTabIndex%
    Call PruefeBtmKosten
    sBrutto$ = RezAddGesamtBrutto$(0)
    Call PaintRezept
    Call RepaintBtmGebühr
    Call TabIndexChange(iTabIndex%, True)
End If

Call DefErrPop
End Sub

Function PrivatRezeptPreis&()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("PrivatRezeptPreis&")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
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
Dim ok%
Dim ret&, StammVK&

StammVK = 0
SQLStr$ = "SELECT * FROM ARTIKEL WHERE PZN = " + PznString(TaxeRec!pzn)
SQLStr = SQLStr + " AND LagerKz<>0"
FabsErrf = Artikel.OpenRecordset(ArtikelAdoRec, SQLStr)
If (FabsErrf = 0) Then
    StammVK = ast.AVP * 100
End If

If (TaxeRec!MwStKZ) Then
    ActMwst% = para.Mwst(1)
Else
    ActMwst% = para.Mwst(2)
End If

ret = 0
If TaxeRec!AmpvAMG = 2 Then
    ret = TaxeRec!vk
Else
    ok = False
    If (PrivatPreisTyp = PRIVATPREIS_AMPV) Then
        If (TaxeRec!EK > 0) Then
            ret = CalcAMPV(TaxeRec!EK / 100, ActMwst%) * 100
            ok = True
        End If
    Else
        If (StammVK > 0) Then
            ret = StammVK
            ok = True
        ElseIf (TaxeRec!vk > 0) Then
            ret = TaxeRec!vk
            ok = True
        End If
    End If
    
    If Not ok Then                      'hat nicht geklappt, andere Art der Preisermittlung versuchen
        If (PrivatPreisTyp = PRIVATPREIS_AMPV) Then
            If (StammVK > 0) Then
                ret = StammVK
                ok = True
            ElseIf (TaxeRec!vk > 0) Then
                ret = TaxeRec!vk
                ok = True
            End If
        Else
            If (TaxeRec!EK > 0) Then
    '            EK = TaxeRec!EK / 100#
                ret = CalcAMPV(TaxeRec!EK / 100, ActMwst%) * 100
                ok = True
            End If
        End If
    End If
End If

PrivatRezeptPreis = ret

Call DefErrPop
End Function
    
Sub MKdurchKK(row%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("MKdurchKK")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
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

SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + RezArtikel(row%).pzn
'Set TaxeRec = TaxeDB.OpenRecordset(SQLStr$)
On Error Resume Next
TaxeRec.Close
Err.Clear
On Error GoTo DefErr
TaxeRec.Open SQLStr, taxeAdoDB.ActiveConn
If (TaxeRec.EOF) Then
    Call DefErrPop: Exit Sub
End If
  

If (PrivatRezept%) Then
ElseIf (PrivatRezept% = 0) And (TaxeRec!FestKz) And (TaxeRec!FESTBETRAG < TaxeRec!vk) And (ActKasse% <> 7) Then
    If (MessageBox("Mehrkostenübernahme durch Krankenkasse ?", vbQuestion Or vbYesNo Or vbDefaultButton1, Trim(RezArtikel(row).text)) = vbYes) Then
        With RezArtikel(row%)
            RezArtikel(row%).PlusMehrKosten = 1
            RezArtikel(row%).Preis = TaxeRec!vk / 100#
            Call PaintArtikel(row%)
            Call TabIndexChange(AktTabIndex%, False)
        End With
    End If
End If

Call DefErrPop
End Sub
    
Sub WunschArtikel(row%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("WunschArtikel")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
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

If (MessageBox("Handelt es sich um ein Wunscharzneimittel?", vbQuestion Or vbYesNo Or vbDefaultButton2, Trim(RezArtikel(row).text)) = vbYes) Then
    Debug.Print "Wunsch" + Str(row)
    With RezArtikel(row%)
        RezArtikel(row%).Preis = 0#
        RezArtikel(row%).Zuz = 0
        Call PaintArtikel(row%)
        Call TabIndexChange(AktTabIndex%, False)
    End With
    Call DefErrPop: Exit Sub
    
    If (ParentForm.ActiveControl.Name = ParentForm.cboVerfügbarkeit(0).Name) Then
        Call TabIndexChange((row * 4) + 3, True)
        Call PaintArtikel(row%)
        ParentForm.picRezept.SetFocus
    End If
End If

Call DefErrPop
End Sub

Function CheckSonderPzn%(pzn$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("CheckSonderPzn%")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
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
    
CheckSonderPzn% = -1
For i% = 0 To UBound(SonderPzn$)
    If (Left$(SonderPzn$(i%), 8) = pzn$) Then
        CheckSonderPzn% = i%
        Exit For
    End If
Next i%

Call DefErrPop
End Function

Function CheckBtmSonderPzn%(pzn$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("CheckBtmSonderPzn%")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
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
    
CheckBtmSonderPzn% = -1
For i% = 0 To UBound(BtmSonderPzn$)
    If (Left$(BtmSonderPzn$(i%), 8) = pzn$) Then
        CheckBtmSonderPzn% = i%
        Exit For
    End If
Next i%

Call DefErrPop
End Function


Sub PaintArtikel(row%, Optional MitTab% = True)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("PaintArtikel")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
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
Dim i%, j%, BlinkFlag%, IstRot%, hImp123%
Dim h$, h2$
        
If (Selbsterklaerung = 0) Then
    j% = 3 + row% * 4
    For i% = 0 To 3
        Call PaintRezeptBox(j% + i%)
    Next i%
    If (row% <= 2) Then
        Call PaintRezeptBox(43 + row%)
        Call PaintRezeptBox(50 + row%)
    End If
End If

BlinkFlag% = RezArtikel(row%).flag And REZ_BLINK


h$ = RezArtikel(row%).pzn
'If (Left$(h$, 1) = "@") Then h$ = RezArtikel(row%).ScreenPzn
If (Left$(h$, 1) = "-") Then
    h$ = RezArtikel(row%).ScreenPzn
End If
Call PaintRezeptWert(j% + 1, h$, BlinkFlag%)

h$ = RezArtikel(row%).pzn
If (Left(h, 1) = "-") Then
    h = Mid(h, 2) + "0"
    h = MakePzn(h)
End If

h2 = RezArtikel(row%).text
'If (Val(h) > 0) And (RezArtikel(row).pzn <> "-ZUSG_HM") Then
If (Val(h) > 0) And (RezArtikel(row).pzn <> "-9999999") Then
    SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + h$
    'Set TaxeRec = TaxeDB.OpenRecordset(SQLStr$)
    On Error Resume Next
    TaxeRec.Close
    Err.Clear
    On Error GoTo DefErr
    TaxeRec.Open SQLStr, taxeAdoDB.ActiveConn
    If Not (TaxeRec.EOF) Then
        h2$ = Trim(TaxeRec!Name) + TaxeRec!menge + TaxeRec!einheit
    End If
End If
If (RezArtikel(row%).N0) Then
    h2 = h2 + " (N0)"
End If
Call PaintRezeptWert(j%, h2, BlinkFlag%)
'Call PaintRezeptWert(j% + 1, h$, BlinkFlag%)

If (RezArtikel(row%).MietDauer > 0) Then
    Call PaintRezeptWert(j% + 2, Format(RezArtikel(row%).MietDauer, "0"), BlinkFlag%)
Else
    Call PaintRezeptWert(j% + 2, Format(RezArtikel(row%).Faktor, "0"), BlinkFlag%)
End If
Call PaintRezeptWert(j% + 3, Format(RezArtikel(row%).Preis * 100 * RezArtikel(row%).Faktor, "0"), BlinkFlag%)

If (row% <= 2) Then
    h$ = ""
    If (RezArtikel(row%).AutIdem) Then h$ = "!"
    If (RezArtikel(row%).AutIdemKreuz = 2) Then h$ = Chr(67)
    
    If (para.Newline) And (RezArtikel(row%).AutIdem = 3) Then
        h$ = Chr$(214)
'        Call PaintRezeptBox(43 + row%, RGB(206, 185, 23))
        Call PaintRezeptBox(43 + row%, vbGreen)
        Call PaintRezeptWert(43 + row%, h$, False, 2)
    ElseIf (para.Newline) And (RezArtikel(row%).AutIdem = 4) Then
'        Call PaintRezeptBox(43 + row%, RGB(135, 61, 52))
        Call PaintRezeptBox(43 + row%, vbRed)
        Call PaintRezeptWert(43 + row%, h$, False, 2)
    ElseIf (RezArtikel(row%).AutIdem = 2) Then
        Call PaintRezeptBox(43 + row%, vbYellow)
        Call PaintRezeptWert(43 + row%, h$, False, True)
    ElseIf (RezArtikel(row%).AutIdem = 3) Then
        Call PaintRezeptBox(43 + row%, vbGreen)
        Call PaintRezeptWert(43 + row%, h$, False, True)
    ElseIf (RezArtikel(row%).AutIdem = 4) Then  '2
        Call PaintRezeptBox(43 + row%, vbRed)
        Call PaintRezeptWert(43 + row%, h$, False, True)
    Else
        Call PaintRezeptWert(43 + row%, h$)
    End If
    
    h$ = " "
    IstRot% = False
    hImp123% = RezArtikel(row%).Imp123
    If (hImp123% > 0) Then
        If (hImp123% And 4) Then
            h$ = "i"
        Else
            h$ = "I"
        End If
        If (hImp123% And 1) Then
            IstRot% = True
        End If
    End If
    If (IstRot%) Then
        Call PaintRezeptBox(50 + row%, vbRed)
        Call PaintRezeptWert(50 + row%, h$, False, True)
    Else
        Call PaintRezeptWert(50 + row%, h$)
    End If
End If

If (MitTab) Then
    Call TabIndexChange(AktTabIndex%, False)
End If
Call CalcRezeptWerte

Call DefErrPop
End Sub

Sub InsertArtikel()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("InsertArtikel")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
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
            
If (AktTabIndex% >= 3) Then
    row% = (AktTabIndex% - 3) \ 4
    If (row% < AnzRezeptArtikel%) Then
        For i% = UBound(RezArtikel) - 1 To row% Step -1
            RezArtikel(i% + 1) = RezArtikel(i%)
        Next i%
        Call InitRezeptArtikel(row%)
        Call InitKontrollStop(row%)
            
        AMverfügbarkeit$ = CheckVerfügbarkeit
        If (AMverfügbarkeit$ <> "") And (row% <= 2) Then
            If (row% = 0) Then
                AMverfügbarkeit$ = "1" + Left$(AMverfügbarkeit$, 2)
            ElseIf (row% = 1) Then
                AMverfügbarkeit$ = Left$(AMverfügbarkeit$, 1) + "1" + Mid$(AMverfügbarkeit$, 2, 1)
            Else
                AMverfügbarkeit$ = Left$(AMverfügbarkeit$, 2) + "1"
            End If
        End If
        
        AnzRezeptArtikel% = AnzRezeptArtikel% + 1
        Call PaintRezept
        Call TabIndexChange(3 + (row% * 4), True)
    End If
End If

Call DefErrPop
End Sub

Sub LoescheArtikel()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("LoescheArtikel")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
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
Dim i%, row%, erg%
            
If (AktTabIndex = 59) Then
    erg% = iMsgBox("Wollen Sie die IK-Nummer tatsächlich rücksetzen/entfernen ?", vbYesNo Or vbDefaultButton2 Or vbInformation)
    If (erg% = vbYes) Then
        AutIdemIk& = 0
        AutIdemKbvNr& = 0
        Call PaintRezeptBox(59)
        Call TabIndexChange(AktTabIndex%, False)
    End If
ElseIf (AktTabIndex% >= 3) Then
    erg% = iMsgBox("Wollen Sie den Artikel tatsächlich aus dem Rezept entfernen ?", vbYesNo Or vbDefaultButton2 Or vbInformation)
    If (erg% = vbYes) Then
        row% = (AktTabIndex% - 3) \ 4
        If (row% < AnzRezeptArtikel%) Then
            For i% = row% To UBound(RezArtikel) - 1
                RezArtikel(i%) = RezArtikel(i% + 1)
            Next i%
            Call InitRezeptArtikel(UBound(RezArtikel))
            
            AMverfügbarkeit$ = CheckVerfügbarkeit
            If (AMverfügbarkeit$ <> "") And (row% <= 2) Then
                If (row% = 0) Then
                    AMverfügbarkeit$ = Mid$(AMverfügbarkeit$, 2) + "1"
                ElseIf (row% = 1) Then
                    AMverfügbarkeit$ = Left$(AMverfügbarkeit$, 1) + Mid$(AMverfügbarkeit$, 3, 1) + "1"
                Else
                    AMverfügbarkeit$ = Left$(AMverfügbarkeit$, 2) + "1"
                End If
            End If
            AnzRezeptArtikel% = AnzRezeptArtikel% - 1
            Call PaintRezept
            Call TabIndexChange(3 + (row% * 4), True)
            
            Call CheckVerordnungMDB
        End If
    End If
End If

Call DefErrPop
End Sub

Sub InitRezeptArtikel(row%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("InitRezeptArtikel")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
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
        
With RezArtikel(row%)
    .pzn = Space$(Len(.pzn))
    .text = Space$(Len(.text))
    .Preis = 0
    .flag = 0
    .Imp123 = 0
    .ImpErspart = 0
    .Fam = 0
    .AutIdem = 0
    .N0 = 0
    .MagIndex = 0
    .Zuz = 0
    .appli = 0
    .Faktor = 0
    .zusatz(0) = Space$(Len(.zusatz(0)))
    .zusatz(1) = Space$(Len(.zusatz(1)))
    .NichtInTaxe = 0
    .MietDauer = 0
    .HerstRabattPrivat130Brutto = 0
    .VdbPauschale = 0
    .TkkPznDruck = 0
    .VebNr = 0
    .PauschaleNr = 0
    .Fortsetzung = 0
    .HmAbrechnungsKz = Space(2)
    .LEGS = Space(7)
End With

Call DefErrPop
End Sub

Sub InitKontrollStop(row%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("InitKontrollStop")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
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
Dim l%
       
With RezArtikel(row%)
    l% = Len(.text)
    .pzn = "0" + String(7, "9")
    .text = Left$("KONTROLL-STOP       (PZN?)" + Space$(l%), l%)
    .Preis = 0
    .flag = REZ_K_STOP Or REZ_BLINK
    .Faktor = 1
    .Zuz = ZuzFeld!(3)
End With

Call DefErrPop
End Sub

Sub InitRezeptDruck()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("InitRezeptDruck")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
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
Dim i%, j%, k%, hFaktor%, hFaktor2%, ZusatzKompontenFaktor%, ind%, gef%, AvpMonat%, HASH_ID%, MaxGruppenLen%, MaxGruppenInd%, Noctu%, BtmApoDa%
Dim iMilli&, iSekunden&, iMinuten&, iStunden&, iTage&, iMonate&, l&
Dim lPreis&, SollPreis&, IstPreis&, pz&, Pruefsumme&, TaskId&
Dim dHash#, dStückPreis#
Dim ch$, h$, h2$, PznStr$, PznStr2$, hPosNr$, SQLStr$, HashDezimal$, OrgHash$, HashHex$, GruppeNr$, ProduktGruppen$
Dim AvpRec As Recordset
Dim HerstellDat As Date
      
RezDruckInd% = 0
MagSpeicherIndex% = -1
Noctu = 0
ZusatzKompontenFaktor = 1

If (SonderBelegRezept%) Then
    With SonderBelege(SonderBelegRezept% - 1)
        Call SpeicherDruckZeile(Trim(.KkBez))
        Call SpeicherDruckZeile(Trim(RezKundenInfo$(0)))
        Call SpeicherDruckZeile(Trim("Sonderbeleg"))
        Call SpeicherDruckZeile(Trim(.KassenId))
        Call SpeicherDruckZeile(Trim(RezKundenInfo$(3)))
        Call SpeicherDruckZeile(Trim(.Status))
        Call SpeicherDruckZeile(Trim(RezApoNr$))
        Call SpeicherDruckZeile(Trim(.GültigBis))
        Call SpeicherDruckZeile(Trim(VerkDatum$))
    End With
End If

Dim HashFlag As Boolean
HashFlag = False
If (ParenteralRezept >= 0) Or (AuseinzelungPzn(0) <> "") Then
    HashFlag = True
ElseIf (AlleRezepturenMitHash) And (RezArtikel(0).MagIndex > 0) And (RezArtikel(0).Preis > 0) Then
    ParenteralRezept = -99
    HashFlag = True
ElseIf (LennartzPzn > 0) Then
    HashFlag = True
End If

If (AvpTeilnahme%) Then
    If (AvpRezeptNr$ = "") Then
        AvpLaufNr& = 1
'        SQLStr$ = "SELECT MAX(AvpLaufNr) as MaxAvpLaufNr FROM Rezepte WHERE AvpMonat=" + Str$(AvpMonat%)
        SQLStr$ = "SELECT MAX(AvpLaufNr) as MaxAvpLaufNr FROM Rezepte"
        Set AvpRec = RezSpeicherDB.OpenRecordset(SQLStr$)
        If (AvpRec.RecordCount > 0) Then
            AvpLaufNr& = CheckNullLong(AvpRec!MaxAvpLaufNr) + 1
        End If
        AvpRezeptNr$ = Format(AvpLaufNr& Mod 1000000, "000000") + Mid$(VerkDatum$, 3)
    End If
    
    pz& = 0
    For i% = 1 To 10
        ch$ = Mid$(AvpRezeptNr$, i%, 1)
        pz& = pz& + (i% + 1) * Val(ch$)
    Next i%
    pz& = pz& Mod 11
    If (pz& = 10) Then
        pz& = 0
    End If
    AvpRezeptNr$ = Left$(AvpRezeptNr$, 10) + Format(pz&, "0")

    If (Selbsterklaerung) Then
        Call SpeicherDruckZeile("X" + "Selbsterklaerung" + "X")
    Else
        Call SpeicherDruckZeile("X" + AvpRezeptNr$ + "X")
    End If
ElseIf (FiveRxFlag) Or (Code128Flag) Then

'ElseIf (ParenteralRezept >= 0) Then
'    If (TransaktionsID = 0) Then
'        TransaktionsID& = 1
'        SQLStr$ = "SELECT MAX(TransaktionsID) as MaxTransaktionsID FROM Rezepte"
'        Set AvpRec = RezSpeicherDB.OpenRecordset(SQLStr$)
'        If (AvpRec.RecordCount > 0) Then
'            TransaktionsID& = CheckNullLong(AvpRec!Maxtransaktionsid) + 1
'        End If
'    End If
        
    If (TransaktionsID = 0) Then
        h = Format((Timer * 1000) Mod 1000, "000") + Format(Now, "SSNNHHDDMM")
        'h = "6784515161501"
        
        iMilli = Int(Val(Left$(h, 3)) / 334)
        iSekunden = Val(Mid$(h, 4, 2))
        iMinuten = Val(Mid$(h, 6, 2))
        iStunden = Val(Mid$(h, 8, 2))
        iTage = Val(Mid$(h, 10, 2))
        iMonate = Val(Mid$(h, 12, 2))
        
        TransaktionsID = iMilli + (iSekunden + (iMinuten * 60) + (iStunden * 3600) + ((iTage - 1) * 86400) + ((iMonate - 1) * 2678400)) * 3
        Do
            SQLStr$ = "SELECT * FROM Rezepte WHERE TransaktionsID=" + CStr(TransaktionsID)
            Set AvpRec = RezSpeicherDB.OpenRecordset(SQLStr$)
            If (AvpRec.EOF) Then
                Exit Do
            Else
                TransaktionsID = TransaktionsID + 1
            End If
        Loop
    End If
'    TransaktionsID = 33804244

    h$ = Right$(String(8, "0") + CStr(TransaktionsID), 8)
    Pruefsumme = 0
    For i = 1 To 8
      Pruefsumme = Pruefsumme + Val(Mid$(h, i, 1)) * (1 + 2 * ((i + 1) Mod 2))
    Next i
    h = h + Right$(CStr(Pruefsumme Mod 10), 1)
    
    Call SpeicherDruckZeile(h$)
    
    If (FiveRxFlag) Then
        '    h2$ = "20100115:161500"g
        '    h = "038044377"
        '    RezApoNr = "301234561"
        '    ParenteralPara = RezApoNr + h + h2$ + ":000"
        '    ParenteralPara = RezApoNr + h + Format(Now, "YYYYMMDD:HHNNSS") + ":000"
            
        '    ParenteralPara = RezApoNrPraefix + RezApoNr + h + Format(Now, "YYYYMMDD:000000") + ":000"
    
    '    HashErstellDat = Now
    
        h2 = VerkDatum
        HashErstellDat = Left(h2, 2) + "." + Mid(h2, 3, 2) + ".20" + Mid(h2, 5, 2) + " 10:00"
        
        HerstellDat = DateAdd("h", -1, HashErstellDat)
        
        If (ParEnteralHerstellerKey < 1) Then
            ParEnteralHerstellerKey = 3
        End If
        If (ParenteralRezept = -99) Or (LennartzPzn > 0) Then
            ParEnteralHerstellerKey = 3
            HashErstellDat = Left(h2, 2) + "." + Mid(h2, 3, 2) + ".20" + Mid(h2, 5, 2) + " 00:01"
            HerstellDat = HashErstellDat
            pCharge = CStr(ParEnteralHerstellerKey) + Right(String(9, "0") + RezApoNrPraefix + RezApoNr, 9) + Format(HerstellDat, "YYYYMMDD") + ":0000" + "01" + "01"
        ElseIf (ParenteralRezept > 15) Then
            ParEnteralHerstellerKey = 3
            
            If (ParenteralRezept < 24) Then
                HashErstellDat = Left(h2, 2) + "." + Mid(h2, 3, 2) + ".20" + Mid(h2, 5, 2) + " 00:02"
            End If
            
            pCharge = CStr(ParEnteralHerstellerKey) + Right(String(9, "0") + RezApoNrPraefix + RezApoNr, 9) + Format(HerstellDat, "YYYYMMDD") + IIf(ParenteralRezept < 24, ":0001", ":0000") + "01" + "01"
        Else
            pCharge = CStr(ParEnteralHerstellerKey) + Right(String(9, "0") + ParEnteralHerstellerKz(ParEnteralHerstellerKey), 9) + Format(HerstellDat, "YYYYMMDD:HHMM") + "01" + "01"
        End If
'        If (Trim(ParEnteralHerstellerKz(ParEnteralHerstellerKey)) = "") Then
'            pCharge = ""
'        End If
        ParenteralPara = RezApoNrPraefix + RezApoNr + h + Format(HashErstellDat, "YYYYMMDD:HHMM") + "00:000"
        
        If (ParenteralRezept > 15) And (ParenteralRezept < 24) Then
        Else
            ParenteralPara = ParenteralPara + pCharge
        End If
    
        TA1_V37 = (DateDiff("d", "01.07.2022", HashErstellDat) >= 0)
        
    '    ParenteralPara = "20" + RezApoNr + h + Format(Now, "YYYYMMDD:000000") + ":000"
    '    ParenteralPara = "20000002327855284720100415:000000:000"
    End If
End If

Call SpeicherDruckZeile("+" + RezApoNr$ + "+")

If (PrivatRezept%) Or (SonderBelegRezept%) Or ((Selbsterklaerung) And (SchutzMasken = 0)) Then
    h$ = "       "
ElseIf (RezGebSumme# = 0#) Then
    h$ = "   0   "
Else
    h$ = Format(RezGebSumme#, "   0.00")
    For i% = 1 To Len(h$)
        If (Mid$(h$, i%, 1) = ".") Then
            Mid$(h$, i%, 1) = ","
        End If
    Next i%
End If
Call SpeicherDruckZeile(h$)


h$ = Format(RezSumme#, "0.00")
For i% = 1 To Len(h$)
    If (Mid$(h$, i%, 1) = ".") Then
        Mid$(h$, i%, 1) = ","
    End If
Next i%
Call SpeicherDruckZeile(h$)
      
BtmApoDa = 0

'If (ParenteralRezept >= 0) Or (AuseinzelungPzn(0) <> "") Then
If (HashFlag) Then
    If (ParenteralRezept >= 0) Then
        If (MagSpeicherIndex% < 0) And ((RezepturDruck%) Or (ParenteralRezept > 15)) Then
            If (RezArtikel(0).MagIndex > 0) Then    'And (RezArtikel(0).Preis > 0) Then
                MagSpeicherIndex% = RezArtikel(0).MagIndex
            End If
        End If
    ElseIf (ParenteralRezept = -99) Then
        If (MagSpeicherIndex% < 0) Then
            If (RezArtikel(0).MagIndex > 0) And (RezArtikel(0).Preis > 0) Then
                MagSpeicherIndex% = RezArtikel(0).MagIndex
            End If
        End If
    End If

    PznStr$ = RezArtikel(0).pzn
    hFaktor% = RezArtikel(0).Faktor
    hFaktor2% = hFaktor%
        
    If (PznStr = "06460749") Then
        Call SpeicherDruckZeile("06460748")
    ElseIf (PznStr = "02567114") Then
        Call SpeicherDruckZeile("02567113")
    ElseIf (PznStr = "02567115") Then
        Call SpeicherDruckZeile("02567113")
    Else
        Call SpeicherDruckZeile(RezArtikel(0).pzn)
    End If
    
    
'    If (hFaktor2% = 1) Then
'        h$ = "   "
'    Else
        h$ = Format(hFaktor2%, "  0")
'    End If
    Call SpeicherDruckZeile(h$)
    
'    lPreis& = CLng(RezArtikel(0).Preis * hFaktor% * 100#)
    h$ = Format(RezArtikel(0).Preis * hFaktor% * 100#, "0")
    Call SpeicherDruckZeile(h$)
    
    If (AuseinzelungPzn(0) <> "") Then
        h = ""
        
        k = 0
'        If (Len(AuseinzelungPzn(k)) = 8) Then
'            h$ = h + AuseinzelungPzn(k) + "11" + Format(AuseinzelungFaktor(k), "00000") + "11" + Format(AuseinzelungPreis(k) * 100, "000000000")
'        Else
'            h$ = h + "0" + AuseinzelungPzn(k) + "11" + Format(AuseinzelungFaktor(k), "00000") + "11" + Format(AuseinzelungPreis(k) * 100, "000000000")
'        End If
        If (Len(AuseinzelungPzn(k)) = 8) Then
            h$ = h + AuseinzelungPzn(k) + "11" + HashFaktor(CDbl(AuseinzelungFaktor(k))) + "11" + HashPreis(AuseinzelungPreis(k) * 100)
        Else
            h$ = h + "0" + AuseinzelungPzn(k) + "11" + HashFaktor(CDbl(AuseinzelungFaktor(k))) + "11" + HashPreis(AuseinzelungPreis(k) * 100)
        End If
        RezArtikel(0).zusatz(1) = h + IIf(AuseinzelungBtm, "1", "0")
        
        k = 1
        If (AuseinzelungPzn(k) <> "") Then
'            If (Len(AuseinzelungPzn(k)) = 8) Then
'                h$ = h + AuseinzelungPzn(k) + "11" + Format(AuseinzelungFaktor(k), "00000") + "11" + Format(AuseinzelungPreis(k) * 100, "000000000")
'            Else
'                h$ = h + "0" + AuseinzelungPzn(k) + "11" + Format(AuseinzelungFaktor(k), "00000") + "11" + Format(AuseinzelungPreis(k) * 100, "000000000")
'            End If
            If (Len(AuseinzelungPzn(k)) = 8) Then
                h$ = h + AuseinzelungPzn(k) + "11" + HashFaktor(CDbl(AuseinzelungFaktor(k))) + "11" + HashPreis(AuseinzelungPreis(k) * 100)
            Else
                h$ = h + "0" + AuseinzelungPzn(k) + "11" + HashFaktor(CDbl(AuseinzelungFaktor(k))) + "11" + HashPreis(AuseinzelungPreis(k) * 100)
            End If
        End If
        
        If (AuseinzelungBtm) Then
'            h$ = h + "02567001" + "11" + Format(1000, "00000") + "81" + Format(3.58 * 100, "000000000")
            h$ = h + "02567001" + "11" + HashFaktor(1000) + "81" + HashPreis(3.58 * 100)
        End If
        
        ParenteralPara = ParenteralPara + h$
    ElseIf (LennartzPzn > 0) Then
        Call LennartzRezeptur(0)
    ElseIf (MagSpeicherIndex > 0) Then
        Load frmTaxieren
    End If
    
    If (ParenteralHash = "") Then
'        Call MsgBox(ParenteralPara)
'        HashHex = "d55d95cbc247045a452928da59cb5a05"
'        HashHex = "aaaaaaaaaaaaaaaa5555555555555555"
    
        h = "Hash" + para.User + ".dat"
        On Error Resume Next
        Kill h
        On Error GoTo DefErr
        
        Dim Hash_Log%
        Hash_Log = FreeFile
        Open "Hash_Op.log" For Append As #Hash_Log
        Print #Hash_Log%, Format(Now, "DD.MM.YYYY HH:MM:SS  ") + "<"; ParenteralPara + ">"
        Close #Hash_Log%

        
'ParenteralPara = "30123456103804437720100115:161500:0001555734110010011000000138149441211000831100003039214943231100083110000034761494375110008311000015800"
'ParenteralPara = "30000002327855284720100415:000000:00036420581101000110001500000954739110005011000000004"
'ParenteralPara = "20000002327858937520100415:000000:0007283604110100011000026406336923711010001100001763109547451100050110000000094633825110004211000003350"
'ParenteralPara = "30999999744177095920100616:000000:000086864911010001100000647388715621101000110000002633417746110033311000000267341777511008001100000022730409801100050110000004890117162110008311000033600"
'ParenteralPara = "30790262132779846220210503:100000:000200000000020210503:0900010106460518110100074000003000103947731100400140000464890304098011008001400000051211241847110100014000003050"
'ParenteralPara = "30790262132785261020210503:100000:000200000000020210503:090001010646051811010007400000300010394773110040014000010086"
'ParenteralPara = "30080255632760990020210503:100000:000299930024520210503:0900010106460518110100074000005000103947731100400140000464890304098011008001400000050511241847110100014000003050"
TaskId& = Shell("Md5Code.exe " + ParenteralPara, vbNormalFocus)
        Call WarteAufTaskEnde(TaskId&, frmAction) ', frmAction)
        If (Dir(h) <> "") Then
            HASH_ID = FileOpen(h, "I")
            Line Input #HASH_ID, HashHex
            Close #HASH_ID
        End If
        h = HashHex
        HashDezimal = "0"
        For i% = 1 To 32
            OrgHash = HashDezimal
            For j = 1 To 15
                HashDezimal = StringAdd(HashDezimal, OrgHash)
            Next j
            
            ch = "&H" + Mid$(h$, i, 1)
            HashDezimal = StringAdd(HashDezimal, Format(Val(ch), "00"))
        Next i
        HashDezimal = Right(String(40, "0") + HashDezimal, 40)
        ParenteralHash = HashDezimal
        
'        h$ = Left$(HashHex, 16)
'        HashHex = Mid(HashHex, 17)
'        For k = 0 To 1
'            HashDezimal = "0"
'            For i% = 1 To 16
'                OrgHash = HashDezimal
'                For j = 1 To 15
'                    HashDezimal = StringAdd(HashDezimal, OrgHash)
'                Next j
'
'                ch = "&H" + Mid$(h$, i, 1)
'                HashDezimal = StringAdd(HashDezimal, Format(Val(ch), "00"))
'            Next i
'            HashDezimal = Right(String(20, "0") + HashDezimal, 20)
'            ParenteralHash = ParenteralHash + HashDezimal
'
'            h$ = Left$(HashHex, 16)
'        Next k
    Else
        BereitsGedruckt = True
    End If

    For k = 0 To 1
        HashDezimal = Mid$(ParenteralHash, k * 20 + 1, 20)
        Call SpeicherDruckZeile(Left(HashDezimal, 10))
        Call SpeicherDruckZeile(Mid(HashDezimal, 11, 3))
        Call SpeicherDruckZeile(Mid(HashDezimal, 14, 7))
    Next k
    
    j = 3
    
Else
    For j% = 0 To (AnzRezeptArtikel% - 2)
        If (Val(RezArtikel(j%).pzn) = 2567018) Then
            If (Trim(RezArtikel(0).text) = "Nichtverfügbarkeit von AM") Then
                If (j <= 3) Then
                    h = Format(RezArtikel(0).Faktor, "000")
                    If (j = 1) Then
                        h = Mid(h, 2) + "1"
                    ElseIf (j = 2) Then
                        h = Mid(h, 1, 1) + Mid(h, 3) + "1"
                    Else
                        h = Mid(h, 1, 2) + "1"
                    End If
                    RezArtikel(0).Faktor = Val(h)
                End If
            End If
            
            RezArtikel(AnzRezeptArtikel) = RezArtikel(j)
            For k = j To (AnzRezeptArtikel - 1)
                RezArtikel(k) = RezArtikel(k + 1)
            Next k
            Exit For
        End If
    Next j
    
'    If (RezeptMitHilfsmittelNr% = 1) Then
'        Dim AnzMKs%, MKind%
'        AnzMKs = 0
'        For j% = 0 To (AnzRezeptArtikel% - 1)
'            If (Left$(RezArtikel(j%).pzn, 1) <> "-") Then
'                RezArtikel(j).zusatz(1) = Space$(Len(RezArtikel(j).zusatz(1)))
'            End If
'
'            If (InStr(RezArtikel(j).zusatz(1), "MK ") > 0) Then
'                AnzMKs = AnzMKs + 1
'            End If
'        Next
'        For j% = 0 To (AnzRezeptArtikel% - 1)
'            MKind = InStr(RezArtikel(j).zusatz(1), "MK ")
'            If (MKind > 0) Then
'                Call SpeicherDruckZeile("06460725")
'
'                h$ = Format(AnzMKs + j + 1, "  0")
'                Call SpeicherDruckZeile(h$)
'
'                lPreis& = CLng(xVal(Mid(RezArtikel(j).zusatz(1), MKind + 3)) * RezArtikel(j).Faktor * 100#)
'                h$ = Format(lPreis, "0")
'                Call SpeicherDruckZeile(h$)
'            End If
'        Next
'    End If
    
    Dim AnzMKs%, MKind%
    Dim Druck_HmAbrechnungsKz$
    AnzMKs = 0
    Druck_HmAbrechnungsKz$ = ""
    For j% = 0 To (AnzRezeptArtikel% - 1)
        If (InStr(RezArtikel(j).zusatz(1), "MK ") > 0) Then
            AnzMKs = AnzMKs + 1
        End If
        If (Trim(RezArtikel(j).HmAbrechnungsKz) <> "") Then
'            Druck_HmAbrechnungsKz$ = Druck_HmAbrechnungsKz$ + Trim(RezArtikel(j).HmAbrechnungsKz) + " "
            Druck_HmAbrechnungsKz$ = Druck_HmAbrechnungsKz$ + Trim(RezArtikel(j).LEGS) + " "
        End If
    Next
    Druck_HmAbrechnungsKz$ = Trim(Druck_HmAbrechnungsKz$)
    For j% = 0 To (AnzRezeptArtikel% - 1)
        MKind = InStr(RezArtikel(j).zusatz(1), "MK ")
        If (MKind > 0) Then
            Call SpeicherDruckZeile("06460725")
            
            h$ = Format(AnzMKs + j + 1, "  0")
            Call SpeicherDruckZeile(h$)
            
            lPreis& = CLng(xVal(Mid(RezArtikel(j).zusatz(1), MKind + 3)) * RezArtikel(j).Faktor * 100#)
            h$ = Format(lPreis, "0")
            Call SpeicherDruckZeile(h$)
        End If
    Next
    
    'SollPreis& = 0
    For j% = 0 To (AnzRezeptArtikel% - 1)
        If (MagSpeicherIndex% < 0) And (RezepturDruck%) Then
            If (RezArtikel(j%).MagIndex > 0) And (RezArtikel(j%).Preis > 0) Then
                MagSpeicherIndex% = RezArtikel(j%).MagIndex
            End If
        End If
        
        If (j% >= 7) Then Exit For
        If (BtmRezept%) Then
            If (j% >= 4) Then Exit For
            
            If (j = 3) And (Val(RezArtikel(0).pzn) = SONDERPZN_WVO) Then
                BtmApoDa = True
                Call SpeicherDruckZeile("@BTM2@" + BtmRezDruckName$)
            Else
                If (j% = 3) And (Val(RezArtikel(0).pzn) <> 2567024) Then Exit For
            End If
        End If
        
        If (j = 2) And (Selbsterklaerung) Then
            h2 = BoxenWert(52)
            Call SpeicherDruckZeile("@ABDATUM@" + h2$)
        End If
    
        If (j = 3) And (BtmRezept% = 0) And (Trim(RezNr) <> "") And (RezeptNrPositionAlt = 0) And (Selbsterklaerung = 0) Then
            h2 = RezNr
            'h2 = "9892002021470"

            If (Left$(h2$, 7) = "9999998") Then h2$ = Mid$(h2$, 8, 5)
            Call SpeicherDruckZeile("@REZNR@" + h2$ + IIf(Druck_HmAbrechnungsKz$ <> "", "  LEGS:" + Druck_HmAbrechnungsKz$, ""))
        End If
    
        PznStr$ = "       "
        If (Val(RezArtikel(j%).pzn) <> 9999999) Then
            PznStr$ = RezArtikel(j%).pzn
        End If
        
        hFaktor% = RezArtikel(j%).Faktor
        If (Val(RezArtikel(j%).pzn) = 9999057) Then
            hFaktor% = 0
        End If
        
        hFaktor2% = hFaktor%
    
'        If (RezArtikel(j%).pzn = "2567024") And (AMverfügbarkeitsPzn = 0) Then
        If (Val(RezArtikel(j%).pzn) = 2567018) Then
            Noctu = True
        End If
        
        gef% = 0
        dStückPreis = 0
'        If (Trim(PznStr$) <> "") And (Left$(PznStr, 1) <> "@") Then
        PznStr2 = PznStr
        If (Left$(PznStr2, 1) = "-") Then
            PznStr2 = Mid(PznStr2, 2) + "0"
            PznStr2 = MakePzn(PznStr2)
        End If
        
'        If (Trim(PznStr2$) <> "") And (Left$(PznStr2, 1) <> "-") Then
'            If (AplusVOk) Then
'                SQLStr$ = "SELECT * FROM VdbArtikel WHERE PZN = " + PznStr2
''                Set VdbPznRec = AplusVDB.OpenRecordset(SQLStr$)
'                FabsErrf = AplusVDB.OpenRecordset(VdbPznRec, SQLStr)
'                If Not (VdbPznRec.EOF) Then
'                    MaxGruppenLen = 0
'                    hPosNr$ = Trim(Left$(CheckNullStr(VdbPznRec!HmPositionsNr1), 10))
'                    If (hPosNr$ <> "") Then
'                        For k = 0 To 2
'                            h$ = Space$(100)
'                            l& = GetPrivateProfileString("ProduktGruppen", "Kassentyp" + CStr(ActKasse) + "," + CStr(k + 1), h$, h$, 101, CurDir + "\HmDruck.ini")
'                            h$ = Left$(h$, l&)
'                            ProduktGruppen$ = RTrim$(h$)
'
'                            If (ProduktGruppen <> "") Then
'                                ProduktGruppen = ProduktGruppen + ","
'                                Do
'                                    ind = InStr(ProduktGruppen, ",")
'                                    If (ind <= 0) Then
'                                        Exit Do
'                                    End If
'
'                                    h = Trim(Left$(ProduktGruppen, ind - 1))
'                                    l = Len(h)
'                                    If (l > MaxGruppenLen) Then
'                                        If (Left(hPosNr, l) = h) Then
'                                            MaxGruppenInd = k
'                                            MaxGruppenLen = l
'                                        End If
'                                    End If
'
'                                    ProduktGruppen = Mid(ProduktGruppen, ind + 1)
'                                Loop
'                            End If
'                        Next k
'                    End If
'                    If (MaxGruppenLen > 0) Then
''                        PznStr$ = hPosNr$
'
'                        If (MaxGruppenInd = 2) Then
'                            HmStückPreis#(j) = RezArtikel(j%).Preis * hFaktor%
'                            HmFaktor(j) = hFaktor
'                            HmNummer$(j) = hPosNr   ' PznStr
'                        Else
'
'                            h$ = Trim(CheckNullStr(VdbPznRec!menge))
'                            If (h$ <> "") Then
'                                If (Val(h$) > 0) Then
'                                    hFaktor2% = Val(h$)
'                                    gef% = True
'                                End If
'                            End If
'
'                            If (gef = 0) Then
'                                SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + RezArtikel(j%).pzn
'                                'Set TaxeRec = TaxeDB.OpenRecordset(SQLStr$)
'                                On Error Resume Next
'                                TaxeRec.Close
'                                Err.Clear
'                                On Error GoTo DefErr
'                                TaxeRec.Open SQLStr, taxeAdoDB.ActiveConn
'                                If (TaxeRec.EOF = False) Then
'                                    hFaktor2% = MengeErmitteln(TaxeRec!menge)
'                                    gef% = True
'                                End If
'                            End If
'
'                            If (gef%) Then
'                                If (MaxGruppenInd = 0) Then
'                                    dStückPreis = RezArtikel(j%).Preis / hFaktor2
'                                Else
'                                    dStückPreis = RezArtikel(j%).Preis * hFaktor%
'                                End If
'                                HmStückPreis#(j) = dStückPreis
'                                HmNummer$(j) = hPosNr   'PznStr
'                                hFaktor2% = hFaktor2% * hFaktor%
'                            End If
'                        End If
'                    End If
'                End If
'            End If
'        End If
        
        If (RezeptMitHilfsmittelNr% = 1) Then
            If (Trim(PznStr2$) <> "") And (Left$(PznStr2, 1) <> "-") Then
                If (AplusVOk) Then
                    SQLStr$ = "SELECT * FROM VdbArtikel WHERE PZN = " + PznStr2
    '                Set VdbPznRec = AplusVDB.OpenRecordset(SQLStr$)
                    FabsErrf = AplusVDB.OpenRecordset(VdbPznRec, SQLStr)
                    If Not (VdbPznRec.EOF) Then
                        MaxGruppenLen = 0
                        hPosNr$ = Trim(Left$(CheckNullStr(VdbPznRec!HmPositionsNr1), 10))
                        If (hPosNr$ <> "") Then
                            MaxGruppenInd = 1
                            
                            h$ = Trim(CheckNullStr(VdbPznRec!menge))
                            If (h$ <> "") Then
                                If (Val(h$) > 0) Then
                                    hFaktor2% = Val(h$)
                                    gef% = True
                                End If
                            End If
                            
                            If (gef = 0) Then
                                SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + PznStr2$  ' RezArtikel(j%).pzn
                                'Set TaxeRec = TaxeDB.OpenRecordset(SQLStr$)
                                On Error Resume Next
                                TaxeRec.Close
                                Err.Clear
                                On Error GoTo DefErr
                                TaxeRec.Open SQLStr, taxeAdoDB.ActiveConn
                                If (TaxeRec.EOF = False) Then
                                    hFaktor2% = MengeErmitteln(TaxeRec!menge)
                                    If (RezArtikel(j).ScreenPzn = ZusatzKomponentenBasisNr) Then
                                        ZusatzKompontenFaktor = hFaktor2
                                    End If
                                    gef% = True
                                End If
                            End If
                        
                            If (gef%) Then
                                If (MaxGruppenInd = 0) Then
                                    dStückPreis = RezArtikel(j%).Preis / hFaktor2
                                Else
                                    If (AbrechnungsVerfahren_2_Aktiv) And (RezArtikel(j%).VdbPauschale = 1) Then
                                        dStückPreis = RezArtikel(j%).Preis
                                    Else
                                        dStückPreis = RezArtikel(j%).Preis * hFaktor%
                                    End If
                                End If
                                HmStückPreis#(j) = dStückPreis
                                HmNummer$(j) = hPosNr   'PznStr
                                hFaktor2% = hFaktor2% * hFaktor%
                                PznStr$ = hPosNr
                            End If
                        End If
                    End If
                End If
            End If
        End If
        
        If (dStückPreis > 0) Then
        ElseIf (RezArtikel(j%).MietDauer > 0) Then
            PznStr$ = "09999063"
'        ElseIf (Left$(PznStr$, 1) = "@") Then
        ElseIf (Left$(PznStr$, 1) = "-") Then
            PznStr$ = RezArtikel(j%).ScreenPzn
'        ElseIf (((ActKasse% = 1) Or (ActKasse% = 27)) And (ActBundesland% = 3)) Then
'            If (Trim(PznStr$) <> "") Then
'                If (AplusVOk) Then
'                    SQLStr$ = "SELECT * FROM VdbArtikel WHERE PZN = " + PznStr
'                    FabsErrf = AplusVDB.OpenRecordset(VdbPznRec, SQLStr)
'                    If Not (VdbPznRec.EOF) Then
'                        gef% = 0
'                        hPosNr$ = Trim(Left$(CheckNullStr(VdbPznRec!HmPositionsNr1), 10))
'                        If (hPosNr$ <> "") Then
'                            PznStr$ = hPosNr$
'
'                            h$ = Trim(CheckNullStr(VdbPznRec!menge))
'                            If (h$ <> "") Then
'                                If (Val(h$) > 0) Then
'                                    hFaktor2% = Val(h$)
'                                    gef% = True
'                                End If
'                            End If
'
'                            If (gef = 0) Then
'                                SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + RezArtikel(j%).pzn
'                                'Set TaxeRec = TaxeDB.OpenRecordset(SQLStr$)
'                                On Error Resume Next
'                                TaxeRec.Close
'                                Err.Clear
'                                On Error GoTo DefErr
'                                TaxeRec.Open SQLStr, taxeAdoDB.ActiveConn
'                                If (TaxeRec.EOF = False) Then
'                                    hFaktor2% = MengeErmitteln(TaxeRec!menge)
'                                    gef% = True
'                                End If
'                            End If
'
'                            If (gef%) Then
'                                hFaktor2% = hFaktor2% * hFaktor%
'                            End If
'                        End If
'                    End If
'                Else
'                    FabsErrf% = VmPzn.IndexSearch(0, PznStr$, FabsRecno&)
'                    If (FabsErrf% = 0) Then
'                        Call VmPzn.GetRecord(FabsRecno&)
'
'                        gef% = 0
'                        hPosNr$ = Trim(Left$(VmPzn.HmPositionsNr1, 10))
'                        If (hPosNr$ <> "") Then
''                            If (Trim(Left$(VmPzn.HmPositionsNr2, 10)) = "") Then
'                                PznStr$ = hPosNr$
''                            End If
'
'                            h$ = Trim(VmPzn.menge)
'                            If (h$ <> "") Then
'                                If (Val(h$) > 0) Then
'                                    hFaktor2% = Val(h$)
'                                    gef% = True
'                                End If
'                            End If
'
'                            If (gef = 0) Then
'                                SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + RezArtikel(j%).pzn
'                                'Set TaxeRec = TaxeDB.OpenRecordset(SQLStr$)
'                                On Error Resume Next
'                                TaxeRec.Close
'                                Err.Clear
'                                On Error GoTo DefErr
'                                TaxeRec.Open SQLStr, taxeAdoDB.ActiveConn
'                                If (TaxeRec.EOF = False) Then
'                                    hFaktor2% = MengeErmitteln(TaxeRec!menge)
'                                    gef% = True
'                                End If
'                            End If
'
'                            If (gef%) Then
'                                hFaktor2% = hFaktor2% * hFaktor%
'                            End If
'                        End If
'                    End If
'                End If
'            End If
        End If
        Call SpeicherDruckZeile(PznStr$)
        
        If (SonderBelegRezept%) Then
            h$ = Format(1, "  0")
        ElseIf (RezArtikel(j%).MietDauer > 0) Then
            h$ = Format(RezArtikel(j%).MietDauer, "  0")
'        ElseIf (hFaktor2% = 1) Then
'            h$ = "   "
        ElseIf (ZusatzKomponentenNr <> "") And (RezArtikel(j).ScreenPzn = ZusatzKomponentenNr) Then
            h$ = Format(ZusatzKompontenFaktor, "  0")
        Else
            h$ = Format(hFaktor2%, "  0")
        End If
        Call SpeicherDruckZeile(h$)
        
        If (Val(RezArtikel(j%).pzn) = 9999057) Or (RezArtikel(j).VdbPauschale) Then
            hFaktor% = 1
        End If
        If (dStückPreis > 0) Then
            lPreis& = CLng(dStückPreis * 100#)
            HmFaktor(j) = hFaktor2
        Else
            lPreis& = CLng(RezArtikel(j%).Preis * hFaktor% * 100#)
        End If
        h$ = Format(lPreis, "0")
        Call SpeicherDruckZeile(h$)
        
    '    SollPreis& = SollPreis& + lPreis&
    Next j%
    
    j = j + AnzMKs
End If

If (BtmRezept%) Then
    Do
        If (j% >= 3) Then Exit Do
        
        For i% = 0 To 2
            Call SpeicherDruckZeile("   ")
        Next i%
        
        j% = j% + 1
    Loop
    If (BtmApoDa) Then
        Call SpeicherDruckZeile("   ")
    Else
        Call SpeicherDruckZeile("@BTM@" + BtmRezDruckName$)
    End If
    Call SpeicherDruckZeile("   ")
    Call SpeicherDruckZeile("   ")
    j% = j% + 1
End If

Do
    If (j% >= 7) Then Exit Do
    
    If (j = 2) And ((Selbsterklaerung) And (SchutzMasken = 0)) Then
        h2 = BoxenWert(52)
        Call SpeicherDruckZeile("@ABDATUM@" + h2$)
    End If
    
    If (j = 3) And (SchutzMasken) Then
        Call SpeicherDruckZeile("@SCHUTZMASKEN@" + "Schutzmasken")
    ElseIf (j = 3) And (BtmRezept% = 0) And (Trim(RezNr) <> "") And (RezeptNrPositionAlt = 0) And (Selbsterklaerung = 0) Then
        h2 = RezNr
        'h2 = "9892002021470"
        If (Left$(h2$, 7) = "9999998") Then h2$ = Mid$(h2$, 8, 5)
        Call SpeicherDruckZeile("@REZNR@" + h2$ + IIf(Druck_HmAbrechnungsKz$ <> "", "  LEGS:" + Druck_HmAbrechnungsKz$, ""))
    End If
    
    For i% = 0 To 2
        Call SpeicherDruckZeile("   ")
    Next i%
    
    j% = j% + 1
Loop

If ((AnzRezeptArtikel% >= 7) Or (lRezNr& = 0&) Or PrivatRezept% Or SonderBelegRezept% Or Selbsterklaerung) Then
    h$ = "   "
Else
    h$ = ""
    If (BenutzerNr% > 0) Then
        h$ = h$ + Format(BenutzerNr%, "00")
    End If
    If (BtmRezept%) Or (RezeptNrPositionAlt) Then
        h$ = h$ + "(" + RezNr$ + ")"
    End If
    If (PersonalNr% > 0) Then
        h$ = h$ + Format(PersonalNr%, "00")
    End If
End If
Call SpeicherDruckZeile(h$)


'If (BtmRezept%) Then
'    h$ = ""
'Else
'    h$ = VerkDatum$
'    If (Noctu) Then
'        h$ = h$ + " " + Format(VerkZeit, "0000")
'    End If
'End If
'Call SpeicherDruckZeile(h$)
'
'If (BtmRezept%) Then
'    h$ = VerkDatum$
'    h$ = Right$(Space$(38) + h$, 38)
'Else
'    h$ = RezApoDruckName$
'    Call CharToOem(h$, h$)
'    h$ = Left$(h$ + Space$(38), 38)
'End If
'Call SpeicherDruckZeile(h$)

If (BtmRezept) Then
    If (BtmRezeptAlt%) Then
        h$ = ""
        Call SpeicherDruckZeile(h$)
        h$ = VerkDatum$
        h$ = Right$(Space$(38) + h$, 38)
        Call SpeicherDruckZeile(h$)
    Else
        h$ = VerkDatum$
        Call SpeicherDruckZeile(h$)
        h = ""
        Call SpeicherDruckZeile(h$)
    End If
Else
    h$ = VerkDatum$
    If (Noctu) Then
        h$ = h$ + " " + Format(VerkZeit, "00:00")
    End If
    If (PrivatRezept) And (DatumObenPrivatRezept) Then
        h = "@PRIVDATUM@" + h
    End If
    Call SpeicherDruckZeile(h$)
    
    If (SchutzMasken) Then
        h = ""
    Else
        h$ = RezApoDruckName$
    End If
    Call CharToOem(h$, h$)
    h$ = Left$(h$ + Space$(38), 38)
    Call SpeicherDruckZeile(h$)
End If


h$ = ""
For i% = 0 To RezDruckInd%
    h$ = h$ + RezDruckZeilen$(i%) + vbCrLf
Next i%
'Call MsgBox(h$)

'IstPreis& = CLng(RezSumme# * 100)
'If (SollPreis& <> IstPreis&) Then
'    Call MsgBox("Achtung: Rezeptdruck u.U. nicht korrekt!", vbInformation)
'End If

Call DefErrPop
End Sub

Function StringAdd$(s1$, s2$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("StringAdd$")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
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
Dim i%, l1%, l2%, iVal%, Übertrag%
Dim ret$

l1 = Len(s1)
l2 = Len(s2)
If (l2 > l1) Then
    l1 = l2
End If
l1 = l1 + 1

s1 = Right(String(l1, "0") + s1, l1)
s2 = Right(String(l1, "0") + s2, l1)

ret$ = ""
Übertrag = 0
For i = l1 To 1 Step -1
    iVal = Val(Mid$(s1, i, 1)) + Val(Mid$(s2, i, 1)) + Übertrag
    ret = CStr(iVal Mod 10) + ret
    Übertrag = iVal \ 10
Next i

Do
    If (Left(ret, 1) = "0") Then
        ret = Mid$(ret, 2)
    Else
        Exit Do
    End If
Loop
If (ret = "") Then
    ret = "0"
End If

StringAdd = ret$

Call DefErrPop
End Function

Sub SpeicherDruckZeile(s$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("SpeicherDruckZeile")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
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

RezDruckZeilen$(RezDruckInd%) = s$
RezDruckInd% = RezDruckInd% + 1

Call DefErrPop
End Sub

Sub FlxBarVerkaufBefuellen()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("FlxBarVerkaufBefuellen")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
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

With frmBarVerkauf.flxBarVerkauf
    For i% = 0 To UBound(BarVerkauf$)
        .TextMatrix(i%, 0) = BarVerkauf$(i%)
    Next i%
End With

Call DefErrPop
End Sub

Sub HoleKundenInfo(KundenNr&)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("HoleKundenInfo")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
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
Dim clsKunden1 As New clsKunden
Dim clsKundenZus1 As New clsKundZusatz
Dim KundenDB As Database
Dim KundenRec As Recordset
Dim i%

Dim KundenSQL%
Dim KundenAdoRec As New ADODB.Recordset
Dim KundenAdoDB As New clsKundenDB

Dim h$, SQLStr$

RezKundenInfo$(4) = Format(KundenNr, "0")
    
If (para.Land = "D") Or (Dir("KunBezug.mdb") <> "") Then
    KundenSQL = KundenAdoDB.DBvorhanden
    If (KundenSQL) Then
        KundenSQL = KundenAdoDB.OpenDB

        SQLStr$ = "SELECT * FROM Kunden WHERE KundenNr =" + Str$(KundenNr)
        On Error Resume Next
        KundenAdoRec.Close
        Err.Clear
        On Error GoTo DefErr
        KundenAdoRec.Open SQLStr, KundenAdoDB.ActiveConn
        If (KundenAdoRec.EOF = False) Then
            h$ = Trim(CheckNullStr(KundenAdoRec!Anrede))
            h$ = Trim(h$ + " " + Trim(CheckNullStr(KundenAdoRec!Titel)))
            h$ = Trim(h$ + " " + Trim(CheckNullStr(KundenAdoRec!Vorname)))
            h$ = Trim(h$ + " " + Trim(CheckNullStr(KundenAdoRec!Name)))
            RezKundenInfo$(0) = h$

            h$ = Trim(CheckNullStr(KundenAdoRec!strasse))
            h$ = Trim(h$ + " " + Trim(CheckNullStr(KundenAdoRec!HausNr)))
            RezKundenInfo$(1) = h$

            h$ = Trim(CheckNullStr(KundenAdoRec!PLZ))
            h$ = Trim(h$ + " " + Trim(CheckNullStr(KundenAdoRec!Ort)))
            RezKundenInfo$(2) = h$

            RezKundenInfo$(3) = Trim(CheckNullStr(KundenAdoRec!Versichertennummer))

            For i% = 0 To 2
                Call CharToOem(RezKundenInfo$(i), RezKundenInfo$(i))
            Next i%

            RezKundenInfo$(5) = Trim(CheckNullStr(KundenAdoRec!KkName))
            h$ = Format(KundenAdoRec!KkNummer, "0")

        End If
        KundenAdoDB.CloseDB
    Else
        h$ = "Kunden.mdb"
        Set KundenDB = OpenDatabase(h$, False, True)
        SQLStr$ = "SELECT * FROM Kunden WHERE KundenNr =" + Str$(KundenNr)
        Set KundenRec = KundenDB.OpenRecordset(SQLStr$)
        If (KundenRec.EOF = False) Then
            h$ = Trim(CheckNullStr(KundenRec!Anrede))
            h$ = Trim(h$ + " " + Trim(CheckNullStr(KundenRec!Titel)))
            h$ = Trim(h$ + " " + Trim(CheckNullStr(KundenRec!Vorname)))
            h$ = Trim(h$ + " " + Trim(CheckNullStr(KundenRec!Name)))
            RezKundenInfo$(0) = h$
            
            h$ = Trim(CheckNullStr(KundenRec!strasse))
            h$ = Trim(h$ + " " + Trim(CheckNullStr(KundenRec!HausNr)))
            RezKundenInfo$(1) = h$
            
            h$ = Trim(CheckNullStr(KundenRec!PLZ))
            h$ = Trim(h$ + " " + Trim(CheckNullStr(KundenRec!Ort)))
            RezKundenInfo$(2) = h$
            
            RezKundenInfo$(3) = Trim(CheckNullStr(KundenRec!Versichertennummer))
            
            For i% = 0 To 2
                Call CharToOem(RezKundenInfo$(i), RezKundenInfo$(i))
            Next i%
            
            RezKundenInfo$(5) = Trim(CheckNullStr(KundenRec!KkName))
            h$ = Format(KundenRec!KkNummer, "0")
        End If
        KundenDB.Close
    End If
Else
    clsKunden1.OpenDatei
    clsKunden1.GetRecord (KundenNr + 1)
    h$ = RTrim$(clsKunden1.Wert("name", 0))
    RezKundenInfo$(0) = h$ + " " + RTrim$(clsKunden1.Wert("name", 1))
    
    RezKundenInfo$(1) = RTrim$(clsKunden1.Wert("name", 2))
    RezKundenInfo$(2) = RTrim$(clsKunden1.Wert("name", 3))
    'RezKundenInfo$(3) = RTrim$(clsKunden1.Wert("name", 3))
    
    clsKunden1.CloseDatei
    Set clsKunden1 = Nothing
    
    clsKundenZus1.OpenDatei
    clsKundenZus1.GetRecord (KundenNr + 1)
    RezKundenInfo$(3) = RTrim$(clsKundenZus1.Vnr)
    
    h$ = Trim(clsKundenZus1.kk)
    clsKundenZus1.CloseDatei
    Set clsKundenZus1 = Nothing
End If

'If (AutIdemSonderregelOk%) And (Val(h$) > 0) Then
'    SQLStr$ = "SELECT * FROM kKassen WHERE IK = " + Str$(Val(h$))
'    Set KkassenRec = kKassenDB.OpenRecordset(SQLStr$)
'    If (KkassenRec.RecordCount > 0) Then
'        AutIdemKbvNr& = CheckNullLong(KkassenRec!KbvNr)
'    End If
'End If

Call DefErrPop
End Sub

Function HoleSonderbelegKundenInfo(KundenNr&, Sonderbeleg As Sonderbeleg_OP) As Boolean
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("HoleSonderbelegKundenInfo as Boolean")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
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
Dim clsKunden1 As New clsKunden
Dim clsKundenZus1 As New clsKundZusatz
Dim KundenDB As Database
Dim KundenRec As Recordset
Dim i%

Dim KundenSQL%
Dim KundenAdoRec As New ADODB.Recordset
Dim KundenAdoDB As New clsKundenDB

Dim bRet As Boolean
Dim h$, SQLStr$

bRet = False

If (para.Land = "D") Or (Dir("KunBezug.mdb") <> "") Then
    With Sonderbeleg
        KundenSQL = KundenAdoDB.DBvorhanden
        If (KundenSQL) Then
            KundenSQL = KundenAdoDB.OpenDB
    
            SQLStr$ = "SELECT * FROM Kunden WHERE KundenNr =" + Str$(KundenNr)
            On Error Resume Next
            KundenAdoRec.Close
            Err.Clear
            On Error GoTo DefErr
            KundenAdoRec.Open SQLStr, KundenAdoDB.ActiveConn
            If (KundenAdoRec.EOF = False) Then
                bRet = True
                
                With .Patient
                    .Id = Trim(CheckNullStr(KundenAdoRec!Versichertennummer))   ' "T555558879"
                    .GeburtsDatum = IIf(IsNull(KundenAdoRec!GeburtsDatum), CDate("31.12.1999"), KundenAdoRec!GeburtsDatum)  'CDate("01.12.1999")
                    With .Name
                        .Vorname = Trim(CheckNullStr(KundenAdoRec!Vorname)) ' "Andreas"
                        .Nachname_ohne_Vor_und_Zusatz = Trim(CheckNullStr(KundenAdoRec!Name))  '"Eder"
                        .Titel = Trim(CheckNullStr(KundenAdoRec!Titel)) ' "Dipl-Ing"
                    End With
                    With .Adresse
                        .PLZ = CheckNullStr(KundenAdoRec!PLZ)   ' "1210"
                        .Ort = Trim(CheckNullStr(KundenAdoRec!Ort)) '"Wien"
                        With .Strasse1
                            .strasse = CheckNullStr(KundenAdoRec!strasse)   ' "Leopoldauer Strasse"
                            .Hausnummer = CheckNullStr(KundenAdoRec!HausNr) ' "68A 3 25"
                        End With
                    End With
                End With
            
                With .Kostentraeger
                    .Typ = "GKV"
                    .IK = Format(KundenAdoRec!KkNummer, "0") ' "101575519"
                    .Name = Trim(CheckNullStr(KundenAdoRec!KkName)) ' "Techniker Krankenkasse"
                End With
            End If
            KundenAdoDB.CloseDB
        Else
            h$ = "Kunden.mdb"
            Set KundenDB = OpenDatabase(h$, False, True)
            SQLStr$ = "SELECT * FROM Kunden WHERE KundenNr =" + Str$(KundenNr)
            Set KundenRec = KundenDB.OpenRecordset(SQLStr$)
            If (KundenRec.EOF = False) Then
                bRet = True
                
                With .Patient
                    .Id = Trim(CheckNullStr(KundenRec!Versichertennummer))   ' "T555558879"
                    .GeburtsDatum = IIf(IsNull(KundenRec!GeburtsDatum), CDate("31.12.1999"), KundenRec!GeburtsDatum)  'CDate("01.12.1999")
                    With .Name
                        .Vorname = Trim(CheckNullStr(KundenRec!Vorname)) ' "Andreas"
                        .Nachname_ohne_Vor_und_Zusatz = Trim(CheckNullStr(KundenRec!Name))  '"Eder"
                        .Titel = Trim(CheckNullStr(KundenRec!Titel)) ' "Dipl-Ing"
                    End With
                    With .Adresse
                        .PLZ = CheckNullStr(KundenRec!PLZ)   ' "1210"
                        .Ort = Trim(CheckNullStr(KundenRec!Ort)) '"Wien"
                        With .Strasse1
                            .strasse = CheckNullStr(KundenRec!strasse)   ' "Leopoldauer Strasse"
                            .Hausnummer = CheckNullStr(KundenRec!HausNr) ' "68A 3 25"
                        End With
                    End With
                End With
            
                With .Kostentraeger
                    .Typ = "GKV"
                    .IK = Format(KundenRec!KkNummer, "0") ' "101575519"
                    .Name = Trim(CheckNullStr(KundenRec!KkName)) ' "Techniker Krankenkasse"
                End With
            End If
            KundenDB.Close
        End If
    End With
'Else
'    clsKunden1.OpenDatei
'    clsKunden1.GetRecord (KundenNr + 1)
'    h$ = RTrim$(clsKunden1.Wert("name", 0))
'    RezKundenInfo$(0) = h$ + " " + RTrim$(clsKunden1.Wert("name", 1))
'
'    RezKundenInfo$(1) = RTrim$(clsKunden1.Wert("name", 2))
'    RezKundenInfo$(2) = RTrim$(clsKunden1.Wert("name", 3))
'    'RezKundenInfo$(3) = RTrim$(clsKunden1.Wert("name", 3))
'
'    clsKunden1.CloseDatei
'    Set clsKunden1 = Nothing
'
'    clsKundenZus1.OpenDatei
'    clsKundenZus1.GetRecord (KundenNr + 1)
'    RezKundenInfo$(3) = RTrim$(clsKundenZus1.Vnr)
'
'    h$ = Trim(clsKundenZus1.kk)
'    clsKundenZus1.CloseDatei
'    Set clsKundenZus1 = Nothing
End If

HoleSonderbelegKundenInfo = bRet

Call DefErrPop
End Function


Sub IndividuelleRezeptur()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("IndividuelleRezeptur")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
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
Dim row%, OrgParenteralRezept%

MagSpeicherIndex% = -1
OrgParenteralRezept = ParenteralRezept
ParenteralRezept = -1
If (ParentForm.picRezept.Visible) Then
    If (AktTabIndex% >= 3) Then
        row% = (AktTabIndex% - 3) \ 4
        If (row% < AnzRezeptArtikel%) Then
            MagSpeicherIndex% = RezArtikel(row%).MagIndex
            ParenteralRezept = OrgParenteralRezept
        Else
            MagSpeicherIndex% = 0
        End If
    End If
End If
        
AutIdemKreuz0 = RezArtikel(0).AutIdemKreuz
frmTaxieren.Show 1

If (MagSpeicherIndex% > 0) Then
    RezArtikel(row%).MagIndex = MagSpeicherIndex%
    
    If (SumPreis# > 0) Then
        If (ParenteralRezept >= 0) And (ParenteralRezept <= 15) Then
            frmParenteralHersteller.Show 1
        End If
        
        With RezArtikel(row%)
            If (ParenteralRezept >= 0) Then
                .pzn = ParenteralPzn(ParenteralRezept)
                .text = ParenteralTxt(ParenteralRezept)
                
                .AutIdemKreuz = IIf(ParenteralRezept = 22, 2, 0)
            ElseIf (UnverarbeiteteAbgabe) Then
                .pzn = "06460702"
                .text = "Rezeptur"
            Else
                .pzn = "09999011"
                .text = "Rezeptur"
            End If
            .Preis = SumPreis#
            
            .Faktor = 1

            .flag = .flag And (REZ_BLINK Xor &HFF)
            .Zuz = CalcZuzaNeu#(SumPreisZuz#)  ' ZuzFeld!(3)
            
            If Not (GebFrei%) And (.Zuz > (SumPreisZuz# / VmRabattFaktor#)) Then
                RezArtikel(row%).Preis = 0#
                RezArtikel(row%).Zuz = 0#
                Call iMsgBox("Artikel unter Zuzahlung, Geldrückgabe ?")
            End If
            
            .AutIdem = 0
            .N0 = 0
            .Imp123 = 0
            .ImpErspart = 0
            .Fam = 0
            .TkkPznDruck = 0
        End With
        
        If (row% >= AnzRezeptArtikel%) Then
            AnzRezeptArtikel% = row% + 1
        End If
        
        Call PaintArtikel(row%)
    End If
End If
       
With ParentForm.flxVerordnung
    .Top = ParentForm.picRezept.Top + KundenBoxY + 60 ' wpara.TitelY
    .Left = ParentForm.picRezept.Left + KundenBoxX + 60 ' wpara.LinksX
End With

Call DefErrPop
End Sub

Public Sub Lennartz()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("Lennartz")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
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
Dim row%
        
frmLennartzAuswahl.Show 1

If (LennartzPzn > 0) Then
    If (SumPreis# > 0) Then
        row% = (AktTabIndex% - 3) \ 4
        With RezArtikel(row%)
            .pzn = PznString(LennartzPzn)
            .text = LennartzTxt
            .Preis = SumPreis#
            
            .Faktor = 1

            .flag = .flag And (REZ_BLINK Xor &HFF)
            .Zuz = CalcZuzaNeu#(SumPreisZuz#)  ' ZuzFeld!(3)
            
            If Not (GebFrei%) And (.Zuz > (SumPreisZuz# / VmRabattFaktor#)) Then
                RezArtikel(row%).Preis = 0#
                RezArtikel(row%).Zuz = 0#
                Call iMsgBox("Artikel unter Zuzahlung, Geldrückgabe ?")
            End If
            
            .AutIdem = 0
            .N0 = 0
            .Imp123 = 0
            .ImpErspart = 0
            .Fam = 0
            .TkkPznDruck = 0
        End With
        
        If (row% >= AnzRezeptArtikel%) Then
            AnzRezeptArtikel% = row% + 1
        End If
        
        Call PaintArtikel(row%)
    End If
End If

Call DefErrPop
End Sub

Sub WriteRezeptSpeicher(Optional PrivRezVK% = 0)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("WriteRezeptSpeicher")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
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
Dim i%, j%, k%, erg%, hFaktor%, VkPrivCheck%, AvpMonat%
Dim RezGesamt#, RezGesamtFAM#, RezImpFähig#, RezImpIst#
Dim Quote#, ImpSoll#, diff#, SaldoV#, RezeptRabatt#(1), ArtikelRabatt#(1), ePreis#, nPreis#
Dim h$, h2$, pzn$, AuswertungMonat$, Unique$, SortDatum$, SQLStr$
Dim hImp123 As Byte
Dim aktvk As Double, OrgVk As Double
Dim SuchPzn As String
Dim RezImpErspart As Double
Dim AvpRec As Recordset

Dim Jahr As Boolean, AbrechVorMon As Boolean
Dim bmk As Variant
Dim w#(4)
Dim FieldNr%(8)
 
    
If (BereitsGedruckt%) Then Call DefErrPop: Exit Sub
If (RezSpeicherOK%) Then
    BereitsGedruckt% = True
    
    If (MarsModus = MARS_REZEPT_KONTROLLE) Then
        Set RezepteRec = MarsRezSpeicherDB.OpenRecordset("Rezepte", dbOpenTable)
        Set ArtikelRec = MarsRezSpeicherDB.OpenRecordset("Artikel", dbOpenTable)
        Set AuswertungRec = MarsRezSpeicherDB.OpenRecordset("Auswertung", dbOpenTable)
    Else
        Set RezepteRec = RezSpeicherDB.OpenRecordset("Rezepte", dbOpenTable)
        Set ArtikelRec = RezSpeicherDB.OpenRecordset("Artikel", dbOpenTable)
        Set AuswertungRec = RezSpeicherDB.OpenRecordset("Auswertung", dbOpenTable)
    End If

    If (SchonDa%) Then
        SchonDa% = False
        With RezepteRec
            .index = "RezeptNr"
            .Seek ">=", RezNr$
            If (.NoMatch = False) Then
                If (RezepteRec!RezeptNr = RezNr$) Then
                    SchonDa% = True
                End If
            End If
        End With
    End If

    If (SchonDa%) Then
        Unique$ = RezepteRec!Unique
    
        If (NurHashCodeDruck) Then
            RezepteRec.Edit
            RezepteRec!DruckDatum = Format(Now, "YYMMDD")
            RezepteRec!DruckZeit = Val(Format(Now, "HHMM"))
            If (ParenteralRezept >= 0) Or (ParenteralRezept = -99) Or (LennartzPzn > 0) Then
                RezepteRec!ParenteralHash = ParenteralHash
            End If
            
            RezepteRec!rzLieferId = ""
            RezepteRec!SendeStatus = 0
            RezepteRec!AbrechnungsStatus = 0
            RezepteRec.Update
            
            If (MAG_SPEICHER% > 0) Then Close (MAG_SPEICHER%)
            Unload frmTaxieren
            
            Call DefErrPop: Exit Sub
        End If
        
        With ArtikelRec
            .index = "Unique"
            For j% = 0 To (RezepteRec!AnzArtikel - 1)
                .Seek "=", Unique$ + Format(j% + 1, "0")
                If (.NoMatch = False) Then
                    .Delete
                End If
            Next j%
        End With
    
        RezepteRec.Delete
    End If
    
    If (Val(RezArtikel(0).pzn) = 2567024) Then
        For i% = 0 To UBound(RezArtikel) - 1
            RezArtikel(i%) = RezArtikel(i% + 1)
        Next i%
        AnzRezeptArtikel% = AnzRezeptArtikel% - 1
    End If
    
    SortDatum$ = Right$(VerkDatum$, 2) + Mid$(VerkDatum$, 3, 2) + Left$(VerkDatum$, 2)
    
    If PrivRezVK% <> 0 Then
        With ArtikelRec
            .index = "Unique"
            j% = 0
            Do
                VkPrivCheck% = True
                Unique$ = SortDatum$ + Format(VerkZeit%, "0000") + Format(j%, "00") + Right$(Space$(2) + VerkUser$, 2)
                For i% = 0 To (AnzRezeptArtikel% - 1)
                    .Seek "=", Unique$ + Format(i% + 1, "0")
                    If (.NoMatch) Then
                        VkPrivCheck% = False
                        Exit For
                    End If
                    If (ArtikelRec!pzn <> Val(RezArtikel(i%).pzn)) Then
                        VkPrivCheck% = False
                        Exit For
                    End If
                Next i%
                If (VkPrivCheck%) Then
                    .Seek "=", Unique$ + Format(AnzRezeptArtikel% + 1, "0")
                    If (.NoMatch = False) Then
                        VkPrivCheck% = False
                    End If
                End If
                If (VkPrivCheck%) Then
                    For i% = 0 To (AnzRezeptArtikel% - 1)
                        .Seek "=", Unique$ + Format(i% + 1, "0")
                        .Delete
                    Next i%

                    With RezepteRec
                        .index = "Unique"
                        .Seek "=", Unique$
                        If (.NoMatch = False) Then
                            If (RezepteRec!Kkasse = "Privat VK") And (RezepteRec!AnzArtikel = AnzRezeptArtikel%) Then
                                .Delete
                            End If
                        End If
                    End With

                End If

                j% = j% + 1
                If (j% >= 100) Then Exit Do
            Loop
        End With
        
        
        With ArtikelRec
            .index = "Unique"
            j% = 0
            Do
                Unique$ = SortDatum$ + Format(VerkZeit%, "0000") + Format(j%, "00") + Right$(Space$(2) + VerkUser$, 2)
                .Seek "=", Unique$ + "1"
                If (.NoMatch) Then Exit Do
                j% = j% + 1
            Loop
        End With
        With RezepteRec
            .index = "Unique"
            Do
                Unique$ = SortDatum$ + Format(VerkZeit%, "0000") + Format(j%, "00") + Right$(Space$(2) + VerkUser$, 2)
                .Seek "=", Unique$
                If (.NoMatch) Then Exit Do
                j% = j% + 1
            Loop
        End With
    Else
        Unique$ = Format(Now, "YYMMDDHHNNSS") + Right$(Space$(2) + para.User, 2)
    End If
    
    RezeptRabatt#(0) = 0
    RezeptRabatt#(1) = 0
    
    For j% = 0 To (AnzRezeptArtikel% - 1)
        ArtikelRec.AddNew
        
        h$ = Unique$ + Format(j% + 1, "0")
        If j% >= 9 And j% <= 243 Then
'            h$ = Unique$ + Chr$(j% + 12)    'für Privatrezepte
            h$ = Unique$ + Chr$(j% + 65)   'für Privatrezepte
        ElseIf j% > 243 Then Exit For      'kann's nicht geben, nur zur Sicherheit
        End If
        ArtikelRec!Unique = h$
        
        pzn$ = "       "
        If (Val(RezArtikel(j%).pzn) <> 9999999) Then
            pzn$ = RezArtikel(j%).pzn
        End If
        ArtikelRec!pzn = Val(RezArtikel(j%).pzn)
    
        ArtikelRec!text = RezArtikel(j%).text
        ArtikelRec!flag = RezArtikel(j%).flag
        ArtikelRec!Zuz = RezArtikel(j%).Zuz
        ArtikelRec!HerstRabattPrivat130Brutto = RezArtikel(j%).HerstRabattPrivat130Brutto
        ArtikelRec!TkkPznDruck = RezArtikel(j%).TkkPznDruck
        ArtikelRec!AutIdemKreuz = RezArtikel(j%).AutIdemKreuz
        ArtikelRec!Imp = RezArtikel(j%).Imp123
        ArtikelRec!ImpErspart = RezArtikel(j%).ImpErspart
        
        hFaktor% = RezArtikel(j%).Faktor
        'If (RezArtikel(j%).pzn = "9999057") Then
        '    hFaktor% = 0
        'End If
        ArtikelRec!Faktor = RezArtikel(j%).Faktor
    
        ePreis# = RezArtikel(j%).Preis
        ArtikelRec!ePreis = RezArtikel(j%).Preis
        RezArtikel(j%).Preis = RezArtikel(j%).Preis * hFaktor%
        ArtikelRec!Preis = RezArtikel(j%).Preis
        
        ArtikelRec!ScreenPzn = RezArtikel(j%).ScreenPzn
        
        ArtikelRec!Warenzeichen = RezArtikel(j%).Warenzeichen
        
        If (PrivRezVK% = 0) And (HmNummer$(j) <> "") Then
            ArtikelRec!HmNummer = HmNummer(j)
            ArtikelRec!HmFaktor = HmFaktor(j)
            ArtikelRec!HmStückPreis = HmStückPreis(j)
        End If
        
        ArtikelRec!Fam = RezArtikel(j).Fam
        ArtikelRec!AutIdem = RezArtikel(j).AutIdem
        ArtikelRec!N0 = RezArtikel(j).N0
        ArtikelRec!MagIndex = RezArtikel(j).MagIndex
        ArtikelRec!appli = RezArtikel(j).appli
        
        ArtikelRec!zusatz1 = Left(RezArtikel(j).zusatz(0), 36)
'        ArtikelRec!zusatz2 = RezArtikel(j).zusatz(1)
        h = Trim(RezArtikel(j).zusatz(1))
        If (Len(h) > 36) Then
            ArtikelRec!Zusatz2 = Left(h, 36)
            ArtikelRec!Zusatz2a = Mid(h, 37)
        Else
            ArtikelRec!Zusatz2 = Left(RezArtikel(j).zusatz(1), 36)
        End If
        
        If (AuseinzelungPzn(0) <> "") Then
            k = 1
            If (AuseinzelungPzn(k) <> "") Then
                If (Len(AuseinzelungPzn(k)) = 8) Then
                    h$ = AuseinzelungPzn(k) + "11" + HashFaktor(CDbl(AuseinzelungFaktor(k))) + "11" + HashPreis(AuseinzelungPreis(k) * 100)
                Else
                    h$ = "0" + AuseinzelungPzn(k) + "11" + HashFaktor(CDbl(AuseinzelungFaktor(k))) + "11" + HashPreis(AuseinzelungPreis(k) * 100)
                End If
            
'                ArtikelRec!zusatz3 = h
                If (Len(h) > 36) Then
                    ArtikelRec!zusatz3 = Left(h, 36)
                    ArtikelRec!Zusatz3a = Mid(h, 37)
                Else
                    ArtikelRec!zusatz3 = h
                End If
            End If
        End If
        
        
        ArtikelRec!NichtInTaxe = RezArtikel(j).NichtInTaxe
        ArtikelRec!IstWg4 = RezArtikel(j).IstWg4
        ArtikelRec!PlusMehrKosten = RezArtikel(j).PlusMehrKosten
        ArtikelRec!ZuzahlungsErlass = RezArtikel(j).ZuzahlungsErlass
        ArtikelRec!MietDauer = RezArtikel(j).MietDauer
        ArtikelRec!VdbPauschale = RezArtikel(j).VdbPauschale
        
        ArtikelRec!VebNr = RezArtikel(j).VebNr
        ArtikelRec!PauschaleNr = RezArtikel(j).PauschaleNr
        ArtikelRec!Fortsetzung = RezArtikel(j).Fortsetzung
        ArtikelRec!HmAbrechnungsKz = RezArtikel(j).HmAbrechnungsKz
        ArtikelRec!LEGS = RezArtikel(j).LEGS
        
        ArtikelRec.Update
        
        RezGesamt# = RezGesamt# + RezArtikel(j%).Preis
        If (RezArtikel(j%).Fam > 0) Then
            RezGesamtFAM# = RezGesamtFAM# + RezArtikel(j%).Preis
        End If
        
        ArtikelRabatt#(0) = 0
        ArtikelRabatt#(1) = 0
        
'        If (Left$(RezArtikel(j%).pzn, 1) <> "@") Then
        If (Left$(RezArtikel(j%).pzn, 1) <> "-") Then
            SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + Str$(Val(RezArtikel(j%).pzn))
            'Set TaxeRec = TaxeDB.OpenRecordset(SQLStr$)
            On Error Resume Next
            TaxeRec.Close
            Err.Clear
            On Error GoTo DefErr
            TaxeRec.Open SQLStr, taxeAdoDB.ActiveConn
            If Not (TaxeRec.EOF) Then
'                If (TaxeRec!ApoRabattKz) Then
'                    'Neuer Abrechnungspreis
'                    Select Case Abs(ePreis#)
'                        Case 0# To 52.46
'                            nPreis# = ePreis# * 0.94
'                        Case 52.47 To 54.8
'                            nPreis# = 49.32
'                        Case 54.81 To 820.22
'                            nPreis# = ePreis# * 0.9
'                        Case Else
'                            diff# = Sgn(ePreis#) * (Abs(ePreis#) - 820.22)
'                            nPreis# = ePreis# * 0.9 - (diff# * 0.06)
'                    End Select
'                    ArtikelRabatt# = ArtikelRabatt# + (ePreis# - nPreis#)
'
'                End If
                If (TaxeRec!ApoRabattKz = 1) Then
                    ArtikelRabatt#(0) = ArtikelRabatt#(0) + 2.05
                ElseIf (TaxeRec!ApoRabattKz = 2) Then
                    ArtikelRabatt#(0) = ArtikelRabatt#(0) + (ePreis# * 0.05)
                End If
                
                If (TaxeRec!HerstRabattKZ) Then
                    ArtikelRabatt#(1) = ArtikelRabatt#(1) + (TaxeRec!HerstRabatt / 100#)
                End If
'                If (TaxeRec!GhRabattKz) Then
'                    ArtikelRabatt#(1) = ArtikelRabatt#(1) + (TaxeRec!GhRabatt / 100#)
'                End If
            End If
        End If
        
        For k = 0 To 1
            RezeptRabatt#(k) = RezeptRabatt#(k) + (ArtikelRabatt#(k) * hFaktor%)
        Next k
        
        hImp123 = RezArtikel(j%).Imp123
        If (hImp123 > 0) Then
            If Not (TaxeRec.EOF) Then
                If (TaxeRec!ArtStatus = "S") Then hImp123 = 0
            End If
        End If
        
        If (hImp123 > 0) Then
'            If (hImp123 = 1) Then
                hImp123 = 0
                OrgVk = 0
                SuchPzn = Str$(RezArtikel(j%).pzn)
                SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + Str$(RezArtikel(j%).pzn)
                'Set TaxeRec = TaxeDB.OpenRecordset(SQLStr$)
                On Error Resume Next
                TaxeRec.Close
                Err.Clear
                On Error GoTo DefErr
                TaxeRec.Open SQLStr, taxeAdoDB.ActiveConn
                If Not TaxeRec.EOF Then
                  SuchPzn = CStr(TaxeRec!OriginalPZN)
                  SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + CStr(TaxeRec!OriginalPZN)
                    'Set TaxeRec = TaxeDB.OpenRecordset(SQLStr$)
                    On Error Resume Next
                    TaxeRec.Close
                    Err.Clear
                    On Error GoTo DefErr
                    TaxeRec.Open SQLStr, taxeAdoDB.ActiveConn
                  If Not TaxeRec.EOF Then
                    OrgVk = TaxeRec!vk / 100#
                  End If
                End If
                
                SQLStr$ = "SELECT * FROM TAXE WHERE ORIGINALPZN = " + SuchPzn
                'Set TaxeRec = TaxeDB.OpenRecordset(SQLStr$)
                On Error Resume Next
                TaxeRec.Close
                Err.Clear
                On Error GoTo DefErr
                TaxeRec.Open SQLStr, taxeAdoDB.ActiveConn
                Do
                    If (TaxeRec.EOF) Then Exit Do
                    If (TaxeRec!ArtStatus <> "S") Then
                      aktvk = TaxeRec!vk / 100#
                      If aktvk > 0 And OrgVk > 0 And (aktvk <= (OrgVk * 0.85) Or aktvk <= (OrgVk - 15#)) Then
                        If TaxeRec!pzn = RezArtikel(j%).pzn Then
                          hImp123 = 2
                          If RezArtikel(j%).ImpErspart = 0 Then RezArtikel(j%).ImpErspart = OrgVk - aktvk
                          Exit Do
                        Else
                          hImp123 = 1
                        End If
                      End If
                    End If
                    TaxeRec.MoveNext
                Loop
 '           End If
            
            If (hImp123 > 0) Then RezImpFähig# = RezImpFähig# + RezArtikel(j%).Preis
            
            If (hImp123 = 2) Then
                RezImpIst# = RezImpIst# + RezArtikel(j%).Preis
                RezImpErspart = RezImpErspart + RezArtikel(j%).ImpErspart
            End If
        End If
        
        RezArtikel(j%).Preis = RezArtikel(j%).Preis / hFaktor%
    Next j%
    
    If ((ParenteralRezept >= 0) Or (ParenteralRezept = -99)) And (MagSpeicherIndex > 0) Then
        MAG_SPEICHER% = FileOpen("MagTax" + Trim(para.User), "RW", "B")
        If (MAG_SPEICHER% > 0) Then
            Set ParenteralTaxierungRec = RezSpeicherDB.OpenRecordset("ParenteralTaxierungen", dbOpenTable)

            Seek #MAG_SPEICHER%, ((MagSpeicherIndex% - 1) * CLng(Len(TaxierRec))) + 1
            For i% = 1 To 100
                Get #MAG_SPEICHER%, , TaxierRec
                If (EOF(MAG_SPEICHER%)) Then Exit For
                If (TaxierRec.flag = 255) Then Exit For
            
                ParenteralTaxierungRec.AddNew
                
                ParenteralTaxierungRec!RezepteUnique = Unique
                ParenteralTaxierungRec!AnfMagInd = i%
                
                ParenteralTaxierungRec!ActPreis = TaxierRec.ActPreis
                ParenteralTaxierungRec!ActMenge = TaxierRec.ActMenge
                ParenteralTaxierungRec!einheit = TaxierRec.Meh
                ParenteralTaxierungRec!text = TaxierRec.kurz
                ParenteralTaxierungRec!pzn = Val(TaxierRec.pzn)
                ParenteralTaxierungRec!flag = TaxierRec.flag
                ParenteralTaxierungRec!kp = TaxierRec.kp
                ParenteralTaxierungRec!GStufe = TaxierRec.GStufe
                ParenteralTaxierungRec!Verwurf = TaxierRec.Verwurf
                
                If (Parenteral_AOK_LosGebiet) Then
                    ParenteralTaxierungRec!menge = "1"
                ElseIf (Parenteral_AOK_NordOst) Then
                    ParenteralTaxierungRec!menge = "2"
                End If
                
                ParenteralTaxierungRec.Update
            Next i%
            
            Close (MAG_SPEICHER%)
        End If
    ElseIf (LennartzPzn > 0) Then
        Call LennartzRezeptur(1, Unique)
    End If
    
    
    If (PrivatRezept%) Then
        If PrivRezVK% = 0 Then
            ActKkasse$ = "Privat   "
        Else
            ActKkasse$ = "Privat VK"
        End If
    ElseIf (ActKkasse$ = "") Then
        If (ActKasse% > 0) Then
            ActKkasse$ = Format(ActKasse%, String(9, "0"))
        Else
            ActKkasse$ = "Unbekannt"
        End If
    End If
    
    RezepteRec.AddNew
    
    
    If (RezNr$ = "") Then
        RezNr$ = " "
    End If
    
    If (ActKkasse$ = "") Then ActKkasse$ = " "
    If (VerkUser$ = "") Then VerkUser$ = " "
    
    RezepteRec!Unique = Unique$
    RezepteRec!RezeptNr = RezNr$
    RezepteRec!Kkasse = ActKkasse$
    RezepteRec!GebFrei = GebFrei%
    RezepteRec!RezGebSumme = RezGebSumme#
    RezepteRec!RezSumme = RezSumme#
    RezepteRec!AnzArtikel = AnzRezeptArtikel%
    
    If (IvfRezept%) Then
        RezepteRec!Bundesland = ActBundesland% + 1000
    Else
        RezepteRec!Bundesland = ActBundesland%
    End If
    
    RezepteRec!Kasse = ActKasse%
    RezepteRec!Verordnung = ActVerordnung%
    RezepteRec!VerkDatum = SortDatum$
    RezepteRec!VerkZeit = VerkZeit%
    RezepteRec!VerkComp = Right$(Space$(2) + VerkUser$, 2)
    RezepteRec!VerkLiRe = VerkLiRe
    
    RezepteRec!PersonalNr = PersonalNr%
    RezepteRec!DruckDatum = Left$(Unique$, 6)
    RezepteRec!InstitutsKz = RezApoNr$
    RezepteRec!Fam = RezGesamtFAM#
    RezepteRec!ImpFähig = RezImpFähig#
    RezepteRec!ImpIst = RezImpIst#
    RezepteRec!DruckZeit = Mid$(Unique$, 7, 4)
    RezepteRec!RabattWert = RezeptRabatt#(0)
    RezepteRec!HerstRabatt = RezeptRabatt#(1)
    RezepteRec!ImpErspart = RezImpErspart
    RezepteRec!knr = Val(RezKundenInfo$(4))

    If (AvpTeilnahme%) Then
'        h$ = Mid$(AvpRezeptNr$, 7, 4)
'        h$ = Mid$(h$, 3, 2) + Left$(h$, 2)
'        RezepteRec!AvpMonat = Val(h$)
        RezepteRec!AvpRezeptNr = AvpRezeptNr$
        RezepteRec!AvpLaufNr = AvpLaufNr&
    End If
    If (FiveRxFlag) Then
        RezepteRec!TransaktionsID = TransaktionsID
        If ((ParenteralRezept >= 0) Or (ParenteralRezept = -99) Or (LennartzPzn > 0)) Or (AuseinzelungPzn(0) <> "") Then
            RezepteRec!ParenteralHash = ParenteralHash
            RezepteRec!pCharge = pCharge
            RezepteRec!HashErstellungsDatum = Format(HashErstellDat, "YYYYMMDD:HHMM")
'        ElseIf (AuseinzelungPzn <> "") Then
'            RezepteRec!ParenteralHash = ParenteralHash
'            RezepteRec!pCharge = "Auseinzelung"
        End If
        If (LennartzPzn > 0) Then
            RezepteRec!AvpLaufNr = -99
        End If
    End If
    If (Code128Flag) Then
        RezepteRec!TransaktionsID = TransaktionsID
    End If
    
    RezepteRec!Verfügbarkeit = CheckVerfügbarkeit
    
'    If (AutIdemIk& > 0) Then
'        RezepteRec!kKassenIk = AutIdemIk&
'    ElseIf (ActAVWGik& > 0) Then
'        RezepteRec!kKassenIk = ActAVWGik&
'    End If
    RezepteRec!kKassenIk = AutIdemIk&
    RezepteRec!AutIdemKbvNr = AutIdemKbvNr
    RezepteRec!AVWGik = ActAVWGik&
    
    RezepteRec!rzLieferId = rzLieferId
    RezepteRec!SendeStatus = 0
    RezepteRec!AbrechnungsStatus = 0
    
    RezepteRec!P302 = Abs(RezeptMitHilfsmittelNr% > 0)
    
    RezepteRec!VebNr = ActVebNr
    RezepteRec!PauschaleNr = ActPauschaleNr
    
    If (MarsModus = MARS_REZEPT_KONTROLLE) Then
        RezepteRec!DruckPersonalNr = ActBenutzer
    End If
    
    RezepteRec.Update
    
    'neu ab 5.4.01
    RezeptUnique = Unique
    
    If (SchonDa%) Then Call DefErrPop: Exit Sub
    
    diff# = 0#
    
    'AuswertungMonat$ = Left$(SortDatum$, 4)
    AuswertungMonat$ = AbrechMonat$
    If Left$(SortDatum$, 4) < AbrechMonat$ Then
        With AbrechDatenRec
            .MoveFirst
            Do While Not .EOF
                If AbrechDatenRec!Unique = Val(Mid(SortDatum, 3, 2)) Then
                    If (PrivRezVK% <> 0) Or (DateValue(Left(AbrechDatenRec!Datum, 2) + "." + Mid(AbrechDatenRec!Datum, 3, 2) + "." + Mid(AbrechDatenRec!Datum, 5, 2)) >= DateValue(Now)) Then
                        AuswertungMonat$ = Left$(SortDatum$, 4)
                    End If
                    Exit Do
                End If
                .MoveNext
            Loop
        End With
    End If
    With AuswertungRec
        .index = "Unique"
        .Seek "=", ActKkasse$, AuswertungMonat$
        If (.NoMatch) Then
            .AddNew
            AuswertungRec!Kkasse = ActKkasse$
            AuswertungRec!Monat = AuswertungMonat$
            AuswertungRec!Rez_Gesamt = 0
            AuswertungRec!Rez_GesamtFAM = 0
            AuswertungRec!Rez_ImpFähig = 0
            AuswertungRec!Rez_ImpIst = 0
        Else
            .Edit
        End If
        
        AuswertungRec!Rez_Gesamt = dCheckNull(AuswertungRec!Rez_Gesamt) + RezGesamt#
        AuswertungRec!Rez_GesamtFAM = dCheckNull(AuswertungRec!Rez_GesamtFAM) + RezGesamtFAM#
        AuswertungRec!Rez_ImpFähig = dCheckNull(AuswertungRec!Rez_ImpFähig) + RezImpFähig#
        AuswertungRec!Rez_ImpIst = dCheckNull(AuswertungRec!Rez_ImpIst) + RezImpIst#
        AuswertungRec!RezAnzahl = dCheckNull(AuswertungRec!RezAnzahl) + 1
        AuswertungRec!ImpErspart = dCheckNull(AuswertungRec!ImpErspart) + RezImpErspart
            
        Quote# = ImportQuote#(AuswertungRec!Rez_GesamtFAM, AuswertungRec!Rez_ImpFähig, Left$(AuswertungRec!Monat, 4))
        ImpSoll# = AuswertungRec!Rez_GesamtFAM * (Quote# / 100#)
        diff# = AuswertungRec!Rez_ImpIst - ImpSoll#
        
        SaldoV# = dCheckNull(AuswertungRec!Saldo)
        
        .Update
    
        If (SaldoV = 0#) Then
            If Val(Mid(AuswertungMonat$, 3, 2)) = 1 Then
                .Seek "<=", ActKkasse$, Format(Val(Left(AuswertungMonat$, 2)) - 1, "00") + "12"
            Else
                .Seek "<=", ActKkasse$, Left(AuswertungMonat$, 2) + Format(Val(Mid(AuswertungMonat$, 3, 2)) - 1, "00")
            End If
            If Not .NoMatch Then
                If (AuswertungRec!Kkasse = ActKkasse$) Then
                    SaldoV# = dCheckNull(AuswertungRec!Saldo)
                    If (SaldoV# = 0) Then GoSub SaldoVorMonRechnen
                End If
            End If
        End If
        If (SaldoV# > 0) Then diff# = diff# + SaldoV#
    
    End With
    
    With KasseRec
        .index = "Unique"
        .Seek "=", ActKkasse$
        If (.NoMatch) Then
        Else
            .Edit
            KasseRec!importkz = (diff# < 0)
            .Update
        End If
    End With
    
    If (AutIdemKbvNr& > 0) Then
        Set AutIdemKkRec = RezSpeicherDB.OpenRecordset("AutIdemKrankenkassen", dbOpenTable)
        With AutIdemKkRec
            .index = "Ik"
            .Seek "=", AutIdemIk&
            If (.NoMatch) Then
                .AddNew
                AutIdemKkRec!IK = AutIdemIk&
                AutIdemKkRec!KbvNr = AutIdemKbvNr&
                AutIdemKkRec!AnzRezepte = 0
            
                h$ = " "
                SQLStr$ = "SELECT * FROM kKassen WHERE IK = " + Str$(AutIdemIk&)
'                Set KkassenRec = kKassenDB.OpenRecordset(SQLStr$)
'                If (KkassenRec.RecordCount > 0) Then
                FabsErrf = kKassenDB.OpenRecordset(KkassenRec, SQLStr)
                If Not (KkassenRec.EOF) Then
                    h$ = KkassenRec!Name
                End If
                AutIdemKkRec!Name = h$
            Else
                .Edit
            End If
            AutIdemKkRec!AnzRezepte = AutIdemKkRec!AnzRezepte + 1
            .Update
        End With
    End If
End If
    
Call WriteRezkontrDat
Call DefErrPop: Exit Sub


SaldoVorMonRechnen:
With AuswertungRec
    FieldNr%(1) = .Fields("rez_gesamt").OrdinalPosition
    FieldNr%(2) = .Fields("abr_gesamt").OrdinalPosition
    FieldNr%(3) = .Fields("rez_gesamtFAM").OrdinalPosition
    FieldNr%(4) = .Fields("abr_gesamtFAM").OrdinalPosition
    FieldNr%(5) = .Fields("rez_ImpFähig").OrdinalPosition
    FieldNr%(6) = .Fields("abr_ImpFähig").OrdinalPosition
    FieldNr%(7) = .Fields("rez_ImpIst").OrdinalPosition
    FieldNr%(8) = .Fields("abr_ImpIst").OrdinalPosition
    
    Do
      AbrechVorMon = False
      For i% = 1 To 4
          'fieldnr(1) = rez_gesamt, 3 = rez_gesamtFAM, 5 = rez_ImpFähig, 7 = rez_ImpIst
          If (dCheckNull(AuswertungRec.Fields(FieldNr%(2 * i%))) = 0#) Then
              w#(i%) = w#(i%) + dCheckNull(AuswertungRec.Fields(FieldNr%(2 * i% - 1)))
          Else
              w#(i%) = w#(i%) + AuswertungRec.Fields(FieldNr%(2 * i%))
              If i% = 1 Then AbrechVorMon = True
          End If
      Next i%
      
      If dCheckNull(AuswertungRec!GutHaben) = 0# And Not AbrechVorMon Then
          SaldoV# = (w#(W_IMPIST) - w#(W_FAM) * (ImportQuote#(w#(W_FAM), w#(W_IMPFÄHIG), Left$(AuswertungRec!Monat, 4)) / 100#)) / 10#
      End If
      AuswertungRec.MovePrevious
      If .BOF Then Exit Do
    Loop Until AbrechVorMon Or SaldoV# <> 0 Or AuswertungRec!Kkasse <> ActKkasse$
End With
Return

Call DefErrPop
End Sub


Sub PaintAnzuzeigendeKkassen()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("PaintAnzuzeigendeKkassen")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
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
Dim i%, row%, MaxHeight%, iFontSize%, MaxWi%, wi%
Dim OrgFontSize%
Dim h$, h2$, OrgFontName$
Dim Kkasse As New clsKkassen

RoteKassenDa% = False
With ParentForm.flxKkassen
    .Visible = False
    If (EasyImpFlag%) Then
        .Rows = 0
        
        KasseRec.index = "Name"
        KasseRec.Seek ">=", Space(30)
        Do
            If (KasseRec.NoMatch) Or (KasseRec.EOF) Then
                Exit Do
            ElseIf (KasseRec!Anzeige) Then
'                h$ = KasseRec!Nummer + vbTab + KasseRec!Name + vbTab + Format(KasseRec!Typ, "0") + vbTab
                h$ = KasseRec!Nummer + vbTab + KasseRec!Name + vbTab + Format(KasseRec!VebNr, "0") + vbTab
                If (KasseRec!importkz) Then
                    h2$ = "0"
                Else
                    h2$ = "1"
                End If
                h2$ = h2$ + KasseRec!Name
                h$ = h$ + h2$
                
'                If (KasseRec!importkz) Then h$ = h$ + "*"
                .AddItem h$
            End If
            KasseRec.MoveNext
        Loop
            
        
        
        If (.Rows = 0) Then
            .Visible = False
        Else
            MaxHeight% = ParentForm.ArbeitBackHeight - .Top - 15 - 210
            .Font.Name = wpara.FontName(0)
            .Font.Size = wpara.FontSize(0)
            Do
                If ((.Rows * .RowHeight(0)) > MaxHeight%) Then
                    iFontSize% = .Font.Size - 1
                    If (iFontSize% <= 9) Then
                        .Height = (MaxHeight% \ .RowHeight(0)) * .RowHeight(0) '+ 90
                        .ScrollBars = flexScrollBarVertical
                        Exit Do
                    End If
'                    If (iFontSize% <= 8) Then
'                        .Font.Name = "Small Fonts"
'                    End If
                    .Font.Size = iFontSize%
                Else
                    .Height = (.Rows * .RowHeight(0)) '+ 90
                    Exit Do
                End If
            Loop
            
'            OrgFontName$ = ParentForm.Font.Name
'            OrgFontSize% = ParentForm.Font.Size
'            ParentForm.Font.Name = .Font.Name
'            ParentForm.Font.Size = .Font.Size
'            MaxWi% = 0
'            For i% = 0 To (.Rows - 1)
'                wi% = ParentForm.TextWidth(.TextMatrix(i%, 1)) '+ 90
'                If (wi% > MaxWi%) Then MaxWi% = wi%
'            Next i%
'            If (MaxWi% < ParentForm.txtRezeptNr.Width) Then
'                MaxWi% = ParentForm.txtRezeptNr.Width
'            End If
'            .ColWidth(1) = MaxWi%
'            .Width = MaxWi%
'            ParentForm.Font.Name = OrgFontName$
'            ParentForm.Font.Size = OrgFontSize%
            
            .row = 0
            .col = 3
            .RowSel = .Rows - 1
            .ColSel = .col
            .Sort = 5
            .col = 0
            .ColSel = .Cols - 1
            
            .ColAlignment(1) = flexAlignLeftCenter
           
            row% = 0
            For i% = 0 To (.Rows - 1)
                .row = i%
                .col = 1
                If (Left$(.TextMatrix(i%, 3), 1) = "0") Then
                    .CellForeColor = vbRed
                    RoteKassenDa% = True
                End If
                If (.TextMatrix(i%, 0) = ActKkasse$) Then
                    .CellFontBold = True
                    row% = i%
                Else
                    .CellFontBold = False
                End If
            Next i%
            .row = row%
            .col = 0
            .ColSel = .Cols - 1
            
            .Visible = True
        End If
    End If
End With
   
Call DefErrPop
End Sub

Sub flxKkassenClick()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("flxKkassenClick")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
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
Dim i%, row%, iTabIndex%

If (Chef = False) Then
    Call DefErrPop: Exit Sub
End If
            
With ParentForm.flxKkassen
    .Redraw = False
    row% = .row
    ActKkasse$ = .TextMatrix(.row, 0)
'    ActKasse% = Val(.TextMatrix(.row, 2))
    ActVebNr = Val(.TextMatrix(.row, 2))
    ActPauschaleNr = 0
        
    If (ActVebNr = 0) Then
        frmAvAuswahl.Show 1
        Call PaintAvParameter
        If (ActVebNr > 0) Then
            SQLStr = "Update Kassen SET VebNr=" + CStr(ActVebNr)
            SQLStr = SQLStr + " WHERE Nummer='" + ActKkasse + "'"
            KassenDB.Execute (SQLStr$)
        End If
    End If

        
    For i% = 0 To (.Rows - 1)
        .row = i%
        .col = 1
        If (.TextMatrix(i%, 0) = ActKkasse$) Then
            .CellFontBold = True
        Else
            .CellFontBold = False
        End If
    Next i%
    
    .row = row%
    .col = 0
    .ColSel = .Cols - 1
    .Redraw = True
End With

Call PaintAvParameter

With ParentForm
    If (.picRezept.Visible) Then
        iTabIndex% = AktTabIndex%
        .picRezept.SetFocus
        For i% = 0 To (AnzRezeptArtikel% - 1)
            If (RezArtikel(i%).flag And REZ_WG_4) Then
                AktTabIndex% = 3 + i% * 4
                Call Verbandmittel
            End If
        Next i%
        Call PaintRezept
        Call TabIndexChange(iTabIndex%, True)
    Else
        .txtRezeptNr.SetFocus
    End If
End With
    
'If (ParentForm.picRezept.Visible) Then ParentForm.picRezept.SetFocus
    
Call DefErrPop
End Sub

Function DosRezeptDruck%()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("DosRezeptDruck%")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
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
Dim i%, j%, ret%, hFaktor%, hFaktor2%, IstScreenPzn%, gef%
Dim lPreis&
Dim h$, iRezeptDrucker$, PznStr$, hPosNr$
Dim ReverseChar$, FontStr$, BreitStr$, cpi12Str$, SQLStr$

ret% = False

DosDruckerLpt% = -1
h$ = UCase(Left$(RezeptDruckerPara$, 4))
If (Left$(h$, 3) = "LPT") Then
    DosDruckerLpt% = FreeFile
    Open h$ For Output As DosDruckerLpt%
Else
    ret% = OpenDruckerCom%(ParentForm, RezeptDruckerPara$)
    If (ret% = False) Then
        Call iMsgBox("Serielle Schnittstelle kann momentan nicht geöffnet werden", vbCritical, "DosRezeptDruck")
        DosRezeptDruck% = False:  Call DefErrPop: Exit Function
    End If
End If

IstTm290% = False
NeuTm290% = False
IstTmKombi% = False
IstTmU950% = False

iRezeptDrucker$ = RezeptDrucker$

If (iRezeptDrucker$ = "TM-U950") Then
    IstTmKombi% = True
    IstTmU950% = True
    iRezeptDrucker$ = "TM290-II"
End If

If (iRezeptDrucker$ = "TM5000II") Then
    IstTmKombi% = True
    iRezeptDrucker$ = "TM290-II"
End If

If (Left$(iRezeptDrucker$, 5) = "TM290") Then
    IstTm290% = True
    ReverseChar$ = "h"
    FontStr$ = "!" + Chr$(1)
    BreitStr$ = "!" + Chr$(&H21)
    cpi12Str$ = " " + Chr$(2)
    If (iRezeptDrucker$ = "TM290-II") Then
        NeuTm290% = True
        ReverseChar$ = "F"
    End If
End If


DruckStr$ = Chr$(27) + "@"
Call DosDruckerSend(DruckStr$)

If (IstTmKombi%) Then
    DruckStr$ = Chr$(27) + "c0" + Chr$(4)
    Call DosDruckerSend(DruckStr$)
    DruckStr$ = Chr$(27) + "c1" + Chr$(4)
    Call DosDruckerSend(DruckStr$)
    If (IstTmU950% = False) Then
        DruckStr$ = Chr$(&H1D) + "L" + Chr$(152) + Chr$(1)
        Call DosDruckerSend(DruckStr$)
    End If
    DruckStr$ = Chr$(&H1D) + "P" + Chr$(150) + Chr$(144)
    Call DosDruckerSend(DruckStr$)
    ret% = True
Else
    ret% = WarteAufRezept%
End If

If (ret% = False) Then
    ParentForm.comSenden.PortOpen = False
    DosRezeptDruck% = False:  Call DefErrPop: Exit Function
End If

DruckStr$ = Chr$(27) + FontStr$
Call DosDruckerSend(DruckStr$)
DruckStr$ = Chr$(27) + cpi12Str$
Call DosDruckerSend(DruckStr$)

DruckStr$ = Chr$(27) + "3" + Chr$(13)
Call DosDruckerSend(DruckStr$)
If (NeuTm290%) Then
    DruckStr$ = Chr$(27) + "e" + Chr$(4)
    Call DosDruckerSend(DruckStr$)
    If (IstTmKombi%) Then
        DruckStr$ = Chr$(&H1D) + "P" + Chr$(150) + Chr$(58)
        Call DosDruckerSend(DruckStr$)
        If (IstTmU950%) Then
            DruckStr$ = Chr$(27) + "3" + Chr$(12)
            Call DosDruckerSend(DruckStr$)
            DruckStr$ = vbCrLf
            Call DosDruckerSend(DruckStr$)
        End If
        DruckStr$ = Chr$(27) + "3" + Chr$(13)
        Call DosDruckerSend(DruckStr$)
    End If
Else
    DruckStr$ = Chr$(27) + ReverseChar$ + Chr$(1) + vbCrLf + vbCrLf + vbCrLf + Chr$(27) + ReverseChar$ + Chr$(0)
    Call DosDruckerSend(DruckStr$)
End If


Call TmU950LeftMargin(1)

DruckStr$ = Chr$(27) + cpi12Str$
Call DosDruckerSend(DruckStr$)
If (NeuTm290%) Then
    DruckStr$ = Space$(23)
Else
    DruckStr$ = Space$(20)
End If
Call DosDruckerSend(DruckStr$)

DruckStr$ = Chr$(27) + " "
If (NeuTm290%) Then
    DruckStr$ = DruckStr$ + Chr$(5)
Else
    DruckStr$ = DruckStr$ + Chr$(2)
End If
Call DosDruckerSend(DruckStr$)

DruckStr$ = "+" + RezApoNr$ + "+"
DruckStr$ = DruckStr$ + vbCrLf
DruckStr$ = DruckStr$ + vbCrLf
Call DosDruckerSend(DruckStr$)


Call TmU950LeftMargin(1)


DruckStr$ = Chr$(27) + "3" + Chr$(11)
Call DosDruckerSend(DruckStr$)

If (NeuTm290%) Then
    DruckStr$ = Space$(6)
Else
    DruckStr$ = Space$(7)
End If
Call DosDruckerSend(DruckStr$)

DruckStr$ = Chr$(27) + " "
If (NeuTm290%) Then
    DruckStr$ = DruckStr$ + Chr$(5)
Else
    DruckStr$ = DruckStr$ + Chr$(2)
End If
Call DosDruckerSend(DruckStr$)

If (PrivatRezept%) Then
    DruckStr$ = Space$(7)
ElseIf (RezGebSumme# = 0#) Then
    DruckStr$ = "   0   "
Else
    DruckStr$ = Right$(Space$(7) + Format(RezGebSumme#, "0.00"), 7)
    For i% = 1 To Len(DruckStr$)
        If (Mid$(DruckStr$, i%, 1) = ".") Then
            Mid$(DruckStr$, i%, 1) = ","
        End If
    Next i%
End If
Call DosDruckerSend(DruckStr$)



If (NeuTm290%) Then
    DruckStr$ = Space$(2)
    Call DosDruckerSend(DruckStr$)
End If


If (NeuTm290%) Then
    i% = 12
Else
    i% = 14
End If
DruckStr$ = Right$(Space$(i%) + Format(RezSumme#, "0.00"), i%)
For i% = 1 To Len(DruckStr$)
    If (Mid$(DruckStr$, i%, 1) = ".") Then
        Mid$(DruckStr$, i%, 1) = ","
    End If
Next i%
DruckStr$ = DruckStr$ + vbCrLf
DruckStr$ = DruckStr$ + vbCrLf
Call DosDruckerSend(DruckStr$)

DruckStr$ = Chr$(27) + "3" + Chr$(10)
Call DosDruckerSend(DruckStr$)

For j% = 0 To (AnzRezeptArtikel% - 1)
    Call TmU950LeftMargin(1)

    If (NeuTm290%) Then
        DruckStr$ = Space$(9)
    Else
        DruckStr$ = Space$(10)
    End If
    Call DosDruckerSend(DruckStr$)
    
    PznStr$ = Space$(7)
    If (Val(RezArtikel(j%).pzn) <> 9999999) Then
        PznStr$ = RezArtikel(j%).pzn
    End If
    
    hFaktor% = RezArtikel(j%).Faktor
    If (Val(RezArtikel(j%).pzn) = 9999057) Then
        hFaktor% = 0
    End If
    hFaktor2% = hFaktor%

    
    IstScreenPzn% = 0
'    If (Left$(PznStr$, 1) = "@") Then
    If (Left$(PznStr$, 1) = "-") Then
        PznStr$ = RezArtikel(j%).ScreenPzn
        IstScreenPzn% = True
'    ElseIf (ActKasse% = 1) And (ActBundesland% = 3) Then
'        If (Trim(PznStr$) <> "") Then
'            If (AplusVOk) Then
'                SQLStr$ = "SELECT * FROM VdbArtikel WHERE PZN = " + PznStr
''                Set VdbPznRec = AplusVDB.OpenRecordset(SQLStr$)
'                FabsErrf = AplusVDB.OpenRecordset(VdbPznRec, SQLStr)
'                If Not (VdbPznRec.EOF) Then
'                    gef% = 0
'                    hPosNr$ = Trim(Left$(CheckNullStr(VdbPznRec!HmPositionsNr1), 10))
'                    If (hPosNr$ <> "") Then
''                        If (Trim(Left$(CheckNullStr(VdbPznRec!HmPositionsNr2), 10)) = "") Then
'                            PznStr$ = hPosNr$
'                            IstScreenPzn% = True
''                        End If
'
'                        h$ = Trim(CheckNullStr(VdbPznRec!menge))
'                        If (h$ <> "") Then
'                            If (Val(h$) > 0) Then
'                                hFaktor2% = Val(h$)
'                                gef% = True
'                            End If
'                        End If
'
'                        If (gef = 0) Then
'                            SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + RezArtikel(j%).pzn
'                            'Set TaxeRec = TaxeDB.OpenRecordset(SQLStr$)
'                            On Error Resume Next
'                            TaxeRec.Close
'                            Err.Clear
'                            On Error GoTo DefErr
'                            TaxeRec.Open SQLStr, taxeAdoDB.ActiveConn
'                            If (TaxeRec.EOF = False) Then
'                                hFaktor2% = Val(TaxeRec!menge)
'                                gef% = True
'                            End If
'                        End If
'
'                        If (gef%) Then
'                            hFaktor2% = hFaktor2% * hFaktor%
'                        End If
'                    End If
'                End If
'            Else
'                FabsErrf% = VmPzn.IndexSearch(0, PznStr$, FabsRecno&)
'                If (FabsErrf% = 0) Then
'                    Call VmPzn.GetRecord(FabsRecno&)
'
'                    gef% = 0
'                    hPosNr$ = Trim(Left$(VmPzn.HmPositionsNr1, 10))
'                    If (hPosNr$ <> "") Then
''                        If (Trim(Left$(VmPzn.HmPositionsNr2, 10)) = "") Then
'                            PznStr$ = hPosNr$
'                            IstScreenPzn% = True
''                        End If
'
'                        h$ = Trim(VmPzn.menge)
'                        If (h$ <> "") Then
'                            If (Val(h$) > 0) Then
'                                hFaktor2% = Val(h$)
'                                gef% = True
'                            End If
'                        End If
'
'                        If (gef = 0) Then
'                            SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + RezArtikel(j%).pzn
'                            'Set TaxeRec = TaxeDB.OpenRecordset(SQLStr$)
'                            On Error Resume Next
'                            TaxeRec.Close
'                            Err.Clear
'                            On Error GoTo DefErr
'                            TaxeRec.Open SQLStr, taxeAdoDB.ActiveConn
'                            If (TaxeRec.EOF = False) Then
'                                hFaktor2% = Val(TaxeRec!menge)
'                                gef% = True
'                            End If
'                        End If
'
'                        If (gef%) Then
'                            hFaktor2% = hFaktor2% * hFaktor%
'                        End If
'                    End If
'                End If
'            End If
'        End If
    End If


    If (IstScreenPzn%) Then
        DruckStr$ = Chr$(27) + " "
        If (NeuTm290%) Then
            DruckStr$ = DruckStr$ + Chr$(2)
        Else
            DruckStr$ = DruckStr$ + Chr$(1)
        End If
        Call DosDruckerSend(DruckStr$)
        
        DruckStr$ = PznStr$
        Call DosDruckerSend(DruckStr$)
        
        DruckStr$ = Chr$(27) + " "
        If (NeuTm290%) Then
            DruckStr$ = DruckStr$ + Chr$(5)
        Else
            DruckStr$ = DruckStr$ + Chr$(1)
        End If
        Call DosDruckerSend(DruckStr$)
        
        If (NeuTm290%) Then
            DruckStr$ = Space$(1)
        Else
            DruckStr$ = Space$(2)
        End If
        Call DosDruckerSend(DruckStr$)
    Else
        DruckStr$ = PznStr$
        If (NeuTm290%) Then DruckStr$ = DruckStr$ + Space$(1)
        Call DosDruckerSend(DruckStr$)
        
        If (NeuTm290% = False) Then
            DruckStr$ = Chr$(27) + " " + Chr$(1)
            Call DosDruckerSend(DruckStr$)
        End If
        
        If (NeuTm290%) Then
            DruckStr$ = Space$(1)
        Else
            DruckStr$ = Space$(4)
        End If
        Call DosDruckerSend(DruckStr$)
    End If
    
    
'    If (hFaktor2% = 1) Then
'        DruckStr$ = Space$(4)
'    Else
        DruckStr$ = Right$(Space$(4) + Format(hFaktor2%, "0"), 4)
'    End If
    Call DosDruckerSend(DruckStr$)


    lPreis& = CLng(RezArtikel(j%).Preis * hFaktor% * 100#)

'    /* ab 1.73: weil bei Beträgen > 10000.00 in nächste Zeile gedruckt wird */
    If (NeuTm290% = False) Or (lPreis& <= 999900) Then
        DruckStr$ = Space$(1)
        Call DosDruckerSend(DruckStr$)
    End If
    
    If (NeuTm290% = False) Then
        If (lPreis& > 99900) Then
            DruckStr$ = Chr$(27) + " " + Chr$(1)
            Call DosDruckerSend(DruckStr$)
        End If
        
        DruckStr$ = Right$(Space$(6) + Format(lPreis&, "0"), 6)
        Call DosDruckerSend(DruckStr$)
        
        If (lPreis& > 99900) Then
            DruckStr$ = Chr$(27) + " " + Chr$(1)
            Call DosDruckerSend(DruckStr$)
        End If
    Else
        If (lPreis& > 99900) Then
            DruckStr$ = Chr$(27) + " " + Chr$(2)
            Call DosDruckerSend(DruckStr$)
            
            DruckStr$ = Right$(Space$(6) + Format(lPreis&, "0"), 6)
            Call DosDruckerSend(DruckStr$)
            
            DruckStr$ = Chr$(27) + " " + Chr$(5)
            Call DosDruckerSend(DruckStr$)
        Else
            DruckStr$ = Right$(Space$(5) + Format(lPreis&, "0"), 5)
            Call DosDruckerSend(DruckStr$)
        End If
    End If

    If (j% < 6) Then
        If (j% = 5) Then
            DruckStr$ = Chr$(27) + "3" + Chr$(5)
            Call DosDruckerSend(DruckStr$)
        End If
        
        DruckStr$ = vbCrLf + vbCrLf
        Call DosDruckerSend(DruckStr$)
        
        If (j% = 5) Then
            DruckStr$ = Chr$(27) + "3" + Chr$(10)
            Call DosDruckerSend(DruckStr$)
        End If
    End If
    
    If (NeuTm290% = False) Then
        DruckStr$ = Chr$(27) + cpi12Str$
        Call DosDruckerSend(DruckStr$)
    End If

    If (j% >= 6) Then Exit For
Next j%
    

Do
    If (j% >= 6) Then Exit Do
    
    If (j% = 5) Then
        DruckStr$ = Chr$(27) + "3" + Chr$(5)
        Call DosDruckerSend(DruckStr$)
    End If
    
    DruckStr$ = vbCrLf + vbCrLf
    Call DosDruckerSend(DruckStr$)
    
    If (j% = 5) Then
        DruckStr$ = Chr$(27) + "3" + Chr$(10)
        Call DosDruckerSend(DruckStr$)
    End If
    
    j% = j% + 1
Loop

Call TmU950LeftMargin(1)

DruckStr$ = Chr$(27) + " "
If (NeuTm290%) Then
    DruckStr$ = DruckStr$ + Chr$(2)
Else
    DruckStr$ = DruckStr$ + Chr$(1)
End If
Call DosDruckerSend(DruckStr$)

If ((AnzRezeptArtikel% >= 7) Or (lRezNr& = 0&) Or PrivatRezept%) Then
    DruckStr$ = vbCrLf
Else
    If (PersonalNr% > 0) Then
        DruckStr$ = Space$(17) + "(" + RezNr$ + ")" + Format(PersonalNr%, "00") + vbCrLf
    Else
        DruckStr$ = Space$(19) + "(" + RezNr$ + ")" + vbCrLf
    End If
End If
Call DosDruckerSend(DruckStr$)


Call TmU950LeftMargin(0)

DruckStr$ = Chr$(27) + " "
If (NeuTm290%) Then
    DruckStr$ = DruckStr$ + Chr$(2)
Else
    DruckStr$ = DruckStr$ + Chr$(1)
End If
Call DosDruckerSend(DruckStr$)


'DruckStr$ = Left$(RezApoName$(0) + " " + RezApoName$(1) + Space$(27), 27)
If (BtmRezept%) Then
    DruckStr$ = Space$(27)
Else
    DruckStr$ = Left$(RezApoDruckName$ + Space$(27), 27)
End If
DruckStr$ = DruckStr$ + " " + VerkDatum$ + vbCrLf
Call DosDruckerSend(DruckStr$)

DruckStr$ = Chr$(27) + ReverseChar$ + Chr$(1) + Chr$(12) + Chr$(27) + ReverseChar$ + Chr$(0)
Call DosDruckerSend(DruckStr$)

DruckStr$ = Chr$(27) + "q"
Call DosDruckerSend(DruckStr$)

'neu in 1.0.31 wegen Kombidrucker
If (IstTmKombi%) Then
    DruckStr$ = Chr$(27) + "@"
    Call DosDruckerSend(DruckStr$)
End If

If (DosDruckerLpt% >= 0) Then
    Close #DosDruckerLpt%
Else
    ParentForm.comSenden.PortOpen = False
End If

DosRezeptDruck% = True

Call DefErrPop
End Function

Sub TmU950LeftMargin(MarginModus%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("TmU950LeftMargin")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
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

If (IstTmU950%) Then
    DruckStr$ = Chr$(27) + " " + Chr$(2)
    Call DosDruckerSend(DruckStr$)

    If (MarginModus% = 1) Then
        DruckStr$ = Space$(37)
    Else
        DruckStr$ = Space$(22)
    End If
    Call DosDruckerSend(DruckStr$)

    DruckStr$ = Chr$(27) + " " + Chr$(5)
    Call DosDruckerSend(DruckStr$)
End If

Call DefErrPop
End Sub

Sub DosDruckerSend(DruckStr$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("DosDruckerSend")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
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

If (DosDruckerLpt% >= 0) Then
    Print #DosDruckerLpt%, DruckStr$;
Else
    Call SeriellSend(DruckStr$)
End If

Call DefErrPop
End Sub

Sub PruefeRezkontrDat()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("PruefeRezkontrDat")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
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
Dim i%, j%, k%, RezeptGef%, REZK%, SollAbholerInd%, ErrNo%
Dim TM_NAMEN%, TM_DATEN%, TmOffen%, AnzTm%
Dim SatzPtr&
Dim KistePreis#
Dim h$, SQLStr$, SollRezNr$
Dim iRec As Recordset

If (Dir("rezkontr.dat") <> "") Then
    REZK% = FileOpen("rezkontr.dat", "R", "B")
    
    k% = -1
    For i% = 0 To (AnzRezeptArtikel% - 1)
        h$ = RezArtikel(i%).zusatz(0)
        If (Left$(h$, 6) = "Abhol-") Then
            If (InStr(h$, "Besorger") = 0) Then
                ActKiste% = Val(Mid$(h$, 9, 4))
                
                If (k% = -1) Then
                    RezeptGef% = False
                    
                    If (AbholerSQL) Then
                        Dim iAdoRec As New ADODB.Recordset
                        
                        If (AbholerConnOpen = 0) Then
                        End If
                        
                        If (PrivatRezept = 0) And (Val(RezNr$) > 0) Then
                            SollRezNr$ = RezNr$
                        Else
                            SollRezNr$ = "P" + Space$(12)
                        End If
                    
                        SQLStr$ = "SELECT * FROM Taxierungen WHERE AbholerNummer =" + Str$(ActKiste%)
                        SQLStr$ = SQLStr$ + " ORDER BY AbholerInd+AnfMagInd"
                        On Error Resume Next
                        iAdoRec.Close
                        Err.Clear
                        On Error GoTo DefErr
                        iAdoRec.Open SQLStr, AbholerConn    ', adOpenDynamic, adLockOptimistic
                        Do
                            If (iAdoRec.EOF) Then
                                Exit Do
                            End If
                            
                            If (iAdoRec!RezeptNr = SollRezNr$) Then
                                RezeptGef% = True
                                Exit Do
                            End If
                            
                            iAdoRec.MoveNext
                        Loop
                        
                        If (RezeptGef%) Then
                            MAG_SPEICHER% = FileOpen("MagTax" + Trim(para.User), "RW", "B")
                            If (MAG_SPEICHER% > 0) Then
                                MagSpeicherIndex% = LOF(MAG_SPEICHER%) / Len(TaxierRec) + 1
                                
                                j% = 0
                                SollAbholerInd% = iAdoRec!AbholerInd
                                Seek #MAG_SPEICHER%, ((MagSpeicherIndex% - 1) * CLng(Len(TaxierRec))) + 1
                                Do
                                    If (iAdoRec.EOF) Then
                                        Exit Do
                                    End If
                                    
                                    If (iAdoRec!AbholerInd <> SollAbholerInd%) Then
                                        Exit Do
                                    End If
                                    
                                    TaxierRec.ActPreis = iAdoRec!ActPreis
                                    TaxierRec.ActMenge = iAdoRec!ActMenge
                                    TaxierRec.Meh = iAdoRec!einheit
                                    TaxierRec.kurz = iAdoRec!text
                                    TaxierRec.pzn = PznString(iAdoRec!pzn)
                                    TaxierRec.flag = iAdoRec!flag
                                    TaxierRec.kp = iAdoRec!kp
                                    TaxierRec.GStufe = iAdoRec!GStufe
                                    TaxierRec.Verwurf = 0
                                    
                                    If (Trim(TaxierRec.kurz) <> "") Then
                                        Put #MAG_SPEICHER%, , TaxierRec
                                        j% = j% + 1
                                    End If
                                    
                                    iAdoRec.MoveNext
                                Loop
                                
                                For k% = (j% + 1) To 100
                                    TaxierRec.flag = 255
                                    Put #MAG_SPEICHER%, , TaxierRec
                                Next k%
                                
                                Close #MAG_SPEICHER%
                                
                                RezArtikel(i%).MagIndex = MagSpeicherIndex%
                            End If
                        End If
                        
                        AbholerConn.Close
                    ElseIf (AbholerMdb%) Then
                        If (PrivatRezept = 0) And (Val(RezNr$) > 0) Then
                            SollRezNr$ = RezNr$
                        Else
                            SollRezNr$ = "P" + Space$(12)
                        End If
                    
                        Set AbholerDB = OpenDatabase(ABHOLER_DB, False, True)
                        SQLStr$ = "SELECT * FROM Taxierungen WHERE AbholerNummer =" + Str$(ActKiste%)
                        SQLStr$ = SQLStr$ + " ORDER BY AbholerInd+AnfMagInd"
                        Set iRec = AbholerDB.OpenRecordset(SQLStr$)
                        Do
                            If (iRec.EOF) Then
                                Exit Do
                            End If
                            
                            If (iRec!RezeptNr = SollRezNr$) Then
                                RezeptGef% = True
                                Exit Do
                            End If
                            
                            iRec.MoveNext
                        Loop
                        
                        If (RezeptGef%) Then
                            MAG_SPEICHER% = FileOpen("MagTax" + Trim(para.User), "RW", "B")
                            If (MAG_SPEICHER% > 0) Then
                                MagSpeicherIndex% = LOF(MAG_SPEICHER%) / Len(TaxierRec) + 1
                                
                                j% = 0
                                SollAbholerInd% = iRec!AbholerInd
                                Seek #MAG_SPEICHER%, ((MagSpeicherIndex% - 1) * CLng(Len(TaxierRec))) + 1
                                Do
                                    If (iRec.EOF) Then
                                        Exit Do
                                    End If
                                    
                                    If (iRec!AbholerInd <> SollAbholerInd%) Then
                                        Exit Do
                                    End If
                                    
                                    TaxierRec.ActPreis = iRec!ActPreis
                                    TaxierRec.ActMenge = iRec!ActMenge
                                    TaxierRec.Meh = iRec!einheit
                                    TaxierRec.kurz = iRec!text
                                    TaxierRec.pzn = PznString(iRec!pzn)
                                    TaxierRec.flag = iRec!flag
                                    TaxierRec.kp = iRec!kp
                                    TaxierRec.GStufe = iRec!GStufe
                                    TaxierRec.Verwurf = 0
                                    
                                    If (Trim(TaxierRec.kurz) <> "") Then
                                        Put #MAG_SPEICHER%, , TaxierRec
                                        j% = j% + 1
                                    End If
                                    
                                    iRec.MoveNext
                                Loop
                                
                                For k% = (j% + 1) To 100
                                    TaxierRec.flag = 255
                                    Put #MAG_SPEICHER%, , TaxierRec
                                Next k%
                                
                                Close #MAG_SPEICHER%
                                
                                RezArtikel(i%).MagIndex = MagSpeicherIndex%
                            End If
                        End If
                        
                        AbholerDB.Close
                    End If
                    
                    Seek #REZK%, CLng(ActKiste% - 1) * Len(RezKontrSatz) + 1
                    Get #REZK%, , RezKontrSatz
                    
                    RezeptGef% = False
                    For k% = 0 To 9
                        h$ = Trim(RezKontrSatz.RezEan(k).ean)
                        If (PrivatRezept And (h$ = "P")) Or (h$ = RezNr$) Then
                            RezeptGef% = True
                            Exit For
                        End If
                    Next k%
                End If
            
                If (RezeptGef%) And (RezKontrSatz.fertig = "*") Then
                    KistePreis# = RezKontrSatz.Preis(k%)
                    If (KistePreis# > 0#) Then
                        With RezArtikel(i%)
                            .pzn = "09999011"
                            .text = "Rezeptur"
                            .Preis = KistePreis#
                            
                            .Faktor = 1
                
                            .flag = .flag And (REZ_BLINK Xor &HFF)
                            .Zuz = CalcZuzaNeu#(KistePreis#) ' ZuzFeld!(3)
                            
                            If Not (GebFrei%) And (.Zuz > (KistePreis# / VmRabattFaktor#)) Then
                                RezArtikel(i%).Preis = 0#
                                RezArtikel(i%).Zuz = 0#
                                Call iMsgBox("Artikel unter Zuzahlung, Geldrückgabe ?")
                            End If
                            
                            .AutIdem = 0
                            .N0 = 0
                            .Imp123 = 0
                            .ImpErspart = 0
                            .Fam = 0
                            .TkkPznDruck = 0
                        End With
                    End If
                    
                    RezKontrSatz.Preis(k%) = 0#
                    RezKontrSatz.RezEan(k%).ean = Space$(13)
                    k% = k% + 1
                End If
            
            End If
        End If
    Next i%
    
    Close (REZK%)
End If


TmOffen = 0
If (TaxmusterDBok) Then
    Dim ParenteralTmActMenge%
    
    For i% = 0 To (AnzRezeptArtikel% - 1)
        If (RezArtikel(i%).Warenzeichen = 98) Then
            If (SollTmName(i) <> "") Then
                FormErg = 0
                
                SQLStr$ = "SELECT * FROM ParenteralTm WHERE Id=" + Format(SollTmId(i), "0")
                Set ParenteralTmRec = RezSpeicherDB.OpenRecordset(SQLStr$)
                If Not (ParenteralTmRec.EOF) Then
                    With RezArtikel(i%)
                        .pzn = "09999011"
                        .text = "Rezeptur"
                        KistePreis# = .Preis
                        
                        .Faktor = 1
            
                        .flag = .flag And (REZ_BLINK Xor &HFF)
                        .Zuz = CalcZuzaNeu#(KistePreis#) ' ZuzFeld!(3)
    
                        If Not (GebFrei%) And (.Zuz > (KistePreis# / VmRabattFaktor#)) Then
                            RezArtikel(i%).Preis = 0#
                            RezArtikel(i%).Zuz = 0#
                            Call iMsgBox("Artikel unter Zuzahlung, Geldrückgabe ?")
                        End If
                        
                        .AutIdem = 0
                        .N0 = 0
                        .Imp123 = 0
                        .ImpErspart = 0
                        .Fam = 0
                        .TkkPznDruck = 0
                    End With
                    
                    'ParenteralTmActMenge = ParenteralTmRec!ActMenge
                    FormErg = ParenteralTmRec!Id
                End If
                ParenteralTmRec.Close
                
                If (FormErg%) Then
                    Load frmTaxieren
                    'TmHeader.ActMenge = ParenteralTmActMenge
                    TmMengenFaktor# = 1     'SollTmMenge(i) / ParenteralTmRec!ActMenge
    '                Call HoleTaxMuster
                    
                    SQLStr$ = "SELECT * FROM ParenteralTmZeilen WHERE TmId=" + CStr(FormErg)
                    SQLStr$ = SQLStr$ + " ORDER BY AnfMagInd"
                    Set ParenteralTmRec = RezSpeicherDB.OpenRecordset(SQLStr$)
                    
                    k = 0
                    Do
                        If (ParenteralTmRec.EOF) Then
                            Exit Do
                        End If
                                
                        TaxierRec.ActPreis = 0
                        
                        TmInhalt.pzn = PznString(ParenteralTmRec!pzn)
            
                        If (k = 0) Then
                            For j% = 0 To UBound(ParenteralPzn)
                                If (ParenteralTmRec!pzn = Val(ParenteralPzn(j))) Then
                                    ParenteralRezept = j
                                    
                                    If (ParenteralRezept >= 0) And (ParenteralRezept < 31) Then
                                        MalFaktor# = 1
                                    End If
                                    
                                    With RezArtikel(i%)
                                        .pzn = TmInhalt.pzn
                                        .text = ParenteralTxt(j)
                                    End With
                                    
                                    Exit For
                                End If
                            Next j
                        End If
                        k = k + 1
                                                                       
                        TmInhalt.kurz = ParenteralTmRec!text
                        
                        TmInhalt.ActMenge = ParenteralTmRec!ActMenge
                        TmInhalt.ActPreis = ParenteralTmRec!ActPreis
                        TmInhalt.flag = ParenteralTmRec!flag
                
                        ParEnteralPrimärPackmittel = ParenteralTmRec!packmittel
                        ParEnteralAI = ParenteralTmRec!AI
                        ParEnteralAnzEinheiten = ParenteralTmRec!WirkstoffMenge
                        ParEnteralPrimärPackmittel = ParEnteralPrimärPackmittel

                        Call HoleTaxMusterZeile
                        frmTaxieren.flxTaxieren.AddItem " "
                        Call frmTaxieren.ZeigeTaxierZeile(frmTaxieren.flxTaxieren.Rows - 1)
                        If (TmInhalt.ActPreis = 0) Then
                            With RezArtikel(i%)
                                .flag = .flag Or (REZ_BLINK)
                            End With
'                            MsgBox ("rot")
                        End If
                                
                        ParenteralTmRec.MoveNext
                    Loop
                    ParenteralTmRec.Close
                    
                    MagSpeicherIndex% = LOF(MAG_SPEICHER%) / Len(TaxierRec) + 1
                    RezArtikel(i%).MagIndex = MagSpeicherIndex%
                    If (MagSpeicherIndex% > 0) Then
                        Call frmTaxieren.SpeicherMagSpeicher
                    End If
                    If (MAG_SPEICHER% > 0) Then Close (MAG_SPEICHER%)
                    Unload frmTaxieren
                End If
            End If
        ElseIf (RezArtikel(i%).Warenzeichen = 99) Then
            If (TmOffen = 0) Then
                Set TaxmusterDB = OpenDatabase(TAXMUSTER_DB, False, False)
                TmOffen = True
            End If
            
            If (SollTmName(i) <> "") Then
                FormErg = 0
                
                SQLStr$ = "SELECT * FROM Taxmuster"
                Set TaxmusterRec = TaxmusterDB.OpenRecordset(SQLStr$)
                Do
                    If (TaxmusterRec.EOF) Then
                        Exit Do
                    End If
                    
                    h$ = Trim(CheckNullStr(TaxmusterRec!Bezeichnung))
                    If (h$ = UCase(SollTmName(i))) Then
                        With RezArtikel(i%)
                            .pzn = "09999011"
                            .text = "Rezeptur"
                            KistePreis# = .Preis
                            
                            .Faktor = 1
                
                            .flag = .flag And (REZ_BLINK Xor &HFF)
                            .Zuz = CalcZuzaNeu#(KistePreis#) ' ZuzFeld!(3)

                            If Not (GebFrei%) And (.Zuz > (KistePreis# / VmRabattFaktor#)) Then
                                RezArtikel(i%).Preis = 0#
                                RezArtikel(i%).Zuz = 0#
                                Call iMsgBox("Artikel unter Zuzahlung, Geldrückgabe ?")
                            End If
                            
                            .AutIdem = 0
                            .N0 = 0
                            .Imp123 = 0
                            .ImpErspart = 0
                            .Fam = 0
                            .TkkPznDruck = 0
                        End With
                        
                        FormErg = TaxmusterRec!Id
                        Exit Do
                    End If
                    
                    TaxmusterRec.MoveNext
                Loop
                
                If (FormErg%) Then
                    Load frmTaxieren
                    TmHeader.ActMenge = TaxmusterRec!ActMenge
                    TmMengenFaktor# = SollTmMenge(i) / TaxmusterRec!ActMenge
    '                Call HoleTaxMuster
                    
                    SQLStr$ = "SELECT * FROM TaxmusterZeilen WHERE TaxmusterId=" + CStr(FormErg)
                    SQLStr$ = SQLStr$ + " ORDER BY LaufNr"
                    Set TaxmusterRec = TaxmusterDB.OpenRecordset(SQLStr$)
                    Do
                        If (TaxmusterRec.EOF) Then
                            Exit Do
                        End If
                                
                        TaxierRec.ActPreis = 0
                        TmInhalt.pzn = PznString(TaxmusterRec!pzn)
                        
                        TmInhalt.kurz = TaxmusterRec!Name
                        
                        TmInhalt.ActMenge = TaxmusterRec!ActMenge
                        TmInhalt.ActPreis = TaxmusterRec!ActPreis
                        TmInhalt.flag = TaxmusterRec!flag
                
                        Call HoleTaxMusterZeile
                        frmTaxieren.flxTaxieren.AddItem " "
                        Call frmTaxieren.ZeigeTaxierZeile(frmTaxieren.flxTaxieren.Rows - 1)
                        If (TmInhalt.ActPreis = 0) Then
                            With RezArtikel(i%)
                                .flag = .flag Or (REZ_BLINK)
                            End With
'                            MsgBox ("rot")
                        End If
                                
                        TaxmusterRec.MoveNext
                    Loop
                    
                    MagSpeicherIndex% = LOF(MAG_SPEICHER%) / Len(TaxierRec) + 1
                    RezArtikel(i%).MagIndex = MagSpeicherIndex%
                    If (MagSpeicherIndex% > 0) Then
                        Call frmTaxieren.SpeicherMagSpeicher
                    End If
                    If (MAG_SPEICHER% > 0) Then Close (MAG_SPEICHER%)
                    Unload frmTaxieren
                End If
            End If
        End If
    Next i%
    If (TmOffen) Then
        TaxmusterDB.Close
    End If
Else
    For i% = 0 To (AnzRezeptArtikel% - 1)
        If (RezArtikel(i%).Warenzeichen = 99) Then
            If (TmOffen = 0) Then
                TM_NAMEN% = FileOpen("tmnamen.dat", "RW", "B")
                TM_DATEN% = FileOpen("tmdaten.dat", "RW", "B")
                TmOffen = True
            End If
            
            If (SollTmName(i) <> "") Then
                FormErg = 0
                If (TM_NAMEN% > 0) Then
                    AnzTm% = LOF(TM_NAMEN%) / Len(TmHeader)
                
                    Seek #TM_NAMEN%, 1
                    Get #TM_NAMEN%, , TmHeader
                    For j% = 1 To AnzTm%
                        Get #TM_NAMEN%, , TmHeader
                        If (EOF(TM_NAMEN%)) Then Exit For
                        
                        h$ = Trim(TmHeader.Name)
                        Call OemToChar(h$, h$)
                        h$ = SqlString(h$)
                        
                        If (h$ = UCase(SollTmName(i))) Then
                            With RezArtikel(i%)
                                .pzn = "09999011"
                                .text = "Rezeptur"
                                KistePreis# = .Preis
                                
                                .Faktor = 1
                    
                                .flag = .flag And (REZ_BLINK Xor &HFF)
                                .Zuz = CalcZuzaNeu#(KistePreis#) ' ZuzFeld!(3)
    
                                If Not (GebFrei%) And (.Zuz > (KistePreis# / VmRabattFaktor#)) Then
                                    RezArtikel(i%).Preis = 0#
                                    RezArtikel(i%).Zuz = 0#
                                    Call iMsgBox("Artikel unter Zuzahlung, Geldrückgabe ?")
                                End If
                                
                                .AutIdem = 0
                                .N0 = 0
                                .Imp123 = 0
                                .ImpErspart = 0
                                .Fam = 0
                            End With
                            
                            FormErg = j
                            Exit For
                        End If
                    Next j%
                End If
                If (FormErg%) Then
                    Load frmTaxieren
                    TmMengenFaktor# = SollTmMenge(i) / TmHeader.ActMenge
    '                Call HoleTaxMuster
                
                    SatzPtr& = TmHeader.ErstSatz
                    For j% = 1 To TmHeader.AnzZeilen
                        If (SatzPtr& <= 0&) Then Exit For
                        
                        TaxierRec.ActPreis = 0
                        On Error Resume Next
                        Seek #TM_DATEN%, (SatzPtr& * Len(TmInhalt)) + 1
                        ErrNo% = Err.Number
                        On Error GoTo DefErr
                        Err.Clear
                        If (ErrNo = 0) Then
                            Get #TM_DATEN%, , TmInhalt
                            If (EOF(TM_DATEN%)) Then Exit For
                            Call HoleTaxMusterZeile
                        End If
                        
                        frmTaxieren.flxTaxieren.AddItem " "
                        Call frmTaxieren.ZeigeTaxierZeile(frmTaxieren.flxTaxieren.Rows - 1)
                                
                        SatzPtr& = TmInhalt.NextSatz
                    Next j%
                    MagSpeicherIndex% = LOF(MAG_SPEICHER%) / Len(TaxierRec) + 1
                    RezArtikel(i%).MagIndex = MagSpeicherIndex%
                    If (MagSpeicherIndex% > 0) Then
                        Call frmTaxieren.SpeicherMagSpeicher
                    End If
                    If (MAG_SPEICHER% > 0) Then Close (MAG_SPEICHER%)
                    Unload frmTaxieren
                End If
            End If
        End If
    Next i%
    If (TmOffen) Then
        Close (TM_NAMEN%)
        Close (TM_DATEN%)
    End If
End If

Call DefErrPop
End Sub

Sub WriteRezkontrDat()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("WriteRezkontrDat")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
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
Dim REZK%
Dim lRecs&
Dim SQLStr$, SollRezNr$
   
If (ActKiste% > 0) Then
    REZK% = FileOpen("rezkontr.dat", "RW", "B")
    Seek #REZK%, CLng(ActKiste% - 1) * Len(RezKontrSatz) + 1
    Put #REZK%, , RezKontrSatz
    Close (REZK%)
    
    If (AbholerSQL) Then
        If (Val(RezNr$) > 0) Then
            SollRezNr$ = RezNr$
        Else
            SollRezNr$ = "P" + Space$(12)
        End If
        
        If (AbholerConnOpen = 0) Then
        End If
        
        Set AbholerDB = OpenDatabase(ABHOLER_DB, False, False)
        SQLStr$ = "DELETE Taxierungen WHERE AbholerNummer =" + Str$(ActKiste%)
        SQLStr$ = SQLStr$ + " AND RezeptNr ='" + SollRezNr$ + "'"
        Call AbholerConn.Execute(SQLStr, lRecs, adExecuteNoRecords)
        AbholerConn.Close
    ElseIf (AbholerMdb%) Then
        If (Val(RezNr$) > 0) Then
            SollRezNr$ = RezNr$
        Else
            SollRezNr$ = "P" + Space$(12)
        End If
        
        Set AbholerDB = OpenDatabase(ABHOLER_DB, False, False)
        SQLStr$ = "DELETE * FROM Taxierungen WHERE AbholerNummer =" + Str$(ActKiste%)
        SQLStr$ = SQLStr$ + " AND RezeptNr ='" + SollRezNr$ + "'"
        AbholerDB.Execute (SQLStr$)
        AbholerDB.Close
    End If
End If

Call DefErrPop
End Sub

Sub PruefeBtmKosten()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("PruefeBtmKosten")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
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
Dim i%, ind%, ABG_HANDLE%, erg%, row%
Dim dPreis#, BtmKosten#(2)
Dim h$, BtmName$(2)
        
For i% = 0 To UBound(AbgabeKosten)
    With AbgabeKosten(i%)
        If (Left$(.Name, 3) = "BTM") Then
            dPreis# = .kosten
            If (dPreis# = 0#) Then
                Do
                    dPreis# = Val(MyInputBox("Kosten für '" + .Name + "' eingeben:", "Abgabe-Preise", "0"))
                    If (dPreis# > 0#) Then Exit Do
                Loop
                .kosten = dPreis#

                ABG_HANDLE% = FileOpen("abgpr.dat", "RW", "B")
                If (ABG_HANDLE% > 0) Then
                    h$ = .Name + Right$(Space$(4) + Format(dPreis# * 100#, "0"), 4)
                    Put #ABG_HANDLE%, (.ind * 16) + 1, h$
                    Close #ABG_HANDLE%
                End If
            End If
            BtmKosten#(ind%) = dPreis#
            BtmName$(ind%) = .Name
            ind% = ind% + 1
            If (ind% >= 2) Then Exit For
        End If
    End With
Next i%

ind% = 0
'weg mit 5.0.5
'If (BtmKosten#(0) <> BtmKosten#(1)) Then
'    erg% = iMsgBox("Handelt es sich um ein Rezept für gesetzliche Kassen ?", vbYesNo Or vbInformation, "BTM-Gebühr")
'    If (erg% <> vbYes) Then ind% = 1
'End If
    

If (BtmAlsZeile%) Then

    AnzRezeptArtikel% = AnzRezeptArtikel% + 1
    row% = AnzRezeptArtikel% - 1
    Call InitRezeptArtikel(row%)
    RezArtikel(row%).text = BtmName$(ind%)
    RezArtikel(row%).pzn = "02567001" ' "8150012"
    RezArtikel(row%).Faktor = 1
    RezArtikel(row%).Preis = BtmKosten#(ind%)
    
Else

    RezSummeDazu# = RezSummeDazu# + BtmKosten#(ind%)
    
    Call CalcRezeptWerte
    Call TabIndexChange(AktTabIndex%, False)
    
    For i% = 0 To UBound(RezAddGesamtBrutto$)
        If (RezAddGesamtBrutto$(i%) = "") Then
            RezAddGesamtBrutto$(i%) = BtmName$(ind%) + Format(BtmKosten#(ind%), "   (0.00)")
            Call EchtKurzInfo
            Exit For
        End If
    Next i%
    
End If
        
Call DefErrPop
End Sub

Sub EditHmNummern()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("EditHmNummern")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
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

EasyMatchModus% = 0
frmEasyMatch.Show 1
If (FormErg%) Then Call ActProgram.LadeSonderPzn("hmnummer.dat", HmNummern$)

Call DefErrPop
End Sub

Sub EditSonderPzn()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("EditSonderPzn")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
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

EasyMatchModus% = 1
frmEasyMatch.Show 1
If (FormErg%) Then Call ActProgram.LadeSonderPzn("sondrpzn.dat", SonderPzn$)
    
Call DefErrPop
End Sub

Sub EditF7Kassen()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("EditF7Kassen")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
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

EasyMatchModus% = 2
frmEasyMatch.Show 1
If (FormErg%) Then Call HolVerbandmittel

Call DefErrPop
End Sub

Sub FlxEasyMatchBefuellen()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("FlxEasyMatchBefuellen")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
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

With frmEasyMatch.flxEasyMatch
    .Rows = 1
    If (EasyMatchModus% = 0) Then
        For i% = 0 To UBound(HmNummern$)
            h$ = Mid$(HmNummern$(i%), 11) + vbTab + Left$(HmNummern$(i%), 10)
            .AddItem h$
        Next i%
    ElseIf (EasyMatchModus% = 1) Then
        For i% = 0 To UBound(SonderPzn$)
            If (InStr("0123456789", Mid(SonderPzn(i), 8, 1)) = 0) Then
                h$ = Mid$(SonderPzn$(i%), 8) + vbTab + Left$(SonderPzn$(i%), 7)
            Else
                h$ = Mid$(SonderPzn$(i%), 9) + vbTab + Left$(SonderPzn$(i%), 8)
            End If
            .AddItem h$
        Next i%
    Else
        For i% = 0 To UBound(F7Kassen)
            h$ = F7Kassen(i%).Name + vbTab
            If (i% > 0) Then h$ = h$ + Format(F7Kassen(i%).kosten, "0")
            .AddItem h$
        Next i%
    End If
    .row = 0
End With

Call DefErrPop
End Sub

Sub AuswahlF7Kassen()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("AuswahlF7Kassen")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
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
Dim i%, row%, EditCol%, ActMenge%, ind%, erg%
Dim spBreite&
Dim ActAep#, dPreis#, dZuz#
Dim h2$, pzn$, SQLStr$, txt$

If (AktTabIndex% < 3) Or (AktTabIndex% = 31) Then Call DefErrPop: Exit Sub

row% = (AktTabIndex% - 3) \ 4
EditCol% = (AktTabIndex - 3) Mod 4
    
If (row% >= AnzRezeptArtikel%) Then Call DefErrPop: Exit Sub


pzn$ = RezArtikel(row%).pzn
If (Val(pzn$) = 0) Then Call DefErrPop: Exit Sub

ActMenge% = RezArtikel(row%).Faktor

SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + pzn$
'Set TaxeRec = TaxeDB.OpenRecordset(SQLStr$)
On Error Resume Next
TaxeRec.Close
Err.Clear
On Error GoTo DefErr
TaxeRec.Open SQLStr, taxeAdoDB.ActiveConn
If (TaxeRec.EOF) Then
    Call DefErrPop: Exit Sub
End If
  
If (TaxeRec!MwStKZ) Then
    ActMwst% = para.Mwst(1)
Else
    ActMwst% = para.Mwst(2)
End If

ActAep# = TaxeRec!EK / 100
dZuz# = RezArtikel(row%).Zuz


If (Left$(TaxeRec!Zuzahlung, 1) = "9") Then
    Call iMsgBox("Achtung: Taxierung entsprechend dem gültigen Liefervertrag !", 0, "Kassen-Aufschlag")
End If
        

With frmEdit.flxEdit
    .Left = 0
    .Top = 0
    
    .Rows = 2
    .FixedRows = 1
    .FormatString = "<Name|>Aufschlag"
    .Rows = 1
    .Cols = 2
    
    .ColWidth(0) = frmEdit.TextWidth(String(28, "X"))
    .ColWidth(1) = frmEdit.TextWidth("9999999.99")
    
    For i% = 0 To UBound(F7Kassen)
        dPreis# = F7Kassen(i%).kosten
        If (dPreis# = 0#) Then
            h2$ = ""
        Else
            h2$ = Format(dPreis#, "0")
        End If
        .AddItem F7Kassen(i%).Name + vbTab + h2$
    Next i%
    
    .Height = .RowHeight(0) * .Rows + 90
    
    spBreite& = 0
    For i% = 0 To (.Cols - 1)
        spBreite& = spBreite& + .ColWidth(i%)
    Next i%
    .Width = spBreite& + 90
    
    frmEdit.Height = .Height
    frmEdit.Width = .Width
    frmEdit.Left = ParentForm.Left + (ParentForm.Width - frmEdit.Width) / 2
    frmEdit.Top = ParentForm.Top + (ParentForm.Height - frmEdit.Height) / 2
    
    .row = 1
    .col = 0
    .ColSel = .Cols - 1
    .Visible = True
End With

frmEdit.Show 1

If (EditErg%) Then
    ind% = InStr(EditTxt$, vbTab)
    txt$ = Left$(EditTxt$, ind% - 1)
    
    h2$ = Mid$(EditTxt$, ind% + 1)
    If (h2$ = "") Then
        dPreis# = 0#
    Else
        dPreis# = CDbl(h2$)
    End If
    
    If (dPreis# = 0#) Then
        dPreis# = CalcAMPV(ActAep#, ActMwst%)
    Else
        dPreis# = ActAep# * (1# + (dPreis#) / 100#)
        dPreis# = dPreis# * (1# + (ActMwst%) / 100#)
    End If
        
    If (Left$(TaxeRec!Zuzahlung, 1) = "9") Then
        dZuz# = dPreis# * 0.2
        dZuz# = FNX(dZuz#)
        If (ActKasse% = 7) Then
            dZuz# = 0#
        End If
        RezArtikel(row%).Zuz = dZuz#
    
        erg% = iMsgBox("Ist dieser Betrag mit dem Rabattfaktor zu multiplizieren:", vbYesNo Or vbInformation, "A+V Taxierung")
        If (erg% = vbYes) Then
            dPreis# = dPreis# * VmRabattFaktor#
        End If
    
    End If
   
    RezArtikel(row%).Preis = dPreis#
    RezArtikel(row%).flag = RezArtikel(row%).flag And (REZ_BLINK Xor &HFF)
    
    If Not (GebFrei%) And (dZuz# > (dPreis# / VmRabattFaktor#)) Then
        RezArtikel(row%).Preis = 0
        RezArtikel(row%).Zuz = 0
    End If

    Call PaintArtikel(row%)
End If

Call DefErrPop
End Sub

Sub ZeigeInhaltsStoffe()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("ZeigeInhaltsStoffe")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
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
Dim i%, j%, iRow%, l%, anz%, p%, AnzInhalt%, InhaltAktiv%, Anzgefunden%
Dim lPtr&, lPtr2&, hrow&, row&, grecno&, ArtikelRecno&
Dim h$, pzn$, SQLStr$, ProgChar$, CDPfad$, M2Nr$, s$, Absatz$


If (AktTabIndex% < 3) Or (AktTabIndex% = 31) Then Call DefErrPop: Exit Sub

iRow% = (AktTabIndex% - 3) \ 4
If (iRow% >= AnzRezeptArtikel%) Then Call DefErrPop: Exit Sub


pzn$ = RezArtikel(iRow%).pzn
If (Val(pzn$) = 0) Then Call DefErrPop: Exit Sub

SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + pzn$
'Set TaxeRec = TaxeDB.OpenRecordset(SQLStr$)
On Error Resume Next
TaxeRec.Close
Err.Clear
On Error GoTo DefErr
TaxeRec.Open SQLStr, taxeAdoDB.ActiveConn
If (TaxeRec.EOF) Then
    Call DefErrPop: Exit Sub
End If
  
M2Nr$ = RTrim$(TaxeRec!M2)
If (Val(M2Nr$) = 0) Then
    Call DefErrPop: Exit Sub
End If
  
If (EinlesenHuffmanTree% = False) Then
    s$ = para.CdLw + ":\cdinfo\hufftree.dat"
    Call iMsgBox("Datei " + s$ + " ist nicht auf der Plattenkopie vorhanden !", vbCritical)
    Call DefErrPop: Exit Sub
End If

InhaltsstoffeModus = 0
Load frmInhaltsStoffe
frmInhaltsStoffe.Caption = "Inhaltsstoffe für " + RezArtikel(iRow%).text
frmInhaltsStoffe.Show 1

Call DefErrPop
End Sub

Sub FlxInhaltsStoffeBefuellen()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("FlxInhaltsStoffeBefuellen")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
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
Dim i%, j%, iRow%, l%, anz%, p%, AnzInhalt%, InhaltAktiv%, Anzgefunden%
Dim lPtr&, lPtr2&, hrow&, row&, grecno&, ArtikelRecno&
Dim h$, pzn$, SQLStr$, ProgChar$, CDPfad$, M2Nr$, s$, Absatz$

M2Nr$ = RTrim$(TaxeRec!M2)

ProgChar$ = "F"
Absatz$ = "M2"
CDPfad$ = para.CdLw + ":\cdinfo\" + ProgChar$ + "\" + ProgChar$
h$ = CDPfad$ + "_a.ptr"
APTR% = FileOpen(h$, "R")
h$ = CDPfad$ + "_a.txt"
ATXT% = FileOpen(h$, "R")

l% = Len(M2Nr$)
h$ = "@BEGIN " + Absatz$
lPtr& = 0
Seek #APTR%, 1
Do While Not (EOF(APTR%))
    Get #APTR%, , lPtr&
    Call HolZeile(ATXT%, lPtr& + 1, s$, -1)
    If (Left$(s$, 9) = h$) Then
        Exit Do
    End If
Loop

Do While Not (EOF(APTR%))
    Get #APTR%, , lPtr&
    Call HolZeile(ATXT%, lPtr& + 1, s$, -1)
    If (Left$(s$, 1) = "@") Then
        Call DefErrPop: Exit Sub
    End If
    If ((l% = 0) Or (Left$(s$, 1) = Left$(M2Nr$, 1))) Then
        p% = InStr(s$, "|")
        row& = Val(Mid$(s$, p% + 1, 7))
        anz% = Val(Mid$(s$, p% + 8, 4))
        Exit Do
    End If
Loop

Seek #APTR%, row& * 4 + 1
For j% = 0 To anz%
    Get #APTR%, , lPtr&
    Call HolZeile(ATXT%, lPtr& + 1, s$, -1)
    p% = InStr(s$, "|")
    h$ = Trim$(Left$(s$, p% - 1))
    row& = Val(Mid$(s$, p% + 1, 7))
    anz% = Val(Mid$(s$, p% + 8, 4))
    If ((j% > 0) And ((Left$(s$, 1) = "@") Or (h$ > M2Nr$) Or (l% = 0))) Then
        Exit For
    End If
    hrow& = row&
Next j%

grecno& = hrow&

Close (APTR%)
Close (ATXT%)

'''''''''''''''''''

h$ = CDPfad$ + "_t.ptr"
TPTR% = FileOpen(h$, "R")
h$ = CDPfad$ + "_t.txt"
TTXT% = FileOpen(h$, "R")

Do While True
    Call TtxtZeile(grecno&)
    If (Left$(buf$, 1) = "@") Then
        Call DefErrPop: Exit Sub
    End If
    If ((l% = 0) Or (Left$(buf$, l%) >= Left$(M2Nr$, l%))) Then
        Exit Do
    End If
    grecno& = grecno& + 1
Loop

Close (TPTR%)
Close (TTXT%)


'''''''''''''''''''

h$ = CDPfad$ + "_z.ptr"
ZPTR% = FileOpen(h$, "R")
h$ = CDPfad$ + "_z.plt"
ZPLT% = FileOpen(h$, "R")
    
lPtr& = grecno&
Seek ZPTR%, lPtr& * 4 + 1
Get #ZPTR%, , lPtr&
Get #ZPTR%, , lPtr2&
Anzgefunden% = (lPtr2& - lPtr& - 1) \ 4
   
Seek #ZPLT%, lPtr& + 1
grecno& = lPtr&
Get #ZPLT%, , ArtikelRecno&

Close (ZPTR%)
Close (ZPLT%)

'''''''''''''''''''

h$ = CDPfad$ + "u.ptr"
UPTR% = FileOpen(h$, "R")
h$ = CDPfad$ + "u.dat"
UDAT% = FileOpen(h$, "R")

Seek UPTR%, ArtikelRecno& * 4 + 1
Get #UPTR%, , lPtr&

AnzInhalt% = 0
Seek UDAT%, lPtr& + 1
  
InhaltAktiv% = False
HuffmanInit% = True
Do
    Call Decomp(UDAT%, s$)
    
    h$ = Left$(s$, 2)
    If (s$ = "") Or (InhaltAktiv% And (h$ <> Space$(2))) Then Exit Do
    
    If (h$ = "HA") Then
'        color("Wi");
'        memset(h, ' ', 46);
'        h[46] = '\0';
'        gotoxy(4, 6);
'        cputs(h);
'        gotoxy(4+(46-strlen(s))/2, 6);
'        cputs(&s[4]);
'        color("W");
    ElseIf (h$ = "ST") Then
        InhaltAktiv% = True
        s$ = Mid$(s$, 6)
    End If
    
    If (InhaltAktiv%) Then
        Call ModiZeile(s$)
        frmInhaltsStoffe.flxInhaltsStoffe.AddItem s$
'        cputs(&s[4]);
        AnzInhalt% = AnzInhalt% + 1
    End If
Loop

Close (UPTR%)
Close (UDAT%)

Call DefErrPop
End Sub

Sub HolZeile(Handle%, lPtr&, s$, AnzLesen%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("HolZeile")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
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
Dim s2$, lPos&, i%, s2i%

s$ = ""
s2$ = ""
If (lPtr& <> -1) Then
    Seek Handle%, lPtr&
End If
lPos& = Seek(Handle%)

If (AnzLesen% = -1) Then
    AnzLesen% = 200
End If

s2$ = Space$(AnzLesen%)
Get #Handle%, , s2$

i% = InStr(s2$, Chr$(10))
If (i% > 0) Then s$ = Left$(s2$, i% - 2)
'For i% = 1 To AnzLesen%
'    s2i% = Asc(Mid$(s2$, i%, 1))
'    'If ((s2i% = 13) Or (s2i% = 10)) Then
'    If (s2i% = 10) Then
''        s$ = Left$(s2$, i% - 1)
'        s$ = Left$(s2$, i% - 2)
'        Exit For
'    End If
'Next i%

If (AnzLesen% = 200) Then
'    Seek Handle%, lPos& + i% + 2
    Seek Handle%, lPos& + i%
End If

s$ = Trim$(s$)
Call OemToChar(s$, s$)

Call DefErrPop
End Sub

Sub TtxtZeile(lPtr&)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("TtxtZeile")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
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
Dim lPtr2&, s$, p%

Seek TPTR%, lPtr& * 4 + 1
Get TPTR%, , lPtr2&
Call HolZeile(TTXT%, lPtr2& + 1, s$, -1)

p% = InStr(s$, "|")
If (p% > 0) Then
    s$ = Left$(s$, p% - 1)
End If

buf$ = s$

Call DefErrPop
End Sub

Sub ZeigePlusX()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("ZeigePlusX")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
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
Dim iRow%
Dim TaskId&
Dim pzn$

If (AktTabIndex% < 3) Or (AktTabIndex% = 31) Then Call DefErrPop: Exit Sub

iRow% = (AktTabIndex% - 3) \ 4
If (iRow% >= AnzRezeptArtikel%) Then Call DefErrPop: Exit Sub


pzn$ = RezArtikel(iRow%).pzn
If (Val(pzn$) = 0) Then Call DefErrPop: Exit Sub

TaskId& = Shell("ABDAinfo.exe " + Format(pzn, "00000000"), vbNormalFocus)

Call DefErrPop
End Sub

Sub ModiZeile(s$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("ModiZeile")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
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
Dim ind%, ind2%
Dim s2$
  
ind% = InStr(s$, "#")
If (ind%) Then
    ind2% = InStr(ind% + 1, s$, "#")
    If (ind2%) Then
        s$ = Left$(s$, ind% - 1) + Mid$(s$, ind2% + 1)
    End If
End If

ind% = InStr(s$, "\")
If (ind%) Then
    s$ = Trim(Left$(s$, ind% - 1)) + vbTab + Trim(Mid$(s$, ind% + 1))
End If

Call DefErrPop
End Sub

Sub ShowNichtInTaxe()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("ShowNichtInTaxe")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
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
Dim i%, j%, ShowBox%

For i% = 0 To (AnzRezeptArtikel% - 1)
    With RezArtikel(i%)
        If (.NichtInTaxe) Then
            ShowBox% = True
            
            For j% = 0 To UBound(SonderPzn$)
                If (Left$(SonderPzn$(j%), 8) = .pzn) Then
                    ShowBox% = False
                    Exit For
                End If
            Next j%
            
            If (ShowBox%) Then
                'BTM, BeschKosten, Noctu
                If (Val(.pzn) = 8150006) Or (Val(.pzn) = 2567001) Or (Val(.pzn) = 2567018) Then  ' or (.pzn = "8150012")
                    ShowBox% = False
                End If
            End If
            
            If (ShowBox%) Then
                Call MessageBox("Artikel nicht in Taxe. Erstattung fraglich.", vbInformation, .pzn + "  " + .text)
            End If
        End If
    End With
Next i%
    
Call DefErrPop
End Sub

Sub RezHistorieInit(iForm As Object)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("RezHistorieInit")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
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

Set ParentForm = iForm
    
Call DefErrPop
End Sub

Sub RezHistorieExit()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("RezHistorieExit")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
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

Set ParentForm = frmAction
    
Call DefErrPop
End Sub

Sub DruckeRezept()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("DruckeRezept")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
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
Dim i%, erg%, NeuesRezept%, AbInd%
Dim l&
Dim h$, h2$, sVerfügbarkeit$

If (Selbsterklaerung) Then
    If (DateDiff("d", "01.07.2025", VerkDatum) > 0) Then
        If (MessageBox("Meldung an den NNF-Fonds elektronisch übertragen?", vbQuestion Or vbYesNo Or vbDefaultButton1, "Selbsterklärung") = vbYes) Then
            Call SelbsterklaerungAbrechnung
            Call DefErrPop: Exit Sub
        End If
    End If
End If

If (SonderBelegRezept% > 0) Then
    If (Len(Trim(RezKundenInfo$(3))) = 0) Then  '<>12
        Call MessageBox("Achtung: Ausdruck nicht durchführbar!" + vbCrLf + vbCrLf + "Versicherungs-Nummer des Patienten fehlt!", vbCritical)
        Call DefErrPop: Exit Sub
    End If
End If

BenutzerNr = 0
If (BenutzerSignatur%) Then
    BenutzerNr% = HoleBenutzerSignatur
    If (BenutzerNr <= 0) Then
        Call MessageBox("Achtung: Ausdruck nicht durchführbar!" + vbCrLf + vbCrLf + "Signatur nicht korrekt!", vbCritical)
        Call DefErrPop: Exit Sub
    End If
End If

HochFormatDruck = False
If (PrivatRezept) Then
    If (FormatPrivatRezept) Then
        Dim MsgBoxButtons%, MsgBoxErg%
        MsgBoxButtons = vbYesNoCancel Or vbQuestion
        If PrivRezDruckHoch Then
        Else
            MsgBoxButtons = MsgBoxButtons Or vbDefaultButton2
        End If
        MsgBoxErg = MsgBox("Ausdruck als HOCHFORMAT? ", MsgBoxButtons)
        If (MsgBoxErg = vbCancel) Then
            Call DefErrPop: Exit Sub
        End If
        HochFormatDruck = (MsgBoxErg = vbYes)
    End If
End If

NeuesRezept% = False
If (RezNr$ = "") Then
    NeuesRezept% = True
    
    h$ = "00001"
    l& = GetPrivateProfileString("Rezeptkontrolle", "NeueRezepte", h$, h$, 6, INI_DATEI)
    h$ = Left$(h$, l&)
    lRezNr& = Val(h$)
    RezNr$ = Format(lRezNr&, "00000")

    h$ = Format(lRezNr& + 1, "0")
    l& = WritePrivateProfileString("Rezeptkontrolle", "NeueRezepte", h$, INI_DATEI)
End If

AbInd = 0
If (Val(RezArtikel(0).pzn) = 9999643) Then
    AbInd = 1
End If

If (Left(WVORabattArtikel, 3) <> "111") Then
    For i% = UBound(RezArtikel) - 1 To AbInd Step -1
        RezArtikel(i% + 1) = RezArtikel(i%)
    Next i%
    Call InitRezeptArtikel(AbInd)
    Call InitKontrollStop(AbInd)
    AnzRezeptArtikel% = AnzRezeptArtikel% + 1
    RezArtikel(AbInd).pzn = SONDERPZN_WVO
    RezArtikel(AbInd).text = "WVO RabattArtikel"
    RezArtikel(AbInd).Faktor = Val(Left(WVORabattArtikel, 3))
    RezArtikel(AbInd).Preis = 0#
    RezArtikel(AbInd).flag = 0
    RezArtikel(AbInd).Zuz = 0
End If

AMverfügbarkeitsPzn = 0
sVerfügbarkeit$ = CheckVerfügbarkeit
If (sVerfügbarkeit$ <> "111") Then
    For i% = UBound(RezArtikel) - 1 To AbInd Step -1
        RezArtikel(i% + 1) = RezArtikel(i%)
    Next i%
    Call InitRezeptArtikel(AbInd)
    Call InitKontrollStop(AbInd)
    AnzRezeptArtikel% = AnzRezeptArtikel% + 1
    RezArtikel(AbInd).pzn = "02567024"
    RezArtikel(AbInd).text = "Nichtverfügbarkeit von AM"
    RezArtikel(AbInd).Faktor = Val(sVerfügbarkeit$)
    RezArtikel(AbInd).Preis = 0#
    RezArtikel(AbInd).flag = 0
    RezArtikel(AbInd).Zuz = 0
    
    AMverfügbarkeitsPzn = True
End If

If (BtmRezept) Then
'    BtmRezeptAlt = (iMsgBox("Handelt es sich um ein NEUES BTM-Rezept-Formular ?", vbYesNo Or vbDefaultButton1 Or vbQuestion, "Ausdruck BTM-Rezept") <> vbYes)
    BtmRezeptAlt = 0
End If

For i = 0 To UBound(HmFaktor)
    HmNummer$(i) = ""
    HmFaktor%(i) = 0
    HmStückPreis#(i) = 0
Next i
    
If (IstDosDrucker%) Then
    erg% = DosRezeptDruck%
Else
    Call InitRezeptDruck
    If (eRezeptTaskId = "") Then
        Call RezeptDruck
    End If
    erg% = True
End If
If (erg%) Then
    If (Selbsterklaerung) Then
    Else
        If (NeuesRezept%) Then
            h$ = "9999998" + Format(lRezNr&, "00000")
            Call PruefZiffer(h$)
            RezNr$ = h$
        End If
        
        Call WriteRezeptSpeicher
        Call CreateGdiObjects
        
        If (eRezeptTaskId <> "") Then
            If (FDok = 0) Then
                Set FD_OP = New TI_Back.Fachdienst_OP
                If (FD_OP.Init(FD_OP.NetWindowHandle(ParentForm.hWnd))) Then
                    FDok = True
                End If
            End If

            Dim lPzn&
            For i% = 0 To (AnzRezeptArtikel% - 1)
                lPzn = Val(RezArtikel(i%).pzn)
                If (lPzn = 9999637) Or (lPzn = 6461110) Then  'Beschaffungskosten 'Botendienst
                    eRezeptSonderPZN = True
                    Exit For
                End If
            Next i
       
            bEinzelErgebnis = True
        
            Dim eRezeptStatusWerte2$
            AMverfügbarkeit$ = CheckVerfügbarkeit
            eRezeptStatusWerte2 = RezArtikel(0).pzn + Mid(AMverfügbarkeit, 1, 1) + Format(RezArtikel(0).Preis, "0.00") + Format(RezArtikel(0).Zuz, "0.00") + VerkDatum
            'MsgBox (eRezeptStatusWerte + vbCrLf + eRezeptStatusWerte2)
            If (eRezeptStatusWerte <> eRezeptStatusWerte2) Or (eRezeptSonderPZN) Then
                Call eRezeptSpeichern
                frm_eRezeptEdit.Show 1
                If (eRezeptGespeichert) Then
                End If
            End If

            Dim eRezept2 As New TI_Back.eRezept_OP
            Dim bErgebnis As Boolean
            bErgebnis = True
            Call eRezept2.New2(eRezeptTaskId)
            If (eRezeptOpStatus < 3) Then
                If (FD_OP.TI_RezeptAbgeben(eRezept2.TaskId, eRezept2.Secret)) Then
                    Call MessageBox("eRezept erfolgreich DISPENSIERT !", vbInformation)
                Else
                    Call MessageBox("Probleme beim Dispensieren !", vbCritical)
                    bErgebnis = False
                End If
            End If
            If (bErgebnis) Then
                Set eRezept2 = New TI_Back.eRezept_OP
                Call eRezept2.New2(eRezeptTaskId)
                If (FD_OP.TI_RezeptAbrechnen(eRezept2, False)) Then
    '                    MsgBox (IIf(SollStatus = 2, "Vorabprüfung: ", "Einreichung: ") + CStr(VorabPruefung(sTaskId, (SollStatus = 2))))
                    bErgebnis = VorabPruefung(eRezeptTaskId, False)
                Else
                    Call MessageBox("Problem beim Erstellen der Abgabedaten", vbCritical)
                End If
            End If
        End If
    End If
    ParentForm.cmdEsc(0).Value = True
    
    If (para.MARS) And (PrivatRezept) Then
'        h = "\user\Privat Rezept Beilage.pdf"
'
'        h2 = String(MAX_FILENAME_LEN, 32)
'        'Retrieve the name and handle of the executable, associated with this file
'        i = FindExecutable(h, vbNullString, h2)
'        If i > 32 Then
'        '                MsgBox Left$(s2, InStr(s2, Chr$(0)) - 1)
'
'            Call ShellExecute(ParentForm.hWnd, "Print", h, 0&, 0&, 0)
'        Else
'        '                MsgBox "No association found !"
'        End If
        
        Call DruckKundenListeKopf
    End If
End If
    
Call DefErrPop
End Sub

Sub DruckKundenListeKopf(Optional MitSpalten% = True)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("DruckKundenListeKopf")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
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
Dim RechtsX%
Dim l&, he&, i%, pos%, x%, x2%, y%, gesBreite&
Dim KuNr&
Dim h$, heute$, SeitenNr$, sTitel$, sPara$, sAnrede$, sDruckerOrg$
Dim KundenDB As Database
Dim KundenRec As Recordset

KuNr = Val(RezKundenInfo$(4))
If (KuNr <= 0) Then
    Call DefErrPop: Exit Sub
End If

Dim hPrinter As Printer
sDruckerOrg = Printer.DeviceName
For Each hPrinter In Printers
    If (UCase(StandardDrucker) = UCase(hPrinter.DeviceName)) Then
        Set Printer = hPrinter
        Exit For
    End If
Next

heute$ = Format(Day(Date), "00") + "-"
heute$ = heute$ + Format(Month(Date), "00") + "-"
heute$ = heute$ + Format(Year(Date), "0000")

Printer.PaintPicture LoadPicture(CurDir + "\DocMorris.jpg"), Printer.ScaleWidth - 5100, 600
With Printer
    .Font.Name = "Arial"
    .Font.Size = 8
    
    x = RandAbstandTwips(20, 0)
    .CurrentX = x
    .CurrentY = RandAbstandTwips(40, 1)
    Printer.Print "Versandapotheke DocMorris     52098 Aachen     Deutschland"
    
    .Font.Name = "Arial"
    .Font.Size = 10
    .CurrentY = RandAbstandTwips(50, 1)
    
    Dim KundenSQL%
    Dim KundenAdoRec As New ADODB.Recordset
    Dim KundenAdoDB As New clsKundenDB

    KundenSQL = KundenAdoDB.DBvorhanden
    If (KundenSQL) Then
        KundenSQL = KundenAdoDB.OpenDB

        SQLStr$ = "SELECT * FROM Kunden WHERE KundenNr =" + Str$(KuNr)
        On Error Resume Next
        KundenAdoRec.Close
        Err.Clear
        On Error GoTo DefErr
        KundenAdoRec.Open SQLStr, KundenAdoDB.ActiveConn
        If (KundenAdoRec.EOF = False) Then
            For i% = 0 To 3
                If (i% = 0) Then
'                    If (CheckNullByte(KundenAdoRec!Geschlecht) = 1) Then
'                        sAnrede = "Sehr geehrter "
'                    Else
'                        sAnrede = "Sehr geehrte "
'                    End If
                    sAnrede = "Sehr geehrte(r) "
                    h$ = Trim(CheckNullStr(KundenAdoRec!Anrede))
                    If (h$ <> "") Then
                        h$ = h$ + " "
                    End If
                    h$ = h$ + Trim(CheckNullStr(KundenAdoRec!Titel))
                    sAnrede = sAnrede + h
                ElseIf (i% = 1) Then
                    h$ = Trim(CheckNullStr(KundenAdoRec!Vorname))
                    If (h$ <> "") Then
                        h$ = h$ + " "
                    End If
                    h$ = h$ + Trim(CheckNullStr(KundenAdoRec!Name))
                    sAnrede = sAnrede + " " + h
                ElseIf (i% = 2) Then
                    h$ = Trim(CheckNullStr(KundenAdoRec!strasse))
                    If (h$ <> "") Then
                        h$ = h$ + " "
                    End If
                    h$ = h$ + Trim(CheckNullStr(KundenAdoRec!HausNr))
                Else
                    h$ = Trim(CheckNullStr(KundenAdoRec!PLZ))
                    If (h$ <> "") Then
                        h$ = h$ + " "
                    End If
                    h$ = h$ + Trim(CheckNullStr(KundenAdoRec!Ort))
                End If
                If (h$ <> "") Then
                    .CurrentX = x
                    Printer.Print h$;
                End If

                Printer.Print
            Next i%
        End If
        KundenAdoDB.CloseDB
    Else
        h$ = "Kunden.mdb"
        Set KundenDB = OpenDatabase(h$, False, True)
        SQLStr$ = "SELECT * FROM Kunden WHERE KundenNr =" + Str$(KuNr)
        Set KundenRec = KundenDB.OpenRecordset(SQLStr$)
        If (KundenRec.EOF = False) Then
            For i% = 0 To 3
                If (i% = 0) Then
'                    If (CheckNullByte(KundenRec!Geschlecht) = 1) Then
'                        sAnrede = "Sehr geehrter "
'                    Else
'                        sAnrede = "Sehr geehrte "
'                    End If
                    sAnrede = "Sehr geehrte(r) "
                    h$ = Trim(CheckNullStr(KundenRec!Anrede))
                    If (h$ <> "") Then
                        h$ = h$ + " "
                    End If
                    h$ = h$ + Trim(CheckNullStr(KundenRec!Titel))
                    sAnrede = sAnrede + h
                ElseIf (i% = 1) Then
                    h$ = Trim(CheckNullStr(KundenRec!Vorname))
                    If (h$ <> "") Then
                        h$ = h$ + " "
                    End If
                    h$ = h$ + Trim(CheckNullStr(KundenRec!Name))
                    sAnrede = sAnrede + " " + h
                ElseIf (i% = 2) Then
                    h$ = Trim(CheckNullStr(KundenRec!strasse))
                    If (h$ <> "") Then
                        h$ = h$ + " "
                    End If
                    h$ = h$ + Trim(CheckNullStr(KundenRec!HausNr))
                Else
                    h$ = Trim(CheckNullStr(KundenRec!PLZ))
                    If (h$ <> "") Then
                        h$ = h$ + " "
                    End If
                    h$ = h$ + Trim(CheckNullStr(KundenRec!Ort))
                End If
                If (h$ <> "") Then
                    .CurrentX = x
                    Printer.Print h$;
                End If
        
                Printer.Print
            Next i%
            
        End If
        KundenDB.Close
    End If
    
    .CurrentX = RandAbstandTwips(175, 0)
    Printer.Print heute$
          
    For i = 1 To 6
        Printer.Print
    Next i
    
    .CurrentX = x
    Printer.Print "Privatrezept-Kopien"
    
    For i = 1 To 2
        Printer.Print
    Next i
    
    .CurrentX = x
    Printer.Print sAnrede
    
    For i = 1 To 2
        Printer.Print
    Next i
    
    .CurrentX = x
    Printer.Print "Sie haben auf Privatrezept verordnete Medikamente bei der Apotheke DocMorris bezogen."
    .CurrentX = x
    Printer.Print "Für Ihr Vertrauen danken wir nochmals herzlich."
    
    For i = 1 To 2
        Printer.Print
    Next i
    
    .CurrentX = x
    Printer.Print "Wie Ihnen unsere Mitarbeiterin im Videogespräch bereits mitgeteilt hat, sind wir als niederländische"
    .CurrentX = x
    Printer.Print "Apotheke dazu verpflichtet, alle Original-Rezepte aufzubewahren. Deshalb erhalten Sie heute wie"
    .CurrentX = x
    Printer.Print "vereinbart je zwei Kopien pro Original-Rezept. Rezeptkopien werden von allen deutschen"
    .CurrentX = x
    Printer.Print "Krankenversicherungen für die Abrechnung Ihrer Medikamente akzeptiert."
    
    For i = 1 To 2
        Printer.Print
    Next i
    
    .CurrentX = x
    Printer.Print "Unsere kostenfreie Telefonnummer:   0800 660 5555 6"
    
    For i = 1 To 5
        Printer.Print
    Next i
    .CurrentX = x
    Printer.Print "Herzliche Grüße von"
    .CurrentX = x
    Printer.Print "Ihrer Apotheke DocMorris"
        
    
    .Font.Name = "Arial"
    .Font.Size = 8
    
    y = .ScaleHeight - RandAbstandTwips(40, 1)
    .CurrentX = x
    .CurrentY = y
    Printer.Print "Postanschrift:"
    .CurrentX = x
    Printer.Print "Versandapotheke";
    x2 = .CurrentX
    Printer.Print
    .CurrentX = x
    Printer.Print "DocMorris"
    .CurrentX = x
    Printer.Print "52098 Aachen"
    .CurrentX = x
    Printer.Print "Deutschland"
    
    x = x2 + 390
    .CurrentX = x
    .CurrentY = y
    Printer.Print "Bank: Hypovereinsbank Stuttgart"
    .CurrentX = x
    Printer.Print "D&W Mailorder Service B.V."
    .CurrentX = x
    Printer.Print "Konto-Nr.: 611022468 BLZ:60020290";
    x2 = .CurrentX
    Printer.Print
    .CurrentX = x
    Printer.Print "IBAN: DE11 6002 0290 0611 0224 68"
    .CurrentX = x
    Printer.Print "BIC: HYVEDEMM473"
    
    x = x2 + 390
    .CurrentX = x
    .CurrentY = y
    Printer.Print "Sitz der Gesellschaft:";
    x2 = .CurrentX
    Printer.Print
    .CurrentX = x
    Printer.Print "DocMorris N.V."
    .CurrentX = x
    Printer.Print "Avantisallee 152"
    .CurrentX = x
    Printer.Print "6422 RA Heerlen"
    .CurrentX = x
    Printer.Print "Niederlande"

    x = x2 + 390
    .CurrentX = x
    .CurrentY = y
    Printer.Print "Vorstand:"
    .CurrentX = x
    Printer.Print "Olaf Heinrich (Vorsitz)"
    .CurrentX = x
    Printer.Print "Prof. Dr. Christian Franken";
    x2 = .CurrentX
    Printer.Print
    .CurrentX = x
    Printer.Print "Max Müller"
    .CurrentX = x
    Printer.Print "Michael Veigel"

    x = x2 + 390
    .CurrentX = x
    .CurrentY = y
    Printer.Print "USt-IdNr.:"
    .CurrentX = x
    Printer.Print "DE210797605"
    .CurrentX = x
    Printer.Print "NL-USt-Nr.:"
    .CurrentX = x
    Printer.Print "NL8086.37.642.B.01";
    x2 = .CurrentX
    Printer.Print
    .CurrentX = x
    Printer.Print "KVK-Nr.: 14066093"
    
    Printer.EndDoc
End With

For Each hPrinter In Printers
    If (UCase(sDruckerOrg) = UCase(hPrinter.DeviceName)) Then
        Set Printer = hPrinter
        Exit For
    End If
Next

Call DefErrPop
End Sub

Function MachLinkenRand%(SollRandInMM%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("MachLinkenRand%")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
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
Dim LRandInPixel&, ORandInPixel&
Dim PrTwipsPerPixelX!, PrTwipsPerpixelY!, LRandInMM!, ORandInMM!

'Randbreiten in Pixel
LRandInPixel& = GetDeviceCaps(Printer.hdc, PHYSICALOFFSETX)
ORandInPixel& = GetDeviceCaps(Printer.hdc, PHYSICALOFFSETY)

'Auflösung Drucker
PrTwipsPerPixelX = Printer.TwipsPerPixelX
PrTwipsPerpixelY = Printer.TwipsPerPixelY

'Randbreiten in Millimeter
LRandInMM! = Printer.ScaleX(LRandInPixel& * PrTwipsPerPixelX!, vbTwips, vbMillimeters)
ORandInMM! = Printer.ScaleY(ORandInPixel& * PrTwipsPerpixelY!, vbTwips, vbMillimeters)

ret% = 0
If (SollRandInMM% > LRandInMM!) Then
    ret% = Printer.ScaleX(SollRandInMM% - LRandInMM!, vbMillimeters, vbTwips)
End If

MachLinkenRand% = ret%
    
Call DefErrPop
End Function

Function RandAbstandTwips%(SollRandInMM%, WaagrechtOderSenkrecht%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("MachLinkenRand%")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
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
Dim LRandInPixel&, ORandInPixel&
Dim PrTwipsPerPixelX!, PrTwipsPerpixelY!, LRandInMM!, ORandInMM!

'Randbreiten in Pixel
LRandInPixel& = GetDeviceCaps(Printer.hdc, PHYSICALOFFSETX)
ORandInPixel& = GetDeviceCaps(Printer.hdc, PHYSICALOFFSETY)

'Auflösung Drucker
PrTwipsPerPixelX = Printer.TwipsPerPixelX
PrTwipsPerpixelY = Printer.TwipsPerPixelY

'Randbreiten in Millimeter
LRandInMM! = Printer.ScaleX(LRandInPixel& * PrTwipsPerPixelX!, vbTwips, vbMillimeters)
ORandInMM! = Printer.ScaleY(ORandInPixel& * PrTwipsPerpixelY!, vbTwips, vbMillimeters)

ret% = 0
If (WaagrechtOderSenkrecht = 0) Then
    If (SollRandInMM% > LRandInMM!) Then
        ret% = Printer.ScaleX(SollRandInMM% - LRandInMM!, vbMillimeters, vbTwips)
    End If
Else
    If (SollRandInMM% > ORandInMM!) Then
        ret% = Printer.ScaleY(SollRandInMM% - ORandInMM!, vbMillimeters, vbTwips)
    End If
End If

RandAbstandTwips = ret%
    
Call DefErrPop
End Function

Sub RepaintBtmGebühr()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("RepaintBtmGebühr")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
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
        
If (BtmRezept%) Then
    RezAddGesamtBrutto$(0) = sBrutto$
    Call EchtKurzInfo
End If

Call DefErrPop
End Sub

Function CalcZuzaNeu#(gPreis#, Optional HimiVerbrauch% = 0, Optional iFaktor% = 1)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("CalcZuzaNeu#")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
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
Dim dPreis#

dPreis# = gPreis# * 10#
If (dPreis# < 500#) And (HimiVerbrauch% = 0) Then
    dPreis# = 500#
'ElseIf (dPreis# > 1000#) Then  'ab 4.0.65 nächste Zeile eingebaut
ElseIf (dPreis# > 1000#) And (HimiVerbrauch% = 0) Then
    dPreis# = 1000#
End If

CalcZuzaNeu# = CDbl(Int(dPreis# + 0.501)) / 100#

If (HimiVerbrauch%) Then
    If (RezeptHolenZuz# > 0) Then
        CalcZuzaNeu# = RezeptHolenZuz# * iFaktor
    ElseIf ((CalcZuzaNeu# * iFaktor%) > 10#) Then
        If (MessageBox("Soll die Zuzahlung auf 10 EURO begrenzt werden?", vbQuestion Or vbYesNo Or vbDefaultButton2) = vbYes) Then
'            CalcZuzaNeu# = 10 / iFaktor%     'ab 4.0.65 nächste Zeile eingebaut
            CalcZuzaNeu = 10
        End If
    End If
End If

Call DefErrPop
End Function

Sub DruckTaxierung(x%, y%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("DruckTaxierung")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
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
Dim lf As LOGFONT
Dim result As Long
Dim hOldfont As Long
Dim hPrintDc As Long
Dim hFont As Long
Dim i%, j%, x1%, y1%, Y2%, AbRow%, BisRow%
Dim lDummy&
Dim dPreis#
Dim txt$

'Y% = 2977
'MsgBox (Str(x%) + Str$(Y%))

'x% = CLng(PrinterScaleX(CLng(x%), vbTwips, vbPixels))
'y% = CLng(PrinterScaleY(CLng(y%), vbTwips, vbPixels))

With frmAction.picTaxierungsDruck
    .Font.Name = Printer.Font.Name
    .Font.Size = Printer.Font.Size
End With

hPrintDc = Printer.hdc

lf.lfEscapement = 900 ' 1800
lf.lfHeight = (DESIREDFONTSIZE * -20) / Printer.TwipsPerPixelY
hFont = CreateFontIndirect(lf)
hOldfont = SelectObject(hPrintDc, hFont)

'MAG_SPEICHER% = FileOpen("MagTax" + Trim(para.User), "RW", "B")
'Seek #MAG_SPEICHER%, ((MagSpeicherIndex% - 1) * CLng(Len(TaxierRec))) + 1
'For i% = 1 To 100
'    Get #MAG_SPEICHER%, , TaxierRec
'    If (EOF(MAG_SPEICHER%)) Then Exit For
'    If (TaxierRec.flag = 255) Then Exit For
'
'    txt$ = Format(TaxierRec.ActPreis, "0.00")
'    y1% = CLng(PrinterScaleY(CLng(y%) + Printer.TextWidth(txt$), vbTwips, vbPixels))
'    lDummy& = SelectObject(hPrintDc, hFont)
'    result = TextOut(hPrintDc, x%, y1%, txt$, Len(txt$))
'    x% = x% + 200
'Next i%
'Close (MAG_SPEICHER%)

'Load frmTaxieren
For j% = 0 To 1
    If (j% = 0) Then
        AbRow% = 1
        BisRow% = frmTaxieren.flxTaxieren.Rows - 1
    Else
        Printer.Line (x%, y% + 60)-(x%, y% + 690)
        x% = x% + 90
        AbRow% = 0
        BisRow% = frmTaxieren.flxTaxSumme.Rows - 1
    End If
    For i% = AbRow% To BisRow%
        If (j% = 0) Then
            dPreis# = xVal(frmTaxieren.flxTaxieren.TextMatrix(i%, 0))
        Else
            If (i% = BisRow%) Then
                Printer.Line (x%, y% + 60)-(x%, y% + 690)
                x% = x% + 90
            End If
            dPreis# = xVal(frmTaxieren.flxTaxSumme.TextMatrix(i%, 0))
        End If
        If (dPreis# > 0) Then
            txt$ = Format(dPreis#, "0.00")
            x1% = CLng(PrinterScaleX(CLng(x%), vbTwips, vbPixels))
'            Printer.ScaleMode = vbTwips
            Y2% = frmAction.picTaxierungsDruck.TextWidth(txt$)
'            y1% = CLng(PrinterScaleY(CLng(Y%) + Printer.TextWidth(txt$), vbTwips, vbPixels))
            y1% = CLng(PrinterScaleY(CLng(y%) + Y2%, vbTwips, vbPixels))
'            MsgBox (Str$(y2%))
            lDummy& = SelectObject(hPrintDc, hFont)
            result = TextOut(hPrintDc, x1%, y1%, txt$, Len(txt$))
            x% = x% + 240
        End If
    Next i%
Next j%
frmTaxieren.cmdEsc.Value = True

On Error Resume Next
Close #MAG_SPEICHER%
Err.Clear
On Error GoTo DefErr


result = SelectObject(hPrintDc, hOldfont)
result = DeleteObject(hFont)

Printer.FontBold = False

Call DefErrPop
End Sub

Sub MakeActKkasse(sIk$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("MakeActKkasse")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
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
Dim h$, SQLStr$

Call PaintRezeptBox(59)
Call PaintRezeptWert(59, sIk$)

'ActKkasse$ = sIk$
If (ActKkasse$ = "") Then
    KasseRec.index = "Unique"
    KasseRec.Seek "=", sIk$
    If (KasseRec.NoMatch) Or (KasseRec.EOF) Then
        SQLStr$ = "SELECT * FROM kKassen WHERE IK = " + Str$(Val(sIk$))
'        Set KkassenRec = kKassenDB.OpenRecordset(SQLStr$)
'        If (KkassenRec.RecordCount > 0) Then
        FabsErrf = kKassenDB.OpenRecordset(KkassenRec, SQLStr)
        If Not (KkassenRec.EOF) Then
            h$ = Trim(KkassenRec!Name)
            Call PaintRezeptBox(34)
            Call PaintRezeptWert(34, h$)
        End If
    Else
        h$ = Trim(KasseRec!Name)
        Call PaintRezeptBox(34)
        Call PaintRezeptWert(34, h$)
    End If
End If

Call DefErrPop
End Sub

Sub LadeOptionenAbgabeKosten(iFlex As Object)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("LadeOptionenAbgabeKosten")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
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
Dim dPreis#
Dim h2$

'With frmRezkOptionen.flxOptionen1(5)
With iFlex
    For i% = 1 To UBound(AbgabeKosten)
        dPreis# = AbgabeKosten(i%).kosten
        If (dPreis# = 0#) Then
            h2$ = ""
        Else
            h2$ = Format(dPreis#, "0.00")
        End If
        .TextMatrix(i, 0) = AbgabeKosten(i%).Name
        .TextMatrix(i, 1) = h2$
    Next i%
End With

Call DefErrPop
End Sub

Function CheckRabattAutIdem(pzn As Long, Optional art As Integer = 0) As Boolean
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("CheckRabattAutIdem")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
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
Dim SQLStr As String, SqlSonderregelNr As String, SQLStr2$
Dim krec As New ADODB.Recordset, tmpRec As New ADODB.Recordset, iRec As New ADODB.Recordset
Dim Sonderregelnr As Long, BtmDerivatMenge As Double, ok As Boolean, StdMenge As Long, Ngr As String
'Dim TaxeTmpRec As Recordset
'Dim tRec2 As Recordset
Dim TaxeTmpRec As New ADODB.Recordset
Dim tRec2 As New ADODB.Recordset
'Dim bmk As Variant
'
'On Error Resume Next
'bmk = TaxeRec.Bookmark
'On Error GoTo DefErr

'TaxeRec.index = "PZN"
'TaxeRec.Seek "=", CStr(Pzn)
'If Not TaxeRec.NoMatch Then
                            
    BtmDerivatMenge# = 0
    If (BtmDerivatMengeIgnorieren) Then
    Else
        BtmDerivatMenge = dCheckNull(TaxeRec!BtmDerivatMenge)
    End If
    If (BtmDerivatMenge >= 0.01) Then
        SQLStr$ = "SELECT * FROM M2Inhalt WHERE M2Nr = " + Mid(TaxeRec!M2, 2, 9)
    '                Set AutIdemRec = AutIdemDB.OpenRecordset(SQLStr$)
        On Error Resume Next
        tRec2.Close
        Err.Clear
        On Error GoTo DefErr
        tRec2.Open SQLStr, AutIdemDB.ActiveConn
    '                            If Not tRec2.NoMatch Then
        If Not (tRec2.EOF) Then
            If (InStr(UCase(CheckNullStr(tRec2!Inhalt)), "HYDROMORPHON") > 0) Then
                BtmDerivatMenge = 0
            End If
        End If
        tRec2.Close
    End If
  
  
  StdMenge = TaxeRec!StdMenge
  Ngr = TaxeRec!NGrösse
'End If

Sonderregelnr = 0
SqlSonderregelNr = ""
CheckRabattAutIdem = False
If art = 0 Then
    SQLStr$ = "SELECT * FROM AutIdemSonderregelKrankenkasse WHERE KbvNr = " + CStr(AutIdemKbvNr)
'    Set krec = AutIdemDB.OpenRecordset(SQLStr)
'    If (krec.RecordCount > 0) Then
    krec.Open SQLStr, AutIdemDB.ActiveConn
    If Not (krec.EOF) Then
        krec.MoveFirst
        Do While Not krec.EOF
            Sonderregelnr = dCheckNull(krec!AutIdemSonderregelNr)
            If (Sonderregelnr > 0) Then
                If (SqlSonderregelNr = "") Then
                    SqlSonderregelNr = "SELECT * FROM AutIdemSonderregelArtikel WHERE ("
                Else
                    SqlSonderregelNr = SqlSonderregelNr$ + " OR"
                End If
                SqlSonderregelNr = SqlSonderregelNr + " AutIdemSonderregelNr = " + Format(dCheckNull(krec!AutIdemSonderregelNr), "##0")
            End If
            krec.MoveNext
        Loop
        krec.Close
        Set krec = Nothing
        If SqlSonderregelNr > "" Then
            SQLStr$ = SqlSonderregelNr + ")"
            SQLStr$ = SQLStr$ + " AND NrKz='P' AND ZoNr=" + CStr(pzn)
'            Set tmpRec = AutIdemDB.OpenRecordset(SQLStr$)
'            If (tmpRec.RecordCount > 0) Then
            tmpRec.Open SQLStr, AutIdemDB.ActiveConn
            If Not (tmpRec.EOF) Then
                CheckRabattAutIdem = False
            Else
                SQLStr$ = "SELECT * FROM PZNSUBST WHERE PZN = " + CStr(pzn)
'                Set iRec = AutIdemDB.OpenRecordset(SQLStr$)
                iRec.Open SQLStr, AutIdemDB.ActiveConn
                Do
                    If (iRec.EOF) Then
                        Exit Do
                    End If
                    SQLStr$ = SqlSonderregelNr + ")"
                    SQLStr$ = SQLStr$ + " AND NrKz='P' AND ZoNr=" + CStr(iRec!Substpzn)
                    
                    ok = True
                    If BtmDerivatMenge <> 0 Then
                        SQLStr2$ = "SELECT * FROM TAXE WHERE PZN = " + CStr(iRec!Substpzn)
'                        Set tRec2 = TaxeDB.OpenRecordset(SQLStr2$)
                        On Error Resume Next
                        tRec2.Close
                        Err.Clear
                        On Error GoTo DefErr
                        tRec2.Open SQLStr2, taxeAdoDB.ActiveConn
                        If Not (tRec2.EOF) Then
                            If dCheckNull(tRec2!BtmDerivatMenge) <> BtmDerivatMenge Then ok = False
                        End If
'                        TaxeRec.Seek "=", iRec!Substpzn
'                        If Not TaxeRec.NoMatch Then
'                            If dCheckNull(TaxeRec!BtmDerivatMenge) <> BtmDerivatMenge Then ok = False
'                        End If
                    End If
                    If ok Then
'                        Set tmpRec = AutIdemDB.OpenRecordset(SQLStr$)
'                        If (tmpRec.RecordCount > 0) Then
                        tmpRec.Close
                        tmpRec.Open SQLStr, AutIdemDB.ActiveConn
                        If Not (tmpRec.EOF) Then
                          If (pzn <> iRec!Substpzn) Then
                              CheckRabattAutIdem = True
                          End If
                          Exit Do
                      End If
                    End If
                    iRec.MoveNext
                Loop
                iRec.Close
            End If
            tmpRec.Close
        End If
    End If
                              
'    CheckRabattAutIdem = False
ElseIf art = 1 Then
    If TaxeRec!Identgruppe <> 0 Then
        SQLStr = "SELECT * FROM TAXE WHERE IDENTGRUPPE = " + CStr(TaxeRec!Identgruppe)
        SQLStr = SQLStr + " and Einheit = '" + TaxeRec!einheit + "'"
        SQLStr = SQLStr + " AND (StdMenge=" + CStr(StdMenge) + " OR NGrösse='" + Ngr + "')"
        If BtmDerivatMenge >= 0.01 Then SQLStr = SQLStr + " AND BtmDerivatMenge=" + Str(BtmDerivatMenge)
        
        SQLStr = SQLStr + " ORDER BY MENGE,VK"
'        Set TaxeTmpRec = TaxeDB.OpenRecordset(SQLStr)
        On Error Resume Next
        TaxeTmpRec.Close
        Err.Clear
        On Error GoTo DefErr
        TaxeTmpRec.Open SQLStr, taxeAdoDB.ActiveConn
        
        If Not (TaxeTmpRec.EOF) Then
            TaxeTmpRec.MoveLast
            If (TaxeTmpRec.RecordCount > 1) Then
                TaxeTmpRec.MoveFirst
            End If
            
            Do While Not TaxeTmpRec.EOF
                If TaxeTmpRec!ArtStatus <> "S" Then
                    If IstRabattAutIdem(PznString(TaxeTmpRec!pzn)) Then
                        CheckRabattAutIdem = True
                        Exit Do
                    End If
                End If
                TaxeTmpRec.MoveNext
            Loop
        End If
    End If
    
'    CheckRabattAutIdem = True
End If


'On Error Resume Next
'TaxeRec.Bookmark = bmk
'On Error GoTo DefErr
Call DefErrPop
End Function

Function IstRabattAutIdem(pzn As Long) As Boolean
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("IstRabattAutIdem")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
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
Dim tmpRec As New ADODB.Recordset
Dim SqlSonderregelNr As String, SQLStr As String
Dim Sonderregelnr As Long

Sonderregelnr = 0
SqlSonderregelNr = ""
If (AutIdemKbvNr > 0) Then
    SQLStr$ = "SELECT * FROM AutIdemSonderregelKrankenkasse WHERE KbvNr = " + CStr(AutIdemKbvNr)
'    Set tmpRec = AutIdemDB.OpenRecordset(SQLStr)
'    If (tmpRec.RecordCount > 0) Then
    tmpRec.Open SQLStr, AutIdemDB.ActiveConn
    If Not (tmpRec.EOF) Then
        tmpRec.MoveFirst
        Do While Not tmpRec.EOF
            Sonderregelnr = dCheckNull(tmpRec!AutIdemSonderregelNr)
            If (Sonderregelnr > 0) Then
                If (SqlSonderregelNr = "") Then
                    SqlSonderregelNr = "SELECT * FROM AutIdemSonderregelArtikel WHERE ("
                Else
                    SqlSonderregelNr = SqlSonderregelNr + " OR"
                End If
                SqlSonderregelNr = SqlSonderregelNr + " AutIdemSonderregelNr = " + Format(dCheckNull(tmpRec!AutIdemSonderregelNr), "##0")
            End If
            tmpRec.MoveNext
        Loop
    End If
    tmpRec.Close
'    Set tmpRec = Nothing
    
    If (SqlSonderregelNr > "") Then
        SQLStr$ = SqlSonderregelNr + ")"
        SQLStr$ = SQLStr$ + " AND NrKz='P' AND ZoNr=" + CStr(pzn)
'        Set tmpRec = AutIdemDB.OpenRecordset(SQLStr$)
'        If (tmpRec.RecordCount > 0) Then
        tmpRec.Open SQLStr, AutIdemDB.ActiveConn
        If Not (tmpRec.EOF) Then
            IstRabattAutIdem = True
        End If
        tmpRec.Close
        Set tmpRec = Nothing
    End If
End If

'IstRabattAutIdem = False

Call DefErrPop
End Function


Sub MachGebFrei(Optional sName$ = "")
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("MachGebFrei")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
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

GebFrei = True
If (sName <> "") Then
    IkSpezialName = sName
End If

Call DefErrPop
End Sub

Sub SelbsterklaerungDruck()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("SelbsterklaerungDruck")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
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
Dim TaskId&
Dim h$

h = Format(DateAdd("d", -1, "01." + Format(Now, "MM.YYYY")), "MMYY")
Do
    h$ = MyInputBox("Selbsterklärung drucken für Monat (MMJJ): ", "Selbsterklärung", h$)
    h$ = UCase(Trim(h$))
    If (h$ = "") Then
        Exit Do
    ElseIf (iDate("01" + h$) <> 0) Then
        Exit Do
    End If
Loop

If (h = "") Then
    Call DefErrPop: Exit Sub
End If

Call InitRezept
GebFrei% = True
'ActVerordnung% = 2
Selbsterklaerung = True
SchutzMasken = False

SelbstErklaerungHidden = ",1,3,7,11,15,19,23,27,32,33,43,44,45,46,50,51,52,"
SelbstErklaerungTabstop = ",0,6,31,52,"

RezKundenInfo$(1) = "SONDERBELEG"

AnzRezeptArtikel = 1
ind = 0
Call InitRezeptArtikel(ind)
RezArtikel(ind).pzn = "02567768"
RezArtikel(ind).text = ""
RezArtikel(ind).Faktor = 1
RezArtikel(ind).Preis = 0
RezArtikel(ind).Zuz = 0

'VerkDatum = Format(DateAdd("d", -1, "01." + Format(Now, "MM.YYYY")), "DDMMYY")
VerkDatum = Format(DateAdd("d", -1, DateAdd("m", 1, "01." + Left(h, 2) + ".20" + Mid(h, 3, 2))), "DDMMYY")
BoxenWert(52) = "01" + Mid(VerkDatum, 3)
Call PaintRezept
DoEvents

h = "WinVk ANSG" + Mid(VerkDatum, 3, 2)
WinVkId& = Shell(h, vbNormalFocus)

frmAction.tmrAction.Enabled = True

Call DefErrPop
End Sub

Sub SchutzMaskenDruck(index%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("SchutzMaskenDruck")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
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
Dim TaskId&
Dim h$

'h = Format(DateAdd("d", -1, "01." + Format(Now, "MM.YYYY")), "MMYY")
'Do
'    h$ = MyInputBox("Selbsterklärung drucken für Monat (MMJJ): ", "Selbsterklärung", h$)
'    h$ = UCase(Trim(h$))
'    If (h$ = "") Then
'        Exit Do
'    ElseIf (iDate("01" + h$) <> 0) Then
'        Exit Do
'    End If
'Loop
'
'If (h = "") Then
'    Call DefErrPop: Exit Sub
'End If

h = Format(Now, "MMYY")

Call InitRezept
GebFrei% = True
'ActVerordnung% = 2
Selbsterklaerung = True
SchutzMasken = index

SelbstErklaerungHidden = ",3,7,11,15,19,23,27,32,33,43,44,45,46,50,51,52,"
SelbstErklaerungTabstop = ",0,5,31,"

RezKundenInfo$(1) = "SONDERBELEG"

AnzRezeptArtikel = 1
ind = 0
Call InitRezeptArtikel(ind)

Select Case SchutzMasken
    Case 1
        RezArtikel(ind).pzn = "06461245"
        RezArtikel(ind).Preis = 36
    Case 2
        RezArtikel(ind).pzn = "06461297"
        RezArtikel(ind).Preis = 23.4
    Case 3
        RezArtikel(ind).pzn = "06461305"
        RezArtikel(ind).Preis = 39
End Select
RezArtikel(ind).text = ""
RezArtikel(ind).Faktor = 1
RezArtikel(ind).Zuz = 0

'VerkDatum = Format(DateAdd("d", -1, "01." + Format(Now, "MM.YYYY")), "DDMMYY")
VerkDatum = Format(DateAdd("d", -1, DateAdd("m", 1, "01." + Left(h, 2) + ".20" + Mid(h, 3, 2))), "DDMMYY")
BoxenWert(31) = VerkDatum
Call PaintRezept
DoEvents

'h = "WinVk ANSG" + Mid(VerkDatum, 3, 2)
'WinVkId& = Shell(h, vbNormalFocus)
'frmAction.tmrAction.Enabled = True

Call DefErrPop
End Sub

Function CheckAlteTaxe$()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("CheckAlteTaxe$")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
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
Dim dd%
Dim dbTaxeDate As Date
Dim NeuesteTaxe$, EinspielDatumStempel$, h, h2$, ret$

ret = ""

dbTaxeDate = Left(OrgVerkDatum, 2) + "." + Mid(OrgVerkDatum, 3, 2) + ".20" + Mid(OrgVerkDatum, 5, 2)
dd% = Day(dbTaxeDate)
If (dd% < 15) Then
    dbTaxeDate = "01." + Format(dbTaxeDate, "MM.YYYY")
Else
    dbTaxeDate = "15." + Format(dbTaxeDate, "MM.YYYY")
End If

Dim tRec As New ADODB.Recordset
EinspielDatumStempel = Format(dbTaxeDate, "YYYYMMDD")
h2 = "TAXE_" + EinspielDatumStempel
SQLStr = "SELECT NAME FROM  Sys.DATABASES"
tRec.Open SQLStr$, sqlop.ActiveConn
Do
    If (tRec.EOF) Then
        Exit Do
    End If
    
    h = UCase(tRec!Name)
    If (Left(h, 13) = h2) Then
        h = Mid(h, 6)
        If (h >= ret) Then
            ret = h
        End If
    End If
    
    tRec.MoveNext
Loop
tRec.Close

CheckAlteTaxe = ret

Call DefErrPop
End Function

Sub eRezeptSpeichern()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("eRezeptSpeichern")
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
If (Err.Number = (999 + vbObjectError)) Then End
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
'On Error GoTo 0
'Err.Raise 999 + vbObjectError, "DSK.VBP", "Fehler in DLL"
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

Dim tRec As New ADODB.Recordset
Dim F_Zuzages As Double, F_Zuza As Double, F_Aufz As Double, F_Preis As Double
Dim eRezept As New TI_Back.eRezept_OP
Dim OrgKBV As Long, IkNr As Long, ZA_Ind As Integer
Dim iMulti%
Dim i As Integer, j As Integer, ok As Boolean
Dim lPzn&
'Dim KeinTausch As Boolean

Dim IsNoctu As Boolean

MsgBox ("Neue eRezept-Daten werden gespeichert")
'OrgKBV = KBVNr
'For i = aind To 1 Step -1
'  If a(i).IkNr > 0 Then
'    IkNr = a(i).IkNr
'    Exit For
'  End If
'  If a(i).flag = FLAG_REZEPTANFANG Then Exit For
'Next i
'If IkNr > 0 Then KBVNr = IkNr

F_Preis = RezArtikel(0).Preis
'If Val(a(aind).Rezeptnummer) > 0 And a(aind).RabArtikelVerfügbar = WunschArtikel Then
'    F_Preis = 0
'End If
'If Val(a(aind).Rezeptnummer) > 0 And a(aind).AnzeigePreis = 0 Then
'  F_Preis = Abs(a(aind).RezeptPreis)
'  If F_Preis = 0 And (a(aind).WarAbholNr > 0 Or a(aind).Abholnummer > 0) Then
'    F_Preis = CVD(a(aind).Artikel.AVP)
'    If CVD(a(aind).Artikel.kp) > 0 And CVD(a(aind).Artikel.kp) < F_Preis Then F_Preis = CVD(a(aind).Artikel.kp)
'  End If
'End If
F_Preis = Abs(F_Preis)
'If a(aind).flag = FLAG_GEBUEHR Then F_Preis = a(aind).Zuzahlung

'If a(aind).ZuzaVerzicht Or a(aind).RezeptArt = REZEPTART_BTM_ZUZAFREI Or a(aind).RezeptArt = REZEPTART_ZUZAFREI Or a(aind).RezeptArt = REZEPTART_SPRECHSTUNDENBEDARF Then
'  F_Zuzages = 0
'  F_Zuza = 0
'Else
'  F_Zuzages = Abs(FNX(a(aind).Zuzahlung * a(aind).multi)) * Sgn(a(aind).multi)
'  If a(aind).HMVerbrauch = 2 And Abs(F_Zuzages) > (MAX_HMVERBRAUCHPROMONAT - 0.1) Then F_Zuzages = MAX_HMVERBRAUCHPROMONAT * Sgn(F_Zuzages)
'  F_Zuza = Abs(a(aind).Zuzahlung)
'End If
F_Zuzages = RezArtikel(0).Zuz * RezArtikel(0).Faktor
F_Zuza = F_Zuzages

'If a(aind).Mehrkostenverzicht Then
'  F_Aufz = 0
'Else
'  F_Aufz = a(aind).Aufzahlung
'  F_Preis = F_Preis + F_Aufz
'End If


Call eRezept.New2(eRezeptTaskId)
With eRezept
    .AbgabeDatum = CDate(Left(VerkDatum, 2) + "." + Mid(VerkDatum, 3, 2) + ".20" + Mid(VerkDatum, 5, 2)) 'Now

'    IsNoctu = .Noctu    '!!!
    IsNoctu = False
    For i% = 0 To (AnzRezeptArtikel% - 1)
        lPzn = Val(RezArtikel(i%).pzn)
        If (lPzn = 2567018) Then   'Noctu
            IsNoctu = True
        End If
    Next i
    
    With .Abgabe
        SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + RezArtikel(0).pzn
        'Set TaxeRec = TaxeDB.OpenRecordset(SQLStr$)
        On Error Resume Next
        TaxeRec.Close
        Err.Clear
        On Error GoTo DefErr
        TaxeRec.Open SQLStr, taxeAdoDB.ActiveConn
        If (TaxeRec.EOF = False) Then
            .VerordnungsTyp = VerordnungsTyp_OP_PZNVerordnung
            .pzn = uFormat(TaxeRec!pzn, "00000000")
            .sPackungsgroesse = Trim(TaxeRec!menge)
            .einheit = Trim(TaxeRec!einheit)
            .Verordnungstext = Trim(TaxeRec!Name)

            .Darreichungsform = TaxeRec!DarForm
            .Normgroesse = TaxeRec!NGrösse
'        ElseIf (eRezept.Verordnung.VerordnungsTyp = VerordnungsTyp_OP_Rezepturverordnung) Or (eRezept.Verordnung.VerordnungsTyp = VerordnungsTyp_OP_Freitextverordnung) Then
        ElseIf (eRezept.Abgabe.VerordnungsTyp = VerordnungsTyp_OP_Rezepturverordnung) Then
'            .VerordnungsTyp = VerordnungsTyp_OP_Freitextverordnung
            
            Dim PznStr$
            PznStr = RezArtikel(0).pzn
            If (PznStr = "06460749") Then
                PznStr = "06460748"
            ElseIf (PznStr = "02567114") Then
                PznStr = "02567113"
            ElseIf (PznStr = "02567115") Then
                PznStr = "02567113"
            End If
            .pzn = PznStr   ' RezArtikel(0).pzn
            
            Dim h$
            If (.Verordnungstext = "") Then
                h = RezArtikel(0).text
            Else
                h = .Verordnungstext
            End If
            h = Trim(InputBox("Text der Verordnung: ", "Text für die Dispensierug", h))
            If (h <> "") Then
                .Verordnungstext = h
            End If
        
'            .sPackungsgroesse = eRezept.Verordnung.sPackungsgroesse
'            .Einheit = eRezept.Verordnung.Einheit
'            .Darreichungsform = eRezept.Verordnung.Darreichungsform
        End If
'        TaxeRec.Close
    End With

    iMulti = xVal(.Anzahl)
'    'Gesamtbrutto= gesamter Preis für gesamtens Rezept (Nuctu, ...)
'    .sGesamtBrutto = uFormat(F_Preis * iMulti, "#######0.00")
'    If IsNoctu Then
'        .sGesamtBrutto = uFormat((F_Preis * iMulti) + NOCTU_PREIS, "#######0.00")
'    End If
    .sGesamtBrutto = uFormat(RezSumme - LieferEngPass, "#######0.00")
    'weg am 23.2.24
'    If IsNoctu Then
'        .sGesamtBrutto = uFormat(RezSumme + NOCTU_PREIS, "#######0.00")
'    End If
    
    'Bruttopreis= gesamter Preis für diese Zeile
    .sBruttoPreis = uFormat(F_Preis * iMulti, "#######0.00")
    .sMwst = uFormat(19, "#0.00")

'    If (.ZusatzAttribut(ZusatzattributGruppe_KuenstlicheBefruchtung).Schluessel = 0) Then
    If (.ZusatzAttribut(9).Schluessel = 0) Then
      .sEigenanteil = uFormat(F_Zuza + F_Aufz, "######0.00")
      .sGesamtZuzahlung = uFormat(0, "######0.00")      'für gesamtes Rezept
      .sZuzahlung = uFormat(0, "######0.00")      'für diese Zeile
    Else
      .sGesamtZuzahlung = uFormat(F_Zuzages, "######0.00")
      .sZuzahlung = uFormat(F_Zuzages, "######0.00")
      .sMehrkosten = uFormat(F_Aufz, "######0.00")
    End If

    If (TaxeRec.EOF = False) Then
        ok = CheckRabattAutIdem(Val(RezArtikel(0).pzn), 0)
    End If

'''    If (ImportMarktGilt(RezArtikel(0).pzn, 0)) Then
'''      .Zusatzattribut(ZusatzattributGruppe_Markt).Schluessel = 2      'Solitär = Importmarkt
'''      ZA_Ind = ZusatzattributGruppe_ImportFAM
'''    Else
'''      .Zusatzattribut(ZusatzattributGruppe_Markt).Schluessel = 1      'Generischer Markt
'''    End If
'''    'Mehrfachvertrieb (schlüssel=3) - was könnte das sein?
'''
'''    'bei Zusatzattribut 2-4 heißt Wert=0 irrelevant. Bei den Attributen ab 5 braucht Andi -1 als nein, 1 heißt ja
'''
'''
'''    .Zusatzattribut(ZusatzattributGruppe_Rabattvertragserfuellung).Schluessel = 0
'''    .Zusatzattribut(ZusatzattributGruppe_PreisguenstigesFAM).Schluessel = 0
'''    .Zusatzattribut(ZusatzattributGruppe_ImportFAM).Schluessel = 0
'''    .Zusatzattribut(ZusatzattributGruppe_Wunscharzneimittel).Schluessel = -1
'''    ZA_Ind = ZusatzattributGruppe_Rabattvertragserfuellung
'''
'''    If (RezArtikel(0).AutIdem = 3) Then
'''        .Zusatzattribut(ZusatzattributGruppe_Rabattvertragserfuellung).Schluessel = 1   'sicherheitshalber auch bei Importen
'''    Else
'''        If Not ok Then
'''            ZA_Ind = ZusatzattributGruppe_PreisguenstigesFAM
'''        End If
'''
'''        AMverfügbarkeit$ = CheckVerfügbarkeit
'''        Select Case Val(Mid$(AMverfügbarkeit, 1, 1))
'''        '    .AddItem "1 - ordnungsgem. Abgabe bzw. leere Verord.Zeile"
'''        '    .AddItem "2 - Nichtverfügbarkeit aller Rabattartikel"
'''        '    .AddItem "3 - Nichtverf. d. 4 günstigsten Generika bzw. günstiger Importe"
'''        '    .AddItem "4 - Nichverf. aller Rabattart. u. günstiger Generika/Importe"
'''        '    .AddItem "5 - Notwendigkeit unverzüglicher Abgabe (statt Rabattart.)"
'''        '    .AddItem "6 - Notwendigkeit unverzügl. Abgabe (statt Rabattart.u. günstiger Genrika/Importe)"
'''        '    .AddItem "7 - Wunscharzneimittel"
'''        '    .AddItem "8 - Pharmazeutische Bedenken (statt Rabattart.)"
'''        '    .AddItem "9 - Pharmazeutische Bedenken (statt Rabattart.u. günstiger Generika/Importe)"
'''            Case 2
'''              .Zusatzattribut(ZA_Ind).Schluessel = 2
'''            Case 3, 4
'''              .Zusatzattribut(ZA_Ind).Schluessel = 2
'''            Case 5, 6
'''              .Zusatzattribut(ZA_Ind).Schluessel = 3
'''            Case 7
'''              .Zusatzattribut(ZusatzattributGruppe_Wunscharzneimittel).Schluessel = 0 '!!!! 1
'''              '.sMehrkosten = uFormat(F_Preis * Abs(a(aind).multi) - F_Zuzages, "######0.00")
'''              .sGesamtBrutto = "0.00"
'''              .sBruttoPreis = "0.00"
'''              .sZuzahlung = "0.00"
'''              .sMehrkosten = "0.00"
'''              .sGesamtZuzahlung = "0.00"
'''            Case 8, 9
'''              .Zusatzattribut(ZA_Ind).Schluessel = 4
'''        End Select
'''    End If
'''    .Zusatzattribut(ZA_Ind).Freitext = ""   ' a(aind).RabArtikelVerfFreitext
'''    .Zusatzattribut(ZusatzattributGruppe_Mehrkostenuebernahme).Schluessel = -1
'''
''''    If a(aind).MKVerzichtArt = 1 Then
''''        .Zusatzattribut(ZusatzattributGruppe_Mehrkostenuebernahme).Schluessel = 1
''''    End If
'''    'MK-Übernahme nach Rabttvertrag wäre Schlüssel=2, daten dafür hab ich nicht
'''
''''    If a(aind).RezeptArt = REZEPTART_IVF Then
''''        .Zusatzattribut(ZusatzattributGruppe_KuenstlicheBefruchtung).Schluessel = 0   '!!! 1
''''    Else
''''        .Zusatzattribut(ZusatzattributGruppe_KuenstlicheBefruchtung).Schluessel = -1
''''    End If

    If IsNoctu Then
'        .ZusatzAttribut(ZusatzattributGruppe_AbgabeImNotdienst).Schluessel = 0
'        .ZusatzAttribut(ZusatzattributGruppe_AbgabeImNotdienst).Datum = CDate(Format(Now, "DD.MM.YYYY hh:nn"))
        .ZusatzAttribut(11).Schluessel = 0
        .ZusatzAttribut(11).Datum = CDate(Format(Now, "DD.MM.YYYY hh:nn"))
    Else
        .ZusatzAttribut(11).Schluessel = -1
    End If

    '.Zusatzattribut(ZusatzattributGruppe_OP_PreisguenstigesFAM).Schluessel = 1
    '.Zusatzattribut(ZusatzattributGruppe_ImportFAM).Schluessel = 1

    '.Rezeptaenderung(RezeptaenderungArt_AbweichungDAFO).RueckspracheArzt = RueckspracheArzt_OP_NichtErforderlich
    
    .sBotendienst = uFormat(0, "######0.00")
    .sBeschaffungsKosten = uFormat(0, "######0.00")
    For i% = 0 To (AnzRezeptArtikel% - 1)
        lPzn = Val(RezArtikel(i%).pzn)
        If (lPzn = 9999637) Then  'Beschaffungskosten
            .sBeschaffungsKosten = uFormat(RezArtikel(i).Preis, "######0.00")
        ElseIf (lPzn = 6461110) Then  'Botendienst
            .sBotendienst = uFormat(RezArtikel(i).Preis, "######0.00")
        End If
    Next i

    .SpeicherAbgabedaten
    .SpeicherAbgabeMedikation
    .SpeicherZusatzAttribute
'    .SpeicherRezeptAenderungen


'    If ((ParenteralRezept >= 0) Or (ParenteralRezept = -99)) And (MagSpeicherIndex > 0) Then
    If (MagSpeicherIndex > 0) Or (LennartzPzn > 0) Then
        Dim AvpRec As Recordset
        Dim iBestandteil As TI_Back.RezepturBestandteil_OP
        
'        MsgBox ("Rezeptur: lennartzpzn: " + CStr(LennartzPzn))

        Dim LennartzRezeptur%
        LennartzRezeptur = 0

        Dim HerstInd%
        Dim Unique$, FaktorKz$, PreisKz$
        Unique = RezeptUnique

        Dim sAuseinzelung As String
        Dim sAuseinzelung2 As String
        sAuseinzelung = ""
        sAuseinzelung2 = ""

'        h$ = Unique$ + Format(1, "0")
'        ArtikelRec.Seek "=", h$
'        If Not (ArtikelRec.NoMatch) Then
'            ret$ = ret$ + vbCrLf + "                <pPosition1>"
'    '        ret$ = ret$ + vbCrLf + "                  <pzn>" + CStr(ArtikelRec!pzn) + "</pzn>"
'
'            lPzn = ArtikelRec!pzn
'            If (lPzn = 6460749) Then
'                lPzn = 6460748
'            ElseIf (lPzn = 2567114) Then
'                lPzn = 2567113
'            ElseIf (lPzn = 2567115) Then
'                lPzn = 2567113
'            End If
'
'            ret$ = ret$ + vbCrLf + "                  <pzn>" + Right(String(8, "0") + CStr(lPzn), 8) + "</pzn>"
'    '        ret$ = ret$ + vbCrLf + "                  <posNr>" + CStr(j) + "</posNr>"
'            ret$ = ret$ + vbCrLf + "                  <faktor>" + Format(ArtikelRec!Faktor, "0") + "</faktor>"
'            ret$ = ret$ + vbCrLf + "                  <taxe>" + FiveRxPreis(ArtikelRec!Preis) + "</taxe>"
'    '        ret$ = ret$ + vbCrLf + "                  <autidem>0</autidem>"
'            ret$ = ret$ + vbCrLf + "                </pPosition1>"
'
'            If (ArtikelRec!pzn = 2567053) Then
'                sAuseinzelung = ArtikelRec!Zusatz2 + CheckNullStr(ArtikelRec!Zusatz2a)
'                sAuseinzelung2 = CheckNullStr(ArtikelRec!Zusatz3) + CheckNullStr(ArtikelRec!Zusatz3a)
'            End If
'        End If

        Dim CannabisRezept%
        CannabisRezept = 0
        Dim SubstitutionsRezept%
        SubstitutionsRezept = 0

        If (sAuseinzelung = "") Then
            h$ = Unique$ + "1"

            Set ArtikelRec = RezSpeicherDB.OpenRecordset("Artikel", dbOpenTable)
            ArtikelRec.index = "Unique"

            ArtikelRec.Seek "=", h$
            If Not (ArtikelRec.NoMatch) Then
                lPzn = ArtikelRec!pzn
                If (lPzn = 6460665) Or (lPzn = 6460694) Or (lPzn = 6460748) Or (lPzn = 6460754) Or (lPzn = 6460749) Then
                    CannabisRezept = True
                ElseIf (lPzn = 9999086) Or (lPzn = 2567107) Or (lPzn = 2567113) Or (lPzn = 2567136) Or (lPzn = 2567114) Or (lPzn = 2567115) Then
                    If (lPzn = 2567114) Then
                        lPzn = 2567113
                    ElseIf (lPzn = 2567115) Then
                        lPzn = 2567113
                    End If
                    SubstitutionsRezept = True
                ElseIf (Trim(CheckNullStr(ArtikelRec!text)) = "Cannabis-Rezeptur BG") Then
                    CannabisRezept = True
                End If
            End If
        End If

        If (SubstitutionsRezept) Then
'            Dim SubstitutionsMg As Double
'            Dim SubstitutionsAbgaben(2) As Integer
'            Dim SubstitutionsEinzelpreis As Double
'
'            AnzBtmTags = 0
'
'            BtmRezept = True
'            Dim h3$
'            h3 = "STANDARDREZEPT"
'            ind = InStr(ret, h3)
'            If (ind <= 0) Then
'                h3$ = "SPRECHSTUNDENBEDARF"
'                ind = InStr(ret, h3)
'            End If
'            If (ind > 0) Then
'                ret = Left(ret, ind - 1) + "BTM" + Mid(ret, ind + Len(h3))
'            End If
'
'            SubstitutionsMg = 0
'            For i = 0 To 2
'                SubstitutionsAbgaben(i) = 0
'            Next
'            SubstitutionsEinzelpreis = 0
'
'            SQLStr$ = "SELECT * FROM ParenteralTaxierungen WHERE RezepteUnique='" + Unique + "'"
'            SQLStr$ = SQLStr$ + " ORDER BY AnfMagInd"
'
'            Set AvpRec = RezSpeicherDB.OpenRecordset(SQLStr$)
'            Do
'                If (AvpRec.EOF) Then
'                    Exit Do
'                End If
'
'                If (InStr(UCase(CheckNullStr(AvpRec!text)), "E-DOSIS") > 0) Then
'                    SubstitutionsMg = AvpRec!GStufe
'                    SubstitutionsEinzelpreis = AvpRec!kp
'
'                    Dim SubstitutionsAbgabenStr As Integer
'                    SubstitutionsAbgabenStr = AvpRec!ActMenge
'                    Dim iMulti%
'                    iMulti = IIf(SubstitutionsAbgabenStr < 900, 10, 30)
'                    For i = 2 To 0 Step -1
'                        SubstitutionsAbgaben(i) = SubstitutionsAbgabenStr Mod iMulti
'                        SubstitutionsAbgabenStr = (SubstitutionsAbgabenStr \ iMulti)
'                    Next i
'                    Exit Do
'                End If
'                AvpRec.MoveNext
'            Loop
'            AvpRec.Close
'
'            If (SubstitutionsMg > 0) Then
'                Dim k%, m%
'                Dim dBtmZuschlag#
'                Dim iBtmMenge%
'
'                Set AvpRec = RezSpeicherDB.OpenRecordset(SQLStr$)
'                For m = 0 To 2
'                    If (SubstitutionsAbgaben(m) <= 0) Then
'                        Exit For
'                    End If
'
'                    For k = 1 To SubstitutionsAbgaben(m)
'                        dBtmZuschlag = 0
'                        AnzBtmTags = 0
'
'                        ret$ = ret$ + vbCrLf + "                <pCharge>"
'                        If (pCharge <> "") Then
'                            h$ = Mid(pCharge, 11, 13)
'                            h$ = Left$(h, 4) + "-" + Mid$(h$, 5, 2) + "-" + Mid$(h$, 7, 2) + "T" + Mid(h$, 10, 2) + ":" + Mid$(h$, 12, 2) + ":00.000"
'
'                            ret$ = ret$ + vbCrLf + "                  <herstellerSchluessel>" + Left(pCharge, 1) + "</herstellerSchluessel>"
'                            ret$ = ret$ + vbCrLf + "                  <herstellerNr>" + Mid(pCharge, 2, 9) + "</herstellerNr>"
'                            ret$ = ret$ + vbCrLf + "                  <herstellungsDatum>" + h$ + "</herstellungsDatum>"
'                            ret$ = ret$ + vbCrLf + "                  <chargenNr>" + Format(m + 1, "00") + "</chargenNr>"
'    '                        ret$ = ret$ + vbCrLf + "                  <anzahlApplikationen>" + Format(SubstitutionsAbgaben(m), "00") + "</anzahlApplikationen>"
'                            ret$ = ret$ + vbCrLf + "                  <anzahlApplikationen>" + Format(k, "00") + "</anzahlApplikationen>"
'                        End If
'
'                        AvpRec.MoveFirst
'
'                        i = 1
'                        For j = 1 To 100
'                            If (AvpRec.EOF) Then
'                                Exit For
'                            End If
'
'                            If (AvpRec!flag = MAG_HILFSTAXE) Or (AvpRec!flag = MAG_SPEZIALITAET) Then
'                                ret$ = ret$ + vbCrLf + "                  <pWirkstoff>"
'                                ret$ = ret$ + vbCrLf + "                    <pzn>" + Right(String(8, "0") + CStr(AvpRec!pzn), 8) + "</pzn>"
'                                ret$ = ret$ + vbCrLf + "                    <pPosNr>" + CStr(i) + "</pPosNr>"
'                                ret$ = ret$ + vbCrLf + "                    <faktor>" + HashFaktor(AvpRec!ActMenge / AvpRec!GStufe * 1000) + "</faktor>"
'
'                                FaktorKz = "11"
'                                ret$ = ret$ + vbCrLf + "                    <faktorKennzeichen>" + FaktorKz + "</faktorKennzeichen>"
'
'                                ret$ = ret$ + vbCrLf + "                    <taxe>" + FiveRxPreis(AvpRec!ActPreis) + "</taxe>"
'
'                                PreisKz = "14"
'                                ret$ = ret$ + vbCrLf + "                    <preisKennzeichen>" + PreisKz + "</preisKennzeichen>"
'
'                                ret$ = ret$ + vbCrLf + "                  </pWirkstoff>"
'                                i = i + 1
'    '                        ElseIf (AvpRec!pzn = 2567001) Then
'    '                            dBtmZuschlag = AvpRec!ActPreis
'    '                            iBtmMenge = AvpRec!ActMenge
'
'                            ElseIf (AvpRec!pzn = 2567001) Or (AvpRec!pzn = 9999637) Or (AvpRec!pzn = 6461110) Or (AvpRec!pzn = 2567018) Then
'                                If (AvpRec!pzn = 2567001) Then  'BTM
'                                    PreisKz = "81"
'                                ElseIf (AvpRec!pzn = 9999637) Then  'Beschaffungskosten
'                                    PreisKz = "82"
'                                ElseIf (AvpRec!pzn = 6461110) Then  'Botendienst
'                                    PreisKz = "83"
'                                ElseIf (AvpRec!pzn = 2567018) Then  'Noctu
'                                    PreisKz = "80"
'                                End If
'
'                                sBTM(AnzBtmTags) = vbCrLf + "                  <pWirkstoff>"
'                                sBTM(AnzBtmTags) = sBTM(AnzBtmTags) + vbCrLf + "                    <pzn>" + Right(String(8, "0") + CStr(AvpRec!pzn), 8) + "</pzn>"
'
'                                sBTM2(AnzBtmTags) = vbCrLf + "                    <faktor>" + HashFaktor(1000 * AvpRec!ActMenge) + "</faktor>"
'
'                                FaktorKz = "11"
'                                sBTM2(AnzBtmTags) = sBTM2(AnzBtmTags) + vbCrLf + "                    <faktorKennzeichen>" + FaktorKz + "</faktorKennzeichen>"
'
'                                sBTM2(AnzBtmTags) = sBTM2(AnzBtmTags) + vbCrLf + "                    <taxe>" + FiveRxPreis(AvpRec!ActPreis) + "</taxe>"
'                                sBTM2(AnzBtmTags) = sBTM2(AnzBtmTags) + vbCrLf + "                    <preisKennzeichen>" + PreisKz + "</preisKennzeichen>"
'                                sBTM2(AnzBtmTags) = sBTM2(AnzBtmTags) + vbCrLf + "                  </pWirkstoff>"
'                                AnzBtmTags = AnzBtmTags + 1
'
'                            ElseIf (AvpRec!flag = MAG_GEFAESS) Then
'                                ret$ = ret$ + vbCrLf + "                  <pWirkstoff>"
'                                ret$ = ret$ + vbCrLf + "                    <pzn>" + Right(String(8, "0") + CStr(AvpRec!pzn), 8) + "</pzn>"
'                                ret$ = ret$ + vbCrLf + "                    <pPosNr>" + CStr(i) + "</pPosNr>"
'                                ret$ = ret$ + vbCrLf + "                    <faktor>" + HashFaktor(1000) + "</faktor>"
'
'                                FaktorKz = "11"
'                                ret$ = ret$ + vbCrLf + "                    <faktorKennzeichen>" + FaktorKz + "</faktorKennzeichen>"
'
'                                ret$ = ret$ + vbCrLf + "                    <taxe>" + FiveRxPreis(AvpRec!ActPreis) + "</taxe>"
'                                ret$ = ret$ + vbCrLf + "                    <preisKennzeichen>" + "14" + "</preisKennzeichen>"
'                                ret$ = ret$ + vbCrLf + "                  </pWirkstoff>"
'                                i = i + 1
'                            End If
'
'                            AvpRec.MoveNext
'                        Next j
'
'                        If (k = SubstitutionsAbgaben(m)) Then
'                            ret$ = ret$ + vbCrLf + "                  <pWirkstoff>"
'                            ret$ = ret$ + vbCrLf + "                    <pzn>" + Right(String(8, "0") + CStr(lPzn), 8) + "</pzn>"
'                            ret$ = ret$ + vbCrLf + "                    <pPosNr>" + CStr(i) + "</pPosNr>"
'                            ret$ = ret$ + vbCrLf + "                    <faktor>" + HashFaktor(SubstitutionsMg * 10) + "</faktor>"
'
'                            FaktorKz = "55"
'                            ret$ = ret$ + vbCrLf + "                    <faktorKennzeichen>" + FaktorKz + "</faktorKennzeichen>"
'
'                            ret$ = ret$ + vbCrLf + "                    <taxe>" + FiveRxPreis(SubstitutionsEinzelpreis) + "</taxe>"
'                            ret$ = ret$ + vbCrLf + "                    <preisKennzeichen>" + "14" + "</preisKennzeichen>"
'                            ret$ = ret$ + vbCrLf + "                  </pWirkstoff>"
'                            i = i + 1
'
'                            Dim MitBtm%
'                            MitBtm = (m = 2)
'                            If (MitBtm = 0) Then
'                                MitBtm = (SubstitutionsAbgaben(m + 1) <= 0)
'                            End If
'    '                        If (MitBtm) And (dBtmZuschlag > 0) Then
'    '                            ret = ret + vbCrLf + "                  <pWirkstoff>"
'    '                            ret = ret + vbCrLf + "                    <pzn>" + Right(String(8, "0") + CStr(2567001), 8) + "</pzn>"
'    '                            ret = ret + vbCrLf + "                    <pPosNr>" + CStr(i) + "</pPosNr>"
'    '                            ret = ret + vbCrLf + "                    <faktor>" + Format(1000 * iBtmMenge, "0") + "</faktor>"
'    '
'    '                            FaktorKz = "11"
'    '                            ret = ret + vbCrLf + "                    <faktorKennzeichen>" + FaktorKz + "</faktorKennzeichen>"
'    '
'    '                            ret = ret + vbCrLf + "                    <taxe>" + FiveRxPreis(dBtmZuschlag) + "</taxe>"
'    '                            ret = ret + vbCrLf + "                    <preisKennzeichen>" + "81" + "</preisKennzeichen>"
'    '                            ret = ret + vbCrLf + "                  </pWirkstoff>"
'    '                            'i = i + 1
'    '                        End If
'                            If (MitBtm) And (AnzBtmTags > 0) Then
'                                For j = 1 To AnzBtmTags
'                                    ret = ret + sBTM(j - 1)
'                                    ret = ret + vbCrLf + "                    <pPosNr>" + CStr(i) + "</pPosNr>"
'                                    ret = ret + sBTM2(j - 1)
'                                    i = i + 1
'                                Next j
'                            End If
'                        End If
'
'                        ret$ = ret$ + vbCrLf + "                </pCharge>"
'                    Next k
'                Next m
'
'                AvpRec.Close
'            Else
'                Call MsgBox("Problem: Rezept '" + RezepteRec!RezeptNr + "' ist nicht übertragbar!" + vbCrLf + vbCrLf + "Vorgang wird abgebrochen!", vbCritical)
'                End
'            End If
        Else
            h$ = Mid(pCharge, 11, 13)
'            h$ = Left$(h, 4) + "-" + Mid$(h$, 5, 2) + "-" + Mid$(h$, 7, 2) + "T" + Mid(h$, 10, 2) + ":" + Mid$(h$, 12, 2) + ":00Z"

'            Dim iHerstellung As New TI_Back.RezepturHerstellung_OP
'            With iHerstellung
'                .Charge = Mid(pCharge, 24, 2)
'                .HerstellerKeyTA3 = Left(pCharge, 1)
'                .HerstellerKz = Mid(pCharge, 2, 9)
'                .Herstellungszeit = h
'            End With
'            .AddRezepturHerstellung (iHerstellung)

            HerstInd = .AddRezepturHerstellung(Mid(pCharge, 24, 2), Left(pCharge, 1), Mid(pCharge, 2, 9), h)

        ''''''''''''''''''''''''
            If (sAuseinzelung <> "") Then
'                i = 1
'                ret$ = ret$ + vbCrLf + "                  <pWirkstoff>"
'                ret$ = ret$ + vbCrLf + "                    <pzn>" + Left(sAuseinzelung, 8) + "</pzn>"
'                ret$ = ret$ + vbCrLf + "                    <pPosNr>" + CStr(i) + "</pPosNr>"
'                If (TA1_V37) Then
'                    ret$ = ret$ + vbCrLf + "                    <faktor>" + HashFaktor(Val(Mid(sAuseinzelung, 11, 13))) + "</faktor>"
'                    ret$ = ret$ + vbCrLf + "                    <faktorKennzeichen>" + Mid(sAuseinzelung, 9, 2) + "</faktorKennzeichen>"
'                    ret$ = ret$ + vbCrLf + "                    <taxe>" + FiveRxPreis(xVal(Mid(sAuseinzelung, 26, 12))) + "</taxe>"
'                    ret$ = ret$ + vbCrLf + "                    <preisKennzeichen>" + Mid(sAuseinzelung, 24, 2) + "</preisKennzeichen>"
'                Else
'                    ret$ = ret$ + vbCrLf + "                    <faktor>" + HashFaktor(Val(Mid(sAuseinzelung, 11, 5))) + "</faktor>"
'                    ret$ = ret$ + vbCrLf + "                    <faktorKennzeichen>" + Mid(sAuseinzelung, 9, 2) + "</faktorKennzeichen>"
'                    ret$ = ret$ + vbCrLf + "                    <taxe>" + FiveRxPreis(Val(Mid(sAuseinzelung, 18, 9)) / 100) + "</taxe>"
'                    ret$ = ret$ + vbCrLf + "                    <preisKennzeichen>" + Mid(sAuseinzelung, 16, 2) + "</preisKennzeichen>"
'                End If
'                ret$ = ret$ + vbCrLf + "                  </pWirkstoff>"
'                i = i + 1
'
'                If (sAuseinzelung2 <> "") Then
'                    ret$ = ret$ + vbCrLf + "                  <pWirkstoff>"
'                    ret$ = ret$ + vbCrLf + "                    <pzn>" + Left(sAuseinzelung2, 8) + "</pzn>"
'                    ret$ = ret$ + vbCrLf + "                    <pPosNr>" + CStr(i) + "</pPosNr>"
'                    If (TA1_V37) Then
'                        ret$ = ret$ + vbCrLf + "                    <faktor>" + HashFaktor(Val(Mid(sAuseinzelung2, 11, 13))) + "</faktor>"
'                        ret$ = ret$ + vbCrLf + "                    <faktorKennzeichen>" + Mid(sAuseinzelung2, 9, 2) + "</faktorKennzeichen>"
'                        ret$ = ret$ + vbCrLf + "                    <taxe>" + FiveRxPreis(xVal(Mid(sAuseinzelung2, 26, 12))) + "</taxe>"
'                        ret$ = ret$ + vbCrLf + "                    <preisKennzeichen>" + Mid(sAuseinzelung2, 24, 2) + "</preisKennzeichen>"
'                    Else
'                        ret$ = ret$ + vbCrLf + "                    <faktor>" + HashFaktor(Val(Mid(sAuseinzelung2, 11, 5))) + "</faktor>"
'                        ret$ = ret$ + vbCrLf + "                    <faktorKennzeichen>" + Mid(sAuseinzelung2, 9, 2) + "</faktorKennzeichen>"
'                        ret$ = ret$ + vbCrLf + "                    <taxe>" + FiveRxPreis(Val(Mid(sAuseinzelung2, 18, 9)) / 100) + "</taxe>"
'                        ret$ = ret$ + vbCrLf + "                    <preisKennzeichen>" + Mid(sAuseinzelung2, 16, 2) + "</preisKennzeichen>"
'                    End If
'                    ret$ = ret$ + vbCrLf + "                  </pWirkstoff>"
'                    i = i + 1
'                End If
'
'                If (Mid(sAuseinzelung, IIf(TA1_V37, 38, 27), 1) = "1") Then
'                    ret$ = ret$ + vbCrLf + "                  <pWirkstoff>"
'                    ret$ = ret$ + vbCrLf + "                    <pzn>" + Right(String(8, "0") + CStr(2567001), 8) + "</pzn>"
'                    ret$ = ret$ + vbCrLf + "                    <pPosNr>" + CStr(i) + "</pPosNr>"
'                    ret$ = ret$ + vbCrLf + "                    <faktor>" + HashFaktor(1000) + "</faktor>"
'
'                    FaktorKz = "11"
'                    ret$ = ret$ + vbCrLf + "                    <faktorKennzeichen>" + FaktorKz + "</faktorKennzeichen>"
'
'                    ret$ = ret$ + vbCrLf + "                    <taxe>" + FiveRxPreis(3.58) + "</taxe>"
'                    ret$ = ret$ + vbCrLf + "                    <preisKennzeichen>" + "81" + "</preisKennzeichen>"
'                    ret$ = ret$ + vbCrLf + "                  </pWirkstoff>"
'                    i = i + 1
'                End If

            Else
            End If
        '''''''''''''''''''''''

            SQLStr$ = "SELECT * FROM ParenteralTaxierungen WHERE RezepteUnique='" + Unique + "'"
            SQLStr$ = SQLStr$ + " ORDER BY AnfMagInd"

            Dim ActPreis As Double
            Dim CannabisFixaufschlag#
            CannabisFixaufschlag = 0

            If (CannabisRezept) Then
                Set AvpRec = RezSpeicherDB.OpenRecordset(SQLStr$)
                Do
                    If (AvpRec.EOF) Then
                        Exit Do
                    End If

                    If (InStr(UCase(CheckNullStr(AvpRec!text)), "FIXZUSCHLAG") > 0) Then
                        CannabisFixaufschlag = AvpRec!ActPreis
                        Exit Do
                    End If

                    AvpRec.MoveNext
                Loop
                AvpRec.Close
            End If

            SQLStr$ = "SELECT * FROM ParenteralTaxierungen WHERE RezepteUnique='" + Unique + "'"
            SQLStr$ = SQLStr$ + " ORDER BY AnfMagInd"
            Set AvpRec = RezSpeicherDB.OpenRecordset(SQLStr$)
            If Not (AvpRec.EOF) Then
                i = 1
                For j = 1 To 100
                    If (AvpRec.EOF) Then
                        Exit For
                    End If

                    FaktorKz = "11"
                    If (CheckNullLong(AvpRec!Verwurf)) Then
                        FaktorKz = "99"
                    End If

                    PreisKz = "14"

                    If (LennartzPzn > 0) Then
'                        ret$ = ret$ + vbCrLf + "                  <pWirkstoff>"
'                        ret$ = ret$ + vbCrLf + "                    <pzn>" + Right(String(8, "0") + CStr(AvpRec!pzn), 8) + "</pzn>"
'                        ret$ = ret$ + vbCrLf + "                    <pPosNr>" + CStr(i) + "</pPosNr>"
'                        ret$ = ret$ + vbCrLf + "                    <faktor>" + HashFaktor(AvpRec!kp) + "</faktor>"
'
'                        FaktorKz = "11"
'                        ret$ = ret$ + vbCrLf + "                    <faktorKennzeichen>" + FaktorKz + "</faktorKennzeichen>"
'
'                        ret$ = ret$ + vbCrLf + "                    <taxe>" + FiveRxPreis(AvpRec!ActPreis) + "</taxe>"
'                        ret$ = ret$ + vbCrLf + "                    <preisKennzeichen>" + Format(Val(CheckNullStr(AvpRec!menge)), "00") + "</preisKennzeichen>"
'                        ret$ = ret$ + vbCrLf + "                  </pWirkstoff>"
'                        i = i + 1
                        Set iBestandteil = New TI_Back.RezepturBestandteil_OP
                        With iBestandteil
                            .Artikelnummer = Right(String(8, "0") + CStr(AvpRec!pzn), 8)
                            .ArtikelnummerTyp = IIf((AvpRec!flag = MAG_HILFSTAXE) Or (AvpRec!flag = MAG_SPEZIALITAET) Or (AvpRec!flag = MAG_ANTEILIG), 1, 2) ' AvpRec!flag
                            .Faktor = HashFaktor(AvpRec!kp)
                            .Faktorkennzeichen = FaktorKz
                            .Preiskennzeichen = Format(Val(CheckNullStr(AvpRec!menge)), "00")
                            .Taxe = FiveRxPreis(AvpRec!ActPreis)
                            .Positionsnummer = i

                            Call eRezept.RezepturHerstellung(HerstInd).AddRezepturBestandteil(.Positionsnummer, .Artikelnummer, .ArtikelnummerTyp, .Faktorkennzeichen, .Faktor, .Preiskennzeichen, .Taxe)
                        End With
                        i = i + 1
                    ElseIf (AvpRec!flag = MAG_HILFSTAXE) Or (AvpRec!flag = MAG_SPEZIALITAET) Or (AvpRec!flag = MAG_ANTEILIG) Then
                        ActPreis = AvpRec!ActPreis
                        If (CannabisRezept) Then
                            ActPreis = ActPreis + CannabisFixaufschlag
                            CannabisFixaufschlag = 0
                        End If

'                        PreisKz = "14"
                        If (CheckNullStr(AvpRec!menge) = "1") Then
                            PreisKz = "15"
                        End If

                        Set iBestandteil = New TI_Back.RezepturBestandteil_OP
                        With iBestandteil
                            .Artikelnummer = Right(String(8, "0") + CStr(AvpRec!pzn), 8)
                            .ArtikelnummerTyp = 1   'bei ADV: PZN   'AvpRec!flag
                            .Faktor = HashFaktor(AvpRec!ActMenge / AvpRec!GStufe * 1000)
                            .Faktorkennzeichen = FaktorKz
                            .Preiskennzeichen = PreisKz
                            .Taxe = FiveRxPreis(ActPreis)
                            .Positionsnummer = i

                            Call eRezept.RezepturHerstellung(HerstInd).AddRezepturBestandteil(.Positionsnummer, .Artikelnummer, .ArtikelnummerTyp, .Faktorkennzeichen, .Faktor, .Preiskennzeichen, .Taxe)
                        End With
'                        iHerstellung.AddBestandteil (iBestandteil)

                        i = i + 1
                    ElseIf (AvpRec!flag = MAG_ARBEIT) And ((AvpRec!pzn = 9999092) Or (AvpRec!pzn = 2567478) Or (AvpRec!pzn = 2567461) Or (AvpRec!pzn = 9999146) Or (AvpRec!pzn = 9999169)) Then
                        PreisKz = "74"

                        Set iBestandteil = New TI_Back.RezepturBestandteil_OP
                        With iBestandteil
                            .Artikelnummer = "06460518"
                            .ArtikelnummerTyp = 2   'bei ADV: SonderPzn ' AvpRec!flag
                            .Faktor = HashFaktor(1000)
                            .Faktorkennzeichen = FaktorKz
                            .Preiskennzeichen = PreisKz
                            .Taxe = FiveRxPreis(AvpRec!kp)
                            .Positionsnummer = i

                            Call eRezept.RezepturHerstellung(HerstInd).AddRezepturBestandteil(.Positionsnummer, .Artikelnummer, .ArtikelnummerTyp, .Faktorkennzeichen, .Faktor, .Preiskennzeichen, .Taxe)
                        End With
'                        iHerstellung.AddBestandteil (iBestandteil)

                        i = i + 1
                    ElseIf (CannabisRezept) Or (sAuseinzelung = "") Then
                        If (AvpRec!pzn = 2567001) Or (AvpRec!pzn = 9999637) Or (AvpRec!pzn = 6461110) Or (AvpRec!pzn = 2567018) Then
                            If (AvpRec!pzn = 2567001) Then  'BTM
                                PreisKz = "81"
                            ElseIf (AvpRec!pzn = 9999637) Then  'Beschaffungskosten
                                PreisKz = "82"
                            ElseIf (AvpRec!pzn = 6461110) Then  'Botendienst
                                PreisKz = "83"
                            ElseIf (AvpRec!pzn = 2567018) Then  'Noctu
                                PreisKz = "80"
                            End If

                            Set iBestandteil = New TI_Back.RezepturBestandteil_OP
                            With iBestandteil
                                .Artikelnummer = Right(String(8, "0") + CStr(AvpRec!pzn), 8)
                                .ArtikelnummerTyp = 2   ' AvpRec!flag
                                .Faktor = HashFaktor(1000 * AvpRec!ActMenge)
                                .Faktorkennzeichen = FaktorKz
                                .Preiskennzeichen = PreisKz
                                .Taxe = FiveRxPreis(AvpRec!ActPreis)
                                .Positionsnummer = i

                                Call eRezept.RezepturHerstellung(HerstInd).AddRezepturBestandteil(.Positionsnummer, .Artikelnummer, .ArtikelnummerTyp, .Faktorkennzeichen, .Faktor, .Preiskennzeichen, .Taxe)
                            End With
    '                        iHerstellung.AddBestandteil (iBestandteil)

                            i = i + 1
                        ElseIf (AvpRec!pzn = 9999117) Or (AvpRec!pzn = 9999206) Then
                            Set iBestandteil = New TI_Back.RezepturBestandteil_OP
                            With iBestandteil
                                .Artikelnummer = Right(String(8, "0") + CStr(AvpRec!pzn), 8)
                                .ArtikelnummerTyp = 2   ' AvpRec!flag
                                .Faktor = HashFaktor(1)
                                .Faktorkennzeichen = FaktorKz
                                .Preiskennzeichen = PreisKz
                                .Taxe = FiveRxPreis(AvpRec!ActPreis)
                                .Positionsnummer = i

                                Call eRezept.RezepturHerstellung(HerstInd).AddRezepturBestandteil(.Positionsnummer, .Artikelnummer, .ArtikelnummerTyp, .Faktorkennzeichen, .Faktor, .Preiskennzeichen, .Taxe)
                            End With
    '                        iHerstellung.AddBestandteil (iBestandteil)

                            i = i + 1

                        ElseIf (InStr(UCase(CheckNullStr(AvpRec!text)), "FIX-AUFSCHLAG") > 0) Then
                            PreisKz = IIf(PreisKz_62_70, "70", "74")

                            Set iBestandteil = New TI_Back.RezepturBestandteil_OP
                            With iBestandteil
                                .Artikelnummer = "06460518"
                                .ArtikelnummerTyp = 2   ' AvpRec!flag
                                .Faktor = HashFaktor(1000)
                                .Faktorkennzeichen = FaktorKz
                                .Preiskennzeichen = PreisKz
                                .Taxe = FiveRxPreis(AvpRec!ActPreis)
                                .Positionsnummer = i

                                Call eRezept.RezepturHerstellung(HerstInd).AddRezepturBestandteil(.Positionsnummer, .Artikelnummer, .ArtikelnummerTyp, .Faktorkennzeichen, .Faktor, .Preiskennzeichen, .Taxe)
                            End With
    '                        iHerstellung.AddBestandteil (iBestandteil)

                            i = i + 1
                        ElseIf (AvpRec!flag = MAG_ARBEIT) And (AvpRec!ActPreis > 0) Then
                            PreisKz = IIf(PreisKz_62_70, "62", "74")

                            Set iBestandteil = New TI_Back.RezepturBestandteil_OP
                            With iBestandteil
                                .Artikelnummer = "06460518"
                                .ArtikelnummerTyp = 2   ' AvpRec!flag
                                .Faktor = HashFaktor(1000)
                                .Faktorkennzeichen = FaktorKz
                                .Preiskennzeichen = PreisKz
                                .Taxe = FiveRxPreis(AvpRec!ActPreis)
                                .Positionsnummer = i

                                Call eRezept.RezepturHerstellung(HerstInd).AddRezepturBestandteil(.Positionsnummer, .Artikelnummer, .ArtikelnummerTyp, .Faktorkennzeichen, .Faktor, .Preiskennzeichen, .Taxe)
                            End With
    '                        iHerstellung.AddBestandteil (iBestandteil)

                            i = i + 1
                        ElseIf (AvpRec!flag = MAG_GEFAESS) Then
                            Set iBestandteil = New TI_Back.RezepturBestandteil_OP
                            With iBestandteil
                                .Artikelnummer = IIf(AvpRec!pzn = 0, "06461328", Right(String(8, "0") + CStr(AvpRec!pzn), 8))
                                .ArtikelnummerTyp = 2   ' AvpRec!flag
                                .Faktor = HashFaktor(IIf(AvpRec!pzn = 0, 1, 1000))
                                .Faktorkennzeichen = FaktorKz
                                .Preiskennzeichen = PreisKz
                                .Taxe = FiveRxPreis(AvpRec!ActPreis)
                                .Positionsnummer = i

                                Call eRezept.RezepturHerstellung(HerstInd).AddRezepturBestandteil(.Positionsnummer, .Artikelnummer, .ArtikelnummerTyp, .Faktorkennzeichen, .Faktor, .Preiskennzeichen, .Taxe)
                            End With
    '                        iHerstellung.AddBestandteil (iBestandteil)

                            i = i + 1
                        End If
                    End If

                    AvpRec.MoveNext
                Next j
            End If
        End If
'MsgBox ("30")
        .SpeicherRezepturHerstellungen
'MsgBox ("31")
    End If

End With

'KbvNr = OrgKBV

Call DefErrPop
End Sub

Sub eRezeptRestore(eRezeptOrg As TI_Back.eRezept_OP)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("eRezeptRestore")
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
If (Err.Number = (999 + vbObjectError)) Then End
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
'On Error GoTo 0
'Err.Raise 999 + vbObjectError, "DSK.VBP", "Fehler in DLL"
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

Dim eRezept As New TI_Back.eRezept_OP
Dim iMulti%
Dim i As Integer, ok As Boolean
'Dim KeinTausch As Boolean

MsgBox ("eRezept-Daten werden RÜCKGESETZT")

Call eRezept.New2(eRezeptTaskId)
With eRezept
    .AbgabeDatum = eRezeptOrg.AbgabeDatum

    With .Abgabe
        .VerordnungsTyp = eRezeptOrg.Abgabe.VerordnungsTyp
        .pzn = eRezeptOrg.Abgabe.pzn
        .sPackungsgroesse = eRezeptOrg.Abgabe.sPackungsgroesse
        .einheit = eRezeptOrg.Abgabe.einheit
        .Verordnungstext = eRezeptOrg.Abgabe.Verordnungstext

        .Darreichungsform = eRezeptOrg.Abgabe.Darreichungsform
        .Normgroesse = eRezeptOrg.Abgabe.Normgroesse
    End With

    .sGesamtBrutto = eRezeptOrg.sGesamtBrutto
    .sBruttoPreis = eRezeptOrg.sBruttoPreis
    .sMwst = eRezeptOrg.sMwst
    
    .sEigenanteil = eRezeptOrg.sEigenanteil
    .sGesamtZuzahlung = eRezeptOrg.sGesamtZuzahlung
    .sZuzahlung = eRezeptOrg.sZuzahlung
    .sMehrkosten = eRezeptOrg.sMehrkosten

    For i = 1 To 15
        .ZusatzAttribut(i).Schluessel = eRezeptOrg.ZusatzAttribut(i).Schluessel
    Next

    .SpeicherAbgabedaten
    .SpeicherAbgabeMedikation
    .SpeicherZusatzAttribute
    .SpeicherRezeptAenderungen
End With

'KbvNr = OrgKBV

Call DefErrPop
End Sub

Function ImportMarktGilt%(pzn$, IstImportModus%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("ImportMarktGilt%")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
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
Dim iAdd%, iAdd2%
Dim row%, NeueKbvNr%, ind%, ind2%
Dim ImportGr&
Dim h$

ImportMarktGilt = False

IdentPzn$ = pzn$

SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + pzn$
'Set TaxeRec = TaxeDB.OpenRecordset(SQLStr$)
On Error Resume Next
TaxeRec.Close
Err.Clear
On Error GoTo DefErr
TaxeRec.Open SQLStr, taxeAdoDB.ActiveConn
If (TaxeRec.EOF) Then
    Call DefErrPop: Exit Function
End If

ImportGr = TaxeRec!ImportGruppeNr
If ImportGr = 0 Then
    Call DefErrPop
    Exit Function
End If

If (PrivatRezept% = 0) And (AutIdemKbvNr& = 0) Then
    frmKkMatchcode.Show 1
    If (AutIdemKbvNr& > 0) Then
        NeueKbvNr% = True
    End If
End If

ReDim ArtLRec(0)
Dim iArtLRec%
iArtLRec = -1


If (IstImportModus) Then
'    Dim ast1 As StammdatenType
    Dim OrgPZN As String
    Dim OrgAvp As Double
    Dim AVP As Double
    Dim testavp As Double
    Dim fhImp As Integer
    Dim xLinks As Long
    Dim xRechts As Long
    Dim xMitte As Long
    Dim gefunden As Long
    Dim test As String
    Dim impPzn As String
    Dim Anzahl As Integer
    Dim erg As Integer
    Dim billiger As Boolean
    Dim billigerAbsolut As Boolean
    Dim NurBilligere As Boolean
    Dim SuchLagernde As Boolean
    Dim Preis As Double
    Dim ImportGruppeNr As Long
    Dim tRec As New ADODB.Recordset
    Dim SQL As String
    Dim vkt As Double
    Dim Durchlauf As Integer
    Dim IstBilliger As Boolean
    Dim i, j As Long
    Dim RoteLinieDa As Boolean
    Dim AusgangsPreis As Double
    Dim FB As Double
    Dim PZNDatei As String
    Dim fhout As Long
    Dim AnzVerfügbarkeitsabfrage As Integer
    Dim artPreis As Double
    Dim NoOriginal As Boolean
    Dim ImRezept As Boolean
    Dim ImportAlt As Boolean

    ImportAlt = False

    OrgPZN = pzn
'    If Val(Rezeptnummer) > 0 And Not Sprechstundenbedarf Then ImRezept = True
    ImRezept = True

    testavp = TaxeRec!vk / 100#
    AusgangsPreis = testavp
    'DrS 9.4.2014 Anzeige immer VK - Herstellerrabatt

    '10.6.19 DrS: günstigkeit eines Imports immer nach AVP unabhängig Herstellerrabatt
'    If Val(Rezeptnummer) > 0 And TaxeRec!HerstRabattKZ <> 0 Then
    If TaxeRec!HerstRabattKZ <> 0 Then
        AusgangsPreis = AusgangsPreis - HerstRabattBrutto(TaxeRec)
    End If

    FB = dCheckNull(TaxeRec!Rabwert_130a_2_SGB) / 100#
    If FB > 0 Then    'Impfstoffe: wenn Rabatt nach §130a günstiger als Herstellerrabatt, dann muss dieser abgezogen werden
        If (TaxeRec!vk / 100#) - FB < AusgangsPreis Then AusgangsPreis = (TaxeRec!vk / 100#) - FB
    End If

    If ImRezept Then
      FB = dCheckNull(TaxeRec!FESTBETRAG) / 100#
      If FB > 0 And FB < AusgangsPreis Then AusgangsPreis = FB    '1.4.49
    End If

    NoOriginal = True
    If TaxeRec!OriginalPZN <> 0 Then
        NoOriginal = False
        pzn = PznString(TaxeRec!OriginalPZN)
        If pzn <> OrgPZN Then
            SQLStr = "SELECT * FROM Taxe WHERE PZN=" + pzn
            On Error Resume Next
            TaxeRec.Close
            Err.Clear
            On Error GoTo DefErr
            TaxeRec.Open SQLStr, taxeAdoDB.ActiveConn
            If (TaxeRec.EOF) Then
                Call DefErrPop: Exit Function
            End If
        
            Preis = TaxeRec!vk
            Preis = Preis / 100#
        
            If TaxeRec!HerstRabattKZ <> 0 Then
              Preis = Preis - HerstRabattBrutto(TaxeRec)
            End If
        
            FB = dCheckNull(TaxeRec!Rabwert_130a_2_SGB) / 100#
            If FB > 0 Then    'Impfstoffe: wenn Rabatt nach §130a günstiger als Herstellerrabatt, dann muss dieser abgezogen werden
              If (TaxeRec!vk / 100#) - FB < Preis Then Preis = (TaxeRec!vk / 100#) - FB
            End If
        
            FB = dCheckNull(TaxeRec!FESTBETRAG) / 100#
            If FB > 0 And FB < Preis Then Preis = FB    '1.4.49
        Else
            Preis = AusgangsPreis   ' testavp
        End If
    End If
    ImportGruppeNr = TaxeRec!ImportGruppeNr

    Call Taxe2ast(pzn$)    'OriginalArtikel. Wenn es keine OriginalPZN gibt, dann Ausgangsartikel

    OrgAvp = ast.AVP
    artPreis = OrgAvp
    '10.6.19 DrS: günstigkeit eines Imports immer nach AVP unabhängig Herstellerrabatt
'    If TaxeRec!HerstRabattKZ <> 0 And Val(Rezeptnummer) > 0 Then
    If TaxeRec!HerstRabattKZ <> 0 Then
      artPreis = FNX(artPreis - HerstRabattBrutto(TaxeRec))
    End If

    FB = dCheckNull(TaxeRec!Rabwert_130a_2_SGB) / 100#
    If FB > 0 Then    'Impfstoffe: wenn Rabatt nach §130a günstiger als Herstellerrabatt, dann muss dieser abgezogen werden
      If OrgAvp - FB < artPreis Then artPreis = OrgAvp - FB
    End If

    If ImRezept Then
      FB = dCheckNull(TaxeRec!FESTBETRAG) / 100#
      If FB > 0 And FB < artPreis Then artPreis = FB    '1.4.49
    End If

    ArtL.pzn = Val(ast.pzn)
    ArtL.kurz = ast.kurz
    ArtL.Preis = uFormat(artPreis, Space(Len(ArtL.Preis) - 4) + "0.00")
    Anzahl = 1
    If TaxeRec!OriginalPZN > 0 Then
      ArtL.kz1 = "o"
    Else
      ArtL.kz1 = " "    'Importgruppe, aber keine OriginalPZN
    End If
    ArtL.kz2 = ".-"
    ArtL.GhVerfügbar = 0
    ArtL.rabKZ = " "
    If (Val(ast.pzn) = Val(OrgPZN)) Then ArtL.AusgangsArtikel = 1 Else ArtL.AusgangsArtikel = 0
    If IstRabattAutIdem(Val(ArtL.pzn)) Then ArtL.rabKZ = "r"
    ArtL.RoteLinie = 0

    ArtL.ArtStatus = TaxeRec!ArtStatus

    iArtLRec = iArtLRec + 1
    ReDim Preserve ArtLRec(iArtLRec)
    ArtLRec(iArtLRec) = ArtL
'    PutDataRecord dbArtL, CLng(anzahl)
'
    pzn = ast.pzn
'
    For Durchlauf = 1 To 2
      If Durchlauf = 1 And Not ImportAlt And ImportGruppeNr > 0 Then
        SQLStr = "SELECT * FROM Taxe WHERE ImportgruppeNr=" + CStr(ImportGruppeNr)
      ElseIf (Durchlauf = 2 And Not ImportAlt) Or (Durchlauf = 1 And ImportGruppeNr = 0) Then
        If Anzahl > 1 Then Exit For

        SQLStr = "SELECT * FROM Taxe WHERE OriginalPZN=" + pzn
      Else
        'alte Variante und schon im ersten Durchlauf erledigt
        Exit For
      End If

        On Error Resume Next
        TaxeRec.Close
        Err.Clear
        On Error GoTo DefErr
        TaxeRec.Open SQLStr, taxeAdoDB.ActiveConn

      Do While Not TaxeRec.EOF
        If TaxeRec!pzn <> Val(pzn) Then
          Anzahl = Anzahl + 1
          Call Taxe2ast(PznString(TaxeRec!pzn))
          AVP = ast.AVP
          artPreis = AVP
          'DrS 9.4.2014 Anzeige immer VK - Herstellerrabatt
          '10.6.19 DrS: günstigkeit eines Imports immer nach AVP unabhängig Herstellerrabatt --> artpreis <> AVP
          If TaxeRec!HerstRabattKZ <> 0 Then
            artPreis = FNX(artPreis - HerstRabattBrutto(TaxeRec))
          End If
            If ImRezept Then
              FB = dCheckNull(TaxeRec!FESTBETRAG) / 100#
              If FB > 0 And FB < artPreis Then artPreis = FB '1.4.49
            End If
          ArtL.pzn = Val(ast.pzn)
          ArtL.kurz = ast.kurz
          ArtL.Preis = uFormat(artPreis, Space(Len(ArtL.Preis) - 4) + "0.00")
          ArtL.kz1 = " "
          If ArtL.pzn = Val(pzn) Then
            ArtL.kz2 = "="
          Else
            IstBilliger = False
            If Not NoOriginal Then
              IstBilliger = IstGünstigerImport(artPreis, Preis)
            End If
            If IstBilliger Then
              ArtL.kz2 = " "
              billiger = True
              billigerAbsolut = True
            Else
              ArtL.kz2 = ".-"
              If Not NoOriginal And AVP < OrgAvp Then billigerAbsolut = True
            End If
          End If
          ArtL.GhVerfügbar = 0
          If Val(ast.pzn) = Val(OrgPZN) Then ArtL.AusgangsArtikel = 1 Else ArtL.AusgangsArtikel = 0
          ArtL.rabKZ = " "
          If IstRabattAutIdem(Val(ArtL.pzn)) Then ArtL.rabKZ = "r"
          ArtL.RoteLinie = 0

          ArtL.ArtStatus = TaxeRec!ArtStatus


            iArtLRec = iArtLRec + 1
            ReDim Preserve ArtLRec(iArtLRec)
            ArtLRec(iArtLRec) = ArtL
'          PutDataRecord dbArtL, CLng(anzahl)
        End If
        TaxeRec.MoveNext
      Loop
    Next Durchlauf
'
    If billigerAbsolut = False Then
'      GetDataRecord dbArtL, 1
'      ArtL.kz2 = "  "
'      PutDataRecord dbArtL, 1
      ArtLRec(0).kz2 = "  "
    End If
    If Anzahl > 1 Then
    '    Rabkz, Preis, Kz1 und 1.Stelle Kz2, Kurz
'      Call SortOpenFile(dbArtL.fh, dbArtL.RecLen, 1, CLng(anzahl), "-46,1,33,10,-43,2,5,28", "")
        For i = 1 To Anzahl
            With ArtLRec(i - 1)
                h = Format(999 - Asc(.rabKZ), "000") + .Preis + Format(999 - Asc(.kz1), "000") + Format(999 - Asc(Left(.kz2, 1)), "000") + .kurz '+ vbTab + Str(i)
                .Sort = h
            End With
        Next i

        Dim MinInd%
        For i = 1 To Anzahl
            MinInd = i
            For j = (i + 1) To Anzahl
                If (ArtLRec(j - 1).Sort < ArtLRec(MinInd - 1).Sort) Then
                    MinInd = j
                End If
            Next j
            ArtL = ArtLRec(i - 1)
            ArtLRec(i - 1) = ArtLRec(MinInd - 1)
            ArtLRec(MinInd - 1) = ArtL
    '            h = .List(i - 1)
    '            ind = InStr(h, vbTab)
    '            If (ind > 0) Then
    '                ind2 = Val(Mid(h, ind + 1))
    '                ArtL = ArtLRec(i - 1)
    '                ArtLRec(i - 1) = ArtLRec(ind2 - 1)
    '                ArtLRec(ind2 - 1) = ArtL
    '            End If
        Next i

        'RabKZ, Preis, Kurz
'      ImporteNeu = True
        ImportMarktGilt = True
    End If
    If Anzahl > 0 Then
'      If AutoVerfügbarkeitsabfrage Then
'        PZNDatei = CurDir + "\P" + Format(Now, "nnss") + CStr(ComputerNummer) + Mid("LR", BSSeite + 1, 1) + ".$$$"
'        fhout = fopen(PZNDatei, "o")
'      End If

      For i = 1 To Anzahl
'        GetDataRecord dbArtL, i
            ArtL = ArtLRec(i - 1)
        If ArtL.rabKZ = "r" Then
'          If fhout > 0 And AutoVerfügbarkeitsabfrage Then
'            Print #fhout, PznString(ArtL.PZN)
'            AnzVerfügbarkeitsabfrage = AnzVerfügbarkeitsabfrage + 1
'          End If
        ElseIf Not RoteLinieDa Then

          'If xVal(artl.Preis) > AusgangsPreis Or artl.kz2 = ".-" Then
          'Günstigkeit eines Imports hat nichts mit der Preislinie zu tun! Wenn kein Rabattvertrag, darf sogar teuerster Import abgegeben werden
          'wenn verordnet und es keine nicht-Importe in der Alt+G-Linie gibt
          If xVal(ArtL.Preis) > xVal(uFormat(AusgangsPreis, "########0.00")) Then

            ArtL.RoteLinie = 1
'            PutDataRecord dbArtL, i
            ArtLRec(i - 1) = ArtL
            RoteLinieDa = True
          Else
'            If fhout > 0 And AutoVerfügbarkeitsabfrage Then
'              Print #fhout, PznString(ArtL.PZN)
'              AnzVerfügbarkeitsabfrage = AnzVerfügbarkeitsabfrage + 1
'            End If
          End If
        Else
          Exit For
        End If
      Next i
    End If
    If Anzahl > 1 And NurBilligere Then    '1.2.64
      If billiger Then
      Else
'        ImporteNeu = False
        ImportMarktGilt = 0
        Anzahl = 0
     End If
    End If
    If Anzahl <= 1 Then
'      erg = dbClose(dbArtL)
'      Kill RTrim(dbArtL.FileName)
    Else
'      ArtikelListenArt = LISTE_IMPORT
    End If
Else
    ''''''''''''''''''''''''''''
    Dim satz As Long
    Dim recno As Long
    'Dim menge As Single
    'Dim Füllmenge As Single
    Dim AnyGrün As Boolean
    Dim billigst As Double
    Dim teuerst As Double
    Dim Anz4billigst As Integer
    Dim ViertBilligstPreis As Double
    Dim PreisVorher As Double
    Dim grenze As Double
    Dim OriginalKP As Double
    Dim OriginalAVP As Double
    'Dim art1 As StammdatenType
    Dim AnzGilt As Integer
    Dim t(3) As Double
    Dim b(3) As Double
    Dim zw As Double
    'Dim SQLStr As String
    Dim iSqlStr As String, tSQLStr As String
    Dim TaxeTmpRec As New ADODB.Recordset
    Dim tmpRec As New ADODB.Recordset
    Dim Sonderregelnr As Long
    Dim PznSubst As Boolean
    Dim WirkStoff As String
    Dim PG As String
    Dim PZNok As Boolean
    Dim UnterFB As Boolean
    Dim raus As Boolean
    Dim asatz As Long
    Dim TaxeVK As Double
    Dim GenImpZähler As Integer
    Dim ok As Boolean
    Dim ret As Boolean
    Dim OrgPznIndik As String
    Dim PznIndik As String
    Dim h2 As String
    'Dim h As String
    'Dim ind As Integer
    Dim SqlSonderregelNr As String
    Dim BtmDerivatMenge As Double
    'Dim Preis As Double
    Dim AktHMNr As String
    Dim NoGenerika As Boolean
    'Dim PZNDatei As String
    'Dim fhout As Long
    'Dim AnzVerfügbarkeitsabfrage As Integer
    'Dim RoteLinieDa As Boolean
    'Dim AusgangsPreis As Double
    Dim AnyRVDa As Boolean
    Dim nurImporte As Boolean
    Dim OrgNgr As String
    Dim OrgMenge As Long  'String
    Dim iOk As Boolean

    nurImporte = True

    Anzahl = 0
    ImRezept = True

'    OrgNgr = TaxeRec!NGrösse
'
'    OrgMenge = 0
'    If (TaxeRec!BtmKz) Then
'        OrgMenge = CheckNullLong(TaxeRec!StdMenge)
'    End If

    ImportGr = TaxeRec!ImportGruppeNr
'    If raus Then
'        Call Message("zu diesem Artikel konnten keine Generika gefunden werden!", vbExclamation)
'      Call DefErrPop
'      Exit Sub
'    End If
    If ImportGr = 0 Then
        Call DefErrPop:        Exit Function
    End If
    Call Taxe2ast(pzn)

'      OriginalKP = ast.kp
'      OriginalAVP = ast.AVP
'      If ast.kp > 0 And ast.kp < OriginalAVP Then OriginalAVP = ast.kp
'      'DrS 28.1.2014: bei AutIdems Reihung und Anzeige mit AVP. Bei Importen weiterhin abzügl. Herstellerrabatt
'      'Rahmenvertrag ab 1.7.2019: Herstellerrabatt abziehen
'        If TaxeRec!HerstRabattKZ <> 0 Then
'          OriginalAVP = OriginalAVP - HerstRabattBrutto(TaxeRec)
'        End If
'
'    If Not TaxeRec.EOF Then
''      GenerikaAusgangsArtikel = Trim(TaxeRec!name) + " " + Trim(TaxeRec!menge) + " " + Trim(TaxeRec!einheit) + " (" + Trim(TaxeRec!HerstellerKB) + ")"
'    End If
    BtmDerivatMenge# = 0
    If (BtmDerivatMengeIgnorieren) Then
    Else
        On Error Resume Next
        BtmDerivatMenge# = TaxeRec!BtmDerivatMenge
        If (BtmDerivatMenge >= 0.01) Then
            SQLStr$ = "SELECT * FROM M2Inhalt WHERE M2Nr = " + Mid(TaxeRec!M2, 2, 9)
            On Error Resume Next
            AutIdemRec.Close
            Err.Clear
            On Error GoTo DefErr
            AutIdemRec.Open SQLStr, AutIdemDB.ActiveConn
            If Not (AutIdemRec.EOF) Then
                If (InStr(UCase(CheckNullStr(AutIdemRec!Inhalt)), "HYDROMORPHON") > 0) Then
                    BtmDerivatMenge = 0
                End If
            End If
            AutIdemRec.Close
        End If
        On Error GoTo DefErr
    End If

    SQLStr = ""
    If TaxeRec!AutIdemNeu > 0 Then
        PznSubst = True
        SQLStr = "SELECT * FROM PZNSUBST WHERE PZN = " + CStr(TaxeRec!pzn)
        On Error Resume Next
        AutIdemRec.Close
        On Error GoTo DefErr
        AutIdemRec.Open SQLStr, AutIdemDB.ActiveConn
        SQLStr = ""
        If Not AutIdemRec.EOF Then
            AutIdemRec.MoveFirst
            i = 1
            SQLStr = "SELECT * FROM TAXE WHERE (PZN = "
            Do While Not (AutIdemRec.EOF)
                If (i > 1) Then SQLStr = SQLStr + " OR PZN = "
                SQLStr = SQLStr + CStr(AutIdemRec!Substpzn)
                AutIdemRec.MoveNext
                i = i + 1
                If i > 250 Then Exit Do
            Loop
            If (BtmDerivatMenge >= 0.01) Then
                SQLStr = SQLStr + ") AND (BtmDerivatMenge=" + Str(BtmDerivatMenge)
            End If
        End If
        AutIdemRec.Close
        Set AutIdemRec = Nothing
        If SQLStr > "" Then
    '          If AktKundensatz > 0 And Newline Then BezugSQL = "SELECT TOP 2 * FROM Bezug WHERE Kundennr=" + CStr(AktKundensatz) + " AND " + Mid(SQLStr, 26) + ") ORDER BY Datum DESC"
            SQLStr = SQLStr + ") ORDER BY VK,NAME"
        End If
    End If
    If TaxeRec!AutIdemNeu = 0 Then
'       If ImporteNeu(ast.PZN) Then
        If (ImportMarktGilt(pzn, True)) Then
            ImportMarktGilt = True
            Call DefErrPop
            Exit Function
        End If
    End If

    If SQLStr > "" Then
        On Error Resume Next
        TaxeTmpRec.Close
        Err.Clear
        On Error GoTo DefErr
        TaxeTmpRec.Open SQLStr, taxeAdoDB.ActiveConn
        If (TaxeTmpRec.EOF) Then
            TaxeTmpRec.Close
            Set TaxeTmpRec = Nothing
            Call DefErrPop
            Exit Function
        End If
    Else
      Call DefErrPop
      Exit Function
    End If

    If Not (TaxeTmpRec.EOF) Then
        TaxeTmpRec.MoveLast
    End If
    If (TaxeTmpRec.RecordCount > 1) Then
      TaxeTmpRec.MoveFirst
    End If

    For GenImpZähler = 1 To 2
        '1 ... normale Generikasuche
    
        '2 Importe:
        If GenImpZähler = 2 Then
            If TaxeRec!ImportGruppeNr = 0 Then Exit For
            iSqlStr = "SELECT * FROM TAXE WHERE (PZN = "
            i = 0
            If TaxeRec!ImportGruppeNr > 0 Then
                tSQLStr = "SELECT * FROM Taxe WHERE ImportgruppeNr=" + CStr(TaxeRec!ImportGruppeNr)
                
                On Error Resume Next
                TaxeTmpRec.Close
                Err.Clear
                On Error GoTo DefErr
                TaxeTmpRec.Open tSQLStr, taxeAdoDB.ActiveConn
                Do While Not TaxeTmpRec.EOF
                  If InStr(SQLStr, "PZN = " + CStr(TaxeTmpRec!pzn)) = 0 Then
                    If i > 0 Then iSqlStr = iSqlStr + " OR PZN = "
                    iSqlStr = iSqlStr + CStr(TaxeTmpRec!pzn)
                    i = i + 1
                  End If
                  TaxeTmpRec.MoveNext
                Loop
            End If
            If TaxeRec!OriginalPZN > 0 Then
                tSQLStr = "SELECT * FROM Taxe WHERE OriginalPZN=" + CStr(TaxeRec!OriginalPZN)
            
                On Error Resume Next
                TaxeTmpRec.Close
                Err.Clear
                On Error GoTo DefErr
                TaxeTmpRec.Open tSQLStr, taxeAdoDB.ActiveConn
            
                Do While Not TaxeTmpRec.EOF
                  If InStr(SQLStr, "PZN = " + CStr(TaxeTmpRec!pzn)) = 0 And InStr(iSqlStr, "PZN = " + CStr(TaxeTmpRec!pzn)) = 0 Then
                    If i > 0 Then iSqlStr = iSqlStr + " OR PZN = "
                    iSqlStr = iSqlStr + CStr(TaxeTmpRec!pzn)
                    i = i + 1
                  End If
                  TaxeTmpRec.MoveNext
                Loop
            End If
            If i = 0 Then Exit For   'keine Importe, die nicht eh schon in der Liste wären
            iSqlStr = iSqlStr + ")"
            
            On Error Resume Next
            TaxeTmpRec.Close
            Err.Clear
            On Error GoTo DefErr
            TaxeTmpRec.Open iSqlStr, taxeAdoDB.ActiveConn
        End If

        Do While Not TaxeTmpRec.EOF
          'If TaxeTmpRec!ArtStatus <> "S" Then      'außer Handel-Artikel sollen in die Liste hinein! (RV2019)
            If Not PznSubst Then
                OrgPznIndik$ = GetPznIndik(TaxeRec!pzn)
                ret = True
                If (OrgPznIndik$ > "") And TaxeTmpRec!pzn <> TaxeRec!pzn Then
                    PznIndik$ = GetPznIndik(TaxeTmpRec!pzn)
                    If (PznIndik$ <> "") Then PznIndik$ = "," + PznIndik$
                    
                    h$ = OrgPznIndik$
                    Do
                      ind% = InStr(h$, ",")
                      If (ind% = 0) Then
                        Exit Do
                      Else
                        h2$ = "," + Left$(h$, ind%)
                        h$ = Mid$(h$, ind% + 1)
                        If (InStr(PznIndik$, h2$) = 0) Then
                          ret = False
                          Exit Do
                        End If
                      End If
                    Loop
                End If
                
                '              If (OrgMenge <> 0) Then
                '                  If TaxeTmpRec!StdMenge <> OrgMenge Then ret = False
                '              Else
                '                  If TaxeTmpRec!NGrösse <> OrgNgr Then ret = False
                '              End If
                
                If ret Then
                    If TaxeTmpRec!ImportGruppeNr = 0 Then nurImporte = False
                    If (ImportGr > 0) And (TaxeTmpRec!ImportGruppeNr <> ImportGr) Then nurImporte = False
                    Exit Do
                End If
            Else
'                If (OrgMenge <> 0) Then
'                    iOk = (TaxeTmpRec!StdMenge = OrgMenge)
'                Else
'                    iOk = (TaxeTmpRec!NGrösse = OrgNgr)
'                End If
'                If (iOk) Then
                  If TaxeTmpRec!ImportGruppeNr = 0 Then nurImporte = False
                  If (ImportGr > 0) And (TaxeTmpRec!ImportGruppeNr <> ImportGr) Then nurImporte = False
                  Exit Do
'                End If
            End If
            If Anzahl > 500 Then Exit Do
          'End If
          TaxeTmpRec.MoveNext
        Loop
        If Not nurImporte Then Exit For
    Next GenImpZähler
    
    ImportMarktGilt = nurImporte

    TaxeTmpRec.Close
    Set TaxeTmpRec = Nothing
End If

Call DefErrPop
End Function

Private Sub LennartzRezeptur(modus As Integer, Optional Unique$ = "")
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("LennartzRezeptur")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
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
Dim i%, AnzInhalt%, iFlag%, ErrNo%
Dim SatzPtr&, TmId&
Dim h$, h2$, FaktorKz$, PreisKz$
Dim sXML$, Inhaltsstoffe$, InhSt$, Gefaesse$, Gefaess$, Zuschlaege$, Zuschlag$
               
If (modus = 0) Then
Else
    Set ParenteralTaxierungRec = RezSpeicherDB.OpenRecordset("ParenteralTaxierungen", dbOpenTable)
End If

FaktorKz = "11"
i = 1

sXML = ReadFileToText(LennartzPfad + "\" + LennartzDatei)
Inhaltsstoffe = XmlAbschnitt(sXML, "Inhaltsstoffe")
Do
    InhSt = XmlAbschnitt(Inhaltsstoffe, "InhSt")
    If (InhSt = "") Then
        Exit Do
    End If
    
    With TaxierRec
        .pzn = PznString(XmlAbschnitt(InhSt, "PZN"))
        .kurz = XmlAbschnitt(InhSt, "Bezeichnung")
        .ActMenge = xVal(XmlAbschnitt(InhSt, "Menge"))
        .Meh = XmlAbschnitt(InhSt, "Einheit")
        .kp = XmlAbschnitt(InhSt, "Faktor")
        .menge = XmlAbschnitt(InhSt, "Preiskennzeichen")
        .ActPreis = XmlAbschnitt(InhSt, "Preis")
        .GStufe = 1
        .flag = MAG_ANTEILIG
        
        PreisKz = Format(Val(.menge), "00")
    
        If (modus = 0) Then
            h$ = .pzn + FaktorKz + HashFaktor(.kp) + PreisKz + HashPreis(.ActPreis * 100)
            ParenteralPara = ParenteralPara + h$
        Else
            ParenteralTaxierungRec.AddNew
            
            ParenteralTaxierungRec!RezepteUnique = Unique
            ParenteralTaxierungRec!AnfMagInd = i%
            
            ParenteralTaxierungRec!ActPreis = TaxierRec.ActPreis
            ParenteralTaxierungRec!ActMenge = TaxierRec.ActMenge
            ParenteralTaxierungRec!einheit = TaxierRec.Meh
            ParenteralTaxierungRec!text = TaxierRec.kurz
            ParenteralTaxierungRec!pzn = Val(TaxierRec.pzn)
            ParenteralTaxierungRec!flag = TaxierRec.flag
            ParenteralTaxierungRec!kp = TaxierRec.kp
            ParenteralTaxierungRec!GStufe = TaxierRec.GStufe
            ParenteralTaxierungRec!Verwurf = TaxierRec.Verwurf
            ParenteralTaxierungRec!menge = TaxierRec.menge
            
            ParenteralTaxierungRec.Update
        End If
    End With
    i = i + 1
Loop

Gefaesse = XmlAbschnitt(sXML, "Gefaesse")
Do
    Gefaess = XmlAbschnitt(Gefaesse, "Gefaess")
    If (Gefaess = "") Then
        Exit Do
    End If
    
    With TaxierRec
        .pzn = PznString(XmlAbschnitt(Gefaess, "PZN"))
        .kurz = XmlAbschnitt(Gefaess, "Bezeichnung")
        .ActMenge = xVal(XmlAbschnitt(Gefaess, "Menge"))
        .Meh = XmlAbschnitt(Gefaess, "Einheit")
        .kp = XmlAbschnitt(Gefaess, "Faktor")
        .menge = XmlAbschnitt(Gefaess, "Preiskennzeichen")
        .ActPreis = XmlAbschnitt(Gefaess, "Preis")
        .GStufe = 1
        .flag = MAG_GEFAESS
        
        PreisKz = Format(Val(.menge), "00")
    
        If (modus = 0) Then
            h$ = .pzn + FaktorKz + HashFaktor(.kp) + PreisKz + HashPreis(.ActPreis * 100)
            ParenteralPara = ParenteralPara + h$
        Else
            ParenteralTaxierungRec.AddNew
            
            ParenteralTaxierungRec!RezepteUnique = Unique
            ParenteralTaxierungRec!AnfMagInd = i%
            
            ParenteralTaxierungRec!ActPreis = TaxierRec.ActPreis
            ParenteralTaxierungRec!ActMenge = TaxierRec.ActMenge
            ParenteralTaxierungRec!einheit = TaxierRec.Meh
            ParenteralTaxierungRec!text = TaxierRec.kurz
            ParenteralTaxierungRec!pzn = Val(TaxierRec.pzn)
            ParenteralTaxierungRec!flag = TaxierRec.flag
            ParenteralTaxierungRec!kp = TaxierRec.kp
            ParenteralTaxierungRec!GStufe = TaxierRec.GStufe
            ParenteralTaxierungRec!Verwurf = TaxierRec.Verwurf
            ParenteralTaxierungRec!menge = TaxierRec.menge
            
            ParenteralTaxierungRec.Update
        End If
    End With
    i = i + 1
Loop

Zuschlaege = XmlAbschnitt(sXML, "Zuschlaege")
Do
    Zuschlag = XmlAbschnitt(Zuschlaege, "Zuschlag")
    If (Zuschlag = "") Then
        Exit Do
    End If
    
    With TaxierRec
        .pzn = PznString(XmlAbschnitt(Zuschlag, "PZN"))
        .kurz = ""  'XmlAbschnitt(Zuschlag, "Bezeichnung")
        .ActMenge = 1   'xVal(XmlAbschnitt(Zuschlag, "Menge"))
        .Meh = ""   'XmlAbschnitt(Zuschlag, "Einheit")
        .kp = XmlAbschnitt(Zuschlag, "Faktor")
        .menge = XmlAbschnitt(Zuschlag, "Preiskennzeichen")
        .ActPreis = XmlAbschnitt(Zuschlag, "Preis")
        .GStufe = 1
        .flag = MAG_PREISEINGABE
        
        PreisKz = Format(Val(.menge), "00")
    
        If (modus = 0) Then
            h$ = .pzn + FaktorKz + HashFaktor(.kp) + PreisKz + HashPreis(.ActPreis * 100)
            ParenteralPara = ParenteralPara + h$
        Else
            ParenteralTaxierungRec.AddNew
            
            ParenteralTaxierungRec!RezepteUnique = Unique
            ParenteralTaxierungRec!AnfMagInd = i%
            
            ParenteralTaxierungRec!ActPreis = TaxierRec.ActPreis
            ParenteralTaxierungRec!ActMenge = TaxierRec.ActMenge
            ParenteralTaxierungRec!einheit = TaxierRec.Meh
            ParenteralTaxierungRec!text = TaxierRec.kurz
            ParenteralTaxierungRec!pzn = Val(TaxierRec.pzn)
            ParenteralTaxierungRec!flag = TaxierRec.flag
            ParenteralTaxierungRec!kp = TaxierRec.kp
            ParenteralTaxierungRec!GStufe = TaxierRec.GStufe
            ParenteralTaxierungRec!Verwurf = TaxierRec.Verwurf
            ParenteralTaxierungRec!menge = TaxierRec.menge
            
            ParenteralTaxierungRec.Update
        End If
    End With
    i = i + 1
Loop
        
Call DefErrPop
End Sub

Private Sub CheckPznVorhanden(lSuchPzn As Long, sText As String, dPreis As Double)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("CheckPznVorhanden")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
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
Dim lPzn&
Dim bGef As Boolean

bGef = False
For i% = 0 To (AnzRezeptArtikel% - 1)
    lPzn = Val(RezArtikel(i%).pzn)
    If (lPzn = lSuchPzn) Then
        bGef = True
        Exit For
    End If
Next i
If Not (bGef) Then
    AnzRezeptArtikel = AnzRezeptArtikel + 1
    With RezArtikel(AnzRezeptArtikel% - 1)
        .text = sText
        .pzn = PznString(lSuchPzn)
        .Faktor = 1
        .Preis = dPreis
'        RezArtikel(row%).flag = RezArtikel(row%).flag And (REZ_BLINK Xor &HFF)
    End With
End If
        
Call DefErrPop
End Sub

Sub PaintSelbsterklaerung(sAnzahl As String)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("PaintSelbsterklaerung")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
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

WinVkAnzahl = xVal(sAnzahl)
RezArtikel(0).Preis = WinVkAnzahl / 100#
Call PaintRezeptBox(6)
Call PaintRezeptWert(6, Format(RezArtikel(0).Preis * 100, "0"), 0)
Call CalcRezeptWerte

Call DefErrPop
End Sub


Sub SelbsterklaerungAbrechnung()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("SelbsterklaerungAbrechnung")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
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
Dim Sonderbeleg As New Sonderbeleg_OP
Dim sVerordnung As String
Dim sTaskId As String

WinVkAnzahl = RezArtikel(0).Preis * 100

With Sonderbeleg
    With .Patient
        .Id = "A000000002"
        .GeburtsDatum = CDate("1.1.1900")
        With .Name
            .Vorname = "Sonderbeleg"
            .Nachname_ohne_Vor_und_Zusatz = "NNF"
            .Titel = ""
        End With
        With .Adresse
            .PLZ = "00000"
'            .Ort = ""
'            With .Strasse1
'                .Strasse = " "
'                .Hausnummer = " "
'            End With
            .Land = "D"
        End With
    End With

    With .Kostentraeger
        .Typ = "PKV"
        .IK = "661100310"
        .Name = "Nacht- und Notdienstfonds des DAV"
    End With
End With

If (FDok = 0) Then
    Set FD_OP = New TI_Back.Fachdienst_OP
    If (FD_OP.Init(FD_OP.NetWindowHandle(frmAction.hWnd), (wpara.EntwicklungsUmgebung = 0))) Then
        FDok = True
    End If
End If

With Sonderbeleg
    .pzn = "02567768"
    .Verordnungstext = "Nacht- und Notdienstfonds des DAV"
    
    .FlowType = 990
    .ImpfleistungPZNs = ""
    .sGesamtBrutto = ""
    .sGesamtZuzahlung = ""
    .SelbsterklaerungAnzahl = WinVkAnzahl
    
    .VerordnungsDatum = CDate("01." + Mid(VerkDatum, 3, 2) + ".20" + Mid(VerkDatum, 5, 2))
    
    .Typ = "Selbsterklaerung"
End With

sTaskId = FD_OP.TI_SonderBeleg(Sonderbeleg)
Set Sonderbeleg = Nothing

If (sTaskId <> "") Then
    Dim bErgebnis As Boolean

    Dim eRezept As New TI_Back.eRezept_OP
    Call eRezept.New2(sTaskId)
'    eRezept.ImpfleistungPZNs = Sonderbeleg.ImpfleistungPZNs
                            
    If (FD_OP.TI_RezeptAbrechnen(eRezept, False)) Then
'                    MsgBox (IIf(SollStatus = 2, "Vorabprüfung: ", "Einreichung: ") + CStr(VorabPruefung(sTaskId, (SollStatus = 2))))
        bErgebnis = VorabPruefung(sTaskId, False)
        If (bErgebnis) Then
            ParentForm.cmdEsc(0).Value = True
        End If
    Else
        If (bEinzelErgebnis) Then
            Call MessageBox("Problem beim Erstellen der Abgabedaten", vbCritical)
        End If
    End If
End If
        
Call DefErrPop
End Sub


