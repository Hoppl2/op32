VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsWinRezDr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Const INI_SECTION = "Rezeptkontrolle"
Const INFO_SECTION = "Infobereich Rezeptkontrolle"

Const REZEPTART_KASSE = 1
Const REZEPTART_ZUZAFREI = 2
Const REZEPTART_BTM = 3
Const REZEPTART_BTM_ZUZAFREI = 4
Const REZEPTART_PRIVAT = 5
Const REZEPTART_GRÜN = 6
Const REZEPTART_IVF = 7
Const REZEPTART_SPRECHSTUNDENBEDARF = 8

Const MAX_VERK_ZEILEN% = 10

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
    Appli As Byte
    Faktor As Integer
    zusatz(1) As String * 36
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

Private Type VERKAUF_STRUCT
    pzn As String * 7
    filler1 As String * 1
    text As String * 36
    filler2 As String * 1
    Preis As String * 9
    filler3 As String * 1
    flag As Byte
    Imp123 As Byte
    Zuz As Single
    Appli As Byte
End Type


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
    Ean As String * 13
    dummy As String * 1
End Type
Private Type REZKONTR_STRUCT
    Fertig As String * 1
    Preis(9) As Double
    RezEan(9) As REZKONTR2_STRUCT
End Type

'Private Type IMPFUNG_STRUCT
'    pzn As String * 8
'    text As String * 28
'End Type

Private Type ABGABE_KOSTEN_STRUCT
    Name As String * 12
    kosten As Double
    ind As Integer
End Type
    


Dim GebFrei%
Dim ImpfRez%
Dim PrivatRezept%
Dim BtmRezept%
Dim IvfRezept%
Dim SchonDa%

Dim ActStueck%, ActMenge%, ActMwst%, ActVmRabatt%
Dim ActMietTage%, ActMietWochen%, ActMietMonate%, ActMietDauer%, ActMietZeit%
Dim ActPatientenAlterMonate%, ActPatientenAlterJahre%
Dim ActWert#, ActAep#, ActAvp#, ActVmPreis#, Act20Preis#
Dim ActWarenzeichen$


Dim RezSumme#, RezSummeDazu#, RezGebSumme#
Dim RezGebRechnen%

Dim VerkUser$
Dim VerkLiRe As Byte

Dim RezKundenInfo$(4)

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

'Dim RezeptBoxen(52) As FORMULAR_STRUCT
'Dim BoxenWert$(52)  '36
'Dim MaxPruefBox%


Dim BarVerkauf$(2)

Dim VerkPtr&, OrgVerkPtr&
Dim lPrivEnde&

Dim bmVerkPtr As Variant
Dim bmOrgVerkPtr As Variant
Dim bmPrivEnde As Variant

Dim lRezNr&

Dim PersonalNr%


Dim BeepFlag%
Dim DruckDatum$
Dim VerkDatum$, OrgVerkDatum$
Dim VerkZeit%

Dim AnzRezeptArtikel%

Dim ZusatzKomponentenNr$, ZusatzKomponentenBasisNr$

Dim AktTabIndex%

Dim AvLaender$()
Dim AvKassen$()
Dim AvVerordnungen$()

Dim SonderPzn$()

Dim AbgabeKosten() As ABGABE_KOSTEN_STRUCT
Dim F7Kassen() As ABGABE_KOSTEN_STRUCT

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

Dim SollTmName$(MAX_VERK_ZEILEN)
Dim SollTmMenge#(MAX_VERK_ZEILEN)

Dim TransaktionsID&
Dim rzLieferId$

Dim AvAbrechnungsNr$(2)
Dim AvHmPositionsNr$(2)

Dim ActPreis#(2), ActPreis20#(2)
Dim ActAVrabatt%(2)

Dim AMverfügbarkeit$
Dim WVORabattArtikel$

Dim BtmRezeptAlt%

Private Const DefErrModul = "WINREZDR.CLS"

Public Property Let SetVerkPtr(NewValue&)
VerkPtr& = NewValue&
End Property
Public Property Get GetPrivEnde&()
GetPrivEnde& = lPrivEnde&
End Property

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
Dim i%, iFaktor%, BlinkFlag%, AnzArtikel%, ind%
Dim GesHerstRabattPrivat130Brutto#

AnzBlink% = 0

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
                        
'                        j% = 3 + i * 4
'                        Call PaintRezeptBox(j% + 3)
'                        Call PaintRezeptWert(j% + 3, Format(RezArtikel(i%).Preis * 100 * RezArtikel(i%).Faktor, "0"), 0)
                    End If
                End If
            End If
        Next i
    End If
    

RezSumme# = RezSummeDazu#
'For i% = 0 To UBound(RezArtikel)
AnzArtikel% = (AnzRezeptArtikel% - 1)     'GS 16.05.02
If AnzArtikel% > 10 Then AnzArtikel% = 10
    
    
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

'*** Ver 5.3.0 am 24.8.23
    Dim sVerfügbarkeit2$
'    sVerfügbarkeit2 = CheckVerfügbarkeit
'    MsgBox (sVerfügbarkeit2)
'    If (sVerfügbarkeit2 <> "111") Then
'        For i = 1 To 3
'            If (InStr("234", Mid(sVerfügbarkeit2, i, 1)) > 0) Then
'            MsgBox ("davor: " + CStr(RezSumme))
'                RezSumme = RezSumme + 0.6
'            MsgBox ("danach: " + CStr(RezSumme))
'            End If
'        Next i
'    End If
    
    sVerfügbarkeit2 = Left(AMverfügbarkeit, 3)

    'MsgBox (sVerfügbarkeit2)
    If (sVerfügbarkeit2 <> "111") Then
        For i = 1 To 3
            If (InStr("234", Mid(sVerfügbarkeit2, i, 1)) > 0) Then
            'MsgBox ("davor: " + CStr(RezSumme))
                RezSumme = RezSumme + 0.6
            'MsgBox ("danach: " + CStr(RezSumme))
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
'        Call PaintArtikel(ind, 0)
    ElseIf (ind >= 0) Then
        For i% = ind% To AnzArtikel
            RezArtikel(i%) = RezArtikel(i% + 1)
'            Call PaintArtikel(i, 0)
        Next i%
        Call InitRezeptArtikel(UBound(RezArtikel))
        
        AnzRezeptArtikel% = AnzRezeptArtikel% - 1
'            Call PaintRezept
'            Call TabIndexChange(3 + (ind% * 4), True)
    End If
End If

'For i% = 0 To AnzArtikel%
'    RezSumme# = RezSumme# + (RezArtikel(i%).Preis * RezArtikel(i%).Faktor)
'Next i%

If (RezGebRechnen%) Then
    RezGebSumme# = 0
    If Not (GebFrei%) Then
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
            RezGebSumme# = RezGebSumme# + (RezArtikel(i%).Zuz * iFaktor%)
        Next i%
    End If
End If

AnzBlink% = 0
If (lRezNr& <> 0&) And (PrivatRezept% = 0) And (SonderBelegRezept% = 0) Then
    For i% = 0 To (AnzRezeptArtikel% - 1)
        BlinkFlag% = RezArtikel(i%).flag And REZ_BLINK
        If (BlinkFlag%) Then AnzBlink% = AnzBlink% + 1
    Next i%
End If
If (Trim(RezApoNr$) = "") Then AnzBlink% = 1

If (DSKNurRezeptnummerDrucken%) Then AnzBlink% = 1

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
Dim PrivUser As Byte, PrivLiRe As Byte, pGebuehren As Byte
Dim lAnzVkSaetze&, lPrivLauf&, iKuNr&
Dim dZuz#, dZuzDos#, dPreis#, dAvp#, PrivatBesorgerPreis#
Dim ch$, RezEan$, hStr$, LetztArtikel$, AktArtikel$, wg$, SQLStr$, Unique$, h$
Dim aktavp As Double, OrgAvp As Double

Static fOrgRezNr!

If (VerkaufDbOk%) Then
    RezeptHolen = RezeptHolenMDB(PrivatModus%)
    Call DefErrPop: Exit Function
End If

h$ = RezNr$

Call InitRezept(1)

RezNr$ = Trim(h$)
l% = Len(RezNr)

If (UCase$(Left$(RezNr, 1)) = "P") Then
    VerkPtr& = Val(Mid$(RezNr, 2))
    OrgVerkPtr& = VerkPtr&

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
        lRezNr& = Val(RezNr$)
'        If (lRezNr& <= 0) Or (lRezNr& > 32000) Then
        If (lRezNr& <= 0) Then
            RezeptHolen% = False
            Call DefErrPop: Exit Function
        End If
        hStr$ = "9999998" + Format(lRezNr&, "00000")
        Call PruefZiffer(hStr$)
        RezNr$ = hStr$
    End If

'    lRezNr& = CLng(RezNr$)

    VerkPtr& = 0&
    
    FabsErrf% = RezTab.IndexSearch(0, RezNr$, FabsRecno&)
    If (FabsErrf% = 0) Then
        Call RezTab.GetRecord(FabsRecno&)
        VerkPtr& = RezTab.VerkPtr + 1
    End If

    If (RezeptHuellen%) Then
        SchonDa% = False
    ElseIf (RezSpeicherOK%) Then
        SchonDa% = False
    End If
    
    If (SchonDa% = False) And (VerkPtr& <= 0) Then
        RezeptHolen% = False
        Call DefErrPop: Exit Function
    End If

End If


If (SchonDa%) Then
'    Call HoleAusRezeptSpeicher
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
            ElseIf (vk.pzn = "9999999") And (Left$(UCase(vk.text), 10) = "IK-NUMMER:") Then
                h$ = Trim(vk.KKNr)
                If (AutIdemSonderregelOk%) And (Val(h$) > 0) Then
                    SQLStr$ = "SELECT * FROM kKassen WHERE IK = " + Str$(Val(h$))
'                    Set KkassenRec = kKassenDB.OpenRecordset(SQLStr$)
'                    If (KkassenRec.RecordCount > 0) Then
                    FabsErrf = kKassenDB.OpenRecordset(KkassenRec, SQLStr)
                    If Not (KkassenRec.EOF) Then
                        AutIdemIk& = CheckNullLong(KkassenRec!IK)
                        AutIdemKbvNr& = CheckNullLong(KkassenRec!KbvNr)
                    End If
                End If
            ElseIf (PrivatBesorger% = 1) Then
                k% = k% + 1
                RezArtikel(k%).pzn = vk.pzn
                RezArtikel(k%).text = vk.text
                RezArtikel(k%).Preis = Val(vk.Preis)
                RezArtikel(k%).flag = 0
                If Chr(vk.AVP) = "A" Then RezArtikel(k%).Appli = vk.AVP
                RezArtikel(k%).Zuz = 0
                RezArtikel(k%).Imp123 = 0
                RezArtikel(k%).Fam = 0
                RezArtikel(k%).AutIdem = 0
                RezArtikel(k%).MagIndex = 0
                RezArtikel(k%).Faktor = 1
                RezArtikel(k%).NichtInTaxe = 0
                
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
                If Chr(vk.AVP) = "A" Then RezArtikel(k%).Appli = vk.AVP
                RezArtikel(k%).Zuz = 0
'                If (IvfRezept) Then
'                    RezArtikel(k%).Zuz = VK.gebsumme
'                End If
                RezArtikel(k%).Imp123 = 0
                RezArtikel(k%).Fam = 0
                RezArtikel(k%).AutIdem = 0
                RezArtikel(k%).MagIndex = 0
                RezArtikel(k%).Faktor = 1
                RezArtikel(k%).NichtInTaxe = 0
            
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
                        RezArtikel(k%).Appli = (vk.gebsumme = 0)
                    End If
                End If
                
                RezArtikel(k%).IstWg4 = 0
                RezArtikel(k%).PlusMehrKosten = 0
                RezArtikel(k%).ZuzahlungsErlass = 0
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
                        FabsErrf% = VmPzn.IndexSearch(0, vk.pzn, FabsRecno&)
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
'            If (iKuNr% > 0) Then Call HoleKundenInfo(iKuNr%)


            If (vk.EndeT And &H2) Then
                RezGebSumme# = 0#
                GebFrei% = True
            End If

            If (vk.KKTyp > 0) Then ActKasse% = vk.KKTyp
            
            ActKkasse$ = Trim(vk.KKNr)


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
    If (SonderBelegRezept% > 0) Then
        If (iKuNr > 0) Then Call HoleKundenInfo(iKuNr)
    End If



    If Not (vk.EndeT And &H4) And Not (PrivatRezept%) And (SonderBelegRezept% = 0) Then
        For i% = 0 To UBound(BarVerkauf$)
            Call vk.GetRecord
    
            hhRez% = vk.RezNr
            If (hhRez% = 0) Then
                hStr$ = vk.text + " " + vk.Preis
                BarVerkauf$(i%) = hStr$
            End If
    
            If (hhRez% <> 0) Or (vk.EndeT And &H4) Then Exit For
        Next i%
    End If


    For i% = 0 To (AnzRezeptArtikel% - 1)
        If (PrivatRezept%) Then
            RezArtikel(i%).flag = RezArtikel(i%).flag And (REZ_BLINK Xor &HFF)
        End If
        
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
                    If (RezArtikel(i%).Appli = 0) Then
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
                    
                    If (TaxeRec!NegListe) Then dZuzDos# = 1#
                    If (RezArtikel(i%).Appli) Then dZuzDos# = 0#
                    
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
                
                    If (ImpFlag% = 0) Then
                        RezArtikel(i%).Imp123 = 0
                    ElseIf (TaxeRec!OriginalPZN = TaxeRec!pzn) Then
                        RezArtikel(i%).Imp123 = 2
                        OrgAvp = TaxeRec!vk / 100#
                        SQLStr$ = "SELECT * FROM TAXE WHERE ORIGINALPZN = " + Format(TaxeRec!OriginalPZN, "0000000")
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
                              RezArtikel(i%).Imp123 = 1
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
                        SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + Format(TaxeRec!OriginalPZN, "0000000")
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
                
                    If (TaxeRec!HimiVerbrauch) Then
                        dZuz# = RezArtikel(i%).Zuz / RezArtikel(i%).Faktor
                    End If
                
                    If (RezArtikel(i%).ZuzahlungsErlass) Then
                        dZuz# = RezArtikel(i%).Zuz
                    End If
                
                    If (TaxeRec!BTMKz) Then
                        BtmRezept% = True
                    End If
                    
                    hStr$ = Format$(TaxeRec!Abgabebest, "0")
                    If (InStr("01245", hStr$) <> 0) Then
                        RezArtikel(i%).Fam = 1
                    End If
                
                    Call CheckAutIdem(i%)
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
        Call PruefeBtmKosten
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

'ActVerordnung% = 1

Call CalcRezeptWerte


RezeptHolen% = True

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
Dim i%, j%, k%, l%, ind%, PrivGef%, iUser%, iVerkMult%, hhRez%, IstDrüber%, ind2%, IstWVORabattArtikel%
Dim pTag%, pMonat%, pJahr%, PrivDatum%, iVerkDatum%
Dim PrivUser As Byte, PrivLiRe As Byte, AktPrivRezNr As Byte
Dim lAnzVkSaetze&, lIkNr&, VerkId&, iKuNr&, ll&
Dim dZuz#, dZuzDos#, dPreis#, dAvp#
Dim ch$, RezEan$, hStr$, LetztArtikel$, AktArtikel$, wg$, SQLStr$, Unique$, h$, sPrivDatum$, sPzn$, sText$
Dim aktavp As Double, OrgAvp As Double
Dim iiDat As Date
Dim AvpRec As Recordset

Static fOrgRezNr!
'Static lPrivEnde&

Call WinArtDebug("RezeptHolenMDB: " + RezNr)

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
    VerkId& = Val(Mid$(RezNr, 2))

    bmVerkPtr = Null
    Set VerkaufRec = VerkaufDB.OpenRecordset("Verkauf", dbOpenTable)
    VerkaufRec.index = "Id"
    
    Call WinArtDebug("Suche Privat")

    VerkaufRec.Seek "=", VerkId&
    If (VerkaufRec.NoMatch = False) Then
        bmVerkPtr = VerkaufRec.Bookmark
    Else
        Call iMsgBox("keine weiteren passenden Privat-Rezepte gespeichert !")
        RezeptHolenMDB% = False
        Call DefErrPop: Exit Function
    End If

    If (VerkaufRec!PrivRezLaufNr <= 0) Then
        h$ = "00001"
        ll& = GetPrivateProfileString("Rezeptkontrolle", "NeueRezepte", h$, h$, 6, INI_DATEI)
        h$ = Left$(h$, ll&)
        lRezNr& = Val(h$)
        RezNr$ = Format(lRezNr&, "00000")
    
        h$ = Format(lRezNr& + 1, "0")
        ll& = WritePrivateProfileString("Rezeptkontrolle", "NeueRezepte", h$, INI_DATEI)
    Else
        lRezNr = VerkaufRec!PrivRezLaufNr
    End If
        
    h$ = "9999998" + Format(lRezNr&, "00000")
    Call PruefZiffer(h$)
    RezNr$ = h$
    
    
    Call WinArtDebug("vor Umschalten")
    
    Set VerkaufRec = VerkaufDB.OpenRecordset("Verkauf", dbOpenTable)
    VerkaufRec.index = "Unique"
    
    bmOrgVerkPtr = bmVerkPtr
    
    bmPrivEnde = bmVerkPtr
    lRezNr& = -1&
    PrivatRezept% = True
    GebFrei% = True
Else

    Set VerkaufRec = VerkaufDB.OpenRecordset("Verkauf", dbOpenTable)
    VerkaufRec.index = "RezeptNr"
    
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
'        If (lRezNr& <= 0) Or (lRezNr& > 32000) Then
        If (lRezNr& <= 0) Then
            RezeptHolenMDB% = False
            Call DefErrPop: Exit Function
        End If
        hStr$ = "9999998" + Format(lRezNr&, "00000")
        Call PruefZiffer(hStr$)
        RezNr$ = hStr$
    End If

'    lRezNr& = CLng(RezNr$)

'    ParentForm.txtRezeptNr.text = RezNr$
    
    Call WinArtDebug("Suche ELSE")

    bmVerkPtr = Null
    
'    VerkaufRec.Seek "=", RezNr$
'    If (VerkaufRec.NoMatch = False) Then
'        bmVerkPtr = VerkaufRec.Bookmark
'    End If
    VerkaufRec.Seek ">=", RezNr$    ', Now
    If (VerkaufRec.NoMatch = False) Then
        If (VerkaufRec!RezeptNr = RezNr$) Then
            bmVerkPtr = VerkaufRec.Bookmark
        End If
    End If

    Call WinArtDebug("nach Suche")
    
    If (RezeptHuellen%) Then
        SchonDa% = False
    ElseIf (RezSpeicherOK%) Then
        SchonDa% = False
    End If
    
    If (SchonDa% = False) And (IsNull(bmVerkPtr)) Then
        RezeptHolenMDB% = False
        Call DefErrPop: Exit Function
    End If

End If

Call WinArtDebug("vor Auslesen")
    

If (SchonDa%) Then
'    Call HoleAusRezeptSpeicher
Else
    Call InitRezept(2)
    AMverfügbarkeit = String(10, "1")


    LetztArtikel$ = ""
    k% = -1
    
    Call WinArtDebug("vor Bookmark")
    
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
    
    Call WinArtDebug("vor DO")
    
    Do
        sPzn = PznString(VerkaufRec!pzn)
        
        Call WinArtDebug("Artikel: " + sPzn)
    
        sText = VerkaufRec!text
        dPreis = VerkaufRec!Preis
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
                
                If (VerkaufRec!HimiVerbrauch) Then
'                    RezArtikel(k%).Zuz = RezArtikel(k%).Zuz + VerkaufRec!Zuzahlung
                End If
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
                RezArtikel(k%).Appli = Abs(InStr(UCase(VerkaufRec!Rezeptzeichen), "A") > 0)  'Abs(VerkaufRec!Hilfsmittel)
                
                RezArtikel(k%).Zuz = 0
                RezArtikel(k%).HerstRabattPrivat130Brutto = CheckNullDouble(VerkaufRec!HerstRabattPrivat130Brutto)
'                If (IvfRezept) Then
'                    RezArtikel(k%).Zuz = VerkaufRec!Zuzahlung
'                    RezArtikel(k%).Preis = VerkaufRec!VK
'                End If
                
                RezArtikel(k%).Imp123 = 0
                RezArtikel(k%).Fam = 0
                RezArtikel(k%).AutIdem = 0
                RezArtikel(k%).AutIdemKreuz = CheckNullByte(VerkaufRec!AutIdemKreuz)
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
                            RezArtikel(k%).Appli = 1
                        End If
                    End If
                End If
                
                RezArtikel(k%).IstWg4 = 0
                RezArtikel(k%).PlusMehrKosten = 0
                RezArtikel(k%).ZuzahlungsErlass = 0
                RezArtikel(k%).Warenzeichen = 0
                RezArtikel(k%).MietDauer = 0
                RezArtikel(k%).TkkPznDruck = 0
                
                RezArtikel(k%).VebNr = CheckNullLong(VerkaufRec!VebNr)
                RezArtikel(k%).PauschaleNr = CheckNullLong(VerkaufRec!PauschaleNr)
                
                If (CheckNullLong(VerkaufRec!Verfügbarkeit) > 1) Then
                    Mid(AMverfügbarkeit, k + 1, 1) = CStr(VerkaufRec!Verfügbarkeit)
                End If

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
                
                IstWVORabattArtikel = 0
                On Error Resume Next
'                IstWVORabattArtikel = CheckNullLong(VerkaufRec!WVORabattArtikel)
                IstWVORabattArtikel = CheckNullLong(VerkaufRec!WirkstoffVerordnung)
                On Error GoTo DefErr
                If (IstWVORabattArtikel <> 0) Then
                    Mid(WVORabattArtikel, k + 1, 1) = "2"
                End If
                
                RezArtikel(k%).Fortsetzung = 0
                If (CheckNullLong(VerkaufRec!Fortsetzung)) Then
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
                    Call WinArtDebug("Artikel: in Taxe")
                    
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
                        If (CheckNullDouble(VerkaufRec!ZuzaGes) > 0) Then
                            RezArtikel(k%).Zuz = VerkaufRec!ZuzaGes / RezArtikel(k%).Faktor
                        Else
                            RezArtikel(k%).Zuz = VerkaufRec!Zuzahlung
                        End If
                    End If
                    
                    If (VerkaufRec!MehrkostenVerzichtKZ) Then
                        RezArtikel(k%).PlusMehrKosten = 1
                    End If
                    If (VerkaufRec!ZuzahlungsErlassKZ) Then
                        RezArtikel(k%).ZuzahlungsErlass = 1
                        RezArtikel(k%).Zuz = VerkaufRec!Zuzahlung
                    End If
                    If (VerkaufRec!Mehrkosten > 0) Then
                        If (VerkaufRec!Hilfsmittel) Then
                            RezArtikel(k).zusatz(1) = "MK " + CStr(VerkaufRec!Mehrkosten)
                        End If
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
                Call WinArtDebug("Artikel: vor HoleKundenInfo")
                Call HoleKundenInfo(iKuNr)
                Call WinArtDebug("Artikel: nach HoleKundenInfo")
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


    If Not (VerkaufRec!KundenEnde) And Not (PrivatRezept%) And (SonderBelegRezept% = 0) Then
        For i% = 0 To UBound(BarVerkauf$)
            VerkaufRec.MoveNext
            If (VerkaufRec.EOF) Then
                Exit For
            End If

            If (VerkaufRec!RezeptArt = 0) Then
                hStr$ = VerkaufRec!text + " " + Format(VerkaufRec!Preis, "0.00")
                BarVerkauf$(i%) = hStr$

'                With ParentForm
'                    .mnuBearbeitenInd(MENU_SF3).Enabled = True
'                    .cmdToolbar(MENU_SF3).Enabled = True
'                End With
            End If

            If (VerkaufRec!KundenEnde) Then
                Exit For
            End If
        Next i%
    End If

'    If (PrivatRezept%) Or (SonderBelegRezept%) Then
'    Else
'        AMverfügbarkeit = Left$(String(AnzRezeptArtikel%, "1") + "000", 3)
'    End If

    Call WinArtDebug("vor 2.Schleife")
    
    For i% = 0 To (AnzRezeptArtikel% - 1)
        dZuz# = 0#
        dAvp# = 0
        
        Call WinArtDebug("in 2.Schleife")
    
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
                ElseIf (TaxeRec!BTMKz) Then
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
                    If (RezArtikel(i%).Appli = 0) Then
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
                        dZuz# = CalcZuzaNeu(RezArtikel(i%).Preis)
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
                    Call WinArtDebug("in 2.Schleife, Taxe")
    
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
                    
                    If (TaxeRec!NegListe) Then dZuzDos# = 1#
                    If (RezArtikel(i%).Appli) Then dZuzDos# = 0#
                    
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
                        dZuz# = RezArtikel(i%).Zuz
                    End If
                
                    If (TaxeRec!BTMKz) Then
                        BtmRezept% = True
                    End If
                    
                    hStr$ = Format$(TaxeRec!Abgabebest, "0")
                    If (InStr("01245", hStr$) <> 0) Then
                        RezArtikel(i%).Fam = 1
                    End If
                
                    Call WinArtDebug("in 2.Schleife, vor CheckAutIdem")
                    Call CheckAutIdem(i%)
                    Call WinArtDebug("in 2.Schleife, nach CheckAutIdem")
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
                        Call WinArtDebug("in 2.Schleife, vor Original")
                        RezArtikel(i%).Imp123 = 6
                        OrgAvp = TaxeRec!vk / 100#
                        SQLStr$ = "SELECT * FROM TAXE WHERE ORIGINALPZN = " + Format(TaxeRec!OriginalPZN, "0000000")
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
                        Call WinArtDebug("in 2.Schleife, nach Original")
                        SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + RezArtikel(i%).pzn
                        'Set TaxeRec = TaxeDB.OpenRecordset(SQLStr$)
                        On Error Resume Next
                        TaxeRec.Close
                        Err.Clear
                        On Error GoTo DefErr
                        TaxeRec.Open SQLStr, taxeAdoDB.ActiveConn
                    
                    ElseIf (TaxeRec!OriginalPZN > 0) Then
                        aktavp = TaxeRec!vk / 100
                        SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + Format(TaxeRec!OriginalPZN, "0000000")
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
'        For i% = 0 To UBound(RezAddGesamtBrutto$)
'            RezAddGesamtBrutto$(i%) = ""
'        Next i%
        Call PruefeBtmKosten
'        sBrutto$ = RezAddGesamtBrutto$(0)
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
        Call WinArtDebug("Verbandmittel")
        For i% = 0 To (AnzRezeptArtikel% - 1)
            If (RezArtikel(i%).flag And REZ_WG_4) Then
                RezeptHolenZuz# = RezArtikel(i%).Zuz
                AktTabIndex% = 3 + i% * 4
                Call Verbandmittel
            End If
        Next i%
    End If
    RezeptHolenZuz# = 0
    
    Call CheckVerordnungMDB
End If

'ActVerordnung% = 1

For i% = 0 To (AnzRezeptArtikel% - 1)
    If (Mid(AMverfügbarkeit, i + 1, 1) = "7") Then
        With RezArtikel(i)
            RezArtikel(i).Preis = 0#
            RezArtikel(i).Zuz = 0
        End With
    End If
Next
Call CalcRezeptWerte

Call WinArtDebug("Ende RezeptHolenMDB")

RezeptHolenMDB% = True

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
Dim i%, j%, k%, l%, ind%, PrivGef%, iUser%, iVerkMult%, hhRez%, IstDrüber%, ind2%, IstWVORabattArtikel%
Dim pTag%, pMonat%, pJahr%, PrivDatum%, iVerkDatum%
Dim PrivUser As Byte, PrivLiRe As Byte, AktPrivRezNr As Byte
Dim lAnzVkSaetze&, lIkNr&, VerkId&, iKuNr&, ll&
Dim dZuz#, dZuzDos#, dPreis#, dAvp#
Dim ch$, RezEan$, hStr$, LetztArtikel$, AktArtikel$, wg$, SQLStr$, Unique$, h$, sPrivDatum$, sPzn$, sText$
Dim aktavp As Double, OrgAvp As Double
Dim iiDat As Date
Dim iiLaufNr As Long
Dim AvpRec As Recordset

Static fOrgRezNr!
'Static lPrivEnde&

Call WinArtDebug("RezeptHolenDB: " + RezNr)

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
    VerkId& = Val(Mid$(RezNr, 2))

    bmVerkPtr = Null
'    Set VerkaufRec = VerkaufDB.OpenRecordset("Verkauf", dbOpenTable)
'    VerkaufRec.Index = "Id"
    
    Call WinArtDebug("Suche Privat")

'    VerkaufRec.Seek "=", VerkId&
'    If (VerkaufRec.NoMatch = False) Then
'        bmVerkPtr = VerkaufRec.Bookmark
'    Else
'        Call iMsgBox("keine weiteren passenden Privat-Rezepte gespeichert !")
'        RezeptHolenDB% = False
'        Call DefErrPop: Exit Function
'    End If
    
    SQLStr = "SELECT * FROM Verkauf WHERE Id=" + CStr(VerkId)
'    SQLStr = SQLStr + " ORDER BY Datum DESC,LaufNr DESC,ZeilenNr"
    FabsErrf = VerkaufAdoDB.OpenRecordset(VerkaufAdoRec, SQLStr, 0)
    If (VerkaufAdoRec.EOF) Then
        Call iMsgBox("keine weiteren passenden Privat-Rezepte gespeichert !")
        RezeptHolenDB% = False
        Call DefErrPop: Exit Function
    End If
    
    bmVerkPtr = VerkaufAdoRec.Bookmark
'    iiDat = VerkaufAdoRec!Datum
'    iiLaufNr = VerkaufAdoRec!LaufNr

    If (VerkaufAdoRec!PrivRezLaufNr <= 0) Then
        h$ = "00001"
        ll& = GetPrivateProfileString("Rezeptkontrolle", "NeueRezepte", h$, h$, 6, INI_DATEI)
        h$ = Left$(h$, ll&)
        lRezNr& = Val(h$)
        RezNr$ = Format(lRezNr&, "00000")
    
        h$ = Format(lRezNr& + 1, "0")
        ll& = WritePrivateProfileString("Rezeptkontrolle", "NeueRezepte", h$, INI_DATEI)
    Else
        lRezNr = VerkaufAdoRec!PrivRezLaufNr
    End If
        
    h$ = "9999998" + Format(lRezNr&, "00000")
    Call PruefZiffer(h$)
    RezNr$ = h$
    
    
    Call WinArtDebug("vor Umschalten")
    
'    Set VerkaufRec = VerkaufDB.OpenRecordset("Verkauf", dbOpenTable)
'    VerkaufRec.Index = "Unique"
    
'    SQLStr = "SELECT * FROM Verkauf WHERE Datum='" + Format(iiDat, "DD.MM.YYYY HH:MM:SS") + "'"
'    SQLStr = SQLStr + " AND LaufNr=" + CStr(iiLaufNr)
    SQLStr = "SELECT * FROM Verkauf WHERE PrivRezLaufNr='" + CStr(VerkaufAdoRec!PrivRezLaufNr) + "'"
    SQLStr = SQLStr + " ORDER BY Datum DESC,LaufNr DESC,ZeilenNr"
    FabsErrf = VerkaufAdoDB.OpenRecordset(VerkaufAdoRec, SQLStr, 0)
    If Not (VerkaufAdoRec.EOF) Then
        bmVerkPtr = VerkaufAdoRec.Bookmark
    End If
    
    bmOrgVerkPtr = bmVerkPtr
    
    bmPrivEnde = bmVerkPtr
    lRezNr& = -1&
    PrivatRezept% = True
    GebFrei% = True
Else

'    Set VerkaufRec = VerkaufDB.OpenRecordset("Verkauf", dbOpenTable)
'    VerkaufRec.Index = "RezeptNr"
    
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
'        If (lRezNr& <= 0) Or (lRezNr& > 32000) Then
        If (lRezNr& <= 0) Then
            RezeptHolenDB% = False
            Call DefErrPop: Exit Function
        End If
        hStr$ = "9999998" + Format(lRezNr&, "00000")
        Call PruefZiffer(hStr$)
        RezNr$ = hStr$
    End If

'    lRezNr& = CLng(RezNr$)

'    ParentForm.txtRezeptNr.text = RezNr$
    
    Call WinArtDebug("Suche ELSE")

    bmVerkPtr = Null
    
'    VerkaufRec.Seek ">=", RezNr$    ', Now
'    If (VerkaufRec.NoMatch = False) Then
'        If (VerkaufRec!RezeptNr = RezNr$) Then
'            bmVerkPtr = VerkaufRec.Bookmark
'        End If
'    End If
    SQLStr = "SELECT * FROM Verkauf WHERE RezeptNr='" + RezNr + "'"
    SQLStr = SQLStr + " ORDER BY Datum DESC,LaufNr DESC,ZeilenNr"
    FabsErrf = VerkaufAdoDB.OpenRecordset(VerkaufAdoRec, SQLStr, 0)
    If Not (VerkaufAdoRec.EOF) Then
        bmVerkPtr = VerkaufAdoRec.Bookmark
    End If

    Call WinArtDebug("nach Suche")
    
    If (RezeptHuellen%) Then
        SchonDa% = False
    ElseIf (RezSpeicherOK%) Then
        SchonDa% = False
    End If
    
    If (SchonDa% = False) And (IsNull(bmVerkPtr)) Then
        RezeptHolenDB% = False
        Call DefErrPop: Exit Function
    End If

End If

Call WinArtDebug("vor Auslesen")
    

If (SchonDa%) Then
'    Call HoleAusRezeptSpeicher
Else
    Call InitRezept(2)
    AMverfügbarkeit = String(10, "1")

    LetztArtikel$ = ""
    k% = -1
    
    Call WinArtDebug("vor Bookmark")
    
    VerkaufAdoRec.Bookmark = bmVerkPtr
    If (PrivatRezept%) Then
        AktPrivRezNr = VerkaufAdoRec!PrivRezNr
    End If
    If (VerkaufAdoRec!Bediener > 0) Then
        PersonalNr% = VerkaufAdoRec!Bediener
    End If
                
    lIkNr = CheckNullLong(VerkaufAdoRec!IkNr)
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
    
    IvfRezept% = (VerkaufAdoRec!RezeptArt = 7)
    
    Call WinArtDebug("vor DO")
    
    Do
        sPzn = PznString(VerkaufAdoRec!pzn)
        
        Call WinArtDebug("Artikel: " + sPzn)
    
        sText = VerkaufAdoRec!text
        dPreis = VerkaufAdoRec!Preis
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
        
        If (Mid$(sText, 15, 4) = "PZN:") Then
            'Stückelung
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
'                If ((PrivatRezept%) And ((VerkaufAdoRec!PrivRezNr <> AktPrivRezNr) Or (Val(VerkaufAdoRec!RezeptNr) > 0))) Then
'                    k% = k% + 1
'                Else
'                    RezArtikel(k% + 1).zusatz(0) = sText
'                End If
'                PrivatBesorger% = 0
'            ElseIf (PrivatRezept%) And (Left$(sText, 1) = "(") Then
'                RezArtikel(k%).zusatz(1) = sText
'                PrivatBesorger% = 0
            ElseIf (AktArtikel$ = LetztArtikel$) And (Val(sPzn) <> 9999999) And _
                ((PrivatRezept% = False) Or ((VerkaufAdoRec!PrivRezNr = AktPrivRezNr) And (Val(VerkaufAdoRec!RezeptNr) = 0))) Then
                RezArtikel(k%).Faktor = RezArtikel(k%).Faktor + VerkaufAdoRec!Anzahl
                
                If (VerkaufAdoRec!HimiVerbrauch) Then
'                    RezArtikel(k%).Zuz = RezArtikel(k%).Zuz + VerkaufAdoRec!Zuzahlung
                End If
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
                RezArtikel(k%).text = Left$(sText, 28) + " " + CheckNullStr(VerkaufAdoRec!menge) + CheckNullStr(VerkaufAdoRec!einheit)
                RezArtikel(k%).Preis = dPreis
                RezArtikel(k%).Faktor = VerkaufAdoRec!Anzahl
                RezArtikel(k%).flag = 0
                RezArtikel(k%).Appli = Abs(InStr(UCase(VerkaufAdoRec!Rezeptzeichen), "A") > 0)  'Abs(VerkaufAdoRec!Hilfsmittel)
                
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
                            RezArtikel(k%).Appli = 1
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
                
                If (CheckNullLong(VerkaufAdoRec!Verfügbarkeit) > 1) Then
                    Mid(AMverfügbarkeit, k + 1, 1) = CStr(VerkaufAdoRec!Verfügbarkeit)
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
                    RezArtikel(k%).Warenzeichen = 99
                    SollTmName(k%) = UCase(sText)
                    SollTmMenge(k%) = xVal(CheckNullStr(VerkaufAdoRec!menge))
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
                    Call WinArtDebug("Artikel: in Taxe")
                    
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
                        If (CheckNullDouble(VerkaufAdoRec!ZuzaGes) > 0) Then
                            RezArtikel(k%).Zuz = VerkaufAdoRec!ZuzaGes / RezArtikel(k%).Faktor
                        Else
                            RezArtikel(k%).Zuz = VerkaufAdoRec!Zuzahlung
                        End If
                    End If
                    
                    If (VerkaufAdoRec!MehrkostenVerzichtKZ) Then
                        RezArtikel(k%).PlusMehrKosten = 1
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

'          DruckDatum = 0L;

            VerkZeit% = Val(Format(VerkaufAdoRec!Datum, "HHMM"))
            
            ActVerordnung% = 1
            If (VerkaufAdoRec!RezeptArt = REZEPTART_SPRECHSTUNDENBEDARF) Then ActVerordnung% = 2

            iUser% = VerkaufAdoRec!Computer
            VerkUser$ = Format(iUser%, " 0")
            VerkLiRe = VerkaufAdoRec!Seite
    
            iKuNr = VerkaufAdoRec!KundenNr
            If (iKuNr > 0) Then
                Call WinArtDebug("Artikel: vor HoleKundenInfo")
                Call HoleKundenInfo(iKuNr)
                Call WinArtDebug("Artikel: nach HoleKundenInfo")
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
            
'            If (AutIdemSonderregelOk%) And (Val(ActKkasse$) > 0) Then
'                SQLStr$ = "SELECT * FROM kKassen WHERE IK = " + Str$(Val(ActKkasse$))
'                Set KkassenRec = kKassenDB.OpenRecordset(SQLStr$)
'                If (KkassenRec.RecordCount > 0) Then
'                    AutIdemIk& = CheckNullLong(KkassenRec!IK)
'                    AutIdemKbvNr& = CheckNullLong(KkassenRec!KbvNr)
'                End If
'            End If


'            If ((PrivatRezept% = False) Or ((VerkaufAdoRec!PrivRezNr = AktPrivRezNr) And (Val(VerkaufAdoRec!RezeptNr) = 0))) Then
'                k% = k% + 1
'            Else
'                lPrivEnde& = lPrivEnde& - 1
'            End If
            
            bmPrivEnde = VerkaufAdoRec.Bookmark
            k% = k% + 1
            
            Exit Do
        End If
        
        VerkaufAdoRec.MoveNext
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


    If Not (VerkaufAdoRec!KundenEnde) And Not (PrivatRezept%) And (SonderBelegRezept% = 0) Then
        For i% = 0 To UBound(BarVerkauf$)
            VerkaufAdoRec.MoveNext
            If (VerkaufAdoRec.EOF) Then
                Exit For
            End If

            If (VerkaufAdoRec!RezeptArt = 0) Then
                hStr$ = VerkaufAdoRec!text + " " + Format(VerkaufAdoRec!Preis, "0.00")
                BarVerkauf$(i%) = hStr$

'                With ParentForm
'                    .mnuBearbeitenInd(MENU_SF3).Enabled = True
'                    .cmdToolbar(MENU_SF3).Enabled = True
'                End With
            End If

            If (VerkaufAdoRec!KundenEnde) Then
                Exit For
            End If
        Next i%
    End If

'    If (PrivatRezept%) Or (SonderBelegRezept%) Then
'    Else
'        AMverfügbarkeit = Left$(String(AnzRezeptArtikel%, "1") + "000", 3)
'        For i = 1 To AnzRezeptArtikel%
'            If (Val(Mid(AMverfügbarkeit2, i, 1)) > 1) Then
'                Mid(AMverfügbarkeit, i, 1) = Mid(AMverfügbarkeit2, i, 1)
'            End If
'        Next i
'    End If

    Call WinArtDebug("vor 2.Schleife")
    
    For i% = 0 To (AnzRezeptArtikel% - 1)
        dZuz# = 0#
        dAvp# = 0
        
        Call WinArtDebug("in 2.Schleife")
    
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
                ElseIf (TaxeRec!BTMKz) Then
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
                    If (RezArtikel(i%).Appli = 0) Then
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
                    Call WinArtDebug("in 2.Schleife, Taxe")
    
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
                    
                    If (TaxeRec!NegListe) Then dZuzDos# = 1#
                    If (RezArtikel(i%).Appli) Then dZuzDos# = 0#
                    
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
                        dZuz# = RezArtikel(i%).Zuz
                    End If
                
                    If (TaxeRec!BTMKz) Then
                        BtmRezept% = True
                    End If
                    
                    hStr$ = Format$(TaxeRec!Abgabebest, "0")
                    If (InStr("01245", hStr$) <> 0) Then
                        RezArtikel(i%).Fam = 1
                    End If
                
                    Call WinArtDebug("in 2.Schleife, vor CheckAutIdem")
                    Call CheckAutIdem(i%)
                    Call WinArtDebug("in 2.Schleife, nach CheckAutIdem")
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
                        Call WinArtDebug("in 2.Schleife, vor Original")
                        RezArtikel(i%).Imp123 = 6
                        OrgAvp = TaxeRec!vk / 100#
                        SQLStr$ = "SELECT * FROM TAXE WHERE ORIGINALPZN = " + Format(TaxeRec!OriginalPZN, "0000000")
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
                        Call WinArtDebug("in 2.Schleife, nach Original")
                        SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + RezArtikel(i%).pzn
                        'Set TaxeRec = TaxeDB.OpenRecordset(SQLStr$)
                        On Error Resume Next
                        TaxeRec.Close
                        Err.Clear
                        On Error GoTo DefErr
                        TaxeRec.Open SQLStr, taxeAdoDB.ActiveConn
                    
                    ElseIf (TaxeRec!OriginalPZN > 0) Then
                        aktavp = TaxeRec!vk / 100
                        SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + Format(TaxeRec!OriginalPZN, "0000000")
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
'        For i% = 0 To UBound(RezAddGesamtBrutto$)
'            RezAddGesamtBrutto$(i%) = ""
'        Next i%
        Call PruefeBtmKosten
'        sBrutto$ = RezAddGesamtBrutto$(0)
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
        Call WinArtDebug("Verbandmittel")
        For i% = 0 To (AnzRezeptArtikel% - 1)
            If (RezArtikel(i%).flag And REZ_WG_4) Then
                RezeptHolenZuz# = RezArtikel(i%).Zuz
                AktTabIndex% = 3 + i% * 4
                Call Verbandmittel
            End If
        Next i%
    End If
    RezeptHolenZuz# = 0
    
    Call CheckVerordnungMDB
End If

'ActVerordnung% = 1

For i% = 0 To (AnzRezeptArtikel% - 1)
    If (Mid(AMverfügbarkeit, i + 1, 1) = "7") Then
        With RezArtikel(i)
            RezArtikel(i).Preis = 0#
            RezArtikel(i).Zuz = 0
        End With
    End If
Next
Call CalcRezeptWerte

Call WinArtDebug("Ende RezeptHolenDB")

RezeptHolenDB% = True

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
Dim SQLStr$, SQLStr2$, SqlSonderregelNr$, h$
Dim iRec As New ADODB.Recordset
'Dim tRec2 As Recordset
Dim tRec2 As New ADODB.Recordset
Dim RabattAutidem As Boolean
        
RezArtikel(pos%).N0 = Abs(TaxeRec!NGrösse = "N0")

'If (AutIdemOk%) Then
'    If (TaxeRec!AutIdemPflicht) Then
'        RezArtikel(pos%).AutIdem = 2
'    ElseIf (TaxeRec!AutIdemNeu) Then
'        RezArtikel(pos%).AutIdem = 1
'    ElseIf (TaxeRec!IdentGruppe > 0&) Then
'        If (GibtsOpAutIdem%) Then RezArtikel(pos%).AutIdem = 1
'    End If
'ElseIf (TaxeRec!IdentGruppe > 0&) Then
'    IstDrüber% = False
'    If (TaxeRec!FestKz) And (TaxeRec!FESTBETRAG < TaxeRec!vk) Then IstDrüber% = True
'
'    If (GibtsOpAutIdem%) Then
'        RezArtikel(pos%).AutIdem = 1
'        If (IstDrüber% And (TaxeRec!vk <= TaxeRec!FESTBETRAG)) Then
'            RezArtikel(pos%).AutIdem = 2
'        End If
'    End If
'End If
       
If (AutIdemOk%) Then
    If (TaxeRec!AutIdemPflicht) Or (TaxeRec!AutIdemNeu) Then
        AusnahmeErsetzungFl = 0
        On Error Resume Next
        AusnahmeErsetzungFl = (CheckNullLong(TaxeRec!AusnahmeErsetzungFl) = 1)
        On Error GoTo DefErr
        If (AusnahmeErsetzungFl) Then
            Beep
        Else
            RezArtikel(pos%).AutIdem = 1
            If (TaxeRec!AutIdemPflicht) Then
                RezArtikel(pos%).AutIdem = 2
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
                    SQLStr$ = "SELECT * FROM PZNSUBST WHERE PZN = " + Str$(TaxeRec!pzn)
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
    Else
        If (TaxeRec!IdentGruppe > 0&) Then
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

'    If (PrivatRezept% = 0) And (kKassenOk%) And (AutIdemKbvNr > 0) And ((RezArtikel(pos%).flag And REZ_WG_4) = 0) Then
'        RabattAutidem = CheckRabattAutIdem(TaxeRec!pzn)
'        If Not RabattAutidem And Not IstRabattAutIdem(TaxeRec!pzn) Then
'            If (CheckRabattAutIdem(TaxeRec!pzn, 1)) Then
'                Call iMsgBox("Rabattartikel mit abweichender DAR oder N-Größe vorhanden", vbInformation, Trim(RezArtikel(pos).text))
'            End If
'        End If
'    End If
ElseIf (TaxeRec!IdentGruppe > 0&) Then
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
Dim ret%
Dim SQLStr$

ret% = False

SQLStr$ = "SELECT * FROM TAXE WHERE IDENTGRUPPE = " + Str$(TaxeRec!IdentGruppe)
SQLStr$ = SQLStr$ + " and Menge = '" + TaxeRec!menge + "' and Einheit = '" + TaxeRec!einheit + "'"
SQLStr$ = SQLStr$ + " and Festbetrag = " + Str$(TaxeRec!FESTBETRAG)
SQLStr$ = SQLStr$ + " ORDER BY VK"
'Set TaxeRec = TaxeDB.OpenRecordset(SQLStr$)
On Error Resume Next
TaxeRec.Close
Err.Clear
On Error GoTo DefErr
TaxeRec.Open SQLStr, taxeAdoDB.ActiveConn
If Not (TaxeRec.EOF) Then
    TaxeRec.MoveLast
    If (TaxeRec.RecordCount > 1) Then ret% = True
    TaxeRec.MoveFirst
End If

GibtsOpAutIdem% = ret%

Call DefErrPop
End Function

Sub PruefZiffer(s$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("PruefZiffer")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim i%, sum%, zw%

sum% = 0
For i% = 1 To 12
    zw% = Val(Mid$(s$, i%, 1))
    sum% = sum% + (zw% * (1 + 2 * ((i% + 1) Mod 2)))
Next i%

zw% = 10 - (sum% Mod 10)
If (zw% = 10) Then zw% = 0

s$ = Left$(s$, 12) + Format(zw%, "0")

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

If (InitModus% And &H1) Then
    GebFrei% = False
    ImpfRez% = False
    ActBundesland% = OrgBundesland%
    ActKasse% = -1
    ActKkasse$ = ""
    PrivatRezept% = False
    BtmRezept% = False
    IvfRezept% = False
    SonderBelegRezept% = False
    SchonDa% = False
    
    BereitsGedruckt% = False
    
    ActKiste% = 0
    
    On Error Resume Next
    Kill "MagTax" + Trim(para.User)
    On Error GoTo DefErr
    
    RezNr$ = ""
    lRezNr& = 0
    DruckDatum$ = Space$(6)
    
    AvpRezeptNr$ = ""
    AutIdemKbvNr& = 0
    AutIdemIk& = 0
    
'    RezeptMitHilfsmittelNr% = -1
    
    TransaktionsID& = 0
    ParenteralRezept = -1
'    ParenteralHash = ""
    rzLieferId = ""
    
    
    ActVebNr = 0
    ActPauschaleNr = 0
    
    RezeptMitHilfsmittelNr% = -1
    
    WVORabattArtikel = String(10, "1")
    
    ZusatzKomponentenNr = ""
    ZusatzKomponentenBasisNr = ""
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
    
    For i% = 0 To 4
        RezKundenInfo$(i%) = ""
    Next i%
    
    AnzRezeptArtikel% = 0
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
  
Call WinArtDebug("RezKontrInit")

'MsgBox ("init")

    RezSpeicherOK% = OpenRezeptSpeicher%
    If (RezSpeicherOK% = False) Then
        Call iMsgBox("Rezeptspeicher momentan nicht verwendbar !")
    End If

Call WinArtDebug("nach OpenRezeptSpeicher")

    
'    ParentForm.Caption = "Rezeptkontrolle - Abrechnungsmonat " + Format(CDate("01." + Mid(AbrechMonat$, 3, 2) + "." + Left(AbrechMonat$, 2)), "mmmm") + " bis " + Mid(AbrechAbgabeDatum$, 5, 2) + "." + Mid(AbrechAbgabeDatum$, 3, 2) + "." + Left(AbrechAbgabeDatum$, 2)
    
    If (InStr(para.Benutz, "V") > 0) And (para.Land = "D") Then
        AllesFlag% = (InStr(para.Benutz, "R") > 0)
        ImpFlag% = (InStr(para.Benutz, "!") > 0)
        EasyImpFlag% = (InStr(para.Benutz, "i") > 0)
        If (EasyImpFlag%) Then ImpFlag% = True
        VmFlag% = (InStr(para.Benutz, "?") > 0)
        FiveRxFlag% = (InStr(para.Benutz, "Z") > 0) Or (InStr(para.Benutz, "z") > 0)
    Else
        RezKontrInit% = False: Call DefErrPop: Exit Function
    End If
    
    Call HolVerbandmittel
    
    Call LadeSonderPzn("sondrpzn.dat", SonderPzn$)
    
    Call LadeAbgabePreise
    
    h$ = "0"
    l& = GetPrivateProfileString("Allgemein", "RezeptHuellen", "0", h$, 2, "\user\dp.ini")
    h$ = Left$(h$, l&)
    If (Val(h$) <> 0) Then
        RezeptHuellen% = True
    Else
        RezeptHuellen% = False
    End If

    h$ = "1"
    l& = GetPrivateProfileString("AlleKassen", "BTMDerivatmengeIgnorieren", "1", h$, 2, "\user\vkpara.ini")
    h$ = Left$(h$, l&)
    If (Val(h$) <> 0) Then
        BtmDerivatMengeIgnorieren = True
    Else
        BtmDerivatMengeIgnorieren = False
    End If
    

    
Call WinArtDebug("vor param.dat")

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
    
Call WinArtDebug("nach param.dat")

    If (RezApoDruckName$ = "") Then
        RezApoDruckName$ = RezApoName$(0) + " " + RezApoName$(1)
        Call OemToChar(RezApoDruckName$, RezApoDruckName$)
    End If
    
    If (RezeptDrucker$ = "") And (DosDruckerInd% > 0) Then
        erg% = GetGeraetParameter(DosDruckerInd%, RezeptDrucker$, RezeptDruckerPara$)
        If (RezeptDrucker$ = "WINPRINT") Then RezeptDrucker$ = ""
        If (UCase(Left$(RezeptDruckerPara$, 3)) = "LPT") Then RezeptDrucker$ = ""
    End If
    If (RezeptDrucker$ = "") Then RezeptDrucker$ = Printer.DeviceName
    Call frmAction.ErzeugeDruckerAuswahl
 
    Call OpenArbEmb
    
Call WinArtDebug("RezKontrInit ENDE")

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

Call WinArtDebug("HolVerbandmittel")

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

Call WinArtDebug("HolVerbandmittel" + " ENDE")

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
  
Call WinArtDebug("LadeAvDatei: " + SrcDatei$)

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
    
Call WinArtDebug("LadeAvDatei: " + SrcDatei$ + " ENDE")

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
Dim h$, SQLStr$
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
    h$ = h$ + CheckNullStr(avRec!Name)
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
  
If (TaxeRec!MwstKz) Then
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

dZuz# = CalcTaxeZuzahlung

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
    
    If Not (GebFrei%) And (dZuz# > (RezArtikel(row%).Preis / VmRabattFaktor#)) Then
        RezArtikel(row%).Preis = 0
        RezArtikel(row%).Zuz = 0
    End If

    If (RezArtikel(row%).Appli) And (Val(Left$(TaxeRec!Zuzahlung, 1)) <> 9) Then
        RezArtikel(row%).Zuz = 0
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
Dim i%, row%, ind%, erg%, l%, iGenehmPflicht%, iRow%, ZweiteZeile%, Paragraph302%, AbrechnungsverfahrenKz%
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
  
If (TaxeRec!MwstKz) Then
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

''If (ActKasse% < 0) Then
''    frmAvAuswahl.Show 1
''    Call PaintAvParameter
''End If
'If (ActKasse% < 0) Then
'    Call DefErrPop: Exit Function
'End If

If (ActVebNr = 0) And (ActPauschaleNr = 0) Then
    Call DefErrPop: Exit Function
End If

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
    
    AvAbrechnungsNr(0) = Trim(CheckNullStr(VdbPauschaleRec!AbrechnungsNr))
        
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

    If (RezArtikel(row%).Appli) And (Val(Left$(TaxeRec!Zuzahlung, 1)) <> 9) Then
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
    If (VdbBedRec!GenehmigungspflichtFl) Then
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
    erg% = MachBerechnungMDB%(VdbBedRec!BedNr)
    If (erg% = 0) Then
        RezArtikel(row%).Zuz = 0
    End If
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
                AnzRezeptArtikel% = AnzRezeptArtikel% + 1
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
                End If
                RezArtikel(iRow%).Faktor = 1
                RezArtikel(iRow%).Preis = ActPreis(i)
                VmAddPreis = ActPreis(i)
                
                RezArtikel(iRow%).Zuz = 0
                
                VerbandmittelMDB% = True
                
                If (ActWarenzeichen$ = "J") Then
                    RezArtikel(iRow%).Warenzeichen = 1
                End If
                
'                Call PaintArtikel(iRow)
                
                If (ZweiteZeile) Then
                    Exit For
                End If
                ZweiteZeile = True
            End If
        Next i
    End If
    
'        ZweiteZeile = 0
'        For i = 0 To 2
'            If (AvHmPositionsNr(i) = "") Then
'                ActVmPreis = ActPreis(i)
'                h = RezArtikel(row%).pzn
'                RezArtikel(row%).pzn = "-" + Left(h, 7)
'                RezArtikel(row%).ScreenPzn = CheckNullStr(VdbPznRec!HmPositionsNr1)
'                If (ZweiteZeile) Then
'                    Exit For
'                End If
'                ZweiteZeile = True
''                Exit For
'            ElseIf (AvHmPositionsNr(i) = VdbPznRec!HmPositionsNr2) Then
'                AnzRezeptArtikel% = AnzRezeptArtikel% + 1
'                iRow% = AnzRezeptArtikel% - 1
'                Call InitRezeptArtikel(iRow%)
'                RezArtikel(iRow%).text = "Zusatzkomponente"
''                RezArtikel(iRow%).pzn = "-ZUSG_HM"
'                RezArtikel(iRow%).pzn = "-9999999"
'                RezArtikel(iRow%).ScreenPzn = AvHmPositionsNr(i)
'                RezArtikel(iRow%).Faktor = 1
'                RezArtikel(iRow%).Preis = ActPreis(i)
'                VmAddPreis = ActPreis(i)
'
''                dZuzPreis# = ActPreis(i)
''                If (ActVmRabatt% = 1) Or (ActVmRabatt% = 2) Then
''                    dZuzPreis# = dZuzPreis# * 0.94
''                ElseIf (ActVmRabatt% = 8) Then
''                    dZuzPreis# = dZuzPreis# * 0.95
''                End If
''                dZuz# = CalcZuzaNeu(dZuzPreis#, 0, 1)
''                RezArtikel(iRow%).Zuz = dZuz#
'                RezArtikel(iRow%).Zuz = 0
'
'                VerbandmittelMDB% = True
'
'                If (ActWarenzeichen$ = "J") Then
''???                    RezArtikel(iRow%).Warenzeichen = 1
'                End If
'
''                Call PaintArtikel(iRow)
'                If (ZweiteZeile) Then
'                    Exit For
'                End If
'                ZweiteZeile = True
'            End If
'        Next i
'    End If
    
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

    If (RezArtikel(row%).Appli) And (Val(Left$(TaxeRec!Zuzahlung, 1)) <> 9) Then
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

'Function SucheInGruppeMDB%(sGruppe$, Optional iBundesland% = 0)
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Call DefErrFnc("SucheInGruppeMDB%")
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
'Dim ErstFlag%, erg%, iWert%
'Dim ActRecNo&
'Dim ActBedingung$, h$, h2$, BedKey$, SQLStr$
'
''  int GefFlag, ErstFlag, erg, erg2;
''  long ActRecno;
''  char h[25], jnStr[2], ActBedingung[101];
'
'If (iBundesland = 0) Then
'    iBundesland = ActBundesland
'End If
'BedKey$ = sGruppe$ + Format(iBundesland, "00") + Format(ActKasse, "00") + Format(ActVerordnung, "00")
'
''If (VmDebugAktiv%) Then
''    frmAvDebug.flxAvDebug(1).AddItem "KEY" + vbTab + BedKey$
''End If
'
'SQLStr = "SELECT * FROM VdbBedingungen WHERE GruppeNr='" + sGruppe + "'"
'SQLStr = SQLStr + " AND KostentrNr=" + CStr(ActKasse)
'SQLStr = SQLStr + " AND LandNr=" + CStr(iBundesland)
'SQLStr = SQLStr + " AND VOrdTypNr=" + CStr(ActVerordnung)
''Set VdbBedRec = AplusVDB.OpenRecordset(SQLStr)
'FabsErrf = AplusVDB.OpenRecordset(VdbBedRec, SQLStr)
'Do
'    If (VdbBedRec.EOF) Then
'        SucheInGruppeMDB% = 0
'        Call DefErrPop: Exit Function
'    End If
'
'    ActBedingung$ = Trim(CheckNullStr(VdbBedRec!Bedingungsausdruck))
'
'    erg% = PruefeBedingung%(ActBedingung$)
'
'    If (erg%) Then
'        If (InStr("JN", VdbBedRec!Warenzeichen) > 0) Then
'            If (ActWarenzeichen$ = " ") Then
'                erg% = iMsgBox("Wurde das Produkt als Markenprodukt verordnet ?", vbYesNo Or vbInformation, "A+V Taxierung")
'                If (erg% = vbYes) Then
'                    ActWarenzeichen$ = "J"
'                Else
'                    ActWarenzeichen$ = "N"
'                End If
'            End If
'
'            If (VdbBedRec!Warenzeichen <> ActWarenzeichen$) Then
'                erg% = False
'            End If
'        End If
'    End If
'
'    If (erg%) Then Exit Do
'
'    VdbBedRec.MoveNext
'Loop
'
'SucheInGruppeMDB% = True
'
'Call DefErrPop
'End Function

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
Dim i%, ErstFlag%, erg%, iWert%
Dim ActRecNo&
Dim ActBedingung$, h$, h2$, BedKey$
  
'  int GefFlag, ErstFlag, erg, erg2;
'  long ActRecno;
'  char h[25], jnStr[2], ActBedingung[101];

'If (iBundesland = 0) Then
'    iBundesland = ActBundesland
'End If
'BedKey$ = sGruppe$ + Format(iBundesland, "00") + Format(ActKasse, "00") + Format(ActVerordnung, "00")
BedKey$ = sGruppe$ + Format(iVebNr, "000000") + Format(ActVerordnung, "00")

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

    erg% = PruefeBedingung%(ActBedingung$)

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
Dim AnzStr$, h$, SQLStr$

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
    h$ = InputBox("Dauer der Miete (in " + AnzStr$ + "): ", "A+V Taxierung", "1")
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
            SQLStr = "SELECT * FROM VdbBerechnungsformeln WHERE BedNr=" + CStr(VdbBedRec!BedNr)
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
    dWert# = 8.56 + (dAep# * 0.03)   'ab 31.7.2020
    
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
    
  Case "VP", "FB"
    If (ActWert# < 0) Then
        SQLStr = "SELECT * FROM VdbBerechnungsformeln WHERE BedNr=" + CStr(VdbBedRec!BedNr)
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
                      erg# = wert1 + (8.56 + (wert1 * 0.03))    'ab 31.7.2020
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
  
AnzGef% = 0

For i% = 1 To 1000

    ActKeyBed$ = KeyBed$ + Format(i%, "00000")
    FabsErrf% = VmRech.IndexSearch(0, ActKeyBed$, FabsRecno&)
    If (FabsErrf% > 0) Then Exit For
        
    Call VmRech.GetRecord(FabsRecno&)
    
    ActWert# = VmRech.RechWert / 100#
    
    ActBerechnung$ = Trim(VmRech.Berechnung)
    
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
    
    VkBerechnung% = False
    If (InStr(VmRech.Berechnung, "VK") > 0) Then VkBerechnung% = True
    If (VmRech.Mwst = "J") And (VkBerechnung% = False) Then
        ActPreis#(AnzGef%) = ActPreis#(AnzGef%) * (1# + (ActMwst) / 100#)
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
        
    If (VmRech.Rabatt = "1") Or (VmRech.Rabatt = "2") Then
        ActPreis#(AnzGef%) = ActPreis#(AnzGef%) * VmRabattFaktor
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
Dim i%, AnzGef%, VkBerechnung%, ind%, ind2%, iRabatt%, iWert%, p1%, p2%, AbrechnungsverfahrenKz%
Dim EinzelPfennig#, KlammerPreis#, dPreis#, diff#, erg#
Dim ActKeyBed$, ActBerechnung$, ActBer2$, h$, h2$, ch$, SQLStr$, c$, x$
  
'If (VmDebugAktiv%) Then
'    With frmAvDebug.flxAvDebug(2)
'        .AddItem "BERECH"  '18
'        .AddItem "WERT"    '19
'        .AddItem "EINZEL"    '20
'        .AddItem "MWST"    '21
'        .AddItem "RABATT"    '22
'        .AddItem "RUNDEN"    '23
'    End With
'End If


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
    
    
    
    VkBerechnung% = False
    If (InStr(CheckNullStr(VdbBerRec!Formelausdruck), "VK") > 0) Then VkBerechnung% = True
    If (Abs(VdbBerRec!MehrwertSteuerFl) > 0) And (VkBerechnung% = False) Then
        ActPreis#(AnzGef%) = ActPreis#(AnzGef%) * (1# + (ActMwst) / 100#)
    End If
    ActPreis#(AnzGef%) = FNX(ActPreis#(AnzGef%))
    
'    If (VmDebugAktiv%) Then
'        ind% = AnzGef% * 3 + 1
'        frmAvDebug.flxAvDebug(2).TextMatrix(4, ind%) = Format(ActPreis#(AnzGef%), "0.00")
'    End If
    
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
        
'        If (VmDebugAktiv%) Then
'            ind% = AnzGef% * 3 + 1
'            frmAvDebug.flxAvDebug(2).TextMatrix(5, ind%) = Format(ActPreis#(AnzGef%), "0.00")
'        End If
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
        
'        If (VmDebugAktiv%) Then
'            ind% = AnzGef% * 3 + 1
'            frmAvDebug.flxAvDebug(2).TextMatrix(5, ind%) = Format(ActPreis#(AnzGef%), "0.00")
'        End If
    End If
    ActPreis#(AnzGef%) = FNX(ActPreis#(AnzGef%))
    
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


If (VdbBedRec!RundungstypFl) Then
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
Else
'        ActVmPreis# = (ActVmPreis# * 100#) '+ 0.501
'        ActVmPreis# = CLng(ActVmPreis#)
'        ActVmPreis# = ActVmPreis# / 100#
    ActVmPreis# = CDbl(Int(Abs(ActVmPreis#) * 100# + 0.501) / 100#)
End If

'If (VmDebugAktiv%) Then
'    frmAvDebug.flxAvDebug(2).TextMatrix(6, 1) = Format(ActVmPreis#, "0.00")
'End If

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
'Dim i%, AnzGef%, VkBerechnung%, ind%, ind2%, iRabatt%, iWert%
'Dim EinzelPfennig#, KlammerPreis#, dPreis#, diff#
'Dim ActKeyBed$, ActBerechnung$, ActBer2$, h$, h2$, ch$, SQLStr$
'
''If (VmDebugAktiv%) Then
''    With frmAvDebug.flxAvDebug(2)
''        .AddItem "BERECH"  '18
''        .AddItem "WERT"    '19
''        .AddItem "EINZEL"    '20
''        .AddItem "MWST"    '21
''        .AddItem "RABATT"    '22
''        .AddItem "RUNDEN"    '23
''    End With
''End If
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
''    If (VmDebugAktiv%) Then
''        With frmAvDebug.flxAvDebug(2)
''            ind% = AnzGef% * 3 + 1
''            .TextMatrix(1, ind%) = ActBerechnung$
''
''            iWert% = Abs(VdbBerRec!MehrwertsteuerFl)
''            If (iWert% = 1) Then
''                h2$ = "J"
''            Else
''                h2$ = "N"
''            End If
''            .TextMatrix(1, ind% + 1) = h2$
''
''            .TextMatrix(1, ind% + 2) = CStr(VdbBerRec!Kassenrabatt)
''
''            If (ActWert# <> 0#) Then
''                .TextMatrix(2, ind%) = Format(ActWert#, "0.00")
''            End If
''        End With
''    End If
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
''    If (VmDebugAktiv%) Then
''        ind% = AnzGef% * 3 + 1
''        frmAvDebug.flxAvDebug(2).TextMatrix(3, ind%) = Format(ActPreis#(AnzGef%), "0.00")
''    End If
'
'    VkBerechnung% = False
'    If (InStr(CheckNullStr(VdbBerRec!Formelausdruck), "VK") > 0) Then VkBerechnung% = True
'    If (Abs(VdbBerRec!MehrwertSteuerFl) > 0) And (VkBerechnung% = False) Then
'        ActPreis#(AnzGef%) = ActPreis#(AnzGef%) * (1# + (ActMwst) / 100#)
'    End If
'
''    If (VmDebugAktiv%) Then
''        ind% = AnzGef% * 3 + 1
''        frmAvDebug.flxAvDebug(2).TextMatrix(4, ind%) = Format(ActPreis#(AnzGef%), "0.00")
''    End If
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
''        If (VmDebugAktiv%) Then
''            ind% = AnzGef% * 3 + 1
''            frmAvDebug.flxAvDebug(2).TextMatrix(5, ind%) = Format(ActPreis#(AnzGef%), "0.00")
''        End If
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
''        If (VmDebugAktiv%) Then
''            ind% = AnzGef% * 3 + 1
''            frmAvDebug.flxAvDebug(2).TextMatrix(5, ind%) = Format(ActPreis#(AnzGef%), "0.00")
''        End If
'    End If
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
''        ActVmPreis# = (ActVmPreis# * 100#) '+ 0.501
''        ActVmPreis# = CLng(ActVmPreis#)
''        ActVmPreis# = ActVmPreis# / 100#
'    ActVmPreis# = CDbl(Int(Abs(ActVmPreis#) * 100# + 0.501) / 100#)
'End If
'
''If (VmDebugAktiv%) Then
''    frmAvDebug.flxAvDebug(2).TextMatrix(6, 1) = Format(ActVmPreis#, "0.00")
''End If
'
'MachBerechnungMDB% = True
'
'Call DefErrPop
'End Function

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
  
Call WinArtDebug("LadeSonderPzn: " + SrcDatei$)

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

Call WinArtDebug("LadeSonderPzn: " + SrcDatei$ + " ENDE")

Call DefErrPop
End Sub

Sub LadeAbgabePreise()
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
  
Call WinArtDebug("LadeAbgabePreise")

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

Call WinArtDebug("LadeAbgabePreise: " + " ENDE")

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
    .Fam = 0
    .AutIdem = 0
    .MagIndex = 0
    .Zuz = 0
    .Appli = 0
    .Faktor = 0
    .zusatz(0) = Space$(Len(.zusatz(0)))
    .zusatz(1) = Space$(Len(.zusatz(1)))
    .HerstRabattPrivat130Brutto = 0
    .VdbPauschale = 0
    .TkkPznDruck = 0
    .VebNr = 0
    .PauschaleNr = 0
    .Fortsetzung = 0
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
Dim i%, j%, k%, hFaktor%, hFaktor2%, ZusatzKompontenFaktor%, ind%, gef%, Noctu%, BtmApoDa%, MaxGruppenLen%, MaxGruppenInd%
Dim iMilli&, iSekunden&, iMinuten&, iStunden&, iTage&, iMonate&
Dim lPreis&, SollPreis&, IstPreis&, pz&, Pruefsumme&
Dim dStückPreis#
Dim ch$, h$, h2$, PznStr$, PznStr2$, hPosNr$, SQLStr$
Dim AvpRec As Recordset
      
If (AnzRezepte% > 1) Then
    Call RezeptPause(RezeptDruckPause%)
End If

RezDruckInd% = 0
'MagSpeicherIndex% = -1
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

    Call SpeicherDruckZeile("X" + AvpRezeptNr$ + "X")
ElseIf (FiveRxFlag) Or (Code128Flag) Then
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

    h$ = Right$(String(8, "0") + CStr(TransaktionsID), 8)
    Pruefsumme = 0
    For i = 1 To 8
      Pruefsumme = Pruefsumme + Val(Mid$(h, i, 1)) * (1 + 2 * ((i + 1) Mod 2))
    Next i
    h = h + Right$(CStr(Pruefsumme Mod 10), 1)
    
    Call SpeicherDruckZeile(h$)
    
'    TA1_V37 = (DateDiff("d", "01.07.2022", HashErstellDat) >= 0)
    TA1_V37 = True
        
'    ParenteralPara = "30" + RezApoNr + h + Format(Now, "YYYYMMDD:000000") + ":000"
End If

BtmApoDa = 0
If (AnzBlink% > 0) Then
    For j% = 0 To 2
        Call SpeicherDruckZeile(" ")
    Next j%
    
    For j% = 0 To 6
        For i% = 0 To 2
            Call SpeicherDruckZeile(" ")
        Next i%
    Next j%
    
    h$ = "(" + RezNr$ + ")"
    Call SpeicherDruckZeile(h$)
    
    For j% = 0 To 1
        Call SpeicherDruckZeile(" ")
    Next j%
Else
    Call SpeicherDruckZeile("+" + RezApoNr$ + "+")
    
    If (PrivatRezept%) Or (SonderBelegRezept%) Then
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
    
    Dim AnzMKs%, MKind%
    AnzMKs = 0
    For j% = 0 To (AnzRezeptArtikel% - 1)
        If (InStr(RezArtikel(j).zusatz(1), "MK ") > 0) Then
            AnzMKs = AnzMKs + 1
        End If
    Next
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
    
    For j% = 0 To (AnzRezeptArtikel% - 1)
'        If (BtmRezept%) And (j% >= 3) Then Exit For
        If (BtmRezept%) Then
            If (j = 3) And (Val(RezArtikel(0).pzn) = SONDERPZN_WVO) Then
                BtmApoDa = True
                Call SpeicherDruckZeile("@BTM2@" + BtmRezDruckName$)
            Else
                If (j% >= 3) Then Exit For
            End If
        End If
        
        If (j = 3) And (BtmRezept% = 0) And (Trim(RezNr) <> "") And (RezeptNrPositionAlt = 0) Then
            h2 = RezNr
            If (Left$(h2$, 7) = "9999998") Then h2$ = Mid$(h2$, 8, 5)
            Call SpeicherDruckZeile("@REZNR@" + h2$)
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
'                                    dStückPreis = RezArtikel(j%).Preis * hFaktor%
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
    Next j%
    j = j + AnzMKs
        
    
    If (BtmRezept%) Then
        Do
            If (j% >= 3) Then Exit Do
            
            For i% = 0 To 2
                Call SpeicherDruckZeile("   ")
            Next i%
            
            j% = j% + 1
        Loop
'        Call SpeicherDruckZeile("@BTM@" + BtmRezDruckName$)
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
        
        If (j = 3) And (BtmRezept% = 0) And (Trim(RezNr) <> "") And (RezeptNrPositionAlt = 0) Then
            h2 = RezNr
            If (Left$(h2$, 7) = "9999998") Then h2$ = Mid$(h2$, 8, 5)
            Call SpeicherDruckZeile("@REZNR@" + h2$)
        End If
    
        For i% = 0 To 2
            Call SpeicherDruckZeile("   ")
        Next i%
        
        j% = j% + 1
    Loop
    
    If ((AnzRezeptArtikel% >= 7) Or (lRezNr& = 0&) Or PrivatRezept% Or SonderBelegRezept%) Then
        h$ = "   "
    Else
        h$ = ""
        If (BtmRezept%) Or (RezeptNrPositionAlt) Then
            h$ = "(" + RezNr$ + ")"
        End If
        If (PersonalNr% > 0) Then
            h$ = h$ + Format(PersonalNr%, "00")
        End If
    End If
    Call SpeicherDruckZeile(h$)
    
'    If (BtmRezept%) Then
'        h$ = ""
'    Else
'        h$ = VerkDatum$
'        If (Noctu) Then
'            h$ = h$ + " " + Format(VerkZeit, "00:00")
'        End If
'    End If
'    Call SpeicherDruckZeile(h$)
'
'    If (BtmRezept%) Then
'        h$ = VerkDatum$
'        h$ = Right$(Space$(38) + h$, 38)
'    Else
'        h$ = RezApoDruckName$
'        Call CharToOem(h$, h$)
'        h$ = Left$(h$ + Space$(38), 38)
'    End If
'    Call SpeicherDruckZeile(Left$(h$ + Space$(38), 38))
    
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
        
        h$ = RezApoDruckName$
        Call CharToOem(h$, h$)
        h$ = Left$(h$ + Space$(38), 38)
        Call SpeicherDruckZeile(h$)
    End If
    
    h$ = ""
    For i% = 0 To RezDruckInd%
        h$ = h$ + RezDruckZeilen$(i%) + vbCrLf
    Next i%
'    MsgBox (h)
End If

Call DefErrPop
End Sub

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
Dim i%, j%, k%, erg%, hFaktor%, VkPrivCheck%
Dim RezGesamt#, RezGesamtFAM#, RezImpFähig#, RezImpIst#
Dim Quote#, ImpSoll#, diff#, SaldoV#, RezeptRabatt#(1), ArtikelRabatt#(1), ePreis#, nPreis#
Dim h$, pzn$, AuswertungMonat$, Unique$, SortDatum$, SQLStr$
Dim hImp123 As Byte
Dim aktvk As Double, OrgVk As Double
Dim SuchPzn As String
Dim RezImpErspart As Double

Dim Jahr As Boolean, AbrechVorMon As Boolean
Dim bmk As Variant
Dim w#(4)
Dim FieldNr%(8)
 
    
If (BereitsGedruckt%) Then Call DefErrPop: Exit Sub
If (RezSpeicherOK%) Then
    BereitsGedruckt% = True
    
    Set RezepteRec = RezSpeicherDB.OpenRecordset("Rezepte", dbOpenTable)
    Set ArtikelRec = RezSpeicherDB.OpenRecordset("Artikel", dbOpenTable)
    Set AuswertungRec = RezSpeicherDB.OpenRecordset("Auswertung", dbOpenTable)

'    If (SchonDa%) Then
    SchonDa% = False
    If (Len(Trim(RezNr)) > 2) Then
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
'        MsgBox (RezNr)
        
        Unique$ = RezepteRec!Unique
    
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
        
'        ArtikelRec!Warenzeichen = 0
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
        ArtikelRec!Appli = RezArtikel(j).Appli
        
        ArtikelRec!zusatz1 = RezArtikel(j).zusatz(0)
        ArtikelRec!zusatz2 = RezArtikel(j).zusatz(1)
        
        ArtikelRec!NichtInTaxe = RezArtikel(j).NichtInTaxe
        ArtikelRec!IstWg4 = RezArtikel(j).IstWg4
        ArtikelRec!PlusMehrKosten = RezArtikel(j).PlusMehrKosten
        ArtikelRec!ZuzahlungsErlass = RezArtikel(j).ZuzahlungsErlass
        ArtikelRec!MietDauer = RezArtikel(j).MietDauer
        ArtikelRec!VdbPauschale = RezArtikel(j).VdbPauschale
        
        ArtikelRec!VebNr = RezArtikel(j).VebNr
        ArtikelRec!PauschaleNr = RezArtikel(j).PauschaleNr
        ArtikelRec!Fortsetzung = RezArtikel(j).Fortsetzung
        
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
'    If (SonderBelegRezept% > 0) Then
        RezepteRec!knr = Val(RezKundenInfo$(4))
'    End If
    
    If (AvpTeilnahme%) Then
        RezepteRec!AvpRezeptNr = AvpRezeptNr$
        RezepteRec!AvpLaufNr = AvpLaufNr&
    End If
    If (FiveRxFlag) Then
        RezepteRec!TransaktionsID = TransaktionsID
'        If (ParenteralRezept >= 0) Then
'            RezepteRec!ParenteralHash = ParenteralHash
'        End If
    End If
    If (Code128Flag) Then
        RezepteRec!TransaktionsID = TransaktionsID
    End If
    
'    If (AutIdemIk& > 0) Then
'        RezepteRec!kKassenIk = AutIdemIk&
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
    
    RezepteRec.Update
    
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
        
        SaldoV# = dCheckNull(AuswertungRec!saldo)
        
        .Update
    
        If (SaldoV = 0#) Then
            If Val(Mid(AuswertungMonat$, 3, 2)) = 1 Then
                .Seek "<=", ActKkasse$, Format(Val(Left(AuswertungMonat$, 2)) - 1, "00") + "12"
            Else
                .Seek "<=", ActKkasse$, Left(AuswertungMonat$, 2) + Format(Val(Mid(AuswertungMonat$, 3, 2)) - 1, "00")
            End If
            If Not .NoMatch Then
                If (AuswertungRec!Kkasse = ActKkasse$) Then
                    SaldoV# = dCheckNull(AuswertungRec!saldo)
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
    
    RezSpeicherDB.Close
    Set RezSpeicherDB = OpenDatabase(REZ_SPEICHER, False, False)
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
    ret% = OpenDruckerCom%(frmAction, RezeptDruckerPara$)
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
    frmAction.comSenden.PortOpen = False
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
    ElseIf (ActKasse% = 1) And (ActBundesland% = 3) Then
        If (Trim(PznStr$) <> "") Then
            FabsErrf% = VmPzn.IndexSearch(0, PznStr$, FabsRecno&)
            If (FabsErrf% = 0) Then
                Call VmPzn.GetRecord(FabsRecno&)
        
                hPosNr$ = Trim(Left$(VmPzn.HmPositionsNr1, 10))
                If (hPosNr$ <> "") Then
                    If (Trim(Left$(VmPzn.HmPositionsNr2, 10)) = "") Then
                        PznStr$ = hPosNr$
                        IstScreenPzn% = True
                    End If
                
                    gef% = 0
                    h$ = Trim(VmPzn.menge)
                    If (h$ <> "") Then
                        If (Val(h$) > 0) Then
                            hFaktor2% = Val(h$)
                            gef% = True
                        End If
                    End If
                    
                    If (gef = 0) Then
                        SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + RezArtikel(j%).pzn
                        'Set TaxeRec = TaxeDB.OpenRecordset(SQLStr$)
                        On Error Resume Next
                        TaxeRec.Close
                        Err.Clear
                        On Error GoTo DefErr
                        TaxeRec.Open SQLStr, taxeAdoDB.ActiveConn
                        If (TaxeRec.EOF = False) Then
                            hFaktor2% = Val(TaxeRec!menge)
                            gef% = True
                        End If
                    End If
                    
                    If (gef%) Then
                        hFaktor2% = hFaktor2% * hFaktor%
                    End If
                End If
            End If
        End If
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
    frmAction.comSenden.PortOpen = False
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
Dim TmOffen%, AnzTm%
Dim KistePreis#
Dim h$, SQLStr$, SollRezNr$
Dim iRec As Recordset

If (Dir("rezkontr.dat") = "") Then Call DefErrPop: Exit Sub

REZK% = FileOpen("rezkontr.dat", "R", "B")

k% = -1
For i% = 0 To (AnzRezeptArtikel% - 1)
    h$ = RezArtikel(i%).zusatz(0)
    If (Left$(h$, 6) = "Abhol-") Then
        If (InStr(h$, "Besorger") = 0) Then
            ActKiste% = Val(Mid$(h$, 9, 4))

            If (k% = -1) Then
                Seek #REZK%, CLng(ActKiste% - 1) * Len(RezKontrSatz) + 1
                Get #REZK%, , RezKontrSatz

                RezeptGef% = False
                For k% = 0 To 9
                    h$ = Trim(RezKontrSatz.RezEan(k).Ean)
                    If (h$ = RezNr$) Then
                        RezeptGef% = True
                        Exit For
                    End If
                Next k%
            End If

            If (RezeptGef%) And (RezKontrSatz.Fertig = "*") Then
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
                        .Imp123 = 0
                        .ImpErspart = 0
                        .Fam = 0
                    End With

                End If

                RezKontrSatz.Preis(k%) = 0#
                RezKontrSatz.RezEan(k%).Ean = Space$(13)
                k% = k% + 1
            End If

        End If
    End If
Next i%

Close (REZK%)


TmOffen = 0
'If (TaxmusterDBok) Then
    For i% = 0 To (AnzRezeptArtikel% - 1)
        If (RezArtikel(i%).Warenzeichen = 99) Then
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
'                    MsgBox ("mFaktor: " + Str(TmMengenFaktor))
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
'End If



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
Dim SQLStr$
   
If (ActKiste% > 0) Then
    REZK% = FileOpen("rezkontr.dat", "RW", "B")
    Seek #REZK%, CLng(ActKiste% - 1) * Len(RezKontrSatz) + 1
    Put #REZK%, , RezKontrSatz
    Close (REZK%)
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
                    dPreis# = Val(InputBox("Kosten für '" + .Name + "' eingeben:", "Abgabe-Preise", "0"))
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
            If ind% >= 2 Then Exit For
        End If
    End With
Next i%

ind% = 0
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
    
End If
        
Call CalcRezeptWerte

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
                If (.pzn = "8150006") Or (.pzn = "2567001") Or (.pzn = "2567018") Then ' or (.pzn = "8150012")
                    ShowBox% = False
                End If
            End If
            
            If (ShowBox%) Then
                Call iMsgBox("Artikel nicht in Taxe. Erstattung fraglich.", vbInformation, .pzn + "  " + .text)
            End If
        End If
    End With
Next i%
    
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
Dim i%, erg%, AbInd%
Dim l&
Dim h$, sVerfügbarkeit$
        
AbInd = 0
If (Val(RezArtikel(0).pzn) = 9999643) Then
    AbInd = 1
End If

If (Left(WVORabattArtikel, 3) <> "111") Then
    For i% = UBound(RezArtikel) - 1 To AbInd Step -1
        RezArtikel(i% + 1) = RezArtikel(i%)
    Next i%
    Call InitRezeptArtikel(AbInd)
'    Call InitKontrollStop(0)
    AnzRezeptArtikel% = AnzRezeptArtikel% + 1
    RezArtikel(AbInd).pzn = SONDERPZN_WVO
    RezArtikel(AbInd).text = "WVO RabattArtikel"
    RezArtikel(AbInd).Faktor = Val(Left(WVORabattArtikel, 3))
    RezArtikel(AbInd).Preis = 0#
    RezArtikel(AbInd).flag = 0
    RezArtikel(AbInd).Zuz = 0
End If

'sVerfügbarkeit$ = CheckVerfügbarkeit
sVerfügbarkeit = Left(AMverfügbarkeit, 3)
If (sVerfügbarkeit$ <> "111") Then
    For i% = UBound(RezArtikel) - 1 To AbInd Step -1
        RezArtikel(i% + 1) = RezArtikel(i%)
    Next i%
    Call InitRezeptArtikel(AbInd)
'    Call InitKontrollStop(0)
    AnzRezeptArtikel% = AnzRezeptArtikel% + 1
    RezArtikel(AbInd).pzn = "02567024"
    RezArtikel(AbInd).text = "Nichtverfügbarkeit von AM"
    RezArtikel(AbInd).Faktor = Val(sVerfügbarkeit$)
    RezArtikel(AbInd).Preis = 0#
    RezArtikel(AbInd).flag = 0
    RezArtikel(AbInd).Zuz = 0
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
    
Call WinArtDebug("DruckeRezept")
If (IstDosDrucker%) Then
    erg% = DosRezeptDruck%
    Call WinArtDebug("nach DosRezeptDruck")
Else
    Call WinArtDebug("vor InitRezeptDruck")
    Call InitRezeptDruck
    Call WinArtDebug("vor RezeptDruck")
    Call RezeptDruck
    Call WinArtDebug("nach RezeptDruck")
    erg% = True
End If
If (erg%) And (AnzBlink% = 0) Then
    Call WinArtDebug("vor WriteRezeptSpeicher")
    Call WriteRezeptSpeicher
    Call WinArtDebug("nach WriteRezeptSpeicher")
End If

Call DefErrPop
End Sub

Sub RezeptPause(PauseSek%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("RezeptPause")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim StartSek, IstSek

StartSek = Timer
Do
    If ((Timer - StartSek) > PauseSek%) Then Exit Do
    DoEvents
Loop

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
ElseIf (dPreis# > 1000#) Then
    dPreis# = 1000#
End If

'If (TaxeRec!ZuzFrei) Then
'    dpreis# = 0
'End If

CalcZuzaNeu# = CDbl(Int(dPreis# + 0.501)) / 100#

If (HimiVerbrauch%) Then
    If (RezeptHolenZuz# > 0) Then
'        CalcZuzaNeu# = RezeptHolenZuz#
        CalcZuzaNeu# = RezeptHolenZuz# * iFaktor
    ElseIf ((CalcZuzaNeu# * iFaktor%) > 10#) Then
        If (iMsgBox("Soll die Zuzahlung auf 10 EURO begrenzt werden?", vbQuestion Or vbYesNo Or vbDefaultButton2) = vbYes) Then
            CalcZuzaNeu# = 10 / iFaktor%
        End If
    End If
End If

Call DefErrPop
End Function

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
            h$ = Trim(h$ + " " + Trim(CheckNullStr(KundenAdoRec!VorName)))
            h$ = Trim(h$ + " " + Trim(CheckNullStr(KundenAdoRec!Name)))
            RezKundenInfo$(0) = h$

            h$ = Trim(CheckNullStr(KundenAdoRec!strasse))
            h$ = Trim(h$ + " " + Trim(CheckNullStr(KundenAdoRec!HausNr)))
            RezKundenInfo$(1) = h$

            h$ = Trim(CheckNullStr(KundenAdoRec!PLZ))
            h$ = Trim(h$ + " " + Trim(CheckNullStr(KundenAdoRec!Ort)))
            RezKundenInfo$(2) = h$

            RezKundenInfo$(3) = Trim(CheckNullStr(KundenAdoRec!VersichertenNummer))

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
            h$ = Trim(h$ + " " + Trim(CheckNullStr(KundenRec!VorName)))
            h$ = Trim(h$ + " " + Trim(CheckNullStr(KundenRec!Name)))
            RezKundenInfo$(0) = h$
            
            h$ = Trim(CheckNullStr(KundenRec!strasse))
            h$ = Trim(h$ + " " + Trim(CheckNullStr(KundenRec!HausNr)))
            RezKundenInfo$(1) = h$
            
            h$ = Trim(CheckNullStr(KundenRec!PLZ))
            h$ = Trim(h$ + " " + Trim(CheckNullStr(KundenRec!Ort)))
            RezKundenInfo$(2) = h$
            
            RezKundenInfo$(3) = Trim(CheckNullStr(KundenRec!VersichertenNummer))
            
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
  
If (BtmRezept) Then 'And (BtmFaktorManuell = 0) Then
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
                If (CheckSonderPzn%(RezArtikel(i%).pzn) >= 0) Then
                    j = j + 1
                End If
            Else
                If (tRec!BTMKz) Then
                    j = j + 1
                End If
            End If
        End If
    Next i%
    If (ind >= 0) Then
        RezArtikel(ind).Faktor = j

'        iTabIndex% = AktTabIndex%
'        Call PaintArtikel(ind%)
'        Call TabIndexChange(iTabIndex%, False)
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
        
        iKeyPkl = CheckNullLong(VdbVOKGRec!Key_PKL)
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
                        
                        If (InStr(SollKeyPkl, "," + CStr(CheckNullLong(VdbVAPRec!Key_PKL)) + ",") > 0) Then
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

''?????????????
'SQLStr = "SELECT * FROM VdbVOKG"
''SQLStr = SQLStr + " WHERE KostentrNr=" + CStr(ActKasse)
''SQLStr = SQLStr + " AND LandNr=" + CStr(ActBundesland)
'SQLStr = SQLStr + " WHERE VebNr=" + CStr(ActVebNr)
'SQLStr = SQLStr + " AND VOrdTypNr=" + CStr(ActVerordnung)
'SQLStr = SQLStr + " AND Folgeausdruck<>'ABR302'"
'SQLStr = SQLStr + " ORDER BY Key_VOKG"
''Set VdbVOKGRec = AplusVDB.OpenRecordset(SQLStr)
'VdbVOKGRec.Open SQLStr, AplusVDB.ActiveConn
'If (VdbVOKGRec.EOF) Then
'    VerweisVebNr = 0
'    SQLStr = "SELECT * FROM VdbVereinbarungen WHERE VebNr=" + CStr(ActVebNr)
'    FabsErrf = AplusVDB.OpenRecordset(tRec, SQLStr)
'    If Not (tRec.EOF) Then
'        VerweisVebNr = CheckNullLong(tRec!VerweisVebNr)
'    End If
'    tRec.Close
'
'    If (VerweisVebNr > 0) Then
'        VdbVOKGRec.Close
'
'        SQLStr = "SELECT * FROM VdbVOKG"
'    '    SQLStr = SQLStr + " WHERE KostentrNr=" + CStr(ActKasse)
'    '    SQLStr = SQLStr + " AND LandNr=" + CStr(18)
'        SQLStr = SQLStr + " WHERE VebNr=" + CStr(VerweisVebNr)
'        SQLStr = SQLStr + " AND VOrdTypNr=" + CStr(ActVerordnung)
'        SQLStr = SQLStr + " AND Folgeausdruck<>'ABR302'"
'        SQLStr = SQLStr + " ORDER BY Key_VOKG"
'    '    Set VdbVOKGRec = AplusVDB.OpenRecordset(SQLStr)
'        VdbVOKGRec.Open SQLStr, AplusVDB.ActiveConn
'    End If
'End If
'
'If (VdbVOKGRec.EOF) Then
'    CheckVerordnungMDB = 0
'    Call DefErrPop: Exit Function
''ElseIf (CheckNullStr(VdbVOKGRec!Folgeausdruck) = "ABR302") Then
''    RezeptMitHilfsmittelNr% = 1
''    CheckVerordnungMDB = 0
''
''    For i% = 0 To (AnzRezeptArtikel% - 1)
''        h = RezArtikel(i).pzn
''        If (h = "-9999999") Then
''            RezeptMitHilfsmittelNr% = -1
''        ElseIf (Left(h, 1) <> "-") Then
''            SQLStr$ = "SELECT * FROM VdbArtikel WHERE PZN = " + h
''            FabsErrf = AplusVDB.OpenRecordset(VdbPznRec2, SQLStr)
''            If Not (VdbPznRec2.EOF) Then
''                hPosNr$ = Trim(Left$(CheckNullStr(VdbPznRec2!HmPositionsNr1), 10))
''                If (hPosNr$ <> "") Then
''                    RezArtikel(i).pzn = "-" + Left(h, 7)
''                    RezArtikel(i).ScreenPzn = hPosNr
''                End If
''            End If
''            VdbPznRec2.Close
''        End If
''    Next i
''
''    Call DefErrPop: Exit Function
'End If
'
'OrgVdbZeilen = ""
'For i% = 0 To (AnzRezeptArtikel% - 1)
'    pzn$ = RezArtikel(i).pzn
'    If (Left(pzn, 1) = "-") Then
'        pzn = Mid(pzn, 2) + "0"
'        pzn = MakePzn(pzn)
'    End If
'
'    SQLStr$ = "SELECT * FROM VdbArtikel WHERE PZN = " + pzn ' RezArtikel(i%).pzn
''    Set VdbPznRec = AplusVDB.OpenRecordset(SQLStr$)
'    FabsErrf = AplusVDB.OpenRecordset(VdbPznRec2, SQLStr)
'    If Not (VdbPznRec2.EOF) Then
'        OrgVdbZeilen = OrgVdbZeilen + "," + CStr(i) + ","
'    End If
'    VdbPznRec2.Close
'
'    h$ = ""
'    SQLStr = "SELECT * FROM VdbVAP"
'    SQLStr = SQLStr + " WHERE PZN=" + pzn   'RezArtikel(i%).pzn
'    SQLStr = SQLStr + " ORDER BY Key_PKL"
''    Set VdbVAPRec = AplusVDB.OpenRecordset(SQLStr)
'    VdbVAPRec.Open SQLStr, AplusVDB.ActiveConn
'    If Not (VdbVAPRec.EOF) Then
'        h$ = ","
'        Do
'            If (VdbVAPRec.EOF) Then
'                Exit Do
'            End If
'
'            h$ = h$ + CStr(CheckNullInt(VdbVAPRec!Key_PKL)) + ","
'
'            VdbVAPRec.MoveNext
'        Loop
'    End If
'    VdbVAPRec.Close
'    PznPkl(i) = h$
'    PznTreffer(i) = ""
'Next i
'
'Do
'    If (VdbVOKGRec.EOF) Then
'        Exit Do
'    End If
'
'    VdbZeilen = OrgVdbZeilen
'
'    SQLStr = "SELECT * FROM VdbVOKZ"
'    SQLStr = SQLStr + " WHERE Key_VOKG=" + CStr(VdbVOKGRec!Key_VOKG)
'    SQLStr = SQLStr + " ORDER BY Zaehler"
''    Set VdbVOKZRec = AplusVDB.OpenRecordset(SQLStr)
'    FabsErrf = AplusVDB.OpenRecordset(VdbVOKZRec, SQLStr)
'    If Not (VdbVOKZRec.EOF) Then
'        VdbZeilen = ""
'        Do
'            If (VdbVOKZRec.EOF) Then
'                Exit Do
'            End If
'
'            iKeyPkl = CheckNullInt(VdbVOKZRec!Key_PKL)
'            If (iKeyPkl > 0) Then
'                AnzKeyPkl(0) = 0
'                For i% = 0 To (AnzRezeptArtikel% - 1)
'                    If (InStr(PznPkl(i), "," + CStr(iKeyPkl) + ",") > 0) Then
'                        AnzKeyPkl(0) = AnzKeyPkl(0) + 1
'                    End If
'                Next i
'
'                AnzKeyPkl(1) = Val(CheckNullStr(VdbVOKZRec!Anz_Zeilen_Zahl))
'                ActOperator$ = CheckNullStr(VdbVOKZRec!Anz_Zeilen_Operator)
'
'                If (((ActOperator = "<") And (AnzKeyPkl(0) < AnzKeyPkl(1))) Or _
'                    ((ActOperator = "=") And (AnzKeyPkl(0) = AnzKeyPkl(1))) Or _
'                    ((ActOperator = ">") And (AnzKeyPkl(0) > AnzKeyPkl(1))) Or _
'                    ((ActOperator = "<=") And (AnzKeyPkl(0) <= AnzKeyPkl(1))) Or _
'                    ((ActOperator = "<>") And (AnzKeyPkl(0) <> AnzKeyPkl(1))) Or _
'                    ((ActOperator = ">=") And (AnzKeyPkl(0) >= AnzKeyPkl(1))) _
'                   ) Then
'                    For i% = 0 To (AnzRezeptArtikel% - 1)
'                        If (InStr(PznPkl(i), "," + CStr(iKeyPkl) + ",") > 0) Then
'                            VdbZeilen = VdbZeilen + "," + CStr(i) + ","
'                        End If
'                    Next i
'                Else
'                    CheckVerordnungMDB = 0
'                    Call DefErrPop: Exit Function
'                End If
'            End If
'
'            VdbVOKZRec.MoveNext
'        Loop
'    End If
'
'    ActBedingung$ = Trim(CheckNullStr(VdbVOKGRec!Bedingungsausdruck))
'    If (UCase(Left(ActBedingung$, 8)) = "VO_SUMME") Then
'        dWert(0) = 0
'        For i% = 0 To (AnzRezeptArtikel% - 1)
'            If (InStr(VdbZeilen, "," + CStr(i) + ",") > 0) Then
'                dWert(0) = dWert(0) + (RezArtikel(i%).Preis * RezArtikel(i%).Faktor)
'            End If
'        Next i
'
'        ActBedingung$ = Mid$(ActBedingung$, 9)
'        ActOperator = Left$(ActBedingung, 1)
'        ActBedingung$ = Mid$(ActBedingung$, 2)
'        If (Left$(ActBedingung, 1) = "=") Then
'            ActOperator = ActOperator + "="
'            ActBedingung$ = Mid$(ActBedingung$, 2)
'        End If
'        dWert(1) = xVal(ActBedingung)
'
'        If (((ActOperator = "<") And (dWert(0) < dWert(1))) Or _
'            ((ActOperator = "=") And (dWert(0) = dWert(1))) Or _
'            ((ActOperator = ">") And (dWert(0) > dWert(1))) Or _
'            ((ActOperator = "<=") And (dWert(0) <= dWert(1))) Or _
'            ((ActOperator = "<>") And (dWert(0) <> dWert(1))) Or _
'            ((ActOperator = ">=") And (dWert(0) >= dWert(1))) _
'           ) Then
'        Else
'            CheckVerordnungMDB = 0
'            Call DefErrPop: Exit Function
'        End If
'    End If
'
'    For i% = 0 To (AnzRezeptArtikel% - 1)
'        If (InStr(VdbZeilen, "," + CStr(i) + ",") > 0) Then
'            PznTreffer(i) = PznTreffer(i) + CStr(VdbVOKGRec!Key_VOKG) + ","
'        End If
'    Next i
'
'    VdbVOKGRec.MoveNext
'Loop
'
'erg = 0
'For i% = 0 To (AnzRezeptArtikel% - 1)
'    If (PznTreffer(i) <> "") Then
'        erg = True
'        Exit For
'    End If
'Next i
'
'If (erg = 0) Then
'    CheckVerordnungMDB = 0
'    Call DefErrPop: Exit Function
'End If
'
'
'InhaltsstoffeModus = 1
'Load frmInhaltsStoffe
'frmInhaltsStoffe.Caption = "Zeilenübergreifende Regelungen"
'
'With frmInhaltsStoffe.flxInhaltsStoffe
'    .Rows = .FixedRows
'    For i% = 0 To (AnzRezeptArtikel% - 1)
'        If (PznTreffer(i) <> "") Then
'            k = 0
'            .AddItem RezArtikel(i).text
'
'            h = PznTreffer(i)
'            Do
'                ind% = InStr(h, ",")
'                If (ind > 0) Then
'                    KeyVOKG& = Val(Left(h, ind - 1))
'
'                    SQLStr = "SELECT * FROM VdbVOKG"
'                    SQLStr = SQLStr + " WHERE Key_VOKG=" + CStr(KeyVOKG)
''                    Set VdbVOKGRec = AplusVDB.OpenRecordset(SQLStr)
'                    FabsErrf = AplusVDB.OpenRecordset(VdbVOKGRec, SQLStr)
'                    If Not (VdbVOKGRec.EOF) Then
'                        h2$ = CheckNullStr(VdbVOKGRec!Folgeausdruck)
'                        If (h2 = "G") Then
'                            MsgStr = "Genehmigungspflicht"
'                        ElseIf (h2 = "K") Then
'                            MsgStr = "Kostenvoranschlag"
'                        End If
'                        If (k > 0) Then
'                            .AddItem ""
'                        End If
'                        .TextMatrix(.Rows - 1, 1) = MsgStr
'                        .row = .Rows - 1
'                        .FillStyle = flexFillRepeat
'                        .col = 0
'                        .ColSel = .Cols - 1
'                        .CellFontBold = True
'                        .FillStyle = flexFillSingle
'                        k = k + 1
'
'                        ActBedingung$ = Trim(CheckNullStr(VdbVOKGRec!Bedingungsausdruck))
'                        If (ActBedingung <> "") Then
'                            h2 = " bei "
'                            If (UCase(Left(ActBedingung$, 8)) = "VO_SUMME") Then
'                                h2 = h2 + "Verordnungssumme" + Mid$(ActBedingung, 9)
'                            Else
'                                h2 = h2 + ActBedingung
'                            End If
'                            .AddItem vbTab + h2
'                        End If
'
'                        SQLStr = "SELECT * FROM VdbVOKZ"
'                        SQLStr = SQLStr + " WHERE Key_VOKG=" + CStr(VdbVOKGRec!Key_VOKG)
'                        SQLStr = SQLStr + " ORDER BY Zaehler"
''                        Set VdbVOKZRec = AplusVDB.OpenRecordset(SQLStr)
'                        FabsErrf = AplusVDB.OpenRecordset(VdbVOKZRec, SQLStr)
'                        If Not (VdbVOKZRec.EOF) Then
'                            j = 0
'                            h2 = ""
'                            Do
'                                If (VdbVOKZRec.EOF) Then
'                                    Exit Do
'                                End If
'
'                                h2 = ""
'                                If (j = 0) Then
'                                    h2 = h2 + "wenn "
'                                Else
'                                    h2 = h2 + "und "
'                                End If
'                                ActOperator = CheckNullStr(VdbVOKZRec!Anz_Zeilen_Operator)
'                                AnzKeyPkl(0) = Val(CheckNullStr(VdbVOKZRec!Anz_Zeilen_Zahl))
'                                h2 = h2 + ActOperator
'                                h2 = h2 + " " + CStr(AnzKeyPkl(0))
'                                h2 = h2 + " Zeilen mit Produktklasse "
'
'                                iKeyPkl = CheckNullInt(VdbVOKZRec!Key_PKL)
'                                SQLStr = "SELECT * FROM VdbPKL"
'                                SQLStr = SQLStr + " WHERE Key_PKL=" + CStr(iKeyPkl)
''                                Set VdbPKLRec = AplusVDB.OpenRecordset(SQLStr)
'                                FabsErrf = AplusVDB.OpenRecordset(VdbPKLRec, SQLStr)
'                                If (VdbPKLRec.EOF) Then
'                                    h2 = h2 + "(" + CStr(iKeyPkl) + ")"
'                                Else
'                                    h2 = h2 + CheckNullStr(VdbPKLRec!Name)
'                                End If
'                                .AddItem vbTab + h2
'
'                                j = j + 1
'                                VdbVOKZRec.MoveNext
'                            Loop
'                            If (j = 1) Then
'                                ind2 = InStr(h2, "> 0")
'                                If (ind2 = 0) Then
'                                    ind2 = InStr(h2, ">= 1")
'                                End If
'                                If (ind2 > 0) Then
'                                    h2 = "(Produktklasse "
'                                    If (VdbPKLRec.EOF) Then
'                                        h2 = h2 + CStr(iKeyPkl)
'                                    Else
'                                        h2 = h2 + CheckNullStr(VdbPKLRec!Name)
'                                    End If
'                                    h2 = h2 + ")"
'                                    .Rows = .Rows - 1
'                                    .AddItem vbTab + h2
'                                    End If
'                            End If
'                        End If
'                    End If
'
'                    h = Mid$(h, ind + 1)
'                Else
'                    Exit Do
'                End If
'            Loop
'        End If
'    Next i
'    .row = 1
'    .col = 0
'    .ColSel = .Cols - 1
'End With
'
'frmInhaltsStoffe.Show 1
'??????

CheckVerordnungMDB = True

Call DefErrPop
End Function



Function CheckNullInt%(s As Variant)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("CheckNullInt%")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
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

If (IsNull(s)) Then
    ret% = 0
Else
    ret% = s
End If

CheckNullInt% = ret%

Call DefErrPop
End Function


