VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsKunden"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Const SATZLEN = 255

Private Type KundenErstSatz
    erstmax As Integer
    erstrest As String * 98
End Type


Private Type Kun11
    Name            As String * 30
End Type
Private Type Kun2
    ZahlungBis      As Integer
    skonto          As Single
End Type
Private Type KundenStruct
    kName(3)          As Kun11
    kurz              As String * 6
    KuPr              As String * 1
    Prozente          As Single
    Rabatt            As Single
    kopien            As Integer
    ausdruck          As String * 1
    Zahlung(2)        As Kun2
    SkontoVorschlag   As Single
    KtoVor            As String * 6     'Konto Vorschlag
    UmsatzIst         As Double
    OffenePosten      As Double
    leRecDat          As String * 6     'letzte Rechnung Datum
    Beleg             As String * 13    'Belegnummer
    EndBetrag         As Double
    Auswahl           As String * 20
    geburt            As String * 6
    ProzentGruppe     As String * 1
    UmsList           As String * 1     'Umsatzliste ("J"/"N")
    SammKnr           As Integer        'Sammelkundennummer
    Offen             As Double
    AutoMahn          As String * 1
    PhCare            As String * 1
    Profil            As Integer
    BezLink           As Integer
    link              As Integer
End Type
Dim kun As KundenStruct
Dim kunErst As KundenErstSatz

Dim iName$(2)
Dim iHandle%(2)
Dim iAnzIndex%

Dim dKurz$
Dim dBezeichnung$
Dim dName$
Dim dZugriff$
Dim dModus$
Dim dSatzLen%
Dim dHandle%
Dim dAnzFields%

Dim fKurz$()
Dim fBezeichnung$()
Dim fStart%()
Dim fLaenge%()
Dim fTyp$()
Dim fKonv%()

Dim FabsErrf%
Dim iFabsRecno&

Dim AktData As String * SATZLEN



Private Const DefErrModul = "KUNDEN.CLS"

Private Sub Class_Initialize()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("Class_Initialize")
Call clsError.DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case clsError.DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call ProjektForm.EndeDll
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
dKurz$ = "KUNDEN"
dBezeichnung$ = "Kunden"
dName$ = "kdatei.dat"
dZugriff$ = "RW"
dModus$ = "R"
dSatzLen% = SATZLEN 'Len(KundenRec)

dAnzFields% = 34
ReDim fKurz$(dAnzFields% - 1)
ReDim fBezeichnung$(dAnzFields% - 1)
ReDim fStart%(dAnzFields% - 1)
ReDim fLaenge%(dAnzFields% - 1)
ReDim fTyp$(dAnzFields% - 1)
ReDim fKonv%(dAnzFields% - 1)

fKurz$(0) = "NAME(0)"
fBezeichnung$(0) = "Name(0)"
fStart%(0) = 1
fLaenge%(0) = 30
fTyp$(0) = "T"
fKonv%(0) = 0

fKurz$(1) = "NAME(1)"
fBezeichnung$(1) = "Name(1)"
fStart%(1) = 31
fLaenge%(1) = 30
fTyp$(1) = "T"
fKonv%(1) = 0

fKurz$(2) = "NAME(2)"
fBezeichnung$(2) = "Name(2)"
fStart%(2) = 61
fLaenge%(2) = 30
fTyp$(2) = "T"
fKonv%(2) = 0

fKurz$(3) = "NAME(3)"
fBezeichnung$(3) = "Name(3)"
fStart%(3) = 91
fLaenge%(3) = 30
fTyp$(3) = "T"
fKonv%(3) = 0

fKurz$(4) = "KURZ"
fBezeichnung$(4) = "Kurz"
fStart%(4) = 121
fLaenge%(4) = 6
fTyp$(4) = "T"
fKonv%(4) = 0

fKurz$(5) = "KUPR"
fBezeichnung$(5) = "Kupr"
fStart%(5) = 127
fLaenge%(5) = 1
fTyp$(5) = "T"
fKonv%(5) = 0

fKurz$(6) = "PROZENTE"
fBezeichnung$(6) = "Prozente"
fStart%(6) = 128
fLaenge%(6) = 4
fTyp$(6) = "N"
fKonv%(6) = 4

fKurz$(7) = "RABATT"
fBezeichnung$(7) = "Rabatt"
fStart%(7) = 132
fLaenge%(7) = 4
fTyp$(7) = "N"
fKonv%(7) = 4

fKurz$(8) = "KOPIEN"
fBezeichnung$(8) = "Kopien"
fStart%(8) = 136
fLaenge%(8) = 2
fTyp$(8) = "N"
fKonv%(8) = 5

fKurz$(9) = "AUSDRUCK"
fBezeichnung$(9) = "Ausdruck"
fStart%(9) = 138
fLaenge%(9) = 1
fTyp$(9) = "T"
fKonv%(9) = 0

fKurz$(10) = "ZAHLUNGBIS(0)"
fBezeichnung$(10) = "ZahlungBis(0)"
fStart%(10) = 139
fLaenge%(10) = 2
fTyp$(10) = "N"
fKonv%(10) = 5

fKurz$(11) = "SKONTO(0)"
fBezeichnung$(11) = "Skonto(0)"
fStart%(11) = 141
fLaenge%(11) = 4
fTyp$(11) = "N"
fKonv%(11) = 4

fKurz$(12) = "ZAHLUNGBIS(1)"
fBezeichnung$(12) = "ZahlungBis(1)"
fStart%(12) = 145
fLaenge%(12) = 2
fTyp$(12) = "N"
fKonv%(12) = 5

fKurz$(13) = "SKONTO(1)"
fBezeichnung$(13) = "Skonto(1)"
fStart%(13) = 147
fLaenge%(13) = 4
fTyp$(13) = "N"
fKonv%(13) = 4

fKurz$(14) = "ZAHLUNGBIS(2)"
fBezeichnung$(14) = "ZahlungBis(2)"
fStart%(14) = 151
fLaenge%(14) = 2
fTyp$(14) = "N"
fKonv%(14) = 5

fKurz$(15) = "SKONTO(2)"
fBezeichnung$(15) = "Skonto(2)"
fStart%(15) = 153
fLaenge%(15) = 4
fTyp$(15) = "N"
fKonv%(15) = 4

fKurz$(16) = "SKTOVOR"
fBezeichnung$(16) = "Skonto Vorschlag"
fStart%(16) = 157
fLaenge%(16) = 4
fTyp$(16) = "N"
fKonv%(16) = 4

fKurz$(17) = "KTOVOR"
fBezeichnung$(17) = "Konto Vorschlag"
fStart%(17) = 161
fLaenge%(17) = 6
fTyp$(17) = "T"
fKonv%(17) = 0

fKurz$(18) = "UMSATZIST"
fBezeichnung$(18) = "Umsatz-Ist"
fStart%(18) = 167
fLaenge%(18) = 8
fTyp$(18) = "N"
fKonv%(18) = 2

fKurz$(19) = "OFP"
fBezeichnung$(19) = "Offene Posten"
fStart%(19) = 175
fLaenge%(19) = 8
fTyp$(19) = "N"
fKonv%(19) = 2

fKurz$(20) = "LERECDAT"
fBezeichnung$(20) = "Letzte Rechnung Datum"
fStart%(20) = 183
fLaenge%(20) = 6
fTyp$(20) = "T"
fKonv%(20) = 0

fKurz$(21) = "BELEG"
fBezeichnung$(21) = "Belegnummer"
fStart%(21) = 189
fLaenge%(21) = 13
fTyp$(21) = "T"
fKonv%(21) = 0

fKurz$(22) = "ENDBETRAG"
fBezeichnung$(22) = "Endbetrag"
fStart%(22) = 202
fLaenge%(22) = 8
fTyp$(22) = "N"
fKonv%(22) = 2

fKurz$(23) = "AUSWAHL"
fBezeichnung$(23) = "Auswahl"
fStart%(23) = 210
fLaenge%(23) = 20
fTyp$(23) = "T"
fKonv%(23) = 0

fKurz$(24) = "GEBURT"
fBezeichnung$(24) = "Geburt"
fStart%(24) = 230
fLaenge%(24) = 6
fTyp$(24) = "T"
fKonv%(24) = 0

fKurz$(25) = "PRZGRP"
fBezeichnung$(25) = "Prozentgruppe"
fStart%(25) = 236
fLaenge%(25) = 1
fTyp$(25) = "T"
fKonv%(25) = 0

fKurz$(26) = "UMSLIST"
fBezeichnung$(26) = "Umsatzliste"
fStart%(26) = 237
fLaenge%(26) = 1
fTyp$(26) = "T"
fKonv%(26) = 0

fKurz$(27) = "SAMMKNR"
fBezeichnung$(27) = "Sammelkundennummer"
fStart%(27) = 238
fLaenge%(27) = 2
fTyp$(27) = "N"
fKonv%(27) = 5

fKurz$(28) = "OFFEN"
fBezeichnung$(28) = "Offen"
fStart%(28) = 240
fLaenge%(28) = 8
fTyp$(28) = "N"
fKonv%(28) = 2

fKurz$(29) = "AUTOMAHN"
fBezeichnung$(29) = "AutoMahn"
fStart%(29) = 248
fLaenge%(29) = 1
fTyp$(29) = "T"
fKonv%(29) = 0

fKurz$(30) = "PHCARE"
fBezeichnung$(30) = "PhCare"
fStart%(30) = 249
fLaenge%(30) = 1
fTyp$(30) = "T"
fKonv%(30) = 0

fKurz$(31) = "PROFIL"
fBezeichnung$(31) = "ProfilNr"
fStart%(31) = 250
fLaenge%(31) = 2
fTyp$(31) = "N"
fKonv%(31) = 5

fKurz$(32) = "BEZLINK"
fBezeichnung$(32) = "BezLink"
fStart%(32) = 252
fLaenge%(32) = 2
fTyp$(32) = "N"
fKonv%(32) = 5

fKurz$(33) = "LINK"
fBezeichnung$(33) = "Link"
fStart%(33) = 254
fLaenge%(33) = 2
fTyp$(33) = "N"
fKonv%(33) = 5



iAnzIndex% = 2
iName$(0) = "KDATEI.IDX"
iName$(1) = "KUVNR.IDX"

If (Para1.KuEan) Then
    iAnzIndex% = 3
    iName$(2) = "KUEAN.IDX"
End If

Set Kun1 = Me

Call clsError.DefErrPop
End Sub

Sub FreeClass()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("FreeClass")
Call clsError.DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case clsError.DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call ProjektForm.EndeDll
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

Set Kun1 = Nothing

Call clsError.DefErrPop
End Sub

Sub OpenDatei()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("OpenDatei")
Call clsError.DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case clsError.DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call ProjektForm.EndeDll
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim i%
Dim iZugriff$

iZugriff$ = ""
If (iZugriff$ = "") Then
    iZugriff$ = RTrim$(dZugriff$)
End If

If (RTrim$(dName$) <> "") Then
    dHandle% = clsDat.FileOpen(RTrim$(dName$), iZugriff$, dModus$, dSatzLen%)
'    If (dHandle% = 0) Then
'        OpenDatei% = 1
'        Call clsError.DefErrPop: Exit Sub
'    End If
End If
For i% = 0 To (iAnzIndex% - 1)
    If (Para1.Land = "D") Or (i% <> 1) Then
        FabsErrf% = clsDat.IndexOpen(iName$(i%), iHandle%(i%))
'       If (FabsErrf%) Then
'           OpenDatei% = 1
'           Call clsError.DefErrPop: Exit Sub
'        End If
    End If
Next i%

'OpenDatei% = 0

Call clsError.DefErrPop
End Sub

Sub CloseDatei()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("CloseDatei")
Call clsError.DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case clsError.DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call ProjektForm.EndeDll
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim i%

If (dHandle% > 0) Then
    Close #dHandle%
    dHandle% = 0
End If
For i% = 0 To (iAnzIndex% - 1)
    If (Para1.Land = "D") Or (i% <> 1) Then
        FabsErrf% = clsDat.IndexClose(iHandle%(i%))
    End If
Next i%

Call clsError.DefErrPop
End Sub

Sub SatzLock(recno&)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("SatzLock")
Call clsError.DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case clsError.DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call ProjektForm.EndeDll
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsDat.iLock(dHandle%, recno&)
Call clsError.DefErrPop
End Sub

Sub SatzUnLock(recno&)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("SatzUnLock")
Call clsError.DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case clsError.DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call ProjektForm.EndeDll
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsDat.iUnLock(dHandle%, recno&)
Call clsError.DefErrPop
End Sub

Sub GetRecord(recno&)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("GetRecord")
Call clsError.DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case clsError.DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call ProjektForm.EndeDll
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Seek #dHandle%, recno&
If (recno& = 1) Then
'    Call GetFirstRecord
    Get #dHandle%, 1, kunErst
Else
    Get #dHandle%, , kun
    Seek #dHandle%, recno&
    Get #dHandle%, , AktData
End If
Call clsError.DefErrPop
End Sub
'Sub GetRecord(recno&)
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Call clsError.DefErrFnc("GetRecord")
'Call clsError.DefErrMod(DefErrModul)
'On Error GoTo DefErr
'GoTo DefErrEnd
'DefErr:
'Select Case clsError.DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
'Case vbRetry
'  Resume
'Case vbIgnore
'  Resume Next
'End Select
'Call ProjektForm.EndeDll
'DefErrEnd:
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Seek #dHandle%, recno&
'Get #dHandle%, , AktData
'Call clsError.DefErrPop
'End Sub

Function IndexSearch%(ind%, key$, recno&)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("IndexSearch%")
Call clsError.DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case clsError.DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call ProjektForm.EndeDll
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
FabsErrf% = clsDat.IndexSearch%(iHandle%(ind%), key$, recno&)
iFabsRecno& = recno&
IndexSearch% = FabsErrf%
Call clsError.DefErrPop
End Function

Function IndexGeneric%(ind%, key$, recno&)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("IndexGeneric%")
Call clsError.DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case clsError.DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call ProjektForm.EndeDll
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
FabsErrf% = clsDat.IndexGeneric%(iHandle%(ind%), key$, recno&)
iFabsRecno& = recno&
IndexGeneric% = FabsErrf%
Call clsError.DefErrPop
End Function

Function IndexNext%(ind%, AltRecno&, key$, recno&)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("IndexNext%")
Call clsError.DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case clsError.DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call ProjektForm.EndeDll
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
FabsErrf% = clsDat.IndexNext%(iHandle%(ind%), AltRecno&, key$, recno&)
iFabsRecno& = recno&
IndexNext% = FabsErrf%
Call clsError.DefErrPop
End Function

Function IndexPrevious%(ind%, AltRecno&, key$, recno&)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("IndexPrevious%")
Call clsError.DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case clsError.DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call ProjektForm.EndeDll
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
FabsErrf% = clsDat.IndexPrevious%(iHandle%(ind%), AltRecno&, key$, recno&)
iFabsRecno& = recno&
IndexPrevious% = FabsErrf%
Call clsError.DefErrPop
End Function

Function IndexInsert%(ind%, key$, recno&)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("IndexInsert%")
Call clsError.DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case clsError.DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call ProjektForm.EndeDll
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
FabsErrf% = clsDat.IndexInsert%(iHandle%(ind%), key$, recno&)
iFabsRecno& = recno&
IndexInsert% = FabsErrf%
Call clsError.DefErrPop
End Function

Function IndexDelete%(ind%, AltRecno&, key$, recno&)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("IndexDelete%")
Call clsError.DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case clsError.DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call ProjektForm.EndeDll
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
FabsErrf% = clsDat.IndexDelete%(iHandle%(ind%), AltRecno&, key$, recno&)
iFabsRecno& = recno&
IndexDelete% = FabsErrf%
Call clsError.DefErrPop
End Function

Function SucheFeldInd%(Typ%, SuchFeld$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("SucheFeldInd%")
Call clsError.DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case clsError.DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call ProjektForm.EndeDll
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim i%, ret%

ret% = -1

For i% = 0 To (dAnzFields% - 1)
    If (Typ%) Then
        If (fBezeichnung$(i%) = SuchFeld$) Then
            ret% = i%
            Exit For
        End If
    Else
        If (fKurz$(i%) = SuchFeld$) Then
            ret% = i%
            Exit For
        End If
    End If
Next i%

SucheFeldInd% = ret%

Call clsError.DefErrPop
End Function

Function Info$(iFeld$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("Info$")
Call clsError.DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case clsError.DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call ProjektForm.EndeDll
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim iValue As Variant
Dim ind%
Dim ret$, s$

ret$ = ""

s$ = UCase$(iFeld$)

ind% = SucheFeldInd%(0, s$)
If (ind% >= 0) Then
    iValue = clsDat.FeldValue(Mid$(AktData$, fStart%(ind%), fLaenge%(ind%)), fTyp$(ind%), fKonv%(ind%))
    ret$ = clsDat.DecodFeld(iValue, fTyp$(ind%), fKonv%(ind%))
End If

Info$ = ret$

Call clsError.DefErrPop
End Function

Public Property Get FeldBezeichnung$(iFeld As Variant)
Dim ind%
Dim ret$, s$

ret$ = ""

If (VarType(iFeld) = vbString) Then
    s$ = UCase$(iFeld)
    ind% = SucheFeldInd%(0, s$)
Else
    ind% = iFeld
End If
If (ind% >= 0) Then
    ret$ = fBezeichnung$(ind%)
End If

FeldBezeichnung$ = ret$

End Property

Public Property Get FeldKurz$(iFeld As Variant)
Dim ind%
Dim ret$, s$

ret$ = ""

If (VarType(iFeld) = vbString) Then
    s$ = UCase$(iFeld)
    ind% = SucheFeldInd%(1, s$)
Else
    ind% = iFeld
End If
If (ind% >= 0) Then
    ret$ = fKurz$(ind%)
End If

FeldKurz = ret$

End Property

Public Property Get DateiBezeichnung$()
DateiBezeichnung$ = dBezeichnung$
End Property

Public Property Get DateiKurz$()
DateiKurz$ = dKurz$
End Property

Public Property Get AnzFelder%()
AnzFelder% = dAnzFields%
End Property

Public Property Get DateiLen&()
DateiLen& = LOF(dHandle%)
End Property

Public Property Get Wert(iFeld$, Optional iIndex% = -1) As Variant
Attribute Wert.VB_UserMemId = 0
Dim ind%
Dim s$
Dim ret As Variant

ret = ""

s$ = UCase$(iFeld$)

If (iIndex% >= 0) Then
    s$ = s$ + "(" + Mid$(Str$(iIndex%), 2) + ")"
End If

ind% = SucheFeldInd%(0, s$)
If (ind% >= 0) Then
    ret = clsDat.FeldValue(Mid$(AktData$, fStart%(ind%), fLaenge%(ind%)), fTyp$(ind%), fKonv%(ind%))
End If

Wert = ret
End Property

Public Property Get RawData$()
Dim temp$

temp$ = String(SATZLEN, 0)
CopyMemory ByVal temp$, kun, SATZLEN
RawData$ = temp$

End Property




'
'Public Property Get erstmax%()
'erstmax% = kunErst.erstmax
'End Property
'Public Property Let erstmax(ByVal vNewValue%)
'kunErst.erstmax = vNewValue%
'End Property
'
'Public Property Get erstrest$()
'erstrest$ = kunErst.erstrest
'End Property
'Public Property Let erstrest(ByVal vNewValue$)
'kunErst.erstrest = vNewValue$
'End Property
'
'
'
'Public Property Get Name$(ByVal ind%)
'Dim h$
'h$ = kun.kName(ind%).Name
'Call OemToChar(h$, h$)
'Name$ = h$
'End Property
'Public Property Let Name(ByVal ind%, ByVal vNewValue$)
'Dim h$
'h$ = vNewValue$
'Call CharToOem(h$, h$)
'kun.kName(ind%).Name = h$
'End Property
'
'Public Property Get kurz$()
'Dim h$
'h$ = kun.kurz
'Call OemToChar(h$, h$)
'kurz$ = h$
'End Property
'Public Property Let kurz(ByVal vNewValue$)
'Dim h$
'h$ = vNewValue$
'Call CharToOem(h$, h$)
'kun.kurz = h$
'End Property
'
'Public Property Get KuPr$()
'KuPr$ = kun.KuPr
'End Property
'Public Property Let KuPr(ByVal vNewValue$)
'kun.KuPr = vNewValue$
'End Property
'
'Public Property Get Prozente!()
'Dim a!
'a! = kun.Prozente
'Call DxToIEEEs(a!)
'Prozente! = a!
'End Property
'Public Property Let Prozente(ByVal vNewValue!)
'Dim a!
'a! = vNewValue!
'Call DxToMBFs(a!)
'kun.Prozente = a!
'End Property
'
'Public Property Get Rabatt!()
'Dim a!
'a! = kun.Rabatt
'Call DxToIEEEs(a!)
'Rabatt! = a!
'End Property
'Public Property Let Rabatt(ByVal vNewValue!)
'Dim a!
'a! = vNewValue!
'Call DxToMBFs(a!)
'kun.Rabatt = a!
'End Property
'
'Public Property Get kopien%()
'kopien% = kun.kopien
'End Property
'Public Property Let kopien(ByVal vNewValue%)
'kun.kopien = vNewValue%
'End Property
'
'Public Property Get ausdruck$()
'ausdruck$ = kun.ausdruck
'End Property
'Public Property Let ausdruck(ByVal vNewValue$)
'kun.ausdruck = vNewValue$
'End Property
'
'Public Property Get ZahlungBis%(ByVal ind%)
'ZahlungBis% = kun.Zahlung(ind%).ZahlungBis
'End Property
'Public Property Let ZahlungBis(ByVal ind%, ByVal vNewValue%)
'kun.Zahlung(ind%).ZahlungBis = vNewValue%
'End Property
'
'Public Property Get ZahlungSkonto!(ByVal ind%)
'Dim a!
'a! = kun.Zahlung(ind%).skonto
'Call DxToIEEEs(a!)
'ZahlungSkonto! = a!
'End Property
'Public Property Let ZahlungSkonto(ByVal ind%, ByVal vNewValue!)
'Dim a!
'a! = vNewValue!
'Call DxToMBFs(a!)
'kun.Zahlung(ind%).skonto = a!
'End Property
'
'Public Property Get SkontoVorschlag!()
'Dim a!
'a! = kun.SkontoVorschlag
'Call DxToIEEEs(a!)
'SkontoVorschlag! = a!
'End Property
'Public Property Let SkontoVorschlag(ByVal vNewValue!)
'Dim a!
'a! = vNewValue!
'Call DxToMBFs(a!)
'kun.SkontoVorschlag = a!
'End Property
'
'Public Property Get KtoVor$()
'KtoVor$ = kun.KtoVor
'End Property
'Public Property Let KtoVor(ByVal vNewValue$)
'kun.KtoVor = vNewValue$
'End Property
'
'Public Property Get UmsatzIst#()
'Dim a#
'a# = kun.UmsatzIst
'Call DxToIEEEd(a#)
'UmsatzIst# = a#
'End Property
'Public Property Let UmsatzIst(ByVal vNewValue#)
'Dim a#
'a# = vNewValue#
'Call DxToMBFd(a#)
'kun.UmsatzIst = a#
'End Property
'
'Public Property Get OffenePosten#()
'Dim a#
'a# = kun.OffenePosten
'Call DxToIEEEd(a#)
'OffenePosten# = a#
'End Property
'Public Property Let OffenePosten(ByVal vNewValue#)
'Dim a#
'a# = vNewValue#
'Call DxToMBFd(a#)
'kun.OffenePosten = a#
'End Property
'
'Public Property Get leRecDat$()
'leRecDat$ = kun.leRecDat
'End Property
'Public Property Let leRecDat(ByVal vNewValue$)
'kun.leRecDat = vNewValue$
'End Property
'
'Public Property Get Beleg$()
'Beleg$ = kun.Beleg
'End Property
'Public Property Let Beleg(ByVal vNewValue$)
'kun.Beleg = vNewValue$
'End Property
'
'Public Property Get EndBetrag#()
'Dim a#
'a# = kun.EndBetrag
'Call DxToIEEEd(a#)
'EndBetrag# = a#
'End Property
'Public Property Let EndBetrag(ByVal vNewValue#)
'Dim a#
'a# = vNewValue#
'Call DxToMBFd(a#)
'kun.EndBetrag = a#
'End Property
'
'Public Property Get Auswahl$()
'Dim h$
'h$ = kun.Auswahl
'Call OemToChar(h$, h$)
'Auswahl$ = h$
'End Property
'Public Property Let Auswahl(ByVal vNewValue$)
'Dim h$
'h$ = vNewValue$
'Call CharToOem(h$, h$)
'kun.Auswahl = h$
'End Property
'
'Public Property Get geburt$()
'geburt$ = kun.geburt
'End Property
'Public Property Let geburt(ByVal vNewValue$)
'kun.geburt = vNewValue$
'End Property
'
'Public Property Get ProzentGruppe$()
'ProzentGruppe$ = kun.ProzentGruppe
'End Property
'Public Property Let ProzentGruppe(ByVal vNewValue$)
'kun.ProzentGruppe = vNewValue$
'End Property
'
'Public Property Get UmsatzListe() As Byte
'UmsatzListe = Abs(UCase(kun.UmsList) = "J")
'End Property
'Public Property Let UmsatzListe(ByVal vNewValue As Byte)
'If (vNewValue) Then
'    kun.UmsList = "J"
'Else
'    kun.UmsList = "N"
'End If
'End Property
'
'Public Property Get SammelKnr%()
'SammelKnr% = kun.SammKnr
'End Property
'Public Property Let SammelKnr(ByVal vNewValue%)
'kun.SammKnr = vNewValue%
'End Property
'
'Public Property Get Offen#()
'Dim a#
'a# = kun.Offen
'Call DxToIEEEd(a#)
'Offen# = a#
'End Property
'Public Property Let Offen(ByVal vNewValue#)
'Dim a#
'a# = vNewValue#
'Call DxToMBFd(a#)
'kun.Offen = a#
'End Property
'
'Public Property Get AutoMahn() As Byte
'AutoMahn = Abs(UCase(kun.AutoMahn) = "J")
'End Property
'Public Property Let AutoMahn(ByVal vNewValue As Byte)
'If (vNewValue) Then
'    kun.AutoMahn = "J"
'Else
'    kun.AutoMahn = "N"
'End If
'End Property
'
'Public Property Get PhCare() As Byte
'PhCare = Abs(UCase(kun.PhCare) = "J")
'End Property
'Public Property Let PhCare(ByVal vNewValue As Byte)
'If (vNewValue) Then
'    kun.PhCare = "J"
'Else
'    kun.PhCare = "N"
'End If
'End Property
'
'Public Property Get Profil%()
'Profil% = kun.Profil
'End Property
'Public Property Let Profil(ByVal vNewValue%)
'kun.Profil = vNewValue%
'End Property
'
'Public Property Get BezLink%()
'BezLink% = kun.BezLink
'End Property
'Public Property Let BezLink(ByVal vNewValue%)
'kun.BezLink = vNewValue%
'End Property
'
'Public Property Get link%()
'link% = kun.link
'End Property
'Public Property Let link(ByVal vNewValue%)
'kun.link = vNewValue%
'End Property
'
'
