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


Private Const DefErrModul = "clsKunden.cls"

Private Sub Class_Initialize()

dKurz$ = "KUNDEN"
dBezeichnung$ = "Kunden"
dName$ = "kdatei.dat"
dZugriff$ = "R"
dModus$ = "R"
dSatzLen% = SATZLEN 'Len(KundenRec)

dAnzFields% = 33
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
fTyp$(6) = "T"
fKonv%(6) = 0

fKurz$(7) = "RABATT"
fBezeichnung$(7) = "Rabatt"
fStart%(7) = 132
fLaenge%(7) = 4
fTyp$(7) = "T"
fKonv%(7) = 0

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
fTyp$(11) = "T"
fKonv%(11) = 0

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
fTyp$(13) = "T"
fKonv%(13) = 0

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
fTyp$(15) = "T"
fKonv%(15) = 0

fKurz$(16) = "SKTOVOR"
fBezeichnung$(16) = "Skonto Vorschlag"
fStart%(16) = 157
fLaenge%(16) = 4
fTyp$(16) = "T"
fKonv%(16) = 0

fKurz$(17) = "KTOVOR"
fBezeichnung$(17) = "Konto Vorschlag"
fStart%(17) = 161
fLaenge%(17) = 6
fTyp$(17) = "T"
fKonv%(17) = 0

fKurz$(18) = "UMSIST"
fBezeichnung$(18) = "Umsatz-Ist"
fStart%(18) = 167
fLaenge%(18) = 8
fTyp$(18) = "T"
fKonv%(18) = 0

fKurz$(19) = "OFP"
fBezeichnung$(19) = "Offene Posten"
fStart%(19) = 175
fLaenge%(19) = 8
fTyp$(19) = "T"
fKonv%(19) = 0

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

fKurz$(22) = "ENDBET"
fBezeichnung$(22) = "Endbetrag"
fStart%(22) = 202
fLaenge%(22) = 8
fTyp$(22) = "T"
fKonv%(22) = 0

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
fTyp$(28) = "T"
fKonv%(28) = 0

fKurz$(29) = "AUTOMAHN"
fBezeichnung$(29) = "AutoMahn"
fStart%(29) = 248
fLaenge%(29) = 1
fTyp$(29) = "T"
fKonv%(29) = 0

fKurz$(30) = "FILLER"
fBezeichnung$(30) = "Filler"
fStart%(30) = 249
fLaenge%(30) = 1
fTyp$(30) = "T"
fKonv%(30) = 0

fKurz$(31) = "BEZLINK"
fBezeichnung$(31) = "BezLink"
fStart%(31) = 250
fLaenge%(31) = 4
fTyp$(31) = "T"
fKonv%(31) = 0

fKurz$(32) = "LINK"
fBezeichnung$(32) = "Link"
fStart%(32) = 254
fLaenge%(32) = 2
fTyp$(32) = "N"
fKonv%(32) = 5




iAnzIndex% = 0

    
'Set clsKunden1 = Me

End Sub

Sub FreeClass()
'Set clsKunden1 = Nothing
End Sub

Sub OpenDatei()
Dim i%

If (RTrim$(dName$) <> "") Then
    dHandle% = clsDat.FileOpen(RTrim$(dName$), RTrim$(dZugriff$), dModus$, dSatzLen%)
End If
For i% = 0 To (iAnzIndex% - 1)
     FabsErrf% = clsDat.IndexOpen(iName$(i%), iHandle%(i%))
Next i%

End Sub

Sub SatzLock(Recno&)
Call clsDat.iLock(dHandle%, Recno&)
End Sub

Sub SatzUnLock(Recno&)
Call clsDat.iUnLock(dHandle%, Recno&)
End Sub

Sub GetRecord(Recno&)
Seek #dHandle%, Recno&
Get #dHandle%, , AktData
End Sub

Function IndexSearch%(ind%, key$, Recno&)
FabsErrf% = clsDat.IndexSearch%(iHandle%(ind%), key$, Recno&)
iFabsRecno& = Recno&
IndexSearch% = FabsErrf%
End Function

Function IndexGeneric%(ind%, key$, Recno&)
FabsErrf% = clsDat.IndexGeneric%(iHandle%(ind%), key$, Recno&)
iFabsRecno& = Recno&
IndexGeneric% = FabsErrf%
End Function

Function IndexNext%(ind%, AltRecno&, key$, Recno&)
FabsErrf% = clsDat.IndexNext%(iHandle%(ind%), AltRecno&, key$, Recno&)
iFabsRecno& = Recno&
IndexNext% = FabsErrf%
End Function

Function IndexPrevious%(ind%, AltRecno&, key$, Recno&)
FabsErrf% = clsDat.IndexPrevious%(iHandle%(ind%), AltRecno&, key$, Recno&)
iFabsRecno& = Recno&
IndexPrevious% = FabsErrf%
End Function

Function IndexInsert%(ind%, key$, Recno&)
FabsErrf% = clsDat.IndexInsert%(iHandle%(ind%), key$, Recno&)
iFabsRecno& = Recno&
IndexInsert% = FabsErrf%
End Function

Function SucheFeldInd%(Typ%, SuchFeld$)
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

End Function

Function Info$(iFeld$)
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

Public Property Get wert(iFeld$, Optional iIndex% = -1) As Variant
Attribute wert.VB_UserMemId = 0
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

wert = ret
End Property



