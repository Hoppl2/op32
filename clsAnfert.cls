VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsAnfert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

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
Dim FabsRecno&

Dim AktData As String * 48


Private Const DefErrModul = "clsAnfert.cls"

Private Sub Class_Initialize()

dKurz$ = "ANFERT"
dBezeichnung$ = "Anfert"
dName$ = "anfert.dat"
dZugriff$ = "R"
dModus$ = "R"
dSatzLen% = 48 'Len(anfert)

dAnzFields% = 21
ReDim fKurz$(dAnzFields%)
ReDim fBezeichnung$(dAnzFields%)
ReDim fStart%(dAnzFields%)
ReDim fLaenge%(dAnzFields%)
ReDim fTyp$(dAnzFields%)
ReDim fKonv%(dAnzFields%)


fKurz$(0) = "STATUS"
fBezeichnung$(0) = "Status"
fStart%(0) = 1
fLaenge%(0) = 1
fTyp$(0) = "T"
fKonv%(0) = 0

fKurz$(1) = "VONWANN"
fBezeichnung$(1) = "VonWann"
fStart%(1) = 2
fLaenge%(1) = 4
fTyp$(1) = "T"
fKonv%(1) = 0

fKurz$(2) = "VONWO"
fBezeichnung$(2) = "VonWo"
fStart%(2) = 6
fLaenge%(2) = 1
fTyp$(2) = "N"
fKonv%(2) = 1

fKurz$(3) = "MITARBEITER"
fBezeichnung$(3) = "Mitarbeiter"
fStart%(3) = 7
fLaenge%(3) = 1
fTyp$(3) = "T"
fKonv%(3) = 0

fKurz$(4) = "ARBEIT"
fBezeichnung$(4) = "Arbeit"
fStart%(4) = 8
fLaenge%(4) = 5
fTyp$(4) = "T"
fKonv%(4) = 0

fKurz$(5) = "ANGEF"
fBezeichnung$(5) = "Angef"
fStart%(5) = 13
fLaenge%(5) = 1
fTyp$(5) = "T"
fKonv%(5) = 0

fKurz$(6) = "ANGEFWANN"
fBezeichnung$(6) = "AngefWann"
fStart%(6) = 14
fLaenge%(6) = 4
fTyp$(6) = "T"
fKonv%(6) = 0

fKurz$(7) = "ANGEFVONWO"
fBezeichnung$(7) = "AngefVonWo"
fStart%(7) = 18
fLaenge%(7) = 1
fTyp$(7) = "T"
fKonv%(7) = 0

fKurz$(8) = "ANFBLOCK(0)"
fBezeichnung$(8) = "Anfblock(0)"
fStart%(8) = 19
fLaenge%(8) = 2
fTyp$(8) = "N"
fKonv%(8) = 5

fKurz$(9) = "ANFBLOCK(1)"
fBezeichnung$(9) = "Anfblock(1)"
fStart%(9) = 21
fLaenge%(9) = 2
fTyp$(9) = "N"
fKonv%(9) = 5

fKurz$(10) = "ANFBLOCK(2)"
fBezeichnung$(10) = "Anfblock(2)"
fStart%(10) = 23
fLaenge%(10) = 2
fTyp$(10) = "N"
fKonv%(10) = 5

fKurz$(11) = "ANFBLOCK(3)"
fBezeichnung$(11) = "Anfblock(3)"
fStart%(11) = 25
fLaenge%(11) = 2
fTyp$(11) = "N"
fKonv%(11) = 5

fKurz$(12) = "ANFBLOCK(4)"
fBezeichnung$(12) = "Anfblock(4)"
fStart%(12) = 27
fLaenge%(12) = 2
fTyp$(12) = "N"
fKonv%(12) = 5

fKurz$(13) = "ANFBLOCK(5)"
fBezeichnung$(13) = "Anfblock(5)"
fStart%(13) = 29
fLaenge%(13) = 2
fTyp$(13) = "N"
fKonv%(13) = 5

fKurz$(14) = "ANFBLOCK(6)"
fBezeichnung$(14) = "Anfblock(6)"
fStart%(14) = 31
fLaenge%(14) = 2
fTyp$(14) = "N"
fKonv%(14) = 5

fKurz$(15) = "ANFBLOCK(7)"
fBezeichnung$(15) = "Anfblock(7)"
fStart%(15) = 33
fLaenge%(15) = 2
fTyp$(15) = "N"
fKonv%(15) = 5

fKurz$(16) = "ANFBLOCK(8)"
fBezeichnung$(16) = "Anfblock(8)"
fStart%(16) = 35
fLaenge%(16) = 2
fTyp$(16) = "N"
fKonv%(16) = 5

fKurz$(17) = "ANFBLOCK(9)"
fBezeichnung$(17) = "Anfblock(9)"
fStart%(17) = 37
fLaenge%(17) = 2
fTyp$(17) = "N"
fKonv%(17) = 5

fKurz$(18) = "KUNDENNR"
fBezeichnung$(18) = "KundenNr"
fStart%(18) = 39
fLaenge%(18) = 2
fTyp$(18) = "N"
fKonv%(18) = 5

fKurz$(19) = "REST"
fBezeichnung$(19) = "Rest"
fStart%(19) = 41
fLaenge%(19) = 8
fTyp$(19) = "T"
fKonv%(19) = 0

fKurz$(20) = "ERST"
fBezeichnung$(20) = "Erst"
fStart%(20) = 1
fLaenge%(20) = 48
fTyp$(20) = "T"
fKonv%(20) = 0



iAnzIndex% = 0


Set clsAnfert1 = Me

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

Function IndexSearch%(ind%, key$, Recno&)

FabsErrf% = clsDat.IndexSearch%(iHandle%(ind%), key$, Recno&)

FabsRecno& = Recno&

IndexSearch% = FabsErrf%

End Function

Sub GetRecord(Recno&)

Seek #dHandle%, Recno&
Get #dHandle%, , AktData

End Sub

Function SucheFeldInd%(ByVal typ%, ByVal SuchFeld$)
Dim i%, ret%

ret% = -1

For i% = 0 To (dAnzFields% - 1)
    If (typ%) Then
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

Function Value(iFeld$) As Variant
Dim ind%
Dim ret As Variant

ret = ""

ind% = SucheFeldInd%(0, iFeld$)
If (ind% >= 0) Then
    ret = clsDat.FeldValue(Mid$(AktData$, fStart%(ind%), fLaenge%(ind%)), fTyp$(ind%), fKonv%(ind%))
End If

Value = ret

End Function

Function Info$(iFeld$)
Dim iValue As Variant
Dim ind%
Dim ret$

ret$ = ""

ind% = SucheFeldInd%(0, iFeld$)
If (ind% >= 0) Then
    iValue = clsDat.FeldValue(Mid$(AktData$, fStart%(ind%), fLaenge%(ind%)), fTyp$(ind%), fKonv%(ind%))
    ret$ = clsDat.DecodFeld(iValue, fTyp$(ind%), fKonv%(ind%))
End If

Value$ = ret$

End Function

Public Property Get FeldBezeichnung$(iFeld As Variant)
Dim ind%
Dim ret$

ret$ = ""

If (VarType(iFeld) = vbString) Then
    ind% = SucheFeldInd%(0, iFeld)
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
Dim ret$

ret$ = ""

If (VarType(iFeld) = vbString) Then
    ind% = SucheFeldInd%(1, iFeld)
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

Public Property Get Wert(ByVal iFeld$, Optional iIndex% = -1) As Variant
Attribute Wert.VB_UserMemId = 0
Dim ind%
Dim ret As Variant

ret = ""

If (iIndex% >= 0) Then
    iFeld$ = iFeld$ + "(" + Mid$(Str$(iIndex%), 2) + ")"
End If

ind% = SucheFeldInd%(0, iFeld$)
If (ind% >= 0) Then
    ret = clsDat.FeldValue(Mid$(AktData$, fStart%(ind%), fLaenge%(ind%)), fTyp$(ind%), fKonv%(ind%))
End If

Wert = ret
End Property


