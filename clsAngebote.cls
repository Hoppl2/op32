VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsAngebote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Const SATZLEN = 34

Private Type GhAngebote1
    bm As Byte
    mp As Long
End Type
Private Type GhAngeboteStruct
    pzn As String * 7
    ang(4) As GhAngebote1
    gh As Byte
    st As String * 1
End Type
Dim GhAnbot As GhAngeboteStruct

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


Private Const DefErrModul = "clsAngebote.cls"


Private Sub Class_Initialize()

dKurz$ = "ANGEBOTE"
dBezeichnung$ = "Gh-Angebote"
dName$ = "gh-anbot.dat"
dZugriff$ = "R"
dModus$ = "R"
dSatzLen% = SATZLEN 'Len(ghanbot)

dAnzFields% = 13
ReDim fKurz$(dAnzFields% - 1)
ReDim fBezeichnung$(dAnzFields% - 1)
ReDim fStart%(dAnzFields% - 1)
ReDim fLaenge%(dAnzFields% - 1)
ReDim fTyp$(dAnzFields% - 1)
ReDim fKonv%(dAnzFields% - 1)

fKurz$(0) = "PZN"
fBezeichnung$(0) = "Pzn"
fStart%(0) = 1
fLaenge%(0) = 7
fTyp$(0) = "T"
fKonv%(0) = 0

fKurz$(1) = "BM(0)"
fBezeichnung$(1) = "Bm(0)"
fStart%(1) = 8
fLaenge%(1) = 1
fTyp$(1) = "N"
fKonv%(1) = 1

fKurz$(2) = "MP(0)"
fBezeichnung$(2) = "Menge/Preis(0)"
fStart%(2) = 9
fLaenge%(2) = 4
fTyp$(2) = "N"
fKonv%(2) = 7

fKurz$(3) = "BM(1)"
fBezeichnung$(3) = "Bm(1)"
fStart%(3) = 13
fLaenge%(3) = 1
fTyp$(3) = "N"
fKonv%(3) = 1

fKurz$(4) = "MP(1)"
fBezeichnung$(4) = "Menge/Preis(1)"
fStart%(4) = 14
fLaenge%(4) = 4
fTyp$(4) = "N"
fKonv%(4) = 7

fKurz$(5) = "BM(2)"
fBezeichnung$(5) = "Bm(2)"
fStart%(5) = 18
fLaenge%(5) = 1
fTyp$(5) = "N"
fKonv%(5) = 1

fKurz$(6) = "MP(2)"
fBezeichnung$(6) = "Menge/Preis(2)"
fStart%(6) = 19
fLaenge%(6) = 4
fTyp$(6) = "N"
fKonv%(6) = 7

fKurz$(7) = "BM(3)"
fBezeichnung$(7) = "Bm(3)"
fStart%(7) = 23
fLaenge%(7) = 1
fTyp$(7) = "N"
fKonv%(7) = 1

fKurz$(8) = "MP(3)"
fBezeichnung$(8) = "Menge/Preis(3)"
fStart%(8) = 24
fLaenge%(8) = 4
fTyp$(8) = "N"
fKonv%(8) = 7

fKurz$(9) = "BM(4)"
fBezeichnung$(9) = "Bm(4)"
fStart%(9) = 28
fLaenge%(9) = 1
fTyp$(9) = "N"
fKonv%(9) = 1

fKurz$(10) = "MP(4)"
fBezeichnung$(10) = "Menge/Preis(4)"
fStart%(10) = 29
fLaenge%(10) = 4
fTyp$(10) = "N"
fKonv%(10) = 7

fKurz$(11) = "GH"
fBezeichnung$(11) = "Gh"
fStart%(11) = 33
fLaenge%(11) = 1
fTyp$(11) = "N"
fKonv%(11) = 1

fKurz$(12) = "ST"
fBezeichnung$(12) = "St"
fStart%(12) = 34
fLaenge%(12) = 1
fTyp$(12) = "T"
fKonv%(12) = 0


iAnzIndex% = 0

    
Set clsAngebote1 = Me

End Sub

Sub FreeClass()
Set clsAngebote1 = Nothing
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

Sub CloseDatei()
Dim i%

If (dHandle% > 0) Then
    Close #dHandle%
End If
For i% = 0 To (iAnzIndex% - 1)
     FabsErrf% = clsDat.IndexClose(iHandle%(i%))
Next i%

End Sub

Sub SatzLock(Recno&)
Call clsDat.iLock(dHandle%, Recno&)
End Sub

Sub SatzUnLock(Recno&)
Call clsDat.iUnLock(dHandle%, Recno&)
End Sub

Sub GetRecord(Optional Recno& = 0)
If (Recno& > 0) Then
    Seek #dHandle%, Recno&
End If
Get #dHandle%, , GhAnbot
End Sub

Sub PutRecord(Optional Recno& = 0)
If (Recno& > 0) Then
    Seek #dHandle%, Recno&
End If
Put #dHandle%, , GhAnbot
End Sub

Sub GetInfoRecord(Optional Recno& = 0)
If (Recno& > 0) Then
    Seek #dHandle%, Recno&
End If
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

Public Property Let wert(iFeld$, Optional iIndex% = -1, vNewValue As Variant)
Dim ind%
Dim s$, ret$

ret$ = ""

s$ = UCase$(iFeld$)

If (iIndex% >= 0) Then
    s$ = s$ + "(" + Mid$(Str$(iIndex%), 2) + ")"
End If

ind% = SucheFeldInd%(0, s$)
If (ind% >= 0) Then
    ret$ = clsDat.MakeFeld(vNewValue, fTyp$(ind%), fKonv%(ind%))
    Mid$(AktData$, fStart%(ind%), fLaenge%(ind%)) = Left$(ret$ + Space$(fLaenge%(ind%)), fLaenge%(ind%))
End If

End Property



Public Property Get pzn$()
pzn$ = GhAnbot.pzn
End Property
Public Property Let pzn(ByVal vNewValue$)
GhAnbot.pzn = vNewValue$
End Property

Public Property Get bm(ind%) As Byte
bm = GhAnbot.ang(ind%).bm
End Property
Public Property Let bm(ind%, ByVal vNewValue As Byte)
GhAnbot.ang(ind%).bm = vNewValue
End Property

Public Property Get mp&(ind%)
mp& = GhAnbot.ang(ind%).mp
End Property
Public Property Let mp(ind%, ByVal vNewValue&)
GhAnbot.ang(ind%).mp = vNewValue&
End Property

Public Property Get gh() As Byte
gh = GhAnbot.gh
End Property
Public Property Let gh(ByVal vNewValue As Byte)
GhAnbot.gh = vNewValue
End Property

Public Property Get st$()
st$ = GhAnbot.st
End Property
Public Property Let st(ByVal vNewValue$)
GhAnbot.st = vNewValue$
End Property

