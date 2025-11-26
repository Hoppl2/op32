VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsEtiketten"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Const SATZLEN = 64

Private Type EtikettenErstSatz
    Max As Integer
    Ges As Integer
    rest As String * 60
End Type
Private Type EtikettenStruct
    pzn As String * 7
    Menge As Integer
    druck As String * 1
    asatz As Integer
    zusatz As String * 2
    knr As String * 2
    lief As Integer     'string * 2
    Datum As String * 2
    rest As String * 44
End Type
Dim etikErst As EtikettenErstSatz
Dim etik As EtikettenStruct


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


Private Const DefErrModul = "clsEtiketten.cls"

Private Sub Class_Initialize()

dKurz$ = "POSDRUCK"
dBezeichnung$ = "Etikettendatei"
dName$ = "posdruck.dat"
dZugriff$ = "RW"
dModus$ = "R"
dSatzLen% = SATZLEN 'Len(etik)

dAnzFields% = 11
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

fKurz$(1) = "MENGE"
fBezeichnung$(1) = "Menge"
fStart%(1) = 8
fLaenge%(1) = 2
fTyp$(1) = "N"
fKonv%(1) = 5

fKurz$(2) = "DRUCK"
fBezeichnung$(2) = "Druck"
fStart%(2) = 10
fLaenge%(2) = 1
fTyp$(2) = "T"
fKonv%(2) = 0

fKurz$(3) = "ASATZ"
fBezeichnung$(3) = "Asatz"
fStart%(3) = 11
fLaenge%(3) = 2
fTyp$(3) = "N"
fKonv%(3) = 5

fKurz$(4) = "ZUSATZ"
fBezeichnung$(4) = "Zusatz"
fStart%(4) = 13
fLaenge%(4) = 2
fTyp$(4) = "T"
fKonv%(4) = 0

fKurz$(5) = "KNR"
fBezeichnung$(5) = "Knr"
fStart%(5) = 15
fLaenge%(5) = 2
fTyp$(5) = "T"
fKonv%(5) = 0

fKurz$(6) = "LIEF"
fBezeichnung$(6) = "Lief"
fStart%(6) = 17
fLaenge%(6) = 2
fTyp$(6) = "N"
fKonv%(6) = 5

fKurz$(7) = "DATUM"
fBezeichnung$(7) = "Datum"
fStart%(7) = 19
fLaenge%(7) = 2
fTyp$(7) = "T"
fKonv%(7) = 0

fKurz$(8) = "REST"
fBezeichnung$(8) = "Rest"
fStart%(8) = 21
fLaenge%(8) = 44
fTyp$(8) = "T"
fKonv%(8) = 0

fKurz$(9) = "MAX"
fBezeichnung$(9) = "Max"
fStart%(9) = 1
fLaenge%(9) = 2
fTyp$(9) = "N"
fKonv%(9) = 5

fKurz$(10) = "GES"
fBezeichnung$(10) = "Ges"
fStart%(10) = 3
fLaenge%(10) = 2
fTyp$(10) = "N"
fKonv%(10) = 5


iAnzIndex% = 0

    
Set Etik1 = Me

End Sub

Sub FreeClass()
Set Etik1 = Nothing
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
If (Recno& = 1) Then
    Get #dHandle%, , etikErst
Else
    Get #dHandle%, , etik
End If
End Sub

Sub PutRecord(Optional Recno& = 0)
If (Recno& > 0) Then
    Seek #dHandle%, Recno&
End If
If (Recno& = 1) Then
    Put #dHandle%, , etikErst
Else
    Put #dHandle%, , etik
End If
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

Public Property Get erstmax%()
erstmax% = etikErst.Max
End Property
Public Property Let erstmax(ByVal vNewValue%)
etikErst.Max = vNewValue%
End Property

Public Property Get erstges%()
erstges% = etikErst.Ges
End Property
Public Property Let erstges(ByVal vNewValue%)
etikErst.Ges = vNewValue%
End Property

Public Property Get erstrest$()
erstrest$ = etikErst.rest
End Property
Public Property Let erstrest(ByVal vNewValue$)
etikErst.rest = vNewValue$
End Property



Public Property Get pzn$()
pzn$ = etik.pzn
End Property
Public Property Let pzn(ByVal vNewValue$)
etik.pzn = vNewValue$
End Property

Public Property Get Menge%()
Menge% = etik.Menge
End Property
Public Property Let Menge(ByVal vNewValue%)
etik.Menge = vNewValue%
End Property

Public Property Get druck$()
druck$ = etik.druck
End Property
Public Property Let druck(ByVal vNewValue$)
etik.druck = vNewValue$
End Property

Public Property Get asatz%()
asatz% = etik.asatz
End Property
Public Property Let asatz(ByVal vNewValue%)
etik.asatz = vNewValue%
End Property

Public Property Get zusatz$()
zusatz$ = etik.zusatz
End Property
Public Property Let zusatz(ByVal vNewValue$)
etik.zusatz = vNewValue$
End Property

Public Property Get knr$()
knr$ = etik.knr
End Property
Public Property Let knr(ByVal vNewValue$)
etik.knr = vNewValue$
End Property

Public Property Get lief%()
lief% = etik.lief
End Property
Public Property Let lief(ByVal vNewValue%)
etik.lief = vNewValue%
End Property

Public Property Get Datum%()
Datum% = clsOpTool.CVDat(etik.Datum)
End Property
Public Property Let Datum(ByVal vNewValue%)
etik.Datum = clsOpTool.MKDate(vNewValue%)
End Property

Public Property Get rest$()
rest$ = etik.rest
End Property
Public Property Let rest(ByVal vNewValue$)
etik.rest = vNewValue$
End Property



Public Property Get FabsRecno&()
FabsRecno& = iFabsRecno&
End Property



