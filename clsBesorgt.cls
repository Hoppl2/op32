VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsBesorgt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Const SATZLEN = 65

Private Type StammlosErstSatz
    erstmax As Integer
    erstrest As String * 63
End Type
Private Type StammlosStruct
    pzn As String * 7
    flag As String * 1
    text As String * 36
    f1 As String * 1
    dat As String * 2
    Leer As String * 4
    f2 As String * 1
    dt As String * 2
    leer1 As String * 4
    lief As Byte
    lm As String * 4
    rest As String * 2
End Type
Dim besErst As StammlosErstSatz
Dim bes As StammlosStruct

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


Private Const DefErrModul = "clsBesorgt.cls"

Private Sub Class_Initialize()

dKurz$ = "BESORGT"
dBezeichnung$ = "Merkzettel"
dName$ = "stammlos.dat"
dZugriff$ = "RW"
dModus$ = "R"
dSatzLen% = SATZLEN 'Len(StammlosFix)

dAnzFields% = 12
ReDim fKurz$(dAnzFields%)
ReDim fBezeichnung$(dAnzFields%)
ReDim fStart%(dAnzFields%)
ReDim fLaenge%(dAnzFields%)
ReDim fTyp$(dAnzFields%)
ReDim fKonv%(dAnzFields%)


fKurz$(0) = "PZN"
fBezeichnung$(0) = "Pzn"
fStart%(0) = 1
fLaenge%(0) = 7
fTyp$(0) = "T"
fKonv%(0) = 0

fKurz$(1) = "FLAG"
fBezeichnung$(1) = "Flag"
fStart%(1) = 8
fLaenge%(1) = 1
fTyp$(1) = "T"
fKonv%(1) = 0


fKurz$(2) = "TEXT"
fBezeichnung$(2) = "Text"
fStart%(2) = 9
fLaenge%(2) = 36
fTyp$(2) = "T"
fKonv%(2) = 0


fKurz$(3) = "F1"
fBezeichnung$(3) = "F1"
fStart%(3) = 45
fLaenge%(3) = 1
fTyp$(3) = "T"
fKonv%(3) = 0

fKurz$(4) = "DAT"
fBezeichnung$(4) = "Dat"
fStart%(4) = 46
fLaenge%(4) = 2
fTyp$(4) = "T"
fKonv%(4) = 0


fKurz$(5) = "LEER"
fBezeichnung$(5) = "Leer"
fStart%(5) = 48
fLaenge%(5) = 4
fTyp$(5) = "T"
fKonv%(5) = 0

fKurz$(6) = "F2"
fBezeichnung$(6) = "F2"
fStart%(6) = 52
fLaenge%(6) = 1
fTyp$(6) = "T"
fKonv%(6) = 0

fKurz$(7) = "DT"
fBezeichnung$(7) = "Dt"
fStart%(7) = 53
fLaenge%(7) = 2
fTyp$(7) = "D"
fKonv%(7) = 10


fKurz$(8) = "LEER1"
fBezeichnung$(8) = "Leer1"
fStart%(8) = 55
fLaenge%(8) = 4
fTyp$(8) = "T"
fKonv%(8) = 0

fKurz$(9) = "LIEF"
fBezeichnung$(9) = "Lief"
fStart%(9) = 59
fLaenge%(9) = 1
fTyp$(9) = "N"
fKonv%(9) = 1

fKurz$(10) = "LM"
fBezeichnung$(10) = "Lm"
fStart%(10) = 60
fLaenge%(10) = 4
fTyp$(10) = "N"
fKonv%(10) = 3

fKurz$(11) = "MAX"
fBezeichnung$(11) = "Max"
fStart%(11) = 1
fLaenge%(11) = 2
fTyp$(11) = "N"
fKonv%(11) = 5


iAnzIndex% = 2
iName$(0) = "besorgt.ix1"
iName$(1) = "besorgt.ix2"



Set Besorgt1 = Me

End Sub

Sub FreeClass()
Set Besorgt1 = Nothing
End Sub

Function OpenDatei%(Optional DateiZugriff$ = "")
Dim i%
Dim iZugriff$

iZugriff$ = RTrim$(DateiZugriff$)
If (iZugriff$ = "") Then
    iZugriff$ = RTrim$(dZugriff$)
End If

If (RTrim$(dName$) <> "") Then
    dHandle% = clsDat.FileOpen(RTrim$(dName$), iZugriff$, dModus$, dSatzLen%)
    If (dHandle% = 0) Then
        OpenDatei% = 1
        Exit Function
    End If
End If
For i% = 0 To (iAnzIndex% - 1)
     FabsErrf% = clsDat.IndexOpen(iName$(i%), iHandle%(i%))
     If (FabsErrf%) Then
        OpenDatei% = 1
        Exit Function
    End If
Next i%

OpenDatei% = 0

End Function

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
    Get #dHandle%, , besErst
Else
    Get #dHandle%, , bes
End If
End Sub

Sub PutRecord(Optional Recno& = 0)
If (Recno& > 0) Then
    Seek #dHandle%, Recno&
End If
If (Recno& = 1) Then
    Put #dHandle%, , besErst
Else
    Put #dHandle%, , bes
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
erstmax% = besErst.erstmax
End Property
Public Property Let erstmax(ByVal vNewValue%)
besErst.erstmax = vNewValue%
End Property

Public Property Get erstrest$()
erstrest$ = besErst.erstrest
End Property
Public Property Let erstrest(ByVal vNewValue$)
besErst.erstrest = vNewValue$
End Property




Public Property Get pzn$()
pzn$ = bes.pzn
End Property
Public Property Let pzn(ByVal vNewValue$)
bes.pzn = vNewValue$
End Property

Public Property Get flag$()
flag$ = bes.flag
End Property
Public Property Let flag(ByVal vNewValue$)
bes.flag = vNewValue$
End Property

Public Property Get text$()
text$ = bes.text
End Property
Public Property Let text(ByVal vNewValue$)
bes.text = vNewValue$
End Property

Public Property Get f1$()
f1$ = bes.f1
End Property
Public Property Let f1(ByVal vNewValue$)
bes.f1 = vNewValue$
End Property

Public Property Get dat%()
dat% = clsOpTool.CVDat(bes.dat)
End Property
Public Property Let dat(ByVal vNewValue%)
bes.dat = clsOpTool.MKDate(vNewValue%)
End Property

Public Property Get Leer$()
Leer$ = bes.Leer
End Property
Public Property Let Leer(ByVal vNewValue$)
bes.Leer = vNewValue$
End Property

Public Property Get f2$()
f2$ = bes.f2
End Property
Public Property Let f2(ByVal vNewValue$)
bes.f2 = vNewValue$
End Property

Public Property Get dt%()
dt% = clsOpTool.CVDat(bes.dt)
End Property
Public Property Let dt(ByVal vNewValue%)
bes.dt = clsOpTool.MKDate(vNewValue%)
End Property

Public Property Get leer1$()
leer1$ = bes.leer1
End Property
Public Property Let leer1(ByVal vNewValue$)
bes.leer1 = vNewValue$
End Property

Public Property Get lief() As Byte
lief = bes.lief
End Property
Public Property Let lief(ByVal vNewValue As Byte)
bes.lief = vNewValue
End Property

Public Property Get lm%()
lm% = Val(bes.lm)
End Property
Public Property Let lm(ByVal vNewValue%)
bes.lm = Right$(Space$(4) + Str$(vNewValue%), 4)
End Property

Public Property Get rest$()
rest$ = bes.rest
End Property
Public Property Let rest(ByVal vNewValue$)
bes.rest = vNewValue$
End Property


