VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsWÜ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Const SATZLEN = 128

Private Type BestelErstSatz
    Max As Integer
    rest As String * 126
End Type
Private Type BestelStruct
    at As String * 44
    AEP As Double   'String * 8
    kz As String * 1
    bd As String * 6
    li As String * 1
    bm As Integer
    nm As Integer
    stat As String * 1
    ac As String * 1
    rm As Integer
    lm As Integer
    am As Integer
    as As Integer
    ss As Integer
    abl As String * 6
    la As Integer
    zeit As String * 2  'Integer
    AVP As Double   'String * 8
    km As Integer
    wg As String * 1
    besorger As String * 1
    nnart As Byte   'String * 1
    NNAep As Double     'String * 8
    ldat As Integer
    beleg As String * 10
    AbholNr As Integer
    dummy As String * 7
End Type
Dim wuErst As BestelErstSatz
Dim wu As BestelStruct


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

Private Const DefErrModul = "clsWUE.cls"

Private Sub Class_Initialize()

dKurz$ = "BESTEL"
dBezeichnung$ = "Warenübernahme"
dName$ = "bestel.dat"
dZugriff$ = "RW"
dModus$ = "R"
dSatzLen% = SATZLEN 'Len(best)

dAnzFields% = 28
ReDim fKurz$(dAnzFields%)
ReDim fBezeichnung$(dAnzFields%)
ReDim fStart%(dAnzFields%)
ReDim fLaenge%(dAnzFields%)
ReDim fTyp$(dAnzFields%)
ReDim fKonv%(dAnzFields%)


fKurz$(0) = "AT"
fBezeichnung$(0) = "At"
fStart%(0) = 1
fLaenge%(0) = 44
fTyp$(0) = "T"
fKonv%(0) = 0

fKurz$(1) = "AEP"
fBezeichnung$(1) = "Aep"
fStart%(1) = 45
fLaenge%(1) = 8
fTyp$(1) = "N"
fKonv%(1) = 2

fKurz$(2) = "KZ"
fBezeichnung$(2) = "Kz"
fStart%(2) = 53
fLaenge%(2) = 1
fTyp$(2) = "T"
fKonv%(2) = 0

fKurz$(3) = "BD"
fBezeichnung$(3) = "Bd"
fStart%(3) = 54
fLaenge%(3) = 6
fTyp$(3) = "T"
fKonv%(3) = 0

fKurz$(4) = "LI"
fBezeichnung$(4) = "Li"
fStart%(4) = 60
fLaenge%(4) = 1
fTyp$(4) = "T"
fKonv%(4) = 0

fKurz$(5) = "BM"
fBezeichnung$(5) = "Bm"
fStart%(5) = 61
fLaenge%(5) = 2
fTyp$(5) = "N"
fKonv%(5) = 5

fKurz$(6) = "NM"
fBezeichnung$(6) = "Nm"
fStart%(6) = 63
fLaenge%(6) = 2
fTyp$(6) = "N"
fKonv%(6) = 5

fKurz$(7) = "STAT"
fBezeichnung$(7) = "Stat"
fStart%(7) = 65
fLaenge%(7) = 1
fTyp$(7) = "T"
fKonv%(7) = 0

fKurz$(8) = "AC"
fBezeichnung$(8) = "Ac"
fStart%(8) = 66
fLaenge%(8) = 1
fTyp$(8) = "T"
fKonv%(8) = 0

fKurz$(9) = "RM"
fBezeichnung$(9) = "Rm"
fStart%(9) = 67
fLaenge%(9) = 2
fTyp$(9) = "N"
fKonv%(9) = 5

fKurz$(10) = "LM"
fBezeichnung$(10) = "Lm"
fStart%(10) = 69
fLaenge%(10) = 2
fTyp$(10) = "N"
fKonv%(10) = 5

fKurz$(11) = "AM"
fBezeichnung$(11) = "Am"
fStart%(11) = 71
fLaenge%(11) = 2
fTyp$(11) = "N"
fKonv%(11) = 5

fKurz$(12) = "AS"
fBezeichnung$(12) = "As"
fStart%(12) = 73
fLaenge%(12) = 2
fTyp$(12) = "N"
fKonv%(12) = 5

fKurz$(13) = "SS"
fBezeichnung$(13) = "Ss"
fStart%(13) = 75
fLaenge%(13) = 2
fTyp$(13) = "N"
fKonv%(13) = 5

fKurz$(14) = "ABL"
fBezeichnung$(14) = "Abl"
fStart%(14) = 77
fLaenge%(14) = 6
fTyp$(14) = "T"
fKonv%(14) = 0

fKurz$(15) = "LA"
fBezeichnung$(15) = "La"
fStart%(15) = 83
fLaenge%(15) = 2
fTyp$(15) = "N"
fKonv%(15) = 5

fKurz$(16) = "ZEIT"
fBezeichnung$(16) = "Zeit"
fStart%(16) = 85
fLaenge%(16) = 2
fTyp$(16) = "T"
fKonv%(16) = 0

fKurz$(17) = "AVP"
fBezeichnung$(17) = "Avp"
fStart%(17) = 87
fLaenge%(17) = 8
fTyp$(17) = "N"
fKonv%(17) = 2

fKurz$(18) = "KM"
fBezeichnung$(18) = "Km"
fStart%(18) = 95
fLaenge%(18) = 2
fTyp$(18) = "N"
fKonv%(18) = 5

fKurz$(19) = "WG"
fBezeichnung$(19) = "Wg"
fStart%(19) = 97
fLaenge%(19) = 1
fTyp$(19) = "T"
fKonv%(19) = 0

fKurz$(20) = "BESORGER"
fBezeichnung$(20) = "Besorger"
fStart%(20) = 98
fLaenge%(20) = 1
fTyp$(20) = "T"
fKonv%(20) = 0

fKurz$(21) = "NNART"
fBezeichnung$(21) = "NNArt"
fStart%(21) = 99
fLaenge%(21) = 1
fTyp$(21) = "N"
fKonv%(21) = 1

fKurz$(22) = "NNAEP"
fBezeichnung$(22) = "NNAep"
fStart%(22) = 100
fLaenge%(22) = 8
fTyp$(22) = "N"
fKonv%(22) = 2

fKurz$(23) = "LDAT"
fBezeichnung$(23) = "Ldat"
fStart%(23) = 108
fLaenge%(23) = 2
fTyp$(23) = "N"
fKonv%(23) = 5

fKurz$(24) = "BELEG"
fBezeichnung$(24) = "Beleg"
fStart%(24) = 110
fLaenge%(24) = 10
fTyp$(24) = "T"
fKonv%(24) = 0

fKurz$(25) = "ABHOLNR"
fBezeichnung$(25) = "AbholNr"
fStart%(25) = 120
fLaenge%(25) = 2
fTyp$(25) = "N"
fKonv%(25) = 5

fKurz$(26) = "REST"
fBezeichnung$(26) = "Rest"
fStart%(26) = 122
fLaenge%(26) = 7
fTyp$(26) = "T"
fKonv%(26) = 0

fKurz$(27) = "MAX"
fBezeichnung$(27) = "Max"
fStart%(27) = 1
fLaenge%(27) = 2
fTyp$(27) = "N"
fKonv%(27) = 5


iAnzIndex% = 0


Set Wu1 = Me

End Sub

Sub FreeClass()
Set Wu1 = Nothing
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
    Get #dHandle%, , wuErst
Else
    Get #dHandle%, , wu
End If
End Sub

Sub PutRecord(Optional Recno& = 0)
If (Recno& > 0) Then
    Seek #dHandle%, Recno&
End If
If (Recno& = 1) Then
    Put #dHandle%, , wuErst
Else
    Put #dHandle%, , wu
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
erstmax% = wuErst.Max
End Property
Public Property Let erstmax(ByVal vNewValue%)
wuErst.Max = vNewValue%
End Property

Public Property Get erstrest$()
erstrest$ = wuErst.rest
End Property
Public Property Let erstrest(ByVal vNewValue$)
wuErst.rest = vNewValue$
End Property




Public Property Get at$()
at$ = wu.at
End Property
Public Property Let at(ByVal vNewValue$)
wu.at = vNewValue$
End Property

Public Property Get AEP#()
Dim a#
a# = wu.AEP
Call DxToIEEEd(a#)
AEP# = a#
End Property
Public Property Let AEP(ByVal vNewValue#)
Dim a#
a# = vNewValue#
Call DxToMBFd(a#)
wu.AEP = a#
End Property

Public Property Get kz$()
kz$ = wu.kz
End Property
Public Property Let kz(ByVal vNewValue$)
wu.kz = vNewValue$
End Property

Public Property Get bd$()
bd$ = wu.bd
End Property
Public Property Let bd(ByVal vNewValue$)
wu.bd = vNewValue$
End Property

Public Property Get li$()
li$ = wu.li
End Property
Public Property Let li(ByVal vNewValue$)
wu.li = vNewValue$
End Property

Public Property Get bm%()
bm% = wu.bm
End Property
Public Property Let bm(ByVal vNewValue%)
wu.bm = vNewValue%
End Property

Public Property Get nm%()
nm% = wu.nm
End Property
Public Property Let nm(ByVal vNewValue%)
wu.nm = vNewValue%
End Property

Public Property Get stat$()
stat$ = wu.stat
End Property
Public Property Let stat(ByVal vNewValue$)
wu.stat = vNewValue$
End Property

Public Property Get ac$()
ac$ = wu.ac
End Property
Public Property Let ac(ByVal vNewValue$)
wu.ac = vNewValue$
End Property

Public Property Get rm%()
rm% = wu.rm
End Property
Public Property Let rm(ByVal vNewValue%)
wu.rm = vNewValue%
End Property

Public Property Get lm%()
lm% = wu.lm
End Property
Public Property Let lm(ByVal vNewValue%)
wu.lm = vNewValue%
End Property

Public Property Get am%()
am% = wu.am
End Property
Public Property Let am(ByVal vNewValue%)
wu.am = vNewValue%
End Property

Public Property Get asatz%()
asatz% = wu.as
End Property
Public Property Let asatz(ByVal vNewValue%)
wu.as = vNewValue%
End Property

Public Property Get ssatz%()
ssatz% = wu.ss
End Property
Public Property Let ssatz(ByVal vNewValue%)
wu.ss = vNewValue%
End Property

Public Property Get abl$()
abl$ = wu.abl
End Property
Public Property Let abl(ByVal vNewValue$)
wu.abl = vNewValue$
End Property

Public Property Get la%()
la% = wu.la
End Property
Public Property Let la(ByVal vNewValue%)
wu.la = vNewValue%
End Property

Public Property Get zeit$()
zeit$ = wu.zeit
End Property
Public Property Let zeit(ByVal vNewValue$)
wu.zeit = vNewValue$
End Property

Public Property Get AVP#()
Dim a#
a# = wu.AVP
Call DxToIEEEd(a#)
AVP# = a#
End Property
Public Property Let AVP(ByVal vNewValue#)
Dim a#
a# = vNewValue#
Call DxToMBFd(a#)
wu.AVP = a#
End Property

Public Property Get km%()
km% = wu.km
End Property
Public Property Let km(ByVal vNewValue%)
wu.km = vNewValue%
End Property

Public Property Get wg$()
wg$ = wu.wg
End Property
Public Property Let wg(ByVal vNewValue$)
wu.wg = vNewValue$
End Property

Public Property Get besorger$()
besorger$ = wu.besorger
End Property
Public Property Let besorger(ByVal vNewValue$)
wu.besorger = vNewValue$
End Property

Public Property Get nnart() As Byte
nnart = wu.nnart
End Property
Public Property Let nnart(ByVal vNewValue As Byte)
wu.nnart = vNewValue
End Property

Public Property Get NNAep#()
Dim a#
a# = wu.NNAep
Call DxToIEEEd(a#)
NNAep# = a#
End Property
Public Property Let NNAep(ByVal vNewValue#)
Dim a#
a# = vNewValue#
Call DxToMBFd(a#)
wu.NNAep = a#
End Property

Public Property Get ldat%()
ldat% = wu.ldat
End Property
Public Property Let ldat(ByVal vNewValue%)
wu.ldat = vNewValue%
End Property

Public Property Get beleg$()
beleg$ = wu.beleg
End Property
Public Property Let beleg(ByVal vNewValue$)
wu.beleg = vNewValue$
End Property

Public Property Get AbholNr%()
AbholNr% = wu.AbholNr
End Property
Public Property Let AbholNr(ByVal vNewValue%)
wu.AbholNr = vNewValue%
End Property

Public Property Get rest$()
rest$ = wu.dummy
End Property
Public Property Let rest(ByVal vNewValue$)
wu.dummy = vNewValue$
End Property


