VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsNachbearbeitung"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Const SATZLEN = 200

Private Type NachbeErstSatz
    Max As Integer
    rest As String * 198
End Type
Private Type NachbeStruct
    at As String * 44
    AEP As String * 8
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
    txt As String * 57
    user As String * 10
    alm As Integer
    arm As Integer
    fertig As String * 1
    AVP As String * 8
    km As Integer
    wg As String * 1
    dummy As String * 33
End Type
Dim nbErst As NachbeErstSatz
Dim nb As NachbeStruct


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


Private Const DefErrModul = "clsNachbearbeitung.cls"


Private Sub Class_Initialize()

dKurz$ = "NACHBE"
dBezeichnung$ = "Nachbearbeitung"
dName$ = "nachbe.dat"
dZugriff$ = "R"
dModus$ = "R"
dSatzLen% = SATZLEN 'Len(nachb)

dAnzFields% = 26
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

fKurz$(16) = "TXT"
fBezeichnung$(16) = "Txt"
fStart%(16) = 85
fLaenge%(16) = 57
fTyp$(16) = "T"
fKonv%(16) = 0

fKurz$(17) = "USER"
fBezeichnung$(17) = "User"
fStart%(17) = 142
fLaenge%(17) = 10
fTyp$(17) = "T"
fKonv%(17) = 0

fKurz$(18) = "ALM"
fBezeichnung$(18) = "Alm"
fStart%(18) = 152
fLaenge%(18) = 2
fTyp$(18) = "N"
fKonv%(18) = 5

fKurz$(19) = "ARM"
fBezeichnung$(19) = "Arm"
fStart%(19) = 154
fLaenge%(19) = 2
fTyp$(19) = "N"
fKonv%(19) = 5

fKurz$(20) = "FERTIG"
fBezeichnung$(20) = "Fertig"
fStart%(20) = 156
fLaenge%(20) = 1
fTyp$(20) = "T"
fKonv%(20) = 0

fKurz$(21) = "AVP"
fBezeichnung$(21) = "Avp"
fStart%(21) = 157
fLaenge%(21) = 8
fTyp$(21) = "N"
fKonv%(21) = 2

fKurz$(22) = "KM"
fBezeichnung$(22) = "Km"
fStart%(22) = 165
fLaenge%(22) = 2
fTyp$(22) = "N"
fKonv%(22) = 5

fKurz$(23) = "WG"
fBezeichnung$(23) = "Wg"
fStart%(23) = 167
fLaenge%(23) = 1
fTyp$(23) = "T"
fKonv%(23) = 0

fKurz$(24) = "REST"
fBezeichnung$(24) = "Rest"
fStart%(24) = 168
fLaenge%(24) = 33
fTyp$(24) = "T"
fKonv%(24) = 0

fKurz$(25) = "MAX"
fBezeichnung$(25) = "Max"
fStart%(25) = 1
fLaenge%(25) = 2
fTyp$(25) = "N"
fKonv%(25) = 5


iAnzIndex% = 0


Set Nb1 = Me

End Sub

Sub FreeClass()
Set Nb1 = Nothing
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
    Get #dHandle%, , nbErst
Else
    Get #dHandle%, , nb
End If
End Sub

Sub PutRecord(Optional Recno& = 0)
If (Recno& > 0) Then
    Seek #dHandle%, Recno&
End If
If (Recno& = 1) Then
    Put #dHandle%, , nbErst
Else
    Put #dHandle%, , nb
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
erstmax% = nbErst.Max
End Property
Public Property Let erstmax(ByVal vNewValue%)
nbErst.Max = vNewValue%
End Property

Public Property Get erstrest$()
erstrest$ = nbErst.rest
End Property
Public Property Let erstrest(ByVal vNewValue$)
nbErst.rest = vNewValue$
End Property







Public Property Get at$()
at$ = nb.at
End Property
Public Property Let at(ByVal vNewValue$)
nb.at = vNewValue$
End Property

Public Property Get AEP#()
Dim a#
a# = nb.AEP
Call DxToIEEEd(a#)
AEP# = a#
End Property
Public Property Let AEP(ByVal vNewValue#)
Dim a#
a# = vNewValue#
Call DxToMBFd(a#)
nb.AEP = a#
End Property

Public Property Get kz$()
kz$ = nb.kz
End Property
Public Property Let kz(ByVal vNewValue$)
nb.kz = vNewValue$
End Property

Public Property Get bd$()
bd$ = nb.bd
End Property
Public Property Let bd(ByVal vNewValue$)
nb.bd = vNewValue$
End Property

Public Property Get li$()
li$ = nb.li
End Property
Public Property Let li(ByVal vNewValue$)
nb.li = vNewValue$
End Property

Public Property Get bm%()
bm% = nb.bm
End Property
Public Property Let bm(ByVal vNewValue%)
nb.bm = vNewValue%
End Property

Public Property Get nm%()
nm% = nb.nm
End Property
Public Property Let nm(ByVal vNewValue%)
nb.nm = vNewValue%
End Property

Public Property Get stat$()
stat$ = nb.stat
End Property
Public Property Let stat(ByVal vNewValue$)
nb.stat = vNewValue$
End Property

Public Property Get ac$()
ac$ = nb.ac
End Property
Public Property Let ac(ByVal vNewValue$)
nb.ac = vNewValue$
End Property

Public Property Get rm%()
rm% = nb.rm
End Property
Public Property Let rm(ByVal vNewValue%)
nb.rm = vNewValue%
End Property

Public Property Get lm%()
lm% = nb.lm
End Property
Public Property Let lm(ByVal vNewValue%)
nb.lm = vNewValue%
End Property

Public Property Get am%()
am% = nb.am
End Property
Public Property Let am(ByVal vNewValue%)
nb.am = vNewValue%
End Property

Public Property Get asatz%()
asatz% = nb.as
End Property
Public Property Let asatz(ByVal vNewValue%)
nb.as = vNewValue%
End Property

Public Property Get ssatz%()
ssatz% = nb.ss
End Property
Public Property Let ssatz(ByVal vNewValue%)
nb.ss = vNewValue%
End Property

Public Property Get abl$()
abl$ = nb.abl
End Property
Public Property Let abl(ByVal vNewValue$)
nb.abl = vNewValue$
End Property

Public Property Get la%()
la% = nb.la
End Property
Public Property Let la(ByVal vNewValue%)
nb.la = vNewValue%
End Property

Public Property Get txt$()
txt$ = nb.txt
End Property
Public Property Let txt(ByVal vNewValue$)
nb.txt = vNewValue$
End Property

Public Property Get user$()
user$ = nb.user
End Property
Public Property Let user(ByVal vNewValue$)
nb.user = vNewValue$
End Property

Public Property Get alm%()
alm% = nb.alm
End Property
Public Property Let alm(ByVal vNewValue%)
nb.alm = vNewValue%
End Property

Public Property Get arm%()
arm% = nb.arm
End Property
Public Property Let arm(ByVal vNewValue%)
nb.arm = vNewValue%
End Property

Public Property Get fertig$()
fertig$ = nb.fertig
End Property
Public Property Let fertig(ByVal vNewValue$)
nb.fertig = vNewValue$
End Property

Public Property Get AVP#()
Dim a#
a# = nb.AVP
Call DxToIEEEd(a#)
AVP# = a#
End Property
Public Property Let AVP(ByVal vNewValue#)
Dim a#
a# = vNewValue#
Call DxToMBFd(a#)
nb.AVP = a#
End Property

Public Property Get km%()
km% = nb.km
End Property
Public Property Let km(ByVal vNewValue%)
nb.km = vNewValue%
End Property

Public Property Get wg$()
wg$ = nb.wg
End Property
Public Property Let wg(ByVal vNewValue$)
nb.wg = vNewValue$
End Property

Public Property Get rest$()
rest$ = nb.dummy
End Property
Public Property Let rest(ByVal vNewValue$)
nb.dummy = vNewValue$
End Property



