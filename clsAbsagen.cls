VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsAbsagen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Const SATZLEN = 100
                    
Private Type AbsagenErstSatz
    erstmax As Integer
    erstrest As String * 98
End Type
Private Type AbsagenStruct
    datum As String * 2
    lief As Byte
    pzn As String * 7
    filler As String * 1
    text As String * 36
    menge As Integer
    rest As String * 51
End Type
Dim absagenErst As AbsagenErstSatz
Dim absagen As AbsagenStruct

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


Private Const DefErrModul = "clsAbsagen.cls"

Private Sub Class_Initialize()

dKurz$ = "ABSAGEN"
dBezeichnung$ = "Absagen"
dName$ = "BABSAGE.DAT"
dZugriff$ = "RW"
dModus$ = "R"
dSatzLen% = SATZLEN 'Len(babsagenfix$)

dAnzFields% = 8
ReDim fKurz$(dAnzFields%)
ReDim fBezeichnung$(dAnzFields%)
ReDim fStart%(dAnzFields%)
ReDim fLaenge%(dAnzFields%)
ReDim fTyp$(dAnzFields%)
ReDim fKonv%(dAnzFields%)


fKurz$(0) = "DATUM"
fBezeichnung$(0) = "Datum"
fStart%(0) = 1
fLaenge%(0) = 2
fTyp$(0) = "T"
fKonv%(0) = 0

fKurz$(1) = "LIEF"
fBezeichnung$(1) = "Lief"
fStart%(1) = 3
fLaenge%(1) = 1
fTyp$(1) = "N"
fKonv%(1) = 1

fKurz$(2) = "PZN"
fBezeichnung$(2) = "Pzn"
fStart%(2) = 4
fLaenge%(2) = 7
fTyp$(2) = "T"
fKonv%(2) = 0

fKurz$(3) = "FILLER"
fBezeichnung$(3) = "Filler"
fStart%(3) = 11
fLaenge%(3) = 1
fTyp$(3) = "T"
fKonv%(3) = 0

fKurz$(4) = "TEXT"
fBezeichnung$(4) = "Text"
fStart%(4) = 12
fLaenge%(4) = 36
fTyp$(4) = "T"
fKonv%(4) = 0

fKurz$(5) = "MENGE"
fBezeichnung$(5) = "Menge"
fStart%(5) = 48
fLaenge%(5) = 2
fTyp$(5) = "N"
fKonv%(5) = 5

fKurz$(6) = "REST"
fBezeichnung$(6) = "Rest"
fStart%(6) = 50
fLaenge%(6) = 51
fTyp$(6) = "T"
fKonv%(6) = 0


fKurz$(7) = "MAX"
fBezeichnung$(7) = "Max"
fStart%(7) = 1
fLaenge%(7) = 2
fTyp$(7) = "N"
fKonv%(7) = 5


iAnzIndex% = 0



Set Absagen1 = Me

End Sub

Sub FreeClass()
Set Absagen1 = Nothing
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
    Get #dHandle%, , absagenErst
Else
    Get #dHandle%, , absagen
End If
End Sub

Sub PutRecord(Optional Recno& = 0)
If (Recno& > 0) Then
    Seek #dHandle%, Recno&
End If
If (Recno& = 1) Then
    Put #dHandle%, , absagenErst
Else
    Put #dHandle%, , absagen
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
erstmax% = absagenErst.erstmax
End Property
Public Property Let erstmax(ByVal vNewValue%)
absagenErst.erstmax = vNewValue%
End Property

Public Property Get erstrest$()
erstrest$ = absagenErst.erstrest
End Property
Public Property Let erstrest(ByVal vNewValue$)
absagenErst.erstrest = vNewValue$
End Property



Public Property Get datum%()
datum% = clsOpTool.CVDat(absagen.datum)
End Property
Public Property Let datum(ByVal vNewValue%)
absagen.datum = clsOpTool.MKDate(vNewValue%)
End Property

Public Property Get lief() As Byte
lief = absagen.lief
End Property
Public Property Let lief(ByVal vNewValue As Byte)
absagen.lief = vNewValue
End Property

Public Property Get pzn$()
pzn$ = absagen.pzn
End Property
Public Property Let pzn(ByVal vNewValue$)
absagen.pzn = vNewValue$
End Property

Public Property Get filler$()
filler$ = absagen.filler
End Property
Public Property Let filler(ByVal vNewValue$)
absagen.filler = vNewValue$
End Property

Public Property Get text$()
text$ = absagen.text
End Property
Public Property Let text(ByVal vNewValue$)
absagen.text = vNewValue$
End Property

Public Property Get menge%()
menge% = absagen.menge
End Property
Public Property Let menge(ByVal vNewValue%)
absagen.menge = vNewValue%
End Property

Public Property Get rest$()
rest$ = absagen.rest
End Property
Public Property Let rest(ByVal vNewValue$)
absagen.rest = vNewValue$
End Property

