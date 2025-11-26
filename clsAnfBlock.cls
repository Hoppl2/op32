VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsAnfBlock"
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

Dim AktData As String * 32


Private Const DefErrModul = "clsAnfBlock.cls"


Private Sub Class_Initialize()

dKurz$ = "ANFBLOCK"
dBezeichnung$ = "AnfBlock"
dName$ = "anfblock.dat"
dZugriff$ = "R"
dModus$ = "R"
dSatzLen% = 32 'Len(anfwas)

dAnzFields% = 11
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

fKurz$(1) = "WASTUN"
fBezeichnung$(1) = "WasTun"
fStart%(1) = 2
fLaenge%(1) = 1
fTyp$(1) = "T"
fKonv%(1) = 0

fKurz$(2) = "REZEPTNR"
fBezeichnung$(2) = "RezeptNr"
fStart%(2) = 3
fLaenge%(2) = 2
fTyp$(2) = "N"
fKonv%(2) = 5

fKurz$(3) = "PREIS"
fBezeichnung$(3) = "Preis"
fStart%(3) = 5
fLaenge%(3) = 9
fTyp$(3) = "T"
fKonv%(3) = 0

fKurz$(4) = "ERSTMAGSATZ"
fBezeichnung$(4) = "ErstMagSatz"
fStart%(4) = 14
fLaenge%(4) = 4
fTyp$(4) = "N"
fKonv%(4) = 7

fKurz$(5) = "ACONTO"
fBezeichnung$(5) = "Aconto"
fStart%(5) = 18
fLaenge%(5) = 9
fTyp$(5) = "T"
fKonv%(5) = 0

fKurz$(6) = "REZEAN"
fBezeichnung$(6) = "RezEan"
fStart%(6) = 27
fLaenge%(6) = 6
fTyp$(6) = "T"
fKonv%(6) = 0

fKurz$(7) = "PROTECT"
fBezeichnung$(7) = "Protect"
fStart%(7) = 1
fLaenge%(7) = 1
fTyp$(7) = "N"
fKonv%(7) = 1

fKurz$(8) = "INFOZEILE"
fBezeichnung$(8) = "InfoZeile"
fStart%(8) = 2
fLaenge%(8) = 24
fTyp$(8) = "T"
fKonv%(8) = 0

fKurz$(9) = "REZEPTART"
fBezeichnung$(9) = "RezeptArt"
fStart%(9) = 26
fLaenge%(9) = 1
fTyp$(9) = "N"
fKonv%(9) = 1

fKurz$(10) = "REST"
fBezeichnung$(10) = "Rest"
fStart%(10) = 27
fLaenge%(10) = 6
fTyp$(10) = "T"
fKonv%(10) = 0



iAnzIndex% = 0


Set clsAnfBlock1 = Me

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



