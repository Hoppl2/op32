VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsLieferanten"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Const SATZLEN = 251

Private Type LeerPznStruct
    knr As String * 3
    LeerPzn As String * 29
End Type

Private Type Lief1
    Name As String * 30
End Type
Private Type LieferantenStruct
    LiefName(3) As Lief1
    kurz As String * 6
    WarenBez As String * 16
    telefon As String * 14
    konto As String * 14
    BLZ As String * 8
    vertreter As String * 15
    KontoVorschlag As String * 6
    BankEinzug As String * 1
    Durchwahl As String * 4
    IdfApo As String * 7
    FibuKonto As String * 6
    IdfLieferant As String * 7
    dummy As String * 6
    WarenWert As Double
    direkt As String * 1
    skonto As Single
    saldo As Double
'    rest As String * 125
End Type
Dim lif As LieferantenStruct


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

Dim iLiefNamen$(260)
Dim iAnzLiefNamen%


Private Const DefErrModul = "LIEFERANTEN.CLS"

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
dKurz$ = "LIEF"
dBezeichnung$ = "Lieferanten"
dName$ = "ldatei.dat"
dZugriff$ = "RW"
dModus$ = "R"
dSatzLen% = SATZLEN 'Len(lif)

dAnzFields% = 21
ReDim fKurz$(dAnzFields% - 1)
ReDim fBezeichnung$(dAnzFields% - 1)
ReDim fStart%(dAnzFields% - 1)
ReDim fLaenge%(dAnzFields% - 1)
ReDim fTyp$(dAnzFields% - 1)
ReDim fKonv%(dAnzFields% - 1)

fKurz$(0) = "NA(0)"
fBezeichnung$(0) = "Name 1"
fStart%(0) = 1
fLaenge%(0) = 30
fTyp$(0) = "T"
fKonv%(0) = 0

fKurz$(1) = "NA(1)"
fBezeichnung$(1) = "Name 2"
fStart%(1) = 31
fLaenge%(1) = 30
fTyp$(1) = "T"
fKonv%(1) = 0

fKurz$(2) = "NA(2)"
fBezeichnung$(2) = "Name 3"
fStart%(2) = 61
fLaenge%(2) = 30
fTyp$(2) = "T"
fKonv%(2) = 0

fKurz$(3) = "NA(3)"
fBezeichnung$(3) = "Name 4"
fStart%(3) = 91
fLaenge%(3) = 30
fTyp$(3) = "T"
fKonv%(3) = 0

fKurz$(4) = "KURZ"
fBezeichnung$(4) = "Kurz-Bezeichnung"
fStart%(4) = 121
fLaenge%(4) = 6
fTyp$(4) = "T"
fKonv%(4) = 0

fKurz$(5) = "WBEZ"
fBezeichnung$(5) = "Warenbeschreibung"
fStart%(5) = 127
fLaenge%(5) = 16
fTyp$(5) = "T"
fKonv%(5) = 0

fKurz$(6) = "TEL"
fBezeichnung$(6) = "Telefon"
fStart%(6) = 143
fLaenge%(6) = 14
fTyp$(6) = "T"
fKonv%(6) = 0

fKurz$(7) = "KTO"
fBezeichnung$(7) = "Bankkonto"
fStart%(7) = 157
fLaenge%(7) = 14
fTyp$(7) = "T"
fKonv%(7) = 0

fKurz$(8) = "BLZ"
fBezeichnung$(8) = "Bankleitzahl"
fStart%(8) = 171
fLaenge%(8) = 8
fTyp$(8) = "T"
fKonv%(8) = 0

fKurz$(9) = "VERT"
fBezeichnung$(9) = "Vertreter"
fStart%(9) = 179
fLaenge%(9) = 15
fTyp$(9) = "T"
fKonv%(9) = 0

fKurz$(10) = "KTVOR"
fBezeichnung$(10) = "Konto-Vorschlag"
fStart%(10) = 194
fLaenge%(10) = 6
fTyp$(10) = "T"
fKonv%(10) = 0

fKurz$(11) = "BEIN"
fBezeichnung$(11) = "Bankeinzug"
fStart%(11) = 200
fLaenge%(11) = 1
fTyp$(11) = "T"
fKonv%(11) = 0

fKurz$(12) = "DW"
fBezeichnung$(12) = "Durchwahl"
fStart%(12) = 201
fLaenge%(12) = 4
fTyp$(12) = "T"
fKonv%(12) = 0

fKurz$(13) = "IDA"
fBezeichnung$(13) = "IDF-Apotheke"
fStart%(13) = 205
fLaenge%(13) = 7
fTyp$(13) = "T"
fKonv%(13) = 0

fKurz$(14) = "FIBU"
fBezeichnung$(14) = "FIBU-Konto"
fStart%(14) = 212
fLaenge%(14) = 6
fTyp$(14) = "T"
fKonv%(14) = 0

fKurz$(15) = "IDL"
fBezeichnung$(15) = "IDF-Lieferant"
fStart%(15) = 218
fLaenge%(15) = 7
fTyp$(15) = "T"
fKonv%(15) = 0

fKurz$(16) = "DUMMY"
fBezeichnung$(16) = "Dummy"
fStart%(16) = 225
fLaenge%(16) = 6
fTyp$(16) = "T"
fKonv%(16) = 0

fKurz$(17) = "WNWT"
fBezeichnung$(17) = "Warenwert netto"
fStart%(17) = 231
fLaenge%(17) = 8
fTyp$(17) = "N"
fKonv%(17) = 2

fKurz$(18) = "DIR"
fBezeichnung$(18) = "Direktlieferant"
fStart%(18) = 239
fLaenge%(18) = 1
fTyp$(18) = "T"
fKonv%(18) = 0

fKurz$(19) = "SK"
fBezeichnung$(19) = "Skonto"
fStart%(19) = 240
fLaenge%(19) = 4
fTyp$(19) = "N"
fKonv%(19) = 4

fKurz$(20) = "SALDO"
fBezeichnung$(20) = "Saldo"
fStart%(20) = 244
fLaenge%(20) = 8
fTyp$(20) = "N"
fKonv%(20) = 2



iAnzIndex% = 0


Set Lif1 = Me

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
Set Lif1 = Nothing
Call clsError.DefErrPop
End Sub

Function OpenDatei%(Optional DateiZugriff$ = "")
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("OpenDatei%")
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
Dim LeerData As String * SATZLEN

iZugriff$ = RTrim$(DateiZugriff$)
If (iZugriff$ = "") Then
    iZugriff$ = RTrim$(dZugriff$)
End If

If (RTrim$(dName$) <> "") Then
    dHandle% = clsDat.FileOpen(RTrim$(dName$), iZugriff$, dModus$, dSatzLen%)
    If (dHandle% = 0) Then
        OpenDatei% = 1
        Call clsError.DefErrPop: Exit Function
    End If
End If
For i% = 0 To (iAnzIndex% - 1)
     FabsErrf% = clsDat.IndexOpen(iName$(i%), iHandle%(i%))
     If (FabsErrf%) Then
        OpenDatei% = 1
        Call clsError.DefErrPop: Exit Function
    End If
Next i%

If (InStr(iZugriff$, "W") > 0) Then
    If (AnzRec < 250) Then
        LeerData = String(SATZLEN, 0)
        For i% = (AnzRec + 1) To 250
            Put #dHandle%, i% + 1, LeerData
        Next i%
    End If
End If

OpenDatei% = 0

Call clsError.DefErrPop
End Function

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
End If
For i% = 0 To (iAnzIndex% - 1)
     FabsErrf% = clsDat.IndexClose(iHandle%(i%))
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

Sub GetRecord(Optional recno& = 0)
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
If (recno& > 0) Then
    Seek #dHandle%, recno&
End If
Get #dHandle%, , lif
Call clsError.DefErrPop
End Sub

Sub PutRecord(Optional recno& = 0)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("PutRecord")
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
If (recno& > 0) Then
    Seek #dHandle%, recno&
End If
Put #dHandle%, , lif
Call clsError.DefErrPop
End Sub

Sub GetInfoRecord(Optional recno& = 0)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("GetInfoRecord")
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
If (recno& > 0) Then
    Seek #dHandle%, recno&
End If
Get #dHandle%, , AktData
Call clsError.DefErrPop
End Sub

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
'If (iIndex% >= 0) Then
'    s$ = s$ + "(" + Mid$(Str$(iIndex%), 2) + ")"
'End If


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

Public Property Get AnzRec&()
AnzRec& = (DateiLen& / SATZLEN) - 1
End Property

Public Property Get Wert(iFeld$, Optional iIndex% = -1) As Variant
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

Public Property Let Wert(iFeld$, Optional iIndex% = -1, vNewValue As Variant)
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



Public Property Get Name$(ByVal ind%)
Dim h$
h$ = lif.LiefName(ind%).Name
Call OemToChar(h$, h$)
Name$ = h$
End Property
Public Property Let Name(ByVal ind%, ByVal vNewValue$)
Dim h$
h$ = vNewValue$
Call CharToOem(h$, h$)
lif.LiefName(ind%).Name = h$
End Property

Public Property Get kurz$()
Dim h$
h$ = lif.kurz
Call OemToChar(h$, h$)
kurz$ = h$
End Property
Public Property Let kurz(ByVal vNewValue$)
lif.kurz = vNewValue$
End Property

Public Property Get WarenBez$()
WarenBez$ = lif.WarenBez
End Property
Public Property Let WarenBez(ByVal vNewValue$)
lif.WarenBez = vNewValue$
End Property

Public Property Get telefon$()
telefon$ = lif.telefon
End Property
Public Property Let telefon(ByVal vNewValue$)
lif.telefon = vNewValue$
End Property

Public Property Get konto$()
konto$ = lif.konto
End Property
Public Property Let konto(ByVal vNewValue$)
lif.konto = vNewValue$
End Property

Public Property Get BLZ$()
BLZ$ = lif.BLZ
End Property
Public Property Let BLZ(ByVal vNewValue$)
lif.BLZ = vNewValue$
End Property

Public Property Get vertreter$()
vertreter$ = lif.vertreter
End Property
Public Property Let vertreter(ByVal vNewValue$)
lif.vertreter = vNewValue$
End Property

Public Property Get KontoVorschlag$()
KontoVorschlag$ = lif.KontoVorschlag
End Property
Public Property Let KontoVorschlag(ByVal vNewValue$)
lif.KontoVorschlag = vNewValue$
End Property

Public Property Get BankEinzug$()
BankEinzug$ = lif.BankEinzug
End Property
Public Property Let BankEinzug(ByVal vNewValue$)
lif.BankEinzug = vNewValue$
End Property

Public Property Get Durchwahl$()
Durchwahl$ = lif.Durchwahl
End Property
Public Property Let Durchwahl(ByVal vNewValue$)
lif.Durchwahl = vNewValue$
End Property

Public Property Get IdfApo$()
IdfApo$ = lif.IdfApo
End Property
Public Property Let IdfApo(ByVal vNewValue$)
lif.IdfApo = vNewValue$
End Property

Public Property Get FibuKonto$()
FibuKonto$ = lif.FibuKonto
End Property
Public Property Let FibuKonto(ByVal vNewValue$)
lif.FibuKonto = vNewValue$
End Property

Public Property Get IdfLieferant$()
IdfLieferant$ = lif.IdfLieferant
End Property
Public Property Let IdfLieferant(ByVal vNewValue$)
lif.IdfLieferant = vNewValue$
End Property

Public Property Get dummy$()
dummy$ = lif.dummy
End Property
Public Property Let dummy(ByVal vNewValue$)
lif.dummy = vNewValue$
End Property

Public Property Get WarenWert#()
Dim a#
a# = lif.WarenWert
Call DxToIEEEd(a#)
WarenWert# = a#
End Property
Public Property Let WarenWert(ByVal vNewValue#)
Dim a#
a# = vNewValue#
Call DxToMBFd(a#)
lif.WarenWert = a#
End Property

Public Property Get DirektLieferant$()
DirektLieferant$ = lif.direkt
End Property
Public Property Let DirektLieferant(ByVal vNewValue$)
lif.direkt = vNewValue$
End Property

Public Property Get skonto!()
Dim a!
a! = lif.skonto
Call DxToIEEEs(a!)
skonto! = a!
End Property
Public Property Let skonto(ByVal vNewValue!)
Dim a!
a! = vNewValue!
Call DxToMBFs(a!)
lif.skonto = a!
End Property

Public Property Get saldo#()
Dim a#
a# = lif.saldo
Call DxToIEEEd(a#)
saldo# = a#
End Property
Public Property Let saldo(ByVal vNewValue#)
Dim a#
a# = vNewValue#
Call DxToMBFd(a#)
lif.saldo = a#
End Property


Public Property Get FabsRecno&()
FabsRecno& = iFabsRecno&
End Property

Function LeerPzn$(ByVal ind%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("LeerPzn$")
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
Dim i%, LLEER%
Dim ret$, s$, SuchKnr$
Dim LeerRec As LeerPznStruct

ret$ = ""

s$ = "LLEER.DAT"
If (Dir$(s$) <> "") Then
    LLEER% = clsDat.FileOpen(s$, "R", "R", 32)
    If (LLEER% > 0) Then
        SuchKnr$ = Right$(Space$(3) + Mid$(Str$(ind%), 2), 3)
        For i% = 1 To 255
            Get #LLEER%, , LeerRec
            If (EOF(LLEER%)) Then Exit For
            If (LeerRec.knr = SuchKnr$) Then
                ret$ = LeerRec.LeerPzn
                Exit For
            End If
        Next i%
        Close #LLEER%
    End If
End If

LeerPzn$ = Trim(ret$)

Call clsError.DefErrPop
End Function

Public Property Get LiefName$(ByVal ind%)
LiefName$ = iLiefNamen$(ind%)
End Property
Public Property Let LiefName(ByVal ind%, ByVal vNewValue$)
iLiefNamen$(ind%) = vNewValue$
End Property

Public Property Get AnzLiefNamen%()
AnzLiefNamen% = iAnzLiefNamen%
End Property
Public Property Let AnzLiefNamen(ByVal vNewValue%)
iAnzLiefNamen% = vNewValue%
End Property

