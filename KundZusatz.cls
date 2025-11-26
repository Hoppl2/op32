VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsKundZusatz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Const SATZLEN = 150

Private Type KundenZusatzStruct
    ApoPlus As String * 1
    LieferscheinOhnePreis As String * 1
    Geschlecht As Byte
    MischKunde As String * 1
    EanAlt As String * 1
    telefon As String * 30
    BefStat As String * 1
    BefBis As String * 6
    SamQ As String * 1
    uId As String * 15
    KK As String * 9
    StartMonat As Byte
    BefNr As String * 8
    HausArzt As String * 30
    Vnr As String * 12
    BLZ As String * 8
    KontoNr As String * 10
    RabattES As String * 1
    ean As String * 13
End Type
Dim KunZusatz As KundenZusatzStruct

Private Type KundenZusatzStructA
    Vnr As String * 10
    Leer As String * 4
    BLZ As String * 5
    KontoNr As String * 11
End Type
Dim KunZusatzA As KundenZusatzStructA

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

Dim iAktRecno&

Private Const DefErrModul = "KUNDZUSATZ.CLS"

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
dKurz$ = "KZUSATZ"
dBezeichnung$ = "Kunden-Zusatz"
dName$ = "kzusatz.dat"
dZugriff$ = "RW"
dModus$ = "R"
dSatzLen% = SATZLEN

dAnzFields% = 2
ReDim fKurz$(dAnzFields% - 1)
ReDim fBezeichnung$(dAnzFields% - 1)
ReDim fStart%(dAnzFields% - 1)
ReDim fLaenge%(dAnzFields% - 1)
ReDim fTyp$(dAnzFields% - 1)
ReDim fKonv%(dAnzFields% - 1)


fKurz$(0) = "MINDUMS"
fBezeichnung$(0) = "Mindest-Umsatz"
fStart%(0) = 1
fLaenge%(0) = 8
fTyp$(0) = "N"
fKonv%(0) = 2

fKurz$(1) = "REST"
fBezeichnung$(1) = "Rest"
fStart%(1) = 9
fLaenge%(1) = 120
fTyp$(1) = "T"
fKonv%(1) = 0

iAnzIndex% = 0


Set KunZus1 = Me

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

Set KunZus1 = Nothing

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

If (dHandle% = 0) Then
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
End If

For i% = 0 To (iAnzIndex% - 1)
     FabsErrf% = clsDat.IndexOpen(iName$(i%), iHandle%(i%))
     If (FabsErrf%) Then
        OpenDatei% = 1
        Call clsError.DefErrPop: Exit Function
    End If
Next i%

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
    dHandle% = 0
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
iAktRecno& = recno&
If (recno& > 0) Then
    Seek #dHandle%, recno&
End If
Get #dHandle%, , KunZusatz

If (Para1.Land = "A") Then
    Call KonvertToA
End If

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

If (Para1.Land = "A") Then
    Call KonvertToD
End If

Put #dHandle%, , KunZusatz

Call clsError.DefErrPop
End Sub

Sub DeleteRecord(Optional recno& = 0)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("DeleteRecord")
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
Dim temp$

temp$ = String(SATZLEN, 0)
CopyMemory KunZusatz, ByVal temp$, SATZLEN
KunZusatz.Vnr = Space$(Len(KunZusatz.Vnr))
KunZusatz.ean = Space$(Len(KunZusatz.ean))
KunZusatz.BefStat = "N"
KunZusatzA.Vnr = Space$(Len(KunZusatzA.Vnr))

Call PutRecord(recno&)

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

Sub InitDatei()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("InitDatei")
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
'Dim i%
'Dim LeerData As String * SATZLEN
'
'LeerData = String(SATZLEN, 0)
'For i% = 0 To 200
'    Put #dHandle%, , LeerData
'Next i%
'
Call clsError.DefErrPop
End Sub

Public Property Get FabsRecno&()
FabsRecno& = iFabsRecno&
End Property

Public Property Get ApoPlus() As Byte
ApoPlus = Abs(UCase(KunZusatz.ApoPlus) = "J")
End Property
Public Property Let ApoPlus(ByVal vNewValue As Byte)
If (vNewValue) Then
    KunZusatz.ApoPlus = "J"
Else
    KunZusatz.ApoPlus = "N"
End If
End Property

Public Property Get LieferscheinOhnePreis() As Byte
LieferscheinOhnePreis = Abs(UCase(KunZusatz.LieferscheinOhnePreis) = "J")
End Property
Public Property Let LieferscheinOhnePreis(ByVal vNewValue As Byte)
If (vNewValue) Then
    KunZusatz.LieferscheinOhnePreis = "J"
Else
    KunZusatz.LieferscheinOhnePreis = "N"
End If
End Property

Public Property Get Geschlecht() As Byte
Dim ret As Byte

ret = 0
If (KunZusatz.Geschlecht > 100) Then
    ret = KunZusatz.Geschlecht - 100
End If

Geschlecht = ret
End Property
Public Property Let Geschlecht(ByVal vNewValue As Byte)
KunZusatz.Geschlecht = vNewValue + 100
End Property

Public Property Get MischKunde() As Byte
MischKunde = Abs(UCase(KunZusatz.MischKunde) = "J")
End Property
Public Property Let MischKunde(ByVal vNewValue As Byte)
If (vNewValue) Then
    KunZusatz.MischKunde = "J"
Else
    KunZusatz.MischKunde = "N"
End If
End Property

Public Property Get EanAlt$()
EanAlt$ = KunZusatz.EanAlt
End Property
Public Property Let EanAlt(ByVal vNewValue$)
KunZusatz.EanAlt = vNewValue$
End Property

Public Property Get telefon$()
telefon$ = KunZusatz.telefon
End Property
Public Property Let telefon(ByVal vNewValue$)
KunZusatz.telefon = vNewValue$
End Property

Public Property Get BefStat() As Byte
BefStat = Abs(UCase(KunZusatz.BefStat) = "J")
End Property
Public Property Let BefStat(ByVal vNewValue As Byte)
If (vNewValue) Then
    KunZusatz.BefStat = "J"
Else
    KunZusatz.BefStat = "N"
End If
End Property

Public Property Get BefBis$()
BefBis$ = KunZusatz.BefBis
End Property
Public Property Let BefBis(ByVal vNewValue$)
KunZusatz.BefBis = vNewValue$
End Property

Public Property Get SamQ() As Byte
SamQ = Abs(UCase(KunZusatz.SamQ) = "J")
End Property
Public Property Let SamQ(ByVal vNewValue As Byte)
If (vNewValue) Then
    KunZusatz.SamQ = "J"
Else
    KunZusatz.SamQ = "N"
End If
End Property

Public Property Get uId$()
uId$ = KunZusatz.uId
End Property
Public Property Let uId(ByVal vNewValue$)
KunZusatz.uId = vNewValue$
End Property

Public Property Get KK$()
Dim h$
h$ = KunZusatz.KK
Call OemToChar(h$, h$)
KK$ = h$
End Property
Public Property Let KK(ByVal vNewValue$)
Dim h$
h$ = vNewValue$
Call CharToOem(h$, h$)
KunZusatz.KK = h$
End Property

Public Property Get StartMonat() As Byte
StartMonat = KunZusatz.StartMonat
If (StartMonat > 12) Then
    StartMonat = 0
End If
End Property
Public Property Let StartMonat(ByVal vNewValue As Byte)
KunZusatz.StartMonat = vNewValue
End Property

Public Property Get BefNr$()
BefNr$ = KunZusatz.BefNr
End Property
Public Property Let BefNr(ByVal vNewValue$)
KunZusatz.BefNr = vNewValue$
End Property

Public Property Get HausArzt$()
Dim h$
h$ = KunZusatz.HausArzt
Call OemToChar(h$, h$)
HausArzt$ = h$
End Property
Public Property Let HausArzt(ByVal vNewValue$)
Dim h$
h$ = vNewValue$
Call CharToOem(h$, h$)
KunZusatz.HausArzt = h$
End Property

Public Property Get Vnr$()
If (Para1.Land = "D") Then
    Vnr$ = KunZusatz.Vnr
Else
    Vnr$ = KunZusatzA.Vnr
End If
End Property
Public Property Let Vnr(ByVal vNewValue$)
If (Para1.Land = "D") Then
    KunZusatz.Vnr = vNewValue$
Else
    KunZusatzA.Vnr = vNewValue$
End If
End Property

Public Property Get BLZ$()
If (Para1.Land = "D") Then
    BLZ$ = KunZusatz.BLZ
Else
    BLZ$ = KunZusatzA.BLZ
End If
End Property
Public Property Let BLZ(ByVal vNewValue$)
If (Para1.Land = "D") Then
    KunZusatz.BLZ = vNewValue$
Else
    KunZusatzA.BLZ = vNewValue$
End If
End Property

Public Property Get KontoNr$()
If (Para1.Land = "D") Then
    KontoNr$ = KunZusatz.KontoNr
Else
    KontoNr$ = KunZusatzA.KontoNr
End If
End Property
Public Property Let KontoNr(ByVal vNewValue$)
If (Para1.Land = "D") Then
    KunZusatz.KontoNr = vNewValue$
Else
    KunZusatzA.KontoNr = vNewValue$
End If
End Property

Public Property Get RabattES$()
RabattES$ = KunZusatz.RabattES
End Property
Public Property Let RabattES(ByVal vNewValue$)
KunZusatz.RabattES = vNewValue$
End Property

Public Property Get ean$()
ean$ = KunZusatz.ean
End Property
Public Property Let ean(ByVal vNewValue$)
KunZusatz.ean = vNewValue$
End Property

    
    

Public Property Get AktRecno&()
AktRecno& = iAktRecno&
End Property

Public Property Get RawData$()
Dim temp$

temp$ = String(SATZLEN, 0)
CopyMemory ByVal temp$, KunZusatz, SATZLEN
RawData$ = temp$

End Property

Sub KonvertToA()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("KonvertToA")
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

With KunZusatzA
    .Vnr = Left$(KunZusatz.Vnr, 10)
    .Leer = Mid$(KunZusatz.Vnr, 11) + Left$(KunZusatz.BLZ, 2)
    .BLZ = Mid$(KunZusatz.BLZ, 3, 5)
    .KontoNr = Mid$(KunZusatz.BLZ, 8) + KunZusatz.KontoNr
End With
    
Call clsError.DefErrPop
End Sub

Sub KonvertToD()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("KonvertToD")
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

With KunZusatzA
    KunZusatz.Vnr = .Vnr + Left$(.Leer, 2)
    KunZusatz.BLZ = Mid$(.Leer, 3) + .BLZ + Left$(.KontoNr, 1)
    KunZusatz.KontoNr = Mid$(.KontoNr, 2)
End With
    
Call clsError.DefErrPop
End Sub


Sub PrintFile(Handle%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("PrintFile")
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

If (Para1.Land = "A") Then
    Call KonvertToD
End If

Put #Handle%, , KunZusatz

Call clsError.DefErrPop
End Sub

Sub InputFile(Handle%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("InputFile")
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

Get #Handle%, , KunZusatz

If (Para1.Land = "A") Then
    Call KonvertToA
End If

Call clsError.DefErrPop
End Sub



