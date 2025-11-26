VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsKunBezug"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Const SATZLEN = 192


Private Type BezugErstSatz
    erstmax As Single
    erstrest As String * 188
End Type
Private Type BezugStruct
    RezNr       As Integer
    EndeT       As Byte
    pzn         As String * 7
    text        As String * 36
    Preis       As String * 9
    mw          As String * 1
    bs          As Byte
    datum       As String * 2
    user        As Integer
    LaufNr      As Integer
    gebuehren   As Byte
    wegmark     As String * 1
    AVP         As Byte
    zeit        As Integer
    bon         As String * 1
    wg          As String * 2
    knr         As Integer
    gebsumme    As String * 8
    pflag       As String * 1
    PersCode    As Byte
    angefordert As Byte
    FremdGeld   As Byte
    FremdBetrag As String * 8
    TxtTyp      As String * 1
    RezEan      As String * 13
    KKNr        As String * 9
    KKTyp       As Byte
    ZuzaFlag    As String * 1
    Zustellung  As Byte
    RezeptArt   As Byte
    Storno      As Byte
    rest        As String * 6
    multi       As Byte
    KnrIdx      As String * 5
    DatIdx      As String * 8
    BelegIdx    As String * 10
    LfdIdx      As String * 3
    FaktAnz     As String * 4
    ZahlKnr     As String * 2
    AbgleichKZ  As String * 1
    PhCare      As Byte
    rest2       As String * 30
End Type
Dim bezErst As BezugErstSatz
Dim bez As BezugStruct

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



Private Const DefErrModul = "KUNBEZUG.CLS"

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
dKurz$ = "BEZ"
dBezeichnung$ = "Kunden-Bezüge"
dName$ = "KUNBEZUG.DAT"
dZugriff$ = "R"
dModus$ = "R"
dSatzLen% = SATZLEN

dAnzFields% = 1
ReDim fKurz$(dAnzFields% - 1)
ReDim fBezeichnung$(dAnzFields% - 1)
ReDim fStart%(dAnzFields% - 1)
ReDim fLaenge%(dAnzFields% - 1)
ReDim fTyp$(dAnzFields% - 1)
ReDim fKonv%(dAnzFields% - 1)

fKurz$(0) = "PZN"
fBezeichnung$(0) = "Pharmazentralnummer"
fStart%(0) = 1
fLaenge%(0) = 7
fTyp$(0) = "T"
fKonv%(0) = 0


iAnzIndex% = 1
iName$(0) = "KUNBEZUG.IDX"


Set bez1 = Me

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
Set bez1 = Nothing
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

'If (Para1.Land = "A") Then
'    iName$(1) = "artikel3.idx"
'End If

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

Sub GetFirstRecord()
Dim LockTime As Date
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("GetFirstRecord")
Call clsError.DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:

If Err = 70 Or Err = 75 Then
  If LockTime = 0 Then LockTime = DateAdd("s", 20, Now)
  If LockTime > Now Then
    'Sleep (1)
    Resume
  End If
End If

Select Case clsError.DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  LockTime = 0
  Resume
Case vbIgnore
  Resume Next
End Select
Call ProjektForm.EndeDll
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

Get #dHandle%, 1, bezErst

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
If (recno& = 1) Then
'    Get #dHandle%, , assErst
    Call GetFirstRecord
Else
    Get #dHandle%, , bez
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
If (recno& = 1) Then
    Put #dHandle%, , bezErst
Else
    Put #dHandle%, , bez
End If
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

Public Property Get DateiName$()
DateiName$ = dName$
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
Attribute Wert.VB_UserMemId = 0
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





Public Property Get erstmax!()
Dim a!
a! = bezErst.erstmax
Call DxToIEEEs(a!)
erstmax! = a!
End Property
Public Property Let erstmax(ByVal vNewValue!)
Dim a!
a! = vNewValue!
Call DxToMBFs(a!)
bezErst.erstmax = a!
End Property

Public Property Get erstrest$()
erstrest$ = bezErst.erstrest
End Property
Public Property Let erstrest(ByVal vNewValue$)
bezErst.erstrest = vNewValue$
End Property



Public Property Get RezNr%()
RezNr% = bez.RezNr
End Property
Public Property Let RezNr(ByVal vNewValue%)
bez.RezNr = vNewValue%
End Property

Public Property Get EndeT() As Byte
EndeT = bez.EndeT
End Property
Public Property Let EndeT(ByVal vNewValue As Byte)
bez.EndeT = vNewValue
End Property

Public Property Get pzn$()
pzn$ = bez.pzn
End Property
Public Property Let pzn(ByVal vNewValue$)
bez.pzn = vNewValue$
End Property

Public Property Get text$()
text$ = bez.text
End Property
Public Property Let text(ByVal vNewValue$)
bez.text = vNewValue$
End Property

Public Property Get Preis$()
Preis$ = bez.Preis
End Property
Public Property Let Preis(ByVal vNewValue$)
bez.Preis = vNewValue$
End Property

Public Property Get mw$()
mw$ = bez.mw
End Property
Public Property Let mw(ByVal vNewValue$)
bez.mw = vNewValue$
End Property

Public Property Get bs() As Byte
bs = bez.bs
End Property
Public Property Let bs(ByVal vNewValue As Byte)
bez.bs = vNewValue
End Property

Public Property Get datum%()
datum% = clsOpTool.CVDat(bez.datum)
End Property
Public Property Let datum(ByVal vNewValue%)
bez.datum = clsOpTool.MKDate(vNewValue%)
End Property

Public Property Get user%()
user% = bez.user
End Property
Public Property Let user(ByVal vNewValue%)
bez.user = vNewValue%
End Property

Public Property Get LaufNr%()
LaufNr% = bez.LaufNr
End Property
Public Property Let LaufNr(ByVal vNewValue%)
bez.LaufNr = vNewValue%
End Property

Public Property Get gebuehren() As Byte
gebuehren = bez.gebuehren
End Property
Public Property Let gebuehren(ByVal vNewValue As Byte)
bez.gebuehren = vNewValue
End Property

Public Property Get wegmark$()
wegmark$ = bez.wegmark
End Property
Public Property Let wegmark(ByVal vNewValue$)
bez.wegmark = vNewValue$
End Property

Public Property Get AVP() As Byte
AVP = bez.AVP
End Property
Public Property Let AVP(ByVal vNewValue As Byte)
bez.AVP = vNewValue
End Property

Public Property Get zeit%()
zeit% = bez.zeit
End Property
Public Property Let zeit(ByVal vNewValue%)
bez.zeit = vNewValue%
End Property

Public Property Get bon$()
bon$ = bez.bon
End Property
Public Property Let bon(ByVal vNewValue$)
bez.bon = vNewValue$
End Property

Public Property Get wg$()
wg$ = bez.wg
End Property
Public Property Let wg(ByVal vNewValue$)
bez.wg = vNewValue$
End Property

Public Property Get knr%()
knr% = bez.knr
End Property
Public Property Let knr(ByVal vNewValue%)
bez.knr = vNewValue%
End Property

Public Property Get gebsumme#()
Dim a#
Call CopyMemory(a#, ByVal bez.gebsumme, 8)
Call DxToIEEEd(a#)
gebsumme# = a#
End Property
Public Property Let gebsumme(ByVal vNewValue#)
Dim a#
a# = vNewValue#
Call DxToMBFd(a#)
Call CopyMemory(ByVal bez.gebsumme, a#, 8)
End Property

Public Property Get pflag$()
pflag$ = bez.pflag
End Property
Public Property Let pflag(ByVal vNewValue$)
bez.pflag = vNewValue$
End Property

Public Property Get PersCode() As Byte
PersCode = bez.PersCode
End Property
Public Property Let PersCode(ByVal vNewValue As Byte)
bez.PersCode = vNewValue
End Property

Public Property Get angefordert() As Byte
angefordert = bez.angefordert
End Property
Public Property Let angefordert(ByVal vNewValue As Byte)
bez.angefordert = vNewValue
End Property

Public Property Get FremdGeld() As Byte
FremdGeld = bez.FremdGeld
End Property
Public Property Let FremdGeld(ByVal vNewValue As Byte)
bez.FremdGeld = vNewValue
End Property

Public Property Get FremdBetrag#()
Dim a#
Call CopyMemory(a#, ByVal bez.FremdBetrag, 8)
Call DxToIEEEd(a#)
FremdBetrag# = a#
End Property
Public Property Let FremdBetrag(ByVal vNewValue#)
Dim a#
a# = vNewValue#
Call DxToMBFd(a#)
Call CopyMemory(ByVal bez.FremdBetrag, a#, 8)
End Property

Public Property Get TxtTyp$()
TxtTyp$ = bez.TxtTyp
End Property
Public Property Let TxtTyp(ByVal vNewValue$)
bez.TxtTyp = vNewValue$
End Property

Public Property Get RezEan$()
RezEan$ = bez.RezEan
End Property
Public Property Let RezEan(ByVal vNewValue$)
bez.RezEan = vNewValue$
End Property

Public Property Get KKNr$()
KKNr$ = bez.KKNr
End Property
Public Property Let KKNr(ByVal vNewValue$)
bez.KKNr = vNewValue$
End Property

Public Property Get KKTyp() As Byte
KKTyp = bez.KKTyp
End Property
Public Property Let KKTyp(ByVal vNewValue As Byte)
bez.KKTyp = vNewValue
End Property

Public Property Get ZuzaFlag$()
ZuzaFlag$ = bez.ZuzaFlag
End Property
Public Property Let ZuzaFlag(ByVal vNewValue$)
bez.ZuzaFlag = vNewValue$
End Property

Public Property Get Zustellung() As Byte
Zustellung = bez.Zustellung
End Property
Public Property Let Zustellung(ByVal vNewValue As Byte)
bez.Zustellung = vNewValue
End Property

Public Property Get RezeptArt() As Byte
RezeptArt = bez.RezeptArt
End Property
Public Property Let RezeptArt(ByVal vNewValue As Byte)
bez.RezeptArt = vNewValue
End Property

Public Property Get Storno() As Byte
Storno = bez.Storno
End Property
Public Property Let Storno(ByVal vNewValue As Byte)
bez.Storno = vNewValue
End Property


Public Property Get rest$()
rest$ = bez.rest
End Property
Public Property Let rest(ByVal vNewValue$)
bez.rest = vNewValue$
End Property

Public Property Get multi() As Byte
multi = bez.multi
End Property
Public Property Let multi(ByVal vNewValue As Byte)
bez.multi = vNewValue
End Property


Public Property Get KnrIdx$()
KnrIdx$ = bez.KnrIdx
End Property
Public Property Let KnrIdx(ByVal vNewValue$)
bez.KnrIdx = vNewValue$
End Property

Public Property Get DatIdx$()
DatIdx$ = bez.DatIdx
End Property
Public Property Let DatIdx(ByVal vNewValue$)
bez.DatIdx = vNewValue$
End Property

Public Property Get BelegIdx$()
BelegIdx$ = bez.BelegIdx
End Property
Public Property Let BelegIdx(ByVal vNewValue$)
bez.BelegIdx = vNewValue$
End Property

Public Property Get LfdIdx$()
LfdIdx$ = bez.LfdIdx
End Property
Public Property Let LfdIdx(ByVal vNewValue$)
bez.LfdIdx = vNewValue$
End Property

Public Property Get FaktAnz$()
FaktAnz$ = bez.FaktAnz
End Property
Public Property Let FaktAnz(ByVal vNewValue$)
bez.FaktAnz = vNewValue$
End Property

Public Property Get ZahlKnr$()
ZahlKnr$ = bez.ZahlKnr
End Property
Public Property Let ZahlKnr(ByVal vNewValue$)
bez.ZahlKnr = vNewValue$
End Property

Public Property Get AbgleichKZ$()
AbgleichKZ$ = bez.AbgleichKZ
End Property
Public Property Let AbgleichKZ(ByVal vNewValue$)
bez.AbgleichKZ = vNewValue$
End Property

Public Property Get PhCare() As Byte
PhCare = bez.PhCare
End Property
Public Property Let PhCare(ByVal vNewValue As Byte)
bez.PhCare = vNewValue
End Property

Public Property Get rest2$()
rest2$ = bez.rest2
End Property
Public Property Let rest2(ByVal vNewValue$)
bez.rest2 = vNewValue$
End Property




Public Property Get FabsRecno&()
FabsRecno& = iFabsRecno&
End Property

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
Put #Handle%, , bez
Call clsError.DefErrPop
End Sub


