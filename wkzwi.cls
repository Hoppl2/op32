VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsWkZwi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Const SATZLEN = 128
                    
Private Type WkZwErstSatz
    erstmax As Integer
    erstrest As String * 126
End Type
Private Type WkZwStruct
    pzn As String * 7
    datum As String * 2
    Lief As String * 9
    menge As Integer
    KontrNr As String * 11
    DruKz As String * 1
    Meng2 As String * 5
    Name As String * 2
    IdDatum As String * 2
    Preis As Double
    Pruef As String * 18
    AblaufDatum As String * 2
    rest As String * 59
End Type
Dim WkZwiErst As WkZwErstSatz
Dim WkZwi As WkZwStruct


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



Private Const DefErrModul = "WKZWI.CLS"

Private Sub Class_Initialize()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("Class_Initialize")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

dKurz$ = "WKZWI"
dBezeichnung$ = "WkZwi"
dName$ = "EW-WKZWI.DAT"
dZugriff$ = "RW"
dModus$ = "R"
dSatzLen% = SATZLEN

dAnzFields% = 1
ReDim fKurz$(dAnzFields%)
ReDim fBezeichnung$(dAnzFields%)
ReDim fStart%(dAnzFields%)
ReDim fLaenge%(dAnzFields%)
ReDim fTyp$(dAnzFields%)
ReDim fKonv%(dAnzFields%)


iAnzIndex% = 0

Call DefErrPop
End Sub

Sub FreeClass()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("FreeClass")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

Call DefErrPop
End Sub

Function OpenDatei%(Optional DateiZugriff$ = "")
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("OpenDatei%")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim i%
Dim iZugriff$

iZugriff$ = RTrim$(DateiZugriff$)
If (iZugriff$ = "") Then
    iZugriff$ = RTrim$(dZugriff$)
End If

If (RTrim$(dName$) <> "") Then
    dHandle% = FileOpen(RTrim$(dName$), iZugriff$, dModus$, dSatzLen%)
    If (dHandle% = 0) Then
        OpenDatei% = 1
        Call DefErrPop: Exit Function
    End If
End If

OpenDatei% = 0

Call DefErrPop
End Function

Sub CloseDatei()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("CloseDatei")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

If (dHandle% > 0) Then
    Close #dHandle%
End If

Call DefErrPop
End Sub

Sub SatzLock(Recno&)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("SatzLock")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

Call iLock(dHandle%, Recno&)

Call DefErrPop
End Sub

Sub SatzUnLock(Recno&)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("SatzUnLock")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

Call iUnLock(dHandle%, Recno&)

Call DefErrPop
End Sub

Sub GetFirstRecord()
Dim LockTime As Date
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("GetFirstRecord")
Call DefErrMod(DefErrModul)
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

Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  LockTime = 0
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

Get #dHandle%, 1, WkZwiErst

Call DefErrPop
End Sub

Sub GetRecord(Optional Recno& = 0)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("GetRecord")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

If (Recno& > 0) Then
    Seek #dHandle%, Recno&
End If
If (Recno& = 1) Then
    Call GetFirstRecord
Else
    Get #dHandle%, , WkZwi
End If

Call DefErrPop
End Sub

Sub PutRecord(Optional Recno& = 0)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("PutRecord")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

If (Recno& > 0) Then
    Seek #dHandle%, Recno&
End If
If (Recno& = 1) Then
    Put #dHandle%, , WkZwiErst
Else
    Put #dHandle%, , WkZwi
End If

Call DefErrPop
End Sub


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

Public Property Get RecordLen%()
RecordLen% = dSatzLen%
End Property



Public Property Get erstmax%()
erstmax% = WkZwiErst.erstmax
End Property
Public Property Let erstmax(ByVal vNewValue%)
WkZwiErst.erstmax = vNewValue%
End Property

Public Property Get erstrest$()
erstrest$ = WkZwiErst.erstrest
End Property
Public Property Let erstrest(ByVal vNewValue$)
WkZwiErst.erstrest = vNewValue$
End Property



Public Property Get pzn$()
pzn$ = WkZwi.pzn
End Property
Public Property Let pzn(ByVal vNewValue$)
WkZwi.pzn = vNewValue$
End Property

Public Property Get datum%()
datum% = CVDat(WkZwi.datum)
End Property
Public Property Let datum(ByVal vNewValue%)
WkZwi.datum = MKDate(vNewValue%)
End Property

Public Property Get Lief$()
Lief$ = WkZwi.Lief
End Property
Public Property Let Lief(ByVal vNewValue$)
WkZwi.Lief = vNewValue$
End Property

Public Property Get menge%()
menge% = WkZwi.menge
End Property
Public Property Let menge(ByVal vNewValue%)
WkZwi.menge = vNewValue%
End Property

Public Property Get KontrNr$()
KontrNr$ = WkZwi.KontrNr
End Property
Public Property Let KontrNr(ByVal vNewValue$)
WkZwi.KontrNr = vNewValue$
End Property

Public Property Get DruckKz$()
DruckKz$ = WkZwi.DruKz
End Property
Public Property Let DruckKz(ByVal vNewValue$)
WkZwi.DruKz = vNewValue$
End Property

Public Property Get Meng2$()
Meng2$ = WkZwi.Meng2
End Property
Public Property Let Meng2(ByVal vNewValue$)
WkZwi.Meng2 = vNewValue$
End Property

Public Property Get Name$()
Name$ = WkZwi.Name
End Property
Public Property Let Name(ByVal vNewValue$)
WkZwi.Name = vNewValue$
End Property

Public Property Get IdDatum%()
IdDatum% = CVDat(WkZwi.IdDatum)
End Property
Public Property Let IdDatum(ByVal vNewValue%)
WkZwi.IdDatum = MKDate(vNewValue%)
End Property

Public Property Get Preis#()
Dim a#
a# = WkZwi.Preis
Call DxToIEEEd(a#)
Preis# = a#
End Property
Public Property Let Preis(ByVal vNewValue#)
Dim a#
a# = vNewValue#
Call DxToMBFd(a#)
WkZwi.Preis = a#
End Property

Public Property Get Pruef$()
Pruef$ = WkZwi.Pruef
End Property
Public Property Let Pruef(ByVal vNewValue$)
WkZwi.Pruef = vNewValue$
End Property

Public Property Get AblaufDatum%()
AblaufDatum% = CVDat(WkZwi.AblaufDatum)
End Property
Public Property Let AblaufDatum(ByVal vNewValue%)
WkZwi.AblaufDatum = MKDate(vNewValue%)
End Property

Public Property Get rest$()
rest$ = WkZwi.rest
End Property
Public Property Let rest(ByVal vNewValue$)
WkZwi.rest = vNewValue$
End Property

