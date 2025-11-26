VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsHilfsTaxe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Const SATZLEN = 128

Private Type HilfsTaxErstSatz
    erstmax     As Integer
    erstrest    As String * 126
End Type
Private Type HilfsTaxStruct
    pzn         As String * 7
    kurz        As String * 28
    meng        As String * 5
    meh         As String * 2
    aep         As Double
    kp          As Double
    AVP         As Double
    mw          As String * 1
    ArzneiBuch  As String * 2
    MinAns      As String * 1
'    blank       As String * 1
    Gstufe      As String * 8
    wg          As String * 1
    vc          As String * 1
    vDatum      As String * 4
    lichts      As String * 1
    SepVen      As String * 1
    cave        As String * 1
    aufbew      As String * 2
    sucht       As String * 1
    vet         As String * 1
    rez         As String * 2
    gal         As String * 1
    Status      As String * 1
    lager1      As String * 1
    lager2      As String * 1
    alkohol     As String * 1
    TextSatz    As Single   'String * 4
    Abgrenz     As String * 2
    kas         As String * 3
    Gruppe      As String * 1
    rest        As String * 20
End Type

Dim hTax As HilfsTaxStruct
Dim hTaxErst As HilfsTaxErstSatz

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

Dim AnzRec&

Dim iPzn8$, iKurzNeu$, iMengNeu$

Private Const DefErrModul = "HILFSTAX.CLS"

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
dKurz$ = "HTAXE"
dBezeichnung$ = "HilfsTaxe"
dName$ = "ew-taxe.dat"
dZugriff$ = "R"
dModus$ = "R"
dSatzLen% = SATZLEN

dAnzFields% = 8
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

fKurz$(1) = "KURZ"
fBezeichnung$(1) = "Bezeichnung"
fStart%(1) = 8
fLaenge%(1) = 28
fTyp$(1) = "T"
fKonv%(1) = 0

fKurz$(2) = "MENG"
fBezeichnung$(2) = "Menge"
fStart%(2) = 36
fLaenge%(2) = 5
fTyp$(2) = "T"
fKonv%(2) = 0

fKurz$(3) = "MEH"
fBezeichnung$(3) = "Einheit"
fStart%(3) = 41
fLaenge%(3) = 2
fTyp$(3) = "T"
fKonv%(3) = 0

fKurz$(4) = "AEP"
fBezeichnung$(4) = "Einkaufspreis"
fStart%(4) = 43
fLaenge%(4) = 8
fTyp$(4) = "N"
fKonv%(4) = 2

fKurz$(5) = "KP"
fBezeichnung$(5) = "Kassenpreis"
fStart%(5) = 51
fLaenge%(5) = 8
fTyp$(5) = "N"
fKonv%(5) = 2

fKurz$(6) = "AVP"
fBezeichnung$(6) = "Verkaufspreis"
fStart%(6) = 59
fLaenge%(6) = 8
fTyp$(6) = "N"
fKonv%(6) = 2

fKurz$(7) = "GSTUFE"
fBezeichnung$(7) = "GewichtsStufe"
fStart%(7) = 71
fLaenge%(7) = 8
fTyp$(7) = "T"
fKonv%(7) = 0


iAnzIndex% = 2
iName$(0) = "ew-taxe1.idx"
iName$(1) = "ew-taxe2.idx"

Set hTax1 = Me

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
Set hTax1 = Nothing
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
Dim i%, j%, iFabsErrf%
Dim iZugriff$

iZugriff$ = RTrim$(DateiZugriff$)
If (iZugriff$ = "") Then
    iZugriff$ = RTrim$(dZugriff$)
End If

For j% = 0 To 1
    If (RTrim$(dName$) <> "") Then
        dHandle% = clsDat.FileOpen(RTrim$(dName$), iZugriff$, dModus$, dSatzLen%)
        If (dHandle% = 0) Then
            OpenDatei% = 1
            Call clsError.DefErrPop: Exit Function
        End If
    End If
    For i% = 0 To (iAnzIndex% - 1)
            iFabsErrf% = clsDat.IndexOpen(iName$(i%), iHandle%(i%), True)
        If (iFabsErrf% <> 0) Then
            If (j% = 0) Then
                Call Reorg(False, DateiZugriff$)
            End If
            Exit For
        End If
    Next i%
    If (iFabsErrf% = 0) Then
        Exit For
    End If
Next j%

Call CheckReorg

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

Sub DeleteRecord(recno&, Optional SollPzn$ = "0000000", Optional MitSpeichern% = True)
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
CopyMemory hTax, ByVal temp$, SATZLEN
hTax.pzn = SollPzn$
hTax.kurz = "A" + Space$(Len(hTax.kurz) - 1)
hTax.meng = Space$(Len(hTax.meng))
hTax.meh = Space$(Len(hTax.meh))

If (MitSpeichern%) Then
    Call PutRecord(recno&)
End If

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

Get #dHandle%, 1, hTaxErst

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
    Get #dHandle%, , hTax
    iPzn8 = "0" + hTax.pzn
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
    Put #dHandle%, , hTaxErst
Else
    Put #dHandle%, , hTax
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

Function IndexCreate%(ind%, iKeyLen%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("IndexCreate%")
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

FabsErrf% = clsDat.IndexCreate(iName$(ind%), iHandle%(ind%), iKeyLen%)
IndexCreate% = FabsErrf%

Call clsError.DefErrPop
End Function

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

Function IndexFirst%(ind%, recno&)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("IndexFirst%")
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
FabsErrf% = clsDat.IndexFirst%(iHandle%(ind%), recno&)
iFabsRecno& = recno&
IndexFirst% = FabsErrf%
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

Function IndexReplace%(ind%, AltRecno&, AltKey$, key$, recno&)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("IndexReplace%")
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
FabsErrf% = clsDat.IndexReplace%(iHandle%(ind%), AltRecno&, AltKey$, key$, recno&)
iFabsRecno& = recno&
IndexReplace% = FabsErrf%
Call clsError.DefErrPop
End Function

Function IndexCountRecords%(ind%, recno&)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("IndexCountRecords%")
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

FabsErrf% = clsDat.IndexCountRecords%(iHandle%(ind%), recno&)
iFabsRecno& = recno&
IndexCountRecords% = FabsErrf%

Call clsError.DefErrPop
End Function

Function IndexCountDeleted%(ind%, recno&)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("IndexCountDeleted%")
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

FabsErrf% = clsDat.IndexCountDeleted%(iHandle%(ind%), recno&)
iFabsRecno& = recno&
IndexCountDeleted% = FabsErrf%

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

Public Property Get RecordLen%()
RecordLen% = dSatzLen%
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


Public Property Get erstmax%()
erstmax% = hTaxErst.erstmax
End Property
Public Property Let erstmax(ByVal vNewValue%)
hTaxErst.erstmax = vNewValue%
End Property

Public Property Get erstrest$()
erstrest$ = hTaxErst.erstrest
End Property
Public Property Let erstrest(ByVal vNewValue$)
hTaxErst.erstrest = vNewValue$
End Property


Public Property Get pzn$()
'pzn$ = hTax.pzn
pzn = iPzn8
End Property
Public Property Let pzn(ByVal vNewValue$)
hTax.pzn = Right(vNewValue$, 7)
iPzn8 = vNewValue
End Property

Public Property Get kurz$()
Dim h$
h$ = hTax.kurz
Call OemToChar(h$, h$)
kurz$ = h$
End Property
Public Property Let kurz(ByVal vNewValue$)
Dim h$
h$ = vNewValue$
Call CharToOem(h$, h$)
hTax.kurz = h$
End Property

Public Property Get KurzNeu$()
Dim h$
h$ = iKurzNeu
Call OemToChar(h$, h$)
KurzNeu$ = h$
End Property
Public Property Let KurzNeu(ByVal vNewValue$)
Dim h$
h$ = vNewValue$
Call CharToOem(h$, h$)
iKurzNeu = h$
End Property

Public Property Get meng$()
meng$ = hTax.meng
End Property
Public Property Let meng(ByVal vNewValue$)
hTax.meng = vNewValue$
End Property

Public Property Get MengNeu$()
MengNeu$ = iMengNeu
End Property
Public Property Let MengNeu(ByVal vNewValue$)
iMengNeu = vNewValue$
End Property

Public Property Get meh$()
meh$ = hTax.meh
End Property
Public Property Let meh(ByVal vNewValue$)
hTax.meh = vNewValue$
End Property

Public Property Get aep#()
Dim a#
a# = hTax.aep
Call DxToIEEEd(a#)
aep# = a#
End Property
Public Property Let aep(ByVal vNewValue#)
Dim a#
a# = vNewValue#
Call DxToMBFd(a#)
hTax.aep = a#
End Property

Public Property Get kp#()
Dim a#
a# = hTax.kp
Call DxToIEEEd(a#)
kp# = a#
End Property
Public Property Let kp(ByVal vNewValue#)
Dim a#
a# = vNewValue#
Call DxToMBFd(a#)
hTax.kp = a#
End Property

Public Property Get AVP#()
Dim a#
a# = hTax.AVP
Call DxToIEEEd(a#)
AVP# = a#
End Property
Public Property Let AVP(ByVal vNewValue#)
Dim a#
a# = vNewValue#
Call DxToMBFd(a#)
hTax.AVP = a#
End Property

Public Property Get mw$()
mw$ = hTax.mw
End Property
Public Property Let mw(ByVal vNewValue$)
hTax.mw = vNewValue$
End Property

Public Property Get Gstufe$()
Gstufe$ = hTax.Gstufe
End Property
Public Property Let Gstufe(ByVal vNewValue$)
hTax.Gstufe = vNewValue$
End Property

Public Property Get MinAns$()
MinAns$ = hTax.MinAns
End Property
Public Property Let MinAns(ByVal vNewValue$)
hTax.MinAns = vNewValue$
End Property

Public Property Get vc$()
vc$ = hTax.vc
End Property
Public Property Let vc(ByVal vNewValue$)
hTax.vc = vNewValue$
End Property

Public Property Get vDatum$()
vDatum$ = hTax.vDatum
End Property
Public Property Let vDatum(ByVal vNewValue$)
hTax.vDatum = vNewValue$
End Property

Public Property Get rez$()
rez$ = hTax.rez
End Property
Public Property Let rez(ByVal vNewValue$)
hTax.rez = vNewValue$
End Property

Public Property Get kas$()
kas$ = hTax.kas
End Property
Public Property Let kas(ByVal vNewValue$)
hTax.kas = vNewValue$
End Property

Public Property Get wg$()
wg$ = hTax.wg
End Property
Public Property Let wg(ByVal vNewValue$)
hTax.wg = vNewValue$
End Property

Public Property Get ArzneiBuch$()
ArzneiBuch$ = hTax.ArzneiBuch
End Property
Public Property Let ArzneiBuch(ByVal vNewValue$)
hTax.ArzneiBuch = vNewValue$
End Property

Public Property Get LichtSchutz$()
LichtSchutz$ = hTax.lichts
End Property
Public Property Let LichtSchutz(ByVal vNewValue$)
hTax.lichts = vNewValue$
End Property

Public Property Get SepVen$()
SepVen$ = hTax.SepVen
End Property
Public Property Let SepVen(ByVal vNewValue$)
hTax.SepVen = vNewValue$
End Property

Public Property Get cave$()
cave$ = hTax.cave
End Property
Public Property Let cave(ByVal vNewValue$)
hTax.cave = vNewValue$
End Property

Public Property Get aufbew$()
aufbew$ = hTax.aufbew
End Property
Public Property Let aufbew(ByVal vNewValue$)
hTax.aufbew = vNewValue$
End Property

Public Property Get sucht$()
sucht$ = hTax.sucht
End Property
Public Property Let sucht(ByVal vNewValue$)
hTax.sucht = vNewValue$
End Property

Public Property Get vet$()
vet$ = hTax.vet
End Property
Public Property Let vet(ByVal vNewValue$)
hTax.vet = vNewValue$
End Property

Public Property Get gal$()
gal$ = hTax.gal
End Property
Public Property Let gal(ByVal vNewValue$)
hTax.gal = vNewValue$
End Property

Public Property Get Status$()
Status$ = hTax.Status
End Property
Public Property Let Status(ByVal vNewValue$)
hTax.Status = vNewValue$
End Property

Public Property Get lager1$()
lager1$ = hTax.lager1
End Property
Public Property Let lager1(ByVal vNewValue$)
hTax.lager1 = vNewValue$
End Property

Public Property Get lager2$()
lager2$ = hTax.lager2
End Property
Public Property Let lager2(ByVal vNewValue$)
hTax.lager2 = vNewValue$
End Property

Public Property Get alkohol$()
alkohol$ = hTax.alkohol
End Property
Public Property Let alkohol(ByVal vNewValue$)
hTax.alkohol = vNewValue$
End Property

Public Property Get Abgrenz$()
Abgrenz$ = hTax.Abgrenz
End Property
Public Property Let Abgrenz(ByVal vNewValue$)
hTax.Abgrenz = vNewValue$
End Property

Public Property Get Gruppe$()
Gruppe$ = hTax.Gruppe
End Property
Public Property Let Gruppe(ByVal vNewValue$)
hTax.Gruppe = vNewValue$
End Property

Public Property Get TextSatz!()
Dim a!
a! = hTax.TextSatz
Call DxToIEEEs(a!)
TextSatz! = a!
End Property
Public Property Let TextSatz(ByVal vNewValue!)
Dim a!
a! = vNewValue!
Call DxToMBFs(a!)
hTax.TextSatz = a!
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
Put #Handle%, , hTax
Call clsError.DefErrPop
End Sub


Function Reorg%(Optional MitOpen% = True, Optional OpenPara$ = "")
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("Reorg%")
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
Dim i%, frei%, ret%
Dim SollAnzRec&
Dim h$

ret% = True

If (dHandle% > 0) Then
    For i% = 0 To (iAnzIndex% - 1)
        FabsErrf% = clsDat.IndexClose(iHandle%(i%))
    Next i%
End If

Call GetFirstRecord
'Anzrec& = astErst.erstmax
AnzRec& = (DateiLen& / SATZLEN) - 1

ret% = ReorgStammIdx%
If (ret% And MitOpen%) Then
    Call OpenDatei(OpenPara$)
End If

Reorg% = ret%

Call clsError.DefErrPop
End Function

Function ReorgStammIdx%()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("ReorgStammIdx%")
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
Dim j%, k%, ret%, defekt%
Dim NEU_ID%
Dim i&, NeuAnz&
Dim sKurz$, NeuKurz$, sMatch$, NeuMatch$, hPzn$, SuchTxt$, GespeichertTxt$, AltKey$
Dim MatchHilfsTaxe As clsMatchHilfsTaxe

Load frmReorg
With frmReorg
    .Caption = "Hilfstaxe reorganisieren"
    .lbReorg.Caption = "Indizes für Hilfstaxe ..."
    .MaxAnz (AnzRec&)
    .Show
End With

On Error Resume Next
Kill "EW-TAXEN.$$$"
Kill iName$(0)
Kill iName$(1)
On Error GoTo DefErr

Set MatchHilfsTaxe = New clsMatchHilfsTaxe
Call MatchHilfsTaxe.InitMatch

FabsErrf% = IndexCreate%(0, 7)
If (FabsErrf% = 0) Then
    FabsErrf% = IndexCreate%(1, 10)
End If

If (FabsErrf% = 0) Then
    NEU_ID% = clsDat.FileOpen("EW-TAXEN.$$$", "RW", "R", SATZLEN)
    If (NEU_ID% = 0) Then
        FabsErrf% = 8
    End If
End If

NeuAnz& = 0
If (FabsErrf% = 0) Then
    For i& = 1 To AnzRec&
        Call GetRecord(i& + 1)
    
        defekt% = 0
        hPzn$ = Trim(hTax.pzn)
        If (hPzn$ <> "") Then
            For j = 1 To 7
                If (InStr("0123456789", Mid$(hTax.pzn, j%, 1)) = 0) Then
                    defekt% = True
                    Exit For
                End If
            Next j%
            If (Val(hTax.pzn) = 0) Then
                defekt% = True
            End If
        End If
        If (defekt% = 0) Then
            sKurz$ = Trim(hTax.kurz + hTax.meng + hTax.meh)
            If (sKurz$ = "") Then
                Call DeleteRecord(i& + 1, hTax.pzn, False)
                defekt% = True
            End If
        End If
        If (defekt% = 0) And (hPzn$ <> "") Then
            FabsErrf% = IndexSearch%(0, hTax.pzn, iFabsRecno&)
            If (FabsErrf% = 0) Then
                defekt = True
            End If
        End If
        If (defekt% = 0) Then
            sKurz$ = hTax.kurz
            Call OemToChar(sKurz$, sKurz$)
            NeuKurz$ = MatchHilfsTaxe.MakeMatchCode(sKurz$)
            Call CharToOem(NeuKurz$, NeuKurz$)
            
            FabsErrf% = IndexSearch%(1, NeuKurz$, iFabsRecno&)
            If (FabsErrf% = 0) Then
                SuchTxt$ = Trim(hTax.kurz + hTax.meng + hTax.meh)
                Do
                    If (FabsErrf% > 0) Then
                        Exit Do
                    End If
                    
                    AltKey$ = clsDat.GetLastKey(10)
                    If (AltKey$ <> NeuKurz$) Then
                        Exit Do
                    End If
                        
                    Get #NEU_ID%, iFabsRecno& + 1, hTax
                    GespeichertTxt$ = Trim(hTax.kurz + hTax.meng + hTax.meh)
                    
                    If (SuchTxt$ = GespeichertTxt$) Then
                        defekt% = True
                        Exit Do
                    Else
                        FabsErrf% = IndexNext%(1, iFabsRecno&, AltKey$, iFabsRecno&)
                    End If
                Loop
                If (defekt% = 0) Then
                    Call GetRecord(i& + 1)
                End If
            End If
        End If
        If (defekt% = 0) Then
            FabsErrf% = IndexInsert%(0, hTax.pzn, iFabsRecno&)
            If (FabsErrf% <> 0) Then
                Exit For
            End If
        
            FabsErrf% = IndexInsert%(1, NeuKurz$, iFabsRecno&)
            If (FabsErrf% <> 0) Then
                Exit For
            End If
            
            NeuAnz& = NeuAnz& + 1
            Put #NEU_ID%, NeuAnz& + 1, hTax
        End If
        
        If (i& Mod 100 = 0) Then
            frmReorg.ShowFortschritt (i&)
        End If
    Next i&

    FabsErrf% = clsDat.IndexClose(iHandle%(0))
    FabsErrf% = clsDat.IndexClose(iHandle%(1))
End If

frmReorg.ShowFortschritt (i&)
Unload frmReorg

With hTaxErst
    .erstmax = NeuAnz&
    .erstrest = String(Len(.erstrest), 0)
End With
Put #NEU_ID%, 1, hTaxErst

Close #dHandle%
Close #NEU_ID%

FileCopy dName$, "ew-taxe.oat"
FileCopy "EW-TAXEN.$$$", dName$

ReorgStammIdx% = (FabsErrf% = 0)

Call clsError.DefErrPop
End Function

Sub CheckReorg()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("CheckReorg")
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
Dim Anz1&, Anz2&, SollAnz&, AnzFrei&, DelAnz1&, DelAnz2&
Dim s$

FabsErrf% = IndexCountRecords(0, iFabsRecno&)
Anz1& = iFabsRecno&
FabsErrf% = IndexCountDeleted(0, iFabsRecno&)
DelAnz1& = iFabsRecno&
'FabsErrf% = IndexCountDeleted(0, iFabsRecno&)
'Anz1& = Anz1& - iFabsRecno&

FabsErrf% = IndexCountRecords(1, iFabsRecno&)
Anz2& = iFabsRecno&
FabsErrf% = IndexCountDeleted(1, iFabsRecno&)
DelAnz1& = iFabsRecno&
'FabsErrf% = IndexCountDeleted(1, iFabsRecno&)
'Anz2& = Anz2& - iFabsRecno&

Call GetFirstRecord
AnzRec& = hTaxErst.erstmax
SollAnz& = (DateiLen& / SATZLEN) - 1

s$ = "Hilfstaxe CheckReorg: "
s$ = s$ + Right$(Space$(8) + Format(Anz1&, "0"), 8)
s$ = s$ + Right$(Space$(8) + Format(Anz2&, "0"), 8)
s$ = s$ + Right$(Space$(8) + Format(AnzRec&, "0"), 8)
s$ = s$ + Right$(Space$(8) + Format(SollAnz&, "0"), 8)
s$ = s$ + Right$(Space$(8) + Format(DelAnz1&, "0"), 8)
s$ = s$ + Right$(Space$(8) + Format(DelAnz2&, "0"), 8)
Call clsDialog.WinArtDebug(s$)

If (SollAnz& <> AnzRec&) Or (Anz1& <> AnzRec&) Or (Anz2& <> AnzRec&) Then
'If (Anz1& <> SollAnz&) Or (Anz2& <> SollAnz&) Then
    AnzRec& = SollAnz&
    s$ = "Achtung: Reorganisieren der Hilfstaxe ist notwendig!"
    Call clsDialog.MessageBox(s$, vbInformation)
    Call clsDialog.WinArtDebug(s$)
End If

If (AnzRec& > 0) Then
    FabsErrf% = IndexCountDeleted(0, iFabsRecno&)
    AnzFrei& = iFabsRecno&
    If ((AnzFrei& / AnzRec& * 100) > 10) Then
        s$ = "Achtung: Reorganisieren (Komprimieren) der Hilfstaxe ist notwendig!"
        Call clsDialog.MessageBox(s$, vbInformation)
        Call clsDialog.WinArtDebug(s$)
'        Call MsgBox("Achtung: Komprimieren der Hilfstaxe ist notwendig!", vbInformation)
    End If
End If

Call clsError.DefErrPop
End Sub


