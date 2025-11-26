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
    erstdruck As String * 1
    erstrest As String * 62
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



Private Const DefErrModul = "BESORGT.CLS"

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
Set Besorgt1 = Nothing
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

Get #dHandle%, 1, besErst

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
'    Get #dHandle%, , besErst
    Call GetFirstRecord
Else
    Get #dHandle%, , bes
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
    Put #dHandle%, , besErst
Else
    Put #dHandle%, , bes
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
Dim i%, xLinks%, xMitte%, xRechts%, besMax%
Dim SQLStr$, sName$, AnfangsBuchstabe$, Buchstabe$

If (Para1.Land = "A") Then
    IndexSearch% = 8
    recno& = 0
    iFabsRecno& = recno&
    
    sName$ = ""
    SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + clsOpTool.SqlPzn(key$)
    Set TaxeRec = TaxeDB.OpenRecordset(SQLStr$)
    If (TaxeRec.EOF = False) Then
        sName$ = Trim(TaxeRec!Name)
    End If
    
    If (sName$ <> "") Then
        Get #dHandle%, 1, besErst
        besMax% = besErst.erstmax
  
        AnfangsBuchstabe$ = Left(sName$, 2)
        Buchstabe$ = String(Len(AnfangsBuchstabe$), 0)
  
        xMitte% = 1
        xLinks% = 1
        xRechts% = besMax%
  
        While (xLinks% < xRechts%)
            xMitte% = Int((xLinks% + xRechts%) / 2)
            Get #dHandle%, xMitte% + 1, bes
            Buchstabe$ = Left$(bes.text, 2)
            If (AnfangsBuchstabe$ <= Buchstabe$) Then
                xRechts% = xMitte%
            ElseIf (AnfangsBuchstabe$ > Buchstabe$) Then
                xLinks% = xMitte% + 1
            End If
        Wend
  
        For i% = xRechts% To besMax%
            Get #dHandle, i% + 1, bes
            If (key$ = bes.pzn) Then
                recno& = i%
                iFabsRecno& = recno&
                IndexSearch% = 0
                Exit For
            End If
            If (Left$(bes.text, 2) > AnfangsBuchstabe$) Then Exit For
        Next i%
    End If
Else
    FabsErrf% = clsDat.IndexSearch%(iHandle%(ind%), key$, recno&)
    iFabsRecno& = recno&
    IndexSearch% = FabsErrf%
End If

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
Dim besMax%

If (Para1.Land = "A") Then
    Get #dHandle%, 1, besErst
    besMax% = besErst.erstmax
    If (AltRecno& < besMax%) Then
        recno& = AltRecno& + 1
        iFabsRecno& = recno&
        IndexNext% = 0
    Else
        recno& = 0
        iFabsRecno& = recno&
        IndexNext% = 8
    End If
Else
    FabsErrf% = clsDat.IndexNext%(iHandle%(ind%), AltRecno&, key$, recno&)
    iFabsRecno& = recno&
    IndexNext% = FabsErrf%
End If

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
erstmax% = besErst.erstmax
End Property
Public Property Let erstmax(ByVal vNewValue%)
besErst.erstmax = vNewValue%
End Property

Public Property Get erstdruck$()
erstdruck$ = besErst.erstdruck
End Property
Public Property Let erstdruck(ByVal vNewValue$)
besErst.erstdruck = vNewValue$
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


