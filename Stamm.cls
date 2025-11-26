VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsStamm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Const SATZLEN = 100


Private Type StammErstSatz
    erstmax As Integer
    erstrest As String * 98
End Type
Private Type StammStruct
    pzn         As String * 7
    kurz        As String * 28
    meng        As String * 5
    meh         As String * 2
    aep         As Double
    kp          As Double
    AVP         As Double
    mw          As String * 1
    ka          As String * 1
    wg          As String * 2
    herst       As String * 5
    lic         As String * 1
    vc          As String * 1
    gc          As String * 2
    lac         As String * 1
    rez         As String * 2
    kas         As String * 3
    kasz        As String * 2
    zuza        As Integer
    Vnr         As String * 8
    kz          As String * 1
    abl         As String * 1
    frei        As String * 1
End Type
Dim astErst As StammErstSatz
Dim ast As StammStruct

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

Dim iPzn8$, iKurzNeu$, iMengNeu$

Private Const DefErrModul = "STAMM.CLS"

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
dKurz$ = "AST"
dBezeichnung$ = "Lagerartikel"
dName$ = "artikel.tax"
dZugriff$ = "RW"
dModus$ = "R"
dSatzLen% = SATZLEN 'Len(ast)

dAnzFields% = 24
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

fKurz$(7) = "MW"
fBezeichnung$(7) = "Mehrwertsteuer-Code"
fStart%(7) = 67
fLaenge%(7) = 1
fTyp$(7) = "T"
fKonv%(7) = 0

fKurz$(8) = "KA"
fBezeichnung$(8) = "Kennzeichen Aufschlag"
fStart%(8) = 68
fLaenge%(8) = 1
fTyp$(8) = "T"
fKonv%(8) = 0

fKurz$(9) = "WG1"
fBezeichnung$(9) = "Warengruppe-1"
fStart%(9) = 69
fLaenge%(9) = 1
fTyp$(9) = "N"
fKonv%(9) = 3

fKurz$(10) = "WG2"
fBezeichnung$(10) = "Warengruppe-2"
fStart%(10) = 70
fLaenge%(10) = 1
fTyp$(10) = "N"
fKonv%(10) = 3

fKurz$(11) = "HERST"
fBezeichnung$(11) = "Hersteller"
fStart%(11) = 71
fLaenge%(11) = 5
fTyp$(11) = "T"
fKonv%(11) = 0

fKurz$(12) = "LIC"
fBezeichnung$(12) = "Liefercode"
fStart%(12) = 76
fLaenge%(12) = 1
fTyp$(12) = "T"
fKonv%(12) = 0

fKurz$(13) = "VC"
fBezeichnung$(13) = "Veränderungscode"
fStart%(13) = 77
fLaenge%(13) = 1
fTyp$(13) = "T"
fKonv%(13) = 0

fKurz$(14) = "GC"
fBezeichnung$(14) = "Gültigkeit"
fStart%(14) = 78
fLaenge%(14) = 2
fTyp$(14) = "T"
fKonv%(14) = 0

fKurz$(15) = "LAC"
fBezeichnung$(15) = "Lagercode"
fStart%(15) = 80
fLaenge%(15) = 1
fTyp$(15) = "T"
fKonv%(15) = 0

fKurz$(16) = "REZ"
fBezeichnung$(16) = "Rezeptzeichen"
fStart%(16) = 81
fLaenge%(16) = 2
fTyp$(16) = "T"
fKonv%(16) = 0

fKurz$(17) = "KAS"
fBezeichnung$(17) = "Kassenzeichen"
fStart%(17) = 83
fLaenge%(17) = 3
fTyp$(17) = "T"
fKonv%(17) = 0

fKurz$(18) = "KASZ"
fBezeichnung$(18) = "Kassenzeichen-Zusatz"
fStart%(18) = 86
fLaenge%(18) = 2
fTyp$(18) = "T"
fKonv%(18) = 0

fKurz$(19) = "VNR"
fBezeichnung$(19) = "Verlagsnummer"
fStart%(19) = 90
fLaenge%(19) = 8
fTyp$(19) = "T"
fKonv%(19) = 0

fKurz$(20) = "KZ"
fBezeichnung$(20) = "Kennzeichen (M, N, *)"
fStart%(20) = 98
fLaenge%(20) = 1
fTyp$(20) = "T"
fKonv%(20) = 0

fKurz$(21) = "ABL"
fBezeichnung$(21) = "Ablaufcode"
fStart%(21) = 99
fLaenge%(21) = 1
fTyp$(21) = "T"
fKonv%(21) = 0

fKurz$(22) = "MWP"
fBezeichnung$(22) = "Mehrwertsteuer"
fStart%(22) = 67
fLaenge%(22) = 1
fTyp$(22) = "N"
fKonv%(22) = 13

fKurz$(23) = "ZUZA"
fBezeichnung$(23) = "Zuzahlung"
fStart%(23) = 88
fLaenge%(23) = 2
fTyp$(23) = "N"
fKonv%(23) = 8


iAnzIndex% = 2
iName$(0) = "artikel1.idx"
iName$(1) = "artikel2.idx"


Set Ast1 = Me

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
Set Ast1 = Nothing
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

If (Para1.Land = "A") Then
    iName$(1) = "artikel3.idx"
    iName$(2) = "artikel2.idx"
    iAnzIndex% = 3
End If


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
CopyMemory ast, ByVal temp$, SATZLEN
ast.pzn = SollPzn$
ast.kurz = "A" + Space$(Len(ast.kurz) - 1)
ast.meng = Space$(Len(ast.meng))
ast.meh = Space$(Len(ast.meh))

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

Get #dHandle%, 1, astErst

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
    Get #dHandle%, , ast
    iPzn8 = "0" + ast.pzn
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
    Put #dHandle%, , astErst
Else
    Put #dHandle%, , ast
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






Public Property Get erstmax%()
erstmax% = astErst.erstmax
End Property
Public Property Let erstmax(ByVal vNewValue%)
astErst.erstmax = vNewValue%
End Property

Public Property Get erstrest$()
erstrest$ = astErst.erstrest
End Property
Public Property Let erstrest(ByVal vNewValue$)
astErst.erstrest = vNewValue$
End Property




Public Property Get pzn$()
'pzn$ = ast.pzn
pzn = iPzn8
End Property
Public Property Let pzn(ByVal vNewValue$)
ast.pzn = Right(vNewValue$, 7)
iPzn8 = vNewValue
End Property


Public Property Get kurz$()
kurz$ = ast.kurz
End Property
Public Property Let kurz(ByVal vNewValue$)
ast.kurz = vNewValue$
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
meng$ = ast.meng
End Property
Public Property Let meng(ByVal vNewValue$)
ast.meng = vNewValue$
End Property

Public Property Get MengNeu$()
MengNeu$ = iMengNeu
End Property
Public Property Let MengNeu(ByVal vNewValue$)
iMengNeu = vNewValue$
End Property

Public Property Get meh$()
meh$ = ast.meh
End Property
Public Property Let meh(ByVal vNewValue$)
ast.meh = vNewValue$
End Property

Public Property Get aep#()
Dim a#
a# = ast.aep
Call DxToIEEEd(a#)
aep# = a#
End Property
Public Property Let aep(ByVal vNewValue#)
Dim a#
a# = vNewValue#
Call DxToMBFd(a#)
ast.aep = a#
End Property

Public Property Get kp#()
Dim a#
a# = ast.kp
Call DxToIEEEd(a#)
kp# = a#
End Property
Public Property Let kp(ByVal vNewValue#)
Dim a#
a# = vNewValue#
Call DxToMBFd(a#)
ast.kp = a#
End Property

Public Property Get AVP#()
Dim a#
a# = ast.AVP
Call DxToIEEEd(a#)
AVP# = a#
End Property
Public Property Let AVP(ByVal vNewValue#)
Dim a#
a# = vNewValue#
Call DxToMBFd(a#)
ast.AVP = a#
End Property

Public Property Get mw$()
mw$ = ast.mw
End Property
Public Property Let mw(ByVal vNewValue$)
ast.mw = vNewValue$
End Property

Public Property Get ka$()
ka$ = ast.ka
End Property
Public Property Let ka(ByVal vNewValue$)
ast.ka = vNewValue$
End Property

Public Property Get wg$()
wg$ = ast.wg
End Property
Public Property Let wg(ByVal vNewValue$)
ast.wg = vNewValue$
End Property

Public Property Get wg1$()
wg1$ = Left$(ast.wg, 1)
End Property
Public Property Let wg1(ByVal vNewValue$)
Mid$(ast.wg, 1, 1) = vNewValue$
End Property

Public Property Get wg2$()
wg2$ = Mid$(ast.wg, 2, 1)
End Property
Public Property Let wg2(ByVal vNewValue$)
Mid$(ast.wg, 2, 1) = vNewValue$
End Property

Public Property Get herst$()
herst$ = ast.herst
End Property
Public Property Let herst(ByVal vNewValue$)
ast.herst = vNewValue$
End Property

Public Property Get lic$()
lic$ = ast.lic
End Property
Public Property Let lic(ByVal vNewValue$)
ast.lic = vNewValue$
End Property

Public Property Get vc$()
vc$ = ast.vc
End Property
Public Property Let vc(ByVal vNewValue$)
ast.vc = vNewValue$
End Property

Public Property Get gc$()
gc$ = ast.gc
End Property
Public Property Let gc(ByVal vNewValue$)
ast.gc = vNewValue$
End Property

Public Property Get lac$()
lac$ = ast.lac
End Property
Public Property Let lac(ByVal vNewValue$)
ast.lac = vNewValue$
End Property

Public Property Get rez$()
rez$ = ast.rez
End Property
Public Property Let rez(ByVal vNewValue$)
ast.rez = vNewValue$
End Property

Public Property Get kas$()
kas$ = ast.kas
End Property
Public Property Let kas(ByVal vNewValue$)
ast.kas = vNewValue$
End Property

Public Property Get kasz$()
kasz$ = ast.kasz
End Property
Public Property Let kasz(ByVal vNewValue$)
ast.kasz = vNewValue$
End Property

Public Property Get zuza&()
zuza& = CLng(ast.zuza) - (ast.zuza < 0) * 65536
End Property
Public Property Let zuza(ByVal vNewValue&)
ast.zuza = Int(vNewValue& + (vNewValue& > 32767) * 65536)
End Property

Public Property Get Vnr$()
Vnr$ = ast.Vnr
End Property
Public Property Let Vnr(ByVal vNewValue$)
ast.Vnr = vNewValue$
End Property

Public Property Get kz$()
kz$ = ast.kz
End Property
Public Property Let kz(ByVal vNewValue$)
ast.kz = vNewValue$
End Property

Public Property Get abl$()
abl$ = ast.abl
End Property
Public Property Let abl(ByVal vNewValue$)
ast.abl = vNewValue$
End Property

Public Property Get frei$()
frei$ = ast.frei
End Property
Public Property Let frei(ByVal vNewValue$)
ast.frei = vNewValue$
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
Put #Handle%, , ast
Call clsError.DefErrPop
End Sub


