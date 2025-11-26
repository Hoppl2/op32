VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsNNEK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Const SATZLEN = 200

Private Type NNEKErstSatz
    erstmax As Integer
    erstrest As String * 198
End Type

Private Type NNEKStruct1
    BelegDat As String * 2
    NNAEP As Double
    art As Integer
End Type

Private Type NNEKStruct
    pzn As String * 7
    lieferung(4) As NNEKStruct1
    Beleg(4) As String * 10
    lm(4) As Integer
    rm(4) As Integer
    LiefNr(4) As Byte
    GhNr(4) As Byte
    RetMenge(4) As Integer
    rest As String * 43
End Type
Dim nnek As NNEKStruct
Dim nnekErst As NNEKErstSatz

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



Private Const DefErrModul = "NNEK.CLS"

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
dKurz$ = "NNEK"
dBezeichnung$ = "Netto-Netto-EK"
dName$ = "nnaep.dat"
dZugriff$ = "RW"
dModus$ = "R"
dSatzLen% = SATZLEN 'Len(lif)

dAnzFields% = 42
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

fKurz$(1) = "BELEGDAT(4)"
fBezeichnung$(1) = "Lieferung 1: Beleg-Datum"
fStart%(1) = 8
fLaenge%(1) = 2
fTyp$(1) = "D"
fKonv%(1) = 10

fKurz$(2) = "NNAEP(4)"
fBezeichnung$(2) = "Lieferung 1: Netto-Netto-EK"
fStart%(2) = 10
fLaenge%(2) = 8
fTyp$(2) = "N"
fKonv%(2) = 2

fKurz$(3) = "ART(4)"
fBezeichnung$(3) = "Lieferung 1: Art"
fStart%(3) = 18
fLaenge%(3) = 2
fTyp$(3) = "N"
fKonv%(3) = 5

fKurz$(4) = "BELEGDAT(3)"
fBezeichnung$(4) = "Lieferung 2: Beleg-Datum"
fStart%(4) = 20
fLaenge%(4) = 2
fTyp$(4) = "D"
fKonv%(4) = 10

fKurz$(5) = "NNAEP(3)"
fBezeichnung$(5) = "Lieferung 2: Netto-Netto-EK"
fStart%(5) = 22
fLaenge%(5) = 8
fTyp$(5) = "N"
fKonv%(5) = 2

fKurz$(6) = "ART(3)"
fBezeichnung$(6) = "Lieferung 2: Art"
fStart%(6) = 30
fLaenge%(6) = 2
fTyp$(6) = "N"
fKonv%(6) = 5

fKurz$(7) = "BELEGDAT(2)"
fBezeichnung$(7) = "Lieferung 3: Beleg-Datum"
fStart%(7) = 32
fLaenge%(7) = 2
fTyp$(7) = "D"
fKonv%(7) = 10

fKurz$(8) = "NNAEP(2)"
fBezeichnung$(8) = "Lieferung 3: Netto-Netto-EK"
fStart%(8) = 34
fLaenge%(8) = 8
fTyp$(8) = "N"
fKonv%(8) = 2

fKurz$(9) = "ART(2)"
fBezeichnung$(9) = "Lieferung 3: Art"
fStart%(9) = 42
fLaenge%(9) = 2
fTyp$(9) = "N"
fKonv%(9) = 5

fKurz$(10) = "BELEGDAT(1)"
fBezeichnung$(10) = "Lieferung 4: Beleg-Datum"
fStart%(10) = 44
fLaenge%(10) = 2
fTyp$(10) = "D"
fKonv%(10) = 10

fKurz$(11) = "NNAEP(1)"
fBezeichnung$(11) = "Lieferung 4: Netto-Netto-EK"
fStart%(11) = 46
fLaenge%(11) = 8
fTyp$(11) = "N"
fKonv%(11) = 2

fKurz$(12) = "ART(1)"
fBezeichnung$(12) = "Lieferung 4: Art"
fStart%(12) = 54
fLaenge%(12) = 2
fTyp$(12) = "N"
fKonv%(12) = 5

fKurz$(13) = "BELEGDAT(0)"
fBezeichnung$(13) = "Lieferung 5: Beleg-Datum"
fStart%(13) = 56
fLaenge%(13) = 2
fTyp$(13) = "D"
fKonv%(13) = 10

fKurz$(14) = "NNAEP(0)"
fBezeichnung$(14) = "Lieferung 5: Netto-Netto-EK"
fStart%(14) = 58
fLaenge%(14) = 8
fTyp$(14) = "N"
fKonv%(14) = 2

fKurz$(15) = "ART(0)"
fBezeichnung$(15) = "Lieferung 5: Art"
fStart%(15) = 66
fLaenge%(15) = 2
fTyp$(15) = "N"
fKonv%(15) = 5

fKurz$(16) = "BELEG(4)"
fBezeichnung$(16) = "Lieferung 1: Beleg"
fStart%(16) = 68
fLaenge%(16) = 10
fTyp$(16) = "T"
fKonv%(16) = 0

fKurz$(17) = "BELEG(3)"
fBezeichnung$(17) = "Lieferung 2: Beleg"
fStart%(17) = 78
fLaenge%(17) = 10
fTyp$(17) = "T"
fKonv%(17) = 0

fKurz$(18) = "BELEG(2)"
fBezeichnung$(18) = "Lieferung 3: Beleg"
fStart%(18) = 88
fLaenge%(18) = 10
fTyp$(18) = "T"
fKonv%(18) = 0

fKurz$(19) = "BELEG(1)"
fBezeichnung$(19) = "Lieferung 4: Beleg"
fStart%(19) = 98
fLaenge%(19) = 10
fTyp$(19) = "T"
fKonv%(19) = 0

fKurz$(20) = "BELEG(0)"
fBezeichnung$(20) = "Lieferung 5: Beleg"
fStart%(20) = 108
fLaenge%(20) = 10
fTyp$(20) = "T"
fKonv%(20) = 0

fKurz$(21) = "LM(4)"
fBezeichnung$(21) = "Lieferung 1: LM"
fStart%(21) = 118
fLaenge%(21) = 2
fTyp$(21) = "N"
fKonv%(21) = 5

fKurz$(22) = "LM(3)"
fBezeichnung$(22) = "Lieferung 2: LM"
fStart%(22) = 120
fLaenge%(22) = 2
fTyp$(22) = "N"
fKonv%(22) = 5

fKurz$(23) = "LM(2)"
fBezeichnung$(23) = "Lieferung 3: LM"
fStart%(23) = 122
fLaenge%(23) = 2
fTyp$(23) = "N"
fKonv%(23) = 5

fKurz$(24) = "LM(1)"
fBezeichnung$(24) = "Lieferung 4: LM"
fStart%(24) = 124
fLaenge%(24) = 2
fTyp$(24) = "N"
fKonv%(24) = 5

fKurz$(25) = "LM(0)"
fBezeichnung$(25) = "Lieferung 5: LM"
fStart%(25) = 126
fLaenge%(25) = 2
fTyp$(25) = "N"
fKonv%(25) = 5

fKurz$(26) = "RM(4)"
fBezeichnung$(26) = "Lieferung 1: RM"
fStart%(26) = 128
fLaenge%(26) = 2
fTyp$(26) = "N"
fKonv%(26) = 5

fKurz$(27) = "RM(3)"
fBezeichnung$(27) = "Lieferung 2: RM"
fStart%(27) = 130
fLaenge%(27) = 2
fTyp$(27) = "N"
fKonv%(27) = 5

fKurz$(28) = "RM(2)"
fBezeichnung$(28) = "Lieferung 3: RM"
fStart%(28) = 132
fLaenge%(28) = 2
fTyp$(28) = "N"
fKonv%(28) = 5

fKurz$(29) = "RM(1)"
fBezeichnung$(29) = "Lieferung 4: RM"
fStart%(29) = 134
fLaenge%(29) = 2
fTyp$(29) = "N"
fKonv%(29) = 5

fKurz$(30) = "RM(0)"
fBezeichnung$(30) = "Lieferung 5: RM"
fStart%(30) = 136
fLaenge%(30) = 2
fTyp$(30) = "N"
fKonv%(30) = 5

fKurz$(31) = "LIEF(4)"
fBezeichnung$(31) = "Lieferung 1: Lieferant"
fStart%(31) = 138
fLaenge%(31) = 1
fTyp$(31) = "N"
fKonv%(31) = 1

fKurz$(32) = "LIEF(3)"
fBezeichnung$(32) = "Lieferung 2: Lieferant"
fStart%(32) = 139
fLaenge%(32) = 1
fTyp$(32) = "N"
fKonv%(32) = 1

fKurz$(33) = "LIEF(2)"
fBezeichnung$(33) = "Lieferung 3: Lieferant"
fStart%(33) = 140
fLaenge%(33) = 1
fTyp$(33) = "N"
fKonv%(33) = 1

fKurz$(34) = "LIEF(1)"
fBezeichnung$(34) = "Lieferung 4: Lieferant"
fStart%(34) = 141
fLaenge%(34) = 1
fTyp$(34) = "N"
fKonv%(34) = 1

fKurz$(35) = "LIEF(0)"
fBezeichnung$(35) = "Lieferung 5: Lieferant"
fStart%(35) = 142
fLaenge%(35) = 1
fTyp$(35) = "N"
fKonv%(35) = 1

fKurz$(36) = "GHNR(4)"
fBezeichnung$(36) = "Lieferung 1: Gh-Nr"
fStart%(36) = 143
fLaenge%(36) = 1
fTyp$(36) = "N"
fKonv%(36) = 1

fKurz$(37) = "GHNR(3)"
fBezeichnung$(37) = "Lieferung 2: Gh-Nr"
fStart%(37) = 144
fLaenge%(37) = 1
fTyp$(37) = "N"
fKonv%(37) = 1

fKurz$(38) = "GHNR(2)"
fBezeichnung$(38) = "Lieferung 3: Gh-Nr"
fStart%(38) = 145
fLaenge%(38) = 1
fTyp$(38) = "N"
fKonv%(38) = 1

fKurz$(39) = "GHNR(1)"
fBezeichnung$(39) = "Lieferung 4: Gh-Nr"
fStart%(39) = 146
fLaenge%(39) = 1
fTyp$(39) = "N"
fKonv%(39) = 1

fKurz$(40) = "GHNR(0)"
fBezeichnung$(40) = "Lieferung 5: Gh-Nr"
fStart%(40) = 147
fLaenge%(40) = 1
fTyp$(40) = "N"
fKonv%(40) = 1

fKurz$(41) = "REST"
fBezeichnung$(41) = "Rest"
fStart%(41) = 148
fLaenge%(41) = 53
fTyp$(41) = "T"
fKonv%(41) = 0



iAnzIndex% = 1
iName$(0) = "nnaep.idx"

Set nnek1 = Me

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
Set nnek1 = Nothing
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

Get #dHandle%, 1, nnekErst

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
    Call GetFirstRecord
Else
    Get #dHandle%, , nnek
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
    Put #dHandle%, , nnekErst
Else
    Put #dHandle%, , nnek
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


Public Property Get erstmax&()
erstmax& = CLng(nnekErst.erstmax) - (nnekErst.erstmax < 0) * 65536
End Property
Public Property Let erstmax(ByVal vNewValue&)
nnekErst.erstmax = Int(vNewValue& + (vNewValue& > 32767) * 65536)
End Property

Public Property Get erstrest$()
erstrest$ = nnekErst.erstrest
End Property
Public Property Let erstrest(ByVal vNewValue$)
nnekErst.erstrest = vNewValue$
End Property



Public Property Get pzn$()
pzn$ = nnek.pzn
End Property
Public Property Let pzn(ByVal vNewValue$)
nnek.pzn = vNewValue$
End Property

Public Property Get BelegDat%(ByVal ind%)
BelegDat% = clsOpTool.CVDat(nnek.lieferung(ind%).BelegDat)
End Property
Public Property Let BelegDat(ByVal ind%, ByVal vNewValue%)
nnek.lieferung(ind%).BelegDat = clsOpTool.MKDate(vNewValue%)
End Property

Public Property Get NNAEP#(ByVal ind%)
Dim a#
a# = nnek.lieferung(ind%).NNAEP
Call DxToIEEEd(a#)
NNAEP# = a#
End Property
Public Property Let NNAEP(ByVal ind%, ByVal vNewValue#)
Dim a#
a# = vNewValue#
Call DxToMBFd(a#)
nnek.lieferung(ind%).NNAEP = a#
End Property

Public Property Get art%(ByVal ind%)
art% = nnek.lieferung(ind%).art
End Property
Public Property Let art(ByVal ind%, ByVal vNewValue%)
nnek.lieferung(ind%).art = vNewValue%
End Property

Public Property Get Beleg$(ByVal ind%)
Beleg$ = nnek.Beleg(ind%)
End Property
Public Property Let Beleg(ByVal ind%, ByVal vNewValue$)
nnek.Beleg(ind%) = vNewValue$
End Property

Public Property Get lm%(ByVal ind%)
lm% = nnek.lm(ind%)
End Property
Public Property Let lm(ByVal ind%, ByVal vNewValue%)
nnek.lm(ind%) = vNewValue%
End Property

Public Property Get rm%(ByVal ind%)
rm% = nnek.rm(ind%)
End Property
Public Property Let rm(ByVal ind%, ByVal vNewValue%)
nnek.rm(ind%) = vNewValue%
End Property

Public Property Get LiefNr(ByVal ind%) As Byte
LiefNr = nnek.LiefNr(ind%)
End Property
Public Property Let LiefNr(ByVal ind%, ByVal vNewValue As Byte)
nnek.LiefNr(ind%) = vNewValue
End Property

Public Property Get GhNr(ByVal ind%) As Byte
GhNr = nnek.GhNr(ind%)
End Property
Public Property Let GhNr(ByVal ind%, ByVal vNewValue As Byte)
nnek.GhNr(ind%) = vNewValue
End Property

Public Property Get RetMenge%(ByVal ind%)
Dim ret%

ret = 0
If (nnek.RetMenge(ind) >= 1000) And (nnek.RetMenge(ind) <= 2000) Then
    ret = nnek.RetMenge(ind) - 1000
End If

RetMenge = ret
End Property
Public Property Let RetMenge(ByVal ind%, ByVal vNewValue%)
nnek.RetMenge(ind) = vNewValue + 1000
End Property

Public Property Get rest$()
rest$ = nnek.rest
End Property
Public Property Let rest(ByVal vNewValue$)
nnek.rest = vNewValue$
End Property


