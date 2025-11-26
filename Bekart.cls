VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsBekart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Const SATZLEN = 128

Private Type BekartErstSatz
    max As Integer
    lief As Integer
    Counter As Byte
    rest As String * 123
End Type
Private Type BekartStruct
    pzn As String * 7
    txt As String * 35
    lief As Byte    'String * 1
    bm As Integer
    asatz As Integer
    ssatz As Integer
    best As String * 1
    nm As Integer
    aep As Double
    abl As String * 1
    wg As String * 1
    AVP As Double
    km As Integer
    absage As Byte  'String * 1
    angebot As Byte 'String * 1
    auto As String * 1
    alt As String * 1
    zr As Integer   'String * 2
    besorger As String * 1
    NNAEP As Double
    nnart As Byte   'String * 1
    Lief1 As Byte   'String * 1
    bm1 As Integer
    nm1 As Integer
    AbholNr As Integer
    zugeordnet As String * 1
    zukontrollieren As String * 1
    fixiert As String * 1
    nochzukontrollieren As String * 1
    actkontrolle(5) As Byte
    actzuordnung As Byte
    poslag As Integer
    AngebotInd As Byte
    aktivlief As Byte
    aktivind As Integer
    beklaufnr As Long
    herst As String * 5
    LiefName As String * 6
    dummy As String * 1
End Type
Dim bekErst As BekartErstSatz
Dim bek As BekartStruct

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



Private Const DefErrModul = "BEKART.CLS"

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
dKurz$ = "BEKART"
dBezeichnung$ = "Bestellungen"
dName$ = "bekart.dat"
dZugriff$ = "RW"
dModus$ = "R"
dSatzLen% = SATZLEN 'Len(bek)

dAnzFields% = 47
ReDim fKurz$(dAnzFields% - 1)
ReDim fBezeichnung$(dAnzFields% - 1)
ReDim fStart%(dAnzFields% - 1)
ReDim fLaenge%(dAnzFields% - 1)
ReDim fTyp$(dAnzFields% - 1)
ReDim fKonv%(dAnzFields% - 1)


fKurz$(0) = "PZN"
fBezeichnung$(0) = "Pzn"
fStart%(0) = 1
fLaenge%(0) = 7
fTyp$(0) = "T"
fKonv%(0) = 0

fKurz$(1) = "TXT"
fBezeichnung$(1) = "Txt"
fStart%(1) = 8
fLaenge%(1) = 35
fTyp$(1) = "T"
fKonv%(1) = 0

fKurz$(2) = "LIEF"
fBezeichnung$(2) = "Lief"
fStart%(2) = 43
fLaenge%(2) = 1
fTyp$(2) = "N"
fKonv%(2) = 1

fKurz$(3) = "BM"
fBezeichnung$(3) = "Bm"
fStart%(3) = 44
fLaenge%(3) = 2
fTyp$(3) = "N"
fKonv%(3) = 5

fKurz$(4) = "ASATZ"
fBezeichnung$(4) = "Asatz"
fStart%(4) = 46
fLaenge%(4) = 2
fTyp$(4) = "N"
fKonv%(4) = 5

fKurz$(5) = "SSATZ"
fBezeichnung$(5) = "Ssatz"
fStart%(5) = 48
fLaenge%(5) = 2
fTyp$(5) = "N"
fKonv%(5) = 5

fKurz$(6) = "BEST"
fBezeichnung$(6) = "Best"
fStart%(6) = 50
fLaenge%(6) = 1
fTyp$(6) = "T"
fKonv%(6) = 0

fKurz$(7) = "NM"
fBezeichnung$(7) = "Nm"
fStart%(7) = 51
fLaenge%(7) = 2
fTyp$(7) = "N"
fKonv%(7) = 5

fKurz$(8) = "AEP"
fBezeichnung$(8) = "Aep"
fStart%(8) = 53
fLaenge%(8) = 8
fTyp$(8) = "N"
fKonv%(8) = 2

fKurz$(9) = "ABL"
fBezeichnung$(9) = "Abl"
fStart%(9) = 61
fLaenge%(9) = 1
fTyp$(9) = "T"
fKonv%(9) = 0

fKurz$(10) = "WG"
fBezeichnung$(10) = "Wg"
fStart%(10) = 62
fLaenge%(10) = 1
fTyp$(10) = "T"
fKonv%(10) = 0

fKurz$(11) = "AVP"
fBezeichnung$(11) = "Avp"
fStart%(11) = 63
fLaenge%(11) = 8
fTyp$(11) = "N"
fKonv%(11) = 2

fKurz$(12) = "KM"
fBezeichnung$(12) = "Km"
fStart%(12) = 71
fLaenge%(12) = 2
fTyp$(12) = "N"
fKonv%(12) = 5

fKurz$(13) = "ABSAGE"
fBezeichnung$(13) = "Absage"
fStart%(13) = 73
fLaenge%(13) = 1
fTyp$(13) = "N"
fKonv%(13) = 1

fKurz$(14) = "ANGEBOT"
fBezeichnung$(14) = "Angebot"
fStart%(14) = 74
fLaenge%(14) = 1
fTyp$(14) = "N"
fKonv%(14) = 1

fKurz$(15) = "AUTO"
fBezeichnung$(15) = "Auto"
fStart%(15) = 75
fLaenge%(15) = 1
fTyp$(15) = "T"
fKonv%(15) = 0

fKurz$(16) = "ALT"
fBezeichnung$(16) = "Alt"
fStart%(16) = 76
fLaenge%(16) = 1
fTyp$(16) = "T"
fKonv%(16) = 0

fKurz$(17) = "ZR"
fBezeichnung$(17) = "Zr"
fStart%(17) = 77
fLaenge%(17) = 2
fTyp$(17) = "T"
fKonv%(17) = 0

fKurz$(18) = "BESORGER"
fBezeichnung$(18) = "Besorger"
fStart%(18) = 79
fLaenge%(18) = 1
fTyp$(18) = "T"
fKonv%(18) = 0

fKurz$(19) = "NNAEP"
fBezeichnung$(19) = "NNAep"
fStart%(19) = 80
fLaenge%(19) = 8
fTyp$(19) = "N"
fKonv%(19) = 2

fKurz$(20) = "NNART"
fBezeichnung$(20) = "NNArt"
fStart%(20) = 88
fLaenge%(20) = 1
fTyp$(20) = "N"
fKonv%(20) = 1

fKurz$(21) = "LIEF1"
fBezeichnung$(21) = "Lief1"
fStart%(21) = 89
fLaenge%(21) = 1
fTyp$(21) = "N"
fKonv%(21) = 1

fKurz$(22) = "BM1"
fBezeichnung$(22) = "Bm1"
fStart%(22) = 90
fLaenge%(22) = 2
fTyp$(22) = "N"
fKonv%(22) = 5

fKurz$(23) = "NM1"
fBezeichnung$(23) = "Nm1"
fStart%(23) = 92
fLaenge%(23) = 2
fTyp$(23) = "N"
fKonv%(23) = 5

fKurz$(24) = "ABHOLNR"
fBezeichnung$(24) = "AbholNr"
fStart%(24) = 94
fLaenge%(24) = 2
fTyp$(24) = "N"
fKonv%(24) = 5

fKurz$(25) = "ZUGEORDNET"
fBezeichnung$(25) = "Zugeordnet"
fStart%(25) = 96
fLaenge%(25) = 1
fTyp$(25) = "T"
fKonv%(25) = 0

fKurz$(26) = "ZUKONTROLLIEREN"
fBezeichnung$(26) = "ZuKontrollieren"
fStart%(26) = 97
fLaenge%(26) = 1
fTyp$(26) = "T"
fKonv%(26) = 0

fKurz$(27) = "MUSSKONTROLLIEREN"
fBezeichnung$(27) = "MussKontrollieren"
fStart%(27) = 98
fLaenge%(27) = 1
fTyp$(27) = "T"
fKonv%(27) = 0

fKurz$(28) = "NOCHZUKONTROLLIEREN"
fBezeichnung$(28) = "NochZuKontrollieren"
fStart%(28) = 99
fLaenge%(28) = 1
fTyp$(28) = "T"
fKonv%(28) = 0

fKurz$(29) = "ACTKONTROLLE(0)"
fBezeichnung$(29) = "ActKontrolle0"
fStart%(29) = 100
fLaenge%(29) = 1
fTyp$(29) = "N"
fKonv%(29) = 1

fKurz$(30) = "ACTKONTROLLE(1)"
fBezeichnung$(30) = "ActKontrolle1"
fStart%(30) = 101
fLaenge%(30) = 1
fTyp$(30) = "N"
fKonv%(30) = 1

fKurz$(31) = "ACTKONTROLLE(2)"
fBezeichnung$(31) = "ActKontrolle2"
fStart%(31) = 102
fLaenge%(31) = 1
fTyp$(31) = "N"
fKonv%(31) = 1

fKurz$(32) = "ACTKONTROLLE(3)"
fBezeichnung$(32) = "ActKontrolle3"
fStart%(32) = 103
fLaenge%(32) = 1
fTyp$(32) = "N"
fKonv%(32) = 1

fKurz$(33) = "ACTKONTROLLE(4)"
fBezeichnung$(33) = "ActKontrolle4"
fStart%(33) = 104
fLaenge%(33) = 1
fTyp$(33) = "N"
fKonv%(33) = 1

fKurz$(34) = "ACTKONTROLLE(5)"
fBezeichnung$(34) = "ActKontrolle5"
fStart%(34) = 105
fLaenge%(34) = 1
fTyp$(34) = "N"
fKonv%(34) = 1

fKurz$(35) = "ACTZUORDNUNG"
fBezeichnung$(35) = "ActZuordnung"
fStart%(35) = 110
fLaenge%(35) = 1
fTyp$(35) = "N"
fKonv%(35) = 1

fKurz$(36) = "POSLAG"
fBezeichnung$(36) = "PosLag"
fStart%(36) = 111
fLaenge%(36) = 2
fTyp$(36) = "N"
fKonv%(36) = 5

fKurz$(37) = "ANGEBOTIND"
fBezeichnung$(37) = "AngebotInd"
fStart%(37) = 113
fLaenge%(37) = 1
fTyp$(37) = "N"
fKonv%(37) = 1

fKurz$(38) = "AKTIVLIEF"
fBezeichnung$(38) = "AktivLief"
fStart%(38) = 114
fLaenge%(38) = 1
fTyp$(38) = "N"
fKonv%(38) = 1

fKurz$(39) = "AKTIVIND"
fBezeichnung$(39) = "AktivInd"
fStart%(39) = 115
fLaenge%(39) = 2
fTyp$(39) = "N"
fKonv%(39) = 5

fKurz$(40) = "BEKLAUFNR"
fBezeichnung$(40) = "BekLaufNr"
fStart%(40) = 117
fLaenge%(40) = 4
fTyp$(40) = "N"
fKonv%(40) = 7

fKurz$(41) = "REST"
fBezeichnung$(41) = "Rest"
fStart%(41) = 121
fLaenge%(41) = 8
fTyp$(41) = "T"
fKonv%(41) = 0

fKurz$(42) = "ERSTMAX"
fBezeichnung$(42) = "ErstMax"
fStart%(42) = 1
fLaenge%(42) = 2
fTyp$(42) = "N"
fKonv%(42) = 5

fKurz$(43) = "ERSTLIEF"
fBezeichnung$(43) = "ErstLieferant"
fStart%(43) = 3
fLaenge%(43) = 2
fTyp$(43) = "N"
fKonv%(43) = 5

fKurz$(44) = "ERSTCOUNTER"
fBezeichnung$(44) = "ErstCounter"
fStart%(44) = 5
fLaenge%(44) = 1
fTyp$(44) = "N"
fKonv%(44) = 1

fKurz$(45) = "ERSTREST"
fBezeichnung$(45) = "ErstRest"
fStart%(45) = 6
fLaenge%(45) = 123
fTyp$(45) = "T"
fKonv%(45) = 0


fKurz$(46) = "ALLES"
fBezeichnung$(46) = "Alles"
fStart%(46) = 1
fLaenge%(46) = SATZLEN
fTyp$(46) = "T"
fKonv%(46) = 0




iAnzIndex% = 0


Set Bek1 = Me

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
Set Bek1 = Nothing
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

Sub SatzLock(Recno&)
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
Call clsDat.iLock(dHandle%, Recno&)
Call clsError.DefErrPop
End Sub

Sub SatzUnLock(Recno&)
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
Call clsDat.iUnLock(dHandle%, Recno&)
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

Get #dHandle%, 1, bekErst

Call clsError.DefErrPop
End Sub

Sub GetRecord(Optional Recno& = 0)
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
If (Recno& > 0) Then
    Seek #dHandle%, Recno&
End If
If (Recno& = 1) Then
'    Get #dHandle%, , bekErst
    Call GetFirstRecord
Else
    Get #dHandle%, , bek
End If
Call clsError.DefErrPop
End Sub

Sub PutRecord(Optional Recno& = 0)
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
If (Recno& > 0) Then
    Seek #dHandle%, Recno&
End If
If (Recno& = 1) Then
    Put #dHandle%, , bekErst
Else
    Put #dHandle%, , bek
End If
Call clsError.DefErrPop
End Sub

Sub GetInfoRecord(Optional Recno& = 0)
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
If (Recno& > 0) Then
    Seek #dHandle%, Recno&
End If
Get #dHandle%, , AktData
Call clsError.DefErrPop
End Sub

Function IndexSearch%(ind%, key$, Recno&)
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
FabsErrf% = clsDat.IndexSearch%(iHandle%(ind%), key$, Recno&)
iFabsRecno& = Recno&
IndexSearch% = FabsErrf%
Call clsError.DefErrPop
End Function

Function IndexGeneric%(ind%, key$, Recno&)
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
FabsErrf% = clsDat.IndexGeneric%(iHandle%(ind%), key$, Recno&)
iFabsRecno& = Recno&
IndexGeneric% = FabsErrf%
Call clsError.DefErrPop
End Function

Function IndexNext%(ind%, AltRecno&, key$, Recno&)
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
FabsErrf% = clsDat.IndexNext%(iHandle%(ind%), AltRecno&, key$, Recno&)
iFabsRecno& = Recno&
IndexNext% = FabsErrf%
Call clsError.DefErrPop
End Function

Function IndexPrevious%(ind%, AltRecno&, key$, Recno&)
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
FabsErrf% = clsDat.IndexPrevious%(iHandle%(ind%), AltRecno&, key$, Recno&)
iFabsRecno& = Recno&
IndexPrevious% = FabsErrf%
Call clsError.DefErrPop
End Function

Function IndexInsert%(ind%, key$, Recno&)
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
FabsErrf% = clsDat.IndexInsert%(iHandle%(ind%), key$, Recno&)
iFabsRecno& = Recno&
IndexInsert% = FabsErrf%
Call clsError.DefErrPop
End Function

Function IndexDelete%(ind%, AltRecno&, key$, Recno&)
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
FabsErrf% = clsDat.IndexDelete%(iHandle%(ind%), AltRecno&, key$, Recno&)
iFabsRecno& = Recno&
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

Public Property Get wert(iFeld$, Optional iIndex% = -1) As Variant
Attribute wert.VB_UserMemId = 0
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
erstmax% = bekErst.max
End Property
Public Property Let erstmax(ByVal vNewValue%)
bekErst.max = vNewValue%
End Property

Public Property Get erstlief%()
erstlief% = bekErst.lief
End Property
Public Property Let erstlief(ByVal vNewValue%)
bekErst.lief = vNewValue%
End Property

Public Property Get erstcounter() As Byte
erstcounter = bekErst.Counter
End Property
Public Property Let erstcounter(ByVal vNewValue As Byte)
bekErst.Counter = vNewValue
End Property

Public Property Get erstrest$()
erstrest$ = bekErst.rest
End Property
Public Property Let erstrest(ByVal vNewValue$)
bekErst.rest = vNewValue$
End Property

Public Property Get pzn$()
pzn$ = bek.pzn
End Property
Public Property Let pzn(ByVal vNewValue$)
bek.pzn = vNewValue$
End Property

Public Property Get txt$()
txt$ = bek.txt
End Property
Public Property Let txt(ByVal vNewValue$)
bek.txt = vNewValue$
End Property

Public Property Get lief() As Byte
lief = bek.lief
End Property
Public Property Let lief(ByVal vNewValue As Byte)
bek.lief = vNewValue
End Property

Public Property Get bm%()
bm% = bek.bm
End Property
Public Property Let bm(ByVal vNewValue%)
bek.bm = vNewValue%
End Property

Public Property Get asatz%()
asatz% = bek.asatz
End Property
Public Property Let asatz(ByVal vNewValue%)
bek.asatz = vNewValue%
End Property

Public Property Get ssatz%()
ssatz% = bek.ssatz
End Property
Public Property Let ssatz(ByVal vNewValue%)
bek.ssatz = vNewValue%
End Property

Public Property Get best$()
best$ = bek.best
End Property
Public Property Let best(ByVal vNewValue$)
bek.best = vNewValue$
End Property

Public Property Get nm%()
nm% = bek.nm
End Property
Public Property Let nm(ByVal vNewValue%)
bek.nm = vNewValue%
End Property

Public Property Get aep#()
Dim a#
a# = bek.aep
Call DxToIEEEd(a#)
aep# = a#
End Property
Public Property Let aep(ByVal vNewValue#)
Dim a#
a# = vNewValue#
Call DxToMBFd(a#)
bek.aep = a#
End Property

Public Property Get abl$()
abl$ = bek.abl
End Property
Public Property Let abl(ByVal vNewValue$)
bek.abl = vNewValue$
End Property

Public Property Get wg$()
wg$ = bek.wg
End Property
Public Property Let wg(ByVal vNewValue$)
bek.wg = vNewValue$
End Property

Public Property Get AVP#()
Dim a#
a# = bek.AVP
Call DxToIEEEd(a#)
AVP# = a#
End Property
Public Property Let AVP(ByVal vNewValue#)
Dim a#
a# = vNewValue#
Call DxToMBFd(a#)
bek.AVP = a#
End Property

Public Property Get km%()
km% = bek.km
End Property
Public Property Let km(ByVal vNewValue%)
bek.km = vNewValue%
End Property

Public Property Get absage() As Byte
absage = bek.absage
End Property
Public Property Let absage(ByVal vNewValue As Byte)
bek.absage = vNewValue
End Property

Public Property Get angebot() As Byte
angebot = bek.angebot
End Property
Public Property Let angebot(ByVal vNewValue As Byte)
bek.angebot = vNewValue
End Property

Public Property Get auto$()
auto$ = bek.auto
End Property
Public Property Let auto(ByVal vNewValue$)
bek.auto = vNewValue$
End Property

Public Property Get alt$()
alt$ = bek.alt
End Property
Public Property Let alt(ByVal vNewValue$)
bek.alt = vNewValue$
End Property

Public Property Get zr%()
zr% = bek.zr
End Property
Public Property Let zr(ByVal vNewValue%)
bek.zr = vNewValue%
End Property

Public Property Get besorger$()
besorger$ = bek.besorger
End Property
Public Property Let besorger(ByVal vNewValue$)
bek.besorger = vNewValue$
End Property

Public Property Get NNAEP#()
NNAEP# = bek.NNAEP
End Property
Public Property Let NNAEP(ByVal vNewValue#)
bek.NNAEP = vNewValue#
End Property

Public Property Get nnart() As Byte
nnart = bek.nnart
End Property
Public Property Let nnart(ByVal vNewValue As Byte)
bek.nnart = vNewValue
End Property

Public Property Get Lief1() As Byte
Lief1 = bek.Lief1
End Property
Public Property Let Lief1(ByVal vNewValue As Byte)
bek.Lief1 = vNewValue
End Property

Public Property Get bm1%()
bm1% = bek.bm1
End Property
Public Property Let bm1(ByVal vNewValue%)
bek.bm1 = vNewValue%
End Property

Public Property Get nm1%()
nm1% = bek.nm1
End Property
Public Property Let nm1(ByVal vNewValue%)
bek.nm1 = vNewValue%
End Property

Public Property Get AbholNr%()
AbholNr% = bek.AbholNr
End Property
Public Property Let AbholNr(ByVal vNewValue%)
bek.AbholNr = vNewValue%
End Property

Public Property Get zugeordnet$()
zugeordnet$ = bek.zugeordnet
End Property
Public Property Let zugeordnet(ByVal vNewValue$)
bek.zugeordnet = vNewValue$
End Property

Public Property Get zukontrollieren$()
zukontrollieren$ = bek.zukontrollieren
End Property
Public Property Let zukontrollieren(ByVal vNewValue$)
bek.zukontrollieren = vNewValue$
End Property

Public Property Get fixiert$()
fixiert$ = bek.fixiert
End Property
Public Property Let fixiert(ByVal vNewValue$)
bek.fixiert = vNewValue$
End Property

Public Property Get nochzukontrollieren$()
nochzukontrollieren$ = bek.nochzukontrollieren
End Property
Public Property Let nochzukontrollieren(ByVal vNewValue$)
bek.nochzukontrollieren = vNewValue$
End Property

Public Property Get actkontrolle(ind%) As Byte
actkontrolle = bek.actkontrolle(ind%)
End Property
Public Property Let actkontrolle(ind%, ByVal vNewValue As Byte)
bek.actkontrolle(ind%) = vNewValue
End Property

Public Property Get actzuordnung() As Byte
actzuordnung = bek.actzuordnung
End Property
Public Property Let actzuordnung(ByVal vNewValue As Byte)
bek.actzuordnung = vNewValue
End Property

Public Property Get poslag%()
poslag% = bek.poslag
End Property
Public Property Let poslag(ByVal vNewValue%)
bek.poslag = vNewValue%
End Property

Public Property Get AngebotInd() As Byte
AngebotInd = bek.AngebotInd
End Property
Public Property Let AngebotInd(ByVal vNewValue As Byte)
bek.AngebotInd = vNewValue
End Property

Public Property Get aktivlief() As Byte
aktivlief = bek.aktivlief
End Property
Public Property Let aktivlief(ByVal vNewValue As Byte)
bek.aktivlief = vNewValue
End Property

Public Property Get aktivind%()
aktivind% = bek.aktivind
End Property
Public Property Let aktivind(ByVal vNewValue%)
bek.aktivind = vNewValue%
End Property

Public Property Get beklaufnr&()
beklaufnr& = bek.beklaufnr
End Property
Public Property Let beklaufnr(ByVal vNewValue&)
bek.beklaufnr = vNewValue&
End Property

Public Property Get herst$()
herst$ = bek.herst
End Property
Public Property Let herst(ByVal vNewValue$)
bek.herst = vNewValue$
End Property

Public Property Get LiefName$()
LiefName$ = bek.LiefName
End Property
Public Property Let LiefName(ByVal vNewValue$)
bek.LiefName = vNewValue$
End Property

Public Property Get dummy$()
dummy$ = bek.dummy
End Property
Public Property Let dummy(ByVal vNewValue$)
bek.dummy = vNewValue$
End Property

Public Property Get RawData$()
Dim temp$

temp$ = String(SATZLEN, 0)
CopyMemory ByVal temp$, bek, SATZLEN
RawData$ = temp$

End Property
