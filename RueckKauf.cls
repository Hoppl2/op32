VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRueckKauf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Const SATZLEN = 256

Private Type WawiDatErstSatz
    max As Integer
    lief As Integer
    Counter As Byte
    rest As String * 251
End Type
Private Type WawiDatStruct
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
    DirektTyp As Byte
    actkontrolle(5) As Byte
    actzuordnung As Byte
    poslag As Integer
    AngebotInd As Byte
    aktivlief As Byte
    aktivind As Integer
    beklaufnr As Long
    herst As String * 5
    LiefName As String * 6
    LoeschFlag As Byte
    Status As Byte
    
    WuAEP As Double   'String * 8
    WuBestDatum As String * 6
    WuBm As Integer
    WuNm As Integer
    WuStat As String * 1
    WuRm As Integer
    WuLm As Integer
    WuAm As Integer
    WuAblDatum As String * 6
    WuLa As Integer
    WuBestZeit As String * 2  'Integer
    WuAVP As Double   'String * 8
    WuBelegDatum As Integer
    WuBeleg As String * 10
    WuRetMenge As Integer
    WuNNaepOk As Byte
    WuNNAep As Double
    WuNNart As Byte   'String * 1
    WuNeuLm As Integer
    WuNeuRm As Integer
    WuText As String * 30
    WuNeuZiel As Integer
    
    IstAltLast As Byte
    WuStatus As Byte
    LmStatus As Byte
    RmStatus As Byte
    LmAnzGebucht As Integer
    RmAnzGebucht As Integer
    
    tplatz As Integer
    tlager As Integer
    
    ArtikelKz As Byte
    
    IstSchwellArtikel As Byte
    OrgZeit As Integer
    
    BestellSupportId As Long
    
    PznStelle8Flag As String * 1
    PznStelle8 As String * 1
    dummy As String * 2
End Type
Dim wwErst As WawiDatErstSatz
Dim ww As WawiDatStruct

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



Private Const DefErrModul = "RUECKKAUF.CLS"

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
dKurz$ = "RUECKKAUF"
dBezeichnung$ = "Rückkauf-Anfragen"
dName$ = "ruekkauf.dat"
dZugriff$ = "RW"
dModus$ = "R"
dSatzLen% = SATZLEN 'Len(bek)

dAnzFields% = 48
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

fKurz$(28) = "DIREKTTYP"
fBezeichnung$(28) = "DirektTyp"
fStart%(28) = 99
fLaenge%(28) = 1
fTyp$(28) = "N"
fKonv%(28) = 1

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

fKurz$(42) = "LOESCH"   '????????
fBezeichnung$(42) = "LoeschFlag"
fStart%(42) = 43
fLaenge%(42) = 1
fTyp$(42) = "N"
fKonv%(42) = 1

fKurz$(43) = "ERSTMAX"
fBezeichnung$(43) = "ErstMax"
fStart%(43) = 1
fLaenge%(43) = 2
fTyp$(43) = "N"
fKonv%(43) = 5

fKurz$(44) = "ERSTLIEF"
fBezeichnung$(44) = "ErstLieferant"
fStart%(44) = 3
fLaenge%(44) = 2
fTyp$(44) = "N"
fKonv%(44) = 5

fKurz$(45) = "ERSTCOUNTER"
fBezeichnung$(45) = "ErstCounter"
fStart%(45) = 5
fLaenge%(45) = 1
fTyp$(45) = "N"
fKonv%(45) = 1

fKurz$(46) = "ERSTREST"
fBezeichnung$(46) = "ErstRest"
fStart%(46) = 6
fLaenge%(46) = 123
fTyp$(46) = "T"
fKonv%(46) = 0


fKurz$(47) = "ALLES"
fBezeichnung$(47) = "Alles"
fStart%(47) = 1
fLaenge%(47) = SATZLEN
fTyp$(47) = "T"
fKonv%(47) = 0




iAnzIndex% = 0


'Set Ww1 = Me

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
'Set Ww1 = Nothing
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
Call clsError.DefErrFnc("SatzLock(" + Mid$(Str$(recno&), 2) + ")")
Call clsError.DefErrMod(DefErrModul)
Call clsError.ErrorDebug
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
Call clsError.ErrorDebug("OK")
Call clsError.DefErrPop
End Sub

Sub SatzUnLock(recno&)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("SatzUnLock(" + Mid$(Str$(recno&), 2) + ")")
Call clsError.DefErrMod(DefErrModul)
Call clsError.ErrorDebug
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
Call clsError.ErrorDebug("OK")
Call clsError.DefErrPop
End Sub

Sub GetFirstRecord()
Dim LockTime As Date
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("GetFirstRecord")
Call clsError.DefErrMod(DefErrModul)
Call clsError.ErrorDebug
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

Get #dHandle%, 1, wwErst
Call clsError.ErrorDebug("OK")

Call clsError.DefErrPop
End Sub

'Sub GetRecord(Optional Recno& = 0)
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Call clsError.DefErrFnc("GetRecord(" + Mid$(Str$(Recno&), 2) + ")")
'Call clsError.DefErrMod(DefErrModul)
'On Error GoTo DefErr
'GoTo DefErrEnd
'DefErr:
'Select Case clsError.DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
'Case vbRetry
'  Resume
'Case vbIgnore
'  Resume Next
'End Select
'Call ProjektForm.EndeDll
'DefErrEnd:
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'If (Recno& > 0) Then
'    Seek #dHandle%, Recno&
'End If
'If (Recno& = 1) Then
'    Call GetFirstRecord
'Else
'    Get #dHandle%, , ww
'End If
'Call clsError.DefErrPop
'End Sub

Sub GetRecord(Optional recno& = 0)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("GetRecord(" + Mid$(Str$(recno&), 2) + ")")
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
If (recno& = 1) Then
    Call GetFirstRecord
ElseIf (recno& > 0) Then
    Get #dHandle%, recno&, ww
Else
    Get #dHandle%, , ww
End If
Call clsError.DefErrPop
End Sub

Sub PutRecord(Optional recno& = 0)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("PutRecord(" + Mid$(Str$(recno&), 2) + ")")
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
    Put #dHandle%, , wwErst
Else
    Put #dHandle%, , ww
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

Public Property Get DateiName$()
DateiName$ = dName$
End Property
Public Property Let DateiName(ByVal vNewValue$)
dName$ = vNewValue$
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
erstmax% = wwErst.max
End Property
Public Property Let erstmax(ByVal vNewValue%)
wwErst.max = vNewValue%
End Property

Public Property Get erstlief%()
erstlief% = wwErst.lief
End Property
Public Property Let erstlief(ByVal vNewValue%)
wwErst.lief = vNewValue%
End Property

Public Property Get erstcounter() As Byte
erstcounter = wwErst.Counter
End Property
Public Property Let erstcounter(ByVal vNewValue As Byte)
wwErst.Counter = vNewValue
End Property

Public Property Get erstrest$()
erstrest$ = wwErst.rest
End Property
Public Property Let erstrest(ByVal vNewValue$)
wwErst.rest = vNewValue$
End Property

'Public Property Get pzn$()
'pzn$ = "0" + ww.pzn
'End Property
'Public Property Let pzn(ByVal vNewValue$)
'If (Len(vNewValue) = 8) Then
'    ww.pzn = Mid(vNewValue$, 2)
'Else
'    ww.pzn = vNewValue$
'End If
'End Property
Public Property Get pzn$()
If (ww.PznStelle8Flag = "@") Then
    pzn = ww.pzn + ww.PznStelle8
Else
    pzn$ = "0" + ww.pzn
End If
End Property
Public Property Let pzn(ByVal vNewValue$)
If (Len(vNewValue) = 8) Then
    ww.pzn = Left(vNewValue$, 7)
    ww.PznStelle8 = Mid(vNewValue, 8, 1)
    ww.PznStelle8Flag = "@"
Else
    ww.pzn = vNewValue$
End If
End Property

Public Property Get txt$()
txt$ = ww.txt
End Property
Public Property Let txt(ByVal vNewValue$)
ww.txt = vNewValue$
End Property

Public Property Get lief() As Byte
lief = ww.lief
End Property
Public Property Let lief(ByVal vNewValue As Byte)
ww.lief = vNewValue
End Property

Public Property Get bm%()
bm% = ww.bm
End Property
Public Property Let bm(ByVal vNewValue%)
ww.bm = vNewValue%
End Property

Public Property Get asatz%()
asatz% = ww.asatz
End Property
Public Property Let asatz(ByVal vNewValue%)
ww.asatz = vNewValue%
End Property

Public Property Get ssatz%()
ssatz% = ww.ssatz
End Property
Public Property Let ssatz(ByVal vNewValue%)
ww.ssatz = vNewValue%
End Property

Public Property Get best$()
best$ = ww.best
End Property
Public Property Let best(ByVal vNewValue$)
ww.best = vNewValue$
End Property

Public Property Get nm%()
nm% = ww.nm
End Property
Public Property Let nm(ByVal vNewValue%)
ww.nm = vNewValue%
End Property

Public Property Get aep#()
Dim a#
a# = ww.aep
Call DxToIEEEd(a#)
aep# = a#
End Property
Public Property Let aep(ByVal vNewValue#)
Dim a#
a# = vNewValue#
Call DxToMBFd(a#)
ww.aep = a#
End Property

Public Property Get abl$()
abl$ = ww.abl
End Property
Public Property Let abl(ByVal vNewValue$)
ww.abl = vNewValue$
End Property

Public Property Get wg$()
wg$ = ww.wg
End Property
Public Property Let wg(ByVal vNewValue$)
ww.wg = vNewValue$
End Property

Public Property Get AVP#()
Dim a#
a# = ww.AVP
Call DxToIEEEd(a#)
AVP# = a#
End Property
Public Property Let AVP(ByVal vNewValue#)
Dim a#
a# = vNewValue#
Call DxToMBFd(a#)
ww.AVP = a#
End Property

Public Property Get km%()
km% = ww.km
End Property
Public Property Let km(ByVal vNewValue%)
ww.km = vNewValue%
End Property

Public Property Get absage() As Byte
absage = ww.absage
End Property
Public Property Let absage(ByVal vNewValue As Byte)
If (vNewValue < 100) Then ww.absage = vNewValue
End Property

Public Property Get angebot() As Byte
angebot = ww.angebot
End Property
Public Property Let angebot(ByVal vNewValue As Byte)
ww.angebot = vNewValue
End Property

Public Property Get auto$()
auto$ = ww.auto
End Property
Public Property Let auto(ByVal vNewValue$)
ww.auto = vNewValue$
End Property

Public Property Get alt$()
alt$ = ww.alt
End Property
Public Property Let alt(ByVal vNewValue$)
ww.alt = vNewValue$
End Property

Public Property Get zr!()
zr! = ww.zr / 100!
End Property
Public Property Let zr(ByVal vNewValue!)
ww.zr = CInt(vNewValue! * 100)
End Property

Public Property Get besorger$()
besorger$ = ww.besorger
End Property
Public Property Let besorger(ByVal vNewValue$)
ww.besorger = vNewValue$
End Property

Public Property Get NNAEP#()
NNAEP# = ww.NNAEP
End Property
Public Property Let NNAEP(ByVal vNewValue#)
ww.NNAEP = vNewValue#
End Property

Public Property Get nnart() As Byte
nnart = ww.nnart
End Property
Public Property Let nnart(ByVal vNewValue As Byte)
ww.nnart = vNewValue
End Property

Public Property Get Lief1() As Byte
Lief1 = ww.Lief1
End Property
Public Property Let Lief1(ByVal vNewValue As Byte)
ww.Lief1 = vNewValue
End Property

Public Property Get bm1%()
bm1% = ww.bm1
End Property
Public Property Let bm1(ByVal vNewValue%)
ww.bm1 = vNewValue%
End Property

Public Property Get nm1%()
nm1% = ww.nm1
End Property
Public Property Let nm1(ByVal vNewValue%)
ww.nm1 = vNewValue%
End Property

Public Property Get AbholNr%()
AbholNr% = ww.AbholNr
End Property
Public Property Let AbholNr(ByVal vNewValue%)
ww.AbholNr = vNewValue%
End Property

Public Property Get zugeordnet$()
zugeordnet$ = ww.zugeordnet
End Property
Public Property Let zugeordnet(ByVal vNewValue$)
ww.zugeordnet = vNewValue$
End Property

Public Property Get zukontrollieren$()
zukontrollieren$ = ww.zukontrollieren
End Property
Public Property Let zukontrollieren(ByVal vNewValue$)
ww.zukontrollieren = vNewValue$
End Property

Public Property Get fixiert$()
fixiert$ = ww.fixiert
End Property
Public Property Let fixiert(ByVal vNewValue$)
ww.fixiert = vNewValue$
End Property

Public Property Get DirektTyp() As Byte
DirektTyp = ww.DirektTyp
End Property
Public Property Let DirektTyp(ByVal vNewValue As Byte)
ww.DirektTyp = vNewValue
End Property

Public Property Get actkontrolle(ind%) As Byte
actkontrolle = ww.actkontrolle(ind%)
End Property
Public Property Let actkontrolle(ind%, ByVal vNewValue As Byte)
ww.actkontrolle(ind%) = vNewValue
End Property

Public Property Get actzuordnung() As Byte
actzuordnung = ww.actzuordnung
End Property
Public Property Let actzuordnung(ByVal vNewValue As Byte)
ww.actzuordnung = vNewValue
End Property

Public Property Get poslag%()
poslag% = ww.poslag
End Property
Public Property Let poslag(ByVal vNewValue%)
ww.poslag = vNewValue%
End Property

Public Property Get AngebotInd() As Byte
AngebotInd = ww.AngebotInd
End Property
Public Property Let AngebotInd(ByVal vNewValue As Byte)
ww.AngebotInd = vNewValue
End Property

Public Property Get aktivlief() As Byte
aktivlief = ww.aktivlief
End Property
Public Property Let aktivlief(ByVal vNewValue As Byte)
ww.aktivlief = vNewValue
End Property

Public Property Get aktivind%()
aktivind% = ww.aktivind
End Property
Public Property Let aktivind(ByVal vNewValue%)
ww.aktivind = vNewValue%
End Property

Public Property Get beklaufnr&()
beklaufnr& = ww.beklaufnr
End Property
Public Property Let beklaufnr(ByVal vNewValue&)
ww.beklaufnr = vNewValue&
End Property

Public Property Get herst$()
herst$ = ww.herst
End Property
Public Property Let herst(ByVal vNewValue$)
ww.herst = vNewValue$
End Property

Public Property Get LiefName$()
LiefName$ = ww.LiefName
End Property
Public Property Let LiefName(ByVal vNewValue$)
ww.LiefName = vNewValue$
End Property

Public Property Get loesch() As Byte
loesch = ww.LoeschFlag
End Property
Public Property Let loesch(ByVal vNewValue As Byte)
ww.LoeschFlag = vNewValue
End Property

Public Property Get Status() As Byte
Status = ww.Status
End Property
Public Property Let Status(ByVal vNewValue As Byte)
ww.Status = vNewValue
End Property



Public Property Get WuAEP#()
Dim a#
a# = ww.WuAEP
Call DxToIEEEd(a#)
WuAEP# = a#
End Property
Public Property Let WuAEP(ByVal vNewValue#)
Dim a#
a# = vNewValue#
Call DxToMBFd(a#)
ww.WuAEP = a#
End Property

Public Property Get WuBestDatum$()
WuBestDatum$ = ww.WuBestDatum
End Property
Public Property Let WuBestDatum(ByVal vNewValue$)
ww.WuBestDatum = vNewValue$
End Property

Public Property Get WuBm%()
WuBm% = ww.WuBm
End Property
Public Property Let WuBm(ByVal vNewValue%)
ww.WuBm = vNewValue%
End Property

Public Property Get WuNm%()
WuNm% = ww.WuNm
End Property
Public Property Let WuNm(ByVal vNewValue%)
ww.WuNm = vNewValue%
End Property

Public Property Get WuStat$()
WuStat$ = ww.WuStat
End Property
Public Property Let WuStat(ByVal vNewValue$)
ww.WuStat = vNewValue$
End Property

Public Property Get WuRm%()
WuRm% = ww.WuRm
End Property
Public Property Let WuRm(ByVal vNewValue%)
ww.WuRm = vNewValue%
End Property

Public Property Get WuLm%()
WuLm% = ww.WuLm
End Property
Public Property Let WuLm(ByVal vNewValue%)
ww.WuLm = vNewValue%
End Property

Public Property Get WuAm%()
WuAm% = ww.WuAm
End Property
Public Property Let WuAm(ByVal vNewValue%)
ww.WuAm = vNewValue%
End Property

Public Property Get WuAblDatum$()
WuAblDatum$ = ww.WuAblDatum
End Property
Public Property Let WuAblDatum(ByVal vNewValue$)
ww.WuAblDatum = vNewValue$
End Property

Public Property Get WuLa%()
WuLa% = ww.WuLa
End Property
Public Property Let WuLa(ByVal vNewValue%)
ww.WuLa = vNewValue%
End Property

Public Property Get WuBestZeit$()
WuBestZeit$ = ww.WuBestZeit
End Property
Public Property Let WuBestZeit(ByVal vNewValue$)
ww.WuBestZeit = vNewValue$
End Property

Public Property Get WuAVP#()
Dim a#
a# = ww.WuAVP
Call DxToIEEEd(a#)
WuAVP# = a#
End Property
Public Property Let WuAVP(ByVal vNewValue#)
Dim a#
a# = vNewValue#
Call DxToMBFd(a#)
ww.WuAVP = a#
End Property

Public Property Get WuBelegDatum%()
WuBelegDatum% = ww.WuBelegDatum
End Property
Public Property Let WuBelegDatum(ByVal vNewValue%)
ww.WuBelegDatum = vNewValue%
End Property

Public Property Get WuBeleg$()
WuBeleg$ = ww.WuBeleg
End Property
Public Property Let WuBeleg(ByVal vNewValue$)
ww.WuBeleg = vNewValue$
End Property

Public Property Get WuRetMenge%()
WuRetMenge% = ww.WuRetMenge
End Property
Public Property Let WuRetMenge(ByVal vNewValue%)
ww.WuRetMenge = vNewValue%
End Property

'Public Property Get WuAnzBereitsGebucht%()
'WuAnzBereitsGebucht% = ww.WuAnzBereitsGebucht
'End Property
'Public Property Let WuAnzBereitsGebucht(ByVal vNewValue%)
'ww.WuAnzBereitsGebucht = vNewValue%
'End Property
'
'Public Property Get WuFertig() As Byte
'WuFertig = ww.WuFertig
'End Property
'Public Property Let WuFertig(ByVal vNewValue As Byte)
'ww.WuFertig = vNewValue
'End Property

Public Property Get WuNNaepOk() As Byte
WuNNaepOk = ww.WuNNaepOk
End Property
Public Property Let WuNNaepOk(ByVal vNewValue As Byte)
ww.WuNNaepOk = vNewValue
End Property

Public Property Get WuNNAep#()
WuNNAep# = ww.WuNNAep
End Property
Public Property Let WuNNAep(ByVal vNewValue#)
ww.WuNNAep = vNewValue#
End Property

Public Property Get WuNNart() As Byte
WuNNart = ww.WuNNart
End Property
Public Property Let WuNNart(ByVal vNewValue As Byte)
ww.WuNNart = vNewValue
End Property

Public Property Get WuNeuLm%()
WuNeuLm% = ww.WuNeuLm
End Property
Public Property Let WuNeuLm(ByVal vNewValue%)
ww.WuNeuLm = vNewValue%
End Property

Public Property Get WuNeuRm%()
WuNeuRm% = ww.WuNeuRm
End Property
Public Property Let WuNeuRm(ByVal vNewValue%)
ww.WuNeuRm = vNewValue%
End Property

Public Property Get WuText$()
WuText$ = ww.WuText
End Property
Public Property Let WuText(ByVal vNewValue$)
ww.WuText = vNewValue$
End Property

Public Property Get WuNeuZiel%()
WuNeuZiel% = ww.WuNeuZiel
End Property
Public Property Let WuNeuZiel(ByVal vNewValue%)
ww.WuNeuZiel = vNewValue%
End Property


Public Property Get tplatz%()
tplatz% = ww.tplatz
End Property
Public Property Let tplatz(ByVal vNewValue%)
ww.tplatz = vNewValue%
End Property


Public Property Get tlager%()
tlager% = ww.tlager
End Property
Public Property Let tlager(ByVal vNewValue%)
ww.tlager = vNewValue%
End Property


Public Property Get IstAltLast() As Byte
IstAltLast = ww.IstAltLast
End Property
Public Property Let IstAltLast(ByVal vNewValue As Byte)
ww.IstAltLast = vNewValue
End Property

Public Property Get WuStatus() As Byte
WuStatus = ww.WuStatus
End Property
Public Property Let WuStatus(ByVal vNewValue As Byte)
ww.WuStatus = vNewValue
End Property

Public Property Get LmStatus() As Byte
LmStatus = ww.LmStatus
End Property
Public Property Let LmStatus(ByVal vNewValue As Byte)
ww.LmStatus = vNewValue
End Property

Public Property Get RmStatus() As Byte
RmStatus = ww.RmStatus
End Property
Public Property Let RmStatus(ByVal vNewValue As Byte)
ww.RmStatus = vNewValue
End Property

Public Property Get LmAnzGebucht%()
LmAnzGebucht% = ww.LmAnzGebucht
End Property
Public Property Let LmAnzGebucht(ByVal vNewValue%)
ww.LmAnzGebucht = vNewValue%
End Property

Public Property Get RmAnzGebucht%()
RmAnzGebucht% = ww.RmAnzGebucht
End Property
Public Property Let RmAnzGebucht(ByVal vNewValue%)
ww.RmAnzGebucht = vNewValue%
End Property


Public Property Get ArtikelKz() As Byte
ArtikelKz = ww.ArtikelKz
End Property
Public Property Let ArtikelKz(ByVal vNewValue As Byte)
ww.ArtikelKz = vNewValue
End Property


Public Property Get IstSchwellArtikel() As Byte
IstSchwellArtikel = ww.IstSchwellArtikel
End Property
Public Property Let IstSchwellArtikel(ByVal vNewValue As Byte)
ww.IstSchwellArtikel = vNewValue
End Property

Public Property Get OrgZeit%()
OrgZeit% = ww.OrgZeit
End Property
Public Property Let OrgZeit(ByVal vNewValue%)
ww.OrgZeit = vNewValue%
End Property

Public Property Get BestellSupportId&()
BestellSupportId& = ww.BestellSupportId
End Property
Public Property Let BestellSupportId(ByVal vNewValue&)
ww.BestellSupportId = vNewValue
End Property


Public Property Get dummy$()
dummy$ = ww.dummy
End Property
Public Property Let dummy(ByVal vNewValue$)
ww.dummy = vNewValue$
End Property

Public Property Get RawData$()
Dim temp$

temp$ = String(SATZLEN, 0)
CopyMemory ByVal temp$, ww, SATZLEN
RawData$ = temp$

End Property
