VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsVerkauf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Const SATZLEN = 128

Private Type VerkaufErstSatz
    erstmax As Single
    erstrest As String * 98
End Type

Private Type VerkaufStruct
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
  angefordert As Byte       'String * 1
  FremdGeld   As Byte
  FremdBetrag As String * 8
  TxtTyp      As String * 1
  RezEan      As String * 13
  KKNr        As String * 9
  KKTyp       As Byte
  ZuzaFlag    As String * 1
  Zustellung As Byte
  RezeptArt   As Byte
  Storno      As Byte
  rest        As String * 6
  multi       As Byte
End Type
Dim vk As VerkaufStruct
Dim vkErst As VerkaufErstSatz

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



Private Const DefErrModul = "VERKAUF.CLS"

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
dKurz$ = "VK"
dBezeichnung$ = "Verkaufsdaten"
dName$ = "verkauf.dat"
dZugriff$ = "R"
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


iAnzIndex% = 0

Set Vk1 = Me

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
Set Vk1 = Nothing
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

Get #dHandle%, 1, vkErst

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
    Get #dHandle%, , vk
End If
Call clsError.DefErrPop
End Sub

Function DatumSuche&(search$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("DatumSuche&")
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

'Suche nach Datum

Dim Compare$
Dim xrec&, xLeft&, xRight&, found&
Dim zugriffe%
Dim Grösser As Boolean

Compare$ = Space$(Len(search$))
found& = 0
xrec& = 0
xLeft& = 1

'xRight& = DateiLen& / RecordLen%
xRight& = erstmax!
zugriffe% = 0
While (xLeft& <= xRight&) And (found& = 0)
  zugriffe% = zugriffe% + 1
  xrec& = Int((xLeft& + xRight&) / 2)
  Call GetRecord(xrec&)
  Grösser = False
  If Asc(Left(search$, 1)) > Asc(Left(vk.datum, 1)) Then
    Grösser = True
  ElseIf Asc(Left(search$, 1)) = Asc(Left(vk.datum, 1)) And Asc(Mid(search$, 2, 1)) > Asc(Mid(vk.datum, 2, 1)) Then
    Grösser = True
  End If
  If Grösser Then
    xLeft& = xrec& + 1
  ElseIf search$ = vk.datum Then
    'nun den ersten passenden Satz suchen
    xRight& = xrec&
    If xLeft& = xrec& Then found& = xrec&
  Else
    xRight& = xrec& - 1
  End If
Wend
If found& = 0 Then found& = -xrec&
DatumSuche& = found&
Call clsError.DefErrPop
End Function



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
    Put #dHandle%, , vkErst
Else
    Put #dHandle%, , vk
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






Public Property Get erstmax!()
Dim a!
a! = vkErst.erstmax
Call DxToIEEEs(a!)
erstmax! = a!
End Property
Public Property Let erstmax(ByVal vNewValue!)
Dim a!
a! = vNewValue!
Call DxToMBFs(a!)
vkErst.erstmax = a!
End Property

Public Property Get erstrest$()
erstrest$ = vkErst.erstrest
End Property
Public Property Let erstrest(ByVal vNewValue$)
vkErst.erstrest = vNewValue$
End Property



Public Property Get RezNr%()
RezNr% = vk.RezNr
End Property
Public Property Let RezNr(ByVal vNewValue%)
vk.RezNr = vNewValue%
End Property

Public Property Get EndeT() As Byte
EndeT = vk.EndeT
End Property
Public Property Let EndeT(ByVal vNewValue As Byte)
vk.EndeT = vNewValue
End Property

Public Property Get pzn$()
pzn$ = vk.pzn
End Property
Public Property Let pzn(ByVal vNewValue$)
vk.pzn = vNewValue$
End Property

Public Property Get text$()
text$ = vk.text
End Property
Public Property Let text(ByVal vNewValue$)
vk.text = vNewValue$
End Property

Public Property Get Preis$()
Preis$ = vk.Preis
End Property
Public Property Let Preis(ByVal vNewValue$)
vk.Preis = vNewValue$
End Property

Public Property Get mw$()
mw$ = vk.mw
End Property
Public Property Let mw(ByVal vNewValue$)
vk.mw = vNewValue$
End Property

Public Property Get bs() As Byte
bs = vk.bs
End Property
Public Property Let bs(ByVal vNewValue As Byte)
vk.bs = vNewValue
End Property

Public Property Get datum%()
datum% = clsOpTool.CVDat(vk.datum)
End Property
Public Property Let datum(ByVal vNewValue%)
vk.datum = clsOpTool.MKDate(vNewValue%)
End Property

Public Property Get user%()
user% = vk.user
End Property
Public Property Let user(ByVal vNewValue%)
vk.user = vNewValue%
End Property

Public Property Get LaufNr%()
LaufNr% = vk.LaufNr
End Property
Public Property Let LaufNr(ByVal vNewValue%)
vk.LaufNr = vNewValue%
End Property

Public Property Get gebuehren() As Byte
gebuehren = vk.gebuehren
End Property
Public Property Let gebuehren(ByVal vNewValue As Byte)
vk.gebuehren = vNewValue
End Property

Public Property Get wegmark$()
wegmark$ = vk.wegmark
End Property
Public Property Let wegmark(ByVal vNewValue$)
vk.wegmark = vNewValue$
End Property

Public Property Get AVP() As Byte
AVP = vk.AVP
End Property
Public Property Let AVP(ByVal vNewValue As Byte)
vk.AVP = vNewValue
End Property

Public Property Get zeit%()
zeit% = vk.zeit
End Property
Public Property Let zeit(ByVal vNewValue%)
vk.zeit = vNewValue%
End Property

Public Property Get bon$()
bon$ = vk.bon
End Property
Public Property Let bon(ByVal vNewValue$)
vk.bon = vNewValue$
End Property

Public Property Get wg$()
wg$ = vk.wg
End Property
Public Property Let wg(ByVal vNewValue$)
vk.wg = vNewValue$
End Property

Public Property Get knr%()
knr% = vk.knr
End Property
Public Property Let knr(ByVal vNewValue%)
vk.knr = vNewValue%
End Property

Public Property Get gebsumme#()
Dim a#
Call CopyMemory(a#, ByVal vk.gebsumme, 8)
Call DxToIEEEd(a#)
gebsumme# = a#
End Property
Public Property Let gebsumme(ByVal vNewValue#)
Dim a#
a# = vNewValue#
Call DxToMBFd(a#)
Call CopyMemory(ByVal vk.gebsumme, a#, 8)
End Property

Public Property Get pflag$()
pflag$ = vk.pflag
End Property
Public Property Let pflag(ByVal vNewValue$)
vk.pflag = vNewValue$
End Property

Public Property Get PersCode() As Byte
PersCode = vk.PersCode
End Property
Public Property Let PersCode(ByVal vNewValue As Byte)
vk.PersCode = vNewValue
End Property

Public Property Get angefordert() As Byte
angefordert = vk.angefordert
End Property
Public Property Let angefordert(ByVal vNewValue As Byte)
vk.angefordert = vNewValue
End Property

Public Property Get FremdGeld() As Byte
FremdGeld = vk.FremdGeld
End Property
Public Property Let FremdGeld(ByVal vNewValue As Byte)
vk.FremdGeld = vNewValue
End Property

Public Property Get FremdBetrag#()
Dim a#
Call CopyMemory(a#, ByVal vk.FremdBetrag, 8)
Call DxToIEEEd(a#)
FremdBetrag# = a#
End Property
Public Property Let FremdBetrag(ByVal vNewValue#)
Dim a#
a# = vNewValue#
Call DxToMBFd(a#)
Call CopyMemory(ByVal vk.FremdBetrag, a#, 8)
End Property

Public Property Get TxtTyp$()
TxtTyp$ = vk.TxtTyp
End Property
Public Property Let TxtTyp(ByVal vNewValue$)
vk.TxtTyp = vNewValue$
End Property

Public Property Get RezEan$()
RezEan$ = vk.RezEan
End Property
Public Property Let RezEan(ByVal vNewValue$)
vk.RezEan = vNewValue$
End Property

Public Property Get KKNr$()
KKNr$ = vk.KKNr
End Property
Public Property Let KKNr(ByVal vNewValue$)
vk.KKNr = vNewValue$
End Property

Public Property Get KKTyp() As Byte
KKTyp = vk.KKTyp
End Property
Public Property Let KKTyp(ByVal vNewValue As Byte)
vk.KKTyp = vNewValue
End Property

Public Property Get ZuzaFlag$()
ZuzaFlag$ = vk.ZuzaFlag
End Property
Public Property Let ZuzaFlag(ByVal vNewValue$)
vk.ZuzaFlag = vNewValue$
End Property

Public Property Get Zustellung() As Byte
Zustellung = vk.Zustellung
End Property
Public Property Let Zustellung(ByVal vNewValue As Byte)
vk.Zustellung = vNewValue
End Property

Public Property Get RezeptArt() As Byte
RezeptArt = vk.RezeptArt
End Property
Public Property Let RezeptArt(ByVal vNewValue As Byte)
vk.RezeptArt = vNewValue
End Property

Public Property Get Storno() As Byte
Storno = vk.Storno
End Property
Public Property Let Storno(ByVal vNewValue As Byte)
vk.Storno = vNewValue
End Property


Public Property Get rest$()
rest$ = vk.rest
End Property
Public Property Let rest(ByVal vNewValue$)
vk.rest = vNewValue$
End Property

Public Property Get multi() As Byte
multi = vk.multi
End Property
Public Property Let multi(ByVal vNewValue As Byte)
vk.multi = vNewValue
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
Put #Handle%, , vk
Call clsError.DefErrPop
End Sub
Public Property Get RawData$()
Dim temp$

temp$ = String(SATZLEN, 0)
CopyMemory ByVal temp$, vk, SATZLEN
RawData$ = temp$

End Property
Public Property Let RawData(ByVal vNewValue$)
CopyMemory vk, ByVal vNewValue$, SATZLEN
End Property

