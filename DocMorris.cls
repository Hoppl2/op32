VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDocMorris"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Const SATZLEN = 200

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



Private Const DefErrModul = "DOCMORRIS.CLS"

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
dKurz$ = "DM"
dBezeichnung$ = "DocMorris"
dName$ = "dm.xxx"
dZugriff$ = "RW"
dModus$ = "R"
dSatzLen% = SATZLEN 'Len(lif)

dAnzFields% = 4
ReDim fKurz$(dAnzFields% - 1)
ReDim fBezeichnung$(dAnzFields% - 1)
ReDim fStart%(dAnzFields% - 1)
ReDim fLaenge%(dAnzFields% - 1)
ReDim fTyp$(dAnzFields% - 1)
ReDim fKonv%(dAnzFields% - 1)


fKurz$(0) = "PZN"
fBezeichnung$(0) = "Pharmazentralnummer"
fKurz$(1) = "VK"
fBezeichnung$(1) = "Verkaufspreis"
fKurz$(2) = "NichtAutomatenfaehig"
fBezeichnung$(2) = "Nicht Automatenfähig"
fKurz$(3) = "Gesperrt"
fBezeichnung$(3) = "Gesperrt"


iAnzIndex% = 0
iName$(0) = "nnaep.idx"

Set DM1 = Me

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
Set DM1 = Nothing
Call clsError.DefErrPop
End Sub

'Function OpenDatei%(Optional DateiZugriff$ = "")
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Call clsError.DefErrFnc("OpenDatei%")
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
'Dim i%
'Dim iZugriff$
'
'iZugriff$ = RTrim$(DateiZugriff$)
'If (iZugriff$ = "") Then
'    iZugriff$ = RTrim$(dZugriff$)
'End If
'
'If (RTrim$(dName$) <> "") Then
'    dHandle% = clsDat.FileOpen(RTrim$(dName$), iZugriff$, dModus$, dSatzLen%)
'    If (dHandle% = 0) Then
'        OpenDatei% = 1
'        Call clsError.DefErrPop: Exit Function
'    End If
'End If
'For i% = 0 To (iAnzIndex% - 1)
'     FabsErrf% = clsDat.IndexOpen(iName$(i%), iHandle%(i%))
'     If (FabsErrf%) Then
'        OpenDatei% = 1
'        Call clsError.DefErrPop: Exit Function
'    End If
'Next i%
'
'OpenDatei% = 0
'
'Call clsError.DefErrPop
'End Function
'
'Sub CloseDatei()
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Call clsError.DefErrFnc("CloseDatei")
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
'Dim i%
'
'If (dHandle% > 0) Then
'    Close #dHandle%
'End If
'For i% = 0 To (iAnzIndex% - 1)
'     FabsErrf% = clsDat.IndexClose(iHandle%(i%))
'Next i%
'
'Call clsError.DefErrPop
'End Sub
'
'Sub SatzLock(recno&)
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Call clsError.DefErrFnc("SatzLock")
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
'Call clsDat.iLock(dHandle%, recno&)
'Call clsError.DefErrPop
'End Sub
'
'Sub SatzUnLock(recno&)
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Call clsError.DefErrFnc("SatzUnLock")
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
'Call clsDat.iUnLock(dHandle%, recno&)
'Call clsError.DefErrPop
'End Sub
'
'Sub GetFirstRecord()
'Dim LockTime As Date
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Call clsError.DefErrFnc("GetFirstRecord")
'Call clsError.DefErrMod(DefErrModul)
'On Error GoTo DefErr
'GoTo DefErrEnd
'DefErr:
'
'If Err = 70 Or Err = 75 Then
'  If LockTime = 0 Then LockTime = DateAdd("s", 20, Now)
'  If LockTime > Now Then
'    'Sleep (1)
'    Resume
'  End If
'End If
'
'Select Case clsError.DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
'Case vbRetry
'  LockTime = 0
'  Resume
'Case vbIgnore
'  Resume Next
'End Select
'Call ProjektForm.EndeDll
'DefErrEnd:
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'
'Get #dHandle%, 1, nnekErst
'
'Call clsError.DefErrPop
'End Sub
'
'Sub GetRecord(Optional recno& = 0)
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Call clsError.DefErrFnc("GetRecord")
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
'If (recno& > 0) Then
'    Seek #dHandle%, recno&
'End If
'If (recno& = 1) Then
'    Call GetFirstRecord
'Else
'    Get #dHandle%, , nnek
'End If
'Call clsError.DefErrPop
'End Sub
'
'Sub PutRecord(Optional recno& = 0)
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Call clsError.DefErrFnc("PutRecord")
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
'If (recno& > 0) Then
'    Seek #dHandle%, recno&
'End If
'If (recno& = 1) Then
'    Put #dHandle%, , nnekErst
'Else
'    Put #dHandle%, , nnek
'End If
'Call clsError.DefErrPop
'End Sub
'
'Sub GetInfoRecord(Optional recno& = 0)
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Call clsError.DefErrFnc("GetInfoRecord")
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
'If (recno& > 0) Then
'    Seek #dHandle%, recno&
'End If
'Get #dHandle%, , AktData
'Call clsError.DefErrPop
'End Sub
'
'Function IndexSearch%(ind%, key$, recno&)
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Call clsError.DefErrFnc("IndexSearch%")
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
'FabsErrf% = clsDat.IndexSearch%(iHandle%(ind%), key$, recno&)
'iFabsRecno& = recno&
'IndexSearch% = FabsErrf%
'Call clsError.DefErrPop
'End Function
'
'Function IndexGeneric%(ind%, key$, recno&)
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Call clsError.DefErrFnc("IndexGeneric%")
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
'FabsErrf% = clsDat.IndexGeneric%(iHandle%(ind%), key$, recno&)
'iFabsRecno& = recno&
'IndexGeneric% = FabsErrf%
'Call clsError.DefErrPop
'End Function
'
'Function IndexNext%(ind%, AltRecno&, key$, recno&)
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Call clsError.DefErrFnc("IndexNext%")
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
'FabsErrf% = clsDat.IndexNext%(iHandle%(ind%), AltRecno&, key$, recno&)
'iFabsRecno& = recno&
'IndexNext% = FabsErrf%
'Call clsError.DefErrPop
'End Function
'
'Function IndexPrevious%(ind%, AltRecno&, key$, recno&)
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Call clsError.DefErrFnc("IndexPrevious%")
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
'FabsErrf% = clsDat.IndexPrevious%(iHandle%(ind%), AltRecno&, key$, recno&)
'iFabsRecno& = recno&
'IndexPrevious% = FabsErrf%
'Call clsError.DefErrPop
'End Function
'
'Function IndexInsert%(ind%, key$, recno&)
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Call clsError.DefErrFnc("IndexInsert%")
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
'FabsErrf% = clsDat.IndexInsert%(iHandle%(ind%), key$, recno&)
'iFabsRecno& = recno&
'IndexInsert% = FabsErrf%
'Call clsError.DefErrPop
'End Function
'
'Function IndexDelete%(ind%, AltRecno&, key$, recno&)
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Call clsError.DefErrFnc("IndexDelete%")
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
'FabsErrf% = clsDat.IndexDelete%(iHandle%(ind%), AltRecno&, key$, recno&)
'iFabsRecno& = recno&
'IndexDelete% = FabsErrf%
'Call clsError.DefErrPop
'End Function
'
'Function IndexReplace%(ind%, AltRecno&, AltKey$, key$, recno&)
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Call clsError.DefErrFnc("IndexReplace%")
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
'FabsErrf% = clsDat.IndexReplace%(iHandle%(ind%), AltRecno&, AltKey$, key$, recno&)
'iFabsRecno& = recno&
'IndexReplace% = FabsErrf%
'Call clsError.DefErrPop
'End Function
'
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
'
'Function Info$(iFeld$)
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Call clsError.DefErrFnc("Info$")
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
'Dim iValue As Variant
'Dim ind%
'Dim ret$, s$
'
'ret$ = ""
'
's$ = UCase$(iFeld$)
''If (iIndex% >= 0) Then
''    s$ = s$ + "(" + Mid$(Str$(iIndex%), 2) + ")"
''End If
'
'
'ind% = SucheFeldInd%(0, s$)
'If (ind% >= 0) Then
'    iValue = clsDat.FeldValue(Mid$(AktData$, fStart%(ind%), fLaenge%(ind%)), fTyp$(ind%), fKonv%(ind%))
'    ret$ = clsDat.DecodFeld(iValue, fTyp$(ind%), fKonv%(ind%))
'End If
'
'Info$ = ret$
'
'Call clsError.DefErrPop
'End Function
'
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
'
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


'Public Property Get erstmax&()
'erstmax& = CLng(nnekErst.erstmax) - (nnekErst.erstmax < 0) * 65536
'End Property
'Public Property Let erstmax(ByVal vNewValue&)
'nnekErst.erstmax = Int(vNewValue& + (vNewValue& > 32767) * 65536)
'End Property
'
'Public Property Get erstrest$()
'erstrest$ = nnekErst.erstrest
'End Property
'Public Property Let erstrest(ByVal vNewValue$)
'nnekErst.erstrest = vNewValue$
'End Property
'
'
'
'Public Property Get pzn$()
'pzn$ = nnek.pzn
'End Property
'Public Property Let pzn(ByVal vNewValue$)
'nnek.pzn = vNewValue$
'End Property
'
'Public Property Get BelegDat%(ByVal ind%)
'BelegDat% = clsOpTool.CVDat(nnek.lieferung(ind%).BelegDat)
'End Property
'Public Property Let BelegDat(ByVal ind%, ByVal vNewValue%)
'nnek.lieferung(ind%).BelegDat = clsOpTool.MKDate(vNewValue%)
'End Property
'
'Public Property Get NNAEP#(ByVal ind%)
'Dim a#
'a# = nnek.lieferung(ind%).NNAEP
'Call DxToIEEEd(a#)
'NNAEP# = a#
'End Property
'Public Property Let NNAEP(ByVal ind%, ByVal vNewValue#)
'Dim a#
'a# = vNewValue#
'Call DxToMBFd(a#)
'nnek.lieferung(ind%).NNAEP = a#
'End Property
'
'Public Property Get art%(ByVal ind%)
'art% = nnek.lieferung(ind%).art
'End Property
'Public Property Let art(ByVal ind%, ByVal vNewValue%)
'nnek.lieferung(ind%).art = vNewValue%
'End Property
'
'Public Property Get Beleg$(ByVal ind%)
'Beleg$ = nnek.Beleg(ind%)
'End Property
'Public Property Let Beleg(ByVal ind%, ByVal vNewValue$)
'nnek.Beleg(ind%) = vNewValue$
'End Property
'
'Public Property Get lm%(ByVal ind%)
'lm% = nnek.lm(ind%)
'End Property
'Public Property Let lm(ByVal ind%, ByVal vNewValue%)
'nnek.lm(ind%) = vNewValue%
'End Property
'
'Public Property Get rm%(ByVal ind%)
'rm% = nnek.rm(ind%)
'End Property
'Public Property Let rm(ByVal ind%, ByVal vNewValue%)
'nnek.rm(ind%) = vNewValue%
'End Property
'
'Public Property Get LiefNr(ByVal ind%) As Byte
'LiefNr = nnek.LiefNr(ind%)
'End Property
'Public Property Let LiefNr(ByVal ind%, ByVal vNewValue As Byte)
'nnek.LiefNr(ind%) = vNewValue
'End Property
'
'Public Property Get GhNr(ByVal ind%) As Byte
'GhNr = nnek.GhNr(ind%)
'End Property
'Public Property Let GhNr(ByVal ind%, ByVal vNewValue As Byte)
'nnek.GhNr(ind%) = vNewValue
'End Property
'
'Public Property Get RetMenge%(ByVal ind%)
'Dim ret%
'
'ret = 0
'If (nnek.RetMenge(ind) >= 1000) And (nnek.RetMenge(ind) <= 2000) Then
'    ret = nnek.RetMenge(ind) - 1000
'End If
'
'RetMenge = ret
'End Property
'Public Property Let RetMenge(ByVal ind%, ByVal vNewValue%)
'nnek.RetMenge(ind) = vNewValue + 1000
'End Property
'
'Public Property Get rest$()
'rest$ = nnek.rest
'End Property
'Public Property Let rest(ByVal vNewValue$)
'nnek.rest = vNewValue$
'End Property


Function FeldWert$(FeldName$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("FeldWert$")
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
Dim ret$, h$, iFeldName$
Dim fldRecordset As ADODB.Field

ret$ = ""

iFeldName$ = UCase(FeldName$)
Set fldRecordset = DocMorrisRec.Fields(FeldName$)
If (IsNull(fldRecordset)) Then
    ret$ = ""
Else
    If (fldRecordset.Type = dbDate) Then
        ret$ = Format(clsOpTool.CheckNullDouble(fldRecordset.Value), "DD.MM.YYYY")
'    ElseIf (fldRecordset.Type = dbDouble) Or (fldRecordset.Type = dbFloat) Then
    ElseIf (fldRecordset.Type = adNumeric) Then
        ret$ = Format(clsOpTool.CheckNullDouble(fldRecordset.Value), "0.00")
    ElseIf (fldRecordset.Type = adInteger) Or (fldRecordset.Type = adTinyInt) Or (fldRecordset.Type = adSmallInt) Or (fldRecordset.Type = adUnsignedTinyInt) Then
        ret$ = Format(clsOpTool.CheckNullDouble(fldRecordset.Value), "0")
    Else
        ret$ = LTrim$(clsOpTool.CheckNullStr(fldRecordset.Value))
    End If
    
'    Select Case iFeldName$
'        Case "PZN"
'            ret$ = clsOpTool.PznString(fldRecordset.Value)
'        Case "EK", "VK", "FESTBETRAG", "KVA", "VK", "VKALT", "HAP"
'            ret$ = Format(fldRecordset.Value / 100#, "0.00")
'            If (ret$ = "0.00") Or (ret$ = "0,00") Then
'                ret$ = "0"
'            End If
'        Case Else
'            ret$ = LTrim$(fldRecordset.Value)
'    End Select
End If

FeldWert$ = " " + ret$

Call clsError.DefErrPop
End Function


