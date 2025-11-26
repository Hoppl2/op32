VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsRabattTabellen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Const SATZLEN = 42

Dim RabTab As String * SATZLEN

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

Dim TabKZ%(6)
Dim TabWG%(9)
Dim TabHST$(5)
Dim TabDat1%(10)
Dim TabDat2%(10), TabDat3%(10)
Dim TabDat4#(10), TabDat5!(10)
Dim TabBM#, TabZW#
Dim AktRab!


Dim LRabTab%, LEKRabZ%, LNoRabTab%, UBRabTab%
Dim TabSatz%, ePos%
Dim iTabOk%

Dim lSchwellwert#(9)
Dim lRabatt!(9)

Dim AktLieferant%



Private Const DefErrModul = "RABTAB.CLS"

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
dKurz$ = "RABTAB"
dBezeichnung$ = "Rabatt-Tabellen"
dName$ = "lirabtab.dat"
dZugriff$ = "R"
dModus$ = "R"
dSatzLen% = SATZLEN 'Len(lif)

dAnzFields% = 1
ReDim fKurz$(dAnzFields% - 1)
ReDim fBezeichnung$(dAnzFields% - 1)
ReDim fStart%(dAnzFields% - 1)
ReDim fLaenge%(dAnzFields% - 1)
ReDim fTyp$(dAnzFields% - 1)
ReDim fKonv%(dAnzFields% - 1)

fKurz$(0) = "Eintrag"
fBezeichnung$(0) = "Eintrag"
fStart%(0) = 1
fLaenge%(0) = SATZLEN
fTyp$(0) = "T"
fKonv%(0) = 0

iAnzIndex% = 0

Set RabTab1 = Me

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
Set RabTab1 = Nothing
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
Get #dHandle%, , RabTab
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
Put #dHandle%, , RabTab
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


Public Property Get eintrag$()
eintrag$ = RabTab
End Property
Public Property Let eintrag(ByVal vNewValue$)
RabTab = vNewValue$
End Property


Public Property Get TabOk%()
TabOk% = iTabOk%
End Property

Sub aTabelleEinlesen(Lieferant%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("aTabelleEinlesen")
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
Dim i%, j&

iTabOk% = False

AktLieferant% = Lieferant%

For i% = 0 To 6: TabKZ%(i%) = 0: Next i%
For i% = 0 To 9: TabWG%(i%) = 0: Next i%
For i% = 0 To 5: TabHST$(i%) = "": Next i%
For i% = 0 To 10  '127!!!!!!!!!!!
    TabDat1%(i%) = 0
    TabDat2%(i%) = 0
    TabDat3%(i%) = 0
    TabDat4#(i%) = 0#
    TabDat5!(i%) = 0!
Next i%

LRabTab% = 0: LEKRabZ% = 0: LNoRabTab% = 0

If (InStr(Para1.Benutz, "#") > 0) Or (InStr(Para1.Benutz, "&") > 0) Then
    iTabOk% = True

    Call OpenDatei
    If (LOF(dHandle%) = 0) Then
        RabTab = String$(SATZLEN, 0)
        For j& = 1 To 156
            Call PutRecord(j&)
        Next j&
    End If
    
    
    If (Lieferant% > 5) Then
        TabSatz% = Int((Lieferant% - 6) / 7) + 2
        ePos% = ((Lieferant% - 6) Mod 7) * 6 + 1
    Else
        TabSatz% = 1
        ePos% = ((Lieferant% - 1) Mod 7) * 6 + 13
    End If
    
    Call GetRecord(CLng(TabSatz%))
    
    LRabTab% = Asc(Mid$(RabTab, ePos%, 1))
    If (LRabTab% < 1) Then
        Call clsDialog.MessageBox("Keine Rabattabelle für Lieferant" + Str$(Lieferant%) + " angegeben!")
        iTabOk% = False
    '    art% = 5
    End If
    LEKRabZ% = Asc(Mid$(RabTab, ePos% + 1, 1))
    LNoRabTab% = Asc(Mid$(RabTab, ePos% + 2, 1))
    If (LNoRabTab% < 1) Then
        Call clsDialog.MessageBox("Keine Ausnahmentabelle für Lieferant" + Str$(Lieferant%) + " angegeben!")
    End If
    
    If (LRabTab% > 0) And (LRabTab% < 100) Then Call EKWertTab
    If (LRabTab% > 99) And (LRabTab% < 199) Then Call BMZeilenTab
    If (LNoRabTab% <> 0) Then Call AusnahmenTab
    
    Call CloseDatei
End If

Call clsError.DefErrPop
End Sub

Sub TabelleEinlesen(Lieferant%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("TabelleEinlesen")
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

If (Para1.Land <> "D") Then Call clsError.DefErrPop: Exit Sub

iTabOk% = False

AktLieferant% = Lieferant%

'If (InStr(Para1.Benutz, "#") > 0) Or (InStr(Para1.Benutz, "&") > 0) Then
    iTabOk% = True

    LifZus1.GetRecord (Lieferant% + 1)
    
    If (LifZus1.Rabatt(0) = 0) And (LifZus1.BmEkRabatt(0) = 0) Then
        If (LifZus1.IstDirektLieferant) Then
        Else
            Call clsDialog.MessageBox("Keine Rabattabelle für Lieferant" + Str$(Lieferant%) + " angegeben!")
        End If
        iTabOk% = False
    '    art% = 5
    End If
    
'End If

Call clsError.DefErrPop
End Sub

Sub BMZeilenTab()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("BMZeilenTab")
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
Dim TabNr%, TabStrt%, TabMax%

TabNr% = LRabTab%
TabSatz% = Int((TabNr% - 1) / 21)
ePos% = ((TabNr% - 1) Mod 21) * 2 + 1
Call GetRecord(TabSatz% + 144)

TabStrt% = clsOpTool.CVI(Mid$(RabTab, ePos%, 2))
TabMax% = 0

While (TabStrt% > 0)
    Call GetRecord(TabStrt% + 156)
    TabStrt% = clsOpTool.CVI(Mid$(RabTab, 41, 2))
    
    TabDat1%(TabMax%) = clsOpTool.CVI(Mid$(RabTab, 3, 2))
    TabDat2%(TabMax%) = clsOpTool.CVI(Mid$(RabTab, 5, 2))
    TabDat3%(TabMax%) = Asc(Mid$(RabTab, 15, 1))
    TabDat4#(TabMax%) = clsOpTool.CVD(Mid$(RabTab, 7, 8))
    TabDat5!(TabMax%) = clsOpTool.CVS(Mid$(RabTab, 23, 4))
    TabMax% = TabMax% + 1
Wend
UBRabTab% = TabMax% - 1

Call clsError.DefErrPop
End Sub

Sub EKWertTab()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("EKWertTab")
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
Dim TabNr%, TabStrt%, TabMax%

If (LEKRabZ% = 0) Then
    Call clsDialog.MessageBox("Wert in EK-Wert-Rabattabelle nicht markiert!")
    iTabOk% = False
'    art% = 5
Else
    TabNr% = LRabTab%
    TabSatz% = Int((TabNr% - 1) / 21)
    ePos% = ((TabNr% - 1) Mod 21) * 2 + 1
    Call GetRecord(TabSatz% + 144)
    TabStrt% = clsOpTool.CVI(Mid$(RabTab, ePos%, 2))
    
    'EK-Werttabelle einlesen
    TabMax% = 0
    While (TabStrt% > 0)
        Call GetRecord(TabStrt% + 156)
        TabStrt% = clsOpTool.CVI(Mid$(RabTab, 41, 2))
        If (TabMax% = LEKRabZ% - 1) Then
            AktRab! = clsOpTool.CVS(Mid$(RabTab, 23, 4))
'            Call DxToIEEEs(AktRab!)
            TabStrt% = 0
        End If
        TabMax% = TabMax% + 1
    Wend
End If

Call clsError.DefErrPop
End Sub

Sub AusnahmenTab()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("AusnahmenTab")
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
Dim i%, TabNr%, TabStrt%, TabMax%, x%

TabNr% = LNoRabTab%
TabSatz% = Int((TabNr% - 1) / 21)
ePos% = ((TabNr% - 1) Mod 21) * 2 + 1

Call GetRecord(TabSatz% + 144)

TabStrt% = clsOpTool.CVI(Mid$(RabTab, ePos%, 2))
TabMax% = 0

While (TabStrt% > 0)
    Call GetRecord(TabStrt% + 156)
    TabStrt% = clsOpTool.CVI(Mid$(RabTab, 41, 2))
    If (TabMax% = 0) Then       'WGs & KZs
        For i% = 0 To 9
            TabWG%(i%) = Asc(Mid$(RabTab, 23 + i%, 1))
        Next i%
        x% = clsOpTool.CVI(Mid$(RabTab, 3, 2))
        For i% = 0 To 6
            If (x% And 2 ^ i%) > 0 Then TabKZ%(i%) = True
        Next i%
    ElseIf (TabMax% = 1) Then   'Hersteller
        TabHST$(0) = Mid$(RabTab, 7, 5)
        TabHST$(1) = Mid$(RabTab, 12, 5)
        For i% = 2 To 5
            TabHST$(i%) = Mid$(RabTab, 17 + (i% - 2) * 5, 5)
        Next i%
    ElseIf (TabMax% = 2) Then   'BM / ZeiWert
        TabBM# = clsOpTool.CVD(Mid$(RabTab, 23, 8))
        TabZW# = clsOpTool.CVD(Mid$(RabTab, 31, 8))
    End If
    TabMax% = TabMax% + 1
Wend

If (TabMax% = 0) Then LNoRabTab% = 0

Call clsError.DefErrPop
End Sub

Function CalcNNEK#(lm%, rm%, NNAEP#, art%, text$, aep#, pzn$, Optional WuInDm% = 0)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("CalcNNEK#")
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
Dim i%, j%, k%, l%, gef%, KriegtRabatt%, eingabe%, orgart%, wg%, val2%, GruppeOk%
Dim FabsRecno&
Dim GesNNAEP#
Dim SQLStr$, herst$

eingabe% = False
orgart% = art%
gef% = False

If (iTabOk%) Then
    art% = 0
    
    SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + clsOpTool.SqlPzn(pzn$)
    If (TaxeAdoDBok) Then
        On Error Resume Next
        TaxeAdoRec.Close
        Err.Clear
        On Error GoTo DefErr
        TaxeAdoRec.Open SQLStr, TaxeAdoDB1.ActiveConn
        If (TaxeAdoRec.EOF = False) Then
            Call clsOpTool.Taxe2ast(pzn$)
            gef% = True
        Else
            If (ArtikelDBok) Then
                SQLStr$ = "SELECT * FROM Artikel WHERE PZN = " + pzn$
                FabsErrf = ArtikelDB1.OpenRecordset(ArtikelRec, SQLStr)
                If (FabsErrf% = 0) Then
                    gef% = True
                End If
            Else
                FabsErrf% = Ast1.IndexSearch(0, pzn$, FabsRecno&)
                If (FabsErrf% = 0) Then
                    Ast1.GetRecord (FabsRecno& + 1)
                    gef% = True
                End If
            End If
        End If
    Else
        Set TaxeRec = TaxeDB.OpenRecordset(SQLStr$)
        If (TaxeRec.EOF = False) Then
    '        AEP# = TaxeRec!EK / 100#
            Call clsOpTool.Taxe2ast(pzn$)
            gef% = True
        Else
            FabsErrf% = Ast1.IndexSearch(0, pzn$, FabsRecno&)
            If (FabsErrf% = 0) Then
                Ast1.GetRecord (FabsRecno& + 1)
    '            AEP# = Ast1.AEP
                gef% = True
            End If
        End If
    End If
    If (gef% = False) Or (Ast1.aep = 0#) Then
        eingabe% = True
        Call EkEingabe(text$, aep#, 0, -1)
        If (gef% = False) Then GoTo CalcNNEKfertig
    Else
        aep# = Ast1.aep
        If (WuInDm% = 1) Then aep# = clsOpTool.FNX(aep# * EURO_FAKTOR)
    End If
    
    If (NNAEP# = 0#) Then NNAEP# = aep#
    
    If (AktLieferant% <> LifZus1.AktRecno) Then
        LifZus1.GetRecord (AktLieferant% + 1)
    End If
    
    KriegtRabatt% = CheckAusnahme%(orgart%, lm%, NNAEP#)
    
    If (KriegtRabatt%) Then
        art% = 4
        With LifZus1
            If (.TabTyp = 0) Then
                NNAEP# = NNAEP# * (100! - .PrognoseRabatt) / 100!
            Else
                For i% = 0 To 4
                    If (.BmEkRabatt(i%) = 0#) Then Exit For
                    If (lm% >= .AbBM(i%)) And ((lm% <= .BisBM(i%)) Or (.BisBM(i%) = 0)) Then
                        GesNNAEP# = NNAEP# * lm%
                        If ((.Vergleich(i%) = 0) And (GesNNAEP# < .zWert(i%))) Or _
                           ((.Vergleich(i%) = 1) And (GesNNAEP# <= .zWert(i%))) Or _
                           ((.Vergleich(i%) = 2) And (GesNNAEP# = .zWert(i%))) Or _
                           ((.Vergleich(i%) = 3) And (GesNNAEP# >= .zWert(i%))) Or _
                           ((.Vergleich(i%) > 3) And (GesNNAEP# > .zWert(i%))) Then
                            GruppeOk% = True
                            val2% = .BmEkTabGruppe(i%)
                            
                            If (TaxeAdoDBok) Then
                                If (TaxeAdoRec.EOF = False) And (val2% > 0) Then
                                    GruppeOk% = 0
                                    l% = Format(TaxeAdoRec!AbgabeBest, "0")
                                    k% = 1
                                    For j% = 1 To l%
                                        k% = k% * 2
                                    Next j%
                                    If (val2% And k%) Then
                                        GruppeOk% = True
                                    End If
                                End If
                            Else
                                If (TaxeRec.EOF = False) And (val2% > 0) Then
                                    GruppeOk% = 0
                                    l% = Format(TaxeRec!AbgabeBest, "0")
                                    k% = 1
                                    For j% = 1 To l%
                                        k% = k% * 2
                                    Next j%
                                    If (val2% And k%) Then
                                        GruppeOk% = True
                                    End If
                                End If
                            End If
                            If (GruppeOk%) Then
                                NNAEP# = NNAEP# * (100! - .BmEkRabatt(i%)) / 100!
                                Exit For
                            End If
                        End If
                    End If
                Next i%
            End If
        End With
        
'        If (LRabTab% > 0) And (LRabTab% < 100) Then
'            NNAEP# = NNAEP# * (100! - AktRab!) / 100!
'        Else
'            For i% = 0 To UBRabTab%
'                If (lm% >= TabDat1%(i%)) And (lm% <= TabDat2%(i%) Or TabDat2%(i%) = 0) Then
'                    GesNNAEP# = NNAEP# * lm%
'                    If ((TabDat3%(i%) = 0) And (GesNNAEP# = TabDat4#(i%))) Or _
'                       ((TabDat3%(i%) = 1) And (GesNNAEP# < TabDat4#(i%))) Or _
'                       ((TabDat3%(i%) = 2) And (GesNNAEP# <= TabDat4#(i%))) Or _
'                       ((TabDat3%(i%) = 3) And (GesNNAEP# >= TabDat4#(i%))) Or _
'                       ((TabDat3%(i%) > 3) And (GesNNAEP# > TabDat4#(i%))) Then
'                        NNAEP# = NNAEP# * (100! - TabDat5!(i%)) / 100!
'                        Exit For
'                    End If
'                End If
'            Next i%
'        End If
    Else
        art% = 3
    End If
    
    If (lm% <> 0) Then NNAEP# = (NNAEP# * rm%) / lm%
Else
    art% = 5
End If

CalcNNEKfertig:
NNAEP# = clsOpTool.FNX(NNAEP#)
If (eingabe%) Then art% = 50 + art%

Call clsError.DefErrPop
End Function

Function EkEingabe%(txt$, Preis#, modus%, eingabe%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("EkEingabe%")
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
                  
EkEingabeText$ = txt$
EkEingabePreis# = Preis#
EkEingabeModus% = modus%
EkEingabeArt% = eingabe%

frmEkEingabe.Show 1

If (EkEingabeErg%) Then Preis# = EkEingabePreis#

EkEingabe% = EkEingabeErg%

Call clsError.DefErrPop
End Function

Function HatRabatt%(Lieferant%, pzn$, art%, lm%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("HatRabatt%")
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
Dim i%, gef%, KriegtRabatt%, wg%
Dim j&, FabsRecno&
Dim aep#
Dim SQLStr$, herst$
                  
If (Para1.Land <> "D") Then
    HatRabatt% = 0: Call clsError.DefErrPop: Exit Function
End If


KriegtRabatt% = True

If (Lieferant% <> LifZus1.AktRecno) Then
    LifZus1.GetRecord (Lieferant% + 1)
End If

If (LifZus1.HatAusnahmen) Then
    SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + clsOpTool.SqlPzn(pzn$)
    If (TaxeAdoDBok) Then
        On Error Resume Next
        TaxeAdoRec.Close
        Err.Clear
        On Error GoTo DefErr
        TaxeAdoRec.Open SQLStr, TaxeAdoDB1.ActiveConn
        If (TaxeAdoRec.EOF = False) Then
            Call clsOpTool.Taxe2ast(pzn$)
            gef% = True
        Else
            If (ArtikelDBok) Then
                SQLStr$ = "SELECT * FROM Artikel WHERE PZN = " + pzn$
                FabsErrf = ArtikelDB1.OpenRecordset(ArtikelRec, SQLStr)
                If (FabsErrf% = 0) Then
                    gef% = True
                End If
            Else
                FabsErrf% = Ast1.IndexSearch(0, pzn$, FabsRecno&)
                If (FabsErrf% = 0) Then
                    Ast1.GetRecord (FabsRecno& + 1)
                    gef% = True
                End If
            End If
        End If
    Else
        Set TaxeRec = TaxeDB.OpenRecordset(SQLStr$)
        If (TaxeRec.EOF = False) Then
    '        AEP# = TaxeRec!EK / 100#
            Call clsOpTool.Taxe2ast(pzn$)
            gef% = True
        Else
            FabsErrf% = Ast1.IndexSearch(0, pzn$, FabsRecno&)
            If (FabsErrf% = 0) Then
                Ast1.GetRecord (FabsRecno& + 1)
    '            AEP# = Ast1.AEP
                gef% = True
            End If
        End If
    End If
    
    If (gef%) Then
        aep# = Ast1.aep
        KriegtRabatt% = CheckAusnahme%(art%, lm%, aep#)
    End If
End If
    
HatRabatt% = KriegtRabatt%

Call clsError.DefErrPop
End Function

Function CheckAusnahme%(art%, lm%, aep#)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("CheckAusnahme%")
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
CheckAusnahme% = LifZus1.CheckAusnahme(-1, art%, lm%, aep#)
Call clsError.DefErrPop
End Function

Function GetRabTab%(Lieferant%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("GetRabTab%")
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
Dim i%, iTabSatz%, iPos%, iTabNr%, iTabStrt%, iTabMax%
Dim j&
Dim ret%, dUmsatz#
                  
ret% = 0

'Call OpenDatei
'If (LOF(dHandle%) = 0) Then
'    RabTab = String$(SATZLEN, 0)
'    For j& = 1 To 156
'        Call PutRecord(j&)
'    Next j&
'End If


For i% = 0 To 6: TabKZ%(i%) = 0: Next i%
For i% = 0 To 9: TabWG%(i%) = 0: Next i%
For i% = 0 To 5: TabHST$(i%) = "": Next i%
For i% = 0 To 10  '127!!!!!!!!!!!
    TabDat1%(i%) = 0
    TabDat2%(i%) = 0
    TabDat3%(i%) = 0
    TabDat4#(i%) = 0#
    TabDat5!(i%) = 0!
Next i%
For i% = 0 To 9
    lSchwellwert#(i%) = 0#
    lRabatt!(i%) = 0!
Next i%

If (Lieferant% > 5) Then
    iTabSatz% = Int((Lieferant% - 6) / 7) + 2
    iPos% = ((Lieferant% - 6) Mod 7) * 6 + 1
Else
    iTabSatz% = 1
    iPos% = ((Lieferant% - 1) Mod 7) * 6 + 13
End If

Call GetRecord(CLng(iTabSatz%))

LRabTab% = Asc(Mid$(RabTab, iPos%, 1))
LNoRabTab% = Asc(Mid$(RabTab, iPos% + 2, 1))

If (LRabTab% > 0) Then
    ret% = LRabTab%
    
    If (LRabTab% > 0) And (LRabTab% < 100) Then
        iTabNr% = LRabTab%
        iTabSatz% = Int((iTabNr% - 1) / 21)
        iPos% = ((iTabNr% - 1) Mod 21) * 2 + 1
        Call GetRecord(iTabSatz% + 144)
        iTabStrt% = clsOpTool.CVI(Mid$(RabTab, iPos%, 2))

        'EK-Werttabelle einlesen
        For i% = 0 To 9
            If (iTabStrt% <= 0) Then Exit For
            Call GetRecord(iTabStrt% + 156)
            iTabStrt% = clsOpTool.CVI(Mid$(RabTab, 41, 2))
            lSchwellwert#(i%) = clsOpTool.CVD(Mid$(RabTab, 7, 8))
            lRabatt!(i%) = clsOpTool.CVS(Mid$(RabTab, 23, 4))
        Next i%
        Do
            If (i% >= 10) Then Exit Do
            lSchwellwert#(i%) = 0#
            lRabatt!(i%) = 0!
            i% = i% + 1
        Loop
    Else
        Call BMZeilenTab
        If (TabDat1%(0) = 1) And (TabDat2%(0) = 999) And (TabDat4#(0) = 0#) Then
            LRabTab% = 1
            ret% = LRabTab%
            lSchwellwert#(0) = 0#
            lRabatt!(0) = TabDat5!(0)
            i% = 1
            Do
                If (i% >= 10) Then Exit Do
                lSchwellwert#(i%) = 0#
                lRabatt!(i%) = 0!
                i% = i% + 1
            Loop
        End If
    End If
End If

If (LNoRabTab% <> 0) Then Call AusnahmenTab
    
'Call CloseDatei

GetRabTab% = ret%

Call clsError.DefErrPop
End Function

Public Property Get Schwellwert#(ByVal ind%)
Schwellwert# = lSchwellwert#(ind%)
End Property
Public Property Let Schwellwert(ByVal ind%, ByVal vNewValue#)
lSchwellwert#(ind%) = vNewValue#
End Property


Public Property Get Rabatt!(ByVal ind%)
Rabatt! = lRabatt!(ind%)
End Property
Public Property Let Rabatt(ByVal ind%, ByVal vNewValue!)
lRabatt!(ind%) = vNewValue!
End Property


Public Property Get IstAusnahme%(ByVal ind%)
IstAusnahme% = TabKZ%(ind%)
End Property
Public Property Let IstAusnahme(ByVal ind%, ByVal vNewValue%)
TabKZ%(ind%) = vNewValue%
End Property


Sub KonvertTabelle()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("KonvertTabelle")
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
Dim i%, j%, erg%, MitAusnahmen%, vergl%

Call OpenDatei
If (LOF(dHandle%) > 0) Then
    For i% = 1 To LifZus1.AnzRec
        erg% = GetRabTab%(i%)
        If (erg% > 0) Then
            LifZus1.GetRecord (i% + 1)
            If (erg% < 100) Then
                LifZus1.TabTyp = 0
            Else
                LifZus1.TabTyp = 1
            End If
            For j% = 0 To 4
                LifZus1.Schwellwert(j%) = lSchwellwert#(j%)
                LifZus1.Rabatt(j%) = CInt(lRabatt!(j%))
            Next j%
            For j% = 0 To 4
                LifZus1.AbBM(j%) = TabDat1%(j%)
                LifZus1.BisBM(j%) = TabDat2%(j%)
                
                vergl% = TabDat3%(j%)
                If (vergl% = 0) Then
                    vergl% = 2
                ElseIf (vergl% <= 2) Then
                    vergl% = vergl% - 1
                End If
                LifZus1.Vergleich(j%) = vergl%
                
                LifZus1.zWert(j%) = TabDat4#(j%)
                LifZus1.BmEkRabatt(j%) = CInt(TabDat5!(j%))
            Next j%
            
            MitAusnahmen% = False
            For j% = 0 To 6
                LifZus1.AusnahmenKz(j%) = Abs(TabKZ%(j%))
                If (TabKZ%(j%)) Then MitAusnahmen% = True
            Next j%
            LifZus1.HatAusnahmen = Abs(MitAusnahmen%)
            
            For j% = 0 To 9
                LifZus1.AusnahmenWg(j%) = TabWG%(j%)
            Next j%
            LifZus1.AusnahmenBM = CInt(TabBM#)
            LifZus1.AusnahmenZW = TabZW#
            For j% = 0 To 5
                LifZus1.AusnahmenHerst(j%) = TabHST$(j%)
            Next j%
            LifZus1.PutRecord (i% + 1)
        End If
    Next i%
End If
    
Call clsError.DefErrPop
End Sub


Sub InitBestLief(pos%, iLief%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("InitBestLief")
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
Call LifZus1.InitBestLief(pos%, iLief%)
Call clsError.DefErrPop
End Sub

Function GetBestLief%(sLiefs$, pzn$, art%, lm%, Rabatt#)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("GetBestLief%")
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
GetBestLief% = LifZus1.GetBestLief(sLiefs$, pzn$, art%, lm%, Rabatt#)
Call clsError.DefErrPop
End Function


