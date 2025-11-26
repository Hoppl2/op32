VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDateien"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim FabsErrf%
Dim FabsRecno&
Dim cmd$

Dim InfoPzn$

Private Const DefErrModul = "DATEIEN.CLS"

Function FileOpen%(fName$, fAttr$, Optional modus$ = "B", Optional SATZLEN% = 100)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("FileOpen%")
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
Dim Handle%

On Error Resume Next
FileOpen% = False
Handle% = FreeFile

'Debug.Print fName$; Str$(Handle%)
fAttr$ = UCase(fAttr$)

If (fAttr$ = "R") Then
    If (modus$ = "B") Then
        Open fName$ For Binary Access Read Shared As #Handle%
    Else
        Open fName$ For Random Access Read Shared As #Handle% Len = SATZLEN%
    End If
    If (Err = 0) Then
        If (LOF(Handle%) = 0) Then
            Close #Handle%
            Kill (fName$)
            Err.Raise 53
        Else
            Call iLock(Handle%, 1)
            Call iUnLock(Handle%, 1)
        End If
    End If
ElseIf (fAttr$ = "W") Then
    If (modus$ = "B") Then
        Open fName$ For Binary Access Write As #Handle%
    Else
        Open fName$ For Random Access Write As #Handle% Len = SATZLEN%
    End If
ElseIf (fAttr$ = "RW") Then
    If (modus$ = "B") Then
        Open fName$ For Binary Access Read Write Shared As #Handle%
    Else
        Open fName$ For Random Access Read Write Shared As #Handle% Len = SATZLEN%
    End If
    Call iLock(Handle%, 1)
    Call iUnLock(Handle%, 1)
ElseIf (fAttr$ = "RWL") Then
    If (modus$ = "B") Then
        Open fName$ For Binary Access Read Write Lock Read Write As #Handle%
    Else
        Open fName$ For Random Access Read Write Lock Read Write As #Handle% Len = SATZLEN%
    End If
ElseIf (fAttr$ = "I") Then
    Open fName$ For Input Access Read Shared As #Handle%
ElseIf (fAttr$ = "O") Then
    Open fName$ For Output Access Write Shared As #Handle%
End If

If (Err = 0) Then
    FileOpen% = Handle%
Else
    Call clsError.DefErrAnswer(Err.Source, Err.Number, Err.Description + " (" + fName$ + ")", DefErrModul, 1)
    Call ProjektForm.EndeDll
'    Call MsgBox("Fehler" + Str$(Err) + " beim Öffnen von " + fName$ + vbCr + Err.Description, vbCritical, "FileOpen")
End If

Call clsError.DefErrPop
End Function

Sub iLock(file As Integer, SatzNr&)
Dim LockTime As Date
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("iLock")
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
  Resume
Case vbIgnore
  Resume Next
End Select
Call ProjektForm.EndeDll
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim s As String

If (Para1.MehrPlatz) Then Lock #file, SatzNr&

Call clsError.DefErrPop
End Sub

'Sub iLock(file As Integer, SatzNr&)
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Call clsError.DefErrFnc("iLock")
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
'Dim LockTime As Date
'Dim s As String
'
'On Error GoTo StandardError
'
'If (Para1.MehrPlatz) Then Lock #file, SatzNr&
'
'Call clsError.DefErrPop: Exit Sub
'
''------------------------------------------------------------
'StandardError:
'
'
'If Err = 70 Or Err = 75 Then
'  If LockTime = 0 Then LockTime = DateAdd("s", 10, Now)
'  If LockTime > Now Then
'    'Sleep (1)
'    Resume
'  End If
'End If
'Error Err.Number
'Call clsError.DefErrPop
'End Sub

Sub iUnLock(file As Integer, SatzNr&)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("iUnLock")
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
If (Para1.MehrPlatz) Then Unlock #file, SatzNr&

Call clsError.DefErrPop
End Sub

Function SucheFreienIndex%()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("SucheFreienIndex%")
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

For i% = 20 To 0 Step -1
    If (IndexBelegt%(i%) = 0) Then
        IndexBelegt%(i%) = True
        ret% = i%
        Exit For
    End If
Next i%

SucheFreienIndex% = ret%

Call clsError.DefErrPop
End Function

Function IndexOpen%(iName$, iHandle%, Optional MitReorg% = 0)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("IndexOpen%")
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
Dim ind%, iInd%, ret%
Dim FullName$

ind% = InStr(iName$, "\")
If (ind% <= 0) Then
    FullName$ = CurDir + "\" + iName$
Else
    FullName$ = iName$
End If

ret% = 22
If (Dir$(FullName$) <> "") Then
    If (wPara1.CalcFileSize(FullName$) >= 1024) Then
        ret% = -1
    End If
End If


If (ret% = -1) Then
    Do
        ind% = InStr(FullName$, "\")
        If (ind% > 0) Then
            Mid$(FullName$, ind%, 1) = "/"
        Else
            Exit Do
        End If
    Loop
    
    iInd% = SucheFreienIndex
    If (iInd% >= 0) Then
        iHandle% = iInd%
        cmd$ = "O\" + RTrim$(FullName$) + "\" + Mid$(Str$(iInd%), 2)
'        ret% = clsfabs.Cmnd(cmd$, FabsRecno&)
        ret = 99
    End If
End If

    
If (ret% <> 0) Then
    If (MitReorg%) And ((ret% = 22) Or (ret% = 29)) Then
    Else
        Call clsError.DefErrAnswer("", (ret%), "(" + iName$ + ")", DefErrModul, 1)
        Call ProjektForm.EndeDll
    '    Call MsgBox("Index-Fehler" + Str$(ret%) + " bei Befehl " + vbCr + cmd$, vbCritical, "IndexOpen")
    End If
End If

IndexOpen% = ret%

Call clsError.DefErrPop
End Function

Function IndexClose%(ind%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("IndexClose%")
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
Dim iRecno&

cmd$ = "K\" + Mid$(Str$(ind%), 2)
'FabsErrf% = clsfabs.Cmnd(cmd$, iRecno&)
FabsErrf% = 99
IndexBelegt%(ind%) = False

IndexClose% = FabsErrf%

Call clsError.DefErrPop
End Function

Function IndexCreate%(iName$, iHandle%, iKeyLen%)
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
Dim ind%, iInd%, ret%
Dim FullName$

ret% = -1

ind% = InStr(iName$, "\")
If (ind% <= 0) Then
    FullName$ = CurDir + "\" + iName$
Else
    FullName$ = iName$
End If
On Error Resume Next
Kill FullName$
On Error GoTo DefErr
Do
    ind% = InStr(FullName$, "\")
    If (ind% > 0) Then
        Mid$(FullName$, ind%, 1) = "/"
    Else
        Exit Do
    End If
Loop

iInd% = SucheFreienIndex
If (iInd% >= 0) Then
    iHandle% = iInd%
    cmd$ = "C\" + RTrim$(FullName$) + "\" + Mid$(Str$(iKeyLen%), 2) + "\1\A\" + Mid$(Str$(iInd%), 2)
'    ret% = clsfabs.Cmnd(cmd$, FabsRecno&)
    ret = 99
End If

If (ret% <> 0) Then
    Call clsError.DefErrAnswer("", (ret%), cmd$, DefErrModul, 1)
    Call ProjektForm.EndeDll
'    Call MsgBox("Index-Fehler" + Str$(ret%) + " bei Befehl " + vbCr + cmd$, vbCritical, "IndexOpen")
End If

IndexCreate% = ret%

Call clsError.DefErrPop
End Function

Function IndexSearch%(ind%, key$, FabsRecno&)
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
Dim s$

cmd$ = "S\1\" + Mid$(Str$(ind%), 2) + "\" + key$
'FabsErrf% = clsfabs.Cmnd(cmd$, FabsRecno&)
FabsErrf = 99

IndexSearch% = FabsErrf%

Call clsError.DefErrPop
End Function

Function IndexGeneric%(ind%, key$, FabsRecno&)
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

cmd$ = "G\1\" + Mid$(Str$(ind%), 2) + "\" + key$
'FabsErrf% = clsfabs.Cmnd(cmd$, FabsRecno&)
FabsErrf = 99

IndexGeneric% = FabsErrf%

Call clsError.DefErrPop
End Function

Function IndexFirst%(ind%, FabsRecno&)
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

cmd$ = "F\1\" + Mid$(Str$(ind%), 2)
'FabsErrf% = clsfabs.Cmnd(cmd$, FabsRecno&)
FabsErrf = 99

IndexFirst% = FabsErrf%

Call clsError.DefErrPop
End Function

Function IndexNext%(ind%, AltRecno&, key$, FabsRecno&)
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

cmd$ = "A\1\" + Mid$(Str$(AltRecno&), 2) + "\" + Mid$(Str$(ind%), 2) + "\" + key$
'FabsErrf% = clsfabs.Cmnd(cmd$, FabsRecno&)
FabsErrf = 99

IndexNext% = FabsErrf%

Call clsError.DefErrPop
End Function

Function IndexPrevious%(ind%, AltRecno&, key$, FabsRecno&)
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

cmd$ = "E\1\" + Mid$(Str$(AltRecno&), 2) + "\" + Mid$(Str$(ind%), 2) + "\" + key$
'FabsErrf% = clsfabs.Cmnd(cmd$, FabsRecno&)
FabsErrf = 99

IndexPrevious% = FabsErrf%

Call clsError.DefErrPop
End Function

Function IndexInsert%(ind%, key$, FabsRecno&)
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

cmd$ = "I\" + Mid$(Str$(ind%), 2) + "\" + key$
'FabsErrf% = clsfabs.Cmnd(cmd$, FabsRecno&)
FabsErrf = 99

IndexInsert% = FabsErrf%

Call clsError.DefErrPop
End Function

Function IndexDelete%(ind%, AltRecno&, key$, FabsRecno&)
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

cmd$ = "D\" + Mid$(Str$(AltRecno&), 2) + "\Y\" + Mid$(Str$(ind%), 2) + "\" + key$
'FabsErrf% = clsfabs.Cmnd(cmd$, FabsRecno&)
FabsErrf = 99

IndexDelete% = FabsErrf%

Call clsError.DefErrPop
End Function

Function IndexReplace%(ind%, AltRecno&, AltKey$, key$, FabsRecno&)
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

cmd$ = "R\1\" + Mid$(Str$(AltRecno&), 2) + "\" + Mid$(Str$(ind%), 2) + "\" + AltKey$ + "\" + key$
'FabsErrf% = clsfabs.Cmnd(cmd$, FabsRecno&)
FabsErrf = 99

IndexReplace% = FabsErrf%

Call clsError.DefErrPop
End Function

Function IndexCountRecords%(ind%, FabsRecno&)
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

cmd$ = "T\" + Mid$(Str$(ind%), 2)
'FabsErrf% = clsfabs.Cmnd(cmd$, FabsRecno&)
FabsErrf = 99

IndexCountRecords% = FabsErrf%

Call clsError.DefErrPop
End Function

Function IndexCountDeleted%(ind%, FabsRecno&)
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

cmd$ = "U\" + Mid$(Str$(ind%), 2)
'FabsErrf% = clsfabs.Cmnd(cmd$, FabsRecno&)
FabsErrf = 99

IndexCountDeleted% = FabsErrf%

Call clsError.DefErrPop
End Function




'Function DecodFeld$(RawData$, fTyp$, fKonv%)
'Dim ret$, h$
'
'ret$ = ""
'
'Select Case fTyp$
'    Case "T"
'        ret$ = " " + LTrim$(RawData$)
'    Case "N"
'        Select Case fKonv%
'            Case 1
'                'ASC
'                ret$ = Str$(Asc(RawData$))
'             Case 2
'                'CVD
'                ret$ = Format$(clsOpTool.CVD(RawData$), "0.00")
'            Case 3
'                'VAL
'                ret$ = RawData$
'             Case 4
'                'CVS
'                ret$ = Format$(clsOpTool.CVS!(RawData$), "0.00")
'             Case 5
'                'CVI
'                ret$ = Str$(clsOpTool.CVI%(RawData$))
'             Case 6
'                'CVI / 10
'                ret$ = Format$(clsOpTool.CVI%(RawData$) / 10, "0.00")
'        End Select
'        ret$ = LTrim$(RTrim$(ret$))
'        If (ret$ = "0.00") Or (ret$ = "0,00") Then
'            ret$ = "0"
'        End If
'        ret$ = " " + ret$
'    Case "D"
'        Select Case fKonv%
'            Case 10
'                'CVDatum$
'                h$ = clsOpTool.CVDatum$(RawData$)
'                ret$ = " " + Mid$(h$, 7, 2) + "." + Mid$(h$, 5, 2) + "." + Left$(h$, 4)
'        End Select
'End Select
'
'DecodFeld$ = ret$
'
'End Function

Function DecodFeld$(RawData As Variant, fTyp$, fKonv%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("DecodFeld$")
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
Dim ret$, h$

ret$ = ""

Select Case fTyp$
    Case "T"
        ret$ = " " + LTrim$(RawData)
        Call OemToChar(ret$, ret$)
    Case "N"
        Select Case fKonv%
            Case 1
                'ASC
                ret$ = Str$(RawData)
             Case 2
                'CVD
                ret$ = Format$(RawData, "0.00")
            Case 3
                'VAL
                ret$ = Str$(RawData)
             Case 4
                'CVS
                ret$ = Format$(RawData, "0.00")
             Case 5
                'CVI
                ret$ = Str$(RawData)
             Case 6
                'CVI / 10
                ret$ = Format$(RawData, "0.00")
             Case 7
                'CVL
                ret$ = Str$(RawData)
             Case 8
                'Zuzahlung
                ret$ = RawData
        End Select
        ret$ = LTrim$(RTrim$(ret$))
        If (ret$ = "0.00") Or (ret$ = "0,00") Then
            ret$ = "0"
        End If
        ret$ = " " + ret$
    Case "D"
        Select Case fKonv%
            Case 10
                'CVDatum$
                h$ = clsOpTool.CVDatum2$(RawData)
                ret$ = " " + Mid$(h$, 7, 2) + "." + Mid$(h$, 5, 2) + "." + Left$(h$, 4)
        End Select
End Select

DecodFeld$ = ret$

Call clsError.DefErrPop
End Function

Function FeldValue(RawData$, fTyp$, fKonv%) As Variant
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("FeldValue")
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
Dim ret As Variant
Dim h$

ret = ""

Select Case fTyp$
    Case "T"
        ret = RawData$
    Case "N"
        Select Case fKonv%
            Case 1
                'ASC
                ret = Asc(RawData$)
             Case 2
                'CVD
                ret = clsOpTool.CVD(RawData$)
            Case 3
                'VAL
                ret = Val(RawData$)
             Case 4
                'CVS
                ret = clsOpTool.CVS!(RawData$)
             Case 5
                'CVI
                ret = clsOpTool.CVI%(RawData$)
             Case 6
                'CVI / 10
                ret = clsOpTool.CVI%(RawData$) / 10#
             Case 7
                'CVL
                ret = clsOpTool.CVL(RawData$)
            Case 8
                'Zuzahlung
                ret = clsOpTool.MachStammAuswahlZuzahlung(clsOpTool.CVI%(RawData$))
        End Select
    Case "D"
        Select Case fKonv%
            Case 10
                'CVDatum$
'                h$ = clsOpTool.CVDatum$(RawData$)
'                ret$ = " " + Mid$(h$, 7, 2) + "." + Mid$(h$, 5, 2) + "." + Left$(h$, 4)
                ret = clsOpTool.CVDat(RawData$)
        End Select
End Select

FeldValue = ret

Call clsError.DefErrPop
End Function

Function MakeFeld$(ByVal RawData As Variant, fTyp$, fKonv%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("MakeFeld$")
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
Dim l&
Dim d#
Dim s!
Dim ret$, h$

ret$ = ""

Select Case fTyp$
    Case "T"
        ret$ = RawData
    Case "N"
        Select Case fKonv%
            Case 1
                'ASC
                ret$ = Chr$(RawData)
             Case 2
                'CVD
                d# = RawData
                ret$ = clsOpTool.MKD(d#)
            Case 3
                'VAL
                ret$ = Mid$(Str$(RawData), 2)
             Case 4
                'CVS
                s! = RawData
                ret$ = clsOpTool.MKS(s!)
             Case 5
                'CVI
                i% = RawData
                ret$ = clsOpTool.MKI(i%)
             Case 6
                'CVI / 10
                i% = RawData
                ret$ = clsOpTool.MKI(CInt(i% * 10))
             Case 7
                'CVI
                l& = RawData
                ret$ = clsOpTool.MKL(l&)
        End Select
    Case "D"
        Select Case fKonv%
            Case 10
                'CVDatum$
'                h$ = clsOpTool.CVDatum$(RawData$)
'                ret$ = " " + Mid$(h$, 7, 2) + "." + Mid$(h$, 5, 2) + "." + Left$(h$, 4)
                i% = RawData
                ret = clsOpTool.MKDate(i%)
        End Select
End Select

MakeFeld$ = ret

Call clsError.DefErrPop
End Function

Function FirstLettersUcase(ByVal s As String) As String
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("FirstLettersUcase")
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
Dim i As Integer
Dim c As String
  
s = UCase(Left(s, 1)) + LCase(Mid(s, 2))
For i = 2 To Len(s)
  c = UCase(Mid(s, i - 1, 1))
  If InStr("ABCDEFGHIJKLMNOPQRSTUVWXYZÄÖÜß", c) = 0 Then
    Mid(s, i, 1) = UCase(Mid(s, i, 1))
  End If
Next i
FirstLettersUcase = s

Call clsError.DefErrPop
End Function

Function CheckKlasse%(dKurz$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("CheckKlasse%")
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
Dim ret%
Dim SQLStr$, s$

ret% = False

On Error Resume Next
Select Case dKurz$
    Case "AST", "ASS", "ART"
        If (ArtikelDBok) Then
            On Error Resume Next
            ArtikelRec.Close
            Err.Clear
            On Error GoTo DefErr
            SQLStr$ = "SELECT * FROM ARTIKEL WHERE PZN = 1234567"
            ArtikelRec.Open SQLStr, ArtikelDB1.ActiveConn
        ElseIf (dKurz = "AST") Then
            s$ = Ast1.DateiKurz
        ElseIf (dKurz = "ASS") Then
            s$ = Ass1.DateiKurz
        End If
'    Case "AST"
'        s$ = Ast1.DateiKurz
'    Case "ASS"
'        s$ = Ass1.DateiKurz
    Case "TAXE"
        If (TaxeAdoDBok) Then
            On Error Resume Next
            TaxeAdoRec.Close
            Err.Clear
            On Error GoTo DefErr
            SQLStr$ = "SELECT * FROM TAXE WHERE PZN = 1234567"
            TaxeAdoRec.Open SQLStr, TaxeAdoDB1.ActiveConn
        Else
            SQLStr$ = "SELECT * FROM TAXE WHERE PZN = 1234567"
            Set TaxeRec = TaxeDB.OpenRecordset(SQLStr$)
        End If
    Case "NNEK"
        s$ = nnek1.DateiKurz
    Case "LIEF"
        If (LieferantenDBok) Then
            On Error Resume Next
            LieferantenRec.Close
            Err.Clear
            On Error GoTo DefErr
            SQLStr$ = "SELECT *, Lieferanten.LiefNr AS LiefNr2 FROM Lieferanten LEFT JOIN LieferantenZusatz ON Lieferanten.LiefNr=LieferantenZusatz.LiefNr"
            SQLStr$ = SQLStr + " WHERE Lieferanten.LiefNr = 1"
            LieferantenRec.Open SQLStr, LieferantenDB1.ActiveConn   ' LieferantenConn
        Else
            s$ = Lif1.DateiKurz
        End If
    Case "HTAXE"
        s$ = hTax1.DateiKurz
    Case "SONDERPR"
        If (ArtikelDBok) Then
            On Error Resume Next
            SonderPrAdoRec.Close
            Err.Clear
            On Error GoTo DefErr
            SQLStr$ = "SELECT * FROM SonderPreise WHERE PZN = 1234567"
            SonderPrAdoRec.Open SQLStr, ArtikelDB1.ActiveConn
        Else
            SQLStr$ = "SELECT * FROM SonderPreise WHERE Pzn = '1234567'"
            Set SonderPrRec = SonderPrDB.OpenRecordset(SQLStr$)
        End If
    Case "DM"
        If (ArtikelDBok) Then
            On Error Resume Next
            DocMorrisRec.Close
            Err.Clear
            On Error GoTo DefErr
            SQLStr$ = "SELECT * FROM DocMorris WHERE PZN = 1234567"
            DocMorrisRec.Open SQLStr, ArtikelDB1.ActiveConn
        ElseIf (dKurz = "AST") Then
            s$ = Ast1.DateiKurz
        ElseIf (dKurz = "ASS") Then
            s$ = Ass1.DateiKurz
        End If
    Case "KZ"
        s = "1" 'damit irgendwas dasteht
    Case Else
        Err.Raise 91
End Select

CheckKlasse% = (Err = 0)

Call clsError.DefErrPop
End Function

Function EinlesenDateiDS%(pzn$, dKurz$, AuchStruktur%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("EinlesenDateiDS%")
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
Dim FabsErrf%, ret%, DatenVonLief%
Dim FabsRecno&
Dim SQLStr$

ret% = False

Select Case dKurz$
    Case "AST", "ASS", "ART", "NNEK"
        If (ArtikelDBok) Then
            On Error Resume Next
            ArtikelRec.Close
            Err.Clear
            On Error GoTo DefErr
'            SQLStr$ = "SELECT * FROM ARTIKEL WHERE PZN = " + clsOpTool.SqlPzn(pzn$)
            SQLStr = "SELECT TOP (1) Artikel.*, Abgaben.Datum AS ABG_Datum,Abgaben.AbgabeMenge AS ABG_Menge, Lieferungen.BelegDatum AS LIEF_Datum,Lieferungen.LiefNr AS LIEF_LiefNr,Lieferungen.LM AS LIEF_LM,Lieferungen.RM AS LIEF_RM"
            SQLStr = SQLStr + ",Lieferungen.NR AS LIEF_NR,Lieferungen.RetMenge AS LIEF_RetMenge,Lieferungen.art AS LIEF_Art,Lieferungen.NNAEP AS LIEF_NNAEP,Lieferungen.Beleg AS LIEF_Beleg,Lieferungen.KuNr AS LIEF_KuNr"
'            If (iMARS) Then
'                SQLStr = SQLStr + ",DM.VK AS DM_VK, DM.Automat AS DM_Automat, DM.Gesperrt as DM_GESPERRT"
'            End If
            SQLStr = SQLStr + " FROM Artikel LEFT JOIN"
            SQLStr = SQLStr + " Abgaben ON Artikel.PZN = Abgaben.PZN LEFT JOIN"
            SQLStr = SQLStr + " Lieferungen ON Artikel.PZN = Lieferungen.PZN"
'            If (iMARS) Then
'                SQLStr = SQLStr + " LEFT JOIN DocMorris AS DM ON Artikel.PZN = DM.PZN"
'            End If
            SQLStr = SQLStr + " WHERE Artikel.pzn = " + clsOpTool.SqlPzn(pzn$)
            SQLStr = SQLStr + " ORDER BY Abgaben.Datum DESC, Lieferungen.BelegDatum DESC"
            ArtikelRec.Open SQLStr, ArtikelDB1.ActiveConn
            If (ArtikelRec.EOF = False) Then
                Call ArtikelDB1.DB2DOS(ArtikelRec)
                ret% = True
            End If
        ElseIf (dKurz = "AST") Then
            FabsErrf% = Ast1.IndexSearch(0, pzn$, FabsRecno&)
            If (FabsErrf% = 0) Then
                ret% = True
                Call Ast1.GetInfoRecord(FabsRecno& + 1)
                If (AuchStruktur%) Then
                    Call Ast1.GetRecord(FabsRecno& + 1)
                End If
            End If
        ElseIf (dKurz = "ASS") Then
            FabsErrf% = Ass1.IndexSearch(0, pzn$, FabsRecno&)
            If (FabsErrf% = 0) Then
                ret% = True
                Call Ass1.GetInfoRecord(FabsRecno& + 1)
                If (AuchStruktur%) Then
                    Call Ass1.GetRecord(FabsRecno& + 1)
                End If
            End If
        ElseIf (dKurz = "NNEK") Then
    '        nnek1.OpenDatei
            FabsErrf% = nnek1.IndexSearch(0, pzn$, FabsRecno&)
            If (FabsErrf% = 0) Then
                ret% = True
                Call nnek1.GetInfoRecord(FabsRecno& + 1)
                If (AuchStruktur%) Then
                    Call nnek1.GetRecord(FabsRecno& + 1)
                End If
            End If
    '        nnek1.CloseDatei
        End If
'    Case "AST"
'        FabsErrf% = Ast1.IndexSearch(0, pzn$, FabsRecno&)
'        If (FabsErrf% = 0) Then
'            ret% = True
'            Call Ast1.GetInfoRecord(FabsRecno& + 1)
'            If (AuchStruktur%) Then
'                Call Ast1.GetRecord(FabsRecno& + 1)
'            End If
'        End If
'    Case "ASS"
'        FabsErrf% = Ass1.IndexSearch(0, pzn$, FabsRecno&)
'        If (FabsErrf% = 0) Then
'            ret% = True
'            Call Ass1.GetInfoRecord(FabsRecno& + 1)
'            If (AuchStruktur%) Then
'                Call Ass1.GetRecord(FabsRecno& + 1)
'            End If
'        End If
    Case "TAXE"
        If (TaxeAdoDBok) Then
            On Error Resume Next
            TaxeAdoRec.Close
            Err.Clear
            On Error GoTo DefErr
            SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + clsOpTool.SqlPzn(pzn$)
            TaxeAdoRec.Open SQLStr, TaxeAdoDB1.ActiveConn
            If (TaxeAdoRec.EOF = False) Then
                ret% = True
            End If
        Else
            SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + clsOpTool.SqlPzn(pzn$)
            Set TaxeRec = TaxeDB.OpenRecordset(SQLStr$)
            If (TaxeRec.EOF = False) Then
                ret% = True
            End If
        End If
    Case "LIEF"
        If (LieferantenDBok) Then
            On Error Resume Next
            LieferantenRec.Close
            Err.Clear
            On Error GoTo DefErr
            SQLStr$ = "SELECT *, Lieferanten.LiefNr AS LiefNr2 FROM Lieferanten LEFT JOIN LieferantenZusatz ON Lieferanten.LiefNr=LieferantenZusatz.LiefNr"
            SQLStr$ = SQLStr + " WHERE Lieferanten.LiefNr = " + clsOpTool.SqlPzn(pzn$)
            LieferantenRec.Open SQLStr, LieferantenDB1.ActiveConn   ' LieferantenConn
            If (LieferantenRec.EOF = False) Then
                LieferantenDB1.OrgLiefNr = clsOpTool.CheckNullLong(LieferantenRec!LiefNr)
                LieferantenDB1.OrgKurz = clsOpTool.CheckNullStr(LieferantenRec!kurz)
                DatenVonLief = clsOpTool.CheckNullLong(LieferantenRec!DatenVonLief)
                If (DatenVonLief > 0) Then
                    LieferantenRec.Close
                    SQLStr$ = "SELECT *, Lieferanten.LiefNr AS LiefNr2 FROM Lieferanten LEFT JOIN LieferantenZusatz ON Lieferanten.LiefNr=LieferantenZusatz.LiefNr"
                    SQLStr$ = SQLStr + " WHERE Lieferanten.LiefNr = " + CStr(DatenVonLief)
                    LieferantenRec.Open SQLStr, LieferantenDB1.ActiveConn   ' LieferantenConn
                    ret = (LieferantenRec.EOF = False)
                Else
                    ret% = True
                End If
            End If
        Else
            ret% = True
            Call Lif1.GetInfoRecord(Val(pzn$) + 1)
            If (AuchStruktur%) Then
                Call Lif1.GetRecord(Val(pzn$) + 1)
            End If
        End If
    Case "HTAXE"
        FabsErrf% = hTax1.IndexSearch(0, pzn$, FabsRecno&)
        If (FabsErrf% = 0) Then
            ret% = True
            Call hTax1.GetInfoRecord(FabsRecno& + 1)
            If (AuchStruktur%) Then
                Call hTax1.GetRecord(FabsRecno& + 1)
            End If
        End If
    Case "SONDERPR"
        If (ArtikelDBok) Then
            On Error Resume Next
            SonderPrAdoRec.Close
            Err.Clear
            On Error GoTo DefErr
            SQLStr$ = "SELECT * FROM SonderPreise WHERE PZN = " + clsOpTool.PznString(Val(pzn))
            SonderPrAdoRec.Open SQLStr, ArtikelDB1.ActiveConn
            If (SonderPrAdoRec.EOF = False) Then
                ret% = True
            End If
        Else
            If (SonderPrOk%) Then
                SQLStr$ = "SELECT * FROM SonderPreise WHERE Pzn = '" + clsOpTool.SqlPzn(pzn$, 0) + "'"
                Set SonderPrRec = SonderPrDB.OpenRecordset(SQLStr$)
                If (SonderPrRec.EOF = False) Then
                    ret% = True
                End If
            End If
        End If
    Case "DOCMORRIS"
        If (ArtikelDBok) Then
            On Error Resume Next
            DocMorrisRec.Close
            Err.Clear
            On Error GoTo DefErr
'            SQLStr = SQLStr + "SELECT DM.VK AS DM_VK, DM.Automat AS DM_Automat, DM.Gesperrt as DM_GESPERRT"
            SQLStr = SQLStr + "SELECT *"
            SQLStr = SQLStr + " FROM DocMorris AS DM"
            SQLStr = SQLStr + " WHERE DM.pzn = " + clsOpTool.SqlPzn(pzn$)
            DocMorrisRec.Open SQLStr, ArtikelDB1.ActiveConn
            If (DocMorrisRec.EOF = False) Then
                ret% = True
            End If
        End If
End Select

EinlesenDateiDS% = ret%

Call clsError.DefErrPop
End Function

Sub CheckDateiDS(pzn$, h$, AuchStruktur%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("CheckDateiDS")
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

InfoPzn$ = pzn$

If (h$ = "ART") Or (h$ = "AST") Or (h$ = "ASS") Or (h$ = "NNEK") Then
    If (ArtikelDBok) Then
        If (ArtNeu% = False) Then
            ArtDa% = EinlesenDateiDS(pzn$, "ART", AuchStruktur%)
            AstDa% = ArtDa%
            AssDa% = ArtDa%
            NnekDa% = ArtDa%
            ArtNeu% = True
            AstNeu% = True
            AssNeu% = True
            NnekNeu% = True
            If (iMARS) Then
                DocMorrisDa% = EinlesenDateiDS(pzn$, "DOCMORRIS", AuchStruktur%)
            Else
                DocMorrisDa = False
            End If
            DocMorrisNeu = True
        End If
    ElseIf (h$ = "AST") Then
        If (AstNeu% = False) Then
            AstDa% = EinlesenDateiDS(pzn$, h$, AuchStruktur%)
            AstNeu% = True
        End If
    ElseIf (h$ = "ASS") Then
        If (AssNeu% = False) Then
            AssDa% = EinlesenDateiDS(pzn$, h$, AuchStruktur%)
            AssNeu% = True
        End If
    ElseIf (h$ = "NNEK") Then
        If (NnekNeu% = False) Then
            NnekDa% = EinlesenDateiDS(pzn$, h$, AuchStruktur%)
            NnekNeu% = True
        End If
    End If
ElseIf (h$ = "TAXE") Then
    If (TaxeNeu% = False) Then
        TaxeDa% = EinlesenDateiDS(pzn$, h$, AuchStruktur%)
        TaxeNeu% = True
    End If
ElseIf (h$ = "LIEF") Then
    If (LiefNeu% = False) Then
        LiefDa% = EinlesenDateiDS(pzn$, h$, AuchStruktur%)
        LiefNeu% = True
    End If
ElseIf (h$ = "HTAXE") Then
    If (hTaxNeu% = False) Then
        hTaxDa% = EinlesenDateiDS(pzn$, h$, AuchStruktur%)
        hTaxNeu% = True
    End If
ElseIf (h$ = "SONDERPR") Then
    SonderPrDa% = EinlesenDateiDS(pzn$, h$, AuchStruktur%)
End If

Call clsError.DefErrPop
End Sub

Function HoleDateiValue$(iDatei$, iKurz$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("HoleDateiValue$")
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
Dim ret$

ret$ = ""
    
If (iDatei$ = "ART") Then
    If (ArtDa%) Then
        If (ArtikelDBok) Then
            ret$ = ArtikelDB1.FeldWert(iKurz$)
        End If
    End If
ElseIf (iDatei$ = "AST") Then
    If (AstDa%) Then
        If (ArtikelDBok) Then
            ret$ = ArtikelDB1.FeldWert(iKurz$)
        Else
            ret$ = Ast1.Info(iKurz$)
        End If
    ElseIf (iMARS) Then
        If (iKurz = "AVP") Then
            If (DocMorrisDa) Then
'                ret = Format(clsOpTool.CheckNullDouble(DocMorrisRec!DM_vk), "0.00")
                ret$ = DM1.FeldWert("VK")
            End If
        End If
    End If
ElseIf (iDatei$ = "ASS") Then
    If (AssDa%) Then
        If (ArtikelDBok) Then
            ret$ = ArtikelDB1.FeldWert(iKurz$)
        Else
            ret$ = Ass1.Info(iKurz$)
        End If
    End If
    If (UCase$(iKurz$) = "POSLAG") And (FremdPznOk%) Then
'        FremdPznRec.Seek "=", InfoPzn$
        FremdPznRec.Seek "=", Val(InfoPzn$)
        If (FremdPznRec.NoMatch = False) Then
'            ret$ = "*" + Trim(ret$)
            ret$ = ret$ + "*"
        End If
    End If
ElseIf (iDatei$ = "TAXE") Then
    If (TaxeDa%) Then
        If (TaxeAdoDBok) Then
            ret$ = TaxeAdoDB1.FeldWert(iKurz$)
        Else
            ret$ = Taxe1.TaxeValue(iKurz$)
        End If
    End If
ElseIf (iDatei$ = "NNEK") Then
    If (NnekDa%) Then
        If (ArtikelDBok) Then
            ret$ = ArtikelDB1.FeldWert("NNEK_" + iKurz$)
        Else
            ret$ = nnek1.Info(iKurz$)
        End If
    End If
ElseIf (iDatei$ = "LIEF") Then
    If (LiefDa%) Then
        If (LieferantenDBok) Then
            ret$ = LieferantenDB1.FeldWert(iKurz$)
        Else
            ret$ = Lif1.Info(iKurz$)
        End If
    End If
ElseIf (iDatei$ = "HTAXE") Then
    If (hTaxDa%) Then
        ret$ = hTax1.Info(iKurz$)
    End If
ElseIf (iDatei$ = "DM") Then
    If (DocMorrisDa) Then
        ret$ = DM1.FeldWert(iKurz$)
'        If (iKurz = "VK") Then
'            ret = Format(clsOpTool.CheckNullDouble(DocMorrisRec!DM_vk), "0.00")
'        ElseIf (iKurz = "AUTOMAT") Then
'            ret = Format(clsOpTool.CheckNullLong(DocMorrisRec!DM_Automat), "0.00")
'        ElseIf (iKurz = "GESPERRT") Then
'            ret = Format(clsOpTool.CheckNullLong(DocMorrisRec!DM_Gesperrt), "0.00")
'        End If
    End If
End If

HoleDateiValue$ = ret$

Call clsError.DefErrPop
End Function

Function GetLastKey$(KeyLen%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("GetLastKey$")
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

'GetLastKey$ = clsfabs.LastKey(KeyLen%)
GetLastKey = String("A", KeyLen)

Call clsError.DefErrPop
End Function


