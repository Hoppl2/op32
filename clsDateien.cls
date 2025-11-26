VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
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

Function FileOpen%(fName$, fAttr$, Optional modus$ = "B", Optional SATZLEN% = 100)
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Call DefErrFnc("FileOpen%")
'On Error GoTo DefErr
'GoTo DefErrEnd
'DefErr:
'Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
'Case vbRetry
'  Resume
'Case vbIgnore
'  Resume Next
'End Select
'Call DefErrAbort
'DefErrEnd:
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim Handle%

On Error Resume Next
FileOpen% = False
Handle% = FreeFile


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
ElseIf (fAttr$ = "I") Then
    Open fName$ For Input Access Read Shared As #Handle%
ElseIf (fAttr$ = "O") Then
    Open fName$ For Output Access Write Shared As #Handle%
End If

If (Err = 0) Then
    FileOpen% = Handle%
Else
    Call MsgBox("Fehler" + Str$(Err) + " beim Öffnen von " + fName$ + vbCr + Err.Description, vbCritical, "FileOpen")
'    Call ProgrammEnde
End If

'Call DefErrPop
End Function

Sub iLock(file As Integer, SatzNr&)

Dim LockTime As Date
Dim s As String

On Error GoTo StandardError

If (Para1.MehrPlatz) Then Lock #file, SatzNr&

Exit Sub

'------------------------------------------------------------
StandardError:

If Err = 70 Or Err = 75 Then
  If LockTime = 0 Then LockTime = DateAdd("s", 10, Now)
  If LockTime > Now Then
    'Sleep (1)
    Resume
  End If
End If
Error Err.Number
End Sub

Sub iUnLock(file As Integer, SatzNr&)

If (Para1.MehrPlatz) Then Unlock #file, SatzNr&

End Sub

Function SucheFreienIndex%()
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

End Function

Function IndexOpen%(iName$, iHandle%)
Dim iInd%, ret%

ret% = -1

iInd% = SucheFreienIndex
If (iInd% >= 0) Then
    iHandle% = iInd%
    cmd$ = "O\" + RTrim$(iName$) + "\" + Mid$(Str$(iInd%), 2)
    ret% = clsfabs.Cmnd(cmd$, FabsRecno&)
End If

    
If (ret% <> 0) Then
    Call MsgBox("Index-Fehler" + Str$(ret%) + " bei Befehl " + vbCr + cmd$, vbCritical, "IndexOpen")
End If

IndexOpen% = ret%

End Function

Function IndexClose%(ind%)
Dim iRecno&

cmd$ = "K\" + Mid$(Str$(ind%), 2)
FabsErrf% = clsfabs.Cmnd(cmd$, iRecno&)

IndexClose% = FabsErrf%

End Function

Function IndexSearch%(ind%, key$, FabsRecno&)

cmd$ = "S\1\" + Mid$(Str$(ind%), 2) + "\" + key$
FabsErrf% = clsfabs.Cmnd(cmd$, FabsRecno&)

IndexSearch% = FabsErrf%

End Function

Function IndexGeneric%(ind%, key$, FabsRecno&)

cmd$ = "G\1\" + Mid$(Str$(ind%), 2) + "\" + key$
FabsErrf% = clsfabs.Cmnd(cmd$, FabsRecno&)

IndexGeneric% = FabsErrf%

End Function

Function IndexNext%(ind%, AltRecno&, key$, FabsRecno&)

cmd$ = "A\1\" + Mid$(Str$(AltRecno&), 2) + "\" + Mid$(Str$(ind%), 2) + "\" + key$
FabsErrf% = clsfabs.Cmnd(cmd$, FabsRecno&)

IndexNext% = FabsErrf%

End Function

Function IndexPrevious%(ind%, AltRecno&, key$, FabsRecno&)

cmd$ = "E\1\" + Mid$(Str$(AltRecno&), 2) + "\" + Mid$(Str$(ind%), 2) + "\" + key$
FabsErrf% = clsfabs.Cmnd(cmd$, FabsRecno&)

IndexPrevious% = FabsErrf%

End Function

Function IndexInsert%(ind%, key$, FabsRecno&)

cmd$ = "I\" + Mid$(Str$(ind%), 2) + "\" + key$
FabsErrf% = clsfabs.Cmnd(cmd$, FabsRecno&)

IndexInsert% = FabsErrf%

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

End Function

Function FeldValue(RawData$, fTyp$, fKonv%) As Variant
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

End Function

Function MakeFeld$(ByVal RawData As Variant, fTyp$, fKonv%)
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

End Function

Function FirstLettersUcase(ByVal s As String) As String
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Call DefErrFnc("FirstLettersUcase")
'On Error GoTo DefErr
'GoTo DefErrEnd
'DefErr:
'Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
'Case vbRetry
'  Resume
'Case vbIgnore
'  Resume Next
'End Select
'Call DefErrAbort
'DefErrEnd:
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
  
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

'Call DefErrPop
End Function

