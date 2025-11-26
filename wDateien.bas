Attribute VB_Name = "Module1"
Option Explicit

Private Const DefErrModul = "opdateien.bas"

Function FileOpen%(fName$, fAttr$, Optional modus$ = "B", Optional SatzLen% = 100)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("FileOpen%")
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call DefErrAbort
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim Handle%

On Error Resume Next
FileOpen% = False
Handle% = FreeFile


If (fAttr$ = "R") Then
    If (modus$ = "B") Then
        Open fName$ For Binary Access Read Shared As #Handle%
    Else
        Open fName$ For Random Access Read Shared As #Handle% Len = SatzLen%
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
        Open fName$ For Random Access Write As #Handle% Len = SatzLen%
    End If
ElseIf (fAttr$ = "RW") Then
    If (modus$ = "B") Then
        Open fName$ For Binary Access Read Write Shared As #Handle%
    Else
        Open fName$ For Random Access Read Write Shared As #Handle% Len = SatzLen%
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
    Call ProgrammEnde
End If

Call DefErrPop
End Function

