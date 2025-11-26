Attribute VB_Name = "modWwBereiche"
Option Explicit

Private Const DefErrModul = "wwbereiche.bas"

Sub EditBereichsFarbe(dlg As CommonDialog, Index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("EditBereichsFarbe")
Call DefErrMod(DefErrModul)
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
Dim l&

On Error Resume Next
If (Index = 0) Then
    dlg.Color = FarbeArbeit&
Else
    dlg.Color = FarbeInfo&
End If
dlg.CancelError = True
dlg.Flags = cdlCCFullOpen + cdlCCRGBInit
Call dlg.ShowColor
If (Err = 0) Then
    If (Index = 0) Then
        FarbeArbeit& = dlg.Color
        l& = WritePrivateProfileString(UserSection$, "FarbeArbeit", Hex$(FarbeArbeit&), WINWAWI_INI)
    Else
        FarbeInfo& = dlg.Color
        l& = WritePrivateProfileString(UserSection$, "FarbeInfo", Hex$(FarbeInfo&), WINWAWI_INI)
    End If
    Call InitAlleBereichsFarben
End If

Call DefErrPop
End Sub

