VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsWbestk2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const DefErrModul = "WBESTK2.CLS"

Sub ManuellSpeichern(Optional MitAuslesen% = False)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("ManuellSpeichern")
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
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim erg%

If (ManuellSsatz% = 0) Then
    If (ManuellPzn$ = "9999999") Or (ManuellAsatz% = 0) Then
        ManuellBm% = -ManuellBm%: ManuellNm% = -ManuellNm%
'    Else
'        erg% = MsgBox("Ist dieser Artikel bereits auf Lager ?", vbYesNo Or vbInformation)
'        If (erg% = vbNo) Then
'            ManuellBm% = -ManuellBm%: ManuellNm% = -ManuellNm%
'        End If
    End If
End If

Call NeuerDS

Call DefErrPop
End Sub

