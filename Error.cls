VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsError"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim LockTime As Date

Private Const DefErrModul = "ERROR.CLS"

'Dim DefErrFncStr(50) As String
'Dim DefErrStk As Integer

Function DefErrAnswer(src As String, num As Long, desc As String, modul As String, Optional FehlerModus% = 0) As Integer
Dim WINERROR%, ret%, AutoIgnore%
Dim i%, ind%
Dim AbError&, BisError&
Dim s As String, s2$, aZeit$

If (Err.Number >= 3000) And (Err.Number <> 3020) And (Err.Number <> 3061) And (Err.Number <> 3075) And (Err.Number <> 3167) Then
    If (LockTime = 0) Then
        LockTime = DateAdd("s", 5, Now)
    End If
    If (LockTime > Now) Then
        ret% = vbRetry
        DefErrAnswer = ret%
        Exit Function
    End If
End If
LockTime = 0

AutoIgnore = 0
If (FehlerModus = 9999) Then
    AutoIgnore = True
    FehlerModus = 0
Else
    s = iDefErrAutoErrors
    If (Right$(s, 1) <> ",") Then
        s = s + ","
    End If
    Do
        ind% = InStr(s, ",")
        If (ind% > 0) Then
            s2 = Left$(s, ind% - 1)
            s = Mid$(s, ind% + 1)
        
            ind% = InStr(s2, "-")
            If (ind% > 0) Then
                AbError = Val(Left$(s2, ind% - 1))
                BisError = Val(Mid$(s2, ind% + 1))
            Else
                AbError = Val(s2)
                BisError = AbError
            End If
            If (num >= AbError) And (num <= BisError) Then
                AutoIgnore = True
                Exit Do
            End If
        Else
            Exit Do
        End If
    Loop
End If

On Error Resume Next
src = ProjektForm.Caption
ind% = InStr(src, "-")
If (ind% > 0) Then
    src = Left$(src, ind% - 1)
End If
aZeit$ = Format(Now, "hh:nn:ss")

s = ""
s = s + Format(Now, "dd.mm.yyyy") + "," + aZeit$
s = s + "," + CStr(num)
s = s + "," + CStr(0)
s = s + "," + App.Path + "\" + App.EXEName + ".EXE"
s = s + "," + Format(App.Major) + "." + Format(App.Minor) + "." + Format(App.Revision)
s = s + "," + Para1.user
s = s + "," + src
s = s + "," + modul
s = s + "," + DefErrFncStr(DefErrStk)

WINERROR% = FreeFile
Open "WINERROR.TXT" For Append As #WINERROR%
For i% = 1 To DefErrStk
    s2$ = DefErrModStr(i%) + ":  " + DefErrFncStr(i%)
    If (InStr(UCase(s2$), "INFOBEREICH") > 0) Then
        AutoIgnore = True
    ElseIf (InStr(UCase(s2$), "ARBEITBEREICH") > 0) Then
        AutoIgnore = True
    ElseIf (InStr(UCase(s2$), "EINZELSATZ") > 0) Then
        AutoIgnore = True
    End If
    Print #WINERROR%, "    " + s2$
Next i%
Print #WINERROR%, "    " + "Fehler " + CStr(num) + ":  "; desc
Print #WINERROR%, s
Close #WINERROR%


s = "Folgendes Ereignis erfordert Ihr Eingreifen:" + vbCrLf + vbCrLf
s = s + desc + " (" + CStr(num) + ")"
s = s + vbCrLf
s = s + "(Interne Info: " + UCase(modul) + " / Funktion " + DefErrFncStr(DefErrStk) + ")"
s = s + vbCrLf
s = s + vbCrLf
If num = 482 Then
    s = s + "Drucker: " + Printer.DeviceName + vbCrLf
    s = s + "Prüfen Sie, ob der Drucker online und Papier eingelegt ist; schalten Sie ihn ev. aus und wieder ein."
    s = s + vbCrLf
End If
s = s + vbCrLf


If ((num >= 52) And (num <= 76)) Or (num >= 3000) Then
    s = s + "Bitte verständigen Sie die OPTIPHARM-Hotline"
Else
    s = s + "Sollte das Problem häufiger auftreten, verständigen Sie bitte die OPTIPHARM-Hotline."
End If
s = s + vbCrLf + vbCrLf

If (AutoIgnore) Then
    ret% = vbIgnore
Else
    Call ProjektForm.FlexKurzInfo(0)
    If (FehlerModus%) Then
        s = s + "PROGRAMM WIRD BEENDET !"
        ret% = clsDialog.MessageBox(s, vbCritical Or vbOKOnly Or vbSystemModal, "Problem")
        ret = vbAbort
    Else
        If (num >= 70 And num <= 75) Or num = 482 Or num = 3197 Or num = 3260 Or num = 3356 Then
            s = s + "Möchten Sie den Programmschritt WIEDERHOLEN?" + vbCr
        Else
            s = s + "Möchten Sie das Problem IGNORIEREN und mit der Programmausführung fortfahren?" + vbCr
        End If
        s = s + "Nein = Programmende"
        If (num = 7) Or (num = 372) Then
            ret% = MsgBox(s, vbQuestion Or vbYesNo Or vbDefaultButton2 Or vbSystemModal, CStr(Time) + " " + UCase(src))
        Else
            ret% = clsDialog.MessageBox(s, vbQuestion Or vbYesNo Or vbDefaultButton2 Or vbSystemModal, CStr(Time) + " " + UCase(src))
        End If
        If ret% = vbNo Then
            ret% = vbCancel
        ElseIf ret% = vbYes Then
            ret% = vbIgnore
        End If
    
        If num = 70 And ret% = vbIgnore Then
          ret% = vbRetry
        End If
    End If
    Call ProjektForm.FlexKurzInfo(1)
End If

s2$ = ""
If (AutoIgnore) Then
    s2$ = "A"
ElseIf (ret% = vbAbort) Then
    s2$ = "B"
ElseIf (ret% = vbRetry) Then
    s2$ = "W"
ElseIf (ret% = vbIgnore) Then
    s2$ = "I"
End If

s = ""
s = s + s2$ + " " + Format(Now, "hh:nn:ss") + "," + aZeit$
s = s + "," + CStr(num)
s = s + "," + CStr(0)
s = s + "," + App.Path + "\" + App.EXEName + ".EXE"
s = s + "," + Format(App.Major) + "." + Format(App.Minor) + "." + Format(App.Revision)
s = s + "," + Para1.user
s = s + "," + src
s = s + "," + modul
s = s + "," + DefErrFncStr(DefErrStk)

WINERROR% = FreeFile
Open "WINERROR.TXT" For Append As #WINERROR%
'For i% = 1 To DefErrStk
'    Print #WINERROR%, DefErrModStr(i%) + ":  "; DefErrFncStr(i%)
'Next i%
Print #WINERROR%, s
Close #WINERROR%

DefErrAnswer = ret%

End Function

Function DefErrAnswerOld(src As String, num As Long, desc As String, modul As String, Optional FehlerModus% = 0) As Integer
Dim WINERROR%, ret%
Dim i%, ind%
Dim s As String, s2$, aZeit$

If (Err.Number >= 3000) Then
    If (LockTime = 0) Then
        LockTime = DateAdd("s", 5, Now)
    End If
    If (LockTime > Now) Then
        ret% = vbRetry
        DefErrAnswerOld = ret%
        Exit Function
    End If
End If
LockTime = 0

On Error Resume Next
src = ProjektForm.Caption
ind% = InStr(src, "-")
If (ind% > 0) Then
    src = Left$(src, ind% - 1)
End If
aZeit$ = Format(Now, "hh:nn:ss")

s = ""
s = s + Format(Now, "dd.mm.yyyy") + "," + aZeit$
s = s + "," + CStr(num)
s = s + "," + CStr(0)
s = s + "," + App.Path + "\" + App.EXEName + ".EXE"
s = s + "," + Format(App.Major) + "." + Format(App.Minor) + "." + Format(App.Revision)
s = s + "," + Para1.user
s = s + "," + src
s = s + "," + modul
s = s + "," + DefErrFncStr(DefErrStk)

WINERROR% = FreeFile
Open "WINERROR.TXT" For Append As #WINERROR%
For i% = 1 To DefErrStk
    Print #WINERROR%, "    " + DefErrModStr(i%) + ":  "; DefErrFncStr(i%)
Next i%
Print #WINERROR%, "    " + "Fehler " + CStr(num) + ":  "; desc
Print #WINERROR%, s
Close #WINERROR%


s = "Es ist ein Fehler aufgetreten!" + Chr(13) + Chr(10) + Chr(13) + Chr(10)
s = s + "Programm:" + Chr(9) + UCase(src) + Chr(13) + Chr(10)
s = s + "Modul:" + Chr(9) + Chr(9) + UCase(modul) + Chr(13) + Chr(10)
s = s + "Zähler:" + Chr(9) + Chr(9) + CStr(Abs(App.PrevInstance)) + Chr(13) + Chr(10)
s = s + "Funktion:" + Chr(9) + Chr(9) + DefErrFncStr(DefErrStk) + Chr(13) + Chr(10)
s = s + "Nummer:" + Chr(9) + Chr(9) + CStr(num) + Chr(13) + Chr(10)
s = s + "Text:" + Chr(9) + Chr(9) + desc + Chr(13) + Chr(10)
s = s + "Uhrzeit:" + Chr(9) + Chr(9) + CStr(Time)
Call ProjektForm.FlexKurzInfo(0)
If (FehlerModus%) Then
    s = s + Chr(13) + Chr(10) + Chr(13) + Chr(10)
    s = s + "PROGRAMM WIRD BEENDET !"
    ret% = MsgBox(s, vbCritical Or vbOKOnly Or vbSystemModal, "Problem")
Else
    If (num = 9) Then
        ret% = MsgBox(s, vbCritical Or vbAbortRetryIgnore Or vbDefaultButton2 Or vbSystemModal, "Problem")
    Else
        ret% = MsgBox(s, vbCritical Or vbAbortRetryIgnore Or vbDefaultButton2 Or vbSystemModal, "Problem")
    End If
End If
Call ProjektForm.FlexKurzInfo(1)

s2$ = ""
If (ret% = vbAbort) Then
    s2$ = "B"
ElseIf (ret% = vbRetry) Then
    s2$ = "W"
ElseIf (ret% = vbIgnore) Then
    s2$ = "I"
End If

s = ""
s = s + s2$ + " " + Format(Now, "hh:nn:ss") + "," + aZeit$
s = s + "," + CStr(num)
s = s + "," + CStr(0)
s = s + "," + App.Path + "\" + App.EXEName + ".EXE"
s = s + "," + Format(App.Major) + "." + Format(App.Minor) + "." + Format(App.Revision)
s = s + "," + Para1.user
s = s + "," + src
s = s + "," + modul
s = s + "," + DefErrFncStr(DefErrStk)

WINERROR% = FreeFile
Open "WINERROR.TXT" For Append As #WINERROR%
'For i% = 1 To DefErrStk
'    Print #WINERROR%, DefErrModStr(i%) + ":  "; DefErrFncStr(i%)
'Next i%
Print #WINERROR%, s
Close #WINERROR%

DefErrAnswerOld = ret%

End Function

Sub ErrorDebug(Optional OkStr$ = "  ")
Dim i%, ind%, WINERROR%
Dim s As String, src As String, aZeit$

If (Para1.F70Debug) Then
    src = ProjektForm.Caption
    ind% = InStr(src, "-")
    If (ind% > 0) Then
        src = Left$(src, ind% - 1)
    End If
    aZeit$ = Format(Now, "hh:nn:ss")
    
    s = Left$(Para1.user + Space$(2), 2) + " "
    s = s + OkStr$ + " "
    s = s + DefErrFncStr(DefErrStk)
    s = s + "  (" + src + ")   "
    s = s + Format(Now, "dd.mm.yyyy") + "," + aZeit$
    
    WINERROR% = FreeFile
    Open "WINERROR.TXT" For Append As #WINERROR%
    Print #WINERROR%, "    " + "    ";
    For i% = 1 To (DefErrStk - 1)
        Print #WINERROR%, DefErrModStr(i%) + ":"; DefErrFncStr(i%) + "  ";
    Next i%
    Print #WINERROR%, " "
    Print #WINERROR%, "   " + "    " + s
    Close #WINERROR%
End If

End Sub

Sub DefErrAnswer2(src As String, num As Long, desc As String, modul As String)
Dim s As String

s = "Es ist ein Fehler aufgetreten!" + Chr(13) + Chr(10) + Chr(13) + Chr(10)
s = s + "Programm:" + Chr(9) + UCase(src) + Chr(13) + Chr(10)
s = s + "Modul:" + Chr(9) + Chr(9) + UCase(modul) + Chr(13) + Chr(10)
s = s + "Zähler:" + Chr(9) + Chr(9) + CStr(Abs(App.PrevInstance)) + Chr(13) + Chr(10)
s = s + "Funktion:" + Chr(9) + Chr(9) + DefErrFncStr(DefErrStk) + Chr(13) + Chr(10)
s = s + "Nummer:" + Chr(9) + Chr(9) + CStr(num) + Chr(13) + Chr(10)
s = s + "Text:" + Chr(9) + Chr(9) + desc + Chr(13) + Chr(10)
s = s + "Uhrzeit:" + Chr(9) + Chr(9) + CStr(Time)
Call clsDialog.MessageBox(s, vbCritical Or vbOKOnly, "Problem")

End Sub

Sub DefErrFnc(s As String)

Dim i As Integer

If DefErrStk < 50 Then
  DefErrStk = DefErrStk + 1
  DefErrFncStr(DefErrStk) = s
Else
  For i = 2 To 50
    DefErrFncStr(i - 1) = DefErrFncStr(i)
  Next i
  DefErrFncStr(50) = s
End If

End Sub

Sub DefErrMod(s As String)

DefErrModStr(DefErrStk) = s

End Sub

Sub DefErrPop()

If DefErrStk > 0 Then DefErrStk = DefErrStk - 1

End Sub

Sub DefErrAbort()
'End
End Sub

