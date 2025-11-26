Attribute VB_Name = "modError"
Option Explicit

Private Const DefErrModul = "DefErr.bas"

Dim DefErrFncStr(50) As String
Dim DefErrStk As Integer

Function DefErrAnswer(src As String, num As Integer, desc As String, modul As String) As Integer
Dim WINERROR%, ret%
Dim s As String

s = "Es ist ein Fehler aufgetreten!" + Chr(13) + Chr(10) + Chr(13) + Chr(10)
s = s + "Programm:" + Chr(9) + UCase(src) + Chr(13) + Chr(10)
s = s + "Modul:" + Chr(9) + Chr(9) + UCase(modul) + Chr(13) + Chr(10)
s = s + "Zähler:" + Chr(9) + Chr(9) + CStr(Abs(App.PrevInstance)) + Chr(13) + Chr(10)
s = s + "Funktion:" + Chr(9) + Chr(9) + DefErrFncStr(DefErrStk) + Chr(13) + Chr(10)
s = s + "Nummer:" + Chr(9) + Chr(9) + CStr(num) + Chr(13) + Chr(10)
s = s + "Text:" + Chr(9) + Chr(9) + desc + Chr(13) + Chr(10)
s = s + "Uhrzeit:" + Chr(9) + Chr(9) + CStr(Time)
ret% = MsgBox(s, vbCritical Or vbAbortRetryIgnore Or vbDefaultButton2, "Problem")

s = ""
s = s + Format(Now, "dd.mm.yyyy") + "," + Format(Now, "hh:nn:ss")
s = s + "," + CStr(num)
s = s + "," + CStr(0)
s = s + "," + App.Path + "\" + App.EXEName + ".EXE"
s = s + "," + Format(App.Major) + "." + Format(App.Minor) + "." + Format(App.Revision)
s = s + "," + Para1.user
s = s + "," + src
s = s + "," + modul
s = s + "," + DefErrFncStr(DefErrStk)
s = s + "," + CStr(ret%)

WINERROR% = FreeFile
Open "WINERROR.TXT" For Append As #WINERROR%
Print #WINERROR%, s
Close #WINERROR%

DefErrAnswer = ret%

End Function

Sub DefErrAnswer2(src As String, num As Integer, desc As String, modul As String)
Dim s As String

s = "Es ist ein Fehler aufgetreten!" + Chr(13) + Chr(10) + Chr(13) + Chr(10)
s = s + "Programm:" + Chr(9) + UCase(src) + Chr(13) + Chr(10)
s = s + "Modul:" + Chr(9) + Chr(9) + UCase(modul) + Chr(13) + Chr(10)
s = s + "Zähler:" + Chr(9) + Chr(9) + CStr(Abs(App.PrevInstance)) + Chr(13) + Chr(10)
s = s + "Funktion:" + Chr(9) + Chr(9) + DefErrFncStr(DefErrStk) + Chr(13) + Chr(10)
s = s + "Nummer:" + Chr(9) + Chr(9) + CStr(num) + Chr(13) + Chr(10)
s = s + "Text:" + Chr(9) + Chr(9) + desc + Chr(13) + Chr(10)
s = s + "Uhrzeit:" + Chr(9) + Chr(9) + CStr(Time)
Call MsgBox(s, vbCritical Or vbOKOnly, "Problem")

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

Sub DefErrPop()

If DefErrStk > 0 Then DefErrStk = DefErrStk - 1

End Sub

Sub DefErrAbort()
'End
End Sub

