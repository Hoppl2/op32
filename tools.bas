Attribute VB_Name = "modTools"

'Filename: Tools.bas
'
Option Explicit

Declare Function CharToOem Lib "user32" Alias "CharToOemA" (ByVal lpszSrc As String, ByVal lpszDst As String) As Long
Declare Function OemToChar Lib "user32" Alias "OemToCharA" (ByVal lpszSrc As String, ByVal lpszDst As String) As Long

Declare Sub DxToIEEEd Lib "mbfiee32.dll" (mbf As Double)
Declare Sub DxToIEEEs Lib "mbfiee32.dll" (mbf As Single)
Declare Sub DxToMBFd Lib "mbfiee32.dll" (ieee As Double)
Declare Sub DxToMBFs Lib "mbfiee32.dll" (ieee As Single)

Type str8type
  s As String * 8
End Type

Type num8type
  z As Double
End Type

Type str4type
  s As String * 4
End Type

Type num4type
  z As Single
End Type

Dim s8 As str8type
Dim s4 As str4type
Dim n8 As num8type
Dim n4 As num4type

Private Const DefErrModul = "TOOLS.BAS"

Function dDatum(datum%) As Date
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("dDatum")
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
Dim s$
If datum% > 0 Then
  s$ = sDate(datum%)
  s$ = Left$(s$, 2) + "." + Mid$(s$, 3, 2) + "." + Mid$(s$, 5, 2)
  dDatum = CDate(s$)
End If
Call DefErrPop
End Function

Public Function sCheckNull(f As Field) As String
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("sCheckNull")
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
On Error Resume Next
If Not IsNull(f) Then
  sCheckNull = f
End If
Call DefErrPop

End Function


Public Function CheckDatum(ByVal s As String) As String
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("CheckDatum")
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

Dim i As Integer
Dim d As String
Dim Tag As String
Dim Monat As String
Dim Jahr As String

CheckDatum = ""
s = Trim(s)
If IsDate(s) Then
  CheckDatum = Format(CDate(s), "dd.mm.yyyy")
  Call DefErrPop
  Exit Function
End If

For i = 1 To Len(s)
  If InStr("0123456789.-", Mid(s, i, 1)) > 0 Then
    d = d + Mid(s, i, 1)
  End If
Next i
i = InStr(d, ".")
If i = 0 Then i = InStr(d, "-")
If i > 0 Then
  Tag = Left(d, i - 1)
  d = Mid(d, i + 1)
Else
  Tag = Left(d, 2)
  d = Mid(d, 3)
End If
i = InStr(d, ".")
If i = 0 Then i = InStr(d, "-")
If i > 0 Then
  Monat = Left(d, i - 1)
  Jahr = Mid(d, i + 1)
Else
  Monat = Left(d, 2)
  Jahr = Mid(d, 3)
End If

If Len(Tag) <= 2 And Len(Monat) <= 2 And Len(Jahr) >= 2 And Len(Jahr) <= 4 Then
  d = CStr(Val(Tag)) + "." + CStr(Val(Monat)) + "." + CStr(Val(Jahr))
  If IsDate(d) Then CheckDatum = Format(CDate(d), "dd.mm.yyyy")
End If
Call DefErrPop
End Function


Public Function dCheckNull(f As Field) As Double
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("dCheckNull")
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
On Error Resume Next
If Not IsNull(f) Then
  dCheckNull = f
End If
Call DefErrPop
End Function


Function Chr0RechtsWeg(ByVal s As String) As String

'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("Chr0RechtsWeg")
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
If (Err.Number = (999 + vbObjectError)) Then End
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
'On Error GoTo 0
'Err.Raise 999 + vbObjectError, "DSK.VBP", "Fehler in DLL"
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
While Right(s, 1) = Chr(0)
  s = Left(s, Len(s) - 1)
Wend
Chr0RechtsWeg = s

Call DefErrPop
End Function


Function CVAEDatum(ByVal d As Long) As Date

'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("CVAEDatum")
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
If (Err.Number = (999 + vbObjectError)) Then End
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
'On Error GoTo 0
'Err.Raise 999 + vbObjectError, "DSK.VBP", "Fehler in DLL"
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
CVAEDatum = CDate(d) - 10227&

Call DefErrPop
End Function

Function fnint(x As Long) As Integer

fnint = Int(x + (x > 32767) * 65536)

End Function

Function fnx(x As Double) As Double

fnx = Sgn(x) * CDbl(Int(Abs(x) * 100# + 0.501) / 100#)

End Function

Function IstNumerisch(s As String) As Boolean

'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("IstNumerisch")
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
If (Err.Number = (999 + vbObjectError)) Then End
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
'On Error GoTo 0
'Err.Raise 999 + vbObjectError, "DSK.VBP", "Fehler in DLL"
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim i As Integer

IstNumerisch = True

For i = 1 To Len(s)
  If InStr("0123456789", Mid(s, i, 1)) = 0 Then
    IstNumerisch = False
    Exit For
  End If
Next i

Call DefErrPop
End Function


Function CheckPreis(Preis As String) As String

'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("CheckPreis")
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
If (Err.Number = (999 + vbObjectError)) Then End
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
'On Error GoTo 0
'Err.Raise 999 + vbObjectError, "DSK.VBP", "Fehler in DLL"
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim s As String
Dim i As Integer

For i = 1 To Len(Preis)
  If InStr("0123456789,.-+", Mid(Preis, i, 1)) > 0 Then
    s = s + Mid(Preis, i, 1)
    If Val(s) > 9999999999# Then Exit For
  End If
Next i

CheckPreis = s
Call DefErrPop
End Function

Function Long2Int(zahl As Long) As Integer

'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("Long2Int")
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
If (Err.Number = (999 + vbObjectError)) Then End
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
'On Error GoTo 0
'Err.Raise 999 + vbObjectError, "DSK.VBP", "Fehler in DLL"
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
While zahl > 32767
  zahl = zahl - 65536
Wend
Long2Int = zahl

Call DefErrPop
End Function

Public Function MyUCASE(sOrgText As String) As String

'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("MyUCASE")
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
If (Err.Number = (999 + vbObjectError)) Then End
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
'On Error GoTo 0
'Err.Raise 999 + vbObjectError, "DSK.VBP", "Fehler in DLL"
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim i%, c$, s, UPCHARS As String, LOCHARS As String

UPCHARS = "ABCDEFGHIJKLMNOPQRSTUVWXYZÄÖÜS"
LOCHARS = "abcdefghijklmnopqrstuvwxyzäöüß"

s = sOrgText
For i = 1 To Len(s)
  c = Mid(s, i, 1)
  c = Mid(UPCHARS + c, InStr(LOCHARS + c, c), 1)
  Mid(s, i, 1) = c
Next i
MyUCASE = s

Call DefErrPop
End Function

Public Function Oem2Ansi(sOrgText As String) As String

'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("Oem2Ansi")
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
If (Err.Number = (999 + vbObjectError)) Then End
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
'On Error GoTo 0
'Err.Raise 999 + vbObjectError, "DSK.VBP", "Fehler in DLL"
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim s As String
Dim l As Long

s = Space(Len(sOrgText))
If Len(sOrgText) > 0 Then
  l = OemToChar(sOrgText, s)
End If
Oem2Ansi = s

Call DefErrPop
End Function

Public Function Ansi2Oem(sOrgText As String) As String

'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("Ansi2Oem")
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
If (Err.Number = (999 + vbObjectError)) Then End
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
'On Error GoTo 0
'Err.Raise 999 + vbObjectError, "DSK.VBP", "Fehler in DLL"
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim s As String
Dim l As Long

s = Space(Len(sOrgText))
If Len(sOrgText) > 0 Then
  l = CharToOem(sOrgText, s)
End If
Ansi2Oem = s

Call DefErrPop
End Function

Public Function CVD(s As String) As Double

'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("CVD")
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
If (Err.Number = (999 + vbObjectError)) Then End
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
'On Error GoTo 0
'Err.Raise 999 + vbObjectError, "DSK.VBP", "Fehler in DLL"
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim ieee As Double

Call CopyMemory(ieee, ByVal s, 8)
Call DxToIEEEd(ieee)
CVD = ieee

Call DefErrPop
End Function

Public Function MKD(ByVal ieee As Double) As String

'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("MKD")
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
If (Err.Number = (999 + vbObjectError)) Then End
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
'On Error GoTo 0
'Err.Raise 999 + vbObjectError, "DSK.VBP", "Fehler in DLL"
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim mfp As String * 8

Call DxToMBFd(ieee)
Call CopyMemory(ByVal mfp, ieee, 8)
MKD = mfp

Call DefErrPop
End Function

Public Function CVS(s As String) As Single

'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("CVS")
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
If (Err.Number = (999 + vbObjectError)) Then End
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
'On Error GoTo 0
'Err.Raise 999 + vbObjectError, "DSK.VBP", "Fehler in DLL"
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim ieee As Single

Call CopyMemory(ieee, ByVal s, 4)
Call DxToIEEEs(ieee)
CVS = ieee

Call DefErrPop
End Function

Public Function MKS(ByVal ieee As Single) As String

'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("MKS")
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
If (Err.Number = (999 + vbObjectError)) Then End
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
'On Error GoTo 0
'Err.Raise 999 + vbObjectError, "DSK.VBP", "Fehler in DLL"
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim mfp As String * 4

Call DxToMBFs(ieee)
Call CopyMemory(ByVal mfp, ieee, 4)
MKS = mfp

Call DefErrPop
End Function

Public Function CVL&(sZahl$)

'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("CVL&")
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
If (Err.Number = (999 + vbObjectError)) Then End
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
'On Error GoTo 0
'Err.Raise 999 + vbObjectError, "DSK.VBP", "Fehler in DLL"
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim l As Long
Call CopyMemory(l, ByVal sZahl$, 4)
CVL = l

Call DefErrPop
End Function

Public Function MKL(ByVal l As Long) As String

'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("MKL")
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
If (Err.Number = (999 + vbObjectError)) Then End
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
'On Error GoTo 0
'Err.Raise 999 + vbObjectError, "DSK.VBP", "Fehler in DLL"
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim s As String * 4
Call CopyMemory(ByVal s, l, 4)
MKL = s

Call DefErrPop
End Function

Public Function CVI(s As String) As Integer

'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("CVI")
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
If (Err.Number = (999 + vbObjectError)) Then End
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
'On Error GoTo 0
'Err.Raise 999 + vbObjectError, "DSK.VBP", "Fehler in DLL"
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim i As Integer
Call CopyMemory(i, ByVal s, 2)
CVI = i

Call DefErrPop
End Function

Public Function MKI(ByVal i As Integer) As String

'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("MKI")
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
If (Err.Number = (999 + vbObjectError)) Then End
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
'On Error GoTo 0
'Err.Raise 999 + vbObjectError, "DSK.VBP", "Fehler in DLL"
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim s As String * 2
Call CopyMemory(ByVal s, i, 2)
MKI = s

Call DefErrPop
End Function

Function Bcd2ascii(ByVal bcd As String, Stellen As Integer) As String

'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("Bcd2ascii")
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
If (Err.Number = (999 + vbObjectError)) Then End
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
'On Error GoTo 0
'Err.Raise 999 + vbObjectError, "DSK.VBP", "Fehler in DLL"
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim i As Integer
Dim ascii As String
'ascii = Ergebnis Zahlenstring
'bcd = binärcodierte Zahl
'stellen = Anzahl der Stellen in Ergebnisfeld
'z.B.: stellen = 3   bcd = chr(&H45) + chr(&H67) => ascii = "456"

ascii = String(Stellen, "0")
If bcd <> String(Len(bcd), "*") Then
  For i = 1 To Stellen Step 2
    Mid(ascii, i, 1) = Chr(Int(Asc(Mid(bcd, Int(i / 2) + 1, 1)) / 16) + &H30)
    If i < Stellen Then
      Mid(ascii, i + 1, 1) = Chr((Asc(Mid(bcd, Int(i / 2) + 1, 1)) Mod 16) + &H30)
    End If
  Next i
End If
Bcd2ascii = ascii

Call DefErrPop
End Function

Function Ascii2bcd(ByVal ascii As String, Stellen As Integer) As String

'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("Ascii2bcd")
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
If (Err.Number = (999 + vbObjectError)) Then End
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
'On Error GoTo 0
'Err.Raise 999 + vbObjectError, "DSK.VBP", "Fehler in DLL"
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim i As Integer
Dim bcd As String
Dim bcdwert As Integer

ascii = Right$(String$(Stellen, "0") + ascii, Stellen)
For i = 1 To Stellen
  If InStr("0123456789", Mid$(ascii, i, 1)) = 0 Then
    Mid$(ascii, i, 1) = "0"
  End If
Next i

bcd = String$(Int((Stellen + 1)) / 2, 0)
For i = 1 To Stellen Step 2
  bcdwert = (Asc(Mid(ascii, i, 1)) - &H30) * 16
  If i < Stellen Then
    bcdwert = bcdwert + Asc(Mid$(ascii, i + 1, 1)) - &H30
  End If
  Mid$(bcd$, Int(i / 2) + 1, 1) = Chr$(bcdwert)
Next i

Ascii2bcd = bcd

Call DefErrPop
End Function

Sub EanPruef(EAN As String)

'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("EanPruef")
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
If (Err.Number = (999 + vbObjectError)) Then End
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
'On Error GoTo 0
'Err.Raise 999 + vbObjectError, "DSK.VBP", "Fehler in DLL"
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim i As Integer, Pruefsumme As Integer, PruefZiffer As Integer

If Val(EAN) > 0 And Len(EAN) = 13 Then
  Pruefsumme = 0
  For i = 1 To 12
    Pruefsumme = Pruefsumme + Val(Mid(EAN, i, 1)) * (1 + 2 * ((i + 1) Mod 2))
  Next i
  PruefZiffer = 10 - (Pruefsumme Mod 10)
  Mid(EAN, 13, 1) = Right(Str(PruefZiffer), 1)
End If

Call DefErrPop
End Sub

Public Function fopen(FileName As String, openmode As String) As Integer

'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("fopen")
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
If (Err.Number = (999 + vbObjectError)) Then End
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
'On Error GoTo 0
'Err.Raise 999 + vbObjectError, "DSK.VBP", "Fehler in DLL"
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim fhTmp As Integer
Dim s As String
Dim sharelock As Boolean
Dim RecLen As Integer
Dim erg As Integer

On Error Resume Next

openmode = UCase(openmode)
fhTmp = FreeFile
If InStr(openmode, "I") > 0 Then
  Open FileName For Input As #fhTmp
ElseIf InStr(openmode, "O") > 0 Then
  Open FileName For Output As #fhTmp
ElseIf InStr(openmode, "A") > 0 Then
  Open FileName For Append As #fhTmp
ElseIf InStr(openmode, "W") > 0 Then
  If InStr(openmode, "L") > 0 Then
    Open FileName For Binary Access Read Write Lock Read Write As #fhTmp
  Else
    Open FileName For Binary Access Read Write Shared As #fhTmp
    'GS 30.3.98 Bei Neuanlage ist Datei immer locked -> schließen, öffnen
    If LOF(fhTmp) = 0 Then
      Close #fhTmp
      Open FileName For Binary Access Read Write Shared As #fhTmp
    End If
    sharelock = True
  End If

Else
  If InStr(openmode, "L") > 0 Then
    Open FileName For Binary Access Read Lock Read Write As #fhTmp
  Else
    'GS 30.3.98 Bei Neuanlage ist Datei immer locked -> schließen, öffnen
    Open FileName For Binary Access Read Shared As #fhTmp
    If LOF(fhTmp) = 0 Then
      Close #fhTmp
      Open FileName For Binary Access Read Shared As #fhTmp
    End If
    sharelock = True
  End If
End If
fopen = 0
If Err.Number = 0 Then
  fopen = fhTmp
  If UBound(OpendFileName) < fhTmp Then
    ReDim Preserve OpendFileName(fhTmp)
  End If
  OpendFileName(fhTmp) = FileName
  If sharelock Then
    iiLock fhTmp, 1, 1
    Unlock fhTmp, 1 To 1
  End If
Else
  If Err.Number <> 53 Then
    s = s + "Fehler " + Str(Err.Number) + " beim Öffnen von " + FileName + "(" + openmode + ")" + Chr(13) + Chr(10)
    erg = iMsgBox(s, vbCritical Or vbSystemModal)
  End If
End If

Call DefErrPop
End Function

Function uFormat(ByVal Wert As Double, ByVal maske As String) As String

'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("uFormat")
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
If (Err.Number = (999 + vbObjectError)) Then End
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
'On Error GoTo 0
'Err.Raise 999 + vbObjectError, "DSK.VBP", "Fehler in DLL"
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim s As String
Dim i As Integer
Dim Stellen As Integer

Stellen = Len(maske)
s = Format(Abs(Wert), maske)
s = LTrim(s)
If Wert < 0 Then s = "-" + s
''GS: LTRIM genügt nicht bei negativen Zahlen, da minus ganz vorne steht --> Leerzeichen fressen
'i = InStr(s, " ")
'While Len(s) > Stellen And i > 0
'  s = Left(s, i - 1) + Mid(s, i + 1)
'  i = InStr(s, " ")
'Wend
If Len(s) < Stellen Then s = Right(Space(Stellen) + s, Stellen)
i = InStr(s, ","): If i > 0 Then Mid(s, i, 1) = "."
uFormat = s

Call DefErrPop
End Function

Function CVDatum(ByVal s As String) As Date

'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("CVDatum")
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
If (Err.Number = (999 + vbObjectError)) Then End
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
'On Error GoTo 0
'Err.Raise 999 + vbObjectError, "DSK.VBP", "Fehler in DLL"
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
CVDatum = 26298& + CLng(CVI(Right(Chr(0) + s, 1) + Left(s + Chr(0), 1)))

Call DefErrPop
End Function

Function MKDatum(ByVal d As Date) As String

'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("MKDatum")
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
If (Err.Number = (999 + vbObjectError)) Then End
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
'On Error GoTo 0
'Err.Raise 999 + vbObjectError, "DSK.VBP", "Fehler in DLL"
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim s As String

s = MKI(CLng(d) - 26298&)

MKDatum = Right(s, 1) + Left(s, 1)

Call DefErrPop
End Function

Sub SeparateKomma(ByVal s As String, argv() As Variant)

'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("SeparateKomma")
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
If (Err.Number = (999 + vbObjectError)) Then End
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
'On Error GoTo 0
'Err.Raise 999 + vbObjectError, "DSK.VBP", "Fehler in DLL"
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim i As Integer
Dim argc As Integer

ReDim argv(1)
argc = 0

Do
  i = InStr(s, ",")
  If i > 0 Then
    argc = argc + 1
    ReDim Preserve argv(argc)
    argv(argc) = LTrim(RTrim(Left(s, i - 1)))
    s = Mid(s, i + 1)
  Else
    argc = argc + 1
    ReDim Preserve argv(argc)
    argv(argc) = LTrim(RTrim(s))
    s = ""
  End If
Loop While s <> ""

Call DefErrPop
End Sub

Sub SplitString(ByVal s As String, separator As String, argv() As Variant)

'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("SplitString")
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
If (Err.Number = (999 + vbObjectError)) Then End
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
'On Error GoTo 0
'Err.Raise 999 + vbObjectError, "DSK.VBP", "Fehler in DLL"
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim i As Integer
Dim argc As Integer

ReDim argv(1)
argc = 0

Do
  i = InStr(s, separator)
  If i > 0 Then
    ReDim Preserve argv(argc + 1)
    argv(argc) = LTrim(RTrim(Left(s, i - 1)))
    s = Mid(s, i + Len(separator))
  Else
    ReDim Preserve argv(argc + 1)
    argv(argc) = LTrim(RTrim(s))
    s = ""
  End If
  argc = argc + 1
Loop While s <> ""

Call DefErrPop
End Sub

Sub StringLeeren(s As String)

'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("StringLeeren")
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
If (Err.Number = (999 + vbObjectError)) Then End
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
'On Error GoTo 0
'Err.Raise 999 + vbObjectError, "DSK.VBP", "Fehler in DLL"
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
s = Space(Len(s))

Call DefErrPop
End Sub

Sub SortOpenFile(fhSort, SatzLen%, VonSatz&, BisSatz&, SortFields$, SortAnzeige$)

'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("SortOpenFile")
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
If (Err.Number = (999 + vbObjectError)) Then End
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
'On Error GoTo 0
'Err.Raise 999 + vbObjectError, "DSK.VBP", "Fehler in DLL"
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim v%(10), b%(10), l%(10), Absteigend%(10)
Dim x As Integer
Dim FileName As String
Dim CloseFile As Boolean
Dim SortFelder As Integer
Dim SortLen As Integer
Dim SortPos As String
Dim ix As Integer
Dim SortMax As Long
Dim SatzOffset As Long
Dim SortFeld As String
Dim HilfsFeld As String
Dim FileBuffer As String
Dim HilfsBuffer As String
Dim satz As Long
Dim ia As Long
Dim ib As Long
Dim ic As Long
Dim Id As Long
Dim xl As Integer
Dim j As Integer
Dim c As Integer
Dim erg As Integer

'Sortiere offene Datei
'fhSort     FileHandle; bei 0 Datei erst öffnen!
'SatzLen%    Satzlänge der Datei
'VonSatz&    physischer VonSatz
'BisSatz&    physischer BisSatz
'SortFields$ wie in Sort-all  vonbyte,länge[,vonbyte,länge]
'            -vonbyte heißt absteigend sortieren
'            bei zu öffnenden Dateien beginnend mit "Dateiname:....."

If fhSort = 0 Then
  x% = InStr(SortFields$, ":")
  FileName$ = Left$(SortFields$, x% - 1)
  SortFields$ = Mid$(SortFields$, x% + 1)
  CloseFile = True
  fhSort = -fhSort
  Open FileName$ For Binary Access Read Write Shared As #fhSort
End If

SortFelder = 0
SortLen = 0
SortPos$ = SortFields$

While SortPos$ <> ""
  If SortFelder < 10 Then
    SortFelder = SortFelder + 1
    ix% = InStr(SortPos$, ",")
    v%(SortFelder) = Val(Left$(SortPos$, ix% - 1))
    If v%(SortFelder) < 0 Then
      Absteigend%(SortFelder) = -1
      v%(SortFelder) = -v%(SortFelder)
    End If
    SortPos$ = Mid$(SortPos$, ix% + 1)
    ix% = InStr(SortPos$, ",")
    If ix% = 0 Then ix% = Len(SortPos$) + 1
    l%(SortFelder) = Val(Left$(SortPos$, ix% - 1))
    SortPos$ = Mid$(SortPos$, ix% + 1)
    SortLen = SortLen + l%(SortFelder)
  Else
    erg = iMsgBox("SortOpenFile: nur 10 Felder möglich!", vbCritical)
    GoTo EndSortOpenFile
  End If
Wend

SortMax = BisSatz - VonSatz + 1
SatzOffset = VonSatz - 2

HilfsFeld$ = String$(SortLen, 0)
SortFeld$ = String$(SortLen, 0)

FileBuffer = String$(SatzLen%, 0)
HilfsBuffer = String$(SatzLen%, 0)

If SortMax > 1 Then
  If SortMax = 2 Then
    satz = 1: GoSub GetSatz
    Swap SortFeld$, HilfsFeld$
    Swap FileBuffer, HilfsBuffer
    satz = 2: GoSub GetSatz
    If HilfsFeld$ > SortFeld$ Then
      'SWAP feld$(1), feld$(2)
      Swap SortFeld$, HilfsFeld$
      Swap FileBuffer, HilfsBuffer
      GoSub PutSatz
      Swap SortFeld$, HilfsFeld$
      Swap FileBuffer, HilfsBuffer
      satz = 1: GoSub PutSatz
    End If
  Else
    ib = SortMax - 1
    While ib > 1
      ib = Int(ib * 0.3 + 0.5)
      For ic = 1 To ib
        For Id = ic + ib To SortMax Step ib
          'If SortZ > 0 Then
          '  Call dTimer(t#)
          '  If (t# - LastTimer#) > 0.5 Then
          '    LOCATE SortZ, SortS
          '    z% = Int(t# - StartTimer#)
          '    Print USING; "###:"; (z% \ 60);
          '    Print Right$("00" + Mid$(Str$(z% Mod 60), 2), 2);
          '    LastTimer# = t#
          '  End If
          'End If
          
          'MID$(HilfsFeld$, 1, l%) = feld$(id)
          satz = Id: GoSub GetSatz
          HilfsFeld = SortFeld
          HilfsBuffer = FileBuffer

          For ia = Id - ib To 1 Step -ib

            'IF HilfsFeld$ >= feld$(ia) THEN GOTO JumpOver2
            satz = ia: GoSub GetSatz
            If HilfsFeld$ >= SortFeld$ Then GoTo JumpOver2

            'MID$(feld$(ia + ib), 1, l%) = feld$(ia)
            satz = ia + ib: GoSub PutSatz

          Next ia
JumpOver2:
          'SWAP feld$(ia + ib), HilfsFeld$
          satz = ia + ib: GoSub GetSatz
          Swap SortFeld$, HilfsFeld$
          Swap FileBuffer, HilfsBuffer
          GoSub PutSatz
        Next Id
      Next ic
    Wend
  End If
End If
GoTo EndSortOpenFile

PutSatz:
ActFileName = "SORT-Datei"
Put #fhSort, (satz + SatzOffset) * CDbl(SatzLen%) + 1, FileBuffer
ActFileName = ""
Return

GetSatz:
ActFileName = "Sort-Datei"
Get #fhSort, (satz + SatzOffset) * CDbl(SatzLen%) + 1, FileBuffer
ActFileName = ""
xl% = 1
For j% = 1 To SortFelder
  If Absteigend%(j%) Then
    For c% = 1 To l%(j%)
      Mid$(SortFeld$, xl% + c% - 1, 1) = Chr$(255 - Asc(Mid$(FileBuffer, v%(j%) + c% - 1, 1)))
    Next c%
  Else
    Mid$(SortFeld$, xl%, l%(j%)) = Mid$(FileBuffer, v%(j%), l%(j%))
  End If
  xl% = xl% + l%(j%)
Next j%
Return

EndSortOpenFile:

If CloseFile Then Close #fhSort
Call DefErrPop
End Sub

Sub Swap(s1 As String, s2 As String)

'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("Swap")
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
If (Err.Number = (999 + vbObjectError)) Then End
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
'On Error GoTo 0
'Err.Raise 999 + vbObjectError, "DSK.VBP", "Fehler in DLL"
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim sx As String

sx = s1
s1 = s2
s2 = sx

Call DefErrPop
End Sub

Function FirstLettersUcase(ByVal s As String) As String
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("FirstLettersUcase")
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
If (Err.Number = (999 + vbObjectError)) Then End
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
'On Error GoTo 0
'Err.Raise 999 + vbObjectError, "DSK.VBP", "Fehler in DLL"
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

Call DefErrPop
End Function

Sub SetSound(ByVal freq As Long, ByVal Duration As Single)

'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("SetSound")
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
If (Err.Number = (999 + vbObjectError)) Then End
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
'On Error GoTo 0
'Err.Raise 999 + vbObjectError, "DSK.VBP", "Fehler in DLL"
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
#If Win16 Then
  
  Dim s As Integer
  
  freq = freq * 2 ^ 16
  s = SetVoiceSound(1, freq, Int(Duration * 40))
  s = StartSound()
  While WaitSoundState(1) <> 0: Wend

#End If

Call DefErrPop
End Sub

Sub Sound(ByVal freq As Integer, ByVal Duration As Single)

'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("Sound")
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
If (Err.Number = (999 + vbObjectError)) Then End
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
'On Error GoTo 0
'Err.Raise 999 + vbObjectError, "DSK.VBP", "Fehler in DLL"
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsfabs.Sound(freq, Duration)

Call DefErrPop
End Sub

Function KommaPunkt(ByVal txt As String) As String

'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("KommaPunkt")
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
If (Err.Number = (999 + vbObjectError)) Then End
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
'On Error GoTo 0
'Err.Raise 999 + vbObjectError, "DSK.VBP", "Fehler in DLL"
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim x As Integer

x = InStr(txt, ",")
If x > 0 Then
  KommaPunkt = Left(txt, x - 1) + "." + Mid(txt, x + 1)
Else
  KommaPunkt = txt
End If

Call DefErrPop
End Function

Function uInt(ByVal i As Integer) As Long

uInt = CLng(i) - (i < 0) * 65536

End Function

Function xVal(ByVal x As String) As Double

'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("xVal")
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
If (Err.Number = (999 + vbObjectError)) Then End
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
'On Error GoTo 0
'Err.Raise 999 + vbObjectError, "DSK.VBP", "Fehler in DLL"
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim i As Integer

i = InStr(x, ".")
If (i > 0) Then
    If (i = Len(x) - 3) Then
        x = Left$(x, i - 1) + Mid$(x, i + 1)
    End If
End If


i = InStr(x, ",")
If i > 0 Then Mid(x, i, 1) = "."
xVal = Val(x)
Call DefErrPop
End Function

Function ASC7Bit2Ansi(x As String) As String

'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("ASC7Bit2Ansi")
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
If (Err.Number = (999 + vbObjectError)) Then End
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
'On Error GoTo 0
'Err.Raise 999 + vbObjectError, "DSK.VBP", "Fehler in DLL"
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim i As Integer
Dim p As Integer
Dim s As String

s = x
For i = 1 To Len(s)

  p = InStr("[\]{|}~", Mid(s, i, 1))
  If p > 0 Then
    Mid(s, i, 1) = Mid("ÄÖÜäöüß", p, 1)
  End If
Next i
ASC7Bit2Ansi = s

Call DefErrPop
End Function

Sub iiLock(file As Integer, von As Long, bis As Long)

'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("iiLock")
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
If (Err.Number = (999 + vbObjectError)) Then End
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
'On Error GoTo 0
'Err.Raise 999 + vbObjectError, "DSK.VBP", "Fehler in DLL"
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim LockTime As Date
Dim s As String

On Error GoTo StandardError

ActFileHandle = file
Lock #file, von To bis
'Unlock #file, von To bis
ActFileHandle = 0

Call DefErrPop: Exit Sub

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
Call DefErrPop
End Sub

Sub Sleep(ByVal dauer As Integer)

'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("Sleep")
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
If (Err.Number = (999 + vbObjectError)) Then End
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
'On Error GoTo 0
'Err.Raise 999 + vbObjectError, "DSK.VBP", "Fehler in DLL"
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim zeit As Date

zeit = DateAdd("s", dauer, Now)
Do
  'Doevents ist notwendig, da sonst alle anderen
  'Anwendungen nichts arbeiten können
  OpenForms = DoEvents
Loop While zeit > Now

Call DefErrPop
End Sub

Sub SkipControl(Richtung As String, frmKasse As Form)

'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("SkipControl")
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
If (Err.Number = (999 + vbObjectError)) Then End
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
'On Error GoTo 0
'Err.Raise 999 + vbObjectError, "DSK.VBP", "Fehler in DLL"
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim i As Integer
Dim ActTab As Integer
Dim t As Integer
Dim after As Integer
Dim before As Integer
Dim min As Integer
Dim max As Integer
Dim obj(4) As Integer
Dim ok As Boolean

ActTab = frmKasse.ActiveControl.TabIndex
min = 9999
max = 0

For i = 0 To frmKasse.Controls.Count - 1
  'Debug.Print frmKasse.Controls(i).Name
  Err = 0
  On Error Resume Next
  
  
  If frmKasse.Controls(i).TabIndex > 0 Then
    If Err = 0 Then
      ok = True
      If ok Then If TypeOf frmKasse.Controls(i) Is Label Then ok = 0
      If ok Then If TypeOf frmKasse.Controls(i) Is PictureBox Then ok = 0
      If ok Then If TypeOf frmKasse.Controls(i) Is Frame Then ok = 0
      If ok Then
        If frmKasse.Controls(i).Visible = True Then
          If frmKasse.Controls(i).Enabled = True Then
            t = frmKasse.Controls(i).TabIndex
            If t < min Then min = t: obj(1) = i
            If t < ActTab And t > before Then before = t: obj(2) = i
            If t > ActTab And (t < after Or after = 0) Then after = t: obj(3) = i
            If t > max Then max = t: obj(4) = i
          End If
        End If
      End If
    End If
  End If
Next i

On Error GoTo 0
If UCase(Richtung) = "NEXT" Then
  If after > 0 Then
    t = obj(3)
  Else
    t = obj(1)
  End If
Else
  If before > 0 Then
    t = obj(2)
  Else
    t = obj(4)
  End If
End If

'frmkasse.controls(t).name

frmKasse.Controls(t).SetFocus

Call DefErrPop
End Sub

Function FlushIniFile(s As String) As Integer

'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("FlushIniFile")
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
If (Err.Number = (999 + vbObjectError)) Then End
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
'On Error GoTo 0
'Err.Raise 999 + vbObjectError, "DSK.VBP", "Fehler in DLL"
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
FlushIniFile = WritePrivateProfileString("", "", "", s)

Call DefErrPop
End Function

Function TaxeMatchcode(ByVal text As String) As String

'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("TaxeMatchcode")
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
If (Err.Number = (999 + vbObjectError)) Then End
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
'On Error GoTo 0
'Err.Raise 999 + vbObjectError, "DSK.VBP", "Fehler in DLL"
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim s As String
Dim i As Integer

text = MyUCASE(text)
s = ""
For i = 1 To Len(text)
  If InStr("ABCDEFGHIJKLMNOPQRSTUVWXYZÄÖÜ0123456789", Mid(text, i, 1)) > 0 Then
    s = s + Mid(text, i, 1)
  End If
Next i
TaxeMatchcode = Left(s + Space(10), 10)

Call DefErrPop
End Function

'Function FileOpen%(fName$, fAttr$, Optional modus$ = "B", Optional SatzLen% = 100)
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Call DefErrFnc("FileOpen%")
'Call DefErrMod(DefErrModul)
'On Error GoTo DefErr
'GoTo DefErrEnd
'DefErr:
'Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
'Case vbRetry
'  Resume
'Case vbIgnore
'  Resume Next
'End Select
'End
'DefErrEnd:
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Dim Handle%
'
'On Error Resume Next
'FileOpen% = False
'Handle% = FreeFile
'
'
'If (fAttr$ = "R") Then
'    If (modus$ = "B") Then
'        Open fName$ For Binary Access Read Shared As #Handle%
'    Else
'        Open fName$ For Random Access Read Shared As #Handle% Len = SatzLen%
'    End If
'    If (Err = 0) Then
'        If (LOF(Handle%) = 0) Then
'            Close #Handle%
'            Kill (fName$)
'            Err.Raise 53
'        Else
'            Call iiLock(Handle%, 1, 1)
'            Unlock Handle%, 1 To 1
'        End If
'    End If
'ElseIf (fAttr$ = "W") Then
'    If (modus$ = "B") Then
'        Open fName$ For Binary Access Write As #Handle%
'    Else
'        Open fName$ For Random Access Write As #Handle% Len = SatzLen%
'    End If
'ElseIf (fAttr$ = "RW") Then
'    If (modus$ = "B") Then
'        Open fName$ For Binary Access Read Write Shared As #Handle%
'    Else
'        Open fName$ For Random Access Read Write Shared As #Handle% Len = SatzLen%
'    End If
'    Call iiLock(Handle%, 1, 1)
'    Unlock Handle%, 1 To 1
'ElseIf (fAttr$ = "I") Then
'    Open fName$ For Input Access Read Shared As #Handle%
'ElseIf (fAttr$ = "O") Then
'    Open fName$ For Output Access Write Shared As #Handle%
'End If
'
'If (Err = 0) Then
'    FileOpen% = Handle%
'Else
'    Call iMsgBox("Fehler" + Str$(Err) + " beim Öffnen von " + fName$ + vbCr + Err.Description, vbCritical, "FileOpen")
'    Call Programmende
'End If
'
'Call DefErrPop
'End Function

Public Function aeTrim$(s$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("aeTrim$")
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
Dim i%
Dim ret$

ret$ = s$
For i% = 1 To Len(ret$)
    If (Asc(Mid$(ret$, i%, 1)) = 0) Then
        Mid$(ret$, i%, 1) = " "
    Else
        Exit For
    End If
Next i%
aeTrim$ = Trim(ret$)

Call DefErrPop
End Function


