Attribute VB_Name = "modSecurPhX"
Option Explicit

Public SecPZN As Long
Public SecVerfall As String
Public SecCharge As String
Public SecSerNo As String
Public SecScanStr As String
Public SecMessage As String
Public SecDMCVerify As Boolean
Public SecHttp As String
Public SecCod As String
Public SecState As String
Public SecHid As String
Public SecHidText As String

Public SecHochladedatum$, SecVeribeginn_Pflicht$

Public InCheckSecPh As Boolean
Public GHAbbruch As Boolean

Private Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Private Const DefErrModul = "SECURPHX.BAS"

Public Function SecurPharmAufruf(wastun As Integer, pzn As Long, SN As String, Verfall As String, Optional sWasTun As String = "") As Boolean
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("SecurPharmAufruf")
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
'Dim sHochladedatum$, sVeribeginn_Pflicht$
Dim iVerfallMonate As Integer
Dim bAbgabefähig As Boolean
Dim TaxeRecSP As New ADODB.Recordset
Dim ArtikelRecSP As New ADODB.Recordset
Dim dtDV As Date

bAbgabefähig = True

SN = Trim(SN)

'If (Trim(SN) = "") Then
'    SecurPharmAufruf = bAbgabefähig
'    Call DefErrPop: Exit Function
'End If

SecHochladedatum = ""
SecVeribeginn_Pflicht = ""

SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + CStr(pzn)
'Set TaxeRecSP = TaxeDB.OpenRecordset(SQLStr$)
On Error Resume Next
TaxeRecSP.Close
Err.Clear
On Error GoTo DefErr
TaxeRecSP.open SQLStr, taxeAdoDB.ActiveConn
If (TaxeRecSP.EOF = False) Then
'    SecHochladedatum = ""
    On Error Resume Next
    SecHochladedatum = Trim(CheckNullStr(TaxeRecSP!Hochladedatum))
    Err.Clear
    On Error GoTo DefErr
    
    If (SecHochladedatum <> "") Then
        SecVeribeginn_Pflicht = Trim(CheckNullStr(TaxeRecSP!Veribeginn_Pflicht))
        If (Len(SecVeribeginn_Pflicht) = 6) Then
'            sVeribeginn_Pflicht = Mid(sVeribeginn_Pflicht, 3) + Left(sVeribeginn_Pflicht, 2) + "01"
            
            dtDV = CDate("01." + Left(SecVeribeginn_Pflicht, 2) + "." + Mid(SecVeribeginn_Pflicht, 3))
            dtDV = DateAdd("m", 1, dtDV)
            dtDV = DateAdd("d", -1, dtDV)
            SecVeribeginn_Pflicht = Format(dtDV, "YYYYMMDD")
        Else
            SecVeribeginn_Pflicht = ""
        End If
        
        iVerfallMonate = CheckNullLong(TaxeRecSP!VerfallMonate)
        
        If (Format(Now, "YYYYMMDD") >= SecHochladedatum) Then
'            If (SecurPharmValue(0) = "") Then   'kein DMC eingelesen
            If (SN = "") Then   'kein DMC eingelesen
                If (SecVeribeginn_Pflicht <> "") Then
'                    If (Format(Now, "YYYYMMDD") >= SecVeribeginn_Pflicht) Then
                        bAbgabefähig = False
'                    End If
                End If
            Else
                If (SecVeribeginn_Pflicht = "") Then
                    If (iVerfallMonate > 0) Then
                        dtDV = CDate("09.02.2019")
                        dtDV = DateAdd("m", iVerfallMonate, dtDV)
                        SecVeribeginn_Pflicht = Format(dtDV, "YYYYMMDD")
                    Else
                        SecVeribeginn_Pflicht = "20991231"
                    End If
                End If
                If (Verfall = "") Then
                    SQLStr = "SELECT * FROM QR_SecurPharm WHERE "
                    SQLStr = SQLStr + "(Pzn=" + CStr(pzn) + ") AND (SerienNr='" + SN + "')"
                    FabsErrf = Artikel.OpenRecordset(ArtikelRecSP, SQLStr, 0)
                    If (FabsErrf = 0) Then
                        Verfall = Format(CheckNullDate(ArtikelRecSP!Verfall), "YYMMDD")
                    End If
                End If
'                If (sVeribeginn_Pflicht < "20" + Verfall) Then
'                    Dim bErg As Boolean
'                    bAbgabefähig = SecurPharmAufruf2(wastun, pzn, SN)
'                End If
                Dim bErg As Boolean
                bAbgabefähig = SecurPharmAufruf2(wastun, pzn, SN, (SecVeribeginn_Pflicht > "20" + Verfall), sWasTun)
            End If
        End If
    End If
End If
SecurPharmAufruf = bAbgabefähig

Call DefErrPop
End Function
        
Public Function SecurPharmAufruf2(wastun As Integer, pzn As Long, SN As String, DVgreaterEXP As Boolean, Optional sWasTun As String = "") As Boolean
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("SecurPharmAufruf2")
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
'Dim pzn As Long
Dim S As String, SQL As String, zeile As String, CheckGHFile As String, sAction$
Dim i As Integer
Dim erfolg As Boolean
ReDim argv(0) As Variant
Dim fhtmpname As String
Dim StartTimer As Single

'Vor dem Funktionsaufruf extern SecScanstr bzw. bei undo dispense SecPZN und SecSerno belegen!

'wastun:
'1 ... verifizieren
'2 ... dispense
'3 ... undo dispense

SecMessage = "TIMEOUT"
SecHttp = "???"
SecCod = ""
SecState = ""
SecPZN = pzn
SecSerNo = SN
SecHid = ""
SecHidText = ""

'globale Variablen um extern festzustellen, ob die Funktion noch läuft bzw. ob Benutzer abgebrochen hat
InCheckSecPh = False
GHAbbruch = False


fhtmpname = "AE" + Format(Now, "nnss") + CStr(Val(para.User)) + "L" + ".$$$"
On Error Resume Next
If Dir(CurDir + "\" + fhtmpname) > "" Then Kill CurDir + "\" + fhtmpname
CheckGHFile = CurDir + "\" + fhtmpname

If (sWasTun <> "") Then
    sAction = sWasTun
Else
    Select Case wastun
    Case 1
    '  Shell CurDir + "\SECURPH.EXE 0," + fhtmpname + ",CHECK,DMC=" + SecScanStr
      sAction = "CHECK"
    Case 2
    '  Shell CurDir + "\SECURPH.EXE 0," + fhtmpname + ",DISPENSE,DMC=" + SecScanStr
      sAction = "DISPENSE"
    Case 3
    '  Shell CurDir + "\SECURPH.EXE 0," + fhtmpname + ",UNDO_DISPENSE,PZN=" + Format(SecPZN, "00000000") + ",SN=" + Trim(SecSerNo)
      sAction = "UNDO_DISPENSE"
    End Select
End If

Shell CurDir + "\SECURPH.EXE 0," + fhtmpname + "," + sAction + ",PZN=" + Format(pzn, "00000000") + ",SN=" + Trim(SN) + IIf(DVgreaterEXP, ",DVgreaterEXP", "")

InCheckSecPh = True
StartTimer = Timer

'im folenden Block wird eine "bitte warten-Nachricht mit Abbruch-Möglichkeit angezeigt, bei Abbruch wird GhAbbruch=True gesetzt
'Über globale Variable CheckGHFile wird auch in der Funktion Message die Dateilänge ständig überprüft und abgebrochen wenn >0

'Call AktInfoAnzeigen
'frmKasse.txtEingabe.Enabled = False
'Call Message("Securpharm-Abfrage... bitte warten", vbOKOnly Or vbInformation Or vbSystemModal)
'Call Message("", 0)
'frmKasse.txtEingabe.Enabled = True

'diese Prüfung steht auch in der Funktion Message - hier redundant, aber die Funktion Message ist sehr umfangreich,
' die schick ich dir liebe nicht mit
Do
'  OpenForms = DoEvents
  If Dir(CurDir + "\" + fhtmpname) > "" Then
    If GetFileSize(CurDir + "\" + fhtmpname) > 0 Then
      Exit Do
    End If
  End If
  
  If ((Timer - StartTimer) > 15) Then
    GHAbbruch = True
    Exit Do
  End If
'  If GHAbbruch Then Exit Do
Loop
On Error GoTo DefErr
'If Not GHAbbruch Then
'If GHAbbruch Then
'  Call AktInfoAnzeigen
'  Call DefErrPop
'  Exit Function
'End If

InCheckSecPh = False
'Call AktInfoAnzeigen

On Error Resume Next
erfolg = False

If (GHAbbruch) Then
Else
    S = Space(1024)
    i = GetPrivateProfileSection("Securpharm", S, 1024, CurDir + "\" + fhtmpname)
    If i > 0 Then
      S = Left(S, i)
      Call IniSection(S, argv())
      
      For i = 1 To UBound(argv, 2)
        Select Case argv(0, i)
        Case "HTTP"
            SecHttp = argv(1, i)
'            If (wastun = 2) Then
                erfolg = (argv(1, i) = 200)
'            End If
        Case "cod"
            SecCod = argv(1, i)
        Case "reasons"
        Case "state"
            SecState = argv(1, i)
    '      If wastun = 1 Or wastun = 3 Then
    '        If argv(1, i) = "ACTIVE" Then erfolg = True
    '      ElseIf wastun = 2 Then
    '        If argv(1, i) = "INACTIVE" Then erfolg = True
    '      End If
        Case "mes"
          SecMessage = argv(1, i)
        Case "Pzn"
          SecPZN = argv(1, i)
        Case "Charge"
          SecCharge = argv(1, i)
        Case "Verfalldatum"
          SecVerfall = argv(1, i)
        Case "SerienNr"
          SecSerNo = argv(1, i)
        Case "hid"
          SecHid = argv(1, i)
        Case "hid_text"
          SecHidText = argv(1, i)
        End Select
      Next i
    End If
    
    Kill CurDir + "\" + fhtmpname
    'erfolg = True
End If

If (erfolg = False) Then
    S = "Problem bei " + sAction + ":" + vbCrLf + vbCrLf
    S = S + "PZN: " + CStr(SecPZN) + vbCrLf
    S = S + "SN: " + SecSerNo + vbCrLf
    S = S + vbCrLf
    S = S + "HTTP: " + SecHttp + vbCrLf
    S = S + "Code: " + SecCod + vbCrLf
    S = S + "Message: " + SecMessage + vbCrLf
    S = S + "State: " + SecState + vbCrLf
    If (SecHid <> "") Then
'        s = s + "Hid: " + SecHid + vbCrLf
'        s = s + "HidText: " + SecHidText + vbCrLf
        S = S + vbCrLf + "Handlungsanweisung: " + vbCrLf
        S = S + SecHidText + " (" + SecHid + ")" + vbCrLf
    End If
    
    Call MsgBox(S, vbCritical, "SECURPHARM")
    
    If (GHAbbruch) Then
        erfolg = True
    End If
End If

SecurPharmAufruf2 = erfolg

Call DefErrPop
End Function


Sub IniSection(ByVal txt As String, argv() As Variant)

'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("IniSection")
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
Dim j As Integer
Dim argc As Integer
Dim S As String

ReDim argv(1, 0)
argc = 0

Do
  j = InStr(txt, Chr(0))
  If j > 0 Then
    S = Trim(Left(txt, j - 1))
    txt = Mid(txt, j + 1)
  Else
    S = txt
    txt = ""
  End If
  If S > "" Then
    i = InStr(S, "=")
    If i > 0 Then
      argc = argc + 1
      ReDim Preserve argv(1, argc)
      argv(0, argc) = Left(S, i - 1)
      argv(1, argc) = Mid(S, i + 1)
    End If
  End If
Loop While txt <> ""
Call DefErrPop

End Sub

Function GetFileSize(ByVal strFilename As String) As Long
    On Error Resume Next

    GetFileSize = FileLen(strFilename)

    If Err > 0 Then
        GetFileSize = -1
        Err = 0
    End If
End Function
