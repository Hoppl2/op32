VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMsv3DB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim dKurz$
Dim dBezeichnung$
Dim dName$
Dim dZugriff$
Dim dAnzFields%, dAnzFieldsDOS%

Dim fKurz$()
Dim fBezeichnung$()

Dim iLiefNamen$(1)
Dim iAnzLiefNamen%

Dim iOrgLiefNr%
Dim iOrgKurz$

Dim SqlServerAktiv%

Dim iConn As New ADODB.Connection
'Dim iComm As New ADODB.Command

Private Const DefErrModul = "MSV3DB.CLS"

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
Dim i%

dKurz$ = "MSV3"
dBezeichnung$ = "Msv3"
dName$ = "Msv3.mdb"

'Set ManuellAngeboteDB1 = Me

Call clsError.DefErrPop
End Sub

Function DBvorhanden%()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("DBvorhanden%")
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

SqlServerAktiv = sqlop1.IstSqlServerDB(dBezeichnung)
If (SqlServerAktiv) Then
'    ret = (Dir(iNetworkDataPath + "\" + dBezeichnung + ".mdf") <> "")
    ret = sqlop1.SqlCheckDatabase(dBezeichnung)
Else
    ret = (Dir(dName) <> "")
End If

DBvorhanden = ret

Call clsError.DefErrPop
End Function

Function SqlServerDB%()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("SqlServerDB%")
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

SqlServerDB = SqlServerAktiv

Call clsError.DefErrPop
End Function

Function OpenDB%()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("OpenDB%")
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
Dim i%, j%, iAnz%, erg%, ErrNumber%
Dim h$

SqlServerAktiv = sqlop1.IstSqlServerDB(dBezeichnung)
If (SqlServerAktiv) Then
    h = sqlop1.SqlConnectionString(SqlServerAktiv)
    iConn.ConnectionString = Left(h, Len(h) - 1) + dBezeichnung + "; Data Source=" + sqlop1.SqlServer(SqlServerAktiv)
Else
    iConn.ConnectionString = sqlop1.SqlConnectionString(0) + "Data Source=" + dName
End If

iConn.CursorLocation = adUseClient
On Error Resume Next
Err.Clear
iConn.Open
ErrNumber = Err.Number

On Error GoTo DefErr
If (ErrNumber% = 0) Then
'    ManuelleAngeboteDbOk = True
End If

OpenDB = (ErrNumber = 0)

Call clsError.DefErrPop
End Function

Function CloseDB%()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("CloseDB%")
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

iConn.Close
Set iConn = Nothing

Call clsError.DefErrPop
End Function

Function CreateDB%()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("CreateDB%")
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

ret = sqlop1.SqlCreateDatabase(dBezeichnung, dBezeichnung)
If (ret) Then
    ret = OpenDB    '(False)
End If

CreateDB = ret

Call clsError.DefErrPop
End Function

Function ActiveConn() As ADODB.Connection
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("ActiveConn")
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

Set ActiveConn = iConn

Call clsError.DefErrPop
End Function

Function TaxeValue$(FeldName$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("TaxeValue$")
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
Dim ret$, h$, iFeldName$
Dim fldRecordset As Field

ret$ = ""

iFeldName$ = UCase(FeldName$)
    
'Set fldRecordset = WawiDatRec.Fields(FeldName$)
'Select Case iFeldName$
'    Case "PZN"
'        ret$ = Format(fldRecordset.Value, "0000000")
'    Case "EK", "VK", "FESTBETRAG", "KVA", "VK", "VKALT"
'        ret$ = Format(fldRecordset.Value / 100#, "0.00")
'        If (ret$ = "0.00") Or (ret$ = "0,00") Then
'            ret$ = "0"
'        End If
'    Case "ZUZAHLUNG"
'        ret$ = MachTaxeAuswahlZuzahlung$(TaxeRec)
'    Case "GUELTIG", "LPREISAENDERUNG"
'        If (iFeldName$ = "GUELTIG") Then
'            h$ = LTrim$(fldRecordset.Value)
'        Else
'            h$ = Right$(String(6, "0") + Format(fldRecordset.Value, "0"), 6)
'        End If
'        ret$ = Left$(h$, 2) + "." + Mid$(h$, 3, 2) + "."
'        If (Val(Right$(h$, 2)) < 50) Then
'            ret$ = ret$ + "20"
'        Else
'            ret$ = ret$ + "19"
'        End If
'        ret$ = ret$ + Right$(h$, 2)
'    Case Else
'        ret$ = LTrim$(fldRecordset.Value)
'End Select

TaxeValue$ = " " + ret$

'For i% = 0 To (recTaxe.Fields.Count - 1)
'    Set fldRecordset = recTaxe.Fields(i%)
'    Debug.Print fldRecordset.Name; fldRecordset.Size; fldRecordset.Type; fldRecordset.Value
''    Call FieldOutput("Recordsrt", fldRecordset)
'Next i%

Call clsError.DefErrPop
End Function

Function FeldBezeichnung$(iFeld As Variant)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("FeldBezeichnung$")
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
Dim ind%
Dim ret$

ret$ = ""

If (VarType(iFeld) = vbString) Then
'    ind% = SucheFeldInd%(iFeld)
Else
    ind% = iFeld
End If
If (ind% >= 0) Then
    ret$ = fBezeichnung$(ind%)
End If

FeldBezeichnung$ = ret$

Call clsError.DefErrPop
End Function

Function FeldKurz$(iFeld As Variant)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("FeldKurz$")
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
Dim ind%
Dim ret$

ret$ = ""

If (VarType(iFeld) = vbString) Then
'    ind% = SucheFeldInd%(iFeld)
Else
    ind% = iFeld
End If
If (ind% >= 0) Then
    ret$ = fBezeichnung$(ind%)
End If

FeldKurz$ = ret$

Call clsError.DefErrPop
End Function

Function FeldWert$(FeldName$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("FeldWert$")
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
Dim ret$, h$, iFeldName$
Dim fldRecordset As ADODB.Field

ret$ = ""

'iFeldName$ = UCase(FeldName$)
'If (iFeldName$ = "NAME(0)") Then
'    ret$ = Trim(clsOpTool.CheckNullStr(KundenRec!Anrede))
'    If (ret$ <> "") Then
'        ret$ = ret$ + " "
'    End If
'    ret$ = ret$ + Trim(clsOpTool.CheckNullStr(KundenRec!Titel))
'ElseIf (iFeldName$ = "NAME(1)") Then
'    ret$ = Trim(clsOpTool.CheckNullStr(KundenRec!VorName))
'    If (ret$ <> "") Then
'        ret$ = ret$ + " "
'    End If
'    ret$ = ret$ + Trim(clsOpTool.CheckNullStr(KundenRec!Name))
'ElseIf (iFeldName$ = "NAME(2)") Then
'    ret$ = Trim(clsOpTool.CheckNullStr(KundenRec!Strasse))
'    If (ret$ <> "") Then
'        ret$ = ret$ + " "
'    End If
'    ret$ = ret$ + Trim(clsOpTool.CheckNullStr(KundenRec!HausNr))
'ElseIf (iFeldName$ = "NAME(3)") Then
'    ret$ = Trim(clsOpTool.CheckNullStr(KundenRec!PLZ))
'    If (ret$ <> "") Then
'        ret$ = ret$ + " "
'    End If
'    ret$ = ret$ + Trim(clsOpTool.CheckNullStr(KundenRec!Ort))
'Else
    For i = 1 To dAnzFieldsDOS
        If (fKurz(i - 1) = FeldName) Then
            FeldName = LieferantenRec.Fields(i + 1).Name
            Exit For
        End If
    Next i
'    If (FeldName = "LIEFNR") Then
'        FeldName = "LIEFNR2"
'    End If
    
    If (UCase(FeldName) = "LIEFNR") Then
        ret$ = CStr(iOrgLiefNr)
    ElseIf (UCase(FeldName) = "KURZ") Then
        ret$ = iOrgKurz
    Else
        Set fldRecordset = LieferantenRec.Fields(FeldName$)
        If (IsNull(fldRecordset)) Then
            ret$ = ""
        ElseIf (fldRecordset.Type = dbDate) Then
            ret$ = Format(clsOpTool.CheckNullDouble(fldRecordset.Value), "DD.MM.YYYY")
        ElseIf (fldRecordset.Type = dbDouble) Or (fldRecordset.Type = dbFloat) Then
            ret$ = Format(clsOpTool.CheckNullDouble(fldRecordset.Value), "0.00")
        Else
            ret$ = LTrim$(clsOpTool.CheckNullStr(fldRecordset.Value))
        End If
    End If
'End If

FeldWert$ = " " + ret$

'For i% = 0 To (recTaxe.Fields.Count - 1)
'    Set fldRecordset = recTaxe.Fields(i%)
'    Debug.Print fldRecordset.Name; fldRecordset.Size; fldRecordset.Type; fldRecordset.Value
''    Call FieldOutput("Recordsrt", fldRecordset)
'Next i%

Call clsError.DefErrPop
End Function

Public Property Get DateiBezeichnung$()
DateiBezeichnung$ = dBezeichnung$
End Property

Public Property Get DateiKurz$()
DateiKurz$ = dKurz$
End Property

Public Property Get DateiName$()
DateiName$ = dName$
End Property

Public Property Get AnzFelder%()
AnzFelder% = dAnzFields%
End Property

Public Property Get AnzRec&()
Dim ret&
Dim sRec As New ADODB.Recordset

ret& = 0
sRec.Open "SELECT COUNT(*) AS Anz FROM " + "Absagen", iConn
If (sRec.RecordCount <> 0) Then
    ret& = sRec!anz
End If
sRec.Close

AnzRec& = ret&
End Property
