VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLieferantenDB"
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

Dim iLiefNamen$(260)
Dim iAnzLiefNamen%

Dim iOrgLiefNr%
Dim iOrgKurz$

Dim SqlServerAktiv%

Dim iConn As New ADODB.Connection
'Dim iComm As New ADODB.Command

Private Const DefErrModul = "LIEFERANTENDB.CLS"

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

dKurz$ = "LIEF"
dBezeichnung$ = "Lieferanten"
dName$ = "Liferant.mdb"
dZugriff$ = ""

dAnzFieldsDOS% = 21

ReDim fKurz$(dAnzFieldsDOS% - 1)

fKurz$(0) = "NA(0)"
fKurz$(1) = "NA(1)"
fKurz$(2) = "NA(2)"
fKurz$(3) = "NA(3)"
fKurz$(4) = "KURZ"
fKurz$(5) = "WBEZ"
fKurz$(6) = "TEL"
fKurz$(7) = "KTO"
fKurz$(8) = "BLZ"
fKurz$(9) = "VERT"
fKurz$(10) = "KTVOR"
fKurz$(11) = "BEIN"
fKurz$(12) = "DW"
fKurz$(13) = "IDA"
fKurz$(14) = "FIBU"
fKurz$(15) = "IDL"
fKurz$(16) = "DUMMY"
fKurz$(17) = "WNWT"
fKurz$(18) = "DIR"
fKurz$(19) = "SK"
fKurz$(20) = "SALDO"

Set LieferantenDB1 = Me

Call clsError.DefErrPop
End Sub

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
    LieferantenDBok = True
    
'    LieferantenComm.ActiveConnection = LieferantenConn
    
    LieferantenRec.Open "SELECT * FROM Lieferanten", iConn
    dAnzFields% = LieferantenRec.Fields.Count '+ 1
    ReDim fBezeichnung$(dAnzFields% + 1)
'    fBezeichnung$(0) = "ImportKz"
    For i% = 0 To (dAnzFields% - 1)
        fBezeichnung$(i%) = LieferantenRec.Fields(i%).Name
    Next i%
    LieferantenRec.Close

    LieferantenRec.Open "SELECT * FROM LieferantenZusatz", iConn
    iAnz% = LieferantenRec.Fields.Count '+ 1
    ReDim Preserve fBezeichnung$(dAnzFields% + iAnz + 1)
    For i% = 1 To (iAnz% - 1)
        fBezeichnung$(dAnzFields) = LieferantenRec.Fields(i%).Name
        dAnzFields = dAnzFields + 1
    Next i%
    LieferantenRec.Close
    
'    Set sConn = LieferantenConn
    
'    Exit For
'ElseIf (j% = 1) Then
'    Call Reorg
End If

OpenDB = (ErrNumber = 0)

'''''''''''''''''''''''''''''''''''''''''''''
'For j% = 1 To 2
'    On Error Resume Next
'    Err.Clear
'
''    LieferantenConn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + DBName
'    'LieferantenConn.ConnectionString = "Provider=SQLOLEDB; Integrated Security=SSPI; Initial Catalog=" + "Lieferanten" + "; Data Source=" + "NBNEU\SQLEXPRESS"
'    LieferantenConn.ConnectionString = sConn.ConnectionString
'    LieferantenConn.CursorLocation = adUseClient
'    On Error Resume Next
'    Err.Clear
'    LieferantenConn.Open
'    ErrNumber = Err.Number
'
'    On Error GoTo DefErr
'    If (ErrNumber% = 0) Then
'        LieferantenDBok = True
'
''        Set LieferantenComm = New ADODB.Command
'        LieferantenComm.ActiveConnection = LieferantenConn
'
''        dAnzFields = 0
''        ' Use OpenSchema and get the table names.
''        ' The final argument in the parameter array
''        ' gives the table's name.
''        Set LieferantenRec = LieferantenConn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "LIEFERANTEN"))
''        Do While Not LieferantenRec.EOF
''            ReDim Preserve fBezeichnung$(dAnzFields% + 1)
'''            fBezeichnung$(dAnzFields%) = LieferantenRec!COLUMN_NAME
''            dAnzFields = dAnzFields + 1
''            LieferantenRec.MoveNext
''        Loop
''        LieferantenRec.Close
''
''        Set LieferantenRec = LieferantenConn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "LIEFERANTENZUSATZ"))
''        Do While Not LieferantenRec.EOF
''            ReDim Preserve fBezeichnung$(dAnzFields% + 1)
''            fBezeichnung$(dAnzFields%) = LieferantenRec!COLUMN_NAME
''            dAnzFields = dAnzFields + 1
''            LieferantenRec.MoveNext
''        Loop
''        LieferantenRec.Close
'
'        LieferantenRec.Open "SELECT * FROM Lieferanten", LieferantenConn
'        dAnzFields% = LieferantenRec.Fields.Count '+ 1
'        ReDim fBezeichnung$(dAnzFields% + 1)
'    '    fBezeichnung$(0) = "ImportKz"
'        For i% = 0 To (dAnzFields% - 1)
'            fBezeichnung$(i%) = LieferantenRec.Fields(i%).Name
'        Next i%
'        LieferantenRec.Close
'
'        LieferantenRec.Open "SELECT * FROM LieferantenZusatz", LieferantenConn
'        iAnz% = LieferantenRec.Fields.Count '+ 1
'        ReDim Preserve fBezeichnung$(dAnzFields% + iAnz + 1)
'        For i% = 1 To (iAnz% - 1)
'            fBezeichnung$(dAnzFields) = LieferantenRec.Fields(i%).Name
'            dAnzFields = dAnzFields + 1
'        Next i%
'        LieferantenRec.Close
'
'        Set sConn = LieferantenConn
'
'        Exit For
'    ElseIf (j% = 1) Then
'        Call Reorg
'    End If
'Next j%
'
'OpenDatenbank% = ErrNumber%

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

'Function OpenDatenbank%(DBname$, sConn As ADODB.Connection)
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Call clsError.DefErrFnc("OpenDatenbank%")
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
'Dim i%, j%, iAnz%, erg%, ErrNumber%
'
'If (DBname$ = "") Then
'    DBname$ = dName$
'End If
'
'For j% = 1 To 2
'    On Error Resume Next
'    Err.Clear
'
''    LieferantenConn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + DBName
'    'LieferantenConn.ConnectionString = "Provider=SQLOLEDB; Integrated Security=SSPI; Initial Catalog=" + "Lieferanten" + "; Data Source=" + "NBNEU\SQLEXPRESS"
'    LieferantenConn.ConnectionString = sConn.ConnectionString
'    LieferantenConn.CursorLocation = adUseClient
'    On Error Resume Next
'    Err.Clear
'    LieferantenConn.Open
'    ErrNumber = Err.Number
'
'    On Error GoTo DefErr
'    If (ErrNumber% = 0) Then
'        LieferantenDBok = True
'
''        Set LieferantenComm = New ADODB.Command
'        LieferantenComm.ActiveConnection = LieferantenConn
'
''        dAnzFields = 0
''        ' Use OpenSchema and get the table names.
''        ' The final argument in the parameter array
''        ' gives the table's name.
''        Set LieferantenRec = LieferantenConn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "LIEFERANTEN"))
''        Do While Not LieferantenRec.EOF
''            ReDim Preserve fBezeichnung$(dAnzFields% + 1)
'''            fBezeichnung$(dAnzFields%) = LieferantenRec!COLUMN_NAME
''            dAnzFields = dAnzFields + 1
''            LieferantenRec.MoveNext
''        Loop
''        LieferantenRec.Close
''
''        Set LieferantenRec = LieferantenConn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "LIEFERANTENZUSATZ"))
''        Do While Not LieferantenRec.EOF
''            ReDim Preserve fBezeichnung$(dAnzFields% + 1)
''            fBezeichnung$(dAnzFields%) = LieferantenRec!COLUMN_NAME
''            dAnzFields = dAnzFields + 1
''            LieferantenRec.MoveNext
''        Loop
''        LieferantenRec.Close
'
'        LieferantenRec.Open "SELECT * FROM Lieferanten", LieferantenConn
'        dAnzFields% = LieferantenRec.Fields.Count '+ 1
'        ReDim fBezeichnung$(dAnzFields% + 1)
'    '    fBezeichnung$(0) = "ImportKz"
'        For i% = 0 To (dAnzFields% - 1)
'            fBezeichnung$(i%) = LieferantenRec.Fields(i%).Name
'        Next i%
'        LieferantenRec.Close
'
'        LieferantenRec.Open "SELECT * FROM LieferantenZusatz", LieferantenConn
'        iAnz% = LieferantenRec.Fields.Count '+ 1
'        ReDim Preserve fBezeichnung$(dAnzFields% + iAnz + 1)
'        For i% = 1 To (iAnz% - 1)
'            fBezeichnung$(dAnzFields) = LieferantenRec.Fields(i%).Name
'            dAnzFields = dAnzFields + 1
'        Next i%
'        LieferantenRec.Close
'
'        Set sConn = LieferantenConn
'
'        Exit For
'    ElseIf (j% = 1) Then
'        Call Reorg
'    End If
'Next j%
'
'OpenDatenbank% = ErrNumber%
'
'Call clsError.DefErrPop
'End Function

''Function OpenDatenbank%(DBName$, Optionen%, Schreibgeschuetzt%, DB As Database)
'''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
''Call clsError.DefErrFnc("OpenDatenbank%")
''Call clsError.DefErrMod(DefErrModul)
''On Error GoTo DefErr
''GoTo DefErrEnd
''DefErr:
''Select Case clsError.DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
''Case vbRetry
''  Resume
''Case vbIgnore
''  Resume Next
''End Select
''Call ProjektForm.EndeDll
''DefErrEnd:
'''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
''Dim i%, j%, iAnz%, erg%, ErrNumber%
''
''If (DBName$ = "") Then
''    DBName$ = dName$
''End If
''
''For j% = 1 To 2
''    On Error Resume Next
''    Err.Clear
''    Set LieferantenDB = OpenDatabase(DBName$, Optionen%, Schreibgeschuetzt%)
''    ErrNumber% = Err.Number
''    On Error GoTo DefErr
''    If (ErrNumber% = 0) Then
''        dAnzFields% = LieferantenDB.TableDefs("LIEFERANTEN").Fields.Count '+ 1
''        ReDim fBezeichnung$(dAnzFields% + 1)
''    '    fBezeichnung$(0) = "ImportKz"
''        For i% = 0 To (dAnzFields% - 1)
''            fBezeichnung$(i%) = LieferantenDB.TableDefs("LIEFERANTEN").Fields(i%).Name
''        Next i%
''
''        iAnz% = LieferantenDB.TableDefs("LIEFERANTENZUSATZ").Fields.Count '+ 1
''        ReDim Preserve fBezeichnung$(dAnzFields% + iAnz + 1)
''        For i% = 1 To (iAnz% - 1)
''            fBezeichnung$(dAnzFields) = LieferantenDB.TableDefs("LIEFERANTENZUSATZ").Fields(i%).Name
''            dAnzFields = dAnzFields + 1
''        Next i%
''
''        Set DB = LieferantenDB
''
''        Exit For
''    ElseIf (j% = 1) Then
''        Call Reorg
''    End If
''Next j%
''
''OpenDatenbank% = ErrNumber%
''
''Call clsError.DefErrPop
''End Function

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
sRec.Open "SELECT COUNT(*) AS Anz FROM " + "Lieferanten", iConn
If (sRec.RecordCount <> 0) Then
    ret& = sRec!anz
End If
sRec.Close

AnzRec& = ret&
End Property

Function Reorg%()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("Reorg%")
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

ret% = True

'MerkzettelDB.Close
ret% = ReorgLieferanten%

Reorg% = ret%

Call clsError.DefErrPop
End Function

Function ReorgLieferanten%()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("ReorgLieferanten%")
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
Dim i&, AnzRec&, ErrorNumber&
Dim s$, OrgMdb$, CompactMdb$

AnzRec& = 100

Load frmReorg
With frmReorg
    .Caption = "Lieferanten-Daten reorganisieren"
    .lbReorg.Caption = "Komprimieren der Lieferanten-Daten ..."
    .MaxAnz (AnzRec&)
    .Show
End With

OrgMdb$ = CurDir$ + "\" + dName$
CompactMdb$ = CurDir$ + "\Lifer2.mdb"

If (Dir(CompactMdb$) <> "") Then
    Kill CompactMdb$
End If

'Compact Database to New Name
On Error Resume Next
DBEngine.CompactDatabase OrgMdb$, CompactMdb$
ErrorNumber = Err.Number
Err.Clear
On Error GoTo DefErr

i& = AnzRec&
frmReorg.ShowFortschritt (i&)
Unload frmReorg

If (ErrorNumber <> 0) Then
    s$ = "Die Datei ist noch geöffnet - Reorganisieren nicht möglich!"
    Call clsDialog.MessageBox(s$, vbInformation)
Else
    FileCopy OrgMdb$, CurDir$ + "\Liferant.mdb"
    FileCopy CompactMdb$, OrgMdb$
End If

ReorgLieferanten% = True

Call clsError.DefErrPop
End Function

Function ErsterFreier%()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("ErsterFreier%")
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
Dim i%, frei%, ret%
Dim h$, SQLStr$

On Error Resume Next
LieferantenRec.Close
Err.Clear
On Error GoTo DefErr

ret% = 1
SQLStr$ = "SELECT * FROM Lieferanten ORDER BY LiefNr"
LieferantenRec.Open SQLStr$, iConn
Do
    If (LieferantenRec.EOF) Then
        Exit Do
    End If
    
    If (LieferantenRec!LiefNr <> ret) Then
        Exit Do
    End If
    
    ret = ret + 1
    LieferantenRec.MoveNext
Loop
LieferantenRec.Close

ErsterFreier% = ret%

Call clsError.DefErrPop
End Function

Function IstFrei%(pos%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("IstFrei%")
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
Dim i%, frei%, ret%
Dim h$, SQLStr$

ret% = True

On Error Resume Next
LieferantenRec.Close
Err.Clear
On Error GoTo DefErr

SQLStr$ = "SELECT * FROM Lieferanten WHERE LiefNr =" + Str$(pos)
LieferantenRec.Open SQLStr$, iConn
If Not (LieferantenRec.EOF) Then
    h$ = Trim(clsOpTool.CheckNullStr(LieferantenRec!kurz))
    h$ = UCase(Trim$(h$))
    If (h$ <> "") Then
        If (Asc(Left$(h$, 1)) >= 32) Then
            ret% = 0
        End If
    End If
    
    If (ret%) Then
        h$ = Trim(clsOpTool.CheckNullStr(LieferantenRec!Name1))
        h$ = UCase(Trim$(h$))
        If (h$ <> "") Then
            If (Asc(Left$(h$, 1)) >= 32) Then
                ret% = 0
            End If
        End If
    End If
End If
LieferantenRec.Close

IstFrei% = ret%

Call clsError.DefErrPop
End Function

Sub DeleteRecord(Optional recno& = 0, Optional sSqlServer$ = "")
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("DeleteRecord")
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
Dim lRecs&
Dim sTabelle$(4), SQLStr$

sTabelle(0) = "Lieferanten"
sTabelle(1) = "LieferantenZusatz"
sTabelle(2) = "RabattTabelle"
sTabelle(3) = "BmEkTabelle"
sTabelle(4) = "Ausnahmen"
For i = 0 To 4
    If (sSqlServer <> "") Then
        SQLStr$ = "DELETE " + sTabelle(i) + " WHERE LiefNr =" + Str$(recno)
    Else
        SQLStr$ = "DELETE * FROM " + sTabelle(i) + " WHERE LiefNr =" + Str$(recno)
    End If
'    LieferantenComm.CommandText = SQLStr$
'    LieferantenComm.CommandTimeout = 120
'    LieferantenComm.Execute
    iConn.CommandTimeout = 120
    Call iConn.Execute(SQLStr, lRecs, adExecuteNoRecords)
Next i

Call clsError.DefErrPop
End Sub

Public Property Get LiefName$(ByVal ind%)
LiefName$ = iLiefNamen$(ind%)
End Property
Public Property Let LiefName(ByVal ind%, ByVal vNewValue$)
iLiefNamen$(ind%) = vNewValue$
End Property

Public Property Get AnzLiefNamen%()
AnzLiefNamen% = iAnzLiefNamen%
End Property
Public Property Let AnzLiefNamen(ByVal vNewValue%)
iAnzLiefNamen% = vNewValue%
End Property

Public Property Get OrgLiefNr%()
OrgLiefNr% = iOrgLiefNr
End Property
Public Property Let OrgLiefNr(ByVal vNewValue%)
iOrgLiefNr% = vNewValue%
End Property

Public Property Get OrgKurz$()
OrgKurz = iOrgKurz
End Property
Public Property Let OrgKurz(ByVal vNewValue$)
iOrgKurz = vNewValue$
End Property

Sub Konvert2Dos(ind%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("Konvert2Dos")
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
Dim i%, j%, k%, ind2%, iLief%, iDatenVonLief%, val2%
Dim h$, h2$, SQLStr$
Dim LieferantenRec2 As New ADODB.Recordset

iLief = ind
SQLStr$ = "SELECT * FROM Lieferanten WHERE LiefNr =" + Str$(ind)
LieferantenRec2.Open SQLStr, iConn
If (LieferantenRec2.RecordCount <> 0) Then
    iDatenVonLief = clsOpTool.CheckNullLong(LieferantenRec2!DatenVonLief)
    If (iDatenVonLief > 0) Then
        iLief = iDatenVonLief
    End If
End If
LieferantenRec2.Close

SQLStr$ = "SELECT * FROM LieferantenZusatz WHERE LiefNr =" + Str$(iLief)
LieferantenzusatzRec.Open SQLStr, iConn
If (LieferantenzusatzRec.EOF) Then
    LieferantenzusatzRec.Close
    Call clsError.DefErrPop
    Exit Sub
End If

SQLStr$ = "SELECT * FROM Ausnahmen WHERE LiefNr =" + Str$(ind)
AusnahmenRec.Open SQLStr, iConn
SQLStr$ = "SELECT * FROM RabattTabelle WHERE LiefNr =" + Str$(ind)
SQLStr$ = SQLStr + " ORDER BY Schwellwert"
RabattTabelleRec.Open SQLStr, iConn
SQLStr$ = "SELECT * FROM BmEkTabelle WHERE LiefNr =" + Str$(ind)
SQLStr$ = SQLStr + " ORDER BY AbBM"
BmEkTabelleRec.Open SQLStr, iConn
SQLStr$ = "SELECT * FROM Msv3 WHERE LiefNr =" + Str$(iLief)
Msv3Rec.Open SQLStr, iConn
    

With LifZus1
    .MindestUmsatz = clsOpTool.CheckNullDouble(LieferantenzusatzRec!MindestUmsatz)
    .DirektMindestBM = clsOpTool.CheckNullLong(LieferantenzusatzRec!DirektMindestBM)
    .DirektMindestWert = clsOpTool.CheckNullLong(LieferantenzusatzRec!DirektMindestWert)
    .KostenProLieferung = clsOpTool.CheckNullDouble(LieferantenzusatzRec!KostenProLieferung)

    For j% = 0 To 19
        .LiefFuerHerst(j%) = ""
    Next j%
    j% = 1
    h$ = Trim(clsOpTool.CheckNullStr(LieferantenzusatzRec!LiefFuerHerst))
    For i% = 0 To 20
        ind2 = InStr(h$, ",")
        If (ind2 > 0) Then
            h2 = Trim(Left(h, ind2 - 1))
            .LiefFuerHerst(j%) = h2$
            j% = j% + 1
            h = Mid$(h$, ind2 + 1)
        Else
            Exit For
        End If
    Next i%

    .DepotPflicht = clsOpTool.CheckNullLong(LieferantenzusatzRec!DepotPflicht)
    .Lieferfrist = clsOpTool.CheckNullLong(LieferantenzusatzRec!Lieferfrist)
    .BevorratungsZeitraum = clsOpTool.CheckNullLong(LieferantenzusatzRec!BevorratungsZeitraum)
    .ValutaStellung = clsOpTool.CheckNullLong(LieferantenzusatzRec!ValutaStellung)
    .FakturenRabatt = clsOpTool.CheckNullDouble(LieferantenzusatzRec!FakturenRabatt)
    .FakturenRabattTyp = clsOpTool.CheckNullLong(LieferantenzusatzRec!FakturenRabattTyp)
    .DirektBestFaxKz = clsOpTool.CheckNullLong(LieferantenzusatzRec!DirektBestFaxKz)
    .DirektBestModemKz = clsOpTool.CheckNullLong(LieferantenzusatzRec!DirektBestModemKz)
    .DirektBestMailKz = clsOpTool.CheckNullLong(LieferantenzusatzRec!DirektBestMailKz)
    .DirektBestDruckKz = clsOpTool.CheckNullLong(LieferantenzusatzRec!DirektBestDruckKz)
    .DirektBestComputerFaxKz = clsOpTool.CheckNullLong(LieferantenzusatzRec!DirektBestComputerFaxKz)
    .DirektBestFax = clsOpTool.CheckNullStr(LieferantenzusatzRec!DirektBestFax)
    .DirektBestModem = clsOpTool.CheckNullStr(LieferantenzusatzRec!DirektBestModem)
    .DirektBestMail = clsOpTool.CheckNullStr(LieferantenzusatzRec!DirektBestMail)
    .AngebotWWW = clsOpTool.CheckNullStr(LieferantenzusatzRec!AngebotWWW)
    .TempBevorratungsZeitraum = clsOpTool.CheckNullLong(LieferantenzusatzRec!TempBevorratungsZeitraum)
    .TempValutaStellung = clsOpTool.CheckNullLong(LieferantenzusatzRec!TempValutaStellung)
    .TempFakturenRabatt = clsOpTool.CheckNullDouble(LieferantenzusatzRec!TempFakturenRabatt)
    .DirektMindestBewertung = clsOpTool.CheckNullLong(LieferantenzusatzRec!DirektMindestBewertung)
    .IstDirektLieferant = clsOpTool.CheckNullLong(LieferantenzusatzRec!IstDirektLieferant)
    .ProzentPlus = clsOpTool.CheckNullLong(LieferantenzusatzRec!ProzentPlus)
    .ZusatzAufwand = clsOpTool.CheckNullLong(LieferantenzusatzRec!ZusatzAufwand)
    .ZuordnungenModus = clsOpTool.CheckNullLong(LieferantenzusatzRec!ZuordnungenModus)
    .PreisBasis = clsOpTool.CheckNullLong(LieferantenzusatzRec!PreisBasis)
    
    For i% = 0 To 4
        .AbBM(i%) = 0
        .BisBM(i%) = 0
        .Vergleich(i%) = 0
        .zWert(i%) = 0#
        .BmEkRabatt(i%) = 0#
    Next i
    For i% = 0 To 4
        If (BmEkTabelleRec.EOF) Then
            Exit For
        End If
        
        .AbBM(i) = clsOpTool.CheckNullLong(BmEkTabelleRec!AbBM)
        .BisBM(i) = clsOpTool.CheckNullLong(BmEkTabelleRec!BisBM)
        .Vergleich(i) = clsOpTool.CheckNullLong(BmEkTabelleRec!Vergleich)
        .zWert(i) = clsOpTool.CheckNullLong(BmEkTabelleRec!zWert)
        .BmEkRabatt(i) = clsOpTool.CheckNullDouble(BmEkTabelleRec!Rabatt)
        
'        .BmEkTabGruppe(i) = clsOpTool.CheckNullStr(BmEkTabelleRec!BmEkTabGruppe)
        h$ = Trim(clsOpTool.CheckNullStr(BmEkTabelleRec!BmEkTabGruppe))
        k% = 1
        val2% = 0
        For j% = 1 To 10
            h2$ = Format(j% - 1, "0")
            If (InStr(h$, h2$) > 0) Then
                val2% = val2% + k%
            End If
            k% = k% * 2
        Next j%
        .BmEkTabGruppe(i) = val2%
            
        BmEkTabelleRec.MoveNext
    Next i

    For i% = 0 To 4
        .Schwellwert(i%) = 0#
        .Rabatt(i%) = 0!
    Next i
    For i% = 0 To 4
        If (RabattTabelleRec.EOF) Then
            Exit For
        End If
        
        .Schwellwert(i) = clsOpTool.CheckNullLong(RabattTabelleRec!Schwellwert)
        .Rabatt(i) = clsOpTool.CheckNullDouble(RabattTabelleRec!Rabatt)
        
        RabattTabelleRec.MoveNext
    Next i

    h$ = Trim(clsOpTool.CheckNullStr(AusnahmenRec!AusnahmenKz))
    For i% = 0 To 6
        .AusnahmenKz(i%) = Val(Mid$(h$, i + 1, 1))
    Next i%
    .AusnahmenAmRx = clsOpTool.CheckNullLong(AusnahmenRec!AusnahmenAmRx)
    
    h$ = Trim(clsOpTool.CheckNullStr(AusnahmenRec!AusnahmenWg))
    For i% = 0 To 9
        ind2 = InStr(h$, ",")
        If (ind2 > 0) Then
            h2 = Trim(Left(h, ind2 - 1))
            .AusnahmenWg(i) = Val(h2$)
            h = Mid$(h$, ind2 + 1)
        Else
            Exit For
        End If
    Next i%
    
    .AusnahmenBM = clsOpTool.CheckNullLong(AusnahmenRec!AusnahmenBM)
    .AusnahmenZW = clsOpTool.CheckNullDouble(AusnahmenRec!AusnahmenZW)

    h$ = Trim(clsOpTool.CheckNullStr(AusnahmenRec!AusnahmenHerst))
    For i% = 0 To 5
        ind2 = InStr(h$, ",")
        If (ind2 > 0) Then
            h2 = Trim(Left(h, ind2 - 1))
            .AusnahmenHerst(i%) = h2$
            h = Mid$(h$, ind2 + 1)
        Else
            Exit For
        End If
    Next i%

    .TabTyp = clsOpTool.CheckNullLong(LieferantenzusatzRec!TabTyp)
    .HatAusnahmen = clsOpTool.CheckNullLong(LieferantenzusatzRec!HatAusnahmen)
    .WumsatzLief = clsOpTool.CheckNullLong(LieferantenzusatzRec!WumsatzLief)
    .MvdaProzent = clsOpTool.CheckNullDouble(LieferantenzusatzRec!MvdaProzent)
    .MvdaLief = clsOpTool.CheckNullLong(LieferantenzusatzRec!MvdaLief)
    .IstMvdaLieferant = clsOpTool.CheckNullLong(LieferantenzusatzRec!IstMvdaLieferant)
    .HatFixRabatt = clsOpTool.CheckNullLong(LieferantenzusatzRec!HatFixRabatt)
    For i = 0 To 1
        .FixRabatt(i) = clsOpTool.CheckNullDouble(LieferantenzusatzRec.Fields("FixRabatt" + CStr(i + 1)))
    Next i
        
    .OrgBeobachtSendungen = clsOpTool.CheckNullLong(LieferantenzusatzRec!OrgBeobachtSendungen)
    .BeobachtSendungen = clsOpTool.CheckNullLong(LieferantenzusatzRec!BeobachtSendungen)
    For j = 1 To 2
        .OrgBeobachtUmsatz(j - 1) = clsOpTool.CheckNullDouble(LieferantenzusatzRec.Fields("OrgBeobachtUmsatz" + CStr(j)))
        .BeobachtUmsatz(j - 1) = clsOpTool.CheckNullDouble(LieferantenzusatzRec.Fields("BeobachtUmsatz" + CStr(j)))
        .UmsatzProSendung(j - 1) = clsOpTool.CheckNullDouble(LieferantenzusatzRec.Fields("UmsatzProSendung" + CStr(j)))
        .UmsatzBisher(j - 1) = clsOpTool.CheckNullDouble(LieferantenzusatzRec.Fields("UmsatzBisher" + CStr(j)))
        .PrognoseUmsatz(j - 1) = clsOpTool.CheckNullDouble(LieferantenzusatzRec.Fields("PrognoseUmsatz" + CStr(j)))
    Next j

    .SendungenAlt = clsOpTool.CheckNullLong(LieferantenzusatzRec!SendungenAlt)
    .SendungenPlan = clsOpTool.CheckNullLong(LieferantenzusatzRec!SendungenPlan)
    .AliquotProzent = clsOpTool.CheckNullDouble(LieferantenzusatzRec!AliquotProzent)
    .PrognoseSchwellwert = clsOpTool.CheckNullDouble(LieferantenzusatzRec!PrognoseSchwellwert)
    .PrognoseRabatt = clsOpTool.CheckNullDouble(LieferantenzusatzRec!PrognoseRabatt)
'    .IstAutoDirektLieferant = clsOpTool.CheckNullLong(LieferantenzusatzRec!IstAutoDirektLieferant)
'    .ZuordnungenAktiv = clsOpTool.CheckNullLong(LieferantenzusatzRec!ZuordnungenAktiv)


    .Msv3Aktiv = 0
    .Version20Aktiv = 0
    .AktiveURL = 0
    .EchtURL = ""
    .EchtKundenNr = ""
    .EchtPasswort = ""
    .TestURL = ""
    .TestKundenNr = ""
    .TestPasswort = ""
    .AuftragsartNormal = 0
    .AuftragsartStapel = 0
    .AuftragsartSonder = 0
    .AuftragsartVersand = 0
    .Substitutionsvereinbarung = 0
    .VerfuegbarkeitEinzelnSpezifischeRueckmeldungVereinbart = 0
    .VerfuegbarkeitBulkVereinbart = 0
    .RuecknahmeangebotVereinbart = 0
    .GueltigAb = ""
    For i = 0 To 4
        .VereinbarteAuftragsKennung(i) = ""
    Next i
    .LieferVorgabe = 0
    .AuftragsKennung = ""
    .GhCheckBestellen = 0
    .GhCheckLieferVorgabe = 0
    .GhCheckAuftragKennung = ""

    If Not (Msv3Rec.EOF) Then
        .Msv3Aktiv = Abs(clsOpTool.CheckNullLong(Msv3Rec!Msv3Aktiv))
        .Version20Aktiv = Abs(clsOpTool.CheckNullLong(Msv3Rec!Version20Aktiv))
'        .AktiveURL = Abs(clsOpTool.CheckNullLong(Msv3Rec!AktiveURL))
        .EchtURL = Trim(clsOpTool.CheckNullStr(Msv3Rec!EchtURL))
        .EchtKundenNr = Trim(clsOpTool.CheckNullStr(Msv3Rec!EchtKundenNr))
        .EchtPasswort = Trim(clsOpTool.CheckNullStr(Msv3Rec!EchtPasswort))
'        .TestURL = Trim(clsOpTool.CheckNullStr(Msv3Rec!TestURL))
'        .TestKundenNr = Trim(clsOpTool.CheckNullStr(Msv3Rec!TestKundenNr))
'        .TestPasswort = Trim(clsOpTool.CheckNullStr(Msv3Rec!TestPasswort))
        .AuftragsartNormal = Abs(clsOpTool.CheckNullLong(Msv3Rec!AuftragsartNormal))
        .AuftragsartStapel = Abs(clsOpTool.CheckNullLong(Msv3Rec!AuftragsartStapel))
        .AuftragsartSonder = Abs(clsOpTool.CheckNullLong(Msv3Rec!AuftragsartSonder))
        .AuftragsartVersand = Abs(clsOpTool.CheckNullLong(Msv3Rec!AuftragsartVersand))
        .Substitutionsvereinbarung = Abs(clsOpTool.CheckNullLong(Msv3Rec!Substitutionsvereinbarung))
        .VerfuegbarkeitEinzelnSpezifischeRueckmeldungVereinbart = Abs(clsOpTool.CheckNullLong(Msv3Rec!VerfuegbarkeitEinzelnSpezifischeRueckmeldungVereinbart))
        .VerfuegbarkeitBulkVereinbart = Abs(clsOpTool.CheckNullLong(Msv3Rec!VerfuegbarkeitBulkVereinbart))
        .RuecknahmeangebotVereinbart = Abs(clsOpTool.CheckNullLong(Msv3Rec!RuecknahmeangebotVereinbart))
        .GueltigAb = Trim(clsOpTool.CheckNullStr(Msv3Rec!GueltigAb))
'        For i = 0 To 4
'            .VereinbarteAuftragsKennung(i) = Trim(clsOpTool.CheckNullStr(Msv3Rec.Fields("VereinbarteAuftragsKennung" + CStr(i + 1))))
'        Next i
'        .LieferVorgabe = Abs(clsOpTool.CheckNullLong(Msv3Rec!LieferVorgabe))
'        .AuftragsKennung = Trim(clsOpTool.CheckNullStr(Msv3Rec!AuftragsKennung))
'        .GhCheckBestellen = Abs(clsOpTool.CheckNullLong(Msv3Rec!GhCheckBestellen))
'        .GhCheckLieferVorgabe = Abs(clsOpTool.CheckNullLong(Msv3Rec!GhCheckLieferVorgabe))
'        .GhCheckAuftragKennung = Trim(clsOpTool.CheckNullStr(Msv3Rec!GhCheckAuftragKennung))
    End If
    Msv3Rec.Close
'    If (iDatenVonLief > 0) Then
'        SQLStr$ = "SELECT * FROM Msv3 WHERE LiefNr =" + Str$(ind)
'        Msv3Rec.Open SQLStr, iConn
'        If Not (Msv3Rec.EOF) Then
'            .LieferVorgabe = Abs(clsOpTool.CheckNullLong(Msv3Rec!LieferVorgabe))
'            .AuftragsKennung = Trim(clsOpTool.CheckNullStr(Msv3Rec!AuftragsKennung))
'            .GhCheckBestellen = Abs(clsOpTool.CheckNullLong(Msv3Rec!GhCheckBestellen))
'        End If
'        Msv3Rec.Close
'    End If
End With

LieferantenzusatzRec.Close
AusnahmenRec.Close
RabattTabelleRec.Close
BmEkTabelleRec.Close
'Msv3Rec.Close
    
Call clsError.DefErrPop
End Sub

