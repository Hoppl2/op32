VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsHilfstaxeDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim dKurz$
Dim dBezeichnung$, dPruefBezeichnung$
Dim dName$
Dim dZugriff$
Dim dAnzFields%, dAnzFieldsDOS%

Dim fKurz$()
Dim fBezeichnung$()

Dim SqlServerAktiv%

Dim iConn As New ADODB.Connection
'Dim iComm As New ADODB.Command

Private Const DefErrModul = "HILFSTAXEDB.CLS"

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

dKurz$ = "HTAXE"
dBezeichnung$ = "HilfsTaxe"
dName$ = "Hilfstax.mdb"
dZugriff$ = ""

dPruefBezeichnung = "Artikel"
dBezeichnung = dBezeichnung + sqlop1.SqlGetSuffix

dAnzFieldsDOS% = 8

ReDim fKurz$(dAnzFieldsDOS% - 1)

fKurz$(0) = "PZN"
fKurz$(1) = "KURZ"
fKurz$(2) = "MENG"
fKurz$(3) = "MEH"
fKurz$(4) = "AEP"
fKurz$(5) = "KP"
fKurz$(6) = "AVP"
fKurz$(7) = "GSTUFE"

Set HilfstaxeDB1 = Me

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

Function OpenDB%(Optional MitFeldern% = True)
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
iConn.CommandTimeout = 300
On Error Resume Next
Err.Clear
iConn.Open
ErrNumber = Err.Number

On Error GoTo DefErr
If (MitFeldern) Then
    If (ErrNumber% = 0) Then
        HilfstaxeDBok = True
        
        HilfstaxeRec.Open "SELECT * FROM Artikel", iConn
        dAnzFields% = HilfstaxeRec.Fields.Count '+ 1
        ReDim fBezeichnung$(dAnzFields% + 1)
    '    fBezeichnung$(0) = "ImportKz"
        For i% = 0 To (dAnzFields% - 1)
            fBezeichnung$(i%) = HilfstaxeRec.Fields(i%).Name
        Next i%
        HilfstaxeRec.Close
    
'        LieferantenRec.Open "SELECT * FROM LieferantenZusatz", iConn
'        iAnz% = LieferantenRec.Fields.Count '+ 1
'        ReDim Preserve fBezeichnung$(dAnzFields% + iAnz + 1)
'        For i% = 1 To (iAnz% - 1)
'            fBezeichnung$(dAnzFields) = LieferantenRec.Fields(i%).Name
'            dAnzFields = dAnzFields + 1
'        Next i%
'        LieferantenRec.Close
    End If
End If

OpenDB = (ErrNumber = 0)

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
    ret = OpenDB(False)
End If

CreateDB = ret

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


'Function OpenDatenbank%(DBName$, sConn As ADODB.Connection)
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
'If (DBName$ = "") Then
'    DBName$ = dName$
'End If
'
'For j% = 1 To 2
'    On Error Resume Next
'    Err.Clear
'
'    artikelConn.ConnectionString = sConn.ConnectionString
'    artikelConn.CursorLocation = adUseClient
'    On Error Resume Next
'    Err.Clear
'    artikelConn.Open
'    ErrNumber = Err.Number
'
'    On Error GoTo DefErr
'    If (ErrNumber% = 0) Then
'        ArtikelDBok = True
'
'        artikelComm.ActiveConnection = artikelConn
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
'        Set sConn = artikelConn
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
'
'Function TaxeValue$(FeldName$)
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Call clsError.DefErrFnc("TaxeValue$")
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
'Dim ret$, h$, iFeldName$
'Dim fldRecordset As Field
'
'ret$ = ""
'
'iFeldName$ = UCase(FeldName$)
'
'TaxeValue$ = " " + ret$
'
'Call clsError.DefErrPop
'End Function
'
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
'
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
'
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

For i = 1 To dAnzFieldsDOS
    If (fKurz(i - 1) = FeldName) Then
        FeldName = ArtikelRec.Fields(i).Name
        Exit For
    End If
Next i

If (UCase(FeldName) = "WG1") Then
    ret$ = Left(Format(clsOpTool.CheckNullLong(ArtikelRec!wg), "00"), 1)
ElseIf (UCase(FeldName) = "WG2") Then
    ret$ = Right(Format(clsOpTool.CheckNullLong(ArtikelRec!wg), "00"), 1)
Else
    Set fldRecordset = ArtikelRec.Fields(FeldName$)
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

FeldWert$ = " " + ret$

'For i% = 0 To (recTaxe.Fields.Count - 1)
'    Set fldRecordset = recTaxe.Fields(i%)
'    Debug.Print fldRecordset.Name; fldRecordset.Size; fldRecordset.Type; fldRecordset.Value
''    Call FieldOutput("Recordsrt", fldRecordset)
'Next i%

Call clsError.DefErrPop
End Function
'
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
sRec.Open "SELECT COUNT(*) AS Anz FROM " + "Artikel", iConn
If (sRec.RecordCount <> 0) Then
    ret& = sRec!anz
End If
sRec.Close

AnzRec& = ret&
End Property
'
'Function Reorg%()
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Call clsError.DefErrFnc("Reorg%")
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
'Dim ret%
'
'ret% = True
'
''MerkzettelDB.Close
'ret% = ReorgLieferanten%
'
'Reorg% = ret%
'
'Call clsError.DefErrPop
'End Function
'
'Function ReorgLieferanten%()
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Call clsError.DefErrFnc("ReorgLieferanten%")
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
'Dim i&, AnzRec&, ErrorNumber&
'Dim s$, OrgMdb$, CompactMdb$
'
'AnzRec& = 100
'
'Load frmReorg
'With frmReorg
'    .Caption = "Lieferanten-Daten reorganisieren"
'    .lbReorg.Caption = "Komprimieren der Lieferanten-Daten ..."
'    .MaxAnz (AnzRec&)
'    .Show
'End With
'
'OrgMdb$ = CurDir$ + "\" + dName$
'CompactMdb$ = CurDir$ + "\Lifer2.mdb"
'
'If (Dir(CompactMdb$) <> "") Then
'    Kill CompactMdb$
'End If
'
''Compact Database to New Name
'On Error Resume Next
'DBEngine.CompactDatabase OrgMdb$, CompactMdb$
'ErrorNumber = Err.Number
'Err.Clear
'On Error GoTo DefErr
'
'i& = AnzRec&
'frmReorg.ShowFortschritt (i&)
'Unload frmReorg
'
'If (ErrorNumber <> 0) Then
'    s$ = "Die Datei ist noch geöffnet - Reorganisieren nicht möglich!"
'    Call clsDialog.MessageBox(s$, vbInformation)
'Else
'    FileCopy OrgMdb$, CurDir$ + "\Liferant.mdb"
'    FileCopy CompactMdb$, OrgMdb$
'End If
'
'ReorgLieferanten% = True
'
'Call clsError.DefErrPop
'End Function
'
'Function ErsterFreier%()
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Call clsError.DefErrFnc("ErsterFreier%")
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
'Dim i%, frei%, ret%
'Dim h$, SQLStr$
'
'On Error Resume Next
'LieferantenRec.Close
'Err.Clear
'On Error GoTo DefErr
'
'ret% = 1
'SQLStr$ = "SELECT * FROM Lieferanten ORDER BY LiefNr"
'LieferantenRec.Open SQLStr$, LieferantenConn
'Do
'    If (LieferantenRec.EOF) Then
'        Exit Do
'    End If
'
'    If (LieferantenRec!LiefNr <> ret) Then
'        Exit Do
'    End If
'
'    ret = ret + 1
'    LieferantenRec.MoveNext
'Loop
'LieferantenRec.Close
'
'ErsterFreier% = ret%
'
'Call clsError.DefErrPop
'End Function
'
'Function IstFrei%(pos%)
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Call clsError.DefErrFnc("IstFrei%")
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
'Dim i%, frei%, ret%
'Dim h$, SQLStr$
'
'ret% = True
'
'On Error Resume Next
'LieferantenRec.Close
'Err.Clear
'On Error GoTo DefErr
'
'SQLStr$ = "SELECT * FROM Lieferanten WHERE LiefNr =" + Str$(pos)
'LieferantenRec.Open SQLStr$, LieferantenConn
'If Not (LieferantenRec.EOF) Then
'    h$ = Trim(clsOpTool.CheckNullStr(LieferantenRec!kurz))
'    h$ = UCase(Trim$(h$))
'    If (h$ <> "") Then
'        If (Asc(Left$(h$, 1)) >= 32) Then
'            ret% = 0
'        End If
'    End If
'
'    If (ret%) Then
'        h$ = Trim(clsOpTool.CheckNullStr(LieferantenRec!Name1))
'        h$ = UCase(Trim$(h$))
'        If (h$ <> "") Then
'            If (Asc(Left$(h$, 1)) >= 32) Then
'                ret% = 0
'            End If
'        End If
'    End If
'End If
'LieferantenRec.Close
'
'IstFrei% = ret%
'
'Call clsError.DefErrPop
'End Function
'
'Sub DeleteRecord(Optional recno& = 0, Optional sSqlServer$ = "")
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Call clsError.DefErrFnc("DeleteRecord")
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
'Dim i%
'Dim sTabelle$(4), SQLStr$
'
'sTabelle(0) = "Lieferanten"
'sTabelle(1) = "LieferantenZusatz"
'sTabelle(2) = "RabattTabelle"
'sTabelle(3) = "BmEkTabelle"
'sTabelle(4) = "Ausnahmen"
'For i = 0 To 4
'    If (sSqlServer <> "") Then
'        SQLStr$ = "DELETE " + sTabelle(i) + " WHERE LiefNr =" + Str$(recno)
'    Else
'        SQLStr$ = "DELETE * FROM " + sTabelle(i) + " WHERE LiefNr =" + Str$(recno)
'    End If
'    LieferantenComm.CommandText = SQLStr$
'    LieferantenComm.CommandTimeout = 120
'    LieferantenComm.Execute
'Next i
'
'Call clsError.DefErrPop
'End Sub
'
'Public Property Get LiefName$(ByVal ind%)
'LiefName$ = iLiefNamen$(ind%)
'End Property
'Public Property Let LiefName(ByVal ind%, ByVal vNewValue$)
'iLiefNamen$(ind%) = vNewValue$
'End Property
'
'Public Property Get AnzLiefNamen%()
'AnzLiefNamen% = iAnzLiefNamen%
'End Property
'Public Property Let AnzLiefNamen(ByVal vNewValue%)
'iAnzLiefNamen% = vNewValue%
'End Property
'
'Public Property Get OrgLiefNr%()
'OrgLiefNr% = iOrgLiefNr
'End Property
'Public Property Let OrgLiefNr(ByVal vNewValue%)
'iOrgLiefNr% = vNewValue%
'End Property
'
'Public Property Get OrgKurz$()
'OrgKurz = iOrgKurz
'End Property
'Public Property Let OrgKurz(ByVal vNewValue$)
'iOrgKurz = vNewValue$
'End Property

Public Sub ZusatzBefuellen(pzn$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("ZusatzBefuellen")
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
Dim FabsRecno&
Dim txt$, SQLStr$
Dim lInfoRec As New ADODB.Recordset

With frmZusatz.flxZusatz
    '* einlesen der Infozeilen
'        .Rows = .FixedRows

    SQLStr$ = "SELECT * FROM ArtikelInfo WHERE Pzn = " + pzn$
    SQLStr$ = SQLStr$ + " ORDER BY LaufNr"
'        Set lInfoRec = LieferantenDB.OpenRecordset(SQLStr$)
    lInfoRec.Open SQLStr, iConn   ' LieferantenConn
    For i% = 0 To (ZusatzAnzTxt% - 1)    '99
        If (lInfoRec.EOF) Then
            Exit For
        End If
        
        txt$ = lInfoRec!TextZeile
        .TextMatrix(i% + 1, 0) = txt$
    
        lInfoRec.MoveNext
    Next i
    lInfoRec.Close

    If (.Rows = .FixedRows) Then
        .AddItem " "
    End If
    
    .row = 1
    .col = 0
    .ColSel = .Cols - 1
End With

Call clsError.DefErrPop
End Sub

Public Sub ZusatzSpeichern(pzn$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("ZusatzSpeichern")
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
Dim i%, j%, Leer%
Dim FabsRecno&, lRecs&
Dim txt$, SQLStr$

If (SqlServerAktiv) Then
    SQLStr$ = "DELETE ArtikelInfo WHERE Pzn = " + pzn
Else
    SQLStr$ = "DELETE * FROM ArtikelInfo WHERE Pzn = " + pzn
End If
'ArtikelDB1.ActiveConn.CommandTimeout = 120
Call iConn.Execute(SQLStr, lRecs, adExecuteNoRecords)
    
With frmZusatz.flxZusatz
    j% = 1
    For i% = .FixedRows To (.Rows - 1)
        txt$ = Trim$(.TextMatrix(i%, 0))
        If (txt$ <> "") Then
            SQLStr = "INSERT INTO ArtikelInfo (Pzn,LaufNr,TextZeile)"
            SQLStr = SQLStr + " VALUES (" + pzn + "," + CStr(j) + "," + "'" + clsOpTool.SqlString(txt$) + "'"
            SQLStr = SQLStr + ")"
'            ArtikelDB1.ActiveConn.CommandTimeout = 120
            Call iConn.Execute(SQLStr, lRecs, adExecuteNoRecords)
            
            j = j + 1
        End If
    Next i%
End With

Call clsError.DefErrPop
End Sub

Sub DB2DOS(ArtRec As ADODB.Recordset)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("DB2DOS")
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

With hTax1
    .pzn = clsOpTool.PznString(ArtRec!pzn)
    .kurz = Left$(Trim(ArtRec!Name) + Space$(Len(.kurz)), Len(.kurz))
    .meng = Right(Space(Len(.meng)) + Trim(ArtRec!menge), Len(.meng))
    .meh = Left(Trim(ArtRec!einheit) + Space(Len(.meh)), Len(.meh))
    .Gstufe = Right(Space(Len(.Gstufe)) + Trim(ArtRec!Gstufe), Len(.Gstufe))
   
    .aep = ArtRec!EK
    .kp = ArtRec!kp
    .AVP = ArtRec!vk

    .mw = String(1, 0)
    .ArzneiBuch = String(2, 0)
    .MinAns = String(1, 0)
    .wg = String(1, 0)
    .vc = String(1, 0)
    .vDatum = String(4, 0)
    .LichtSchutz = String(1, 0)
    .SepVen = String(1, 0)
    .cave = String(1, 0)
    .aufbew = String(2, 0)
    .sucht = String(1, 0)
    .vet = String(1, 0)
    .rez = String(2, 0)
    .gal = String(1, 0)
    .Status = String(1, 0)
    .lager1 = String(1, 0)
    .lager2 = String(1, 0)
    .alkohol = String(1, 0)
    .TextSatz = 0
    .Abgrenz = String(2, 0)
    .kas = String(3, 0)
    .Gruppe = String(1, 0)
'    .rest = String(20, 0)
    
'    .KurzNeu = Left$(Trim(ArtRec!Name) + Space(48), 48)
'    .MengNeu = Right(Space(7) + Trim(ArtRec!menge), 7)
    .KurzNeu = Trim(ArtRec!Name)
    .MengNeu = Trim(ArtRec!menge)
End With

Call clsError.DefErrPop
End Sub

Function OpenRecordset%(ArtRec As ADODB.Recordset, SqlCommand$, Optional DosStruktur% = True)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("OpenRecordset%")
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
'Dim ret%

On Error Resume Next
ArtRec.Close
Err.Clear
On Error GoTo DefErr
ArtRec.Open SqlCommand, iConn
FabsErrf = Abs(ArtRec.EOF)
If (FabsErrf = 0) Then
    If (DosStruktur) Then
        Call DB2DOS(ArtRec)
    End If
End If

OpenRecordset = FabsErrf

Call clsError.DefErrPop
End Function



