VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLagerbewegungDB"
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

Private Const DefErrModul = "LAGERBEWDB.CLS"

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

dKurz$ = "LAGBEW"
dBezeichnung$ = "LagerBewegungen" + sqlop1.SqlGetSuffix
dName$ = "lagerbew.mdb"


'dAnzFieldsDOS% = 21
'
'''''''''''
'fKurz$(0) = "DATUM"
'fKurz$(1) = "LIEF"
'fKurz$(2) = "PZN"
'fKurz$(3) = "FILLER"
'fKurz$(4) = "TEXT"
'fKurz$(5) = "MENGE"
'fKurz$(6) = "REST"
'fKurz$(7) = "MAX"
'''''''''''
'ReDim fKurz$(dAnzFieldsDOS% - 1)
'
'fKurz$(0) = "NA(0)"
'fKurz$(1) = "NA(1)"
'fKurz$(2) = "NA(2)"
'fKurz$(3) = "NA(3)"
'fKurz$(4) = "KURZ"
'fKurz$(5) = "WBEZ"
'fKurz$(6) = "TEL"
'fKurz$(7) = "KTO"
'fKurz$(8) = "BLZ"
'fKurz$(9) = "VERT"
'fKurz$(10) = "KTVOR"
'fKurz$(11) = "BEIN"
'fKurz$(12) = "DW"
'fKurz$(13) = "IDA"
'fKurz$(14) = "FIBU"
'fKurz$(15) = "IDL"
'fKurz$(16) = "DUMMY"
'fKurz$(17) = "WNWT"
'fKurz$(18) = "DIR"
'fKurz$(19) = "SK"
'fKurz$(20) = "SALDO"

Set LbDB1 = Me

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
    LagerbewDBok = True
    
'    LieferantenComm.ActiveConnection = LieferantenConn
    
'    LieferantenRec.Open "SELECT * FROM Lieferanten", LieferantenConn
'    dAnzFields% = LieferantenRec.Fields.Count '+ 1
'    ReDim fBezeichnung$(dAnzFields% + 1)
''    fBezeichnung$(0) = "ImportKz"
'    For i% = 0 To (dAnzFields% - 1)
'        fBezeichnung$(i%) = LieferantenRec.Fields(i%).Name
'    Next i%
'    LieferantenRec.Close
'
'    LieferantenRec.Open "SELECT * FROM LieferantenZusatz", LieferantenConn
'    iAnz% = LieferantenRec.Fields.Count '+ 1
'    ReDim Preserve fBezeichnung$(dAnzFields% + iAnz + 1)
'    For i% = 1 To (iAnz% - 1)
'        fBezeichnung$(dAnzFields) = LieferantenRec.Fields(i%).Name
'        dAnzFields = dAnzFields + 1
'    Next i%
'    LieferantenRec.Close
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

