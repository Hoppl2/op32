VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsVerkaufDB"
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

Private Const DefErrModul = "VERKAUFDB.CLS"

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

dKurz$ = "VK"
dBezeichnung$ = "Verkauf"
dName$ = "Verkauf.mdb"
dZugriff$ = ""

dPruefBezeichnung = dBezeichnung
dBezeichnung = dBezeichnung + sqlop1.SqlGetSuffix

Set VerkaufDB1 = Me

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
Dim h$, SQLStr$

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
        VerkaufDbOk = True
        
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

Public Property Get AnzFelderDos%()
AnzFelderDos% = dAnzFieldsDOS%
End Property

Public Property Get AnzRec&()
Dim ret&
Dim sRec As New ADODB.Recordset

ret& = 0
sRec.Open "SELECT COUNT(*) AS Anz FROM " + "Verkauf", iConn
If (sRec.RecordCount <> 0) Then
    ret& = sRec!anz
End If
sRec.Close

AnzRec& = ret&
End Property

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
Dim i%

'With Ast1
'    .pzn = clsOpTool.PznString(ArtRec!pzn)
'    .kurz = Left$(Trim(ArtRec!Name) + Space$(Len(.kurz)), Len(.kurz))
'    .meng = Right(Space(Len(.meng)) + Trim(ArtRec!menge), Len(.meng))
'    .meh = Left(Trim(ArtRec!einheit) + Space(Len(.meh)), Len(.meh))
'    .aep = ArtRec!EK
'    .kp = ArtRec!kp
'    .AVP = ArtRec!vk
'    .mw = ArtRec!Mwst
'    .ka = ArtRec!KzAufschlag
''    .wg = Right(Space(Len(.wg)) + CStr(ArtRec!wg), Len(.wg))
'    .wg = Left(CStr(ArtRec!wg) + Space(Len(.wg)), Len(.wg))
'    .herst = Left$(Trim(ArtRec!HerstellerKB) + Space$(Len(.herst)), Len(.herst))
'    .lic = Left(Trim(ArtRec!lic) + Space(Len(.lic)), Len(.lic))
'    .vc = Left(Trim(ArtRec!vc) + Space(Len(.vc)), Len(.vc))
'    .gc = Left(Trim(ArtRec!gc) + Space(Len(.gc)), Len(.gc))
'    .lac = Left(Trim(ArtRec!lac) + Space(Len(.lac)), Len(.lac))
'    .rez = Left(Trim(ArtRec!rez) + Space(Len(.rez)), Len(.rez))
'    .kas = Left(Trim(ArtRec!kas) + Space(Len(.kas)), Len(.kas))
'    .kasz = Left(Trim(ArtRec!kasz) + Space(Len(.kasz)), Len(.kasz))
'    .zuza = ArtRec!zuza
'    .Vnr = Left(Trim(ArtRec!Vnr) + Space(Len(.Vnr)), Len(.Vnr))
'    .kz = Left(Trim(ArtRec!kz) + Space(Len(.kz)), Len(.kz))
'    .abl = Left(Trim(ArtRec!abl) + Space(Len(.abl)), Len(.abl))
'
'    .KurzNeu = Trim(ArtRec!Name)
'    .MengNeu = Trim(ArtRec!menge)
'End With
'
''If (para.Land = "A") Then
''    SQLStr = SQLStr + ", ArtGruppe=" + CStr(ast.ArtGruppe)
''    If (Left(ast.wg, 1) = "3") Then
''        SQLStr = SQLStr + ", PZN2=" + CStr(Val(ast.pzn2))
''        SQLStr = SQLStr + ", Menge2='" + ast.meng2 + "'"
''    End If
''End If
'
'
'With Ass1
'    .pzn = clsOpTool.PznString(ArtRec!pzn)
'    .poslag = ArtRec!poslag
'    .opt = ArtRec!OptBM
'    .lief = ArtRec!lief
'    .bm = ArtRec!bm
'    .mm = ArtRec!mm
'    .vbm = ArtRec!vbm
'    .vmm = ArtRec!vmm
'    .halt = Left(Trim(clsOpTool.CheckNullStr(ArtRec!halt)) + Space(Len(.halt)), Len(.halt))
'
''    .lld = clsOpTool.iDate(Format(ArtRec!LetztLieferDatum, "DDMMYY"))
'    .lld = clsOpTool.Date2OpDatum(ArtRec!LetztLieferDatum)
'    .llm = ArtRec!LetztLieferMenge
'
''    .abl1 = clsOpTool.iDate(Format(ArtRec!AblaufDatum1, "DDMMYY"))
''    .abl2 = clsOpTool.iDate(Format(ArtRec!AblaufDatum2, "DDMMYY"))
''    .anlage = clsOpTool.iDate(Format(ArtRec!AnlageDatum, "DDMMYY"))
'    .abl1 = clsOpTool.Date2OpDatum(ArtRec!AblaufDatum1)
'    .abl2 = clsOpTool.Date2OpDatum(ArtRec!AblaufDatum2)
'    .anlage = clsOpTool.Date2OpDatum(ArtRec!AnlageDatum)
'
'    .lag = ArtRec!lag
'    .pp = Left(Trim(clsOpTool.CheckNullStr(ArtRec!pp)) + Space(Len(.pp)), Len(.pp))
'    .ean = Right(Space(Len(.ean)) + Trim(clsOpTool.CheckNullStr(ArtRec!ean)), Len(.ean))
'
'    .tplatz = ArtRec!TaraPlatz
'    .tlager = ArtRec!TaraLager
'    .tmm = ArtRec!TaraMM
'
'    .fplatz = ArtRec!FilialePlatz
'    .flager = ArtRec!FilialeLager
'    .fmm = ArtRec!FilialeMM
'
'    .pk = ArtRec!pk
'    .lagbew = Left(Trim(clsOpTool.CheckNullStr(ArtRec!LagerBew)) + Space(Len(.lagbew)), Len(.lagbew))
'    .lac2 = Left(Trim(clsOpTool.CheckNullStr(ArtRec!lac2)) + Space(Len(.lac2)), Len(.lac2))
'    .RowaUeberLagerKz = Left(Trim(clsOpTool.CheckNullStr(ArtRec!Rowa‹berlagerKz)) + Space(Len(.RowaUeberLagerKz)), Len(.RowaUeberLagerKz))
'
'    For i = 1 To 4
'        .max(i - 1) = ArtRec.Fields("TaraEinzelPlatz" + CStr(i))
'        .akt(i - 1) = ArtRec.Fields("TaraEinzelLager" + CStr(i))
'        .min(i - 1) = ArtRec.Fields("TaraEinzelMM" + CStr(i))
'    Next i
'
'End With
'
''With flxStammdaten6(0)
''    For i% = 0 To 3
''        h$ = h$ + vbTab + Format(ass.max(i%), "0")
''    Next i%
''    For i% = 0 To 3
''        h$ = h$ + vbTab + Format(ass.min(i%), "0")
''    Next i%
''    h$ = "Ist"
''    For i% = 0 To 3
''        h$ = h$ + vbTab + Format(ass.akt(i%), "0")
''    Next i%
''End With


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

