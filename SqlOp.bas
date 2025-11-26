Attribute VB_Name = "modSqlOp"
Option Explicit

Private Const PLATFORM_ID_DOS = 300
Private Const PLATFORM_ID_OS2 = 400
Private Const PLATFORM_ID_NT = 500
Private Const PLATFORM_ID_OSF = 600
Private Const PLATFORM_ID_VMS = 700

Private Type WKSTA_INFO_102
   wki100_platform_id As Long
   pwki100_computername As Long
   pwki100_langroup As Long
   wki100_ver_major As Long
   wki100_ver_minor As Long
   pwki102_lanroot As Long
   wki102_logged_on_users As Long
End Type

Declare Function NetWkstaGetInfo Lib "netapi32" (ByVal servername As String, ByVal Level As Long, lpBuf As Any) As Long
'Declare Function NetApiBufferFree Lib "netapi32" (ByVal Buffer As Long) As Long
'Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, ByVal hpvSource As Long, ByVal cbCopy As Long)

Declare Function lstrlenW Lib "kernel32" (ByVal lpString As Long) As Long
Declare Function NetServerEnum Lib "netapi32" (strServername As Any, ByVal Level As Long, bufptr As Long, ByVal prefmaxlen As Long, entriesread As Long, totalentries As Long, ByVal servertype As Long, strDomain As Any, resumehandle As Long) As Long
Declare Function NetApiBufferFree Lib "Netapi32.dll" (ByVal lpBuffer As Long) As Long
'Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Const SV_TYPE_SERVER As Long = &H2
Const SV_TYPE_SQLSERVER As Long = &H4
Type SV_100
    platform As Long
    Name As Long
End Type

Public Const SQL_TAXE = 0
Public Const SQL_LIEFERANTEN = 1

Public SqlServer$
Public ConnectionString$, SqlConnectionstring$(1)
Public SqlServerDataPath$, NetworkDataPath$
Public SqlDatabase$(1), SqlServerAktiv%(1)

Public SqlError&
Public SqlErrorDesc$

Public SqlConnectErg%

Public GlobalConn As New ADODB.Connection
Public GlobalComm As New ADODB.Command

Private Const DefErrModul = "SqlOp.bas"

Function SqlInit%()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("SqlInit%")
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
Dim ret%, erg%, ind%, ind2%, SqlServerNötig%
Dim h$, h2$

Dim l As Long
Dim entriesread As Long
Dim totalentries As Long
Dim hREsume As Long
Dim bufptr As Long
Dim Level As Long
Dim prefmaxlen As Long
Dim lType As Long
Dim domain() As Byte
Dim i As Long
Dim sv100 As SV_100
Dim strComputername$, strWorkgroup$, strSqlServers$
Dim SqlServers$(10)
Dim AnzSqlServers%
Dim tRec As New ADODB.Recordset

Dim sKey$, SQLStr$

'SqlDatenDir = Left$(CurDir$, 1) + ":\user\SqlDaten"
'ret% = CreateDirectory(SqlDatenDir)
'If (ret = 0) Then
'    SqlInit = 0
'    Call DefErrPop: Exit Function
'End If
    
SqlConnectionstring(0) = "Provider=Microsoft.Jet.OLEDB.4.0;"
SqlConnectionstring(1) = ""

SqlDatabase(0) = "Taxe"
SqlDatabase(1) = "Lieferanten"
For i = 0 To 1
    SqlServerAktiv(i) = 0
Next i
    
SqlServerNötig = 0
For i = 0 To 1
    sKey = SqlDatabase(i)
    h$ = "0"
    l& = GetPrivateProfileString("Databases", sKey, h$, h$, 2, CurDir + "\SqlOp.ini")
    h = Trim$(Left$(h$, l&))
    If (Val(h) <> 0) Then
        SqlServerAktiv(i) = True
        SqlServerNötig = True
    End If
Next i

If (SqlServerNötig = 0) Then
    SqlServer = ""
    ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;"
    SqlInit = True
    Call DefErrPop: Exit Function
End If

sKey = "SqlServer"
h$ = Space$(50)
l& = GetPrivateProfileString("SqlServer", sKey, h$, h$, 51, CurDir + "\SqlOp.ini")
h = Trim$(Left$(h$, l&))
If (h = "") Then
    SqlServer = ""
    ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;"
    SqlInit = True
    Call DefErrPop: Exit Function
End If


If (h = "?") Then
    AnzSqlServers = 0
    
    strComputername = Environ("ComputerName")
    strWorkgroup = ""
    If (strComputername <> "") Then
        Dim pWrkInfo As Long, WrkInfo(0) As WKSTA_INFO_102
        Dim lResult As Long
        lResult = NetWkstaGetInfo(StrConv("\\" & strComputername, vbUnicode), 102, pWrkInfo)
        If (lResult = 0) Then
           Dim cname As String
           cname = String$(255, 0)
           CopyMemory WrkInfo(0), ByVal pWrkInfo, ByVal Len(WrkInfo(0))
           CopyMemory ByVal cname, ByVal WrkInfo(0).pwki100_langroup, ByVal 255
           strWorkgroup = StripTerminator(StrConv(cname, vbFromUnicode))
        End If
    End If

    strSqlServers = ""
    If (strWorkgroup <> "") Then
        Level = 100
        prefmaxlen = -1
        lType = SV_TYPE_SQLSERVER
        domain = strWorkgroup & vbNullChar
        l = NetServerEnum(ByVal 0&, Level, bufptr, prefmaxlen, entriesread, totalentries, lType, domain(0), hREsume)
        If l = 0 Or l = 234& Then
            For i = 0 To entriesread - 1
                CopyMemory sv100, ByVal bufptr, Len(sv100)
                strSqlServers = strSqlServers + Pointer2stringw(sv100.Name) + vbCrLf
                SqlServers(AnzSqlServers) = Pointer2stringw(sv100.Name)
                AnzSqlServers = AnzSqlServers + 1
                bufptr = bufptr + Len(sv100)
            Next i
        End If
        NetApiBufferFree bufptr '
    End If
    
    h2 = "Computer:" + vbTab + strComputername
    h2 = h2 + vbCrLf + "Workgroup:" + vbTab + strWorkgroup
    h2 = h2$ + vbCrLf + vbCrLf + "Installierte SQL-Server:" + vbCrLf + strSqlServers
    MsgBox (h2$)
Else
    AnzSqlServers = 1
    SqlServers(0) = h$
End If

ConnectionString = ""

sKey = "Provider"
h$ = Space$(50)
l& = GetPrivateProfileString("Connection", "Provider", h$, h$, 51, CurDir + "\SqlOp.ini")
h = Trim$(Left$(h$, l&))
If (h <> "") Then
    ConnectionString = ConnectionString + sKey + "=" + h$ + ";"
End If

sKey = "Integrated Security"
h$ = Space$(50)
l& = GetPrivateProfileString("Connection", sKey, h$, h$, 51, CurDir + "\SqlOp.ini")
h = Trim$(Left$(h$, l&))
If (h <> "") Then
    ConnectionString = ConnectionString + sKey + "=" + h$ + ";"
Else
    sKey = "Trusted Connection"
    h$ = Space$(50)
    l& = GetPrivateProfileString("Connection", sKey, h$, h$, 51, CurDir + "\SqlOp.ini")
    h = Trim$(Left$(h$, l&))
    If (h <> "") Then
        ConnectionString = ConnectionString + sKey + "=" + h$ + ";"
    End If
    
    sKey = "User Id"
    h$ = Space$(50)
    l& = GetPrivateProfileString("Connection", sKey, h$, h$, 51, CurDir + "\SqlOp.ini")
    h = Trim$(Left$(h$, l&))
    If (h <> "") Then
        ConnectionString = ConnectionString + sKey + "=" + h$ + ";"
    End If
    
    sKey = "Password"
    h$ = Space$(50)
    l& = GetPrivateProfileString("Connection", sKey, h$, h$, 51, CurDir + "\SqlOp.ini")
    h = Trim$(Left$(h$, l&))
    If (h <> "") Then
        ConnectionString = ConnectionString + sKey + "=" + h$ + ";"
    End If
End If

sKey = "Initial Catalog"
h$ = Space$(50)
l& = GetPrivateProfileString("Connection", sKey, h$, h$, 51, CurDir + "\SqlOp.ini")
h = Trim$(Left$(h$, l&))
ConnectionString = ConnectionString + sKey + "=" + h$ + ";"
SqlConnectionstring(1) = ConnectionString

'sKey = "Data Source"
'ConnectionString = ConnectionString + sKey + "=" + ";"

On Error Resume Next
'On Error GoTo DefErr
With GlobalConn
    .CursorLocation = adUseServer
    For i = 1 To AnzSqlServers
        SqlServer = SqlServers(i - 1)
'        .ConnectionString = "Provider=SQLOLEDB; Integrated Security=SSPI; Initial Catalog=" + "''" + "; Data Source=" + h$
'        .ConnectionString = "Provider=SQLOLEDB; Trusted Connection=False; User Id=ELVIRA; Password=JESSICA; Initial Catalog=" + "''" + "; Data Source=" + h$
        .ConnectionString = ConnectionString + "Data Source=" + SqlServer
        .Open
        If (Err) Then
            If (InStr(UCase(SqlServer), "EXPRESS") <= 0) Then
                Err.Clear
                SqlServer = SqlServer + "\SQLEXPRESS"
                .ConnectionString = ConnectionString + "Data Source=" + SqlServer
                .Open
            End If
        End If
        If (Err = 0) Then
            ret = True
'            MsgBox ("Verbindung zu SQl-Server '" + SqlServer + "'")
            On Error GoTo DefErr
            GlobalComm.ActiveConnection = GlobalConn
            Exit For
        End If
    Next i
End With
On Error GoTo DefErr

If (ret) Then
    l& = WritePrivateProfileString("SqlServer", "SqlServer", SqlServer$, CurDir + "\SqlOp.ini")

    sKey = "SqlServer"
    h$ = Space$(50)
    l& = GetPrivateProfileString("DataPath", sKey, h$, h$, 51, CurDir + "\SqlOp.ini")
    h = Trim$(Left$(h$, l&))
    If (h = "") Then
        SQLStr = "select physical_name from sys.database_files"
        tRec.Open SQLStr$, GlobalConn
        If Not (tRec.EOF) Then
            h = tRec!physical_name
            ind = 0
            Do
                ind2 = InStr(ind + 1, h, "\")
                If (ind2 > 0) Then
                    ind = ind2
                Else
                    If (ind > 0) Then
                        h = Left(h, ind - 1)
                    End If
                    Exit Do
                End If
            Loop
            l& = WritePrivateProfileString("DataPath", sKey, h$, CurDir + "\SqlOp.ini")
        End If
        tRec.Close
    End If
    SqlServerDataPath = h$
    
    sKey = "Network"
    h$ = Space$(50)
    l& = GetPrivateProfileString("DataPath", sKey, h$, h$, 51, CurDir + "\SqlOp.ini")
    h = Trim$(Left$(h$, l&))
    If (h = "") Then
        h = "Z" + Mid$(SqlServerDataPath, 2)
        l& = WritePrivateProfileString("DataPath", sKey, h$, CurDir + "\SqlOp.ini")
    End If
    NetworkDataPath = h$
'    erg% = CreateDirectory(NetworkDataPath)
Else
    Call MsgBox("Keine Verbindung zu SqlServer" + SqlServer + "möglich!", vbCritical)
End If

SqlInit = ret

Call DefErrPop
End Function

Private Function Pointer2stringw(ByVal l As Long) As String
Dim Buffer() As Byte
Dim nLen As Long '

nLen = lstrlenW(l) * 2
If nLen Then
    ReDim Buffer(0 To (nLen - 1)) As Byte
    CopyMemory Buffer(0), ByVal l, nLen
    Pointer2stringw = Buffer
End If
End Function

'This function is used to stripoff all the unnecessary chr$(0)'s
Private Function StripTerminator(sInput As String) As String
    Dim ZeroPos As Integer
    'Search the first chr$(0)
    ZeroPos = InStr(1, sInput, vbNullChar)
    If ZeroPos > 0 Then
        StripTerminator = Left$(sInput, ZeroPos - 1)
    Else
        StripTerminator = sInput
    End If
End Function

Function SqlCheckDatabase%(sDatabase$, Optional iSqlAktiv% = True)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("SqlCheckDatabase%")
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
Dim ret%
Dim ErrNumber&
Dim SQLStr$
Dim tRec As New ADODB.Recordset
Dim CheckDB As Database

'SQLStr = "SELECT HOST_NAME() as HostName"
'tRec.Open SQLStr$, GlobalConn
'If Not (tRec.EOF) Then
'    MsgBox ("Hostname:  " + tRec!HostName)
'End If
'tRec.Close
'
'SQLStr = "SELECT SUSER_NAME() as UserName"
'tRec.Open SQLStr$, GlobalConn
'If Not (tRec.EOF) Then
'    MsgBox ("Username:  " + tRec!UserName)
'End If
'tRec.Close

ret% = 0
'If (SqlServerAktiv(iSqlDatabase)) Then
If (iSqlAktiv) Then
    SQLStr = "SELECT NAME FROM  Sys.DATABASES"
    tRec.Open SQLStr$, GlobalConn
    Do
        If (tRec.EOF) Then
            Exit Do
        End If
        
        If (UCase(tRec!Name) = UCase(sDatabase)) Then
            ret = True
            Exit Do
        End If
        
        tRec.MoveNext
    Loop
    tRec.Close
Else
    On Error Resume Next
    Err.Clear
    Set CheckDB = OpenDatabase(sDatabase$, False, False)
    ErrNumber = Err.Number
    On Error GoTo DefErr
    
    If (ErrNumber = 0) Then
        CheckDB.Close
        ret = True
    End If
End If
   
SqlCheckDatabase = ret

Call DefErrPop
End Function

Function SqlCreateDatabase%(adoConn As ADODB.Connection, adoComm As ADODB.Command, sDatabase$, sDateien$, sPfad$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("SqlCreateDatabase%")
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
Dim ret%
Dim SQLStr$, SQLPfad$

'SQLPfad = "C" + Mid(sPfad, 2)
SQLPfad = SqlServerDataPath
'SQLPfad = CurDir

ret% = SqlDeleteDatabase(sDatabase)
If (ret) Then
    SQLStr = "CREATE DATABASE " + sDatabase
    SQLStr = SQLStr + " ON (NAME='" + sDateien + "', FILENAME='" + SQLPfad + "\" + sDatabase + ".mdf" + "')"
    SQLStr = SQLStr + " LOG ON (NAME='" + sDateien + "_log" + "', FILENAME = '" + SQLPfad + "\" + sDatabase + "_log.ldf" + "')"
'    SQLStr = SQLStr + " ON (NAME='" + SqlDatenDir + "\" + sDatabase + "', FILENAME='" + SqlDatenDir + "\" + sDatabase + "555.mdf" + "')"
'    SQLStr = SQLStr + " LOG ON (NAME='" + SqlDatenDir + "\" + sDatabase + "_log" + "', FILENAME = '" + SqlDatenDir + "\" + sDatabase + "_log555.ldf" + "')"
    
    ret = SqlCommand(GlobalComm, SQLStr)
End If
If (ret) Then
    ret = SqlConnect(adoConn, adoComm, sDatabase, True, adUseServer)
End If

SqlCreateDatabase = ret

Call DefErrPop
End Function

Function SqlDeleteDatabase%(sDatabase$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("SqlDeleteDatabase%")
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
Dim ret%, gef%
Dim SQLStr$
    
ret = True

gef% = SqlCheckDatabase(sDatabase)
If (gef) Then
    If (ret) Then
        SQLStr = "ALTER DATABASE " + sDatabase + " SET SINGLE_USER WITH ROLLBACK IMMEDIATE"
        ret = SqlCommand(GlobalComm, SQLStr)
    End If
    If (ret) Then
        SQLStr = "DROP DATABASE " + sDatabase
        ret = SqlCommand(GlobalComm, SQLStr)
    End If
End If
    
SqlDeleteDatabase = ret

Call DefErrPop
End Function

Function SqlCommand%(adoComm As ADODB.Command, sCommand$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("SqlCommand%")
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
Dim ret%

ret = True
With adoComm
    .CommandText = sCommand$
    .CommandTimeout = 300   '120
    On Error Resume Next
    .Execute
    If (Err) Then
        SqlError = Err.Number
        SqlErrorDesc = Err.Description
        Call MsgBox("Fehler bei '" + sCommand + "': " + vbCrLf + Str(Err.Number) + " " + Err.Description, vbCritical)
        ret = 0
    End If
    On Error GoTo DefErr
End With
   
SqlCommand = ret

Call DefErrPop
End Function

Function SqlConnect%(adoConn As ADODB.Connection, adoComm As ADODB.Command, sDatabase$, iSqlDatabase%, Optional CursorLocation% = adUseClient)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("SqlConnect%")
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
Dim ret%, ind%
Dim h$, h2$

ret = True
With adoConn
    If (.State = adStateOpen) Then 'Wenn Verbindung besteht
        .Close 'Verbindung trennen
    End If
'    Set adoConn = Nothing 'Objekt löschen
'    Set adoComm = Nothing

'    Set adoConn = New ADODB.Connection
'    Set adoComm = New ADODB.Command
    
    h = ConnectionString
    If (SqlServerAktiv(1)) Then
        h2 = sDatabase
        Do
            ind = InStr(h2, "\")
            If (ind > 0) Then
                h2 = Mid(h2, ind + 1)
            Else
                Exit Do
            End If
        Loop
        ind = InStr(h2, ".")
        If (ind > 0) Then
            h2 = Left(h2, ind - 1)
        End If
    '    .ConnectionString = "Provider=SQLOLEDB; Integrated Security=SSPI; Initial Catalog=" + sDatabase + "; Data Source=" + SqlServer
        h = SqlConnectionstring(1)
        .ConnectionString = Left(h, Len(h) - 1) + h2 + "; Data Source=" + SqlServer
    Else
        .ConnectionString = SqlConnectionstring(0) + "Data Source=" + sDatabase
    End If
    .CursorLocation = adUseClient   'adUseServer
    .CommandTimeout = 300
    On Error Resume Next
    Err.Clear
    .Open
    If (Err) Then
        SqlError = Err.Number
        SqlErrorDesc = Err.Description
        Call MsgBox("Fehler bei '" + .ConnectionString + "': " + vbCrLf + Str(Err.Number) + " " + Err.Description, vbCritical)
        ret = 0
    End If
    On Error GoTo DefErr
    adoComm.ActiveConnection = adoConn
End With

SqlConnect = ret

Call DefErrPop
End Function

Function SqlRenameDatabase%(sDatabaseOld$, sDatabaseNew$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("SqlRenameDatabase%")
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
Dim ret%
Dim SQLStr$, SQLPfad$
    
'SQLPfad = "C" + Mid(TaxeDBdir, 2)
'SQLPfad = TaxeDBdir

ret% = SqlCheckDatabase(sDatabaseOld)
If (ret) Then
    SQLStr = "ALTER DATABASE " + sDatabaseOld + " SET OFFLINE"
    ret = SqlCommand(GlobalComm, SQLStr)
End If
If (ret) Then
    ret% = SqlDeleteDatabase(sDatabaseNew)
End If
If (ret) Then
'    Name VorabTaxeDir + "\" + sDatabaseNew + ".mdf" As TaxeDBdir + "\" + sDatabaseNew + ".mdf"
'    SQLStr = "ALTER DATABASE " + sDatabaseOld + " MODIFY file(Name='" + sDatabaseNew + "', FILENAME='" + SQLPfad + "\" + sDatabaseNew + ".mdf'" + ")"
    Name NetworkDataPath + "\" + sDatabaseOld + ".mdf" As NetworkDataPath + "\" + sDatabaseNew + ".mdf"
    SQLStr = "ALTER DATABASE " + sDatabaseOld + " MODIFY file(Name='" + sDatabaseNew + "', FILENAME='" + SqlServerDataPath + "\" + sDatabaseNew + ".mdf'" + ")"
    ret = SqlCommand(GlobalComm, SQLStr)
End If
If (ret) Then
'    Name VorabTaxeDir + "\" + sDatabaseNew + "_log.ldf" As TaxeDBdir + "\" + sDatabaseNew + "_log.ldf"
'    SQLStr = "ALTER DATABASE " + sDatabaseOld + " MODIFY file(Name='" + sDatabaseNew + "_log" + "', FILENAME='" + SQLPfad + "\" + sDatabaseNew + "_log.ldf'" + ")"
    Name NetworkDataPath + "\" + sDatabaseNew + "_log.ldf" As NetworkDataPath + "\" + sDatabaseNew + "_log.ldf"
    SQLStr = "ALTER DATABASE " + sDatabaseOld + " MODIFY file(Name='" + sDatabaseNew + "_log" + "', FILENAME='" + SqlServerDataPath + "\" + sDatabaseNew + "_log.ldf'" + ")"
    ret = SqlCommand(GlobalComm, SQLStr)
End If
If (ret) Then
    SQLStr = "ALTER DATABASE " + sDatabaseOld + " SET ONLINE"
    ret = SqlCommand(GlobalComm, SQLStr)
End If

If (ret) Then
    SQLStr = "ALTER DATABASE " + sDatabaseOld + " MODIFY NAME=" + sDatabaseNew
    ret = SqlCommand(GlobalComm, SQLStr)
End If

''If (ret) Then
''    SQLStr = "ALTER DATABASE " + sDatabaseNew + " MODIFY file(Name='" + SqlDatenDir + "\" + sDatabaseOld + "', NEWNAME='" + SqlDatenDir + "\" + sDatabaseNew + "')"
''    ret = SqlCommand(GlobalComm, SQLStr)
''End If
'If (ret) Then
''    SQLStr = "ALTER DATABASE " + sDatabaseNew + " MODIFY file(Name='" + sDatabaseNew + "_log" + "', FILENAME='" + TaxeDBdir + "\" + sDatabaseNew + "_log" + ".mdf'" + ")"
'    SQLStr = "ALTER DATABASE " + sDatabaseNew + " MODIFY file(Name='" + sDatabaseNew + "_log" + "', FILENAME='" + "c:\taxe" + "\" + sDatabaseNew + "_log" + ".mdf'" + ")"
''    SQLStr = "ALTER DATABASE " + sDatabaseNew + " MODIFY file(Name='" + SqlDatenDir + "\" + sDatabaseOld + "_log" + "', NEWNAME='" + SqlDatenDir + "\" + sDatabaseNew + "_log" + "')"
'    ret = SqlCommand(GlobalComm, SQLStr)
'End If
'If (ret) Then
''    If (Dir(aFile$) <> "") Then Kill aFile$
''    Name SqlDatenDir + "\" + sDatabaseOld + ".mdf" As SqlDatenDir + "\" + sDatabaseNew + ".mdf"
''    SQLStr = "ALTER DATABASE " + sDatabaseNew + " MODIFY file(Name='" + sDatabaseOld + "', FILENAME='" + SqlDatenDir + "\" + sDatabaseNew + ".mdf'" + ")"
''    ret = SqlCommand(GlobalComm, SQLStr)
'End If
''If (ret) Then
'''    If (Dir(aFile$) <> "") Then Kill aFile$
''    Name SqlDatenDir + "\" + sDatabaseOld + "_log" + ".mdf" As SqlDatenDir + "\" + sDatabaseNew + "_log" + ".mdf"
''    SQLStr = "ALTER DATABASE " + sDatabaseNew + " MODIFY file(Name='" + sDatabaseOld + "_log" + "', FILENAME='" + SqlDatenDir + "\" + sDatabaseNew + "_log" + ".mdf'" + ")"
''    ret = SqlCommand(GlobalComm, SQLStr)
''End If
''If (ret) Then
''    SQLStr = "ALTER DATABASE " + sDatabaseNew + " MODIFY file(Name=" + sDatabaseOld + "_log" + ", NEWNAME=" + sDatabaseNew + "_log" + ", FILENAME='" + SqlDatenDir + "\" + sDatabaseNew + "_log.ldf'" + ")"
''    ret = SqlCommand(GlobalComm, SQLStr)
''End If

SqlRenameDatabase = ret

Call DefErrPop
End Function

Function OpenLieferantenDB%()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("OpenLieferantenDB%")
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
Dim DbOk%, ret%
Dim ErrNumber&
Dim DBname$, h$, s$

ret% = False

If (SqlServerAktiv(SQL_LIEFERANTEN)) Then
    DBname$ = Lieferanten.DateiBezeichnung
Else
    DBname$ = Lieferanten.DateiName
End If
DbOk = SqlCheckDatabase(DBname, SqlServerAktiv(SQL_LIEFERANTEN))
If (DbOk) Then
'    h = ConnectionString
    If (SqlServerAktiv(SQL_LIEFERANTEN)) Then
        h = SqlConnectionstring(1)
        LieferantenConn.ConnectionString = Left(h, Len(h) - 1) + DBname + "; Data Source=" + SqlServer
    Else
        LieferantenConn.ConnectionString = SqlConnectionstring(0) + "Data Source=" + DBname
    End If
    
'    OpenCreateLieferantenMDB% = Lieferanten.OpenDatenbank("", LieferantenConn)
'    ret = Lieferanten.OpenDatenbank("", LieferantenConn.ConnectionString)
    ret% = (Lieferanten.OpenDatenbank("", LieferantenConn) = 0)
    
'    LieferantenConn.CursorLocation = adUseClient
'    On Error Resume Next
'    Err.Clear
'    LieferantenConn.Open
'    ErrNumber = Err.Number
    LieferantenComm.ActiveConnection = LieferantenConn
End If

If (ret = 0) Then
    s$ = "ACHTUNG:" + vbCrLf + vbCrLf
    s$ = s$ + "Lieferanten-Datenbank NICHT vorhanden!" + vbCrLf + vbCrLf
    s$ = s$ + " Bitte durch Aufruf der Lieferanten-Stammdaten konvertieren!"
    Call MessageBox(s$, vbCritical, "Abholer-Verwaltung")
End If
        
OpenLieferantenDB% = ret

Call DefErrPop
End Function



