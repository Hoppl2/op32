VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSqlTools"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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

Private Declare Function NetWkstaGetInfo Lib "netapi32" (ByVal servername As String, ByVal Level As Long, lpBuf As Any) As Long
'Declare Function NetApiBufferFree Lib "netapi32" (ByVal Buffer As Long) As Long
'Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, ByVal hpvSource As Long, ByVal cbCopy As Long)

Private Declare Function lstrlenW Lib "kernel32" (ByVal lpString As Long) As Long
Private Declare Function NetServerEnum Lib "netapi32" (strServername As Any, ByVal Level As Long, bufptr As Long, ByVal prefmaxlen As Long, entriesread As Long, totalentries As Long, ByVal servertype As Long, strDomain As Any, resumehandle As Long) As Long
Private Declare Function NetApiBufferFree Lib "Netapi32.dll" (ByVal lpBuffer As Long) As Long
'Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Const SV_TYPE_SERVER As Long = &H2
Const SV_TYPE_SQLSERVER As Long = &H4
Private Type SV_100
    platform As Long
    Name As Long
End Type

'Public Const SQL_TAXE = 0
'Public Const SQL_LIEFERANTEN = 1

'Public SqlServer$
'Public SqlConnectionstring$(1)
Dim iSqlServers$(10), iSqlConnectionString$(10)
Dim AnzSqlServers%

Dim iSqlServerDataPath$, iNetworkDataPath$
Dim iSqlServerDB$(100)
Dim iSqlServerDBind%(100)
Dim iSqlSuffix$(10)
Dim iTaxeDatum$

Dim SqlError&
Dim SqlErrorDesc$

Dim SqlConnectErg%

Dim GlobalConn(10) As New ADODB.Connection
Dim GlobalComm(10) As New ADODB.Command

Private Const DefErrModul = "SqlTools.cls"

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

Set sqlop1 = Me

Call clsError.DefErrPop
End Sub

Function SqlInit%(Optional MitStatus% = 0)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("SqlInit%")
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
Dim j%, ret%, erg%, ind%, ind2%, SqlServerNötig%
Dim h$, h2$
Dim ConnectionString$

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
Dim tRec As New ADODB.Recordset

Dim sKey$, sSection$, SQLStr$, SqlServer$, sProtocol$

h$ = Space$(50)
l& = GetPrivateProfileString("LetzteAktionen", "AktiveTaxe", h$, h$, 51, wPara1.LwZ + "\user\SysManag.ini")
iTaxeDatum = Trim$(Left$(h$, l&))
If (iTaxeDatum <> "") Then
    iTaxeDatum = "_" + iTaxeDatum
End If

iSqlConnectionString(0) = "Provider=Microsoft.Jet.OLEDB.4.0;"
For i = 1 To 10
    iSqlServers(i) = ""
    iSqlConnectionString(i) = ""
Next i

For i = 0 To 9
    iSqlSuffix(i + 1) = ""
    sKey = "Suffix" + CStr(i + 1)
    h$ = Space(51)
    l& = GetPrivateProfileString("SqlSuffix", sKey, h$, h$, 50, wPara1.LwZ + "\user\OpDrives.ini")
    h = Trim$(Left$(h$, l&))
    If (Len(h) = 1) Then
        iSqlSuffix(i + 1) = UCase(h)
    End If
Next i

sProtocol = ""
For i = 1 To 10
    If (i = 1) Then
        sSection = "SqlServer"
    Else
        sSection = "SqlServer" + CStr(i)
    End If
    h$ = Space$(50)
    l& = GetPrivateProfileString(sSection, "SqlServer", h$, h$, 51, CurDir + "\SqlOp.ini")
    h = Trim$(Left$(h$, l&))
    If (h <> "") Then
        AnzSqlServers = i
        
        l = InStr(h, "\")
        If l > 0 Then
          iSqlServers(i) = wPara1.ServerRechnerName + Mid(h, l)
        Else
          iSqlServers(i) = h
        End If
        
        If (i = 1) Then
            sSection = "Connection"
        End If
            
        sKey = "Protocol"
        h$ = Space$(50)
        l& = GetPrivateProfileString(sSection, sKey, h$, h$, 51, CurDir + "\SqlOp.ini")
        h = Trim$(Left$(h$, l&))
        If (h <> "") Then
            iSqlServers(i) = h + ":" + iSqlServers(i)
        ElseIf (sProtocol <> "") Then
            iSqlServers(i) = sProtocol + ":" + iSqlServers(i)
        End If
        If (i = 1) Then
            sProtocol = h
        End If
            
        ConnectionString = ""
        
        sKey = "Provider"
        h$ = Space$(50)
        l& = GetPrivateProfileString(sSection, "Provider", h$, h$, 51, CurDir + "\SqlOp.ini")
        h = Trim$(Left$(h$, l&))
        If (h = "") Then
            iSqlConnectionString(i) = iSqlConnectionString(1)
        Else
            ConnectionString = ConnectionString + sKey + "=" + h$ + ";"
            
            sKey = "Integrated Security"
            h$ = Space$(50)
            l& = GetPrivateProfileString(sSection, sKey, h$, h$, 51, CurDir + "\SqlOp.ini")
            h = Trim$(Left$(h$, l&))
            If (h <> "") Then
                ConnectionString = ConnectionString + sKey + "=" + h$ + ";"
            Else
                sKey = "Trusted Connection"
                h$ = Space$(50)
                l& = GetPrivateProfileString(sSection, sKey, h$, h$, 51, CurDir + "\SqlOp.ini")
                h = Trim$(Left$(h$, l&))
                If (h <> "") Then
                    ConnectionString = ConnectionString + sKey + "=" + h$ + ";"
                End If
                
                sKey = "User Id"
                h$ = Space$(50)
                l& = GetPrivateProfileString(sSection, sKey, h$, h$, 51, CurDir + "\SqlOp.ini")
                h = Trim$(Left$(h$, l&))
                If (h <> "") Then
                    ConnectionString = ConnectionString + sKey + "=" + h$ + ";"
                End If
                
                sKey = "Password"
                h$ = Space$(50)
                l& = GetPrivateProfileString(sSection, sKey, h$, h$, 51, CurDir + "\SqlOp.ini")
                h = Trim$(Left$(h$, l&))
                If (h <> "") Then
                    ConnectionString = ConnectionString + sKey + "=" + h$ + ";"
                End If
            End If
        
            sKey = "Initial Catalog"
            h$ = Space$(50)
            l& = GetPrivateProfileString(sSection, sKey, h$, h$, 51, CurDir + "\SqlOp.ini")
            h = Trim$(Left$(h$, l&))
            ConnectionString = ConnectionString + sKey + "=" + h$ + ";"
            iSqlConnectionString(i) = ConnectionString
        End If
    Else
        Exit For
    End If
Next i

For i = 0 To UBound(iSqlServerDB)
    iSqlServerDB(i) = ""
    iSqlServerDBind(i) = 0
Next i
j = 0


On Error Resume Next
For i = 1 To AnzSqlServers
    With GlobalConn(i)
        .CursorLocation = adUseServer
        SqlServer = iSqlServers(i)
        .ConnectionString = iSqlConnectionString(i) + "Data Source=" + SqlServer
        .Open
        If (Err) And (i = 1) Then
            If (InStr(UCase(SqlServer), "EXPRESS") <= 0) Then
                Err.Clear
                SqlServer = SqlServer + "\SQLEXPRESS"
                .ConnectionString = ConnectionString + "Data Source=" + SqlServer
                .Open
            End If
        End If
        If (Err) Then
            Call MsgBox("Keine Verbindung zu SQl-Server '" + SqlServer + "'")
            ret = 0
        Else
            ret = True
            On Error GoTo DefErr
            GlobalComm(i).ActiveConnection = GlobalConn(i)
    
            SQLStr = "SELECT NAME FROM  Sys.DATABASES WHERE database_id>4"
            tRec.Open SQLStr$, GlobalConn(i)
            Do
                If (tRec.EOF) Then
                    Exit Do
                End If
                
                iSqlServerDB(j) = tRec!Name
                iSqlServerDBind(j) = i
                j = j + 1
                If (j >= 99) Then
                    Exit Do
                End If
        
                tRec.MoveNext
            Loop
            tRec.Close
        End If
    End With
Next i
On Error GoTo DefErr

If (MitStatus) Then
    sSection = "ActualSqlDatabases"
    For j = 0 To UBound(iSqlServerDB)
        sKey = "Database" + CStr(j + 1)
        If (iSqlServerDB(j) <> "") Then
            h = iSqlServerDB(j) + "," + CStr(iSqlServerDBind(j))
            l& = WritePrivateProfileString(sSection, sKey, h$, CurDir + "\SqlOp.ini")
        Else
            h = Space(50)
            l& = GetPrivateProfileString(sSection, sKey, h$, h$, 51, CurDir + "\SqlOp.ini")
            h = Trim$(Left$(h$, l&))
            If (h <> "") Then
                l& = WritePrivateProfileString(sSection, sKey, "", CurDir + "\SqlOp.ini")
            End If
        End If
    Next j
End If

If (ret) Then
'    l& = WritePrivateProfileString("SqlServer", "SqlServer", SqlServer$, CurDir + "\SqlOp.ini")

    sKey = "SqlServer"
    h$ = Space$(50)
    l& = GetPrivateProfileString("DataPath", sKey, h$, h$, 51, CurDir + "\SqlOp.ini")
    h = Trim$(Left$(h$, l&))
    If (h = "") Then
        SQLStr = "select physical_name from sys.database_files"
        tRec.Open SQLStr$, GlobalConn(1)
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
    iSqlServerDataPath = h$
    
    sKey = "Network"
    h$ = Space$(50)
    l& = GetPrivateProfileString("DataPath", sKey, h$, h$, 51, CurDir + "\SqlOp.ini")
    h = Trim$(Left$(h$, l&))
    If (h = "") Then
        h = "Z" + Mid$(iSqlServerDataPath, 2)
        l& = WritePrivateProfileString("DataPath", sKey, h$, CurDir + "\SqlOp.ini")
    End If
    iNetworkDataPath = h$
'Else
'    Call MsgBox("Keine Verbindung zu SqlServer" + SqlServer + "möglich!", vbCritical)
End If

SqlInit = ret

Call clsError.DefErrPop
End Function

'Function SqlInit%()
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Call clsError.DefErrFnc("SqlInit%")
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
'Dim j%, ret%, erg%, ind%, ind2%, SqlServerNötig%
'Dim h$, h2$
'
'Dim l As Long
'Dim entriesread As Long
'Dim totalentries As Long
'Dim hREsume As Long
'Dim bufptr As Long
'Dim Level As Long
'Dim prefmaxlen As Long
'Dim lType As Long
'Dim domain() As Byte
'Dim i As Long
'Dim sv100 As SV_100
'Dim strComputername$, strWorkgroup$, strSqlServers$
'Dim SqlServers$(10)
'Dim AnzSqlServers%
'Dim tRec As New ADODB.Recordset
'
'Dim sKey$, SQLStr$
'
'h$ = Space$(50)
'l& = GetPrivateProfileString("LetzteAktionen", "AktiveTaxe", h$, h$, 51, "z:\user\SysManag.ini")
'TaxeDatum = Trim$(Left$(h$, l&))
'If (TaxeDatum <> "") Then
'    TaxeDatum = "_" + TaxeDatum
'End If
'
'SqlConnectionstring(0) = "Provider=Microsoft.Jet.OLEDB.4.0;"
'SqlConnectionstring(1) = ""
'
'qlServerNötig = 0
'SqlServerDB(0) = "Artikel"
'SqlServerDB(1) = "Wawi"
'SqlServerDB(2) = "Verkauf"
'
'SqlServerNötig = True
'j = 3
'For i = 0 To 19
'    sKey = "Database" + CStr(i + 1)
'    h$ = Space(51)
'    l& = GetPrivateProfileString("SqlDatabases", sKey, h$, h$, 50, CurDir + "\SqlOp.ini")
'    h = Trim$(Left$(h$, l&))
'    If (h <> "") Then
'        SqlServerDB(j) = h
'        SqlServerNötig = True
'        j = j + 1
'    End If
'Next i
'
'For i = 0 To 9
'    SqlSuffix(i + 1) = ""
'    sKey = "Suffix" + CStr(i + 1)
'    h$ = Space(51)
'    l& = GetPrivateProfileString("SqlSuffix", sKey, h$, h$, 50, "z:\user\OpDrives.ini")
'    h = Trim$(Left$(h$, l&))
'    If (Len(h) = 1) Then
'        SqlSuffix(i + 1) = UCase(h)
'    End If
'Next i
'
'If (SqlServerNötig = 0) Then
'    SqlServer = ""
'    ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;"
'    SqlInit = True
'    Call clsError.DefErrPop: Exit Function
'End If
'
'sKey = "SqlServer"
'h$ = Space$(50)
'l& = GetPrivateProfileString("SqlServer", sKey, h$, h$, 51, CurDir + "\SqlOp.ini")
'h = Trim$(Left$(h$, l&))
'If (h = "") Then
'    SqlServer = ""
'    ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;"
'    SqlInit = True
'    Call clsError.DefErrPop: Exit Function
'End If
'
'
'If (h = "?") Then
'    AnzSqlServers = 0
'
'    strComputername = Environ("ComputerName")
'    strWorkgroup = ""
'    If (strComputername <> "") Then
'        Dim pWrkInfo As Long, WrkInfo(0) As WKSTA_INFO_102
'        Dim lResult As Long
'        lResult = NetWkstaGetInfo(StrConv("\\" & strComputername, vbUnicode), 102, pWrkInfo)
'        If (lResult = 0) Then
'           Dim cname As String
'           cname = String$(255, 0)
'           CopyMemory WrkInfo(0), ByVal pWrkInfo, ByVal Len(WrkInfo(0))
'           CopyMemory ByVal cname, ByVal WrkInfo(0).pwki100_langroup, ByVal 255
'           strWorkgroup = StripTerminator(StrConv(cname, vbFromUnicode))
'        End If
'    End If
'
'    strSqlServers = ""
'    If (strWorkgroup <> "") Then
'        Level = 100
'        prefmaxlen = -1
'        lType = SV_TYPE_SQLSERVER
'        domain = strWorkgroup & vbNullChar
'        l = NetServerEnum(ByVal 0&, Level, bufptr, prefmaxlen, entriesread, totalentries, lType, domain(0), hREsume)
'        If l = 0 Or l = 234& Then
'            For i = 0 To entriesread - 1
'                CopyMemory sv100, ByVal bufptr, Len(sv100)
'                strSqlServers = strSqlServers + Pointer2stringw(sv100.Name) + vbCrLf
'                SqlServers(AnzSqlServers) = Pointer2stringw(sv100.Name)
'                AnzSqlServers = AnzSqlServers + 1
'                bufptr = bufptr + Len(sv100)
'            Next i
'        End If
'        NetApiBufferFree bufptr '
'    End If
'
'    h2 = "Computer:" + vbTab + strComputername
'    h2 = h2 + vbCrLf + "Workgroup:" + vbTab + strWorkgroup
'    h2 = h2$ + vbCrLf + vbCrLf + "Installierte SQL-Server:" + vbCrLf + strSqlServers
'    MsgBox (h2$)
'Else
'    AnzSqlServers = 1
'    SqlServers(0) = h$
'End If
'
'ConnectionString = ""
'
'sKey = "Provider"
'h$ = Space$(50)
'l& = GetPrivateProfileString("Connection", "Provider", h$, h$, 51, CurDir + "\SqlOp.ini")
'h = Trim$(Left$(h$, l&))
'If (h <> "") Then
'    ConnectionString = ConnectionString + sKey + "=" + h$ + ";"
'End If
'
'sKey = "Integrated Security"
'h$ = Space$(50)
'l& = GetPrivateProfileString("Connection", sKey, h$, h$, 51, CurDir + "\SqlOp.ini")
'h = Trim$(Left$(h$, l&))
'If (h <> "") Then
'    ConnectionString = ConnectionString + sKey + "=" + h$ + ";"
'Else
'    sKey = "Trusted Connection"
'    h$ = Space$(50)
'    l& = GetPrivateProfileString("Connection", sKey, h$, h$, 51, CurDir + "\SqlOp.ini")
'    h = Trim$(Left$(h$, l&))
'    If (h <> "") Then
'        ConnectionString = ConnectionString + sKey + "=" + h$ + ";"
'    End If
'
'    sKey = "User Id"
'    h$ = Space$(50)
'    l& = GetPrivateProfileString("Connection", sKey, h$, h$, 51, CurDir + "\SqlOp.ini")
'    h = Trim$(Left$(h$, l&))
'    If (h <> "") Then
'        ConnectionString = ConnectionString + sKey + "=" + h$ + ";"
'    End If
'
'    sKey = "Password"
'    h$ = Space$(50)
'    l& = GetPrivateProfileString("Connection", sKey, h$, h$, 51, CurDir + "\SqlOp.ini")
'    h = Trim$(Left$(h$, l&))
'    If (h <> "") Then
'        ConnectionString = ConnectionString + sKey + "=" + h$ + ";"
'    End If
'End If
'
'sKey = "Initial Catalog"
'h$ = Space$(50)
'l& = GetPrivateProfileString("Connection", sKey, h$, h$, 51, CurDir + "\SqlOp.ini")
'h = Trim$(Left$(h$, l&))
'ConnectionString = ConnectionString + sKey + "=" + h$ + ";"
'SqlConnectionstring(1) = ConnectionString
'
''sKey = "Data Source"
''ConnectionString = ConnectionString + sKey + "=" + ";"
'
'On Error Resume Next
''On Error GoTo DefErr
'With GlobalConn
'    .CursorLocation = adUseServer
'    For i = 1 To AnzSqlServers
'        SqlServer = SqlServers(i - 1)
''        .ConnectionString = "Provider=SQLOLEDB; Integrated Security=SSPI; Initial Catalog=" + "''" + "; Data Source=" + h$
''        .ConnectionString = "Provider=SQLOLEDB; Trusted Connection=False; User Id=ELVIRA; Password=JESSICA; Initial Catalog=" + "''" + "; Data Source=" + h$
'        .ConnectionString = ConnectionString + "Data Source=" + SqlServer
'        .Open
'        If (Err) Then
'            If (InStr(UCase(SqlServer), "EXPRESS") <= 0) Then
'                Err.Clear
'                SqlServer = SqlServer + "\SQLEXPRESS"
'                .ConnectionString = ConnectionString + "Data Source=" + SqlServer
'                .Open
'            End If
'        End If
'        If (Err = 0) Then
'            ret = True
''            MsgBox ("Verbindung zu SQl-Server '" + SqlServer + "'")
'            On Error GoTo DefErr
'            GlobalComm.ActiveConnection = GlobalConn
'            Exit For
'        End If
'    Next i
'End With
'On Error GoTo DefErr
'
'If (ret) Then
'    l& = WritePrivateProfileString("SqlServer", "SqlServer", SqlServer$, CurDir + "\SqlOp.ini")
'
'    sKey = "SqlServer"
'    h$ = Space$(50)
'    l& = GetPrivateProfileString("DataPath", sKey, h$, h$, 51, CurDir + "\SqlOp.ini")
'    h = Trim$(Left$(h$, l&))
'    If (h = "") Then
'        SQLStr = "select physical_name from sys.database_files"
'        tRec.Open SQLStr$, GlobalConn
'        If Not (tRec.EOF) Then
'            h = tRec!physical_name
'            ind = 0
'            Do
'                ind2 = InStr(ind + 1, h, "\")
'                If (ind2 > 0) Then
'                    ind = ind2
'                Else
'                    If (ind > 0) Then
'                        h = Left(h, ind - 1)
'                    End If
'                    Exit Do
'                End If
'            Loop
'            l& = WritePrivateProfileString("DataPath", sKey, h$, CurDir + "\SqlOp.ini")
'        End If
'        tRec.Close
'    End If
'    iSqlServerDataPath = h$
'
'    sKey = "Network"
'    h$ = Space$(50)
'    l& = GetPrivateProfileString("DataPath", sKey, h$, h$, 51, CurDir + "\SqlOp.ini")
'    h = Trim$(Left$(h$, l&))
'    If (h = "") Then
'        h = "Z" + Mid$(iSqlServerDataPath, 2)
'        l& = WritePrivateProfileString("DataPath", sKey, h$, CurDir + "\SqlOp.ini")
'    End If
'    iNetworkDataPath = h$
''    erg% = CreateDirectory(iNetworkDataPath)
'Else
'    Call MsgBox("Keine Verbindung zu SqlServer" + SqlServer + "möglich!", vbCritical)
'End If
'
'SqlInit = ret
'
'Call clsError.DefErrPop
'End Function

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

Function IstSqlServerDB%(sDatabase$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("IstSqlServerDB%")
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
Dim i%, ret%

ret = 0
For i = 0 To UBound(iSqlServerDB)
    If (UCase(iSqlServerDB(i)) = UCase(sDatabase)) Then
'        ret = True
        ret = iSqlServerDBind(i)
        Exit For
    End If
Next i

IstSqlServerDB = ret

Call clsError.DefErrPop
End Function

Function SqlGetSuffix$()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("SqlGetSuffix$")
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
Dim ret$, sLw$

ret = ""

sLw = UCase(Left(CurDir, 1))
If (sLw <> "Z") Then
    For i = 1 To UBound(iSqlSuffix)
        If (sLw = iSqlSuffix(i)) Then
            ret = "_" + CStr(i)
            Exit For
        End If
    Next i
End If

SqlGetSuffix = ret

Call clsError.DefErrPop
End Function

Function SqlCheckDatabase%(sDatabase$, Optional iSqlAktiv% = True)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("SqlCheckDatabase%")
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
Dim i%, ret%
Dim ErrNumber&
Dim SQLStr$
Dim tRec As New ADODB.Recordset
Dim CheckDB As Database

ret% = 0
If (iSqlAktiv) Then
'    SQLStr = "SELECT NAME FROM  Sys.DATABASES"
'    tRec.Open SQLStr$, GlobalConn
'    Do
'        If (tRec.EOF) Then
'            Exit Do
'        End If
'
'        If (UCase(tRec!Name) = UCase(sDatabase)) Then
'            ret = True
'            Exit Do
'        End If
'
'        tRec.MoveNext
'    Loop
'    tRec.Close

    For i = 0 To UBound(iSqlServerDB)
        If (UCase(iSqlServerDB(i)) = UCase(sDatabase)) Then
'            ret = True
            ret = iSqlServerDBind(i)
            Exit For
        End If
    Next i
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

Call clsError.DefErrPop
End Function
'
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

Set ActiveConn = GlobalConn(1)

Call clsError.DefErrPop
End Function

Function SqlServer$(ind%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("SqlServer$")
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

SqlServer = iSqlServers(ind)

Call clsError.DefErrPop
End Function

Function SqlConnectionString$(ind%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("SqlConnectionsString$")
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

SqlConnectionString = iSqlConnectionString(ind)

Call clsError.DefErrPop
End Function

Function SqlServerDataPath$()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("SqlServerDataPath$")
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

SqlServerDataPath = iSqlServerDataPath

Call clsError.DefErrPop
End Function

Function NetworkDataPath$()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("NetworkDataPath$")
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

NetworkDataPath = iNetworkDataPath

Call clsError.DefErrPop
End Function

'Function TaxeDatum$()
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Call clsError.DefErrFnc("TaxeDatum$")
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
'
'TaxeDatum = iTaxeDatum
'
'Call clsError.DefErrPop
'End Function
Public Property Get TaxeDatum$()
TaxeDatum$ = iTaxeDatum
End Property
Public Property Let TaxeDatum(ByVal vNewValue$)
iTaxeDatum$ = vNewValue$
End Property


Function SqlCreateDatabase%(sDatabase$, Optional sDateien$ = "", Optional gInd% = 1)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("SqlCreateDatabase%")
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
Dim SQLStr$

If (gInd > AnzSqlServers) Then
    gInd = AnzSqlServers
End If

ret% = SqlDeleteDatabase(sDatabase)
If (ret) Then
    SQLStr = "CREATE DATABASE " + sDatabase
    SQLStr = SQLStr + " ON (NAME='" + sDateien + "', FILENAME='" + SqlServerDataPath + "\" + sDatabase + ".mdf" + "')"
    SQLStr = SQLStr + " LOG ON (NAME='" + sDateien + "_log" + "', FILENAME = '" + SqlServerDataPath + "\" + sDatabase + "_log.ldf" + "')"
'    SQLStr = SQLStr + " ON (NAME='" + SqlDatenDir + "\" + sDatabase + "', FILENAME='" + SqlDatenDir + "\" + sDatabase + "555.mdf" + "')"
'    SQLStr = SQLStr + " LOG ON (NAME='" + SqlDatenDir + "\" + sDatabase + "_log" + "', FILENAME = '" + SqlDatenDir + "\" + sDatabase + "_log555.ldf" + "')"

    ret = SqlCommand(SQLStr, gInd)
End If
'If (ret) Then
'    ret = SqlConnect(adoConn, adoComm, sDatabase, True, adUseServer)
'End If

SqlCreateDatabase = ret

Call clsError.DefErrPop
End Function


'Function SqlCreateDatabase%(adoConn As ADODB.Connection, adoComm As ADODB.Command, sDatabase$, sDateien$, sPfad$)
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Call DefErrFnc("SqlCreateDatabase%")
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
'Dim ret%
'Dim SQLStr$, SQLPfad$
'
''SQLPfad = "C" + Mid(sPfad, 2)
'SQLPfad = iSqlServerDataPath
''SQLPfad = CurDir
'
'ret% = SqlDeleteDatabase(sDatabase)
'If (ret) Then
'    SQLStr = "CREATE DATABASE " + sDatabase
'    SQLStr = SQLStr + " ON (NAME='" + sDateien + "', FILENAME='" + SQLPfad + "\" + sDatabase + ".mdf" + "')"
'    SQLStr = SQLStr + " LOG ON (NAME='" + sDateien + "_log" + "', FILENAME = '" + SQLPfad + "\" + sDatabase + "_log.ldf" + "')"
''    SQLStr = SQLStr + " ON (NAME='" + SqlDatenDir + "\" + sDatabase + "', FILENAME='" + SqlDatenDir + "\" + sDatabase + "555.mdf" + "')"
''    SQLStr = SQLStr + " LOG ON (NAME='" + SqlDatenDir + "\" + sDatabase + "_log" + "', FILENAME = '" + SqlDatenDir + "\" + sDatabase + "_log555.ldf" + "')"
'
'    ret = SqlCommand(GlobalComm, SQLStr)
'End If
'If (ret) Then
'    ret = SqlConnect(adoConn, adoComm, sDatabase, True, adUseServer)
'End If
'
'SqlCreateDatabase = ret
'
'Call DefErrPop
'End Function
'
Function SqlDeleteDatabase%(sDatabase$, Optional iSqlAktiv% = True)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("SqlDeleteDatabase%")
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
Dim ret%, gef%
Dim SQLStr$

ret = True

gef% = SqlCheckDatabase(sDatabase, iSqlAktiv)
If (gef) Then
    If (iSqlAktiv) Then
        SQLStr = "ALTER DATABASE " + sDatabase + " SET SINGLE_USER WITH ROLLBACK IMMEDIATE"
        ret = SqlCommand(SQLStr, gef)
        If (ret) Then
            SQLStr = "DROP DATABASE " + sDatabase
            ret = SqlCommand(SQLStr, gef)
        End If
    Else
    End If
End If

SqlDeleteDatabase = ret

Call clsError.DefErrPop
End Function

Function SqlCommand%(sCommand$, gInd%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("SqlCommand%")
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
Dim lRecs&

ret = True
With GlobalConn(gInd)
'    .CommandText = sCommand$
    .CommandTimeout = 300   '120
    On Error Resume Next
    Call .Execute(sCommand, lRecs&, adExecuteNoRecords)
    If (Err) Then
        SqlError = Err.Number
        SqlErrorDesc = Err.Description
        Call MsgBox("Fehler bei '" + sCommand + "': " + vbCrLf + Str(Err.Number) + " " + Err.Description, vbCritical)
        ret = 0
    End If
    On Error GoTo DefErr
End With

SqlCommand = ret

Call clsError.DefErrPop
End Function
'
