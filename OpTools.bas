Attribute VB_Name = "modWinTools"
Option Explicit

Private Const SPI_GETWORKAREA = 48



Public Const REG_SZ As Long = 1
Public Const REG_DWORD As Long = 4

Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003

Public Const ERROR_NONE = 0
Public Const ERROR_BADDB = 1
Public Const ERROR_BADKEY = 2
Public Const ERROR_CANTOPEN = 3
Public Const ERROR_CANTREAD = 4
Public Const ERROR_CANTWRITE = 5
Public Const ERROR_OUTOFMEMORY = 6
Public Const ERROR_ARENA_TRASHED = 7
Public Const ERROR_ACCESS_DENIED = 8
Public Const ERROR_INVALID_PARAMETERS = 87
Public Const ERROR_NO_MORE_ITEMS = 259

Public Const KEY_ALL_ACCESS = &H3F

Public Const REG_OPTION_NON_VOLATILE = 0

Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, _
    ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, ByVal lpSecurityAttributes _
    As Long, phkResult As Long, lpdwDisposition As Long) As Long
Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, _
    ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Declare Function RegQueryValueExString Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As _
    String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Declare Function RegQueryValueExLong Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As _
    String, ByVal lpReserved As Long, lpType As Long, lpData As Long, lpcbData As Long) As Long
Declare Function RegQueryValueExNULL Lib "advapi32.dll" Alias _
"RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData _
    As Long, lpcbData As Long) As Long
Declare Function RegSetValueExString Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, _
    ByVal Reserved As Long, ByVal dwType As Long, ByVal lpValue As String, ByVal cbData As Long) As Long
Declare Function RegSetValueExLong Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, _
    ByVal Reserved As Long, ByVal dwType As Long, lpValue As Long, ByVal cbData As Long) As Long





Private Const DefErrModul = "OPTOOLS.BAS"

Sub iClose(Handle%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("iClose")
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
If (Handle% > 0) Then
    Close #Handle%
    Handle% = 0
End If
Call DefErrPop
End Sub

Function CheckTask%(SuchTaskId&)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("GetTaskId")
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
Dim ThreadID
Dim CurrWnd&, x&

CheckTask% = True

CurrWnd& = GetWindow(frmAction.hWnd, GW_HWNDFIRST)
While (CurrWnd& <> 0)
  x& = GetWindowThreadProcessId(CurrWnd&, ThreadID)
  If (ThreadID = SuchTaskId&) Then Exit Function
  CurrWnd& = GetWindow(CurrWnd&, GW_HWNDNEXT)
Wend

CheckTask% = False

Call DefErrPop
End Function

Function GetTaskId(ByVal TaskName As String) As Long
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("GetTaskId")
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
Dim CurrWnd As Long
Dim Length As Integer
Dim ListItem As String
Dim x As Long
Dim ThreadID As Long
Dim i&

GetTaskId = 0

TaskName = UCase(TaskName)
CurrWnd = GetWindow(frmAction.hWnd, GW_HWNDFIRST)
Do While (CurrWnd <> 0)
    Length = GetWindowTextLength(CurrWnd)
    ListItem = Space(Length + 1)
    Length = GetWindowText(CurrWnd, ListItem, Length + 1)
    If (Length > 0) Then
        If (UCase(Left(ListItem, Len(TaskName))) = TaskName) Then
            x = GetWindowThreadProcessId(CurrWnd, ThreadID)
'            Print #PROTOKOLL%, ListItem; ThreadID
            GetTaskId = ThreadID
            Exit Do
        End If
    End If
    CurrWnd = GetWindow(CurrWnd, GW_HWNDNEXT)
Loop

Call DefErrPop
End Function

Function EntferneTask(ByVal TaskName As String) As Long
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("EntferneTask")
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
Dim CurrWnd As Long
Dim Length As Integer
Dim ListItem As String
Dim x As Long
Dim ThreadID As Long
Dim ProcHandle&
Dim i&

EntferneTask = 0

TaskName = UCase(TaskName)
CurrWnd = GetWindow(frmAction.hWnd, GW_HWNDFIRST)
Do While (CurrWnd <> 0)
    Length = GetWindowTextLength(CurrWnd)
    ListItem = Space(Length + 1)
    Length = GetWindowText(CurrWnd, ListItem, Length + 1)
    If (Length > 0) Then
        If (UCase(Left(ListItem, Len(TaskName))) = TaskName) Then
            x = GetWindowThreadProcessId(CurrWnd, ThreadID)
            ProcHandle& = OpenProcess(PROCESS_SYNCHRONIZE Or PROCESS_TERMINATE, False, ThreadID)
            x = PostMessage(CurrWnd, WM_CLOSE, 0, 0)
            x = WaitForSingleObject(ProcHandle&, 5000&)
            
'            Print #PROTOKOLL%, ListItem; ThreadID; ProcHandle&; X
            EntferneTask = ThreadID
            
            If (x <> 0) Then
                x = TerminateProcess(ProcHandle&, 0)
            End If
            Exit Do
        End If
    End If
    CurrWnd = GetWindow(CurrWnd, GW_HWNDNEXT)
Loop

Call DefErrPop
End Function

'Function CreateDirectory%(DirName$)
'Dim fTmp%, ret%
'Dim TestDateiName$
'
'On Error Resume Next
'
'ret% = True
'
'DirName$ = UCase$(RTrim$(DirName$))
'TestDateiName$ = DirName$ + "\TEST.$$$"
'
'fTmp% = FreeFile
'Open TestDateiName$ For Random As #fTmp%
'
'If Err <> 0 Then
'    Err = 0
'    MkDir DirName$
'    If Err <> 0 Then
'        Call MsgBox("Problem: Das Verzeichnis " + DirName$ + " kann nicht angelegt werden!" + vbCrLf + "Einspielvorgang wird abgebrochen !", vbExclamation)
'        ret% = False
'    End If
'Else
'    Close #fTmp%
'    Kill TestDateiName$
'End If
'
'CreateDirectory% = ret%
'
'End Function

Function CalcDirectorySize&(DirName$, layer%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("CalcDirectorySize&")
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
Dim h$, DirMask$, EntryName$
Dim SearchHandle&, BytesGesamt&, FileSize&
Dim FindDataRec As WIN32_FIND_DATA
Dim erg%, ind%, i%

DirMask$ = DirName$ + "\*.*"

BytesGesamt& = 0
SearchHandle& = FindFirstFile(DirMask$, FindDataRec)

If (SearchHandle& = INVALID_HANDLE_VALUE) Then
    CalcDirectorySize& = 0&
    Call DefErrPop: Exit Function
End If

Do
    h$ = FindDataRec.cFileName
    ind% = InStr(h$, Chr$(0))
    If (ind% > 0) Then h$ = Left$(h$, ind% - 1)
    EntryName$ = h$
    
        FileSize& = FindDataRec.nFileSizeHigh * MAXDWORD + FindDataRec.nFileSizeLow
        BytesGesamt& = BytesGesamt& + FileSize&
'        If (prot% = True) Then
'            For i% = 1 To layer%
'                Print #PROTOKOLL%, vbTab;
'            Next i%
'            Print #PROTOKOLL%, EntryName$; FileSize&
'        End If
    
    If ((EntryName$ = ".") Or (EntryName$ = "..")) Then
    ElseIf (FindDataRec.dwFileAttributes = FILE_ATTRIBUTE_DIRECTORY) Then
        h$ = DirName$ + "\" + EntryName$
        BytesGesamt& = BytesGesamt& + CalcDirectorySize(h$, layer% + 1)
    Else
    End If
    
    erg% = FindNextFile(SearchHandle, FindDataRec)
    If (erg% = 0) Then Exit Do
Loop
erg% = FindClose(SearchHandle&)

'If (prot% = True) Then
'    For i% = 1 To layer%
'        Print #PROTOKOLL%, vbTab;
'    Next i%
'    Print #PROTOKOLL%, DirMask$; BytesGesamt&
'End If

CalcDirectorySize& = BytesGesamt&
Call DefErrPop
End Function
       
Sub DeleteDirectory(DirName$, layer%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("DeleteDirectory")
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
Dim h$, DirMask$, EntryName$
Dim SearchHandle&
Dim FindDataRec As WIN32_FIND_DATA
Dim erg%, ind%, i%

DirMask$ = DirName$ + "\*.*"

SearchHandle& = FindFirstFile(DirMask$, FindDataRec)

If (SearchHandle& = INVALID_HANDLE_VALUE) Then
    Call DefErrPop: Exit Sub
End If

Do
    h$ = FindDataRec.cFileName
    ind% = InStr(h$, Chr$(0))
    If (ind% > 0) Then h$ = Left$(h$, ind% - 1)
    EntryName$ = h$

    If ((EntryName$ = ".") Or (EntryName$ = "..")) Then
    ElseIf (FindDataRec.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) Then
        h$ = DirName$ + "\" + EntryName$
        Call DeleteDirectory(h$, layer% + 1)
    Else
        h$ = DirName$ + "\" + EntryName$
        On Error Resume Next
        Kill h$
        On Error GoTo DefErr
    End If

    erg% = FindNextFile(SearchHandle, FindDataRec)
    If (erg% = 0) Then Exit Do
Loop
erg% = FindClose(SearchHandle&)

On Error Resume Next
RmDir DirName$
On Error GoTo DefErr

Call DefErrPop
End Sub

Sub StartAnimation(hForm As Object, Optional text$ = "Aufgabe wird bearbeitet ...")
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("StartAnimation")
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
With hForm
    .MousePointer = vbHourglass
    .aniAnimation.Open "findfile.avi"
    .lblAnimation.Caption = text$
    .picAnimationBack.Left = (.ScaleWidth - .picAnimationBack.Width) / 2
    .picAnimationBack.Top = (.ScaleHeight - .picAnimationBack.Height) / 2
'    .picAnimationBack.Left = .Left + (.ScaleWidth - .picAnimationBack.Width) / 2
'    .picAnimationBack.Top = .Top + (.ScaleHeight - .picAnimationBack.Height) / 2
    .picAnimationBack.Visible = True
    .aniAnimation.Play
    .Refresh
End With
Call DefErrPop
End Sub

Sub StopAnimation(hForm As Object)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("StopAnimation")
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
With hForm
    .aniAnimation.Stop
    .picAnimationBack.Visible = False
    .MousePointer = vbDefault
End With
Call DefErrPop
End Sub

Sub DruckKopf(header$, Typ$, Optional KopfZeile$ = "", Optional TypHeight% = 18)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("DruckKopf")
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
Dim l&, i%, pos%, h&, x%, y%, heute$, SeitenNr$

If (KopfZeile$ = "") Then
    KopfZeile$ = frmAction.Caption
End If

heute$ = Format(Day(Date), "00") + "-"
heute$ = heute$ + Format(Month(Date), "00") + "-"
heute$ = heute$ + Format(Year(Date), "0000")
heute$ = heute$ + " " + Time$

With Printer
    .CurrentX = 0: .CurrentY = 0
    .Font.Size = 14
    Printer.Print KopfZeile$
    
    DruckSeite% = DruckSeite% + 1
    SeitenNr$ = "-" + Str$(DruckSeite%) + " -"
    l& = .TextWidth(SeitenNr$)
    .CurrentX = (.ScaleWidth - l&) / 2: .CurrentY = 0
    Printer.Print SeitenNr$
    
    l& = .TextWidth(heute$)
    .CurrentX = .ScaleWidth - l& - 10: .CurrentY = 0
    Printer.Print heute$
    
    .CurrentX = 0
    l& = .TextHeight("A") + 500
    .CurrentY = l&
    
    .Font.Size = 18
    l& = .TextWidth(header$)
    h& = .TextHeight("A")
        
    x% = (.ScaleWidth - (l& + 800)) / 2
    y% = .CurrentY
    .DrawWidth = 2
    RoundRect .hdc, x% / .TwipsPerPixelX, y% / .TwipsPerPixelY, (x% + l& + 800) / .TwipsPerPixelX, (y% + h& + 800) / .TwipsPerPixelY, 200, 200
    .CurrentX = (.ScaleWidth - l&) / 2
    .CurrentY = y% + 400
    Printer.Print header$
    
    If (Typ$ <> "") Then
        .CurrentY = 2500
        .Font.Size = TypHeight%
    '    l& = .TextWidth(KopfStr$)
    '    .CurrentX = (.ScaleWidth - l&) / 2
        .CurrentX = 0
        Printer.Print Typ$ + ":"
    End If
    
    .CurrentY = 3100
    
    If (DruckFontSize% > 0) Then
        .Font.Size = DruckFontSize%
    Else
        .Font.Size = 12
    End If
End With

Call DefErrPop
End Sub

Sub DruckFuss(Optional NewPage% = True)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("DruckFuss")
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
Dim l&, h&, y%

With Printer
    .Font.Bold = False
    .Font.Size = 14
    l& = .TextWidth(para.Fistam(0))
    h& = .TextHeight("A")
    
    .CurrentX = 0
    .CurrentY = .ScaleHeight - 3 * h&
    y% = .CurrentY
    Printer.Line (0, y%)-(.ScaleWidth, y%)
    
    .CurrentX = 0
    .CurrentY = .CurrentY + 200
    Printer.Print para.Fistam(0)
    .CurrentX = 0
    Printer.Print para.Fistam(1);
    
    If (NewPage% = True) Then .NewPage
End With
Call DefErrPop
End Sub

Function iTrim$(Source$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("iTrim$")
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
Dim ind%
Dim ret$

ret$ = Source$
Do
    ind% = InStr(ret$, Chr$(0))
    If (ind% = 0) Then
        Exit Do
    Else
        Mid$(ret$, ind%, 1) = " "
    End If
Loop
iTrim$ = ret$

Call DefErrPop
End Function

